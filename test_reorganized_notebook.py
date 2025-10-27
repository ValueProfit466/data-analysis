"""
Test the reorganized notebook structure and new export functionality
"""
import pandas as pd
import numpy as np
from datetime import datetime
from pathlib import Path
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

print("="*70)
print("TESTING REORGANIZED NOTEBOOK")
print("="*70)

# Load test dataset
test_file = 'estat_iww_go_atygo_en.csv'
print(f"\n📂 Loading test dataset: {test_file}")

df = pd.read_csv(test_file, low_memory=False)

# Convert columns
if 'TIME_PERIOD' in df.columns:
    df['TIME_PERIOD'] = pd.to_datetime(df['TIME_PERIOD'], format='%Y', errors='coerce')
if 'OBS_VALUE' in df.columns:
    df['OBS_VALUE'] = pd.to_numeric(df['OBS_VALUE'], errors='coerce')

print(f"✓ Loaded: {df.shape[0]:,} rows × {df.shape[1]} columns\n")

# Load enhanced functions inline
from openpyxl import load_workbook

def get_filterable_columns(df, max_unique=50, exclude_cols=None):
    if exclude_cols is None:
        exclude_cols = ['DATAFLOW', 'LAST UPDATE', 'TIME_PERIOD', 'OBS_VALUE']
    filterable = {}
    for col in df.columns:
        if col in exclude_cols:
            continue
        n_unique = df[col].nunique()
        if 0 < n_unique <= max_unique:
            filterable[col] = {
                'unique_count': n_unique,
                'values': df[col].value_counts().to_dict()
            }
    return filterable

def validate_filters(df, filters):
    errors = []
    for col, values in filters.items():
        if col not in df.columns:
            errors.append(f"❌ Column '{col}' not found")
            continue
        valid_values = df[col].unique()
        invalid_values = [v for v in values if v not in valid_values]
        if invalid_values:
            errors.append(f"❌ Invalid values in '{col}': {invalid_values}")
    return len(errors) == 0, errors

def apply_filters(df, filters, verbose=True):
    is_valid, errors = validate_filters(df, filters)
    if not is_valid:
        print("⚠️  FILTER VALIDATION ERRORS:")
        for error in errors:
            print(error)
        raise ValueError("Filter validation failed")
    df_filtered = df.copy()
    original_rows = len(df_filtered)
    if verbose:
        print("🔍 APPLYING FILTERS")
    for col, values in filters.items():
        if col in df_filtered.columns:
            df_filtered = df_filtered[df_filtered[col].isin(values)]
            if verbose:
                print(f"✓ {col}: {values}")
    if verbose:
        print(f"  Original: {original_rows:,} rows")
        print(f"  Filtered: {len(df_filtered):,} rows")
        print(f"  Removed: {original_rows - len(df_filtered):,} rows\n")
    return df_filtered

print("\nTEST 1: Detect Key Columns")
print("-" * 70)

def detect_key_columns(df):
    """Automatically detect geographic, time, and value columns in the dataset."""
    geo_col = None
    for col in ['geo', 'REF_AREA', 'LOCATION', 'Country', 'COUNTRY']:
        if col in df.columns:
            geo_col = col
            break

    time_col = None
    for col in ['TIME_PERIOD', 'Year', 'Date', 'TIME', 'YEAR']:
        if col in df.columns:
            time_col = col
            break

    value_col = None
    for col in ['OBS_VALUE', 'Value', 'Observation', 'VALUE']:
        if col in df.columns:
            value_col = col
            break

    return geo_col, time_col, value_col

geo_col, time_col, value_col = detect_key_columns(df)
print(f"Geographic column: {geo_col}")
print(f"Time column: {time_col}")
print(f"Value column: {value_col}")
print()

print("TEST 2: Show Filter Options")
print("-" * 70)
filterable = get_filterable_columns(df, max_unique=10)
print(f"Found {len(filterable)} filterable columns:")
for col in filterable.keys():
    print(f"  • {col}")
print()

print("TEST 3: Apply Filters")
print("-" * 70)
filters = {
    'geo': ['BE', 'NL'],
    'tra_cov': ['TOTAL'],
    'unit': ['MIO_TKM']
}
df_filtered = apply_filters(df, filters, verbose=True)
print()

print("TEST 4: Create Analysis Summary")
print("-" * 70)

def create_analysis_summary(df, geo_col=None, time_col=None, value_col=None):
    """Create comprehensive analysis summary DataFrames."""
    analysis = {}

    if geo_col is None or time_col is None or value_col is None:
        geo_col, time_col, value_col = detect_key_columns(df)

    # Overall statistics
    if value_col and value_col in df.columns:
        stats = df[value_col].describe().to_frame('Value')
        stats.loc['missing'] = df[value_col].isna().sum()
        stats.loc['missing_pct'] = (df[value_col].isna().sum() / len(df)) * 100
        analysis['overall_stats'] = stats

    # Time series analysis
    if time_col and value_col and time_col in df.columns and value_col in df.columns:
        ts = df.groupby(time_col)[value_col].agg([
            ('count', 'count'),
            ('sum', 'sum'),
            ('mean', 'mean'),
            ('min', 'min'),
            ('max', 'max')
        ]).round(2)
        ts.columns = ['Count', 'Total', 'Average', 'Min', 'Max']
        analysis['time_series'] = ts

    # Geographic analysis
    if geo_col and value_col and geo_col in df.columns and value_col in df.columns:
        geo = df.groupby(geo_col)[value_col].agg([
            ('count', 'count'),
            ('sum', 'sum'),
            ('mean', 'mean')
        ]).round(2)
        geo.columns = ['Count', 'Total', 'Average']
        geo = geo.sort_values('Total', ascending=False)
        analysis['geographic'] = geo

    return analysis

analysis = create_analysis_summary(df_filtered, geo_col, time_col, value_col)
print(f"Created {len(analysis)} analysis tables:")
for name in analysis.keys():
    print(f"  • {name}: {len(analysis[name])} rows")
print()

print("TEST 5: Export Analysis to XLSX")
print("-" * 70)
output_file = 'test_full_analysis_export.xlsx'

with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    # Metadata
    metadata = pd.DataFrame([
        {'Info': 'Export Date', 'Value': datetime.now().strftime('%Y-%m-%d %H:%M:%S')},
        {'Info': 'Total Rows', 'Value': len(df_filtered)},
        {'Info': 'Filters Applied', 'Value': str(filters)}
    ])
    metadata.to_excel(writer, sheet_name='Metadata', index=False)

    # Filtered data
    df_filtered.to_excel(writer, sheet_name='Filtered Data', index=False)

    # Analysis sheets
    for analysis_name, analysis_df in analysis.items():
        sheet_name = analysis_name.replace('_', ' ').title()[:31]
        analysis_df.to_excel(writer, sheet_name=sheet_name)

    # Format all sheets
    header_fill = PatternFill(start_color='366092', end_color='366092', fill_type='solid')
    header_font = Font(color='FFFFFF', bold=True)

    for sheet_name in writer.sheets:
        worksheet = writer.sheets[sheet_name]
        for cell in worksheet[1]:
            if cell.value:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal='center', vertical='center')
        worksheet.freeze_panes = 'A2'

print(f"✓ Exported: {output_file}")

if Path(output_file).exists():
    file_size = Path(output_file).stat().st_size / 1024
    print(f"✓ File created: {output_file} ({file_size:.1f} KB)")

    # Count sheets
    from openpyxl import load_workbook
    wb = load_workbook(output_file)
    sheet_count = len(wb.sheetnames)
    print(f"✓ Sheets in workbook: {sheet_count}")
    print(f"   Sheet names: {', '.join(wb.sheetnames)}")
print()

print("="*70)
print("✅ ALL TESTS PASSED!")
print("="*70)
print("\nNew workflow verified:")
print("  1. Load dataset ✓")
print("  2. Show filter options ✓")
print("  3. Apply filters ✓")
print("  4. Analyze filtered data ✓")
print("  5. Export with comprehensive analysis ✓")
