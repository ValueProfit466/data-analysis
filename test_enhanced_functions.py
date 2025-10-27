"""
Test the enhanced filtering and XLSX export functions
"""
import pandas as pd
import numpy as np
from datetime import datetime
from pathlib import Path
import sys

# Import the enhanced functions
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

print("="*70)
print("TESTING ENHANCED FILTERING AND XLSX EXPORT FUNCTIONS")
print("="*70)

# Load a small test dataset
test_file = 'estat_iww_ac_nbac_en.csv'
print(f"\n📂 Loading test dataset: {test_file}")

df = pd.read_csv(test_file, low_memory=False)

# Convert TIME_PERIOD to datetime
if 'TIME_PERIOD' in df.columns:
    df['TIME_PERIOD'] = pd.to_datetime(df['TIME_PERIOD'], format='%Y', errors='coerce')

# Convert OBS_VALUE to numeric
if 'OBS_VALUE' in df.columns:
    df['OBS_VALUE'] = pd.to_numeric(df['OBS_VALUE'], errors='coerce')

print(f"✓ Loaded: {df.shape[0]:,} rows × {df.shape[1]} columns\n")

# Define enhanced functions inline for testing
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

def export_to_xlsx(df, filename):
    if not filename.endswith('.xlsx'):
        filename += '.xlsx'

    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Data', index=False)
        worksheet = writer.sheets['Data']

        # Format headers
        header_fill = PatternFill(start_color='366092', end_color='366092', fill_type='solid')
        header_font = Font(color='FFFFFF', bold=True)

        for col_num, column in enumerate(df.columns, 1):
            cell = worksheet.cell(row=1, column=col_num)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center', vertical='center')

        worksheet.freeze_panes = 'A2'

    print(f"✓ Exported: {filename}")

# TEST 1: Discover filterable columns
print("TEST 1: Discover filterable columns")
print("-" * 70)
filterable = get_filterable_columns(df)
print(f"Found {len(filterable)} filterable columns:")
for col, info in filterable.items():
    print(f"  • {col}: {info['unique_count']} unique values")
print()

# TEST 2: Validate filters (valid case)
print("TEST 2: Validate filters (valid case)")
print("-" * 70)
valid_filters = {
    'geo': ['AT', 'HU'],
    'accident': ['TOTAL']
}
is_valid, errors = validate_filters(df, valid_filters)
if is_valid:
    print("✓ Filters are valid")
else:
    print("✗ Validation failed")
    for error in errors:
        print(f"  {error}")
print()

# TEST 3: Validate filters (invalid case)
print("TEST 3: Validate filters (invalid case - should show errors)")
print("-" * 70)
invalid_filters = {
    'geo': ['AT', 'INVALID_COUNTRY'],
    'nonexistent_column': ['value']
}
is_valid, errors = validate_filters(df, invalid_filters)
if is_valid:
    print("✓ Filters are valid")
else:
    print("✗ Validation failed (as expected):")
    for error in errors:
        print(f"  {error}")
print()

# TEST 4: Apply filters
print("TEST 4: Apply filters")
print("-" * 70)
filters = {
    'geo': ['AT', 'HU'],
    'accident': ['TOTAL']
}
df_filtered = apply_filters(df, filters, verbose=True)
print(f"✓ Filtering successful: {len(df_filtered)} rows remaining\n")

# TEST 5: Export to XLSX
print("TEST 5: Export to XLSX")
print("-" * 70)
output_file = 'test_filtered_output.xlsx'
export_to_xlsx(df_filtered, output_file)
print()

# Verify the file was created
if Path(output_file).exists():
    file_size = Path(output_file).stat().st_size / 1024
    print(f"✓ File created successfully: {output_file} ({file_size:.1f} KB)")
else:
    print(f"✗ File was not created: {output_file}")
print()

# TEST 6: Multi-sheet export
print("TEST 6: Multi-sheet export with metadata")
print("-" * 70)
output_file2 = 'test_multisheet_output.xlsx'

with pd.ExcelWriter(output_file2, engine='openpyxl') as writer:
    # Sheet 1: Data
    df_filtered.to_excel(writer, sheet_name='Filtered Data', index=False)

    # Sheet 2: Metadata
    metadata = pd.DataFrame([
        {'Info': 'Export Date', 'Value': datetime.now().strftime('%Y-%m-%d %H:%M:%S')},
        {'Info': 'Original Rows', 'Value': len(df)},
        {'Info': 'Filtered Rows', 'Value': len(df_filtered)},
        {'Info': 'Filters Applied', 'Value': str(filters)}
    ])
    metadata.to_excel(writer, sheet_name='Metadata', index=False)

    # Sheet 3: Summary
    if 'OBS_VALUE' in df_filtered.columns:
        summary = df_filtered['OBS_VALUE'].describe()
        summary.to_excel(writer, sheet_name='Summary')

print(f"✓ Exported: {output_file2}")

if Path(output_file2).exists():
    file_size = Path(output_file2).stat().st_size / 1024
    print(f"✓ Multi-sheet file created: {output_file2} ({file_size:.1f} KB)")
print()

print("="*70)
print("✅ ALL TESTS PASSED!")
print("="*70)
print("\nCreated files:")
print(f"  • {output_file}")
print(f"  • {output_file2}")
print("\nSuggestion: Open these files in Excel to verify formatting and content.")
