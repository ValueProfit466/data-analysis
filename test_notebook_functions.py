"""
Test script to evaluate dataset_analysis_eurostat.ipynb functions on CSV files
"""
import pandas as pd
import numpy as np
from pathlib import Path
import warnings
warnings.filterwarnings('ignore')

# ============================================================================
# LOAD THE FUNCTIONS FROM THE NOTEBOOK
# ============================================================================

def load_eurostat_csv(filepath: str) -> pd.DataFrame:
    """Load Eurostat CSV file with proper data type handling."""
    df = pd.read_csv(filepath, low_memory=False)

    if 'TIME_PERIOD' in df.columns:
        df['TIME_PERIOD'] = pd.to_datetime(df['TIME_PERIOD'], format='%Y', errors='coerce')

    if 'OBS_VALUE' in df.columns:
        df['OBS_VALUE'] = pd.to_numeric(df['OBS_VALUE'], errors='coerce')

    for col in df.columns:
        if df[col].dtype == 'object' and col not in ['LAST UPDATE']:
            num_unique = df[col].nunique()
            if num_unique / len(df) < 0.5:
                df[col] = df[col].astype('category')

    return df


def load_oecd_csv(filepath: str) -> pd.DataFrame:
    """Load OECD CSV file handling the double-header format."""
    with open(filepath, 'r') as f:
        first_line = f.readline()

    if 'STRUCTURE' in first_line:
        df = pd.read_csv(filepath, skiprows=1, low_memory=False)
    else:
        df = pd.read_csv(filepath, low_memory=False)

    if 'TIME_PERIOD' in df.columns:
        df['TIME_PERIOD'] = pd.to_datetime(df['TIME_PERIOD'], format='%Y', errors='coerce')
    elif 'Time period' in df.columns:
        df['TIME_PERIOD'] = pd.to_datetime(df['Time period'], format='%Y', errors='coerce')

    if 'OBS_VALUE' in df.columns:
        df['OBS_VALUE'] = pd.to_numeric(df['OBS_VALUE'], errors='coerce')
    elif 'Observation value' in df.columns:
        df['OBS_VALUE'] = pd.to_numeric(df['Observation value'], errors='coerce')

    for col in df.columns:
        if df[col].dtype == 'object':
            num_unique = df[col].nunique()
            if num_unique / len(df) < 0.5:
                df[col] = df[col].astype('category')

    return df

# ============================================================================
# TEST FUNCTIONS
# ============================================================================

def test_csv_files():
    """Test loading and analyzing all CSV files"""
    csv_files = list(Path('.').glob('*.csv'))

    print("="*80)
    print("TESTING DATASET ANALYSIS NOTEBOOK ON CSV FILES")
    print("="*80)
    print(f"\nFound {len(csv_files)} CSV files\n")

    results = {}

    for filepath in csv_files:
        filename = filepath.name
        print(f"\n{'='*80}")
        print(f"Testing: {filename}")
        print('='*80)

        try:
            # Load the file
            if 'OECD' in filename:
                df = load_oecd_csv(str(filepath))
            else:
                df = load_eurostat_csv(str(filepath))

            print(f"✓ Loaded successfully")
            print(f"  Shape: {df.shape[0]:,} rows × {df.shape[1]} columns")
            print(f"  Memory: {df.memory_usage(deep=True).sum() / 1024**2:.2f} MB")

            # Analyze structure
            print(f"\n📊 COLUMN ANALYSIS:")
            print(f"  Total columns: {len(df.columns)}")

            # Identify key columns
            geo_col = None
            for col in ['geo', 'REF_AREA', 'LOCATION', 'Country']:
                if col in df.columns:
                    geo_col = col
                    break

            time_col = None
            for col in ['TIME_PERIOD', 'Year', 'Date']:
                if col in df.columns:
                    time_col = col
                    break

            value_col = None
            for col in ['OBS_VALUE', 'Value', 'Observation']:
                if col in df.columns:
                    value_col = col
                    break

            if geo_col:
                print(f"  Geographic: {geo_col} ({df[geo_col].nunique()} unique values)")
            if time_col:
                if df[time_col].dtype == 'datetime64[ns]':
                    print(f"  Time: {time_col} ({df[time_col].min().year} - {df[time_col].max().year})")
                else:
                    print(f"  Time: {time_col}")
            if value_col:
                print(f"  Value: {value_col} ({df[value_col].notna().sum():,} non-null)")

            # Show filterable columns (categorical with reasonable unique values)
            print(f"\n🔍 FILTERABLE COLUMNS (for subset selection):")
            skip = ['DATAFLOW', 'LAST UPDATE', 'TIME_PERIOD', 'OBS_VALUE', geo_col]
            filterable = []

            for col in df.columns:
                if col in skip:
                    continue
                n_unique = df[col].nunique()
                if 0 < n_unique <= 50:
                    filterable.append(col)
                    print(f"  • {col}: {n_unique} unique values")
                    # Show first few values
                    top_values = df[col].value_counts().head(3)
                    for val, count in top_values.items():
                        print(f"      - {val}: {count:,} rows")

            if not filterable:
                print("  (No obvious filterable columns found)")

            # Check data quality
            print(f"\n📈 DATA QUALITY:")
            missing = df.isnull().sum().sum()
            print(f"  Missing values: {missing:,} ({missing/df.size*100:.1f}%)")
            print(f"  Duplicates: {df.duplicated().sum():,}")

            results[filename] = {
                'status': 'success',
                'shape': df.shape,
                'geo_col': geo_col,
                'time_col': time_col,
                'value_col': value_col,
                'filterable_cols': filterable
            }

        except Exception as e:
            print(f"✗ Error: {str(e)}")
            results[filename] = {
                'status': 'error',
                'error': str(e)
            }

    return results


def suggest_improvements(results):
    """Provide suggestions based on test results"""
    print("\n\n")
    print("="*80)
    print("SUGGESTIONS FOR IMPROVEMENT")
    print("="*80)

    print("\n1. CURRENT FILTERING APPROACH:")
    print("   ✓ The notebook has a manual filtering section (cell f5650094)")
    print("   ✓ Users need to edit code to change filters")
    print("   ✗ Not user-friendly for non-coders")
    print("   ✗ Filter values are hard-coded")

    print("\n2. XLSX EXPORT:")
    print("   ✗ Current export only supports CSV (cell 23)")
    print("   ✗ No multi-sheet XLSX export for filtered subsets")
    print("   ✗ Missing openpyxl/xlsxwriter support check")

    print("\n3. RECOMMENDED ENHANCEMENTS:")
    print("   📌 Add interactive filter configuration")
    print("   📌 Create filter_and_export() function")
    print("   📌 Support XLSX export with multiple sheets")
    print("   📌 Add filter validation and suggestions")
    print("   📌 Create filter templates for common use cases")
    print("   📌 Add summary statistics to exported files")

    print("\n4. SPECIFIC IMPROVEMENTS NEEDED:")
    print("   a) Filter Management:")
    print("      • Auto-detect filterable columns")
    print("      • Show available values before filtering")
    print("      • Validate filter values exist in data")
    print("      • Support multiple filter sets")

    print("   b) XLSX Export:")
    print("      • Export filtered data to XLSX")
    print("      • Create multiple sheets (raw, summary, charts)")
    print("      • Add auto-formatting (headers, number formats)")
    print("      • Include metadata sheet with filter info")

    print("   c) Workflow Optimization:")
    print("      • Batch filter + export multiple datasets")
    print("      • Save/load filter configurations")
    print("      • Generate filter documentation")

    print("\n5. DATASET-SPECIFIC INSIGHTS:")
    successful = [k for k, v in results.items() if v['status'] == 'success']
    print(f"   • Successfully tested {len(successful)}/{len(results)} files")
    print(f"   • All have similar structure (good for standardization)")
    print(f"   • Filtering columns vary by dataset (need flexible approach)")


if __name__ == '__main__':
    results = test_csv_files()
    suggest_improvements(results)

    print("\n\n" + "="*80)
    print("TEST COMPLETE")
    print("="*80)
