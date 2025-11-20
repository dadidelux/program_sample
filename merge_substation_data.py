"""
Substation Data Merge and Validation Script
Merges SUB1 and SUB2 CSV files, cross-references with TLS Excel file,
and generates validation reports.
"""

import pandas as pd
import openpyxl
import os
from pathlib import Path

# File paths
BASE_DIR = Path(__file__).parent
DATASETS_DIR = BASE_DIR / "Datasets"
OUTPUT_DIR = BASE_DIR / "Final"

SUB1_FILE = DATASETS_DIR / "SUB1.csv"
SUB2_FILE = DATASETS_DIR / "SUB2.csv"
TLS_FILE = DATASETS_DIR / "SUB1-SUB2 115 kV -XcelUpdate.xlsx"

# Output files
MERGED_FILE = OUTPUT_DIR / "SUB1-SUB2 115kV.csv"
HIGHLIGHTED_FILE = OUTPUT_DIR / "SUB1-SUB2 115kV_highlighted.csv"
UPDATED_FILE = OUTPUT_DIR / "SUB1-SUB2 115kV_updated.csv"
SUMMARY_FILE = OUTPUT_DIR / "SUB1-SUB2 115kV_summary_report.csv"

# Column mapping: CSV column -> Excel column (source of truth naming)
# The Excel file is the source of truth, so we rename CSV columns to match Excel naming
COLUMN_RENAME_MAP = {
    'AMP Rating': 'High Rating',      # SN (N)
    'AMP Rating.1': 'High Rating.1',  # SE (A)
    'AMP Rating.2': 'High Rating.2',  # WN (B)
    'AMP Rating.3': 'High Rating.3',  # WE (C)
    'AMP Rating.4': 'High Rating.4',  # 5th rating type (CSV only, doesn't exist in Excel)
    'High KV': 'High kV',
    'Low KV': 'Low kV',
    'Tertiary KV': 'Tertiary kV',
}

# Columns that exist in Excel but not in CSV (will be added with NaN initially)
EXCEL_ONLY_COLUMNS = [
    'Low Rating',      # SN (N)
    'Low Rating.1',    # SE (A)
    'Low Rating.2',    # WN (B)
    'Low Rating.3',    # WE (C)
    'Line Number',
    'Type of Change',  # Will be populated during update
]


def load_csv_files():
    """Load SUB1 and SUB2 CSV files."""
    print("Loading CSV files...")

    try:
        sub1_df = pd.read_csv(SUB1_FILE)
        print(f"  - Loaded SUB1.csv: {len(sub1_df)} rows")

        sub2_df = pd.read_csv(SUB2_FILE)
        print(f"  - Loaded SUB2.csv: {len(sub2_df)} rows")

        return sub1_df, sub2_df
    except Exception as e:
        print(f"Error loading CSV files: {e}")
        raise


def load_tls_file():
    """Load TLS Excel file."""
    print("Loading TLS Excel file...")

    try:
        # Read the CAISO Update sheet which contains the component data
        tls_df = pd.read_excel(TLS_FILE, sheet_name='CAISO Update', engine='openpyxl')
        print(f"  - Loaded TLS file: {len(tls_df)} rows")
        print(f"  - Columns: {list(tls_df.columns)[:10]}...")  # Show first 10 columns
        return tls_df
    except Exception as e:
        print(f"Error loading TLS file: {e}")
        raise


def merge_csv_files(sub1_df, sub2_df):
    """Merge SUB1 and SUB2 dataframes and standardize to Excel column names."""
    print("\nMerging CSV files...")

    # Concatenate dataframes
    merged_df = pd.concat([sub1_df, sub2_df], ignore_index=True)
    print(f"  - Combined rows: {len(merged_df)}")

    # Remove duplicates based on OID
    initial_count = len(merged_df)
    merged_df = merged_df.drop_duplicates(subset=['OID'], keep='first')
    duplicates_removed = initial_count - len(merged_df)

    if duplicates_removed > 0:
        print(f"  - Removed {duplicates_removed} duplicate(s)")

    print(f"  - Final merged rows: {len(merged_df)}")

    # Rename CSV columns to match Excel naming convention (source of truth)
    print("\nStandardizing column names to match Excel (source of truth)...")
    columns_renamed = {k: v for k, v in COLUMN_RENAME_MAP.items() if k in merged_df.columns}
    if columns_renamed:
        merged_df = merged_df.rename(columns=columns_renamed)
        print(f"  - Renamed {len(columns_renamed)} columns:")
        for old, new in columns_renamed.items():
            print(f"    {old} -> {new}")

    # Add Excel-only columns that don't exist in CSV
    print("\nAdding Excel-only columns...")
    added_count = 0
    for col in EXCEL_ONLY_COLUMNS:
        if col not in merged_df.columns:
            merged_df[col] = None
            added_count += 1
            print(f"  - Added: {col}")

    if added_count == 0:
        print("  - No new columns to add")

    return merged_df


def match_component(row, tls_df):
    """
    Match a component from merged data with TLS data.
    Returns matched TLS row or None.
    """
    # Primary match: OID
    if pd.notna(row['OID']):
        oid_match = tls_df[tls_df['OID'] == row['OID']]
        if not oid_match.empty:
            return oid_match.iloc[0]

    # Fallback match: Station Name + Component Description + Additional Info
    station_match = tls_df[tls_df['Station Name'] == row['Station Name']]

    if not station_match.empty:
        desc_match = station_match[
            station_match['Component Description'] == row['Component Description']
        ]

        if not desc_match.empty:
            # Try to match Additional Info if present
            if pd.notna(row['Additional Information']):
                info_match = desc_match[
                    desc_match['Additional Information'] == row['Additional Information']
                ]
                if not info_match.empty:
                    return info_match.iloc[0]

            # Return first description match if no Additional Info match
            return desc_match.iloc[0]

    return None


def compare_rows(merged_row, tls_row):
    """
    Compare two rows and return list of differences.
    Returns list of tuples: (column_name, merged_value, tls_value)
    Note: Columns have already been renamed to match Excel naming convention.
    """
    differences = []

    # Compare all columns that exist in both rows
    for col in merged_row.index:
        if col in ['Mismatch']:  # Skip system columns (but not 'Type of Change' - we want to compare it)
            continue

        merged_val = merged_row[col]
        tls_val = tls_row.get(col) if col in tls_row.index else None

        # Handle NaN comparisons
        merged_is_nan = pd.isna(merged_val)
        tls_is_nan = pd.isna(tls_val)

        if merged_is_nan and tls_is_nan:
            continue
        elif merged_is_nan or tls_is_nan:
            differences.append((col, merged_val, tls_val))
        elif str(merged_val).strip() != str(tls_val).strip():
            differences.append((col, merged_val, tls_val))

    return differences


def add_mismatch_column(merged_df, tls_df):
    """Add Mismatch column to identify discrepancies."""
    print("\nCross-referencing with TLS file...")

    merged_df['Mismatch'] = 'No'
    mismatch_count = 0
    not_in_tls_count = 0

    for idx, row in merged_df.iterrows():
        tls_match = match_component(row, tls_df)

        if tls_match is None:
            # Component not found in TLS
            not_in_tls_count += 1
            continue

        # Compare rows
        differences = compare_rows(row, tls_match)

        if differences:
            merged_df.at[idx, 'Mismatch'] = 'Yes'
            mismatch_count += 1

    print(f"  - Components with mismatches: {mismatch_count}")
    print(f"  - Components not in TLS: {not_in_tls_count}")
    print(f"  - Components matching: {len(merged_df) - mismatch_count - not_in_tls_count}")

    return merged_df


def update_with_tls_data(merged_df, tls_df):
    """Update all entries with TLS data (source of truth) and track changes."""
    print("\nUpdating with TLS data (source of truth)...")

    # Initialize Type of Change if not already present
    if 'Type of Change' not in merged_df.columns:
        merged_df['Type of Change'] = ''

    changes_log = []
    update_count = 0
    not_in_excel_count = 0

    for idx, row in merged_df.iterrows():
        tls_match = match_component(row, tls_df)

        if tls_match is None:
            # Component not found in Excel - retain CSV data as-is
            not_in_excel_count += 1
            continue

        # Find differences and update ALL fields from Excel (source of truth)
        differences = compare_rows(row, tls_match)

        if differences:
            updated_columns = []

            for col, old_val, new_val in differences:
                # Update the value with Excel data
                merged_df.at[idx, col] = new_val
                updated_columns.append(col)

                # Log the change
                changes_log.append({
                    'OID': row['OID'],
                    'Column(s) updated': col,
                    'Old Value': old_val,
                    'New Value': new_val
                })

            # Update Type of Change column
            if updated_columns:
                change_type = f"Updated: {', '.join(updated_columns)}"
                merged_df.at[idx, 'Type of Change'] = change_type
                # Mark as mismatch
                merged_df.at[idx, 'Mismatch'] = 'Yes'
                update_count += 1
        else:
            # No differences - data matches Excel
            merged_df.at[idx, 'Mismatch'] = 'No'

    print(f"  - Updated {update_count} component(s) with Excel data")
    print(f"  - Components not in Excel (retained CSV data): {not_in_excel_count}")
    print(f"  - Total field changes: {len(changes_log)}")

    return merged_df, changes_log


def generate_summary_report(changes_log):
    """Generate summary report of all changes."""
    print("\nGenerating summary report...")

    if changes_log:
        summary_df = pd.DataFrame(changes_log)
        print(f"  - Summary report entries: {len(summary_df)}")
    else:
        # Create empty dataframe with correct columns
        summary_df = pd.DataFrame(columns=['OID', 'Column(s) updated', 'Old Value', 'New Value'])
        print("  - No changes to report")

    return summary_df


def main():
    """Main execution function."""
    print("=" * 60)
    print("Substation Data Merge and Validation")
    print("=" * 60)

    # Create output directory if it doesn't exist
    OUTPUT_DIR.mkdir(exist_ok=True)

    # Step 1: Load files
    sub1_df, sub2_df = load_csv_files()
    tls_df = load_tls_file()

    # Step 2: Merge CSV files and standardize to Excel column names
    merged_df = merge_csv_files(sub1_df, sub2_df)

    # Save merged file (with standardized column names)
    merged_df.to_csv(MERGED_FILE, index=False)
    print(f"\n[OK] Saved: {MERGED_FILE.name}")

    # Step 3: Update with Excel data (source of truth) and identify mismatches
    # This will populate Excel-only columns and mark mismatches
    updated_df, changes_log = update_with_tls_data(merged_df.copy(), tls_df)

    # Save highlighted file (with Mismatch column)
    updated_df.to_csv(HIGHLIGHTED_FILE, index=False)
    print(f"[OK] Saved: {HIGHLIGHTED_FILE.name}")

    # The updated file is the same as highlighted (already contains Excel data)

    # Save updated file
    updated_df.to_csv(UPDATED_FILE, index=False)
    print(f"[OK] Saved: {UPDATED_FILE.name}")

    # Step 5: Generate summary report
    summary_df = generate_summary_report(changes_log)

    # Save summary report
    summary_df.to_csv(SUMMARY_FILE, index=False)
    print(f"[OK] Saved: {SUMMARY_FILE.name}")

    print("\n" + "=" * 60)
    print("Processing complete! All files saved to 'Final' directory.")
    print("=" * 60)


if __name__ == "__main__":
    main()
