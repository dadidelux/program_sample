"""
Test script to verify column mapping is working correctly
"""
import pandas as pd
from pathlib import Path

# File paths
BASE_DIR = Path(__file__).parent
DATASETS_DIR = BASE_DIR / "Datasets"

SUB1_FILE = DATASETS_DIR / "SUB1.csv"
TLS_FILE = DATASETS_DIR / "SUB1-SUB2 115 kV -XcelUpdate.xlsx"

# Column mapping
COLUMN_MAPPING = {
    'AMP Rating': 'High Rating',
    'AMP Rating.1': 'High Rating.1',
    'AMP Rating.2': 'High Rating.2',
    'AMP Rating.3': 'High Rating.3',
    'High KV': 'High kV',
    'Low KV': 'Low kV',
    'Tertiary KV': 'Tertiary kV',
}

# Load files
print("Loading files...")
sub1_df = pd.read_csv(SUB1_FILE)
tls_df = pd.read_excel(TLS_FILE, sheet_name='CAISO Update', engine='openpyxl')

# Find a matching row by OID
test_oid = sub1_df.iloc[1]['OID']  # Get second row
print(f"\nTesting with OID: {test_oid}")

csv_row = sub1_df[sub1_df['OID'] == test_oid].iloc[0]
tls_row = tls_df[tls_df['OID'] == test_oid].iloc[0]

print("\n" + "="*80)
print("COMPARISON TEST - Mapped Columns")
print("="*80)

for csv_col, excel_col in COLUMN_MAPPING.items():
    csv_val = csv_row.get(csv_col)
    excel_val = tls_row.get(excel_col)

    match = "MATCH" if str(csv_val) == str(excel_val) else "DIFF"

    print(f"\n{match} {csv_col:20s} -> {excel_col:20s}")
    print(f"  CSV:   {csv_val}")
    print(f"  Excel: {excel_val}")

print("\n" + "="*80)
print("Excel has additional 'Low Rating' columns not in CSV:")
print("="*80)
for i in ['', '.1', '.2', '.3']:
    col_name = f'Low Rating{i}'
    if col_name in tls_row.index:
        print(f"{col_name:20s}: {tls_row[col_name]}")
