import pandas as pd
import os

data_dir = r'c:\Users\dadidelux\Documents\Programs\program_sample\Datasets'
sub1_path = os.path.join(data_dir, 'SUB1.csv')
sub2_path = os.path.join(data_dir, 'SUB2.csv')
tls_path = os.path.join(data_dir, 'SUB1-SUB2 115 kV -XcelUpdate.xlsx')

print("--- SUB1.csv Columns ---")
try:
    df_sub1 = pd.read_csv(sub1_path)
    print(df_sub1.columns.tolist())
except Exception as e:
    print(e)

print("\n--- SUB2.csv Columns ---")
try:
    df_sub2 = pd.read_csv(sub2_path)
    print(df_sub2.columns.tolist())
except Exception as e:
    print(e)

print("\n--- CAISO Update Sheet Columns ---")
try:
    df_tls = pd.read_excel(tls_path, sheet_name='CAISO Update', engine='openpyxl')
    print(df_tls.columns.tolist())
except Exception as e:
    print(e)
