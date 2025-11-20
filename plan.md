# Plan for Solving Programming Scope Assignment

## Assignment Overview
**Objective**: Merge two substation CSV files (SUB1 and SUB2), cross-reference with a TLS (Transmission Line Segment) Excel file, identify mismatches, update data, and generate reports.

## Input Files
1. **SUB1.csv** - Substation 1 component data (15 rows, 115 kV components)
2. **SUB2.csv** - Substation 2 component data (51 rows, multiple voltage levels: 60kV, 115kV, 230kV)
3. **SUB1-SUB2 115 kV -XcelUpdate.xlsx** - TLS reference data (CAISO Update sheet, 48 rows)

## Solution Implementation

### Step 1: Load and Validate Input Files
- Load both CSV files using `pandas.read_csv()`
- Load Excel file using `pandas.read_excel()` with `openpyxl` engine
- Use the **CAISO Update** sheet from the Excel file (contains component data)
- Validate file formats and column structures
- Handle missing data with `pd.notna()` and `pd.isna()`

### Step 2: Merge CSV Files and Standardize Column Names
- Use `pd.concat()` to combine SUB1 and SUB2 dataframes vertically
- Remove duplicates using `drop_duplicates(subset=['OID'], keep='first')`
- **Rename CSV columns to match Excel naming convention** (Excel is source of truth):
  - `AMP Rating` → `High Rating`
  - `High KV` → `High kV` (and similar for Low/Tertiary)
- **Add Excel-only columns** to the merged dataframe:
  - `Low Rating`, `Low Rating.1`, `Low Rating.2`, `Low Rating.3`
  - `Line Number`
  - `Type of Change`
- **Retain CSV-only columns** that don't exist in Excel:
  - `Group`, `T-Line Name`, `High Rating.4`, etc.
- **Output**: `Final/SUB1-SUB2 115kV.csv` (with standardized column names)

### Step 3: Update with Excel Data (Source of Truth)
- For each component in merged data:
  - Find matching component in Excel using matching strategy (see below)
  - If match found: Compare all fields and update with Excel values
  - If no match: Retain CSV data as-is
- Populate Excel-only columns (Low Rating, Line Number, etc.) from matched rows
- Add `Mismatch` column: "Yes" if any field was updated, "No" if data matches
- Add `Type of Change` column: List which columns were updated
- **Output**:
  - `Final/SUB1-SUB2 115kV_highlighted.csv` (with Mismatch column)
  - `Final/SUB1-SUB2 115kV_updated.csv` (same as highlighted, fully updated with Excel data)

**Matching Strategy** (hierarchical):
1. **Primary Match**: Match by OID column
   - If `OID` is not NaN and exists in both files
2. **Fallback Match** (if OID missing or no match):
   - Match by `Station Name` + `Component Description`
   - If `Additional Information` is present, use it for additional matching

**Comparison Logic**:
- Since columns are now standardized to Excel naming, direct column comparison is possible
- Compare all columns that exist in both datasets
- Handle NaN values properly (two NaN values are considered equal)
- String values are compared after trimming whitespace

### Step 4: Generate Summary Report
Create a detailed change log with:
- **OID**: Component identifier
- **Column(s) updated**: Which field was changed
- **Old Value**: Original value from CSV
- **New Value**: Updated value from Excel (source of truth)

**Output**: `Final/SUB1-SUB2 115kV_summary_report.csv`

## Final Deliverables

All files saved to `Final/` directory:

1. ✅ **SUB1-SUB2 115kV.csv** - Combined data from both substations with:
   - Standardized column names (Excel naming convention)
   - Excel-only columns added (Low Rating, Line Number, etc.)
   - CSV-only columns retained (Group, T-Line Name, etc.)
2. ✅ **SUB1-SUB2 115kV_highlighted.csv** - Same as updated file (with Mismatch and Type of Change columns)
3. ✅ **SUB1-SUB2 115kV_updated.csv** - Final file updated with Excel data (source of truth):
   - All matching components updated with Excel values
   - Excel-only columns populated from source
   - Non-matching components retain CSV data
   - Mismatch and Type of Change columns added
4. ✅ **SUB1-SUB2 115kV_summary_report.csv** - Detailed change log (291 field updates across 23 components)

## Implementation Details

### Technologies Used
- **Python 3.12**
- **pandas** - Data manipulation and CSV/Excel handling
- **openpyxl** - Excel file reading engine

### Key Functions

1. **load_csv_files()** - Load both substation CSV files
2. **load_tls_file()** - Load Excel file (CAISO Update sheet - source of truth)
3. **merge_csv_files()** - Combine, deduplicate, rename columns to Excel convention, and add Excel-only columns
4. **match_component()** - Find matching component in Excel using OID or fallback criteria
5. **compare_rows()** - Compare two components and identify differences (after column standardization)
6. **update_with_tls_data()** - Update all matched entries with Excel data and track changes
7. **generate_summary_report()** - Create change log DataFrame

### Data Flow
```
SUB1.csv  ──┐
            ├──> merge_csv_files() ──> merged.csv ──────────┐
SUB2.csv  ──┘    (standardize columns)                      │
                 (add Excel-only cols)                      │
                                                            ↓
Excel.xlsx ──────────────────────────────────> update_with_tls_data()
(source of truth)                                           ↓
                                             ┌──────────────┴────────────┐
                                             ↓                           ↓
                                    highlighted.csv              summary_report.csv
                                    updated.csv
                                    (same file, both outputs)
```

### Handling Edge Cases
1. **Missing OID**: Use fallback matching by Station Name + Component Description
2. **NaN Values**: Properly handle NaN comparisons (two NaN values are considered equal)
3. **Column Name Mismatches**: Rename CSV columns to match Excel naming convention (source of truth)
4. **Excel-only Columns**: Add columns that exist in Excel but not in CSV (Low Rating, Line Number, etc.)
5. **CSV-only Columns**: Retain columns that exist in CSV but not in Excel (Group, T-Line Name, High Rating.4, etc.)
6. **Components Not in Excel**: Retain CSV data as-is for components not found in Excel
7. **Duplicate OIDs**: Keep first occurrence during merge

## Testing Results

**Execution Summary**:
- SUB1.csv: 15 rows loaded
- SUB2.csv: 51 rows loaded
- Excel file: 48 rows loaded (source of truth)
- **Merged**: 66 total components (no duplicates found)
- **Column standardization**: 8 columns renamed to Excel convention
- **Excel-only columns added**: 6 columns (Low Rating × 4, Line Number, Type of Change)
- **Components updated with Excel data**: 23 components
- **Components not in Excel** (CSV data retained): 43 components
- **Total field changes**: 291 changes across 23 components

## How to Run

```bash
python merge_substation_data.py
```

The script will:
1. Create the `Final/` directory if it doesn't exist
2. Load and merge CSV files
3. Standardize column names to Excel convention (source of truth)
4. Add Excel-only columns to merged data
5. Update all matched components with Excel data
6. Retain CSV data for components not in Excel
7. Display progress and statistics
8. Generate all 4 required output files

## Notes and Considerations

1. **Column Standardization to Excel Convention**: The Excel file is the source of truth, so the script standardizes all column names to match Excel naming:

   **Rating Columns (renamed from CSV):**
   - CSV `AMP Rating` → Excel `High Rating` ✓
   - CSV `AMP Rating.1` → Excel `High Rating.1` ✓
   - CSV `AMP Rating.2` → Excel `High Rating.2` ✓
   - CSV `AMP Rating.3` → Excel `High Rating.3` ✓
   - CSV `AMP Rating.4` → `High Rating.4` (retained, no Excel equivalent)

   **Voltage Columns (renamed from CSV):**
   - CSV `High KV` → Excel `High kV` ✓
   - CSV `Low KV` → Excel `Low kV` ✓
   - CSV `Tertiary KV` → Excel `Tertiary kV` ✓

   **Excel-only Columns (added to merged data):**
   - `Low Rating`, `Low Rating.1`, `Low Rating.2`, `Low Rating.3`
   - `Line Number`
   - `Type of Change`

   **CSV-only Columns (retained in merged data):**
   - `Group`, `T-Line Name`, `High Rating.4`, `MVA Rating`, `MVAr High`, `MVAr Low`, `Con`, etc.

   **Rating Type Pattern:**
   - Both files have 4 rating types with the pattern: Rating Type, High/AMP Rating, [Low Rating], Duration, Note #
   - Rating Type: `SN (N)` - Summer Normal
   - Rating Type.1: `SE (A)` - Summer Emergency
   - Rating Type.2: `WN (B)` - Winter Normal
   - Rating Type.3: `WE (C)` - Winter Emergency
   - CSV files may have a 5th rating type (Rating Type.4) that doesn't exist in Excel

2. **Data Types**: Some fields convert from integers to floats during the update process (e.g., OID: 163125 → 163125.0, AMP Rating: 2000 → 2000.0). This is normal pandas behavior.

3. **Voltage Level**: The script is designed for 115 kV data specifically, as indicated by the TLS filename.

4. **Missing Data Handling**: Empty fields in the TLS file are treated as NaN and properly compared against existing data.

5. **Case Sensitivity**: String comparisons are case-sensitive. Consider adding `.str.lower()` if case-insensitive matching is needed.

## Potential Improvements

1. **Filter by Voltage Level**: Only compare components at matching voltage levels (115 kV with 115 kV)
2. **Smart Column Mapping**: Handle column name differences between CSV and Excel files
3. **Validation Rules**: Add data type validation and range checks for numeric fields
4. **Detailed Logging**: Add logging to file for debugging and audit trail
5. **Excel Output**: Generate Excel files with formatting for easier review
6. **Configuration File**: Externalize file paths and matching rules to a config file

## Conclusion

The solution successfully implements all requirements with Excel as the source of truth:
- ✅ Loads and merges CSV files
- ✅ Standardizes column names to Excel convention (source of truth)
- ✅ Adds Excel-only columns to merged data (Low Rating, Line Number, etc.)
- ✅ Retains CSV-only columns not in Excel (Group, T-Line Name, etc.)
- ✅ Cross-references with Excel file and updates all matched components
- ✅ Identifies and highlights mismatches
- ✅ Updates data from Excel source (source of truth)
- ✅ Retains CSV data for components not found in Excel
- ✅ Generates comprehensive summary report
- ✅ Handles missing data carefully
- ✅ Ensures no duplicates after merging
- ✅ Validates file formats before processing

All output files are ready for review in the `Final/` directory.
