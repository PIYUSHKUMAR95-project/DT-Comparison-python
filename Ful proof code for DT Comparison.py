import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from datetime import datetime

def compare_excels(file1_path, file2_path):
    try:
        # Load first sheet names
        sheet1 = pd.ExcelFile(file1_path).sheet_names[0]
        sheet2 = pd.ExcelFile(file2_path).sheet_names[0]

        # Read File 1 normally
        df1 = pd.read_excel(file1_path, sheet_name=sheet1)

        # Read File 2 with header starting from A19 (zero-indexed row 18)
        df2 = pd.read_excel(file2_path, sheet_name=sheet2, header=18)

        # Column mapping
        column_mapping = {
            'IA_Code': ['IA_Code_1', 'IA_CODE1_DESC_NEW'],
            'Quantity': ['QUANTITY', 'FINAL_LETTERSHOP_QTY'],
            'PRIMARY_SOURCE_CODE': ['PRIMARY_SOURCE_CODE', 'PRIMARY_SOURCE_CODE'],
            'PRIMARY_SPID': ['PRIMARY_SPID', 'PRIMARY_SPID1_NEW'],
            'CAMPAIGN_CODE': ['CAMPAIGN_CODE', 'CAMPAIGN_CODE'],
            'TEMPLATE_CODE': ['TEMPLATE_CODE', 'TEMPLATE_CODE'],
            'EXPIRATION_DATE': ['EXPIRATION_DATE', 'EXPIRATION_DATE'],
            'PRESCREEN_DATE': ['PRESCREEN_DATE', 'PRESCREEN_DATE'],
            'POID': ['POID', 'POID'],
            'CELL_ID': ['CELL_ID', 'CELL_ID']
        }

        # Rename based on mapping
        file1_rename = {}
        file2_rename = {}
        for std, (f1_col, f2_col) in column_mapping.items():
            if f1_col in df1.columns:
                file1_rename[f1_col] = std
            if f2_col in df2.columns:
                file2_rename[f2_col] = std
        df1.rename(columns=file1_rename, inplace=True)
        df2.rename(columns=file2_rename, inplace=True)

        # Determine common columns
        common_cols = list(set(file1_rename.values()) & set(file2_rename.values()))
        if not common_cols:
            common_cols = df1.columns.intersection(df2.columns).tolist()

        if not common_cols:
            raise ValueError("No common columns to compare.")

        # Sort by CELL_ID if available
        for df in [df1, df2]:
            if 'CELL_ID' in df.columns:
                df['CELL_ID'] = df['CELL_ID'].astype(str).str.strip()
                df.sort_values(by='CELL_ID', inplace=True)
                df.reset_index(drop=True, inplace=True)

        # Normalize date fields
        def normalize_date(val):
            if pd.isnull(val): return None
            val = str(val).strip()
            if val.isdigit() and 7 <= len(val) <= 8:
                try: return datetime.strptime(val.zfill(8), "%m%d%Y").date()
                except: pass
            for fmt in ("%m/%d/%Y", "%Y%m%d", "%m%d%Y"):
                try: return datetime.strptime(val, fmt).date()
                except: continue
            return val

        for col in common_cols:
            if "date" in col.lower():
                df1[col] = df1[col].apply(normalize_date)
                df2[col] = df2[col].apply(normalize_date)

        # Align rows
        min_len = min(len(df1), len(df2))
        df1 = df1.iloc[:min_len].reset_index(drop=True)
        df2 = df2.iloc[:min_len].reset_index(drop=True)

        # Comparison
        result_df = df1.copy()

        if 'Quantity' in df1.columns and 'Quantity' in df2.columns:
            diff_series = pd.to_numeric(df1['Quantity'], errors='coerce') - pd.to_numeric(df2['Quantity'], errors='coerce')
            result_df.insert(result_df.columns.get_loc('Quantity') + 1, 'Quantity_Diff', diff_series)

        for col in common_cols:
            result_df[col] = [
                v1 if str(v1) == str(v2) else f'DIFF: {v1} | {v2}'
                for v1, v2 in zip(df1[col].astype(str), df2[col].astype(str))
            ]

        # Save result to file1
        book = load_workbook(file1_path)
        with pd.ExcelWriter(file1_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            writer._book = book
            result_df.to_excel(writer, sheet_name='Comparison_Result', index=False)

        # Apply red font to differences
        wb = load_workbook(file1_path)
        ws = wb['Comparison_Result']
        red_font = Font(color="FF0000")
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                if isinstance(cell.value, str) and cell.value.startswith("DIFF:"):
                    cell.font = red_font

        # Auto-fit columns
        for col_idx, col in enumerate(ws.iter_cols(min_row=1, max_row=ws.max_row), start=1):
            max_len = max((len(str(cell.value)) for cell in col if cell.value), default=0)
            ws.column_dimensions[get_column_letter(col_idx)].width = max_len + 2

        wb.save(file1_path)
        print("✅ Comparison saved to 'Comparison_Result' in:", file1_path)
        return file1_path

    except Exception as e:
        print("❌ Error:", str(e))
        raise

# Replace with your file names
result_path = compare_excels("Data Tab Report.xlsx", "Platinum_Mail Plan.xlsx")
