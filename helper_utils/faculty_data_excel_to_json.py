import pandas as pd
import json
import os

# 🔧 ====== CONFIGURABLE PATHS ======
EXCEL_PATH = r"C:\Users\Lenovo\Downloads\faculty_data.xlsx"   # <-- change this to your Excel file path
OUTPUT_JSON_PATH = r"E:\attendence_filter\flask_app\JSON\\faculty_detail.json"  # <-- change output location
# ==================================

def process_faculty_excel(excel_path):
    # Read both sheets
    sheets = pd.read_excel(excel_path, sheet_name=None)
    all_data = {}

    for sheet_name, df in sheets.items():
        # Normalize column names
        df.columns = [str(col).strip().lower().replace(" ", "_") for col in df.columns]
        
        for _, row in df.iterrows():
            faculty_id = str(row.get("faculty_id", "")).strip()
            if not faculty_id:
                continue

            all_data[faculty_id] = {
                "name": str(row.get("name", "")).strip(),
                "designation": str(row.get("designation", "")).strip(),
                "department": str(row.get("department", "")).strip(),
                "category": sheet_name.strip()
            }
    
    return all_data

def save_json(data, output_path):
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=4, ensure_ascii=False)
    print(f"✅ JSON file saved at: {output_path}")

if __name__ == "__main__":
    faculty_data = process_faculty_excel(EXCEL_PATH)
    save_json(faculty_data, OUTPUT_JSON_PATH)
