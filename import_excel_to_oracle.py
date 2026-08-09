import pandas as pd
import oracledb
import os

excel_sources = {
    "Card_TestCases.xlsx": "Card Based",
    "Cardless_TestCases.xlsx": "Cardless"
}

conn = oracledb.connect(
    user="SYSTEM",
    password="Dhinesh@98",
    dsn="localhost:1521/FREEPDB1"
)
cursor = conn.cursor()

print("Setting up Oracle database schema with all required defect tracking columns...")
try:
    cursor.execute("DROP TABLE uat_test_cases_v2")
except:
    pass

cursor.execute("""
    CREATE TABLE uat_test_cases_v2 (
        tc_id VARCHAR2(50),
        category VARCHAR2(50),
        module_name VARCHAR2(150),
        test_area VARCHAR2(150),
        test_case_description VARCHAR2(500),
        pre_conditions VARCHAR2(1000),
        test_steps VARCHAR2(1000),
        expected_result VARCHAR2(1000),
        path_type VARCHAR2(20),
        actual_result VARCHAR2(1000),
        rrn VARCHAR2(100),
        utano VARCHAR2(100),
        status VARCHAR2(20) DEFAULT 'PENDING',
        fe VARCHAR2(50),
        sibs VARCHAR2(50),
        remarks VARCHAR2(500),
        executed_by VARCHAR2(100),
        executed_date VARCHAR2(20),
        receipt_path VARCHAR2(255),
        photo_path VARCHAR2(255),
        severity VARCHAR2(50) DEFAULT 'Medium',
        priority VARCHAR2(50) DEFAULT 'Medium',
        defect_status VARCHAR2(50) DEFAULT 'Open',
        assigned_to VARCHAR2(100) DEFAULT 'Development Team',
        target_date VARCHAR2(20),
        root_cause VARCHAR2(1000),
        defect_description VARCHAR2(1000)
    )
""")

try:
    cursor.execute("""
        CREATE TABLE crm_machine_status (
            lock_id NUMBER PRIMARY KEY,
            crm_status VARCHAR2(20),
            locked_by VARCHAR2(100),
            last_updated TIMESTAMP
        )
    """)
    cursor.execute("INSERT INTO crm_machine_status VALUES (1, 'AVAILABLE', '', CURRENT_TIMESTAMP)")
    conn.commit()
except:
    pass

def derive_rrn_from_utano(utano_val):
    u_str = str(utano_val).strip()
    if u_str.endswith(".0"):
        u_str = u_str[:-2]
    if len(u_str) > 6 and u_str.lower() not in ["nan", "none", ""]:
        return u_str[6:]
    return ""

total_inserted = 0

for filename, category_name in excel_sources.items():
    if not os.path.exists(filename):
        print(f"⚠️ File '{filename}' not found. Skipping...")
        continue
        
    print(f"\nOpening workbook: {filename} (Category: {category_name})...")
    xls = pd.ExcelFile(filename)
    
    for sheet_name in xls.sheet_names:
        if sheet_name in ["Export Summary", "Summary"]:
            continue
        
        df = pd.read_excel(filename, sheet_name=sheet_name, header=1, dtype=str)
        if df.empty:
            continue
            
        print(f"Processing sheet: {sheet_name}...")
        path_type = "Negative" if "negative" in sheet_name.lower() or "negetive" in sheet_name.lower() or "-neg" in sheet_name.lower() else "Positive"
        
        for index, row in df.iterrows():
            raw_id = row.iloc[0] if len(row) > 0 and pd.notna(row.iloc[0]) else ""
            tc_id = str(raw_id).strip()
            
            if not tc_id or tc_id.lower() == "nan" or not (tc_id.startswith("TC_") or tc_id.startswith("CD-") or tc_id.startswith("CL-") or "-" in tc_id):
                continue
                
            test_area = str(row.iloc[1]).strip() if len(row) > 1 and pd.notna(row.iloc[1]) and str(row.iloc[1]) != "nan" else ""
            description = str(row.iloc[2]).strip() if len(row) > 2 and pd.notna(row.iloc[2]) and str(row.iloc[2]) != "nan" else ""
            pre_cond = str(row.iloc[3]).strip() if len(row) > 3 and pd.notna(row.iloc[3]) and str(row.iloc[3]) != "nan" else ""
            steps = str(row.iloc[4]).strip() if len(row) > 4 and pd.notna(row.iloc[4]) and str(row.iloc[4]) != "nan" else ""
            expected = str(row.iloc[5]).strip() if len(row) > 5 and pd.notna(row.iloc[5]) and str(row.iloc[5]) != "nan" else ""
            
            actual = str(row.iloc[6]).strip() if len(row) > 6 and pd.notna(row.iloc[6]) and str(row.iloc[6]) != "nan" else ""
            rrn = str(row.iloc[7]).strip() if len(row) > 7 and pd.notna(row.iloc[7]) and str(row.iloc[7]) != "nan" else ""
            if rrn.endswith(".0"): rrn = rrn[:-2]
            
            utano = str(row.iloc[8]).strip() if len(row) > 8 and pd.notna(row.iloc[8]) and str(row.iloc[8]) != "nan" else ""
            if utano.endswith(".0"): utano = utano[:-2]
            
            if utano and (not rrn or rrn.lower() == "nan" or rrn == ""):
                rrn = derive_rrn_from_utano(utano)

            status = str(row.iloc[9]).strip() if len(row) > 9 and pd.notna(row.iloc[9]) and str(row.iloc[9]) != "nan" else "PENDING"
            if status.lower() == "nan" or not status:
                status = "PENDING"
                
            fe = str(row.iloc[10]).strip() if len(row) > 10 and pd.notna(row.iloc[10]) and str(row.iloc[10]) != "nan" else ""
            
            sibs = ""
            remarks = ""
            if len(row) >= 13:
                sibs = str(row.iloc[11]).strip() if pd.notna(row.iloc[11]) and str(row.iloc[11]) != "nan" else ""
                remarks = str(row.iloc[12]).strip() if pd.notna(row.iloc[12]) and str(row.iloc[12]) != "nan" else ""
            elif len(row) == 12:
                remarks = str(row.iloc[11]).strip() if pd.notna(row.iloc[11]) and str(row.iloc[11]) != "nan" else ""

            cursor.execute("""
                INSERT INTO uat_test_cases_v2 (
                    tc_id, category, module_name, test_area, test_case_description, pre_conditions, test_steps, 
                    expected_result, path_type, actual_result, rrn, utano, status, fe, sibs, remarks, 
                    executed_by, executed_date, receipt_path, photo_path,
                    severity, priority, defect_status, assigned_to, target_date, root_cause, defect_description
                ) VALUES (:1, :2, :3, :4, :5, :6, :7, :8, :9, :10, :11, :12, :13, :14, :15, :16, :17, :18, :19, :20, :21, :22, :23, :24, :25, :26, :27)
            """, [
                tc_id, category_name, sheet_name, test_area, description, pre_cond, steps, 
                expected, path_type, actual, rrn, utano, status, fe, sibs, remarks, 
                "", "", "", "",
                'Medium', 'Medium', 'Open', 'Development Team', '', '', ''
            ])
            total_inserted += 1

conn.commit()
cursor.close()
conn.close()
print(f"\nSuccessfully imported {total_inserted} test cases into Oracle DB with all defect columns!")