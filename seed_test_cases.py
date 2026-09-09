import pandas as pd
import psycopg2

db_url = "postgresql://postgres.mxqtzquofvbcewtaowxo:Dhineshmelroy@aws-0-ap-southeast-2.pooler.supabase.com:5432/postgres?sslmode=require"

try:
    conn = psycopg2.connect(db_url)
    cursor = conn.cursor()
    print("Successfully connected to Supabase!")

    excel_path = "GRG_CRM_CardLess_Bill_Payment_Test_Execution_Sheet.xlsx"
    xls = pd.ExcelFile(excel_path)
    
    total_inserted = 0
    for sheet_name in xls.sheet_names:
        if sheet_name == "Export Summary":
            continue  
            
        df_raw = pd.read_excel(xls, sheet_name=sheet_name, header=None)
        
        # Find header row index
        header_row_idx = 0
        for idx, row in df_raw.iterrows():
            row_str = str(row.values).lower()
            if 'tc_id' in row_str or 'tc id' in row_str or 'test case' in row_str:
                header_row_idx = idx
                break

        df = pd.read_excel(xls, sheet_name=sheet_name, header=header_row_idx)
        print(f"\nProcessing sheet: '{sheet_name}' (Detected header at row {header_row_idx})")

        for _, row in df.iterrows():
            row_dict = row.to_dict()
            
            # Helper to safely clean and fetch values by partial column name matching
            def get_val(*keywords):
                for col, val in row_dict.items():
                    col_lower = str(col).lower()
                    if any(kw in col_lower for kw in keywords):
                        if pd.notna(val):
                            return str(val).strip()
                return None

            tc_id = get_val("tc_id", "tc id", "test case id", "tc")
            if not tc_id or not tc_id.startswith("TC"):
                continue

            desc = get_val("description") or ""
            biller = get_val("biller", "category") or ""
            consumer = get_val("consumer", "acc", "ref", "account number") or ""
            card_no = get_val("card number") or ""
            card_type = get_val("card type") or ""
            from_acc = get_val("from account") or ""
            to_acc = get_val("to account") or ""
            bill_no = get_val("bill number") or ""
            remarks = get_val("remark", "expected", "note") or ""

            cursor.execute("""
                INSERT INTO uat_test_cases 
                (module_name, tc_id, test_description, biller_category, consumer_acc_ref, 
                 card_number, card_type, from_account, to_account, bill_number, expected_remarks)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                ON CONFLICT (tc_id) DO UPDATE 
                SET test_description = EXCLUDED.test_description,
                    biller_category = EXCLUDED.biller_category,
                    consumer_acc_ref = EXCLUDED.consumer_acc_ref,
                    card_number = EXCLUDED.card_number,
                    card_type = EXCLUDED.card_type,
                    from_account = EXCLUDED.from_account,
                    to_account = EXCLUDED.to_account,
                    bill_number = EXCLUDED.bill_number,
                    expected_remarks = EXCLUDED.expected_remarks;
            """, (sheet_name, tc_id, desc, biller, consumer, card_no, card_type, from_acc, to_acc, bill_no, remarks))
            total_inserted += 1

    conn.commit()
    cursor.close()
    conn.close()
    print(f"\nSuccessfully inserted/updated {total_inserted} test cases with full columns into Supabase!")

except Exception as e:
    print(f"Upload failed: {e}")