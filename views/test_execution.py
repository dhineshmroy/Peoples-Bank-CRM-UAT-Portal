import streamlit as st
import pandas as pd
import io
import psycopg2
from datetime import datetime

def get_db_connection():
    try:
        db_url = st.secrets["postgres"]["url"]
        conn = psycopg2.connect(db_url)
        return conn
    except Exception as e:
        st.error(f"PostgreSQL Connection Failed: {e}")
        return None

def render_test_execution_page():
    st.title("💳 GRG CRM - Test Execution & Cash Loading Management")
    
    tab_exec, tab_cash, tab_export = st.tabs(["📝 Test Execution", "💵 Cash Loading & Receipts", "📊 Finance Export"])

    modules = [
        "GRG_CRM_Cardless_Bill_Payment",
        "GRG_CRM_Cardless_Cash_Deposit",
        "GRG_CRM_Cardbased_Bill_Payment",
        "GRG_CRM_Cardbased_Cash_Deposit",
        "GRG_CRM_Cardbased_Cash_Withdraw",
        "GRG_CRM_Cardbased_Fund_transfer",
        "GRG_CRM_Cardbased_SLIC_Bill_Pay",
        "GRG_CRM_Cardless_SLIC_Bill_Paym",
        "GRG_CRM_Cardbased_Mini_Statemen"
    ]

    # -------------------------------------------------------------------------
    # TAB 1: TEST EXECUTION WITH DYNAMIC FIELDS PER MODULE
    # -------------------------------------------------------------------------
    with tab_exec:
        st.subheader("Execute Pre-Built Test Cases (Dynamic Module Layouts)")
        selected_module = st.selectbox("Select Test Module / Feature", modules, key="exec_mod")

        # Fetch pre-built test cases from Supabase
        conn = get_db_connection()
        tc_list = []
        tc_data_dict = {}
        if conn:
            try:
                tc_df = pd.read_sql(f"SELECT tc_id, test_description, biller_category, expected_remarks FROM uat_test_cases WHERE module_name = '{selected_module}'", conn)
                conn.close()
                if not tc_df.empty:
                    tc_list = tc_df['tc_id'].tolist()
                    for _, row in tc_df.iterrows():
                        tc_data_dict[row['tc_id']] = {
                            "description": row.get('test_description', ''),
                            "biller": row.get('biller_category', ''),
                            "remarks": row.get('expected_remarks', '')
                        }
            except Exception:
                pass

        if not tc_list:
            st.warning("No pre-built test cases found in database for this module.")
            selected_tc = st.text_input("Test Case ID (Manual)", value="TC_001")
            pre_desc, pre_biller, pre_remarks = "", "", ""
        else:
            selected_tc = st.selectbox("Select Test Case ID", tc_list)
            details = tc_data_dict.get(selected_tc, {})
            pre_desc = details.get("description", "")
            pre_biller = details.get("biller", "")
            pre_remarks = details.get("remarks", "")
            st.info(f"**Test Description:** {pre_desc}")

        # --- LIVE INPUTS OUTSIDE FORM FOR REAL-TIME RRN & CALCULATION ---
        col1, col2 = st.columns(2)
        with col1:
            stan = st.text_input("STAN / UTANO", value="260824000228220130", key="stan_input")
            
            # Accurate RRN extraction: e.g., "260824000228220130" -> "000228220130" (taking last 12 digits or slicing index 6)
            auto_rrn = stan[6:] if len(stan) >= 18 else (stan[-12:] if len(stan) >= 12 else stan)
            rrn = st.text_input("RRN (Retrieval Reference Number - Auto)", value=auto_rrn, key="rrn_input")

        form_data = {}

        if selected_module == "GRG_CRM_Cardless_Bill_Payment":
            with col2:
                form_data["biller_name"] = st.text_input("Biller Name / Category", value=pre_biller)
                form_data["consumer_acc"] = st.text_input("Consumer / Acc / Ref No")
            c3, c4 = st.columns(2)
            with c3:
                bill_amount = st.number_input("Bill Txn Amount (LKR)", value=0.00, format="%.2f", key="cardless_bill_amt")
                service_charge = st.number_input("Service Charge (LKR)", value=0.00, format="%.2f", key="cardless_serv_chg")
            with c4:
                actual_paid = bill_amount + service_charge
                form_data["bill_amount"] = bill_amount
                form_data["service_charge"] = service_charge
                form_data["actual_paid"] = actual_paid
                st.markdown(f"### **Actual Paid Txn Amount (LKR) [Auto]:** `LKR {actual_paid:,.2f}`")
                biller_sv_status = st.selectbox("Biller & SV Status", ["UPDATED", "PENDING", "FAILED"])

        elif selected_module == "GRG_CRM_Cardless_Cash_Deposit":
            with col2:
                form_data["account_number"] = st.text_input("Account Number")
            c3, c4 = st.columns(2)
            with c3:
                form_data["before_balance"] = st.number_input("Before Txn Balance (LKR)", value=0.00, format="%.2f")
                form_data["txn_amount"] = st.number_input("Txn Amount (LKR)", value=0.00, format="%.2f")
            with c4:
                form_data["after_balance"] = st.number_input("After Txn Balance (LKR)", value=0.00, format="%.2f")

        elif selected_module == "GRG_CRM_Cardbased_Bill_Payment":
            with col2:
                form_data["bill_number"] = st.text_input("Bill Number")
                form_data["card_number"] = st.text_input("Card Number")
            c3, c4 = st.columns(2)
            with c3:
                form_data["account_number"] = st.text_input("Account Number")
                form_data["card_type"] = st.selectbox("Card Type", ["VISA", "MASTER", "AMEX", "RUPAY"])
                form_data["before_balance"] = st.number_input("Before Txn Balance (LKR)", value=0.00, format="%.2f")
                txn_amt = st.number_input("Txn Amount (LKR)", value=0.00, format="%.2f", key="cb_bill_amt")
            with c4:
                serv_chg = st.number_input("Service Charge (LKR)", value=0.00, format="%.2f", key="cb_bill_serv")
                actual_paid = txn_amt + serv_chg
                form_data["txn_amount"] = txn_amt
                form_data["service_charge"] = serv_chg
                form_data["actual_paid"] = actual_paid
                st.markdown(f"### **Actual Paid Amount (LKR) [Auto]:** `LKR {actual_paid:,.2f}`")
                form_data["after_balance"] = st.number_input("After Txn Balance (LKR)", value=0.00, format="%.2f")

        elif selected_module == "GRG_CRM_Cardbased_Cash_Deposit":
            with col2:
                form_data["account_number"] = st.text_input("Account Number")
                form_data["card_type"] = st.selectbox("Card Type", ["VISA", "MASTER", "AMEX"])
            c3, c4 = st.columns(2)
            with c3:
                form_data["before_balance"] = st.number_input("Before Txn Balance (LKR)", value=0.00, format="%.2f")
                form_data["txn_amount"] = st.number_input("Txn Amount (LKR)", value=0.00, format="%.2f")
            with c4:
                form_data["after_balance"] = st.number_input("After Txn Balance (LKR)", value=0.00, format="%.2f")

        elif selected_module == "GRG_CRM_Cardbased_Cash_Withdraw":
            with col2:
                form_data["card_number"] = st.text_input("Card Number")
                form_data["account_number"] = st.text_input("Account Number")
            c3, c4 = st.columns(2)
            with c3:
                form_data["card_type"] = st.selectbox("Card Type", ["VISA", "MASTER", "AMEX"])
                form_data["before_balance"] = st.number_input("Before Txn Balance (LKR)", value=0.00, format="%.2f")
                txn_amt = st.number_input("Txn Amount (LKR)", value=0.00, format="%.2f", key="cb_wd_amt")
            with c4:
                serv_chg = st.number_input("Service Charge (LKR)", value=0.00, format="%.2f", key="cb_wd_serv")
                actual_txn = txn_amt + serv_chg
                form_data["txn_amount"] = txn_amt
                form_data["service_charge"] = serv_chg
                form_data["actual_txn"] = actual_txn
                st.markdown(f"### **Actual Txn Amount (LKR) [Auto]:** `LKR {actual_txn:,.2f}`")
                form_data["after_balance"] = st.number_input("After Txn Balance (LKR)", value=0.00, format="%.2f")

        elif selected_module == "GRG_CRM_Cardbased_Fund_transfer":
            with col2:
                form_data["card_number"] = st.text_input("Card Number")
                form_data["card_type"] = st.selectbox("Card Type", ["VISA", "MASTER"])
            c3, c4 = st.columns(2)
            with c3:
                form_data["from_acc"] = st.text_input("From Account Number")
                form_data["to_acc"] = st.text_input("To Account Number")
                form_data["before_bal_from"] = st.number_input("Before Txn Balance (LKR) - From Acc", value=0.00, format="%.2f")
                form_data["before_bal_to"] = st.number_input("Before Txn Balance (LKR) - To Acc", value=0.00, format="%.2f")
            with c4:
                form_data["txn_amount"] = st.number_input("Txn Amount (LKR)", value=0.00, format="%.2f")
                form_data["after_bal_from"] = st.number_input("After Txn Balance (LKR) - From Acc", value=0.00, format="%.2f")
                form_data["after_bal_to"] = st.number_input("After Txn Balance (LKR) - To Acc", value=0.00, format="%.2f")

        elif selected_module in ["GRG_CRM_Cardbased_SLIC_Bill_Pay", "GRG_CRM_Cardless_SLIC_Bill_Paym"]:
            with col2:
                form_data["bill_number"] = st.text_input("Bill Number")
                if "Cardbased" in selected_module:
                    form_data["card_number"] = st.text_input("Card Number")
                    form_data["account_number"] = st.text_input("Account Number")
                form_data["card_type"] = st.selectbox("Card Type / Biller Type", ["SLIC", "INSURANCE", "VISA", "MASTER"])
            c3, c4 = st.columns(2)
            with c3:
                txn_amt = st.number_input("Txn Amount (LKR)", value=0.00, format="%.2f", key="slic_amt")
                serv_chg = st.number_input("Service Charge (LKR)", value=0.00, format="%.2f", key="slic_serv")
            with c4:
                actual_paid = txn_amt + serv_chg
                form_data["txn_amount"] = txn_amt
                form_data["service_charge"] = serv_chg
                form_data["actual_paid"] = actual_paid
                st.markdown(f"### **Actual Paid Amount (LKR) [Auto]:** `LKR {actual_paid:,.2f}`")

        elif selected_module == "GRG_CRM_Cardbased_Mini_Statemen":
            with col2:
                form_data["card_number"] = st.text_input("Card Number")
                form_data["account_number"] = st.text_input("Account Number")
            c3, c4 = st.columns(2)
            with c3:
                form_data["card_type"] = st.selectbox("Card Type", ["VISA", "MASTER"])
                form_data["before_balance"] = st.number_input("Before Txn Balance (LKR)", value=0.00, format="%.2f")
            with c4:
                serv_chg = st.number_input("Service Charge (LKR)", value=0.00, format="%.2f", key="mini_serv")
                form_data["service_charge"] = serv_chg
                form_data["actual_txn"] = serv_chg
                st.markdown(f"### **Actual Txn Amount (LKR):** `LKR {serv_chg:,.2f}`")

        # Common Status Controls
        st.markdown("---")
        s_col1, s_col2 = st.columns(2)
        with s_col1:
            fe_status = st.selectbox("Front-End (FE) Status", ["SUCCESS", "FAILED"])
        with s_col2:
            sibs_status = st.selectbox("Core Banking (SIBS) Status", ["UPDATED (SIBS)", "PENDING", "FAILED"])

        overall_status = "PASS" if fe_status == "SUCCESS" and sibs_status == "UPDATED (SIBS)" else "FAIL"
        st.write(f"**Overall Test Status (Auto):** `{overall_status}`")

        remarks = st.text_area("Tester Remarks / Notes", value=pre_remarks)
        
        if st.button("💾 Save Test Execution Record", type="primary"):
            conn = get_db_connection()
            if conn:
                try:
                    cur = conn.cursor()
                    # Included extra_data to match or fallback safely
                    try:
                        cur.execute("""
                            INSERT INTO uat_test_executions 
                            (module_name, tc_id, test_description, rrn, stan_utano, fe_status, sibs_status, overall_status, tester_remarks, executed_by, extra_data)
                            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                        """, (selected_module, selected_tc, pre_desc, rrn, stan, fe_status, sibs_status, overall_status, remarks, st.session_state.get("logged_user", "Tester"), str(form_data)))
                    except Exception:
                        conn.rollback()
                        cur.execute("""
                            INSERT INTO uat_test_executions 
                            (module_name, tc_id, test_description, rrn, stan_utano, fe_status, sibs_status, overall_status, tester_remarks, executed_by)
                            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                        """, (selected_module, selected_tc, pre_desc, rrn, stan, fe_status, sibs_status, overall_status, remarks, st.session_state.get("logged_user", "Tester")))
                    
                    conn.commit()
                    cur.close()
                    conn.close()
                    st.success(f"Successfully recorded execution for **{selected_tc}** under **{selected_module}**!")
                except Exception as e:
                    st.error(f"Error saving execution record: {e}")

    # -------------------------------------------------------------------------
    # TAB 2: CASH LOADING & UNLOADING MANAGEMENT
    # -------------------------------------------------------------------------
    with tab_cash:
        st.subheader("💵 Terminal Cash Loading & Unloading Tracker")
        with st.form("cash_loading_form"):
            c_col1, c_col2 = st.columns(2)
            with c_col1:
                terminal_id = st.text_input("Terminal ID", value="169RB02")
                report_date = st.date_input("Report Date", value=datetime.today())
                loading_session = st.selectbox("Loading Session", ["1st Cash Loading", "2nd Cash Loading", "3rd Cash Loading", "Unloading Session"])
            with c_col2:
                load_time = st.time_input("Loading / Action Time")
                loading_total = st.number_input("Loading / Session Total (LKR)", value=0.00, min_value=0.00)

            sop_file = st.file_uploader("Upload SOP Receipt (.pdf, .png, .jpg)", type=["pdf", "png", "jpg"], key="sop")
            host_file = st.file_uploader("Upload HOST Receipt (.pdf, .png, .jpg)", type=["pdf", "png", "jpg"], key="host")

            if st.form_submit_button("💾 Save Cash Loading Entry"):
                sop_name = sop_file.name if sop_file else "None"
                host_name = host_file.name if host_file else "None"
                conn = get_db_connection()
                if conn:
                    try:
                        cur = conn.cursor()
                        cur.execute("""
                            INSERT INTO terminal_cash_logs 
                            (terminal_id, report_date, loading_session, load_time, loading_total, sop_receipt_path, host_receipt_path, logged_by)
                            VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
                        """, (terminal_id, report_date, loading_session, str(load_time), loading_total, sop_name, host_name, st.session_state.get("logged_user", "Tester")))
                        conn.commit()
                        cur.close()
                        conn.close()
                        st.success(f"Successfully recorded **{loading_session}** of LKR {loading_total:,.2f}!")
                    except Exception as e:
                        st.error(f"Error saving cash log: {e}")

    # -------------------------------------------------------------------------
    # TAB 3: FINANCE EXPORT
    # -------------------------------------------------------------------------
    with tab_export:
        st.subheader("📊 Export Multi-Tab Report for Finance Department")
        if st.button("📥 Generate & Download Finance Excel Workbook"):
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                conn = get_db_connection()
                if conn:
                    try:
                        for mod_name in modules:
                            df_mod = pd.read_sql(f"SELECT tc_id, test_description, rrn, stan_utano, fe_status, sibs_status, overall_status, tester_remarks FROM uat_test_executions WHERE module_name = '{mod_name}'", conn)
                            if df_mod.empty:
                                df_mod = pd.DataFrame(columns=["TC_ID", "Test Description", "RRN", "STAN", "FE Status", "SIBS Status", "Overall Status", "Remarks"])
                            df_mod.to_excel(writer, sheet_name=mod_name[:31], index=False)
                        conn.close()
                    except Exception:
                        for mod_name in modules:
                            pd.DataFrame().to_excel(writer, sheet_name=mod_name[:31], index=False)
            output.seek(0)
            st.download_button(
                label="⬇️ Download Completed Finance Workbook (.xlsx)",
                data=output,
                file_name=f"PeoplesBank_CRM_Finance_Report_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

if __name__ == "__main__":
    render_test_execution_page()