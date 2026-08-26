import streamlit as st
import pandas as pd
import io
import psycopg2
from datetime import datetime
from config import get_db_url

def render_test_execution_page():
    st.title("💳 GRG CRM - Automated Test Execution & Finance Export")
    st.markdown("""
    Execute positive test cases, record live transaction identifiers (RRN, STAN/UTANO, Balances), 
    and export multi-tab execution sheets for **Finance Department Confirmation**.
    """)

    # Module selection matching your Excel tabs
    modules = {
        "GRG_CRM_Cardless_Bill_Payment": [
            "TC_ID", "Test Description", "Biller Name / Category", "Consumer / Acc / Ref No", 
            "RRN (Retrieval Ref)", "STAN / UTANO", "Bill Txn Amount (LKR)", "Service Charge", 
            "Actual Paid Txn Amount (LKR)", "FE Status", "Biller & SV Status", 
            "Core Banking (SIBS) Status", "Overall Test Status", "Expected Result / Remarks"
        ],
        "GRG_CRM_Cardless_Cash_Deposit": [
            "TC_ID", "Test Description", "Account Number", "RRN (Retrieval Ref)", "STAN / UTANO", 
            "Before Txn Balance (LKR)", "Txn Amount (LKR)", "After Txn Balance (LKR)", 
            "Front-End (FE) Status", "Core Banking (SIBS) Status", "Overall Test Status", "Tester Remarks / Notes"
        ],
        "GRG_CRM_Cardbased_Bill_Payment": [
            "TC_ID", "Test Description", "Bill Number", "Card Number", "Account Number", "Card Type", 
            "RRN (Retrieval Ref)", "STAN / UTANO", "Before Txn Balance (LKR)", "Txn Amount (LKR)", 
            "Service Charge", "Actual paid amount (LKR)", "After Txn Balance (LKR)", 
            "Front-End (FE) Status", "Core Banking (SIBS) Status", "Overall Test Status", "Tester Remarks / Notes"
        ],
        "GRG_CRM_Cardbased_Cash_Deposit": [
            "TC_ID", "Test Description", "Account Number", "Card Type", "RRN (Retrieval Ref)", "STAN / UTANO", 
            "Before Txn Balance (LKR)", "Txn Amount (LKR)", "After Txn Balance (LKR)", 
            "Front-End (FE) Status", "Core Banking (SIBS) Status", "Overall Test Status", "Tester Remarks / Notes"
        ],
        "GRG_CRM_Cardbased_Cash_Withdrawal": [
            "TC_ID", "Test Description", "Card Number", "Account Number", "Card Type", "RRN (Retrieval Ref)", 
            "STAN / UTANO", "Before Txn Balance (LKR)", "Txn Amount (LKR)", "Service Charge", 
            "Actual Txn Amount (LKR)", "After Txn Balance (LKR)", "Front-End (FE) Status", 
            "Core Banking (SIBS) Status", "Overall Test Status", "Tester Remarks / Notes"
        ],
        "GRG_CRM_Cardbased_Fund_transfer": [
            "TC_ID", "Test Description", "Card Number", "Card Type", "From Account Number", "To Account Number", 
            "RRN (Retrieval Ref)", "STAN / UTANO", "Before Txn Balance (LKR) - From Acc", "Before Txn Balance (LKR) - To Acc", 
            "Txn Amount (LKR)", "After Txn Balance (LKR) - From Acc", "To After Txn Balance (LKR) To Acc", 
            "Front-End (FE) Status", "Core Banking (SIBS) Status", "Overall Test Status", "Tester Remarks / Notes"
        ],
        "GRG_CRM_Cardbased_SLIC_Bill_Payment": [
            "TC_ID", "Test Description", "Bill Number", "Card Number", "Account Number", "Card Type", 
            "RRN (Retrieval Ref)", "STAN / UTANO", "Txn Amount (LKR)", "Actual paid amount (LKR)", 
            "Front-End (FE) Status", "Core Banking (SIBS) Status", "Overall Test Status", "Tester Remarks / Notes"
        ],
        "GRG_CRM_Cardless_SLIC_Bill_Payment": [
            "TC_ID", "Test Description", "Bill Number", "Card Type", "RRN (Retrieval Ref)", "STAN / UTANO", 
            "Txn Amount (LKR)", "Service Charge", "Actual paid amount (LKR)", "Front-End (FE) Status", 
            "Core Banking (SIBS) Status", "Overall Test Status", "Tester Remarks / Notes"
        ],
        "GRG_CRM_Cardbased_Mini_Statement": [
            "TC_ID", "Test Description", "Card Number", "Account Number", "Card Type", "RRN (Retrieval Ref)", 
            "STAN / UTANO", "Before Txn Balance (LKR)", "Service Charge (LKR)", "Actual Txn Amount (LKR)", 
            "Front-End (FE) Status", "Core Banking (SIBS) Status", "Overall Test Status", "Tester Remarks / Notes"
        ]
    }

    selected_module = st.selectbox("Select Test Module / Feature", list(modules.keys()))
    columns = modules[selected_module]

    st.subheader(f"📝 Execute & Record: {selected_module}")
    
    with st.form(key="execution_form"):
        col1, col2 = st.columns(2)
        with col1:
            tc_id = st.text_input("Test Case ID (e.g., TC_CL_BP_001)", value="TC_CL_BP_001")
            test_desc = st.text_input("Test Description", value="Successful transaction test")
            rrn = st.text_input("RRN (Retrieval Reference Number)")
            stan = st.text_input("STAN / UTANO")
        with col2:
            txn_amount = st.number_input("Transaction Amount (LKR)", value=100.00)
            service_charge = st.number_input("Service Charge (LKR)", value=0.00)
            fe_status = st.selectbox("Front-End (FE) Status", ["SUCCESS", "FAILED"])
            sibs_status = st.selectbox("Core Banking (SIBS) Status", ["UPDATED (SIBS)", "PENDING", "FAILED"])

        remarks = st.text_area("Tester Remarks / Notes", value="Receipt printed successfully")
        
        submit_btn = st.form_submit_button("💾 Save Test Execution Record")
        
        if submit_btn:
            try:
                conn = psycopg2.connect(get_db_url())
                cur = conn.cursor()
                cur.execute("""
                    INSERT INTO uat_test_executions 
                    (module_name, tc_id, test_description, rrn, stan_utano, txn_amount, service_charge, 
                     fe_status, sibs_status, overall_status, tester_remarks, executed_by)
                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                """, (selected_module, tc_id, test_desc, rrn, stan, txn_amount, service_charge,
                      fe_status, sibs_status, "PASS" if fe_status=="SUCCESS" else "FAIL", remarks, "Current Tester"))
                conn.commit()
                cur.close()
                conn.close()
                st.success(f"Test case {tc_id} recorded successfully!")
            except Exception as e:
                st.error(f"Database save error (Check if table exists): {e}")

    st.markdown("---")
    st.subheader("📊 Export Multi-Tab Report for Finance Department")
    st.markdown("Download all executed test cases formatted into the exact multi-sheet workbook structure required by Finance.")

    if st.button("📥 Generate & Download Finance Excel Report"):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # Fetch data from DB or fallback to default templates if empty
            try:
                conn = psycopg2.connect(get_db_url())
                for mod_name, cols in modules.items():
                    df_mod = pd.read_sql(f"SELECT tc_id, test_description, rrn, stan_utano, txn_amount, service_charge, fe_status, sibs_status, overall_status, tester_remarks FROM uat_test_executions WHERE module_name = '{mod_name}'", conn)
                    if df_mod.empty:
                        # Create empty template placeholder row matching original format
                        df_mod = pd.DataFrame(columns=cols)
                    df_mod.to_excel(writer, sheet_name=mod_name[:31], index=False)
                conn.close()
            except Exception:
                # Fallback blank template sheets if DB table isn't created yet
                for mod_name, cols in modules.items():
                    df_blank = pd.DataFrame(columns=cols)
                    df_blank.to_excel(writer, sheet_name=mod_name[:31], index=False)
            
        output.seek(0)
        
        st.download_button(
            label="⬇️ Download Completed Excel Sheet (.xlsx)",
            data=output,
            file_name=f"GRG_CRM_Test_Execution_Finance_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

if __name__ == "__main__":
    render_test_execution_page()