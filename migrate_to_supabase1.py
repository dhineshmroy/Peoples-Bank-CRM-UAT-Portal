import pandas as pd
from sqlalchemy import create_engine

# Replace with your actual Supabase PostgreSQL connection string
# Format: postgresql://postgres:[YOUR-PASSWORD]@db.[YOUR-PROJECT-REF].supabase.co:5432/postgres
DB_URL = "postgresql://postgres.mxqtzquofvbcewtaowxo:Dhineshmelroy@aws-0-ap-southeast-2.pooler.supabase.com:5432/postgres?sslmode=require"
engine = create_engine(DB_URL)

excel_path = "CRM_Pre-Prod_Withdrawal_Visa_Mastercard_JCB_Completion_Report.xlsx"

# 1. Parse and clean Withdrawal - Card Matrix
mat_raw = pd.read_excel(excel_path, sheet_name="Withdrawal - Card Matrix", header=7)
mat_raw["TC ID"] = mat_raw["TC ID"].ffill()
mat_raw["Card Scheme"] = mat_raw["Card Scheme"].ffill()
mat_raw["Card Type"] = mat_raw["Card Type"].ffill()
mat_raw["Issuing Bank"] = mat_raw["Issuing Bank"].ffill()
mat_raw = mat_raw.dropna(subset=["TC ID"])

# Rename columns to match SQL snake_case
mat_raw.columns = [
    "tc_id",
    "card_scheme",
    "card_type",
    "issuing_bank",
    "account_type",
    "withdrawal_amount",
    "atm_crm_id",
    "account_reference_no",
    "rrn",
    "stan_utano",
    "fe_status",
    "sibs_status",
    "receipt_output",
    "overall_status",
    "execution_date",
    "tester",
    "remarks",
]
mat_raw.to_sql(
    "preprod_withdrawal_matrix",
    con=engine,
    if_exists="replace",
    index=False,
)
print("Successfully migrated Withdrawal - Card Matrix to Supabase!")

# 2. Parse and clean Execution Report (All Transactions)
exec_raw = pd.read_excel(excel_path, sheet_name="Execution Report", header=6)
exec_raw["TC ID"] = exec_raw["TC ID"].ffill()
exec_raw["Transaction Category"] = exec_raw["Transaction Category"].ffill()
exec_raw["Transaction / Test Description"] = exec_raw[
    "Transaction / Test Description"
].ffill()
exec_raw = exec_raw.dropna(subset=["TC ID"])

exec_raw.columns = [
    "tc_id",
    "transaction_category",
    "transaction_description",
    "card_type",
    "account_reference_no",
    "amount",
    "rrn",
    "stan_utano",
    "before_balance",
    "after_balance",
    "fe_status",
    "switch_status",
    "sibs_status",
    "receipt_output",
    "overall_status",
    "execution_date",
    "tester",
    "remarks",
]
exec_raw.to_sql(
    "preprod_all_transactions", con=engine, if_exists="replace", index=False
)
print("Successfully migrated All Transactions Execution Report to Supabase!")