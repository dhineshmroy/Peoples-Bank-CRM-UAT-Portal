"""
One-time import of the Pre-Prod Excel report into Supabase.

WARNING: if_exists="replace" DROPS and re-creates the two tables, which erases
test results and uploaded receipts already saved in the app. Only run this on a
fresh setup, or set REPLACE_TABLES = False to append instead.

Every column is created as TEXT (except receipt_output = BYTEA), so blank values
such as ATM/CRM ID, RRN, UTANO or amounts can never break a save again.

Set the connection string as an environment variable - do NOT hard-code the password:
    export SUPABASE_DB_URL="postgresql://postgres.<ref>:<PASSWORD>@aws-0-ap-southeast-2.pooler.supabase.com:5432/postgres?sslmode=require"
"""
import os

import pandas as pd
from sqlalchemy import create_engine
from sqlalchemy.types import LargeBinary, Text

DB_URL = "postgresql://postgres.mxqtzquofvbcewtaowxo:Dhineshmelroy@aws-0-ap-southeast-2.pooler.supabase.com:5432/postgres?sslmode=require"
engine = create_engine(DB_URL)
EXCEL_PATH = "CRM_Pre-Prod_Withdrawal_Visa_Mastercard_JCB_Completion_Report.xlsx"
REPLACE_TABLES = True

engine = create_engine(DB_URL)
if_exists = "replace" if REPLACE_TABLES else "append"


def clean_value(v):
    """Any cell -> clean text, or None (NULL) for blanks / NaN / 'nan'."""
    if v is None:
        return None
    try:
        if pd.isna(v):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(v, float) and v.is_integer():
        v = int(v)                      # 700.0 -> '700', 260922002657 -> no '.0' / no e+11
    text = str(v).strip()
    return None if text.lower() in ("", "nan", "none", "nat", "<na>") else text


def to_text_frame(df: pd.DataFrame) -> pd.DataFrame:
    """Everything -> clean text; blanks -> NULL; receipt_output left empty (filled later by the app)."""
    out = pd.DataFrame(index=df.index)
    for col in df.columns:
        if col == "receipt_output":
            out[col] = pd.Series([None] * len(df), index=df.index, dtype="object")
        else:
            out[col] = pd.Series([clean_value(v) for v in df[col].tolist()], index=df.index, dtype="object")
    return out


def dtype_map(columns):
    return {c: (LargeBinary() if c == "receipt_output" else Text()) for c in columns}


# 1. Withdrawal - Card Matrix
mat = pd.read_excel(EXCEL_PATH, sheet_name="Withdrawal - Card Matrix", header=7)
for c in ["TC ID", "Card Scheme", "Card Type", "Issuing Bank"]:
    mat[c] = mat[c].ffill()
mat = mat.dropna(subset=["TC ID"])
mat.columns = ["tc_id", "card_scheme", "card_type", "issuing_bank", "account_type", "withdrawal_amount",
               "atm_crm_id", "account_reference_no", "rrn", "stan_utano", "fe_status", "sibs_status",
               "receipt_output", "overall_status", "execution_date", "tester", "remarks"]
mat = to_text_frame(mat)
mat.to_sql("preprod_withdrawal_matrix", con=engine, if_exists=if_exists, index=False, dtype=dtype_map(mat.columns))
print("Migrated Withdrawal - Card Matrix:", len(mat), "rows")

# 2. Execution Report (All Transactions)
ex = pd.read_excel(EXCEL_PATH, sheet_name="Execution Report", header=6)
for c in ["TC ID", "Transaction Category", "Transaction / Test Description"]:
    ex[c] = ex[c].ffill()
ex = ex.dropna(subset=["TC ID"])
ex.columns = ["tc_id", "transaction_category", "transaction_description", "card_type", "account_reference_no",
              "amount", "rrn", "stan_utano", "before_balance", "after_balance", "fe_status", "switch_status",
              "sibs_status", "receipt_output", "overall_status", "execution_date", "tester", "remarks"]
ex = to_text_frame(ex)
ex.to_sql("preprod_all_transactions", con=engine, if_exists=if_exists, index=False, dtype=dtype_map(ex.columns))
print("Migrated All Transactions:", len(ex), "rows")
