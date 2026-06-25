
import streamlit as st
import pandas as pd
import zipfile
from datetime import datetime, timedelta

st.set_page_config(page_title="PUMA Pending Order Automation", layout="wide")
st.title("📦 PUMA SG Pending Order Automation")

# ------------------------
# Holiday Calendar
# ------------------------
default_holidays = ["2026-01-01"]
holiday_df = st.data_editor(
    pd.DataFrame({"Holiday Date": pd.to_datetime(default_holidays)}),
    num_rows="dynamic",
    use_container_width=True
)

PUBLIC_HOLIDAYS = pd.to_datetime(
    holiday_df["Holiday Date"], errors="coerce"
).dropna().dt.strftime("%Y-%m-%d").tolist()

def parse_date(value):
    try:
        return pd.to_datetime(value, dayfirst=True, errors="coerce")
    except:
        return pd.NaT

def is_public_holiday(date):
    return date.strftime("%Y-%m-%d") in PUBLIC_HOLIDAYS

def add_working_days(date, days):
    current = date
    added = 0
    while added < days:
        current += timedelta(days=1)
        if current.weekday() != 6 and not is_public_holiday(current):
            added += 1
    return current

def move_next_working_day(date):
    while date.weekday() == 6 or is_public_holiday(date):
        date += timedelta(days=1)
    return date

def calculate_lazada_sla(order_date):
    if pd.isna(order_date):
        return pd.NaT
    sla = order_date if order_date.hour < 13 else add_working_days(order_date, 1)
    return move_next_working_day(sla).replace(hour=23, minute=59, second=59)

def calculate_shopee_sla(order_date, payment_method=""):
    if pd.isna(order_date):
        return pd.NaT
    sla = order_date if order_date.hour < 12 else add_working_days(order_date, 1)
    return move_next_working_day(sla).replace(hour=23, minute=59, second=59)

def calculate_zalora_sla(order_date):
    if pd.isna(order_date):
        return pd.NaT
    sla = add_working_days(order_date, 2)
    return move_next_working_day(sla).replace(hour=23, minute=59, second=59)

def read_file(uploaded_file):
    if uploaded_file is None:
        return None
    name = uploaded_file.name.lower()
    if name.endswith(".csv"):
        return pd.read_csv(uploaded_file, dtype=str)
    if name.endswith(".xlsx"):
        return pd.read_excel(uploaded_file, dtype=str)
    if name.endswith(".zip"):
        frames = []
        with zipfile.ZipFile(uploaded_file) as z:
            for f in z.namelist():
                if f.endswith(".csv"):
                    frames.append(pd.read_csv(z.open(f), dtype=str))
                elif f.endswith(".xlsx"):
                    frames.append(pd.read_excel(z.open(f), dtype=str))
        return pd.concat(frames, ignore_index=True) if frames else None
    return None

c1, c2 = st.columns(2)
with c1:
    tc_file = st.file_uploader("TC Report", type=["xlsx","csv","zip"])
    oms_file = st.file_uploader("OMS Report", type=["xlsx","csv","zip"])
with c2:
    lazada_file = st.file_uploader("Lazada", type=["xlsx","csv","zip"])
    shopee_file = st.file_uploader("Shopee", type=["xlsx","csv","zip"])
    zalora_file = st.file_uploader("Zalora", type=["xlsx","csv","zip"])

if st.button("🚀 Run Automation"):
    tc_df = read_file(tc_file)
    oms_df = read_file(oms_file)
    lazada_df = read_file(lazada_file)
    shopee_df = read_file(shopee_file)
    zalora_df = read_file(zalora_file)

    st.success("Files loaded successfully")

    consolidated = []

    if lazada_df is not None:
        st.write("Lazada records:", len(lazada_df))

    if shopee_df is not None:
        st.write("Shopee records:", len(shopee_df))

    if zalora_df is not None:
        st.write("Zalora records:", len(zalora_df))

    st.info("Part 1 + Part 2 starter application generated. Extend column mappings according to your source files.")
