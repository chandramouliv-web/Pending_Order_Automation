import streamlit as st
import pandas as pd
import zipfile
import io
import requests

from datetime import datetime,timedelta
st.set_page_config(
    page_title="PUMA Pending Order Automation",
    page_icon="📦",
    layout="wide"
)

st.title("📦 PUMA SG Pending Order Automation")
PUBLIC_HOLIDAYS = []
def normalize_marketplace(name):

    name = str(name).lower().strip()

    if "shopee" in name:
        return "Shopee SG"

    if "lazada" in name:
        return "Lazada SG"

    if "zalora" in name:
        return "Zalora SG"

    return name
def normalize_marketplace(name):

    name = str(name).lower().strip()

    if "shopee" in name:
        return "Shopee SG"

    if "lazada" in name:
        return "Lazada SG"

    if "zalora" in name:
        return "Zalora SG"

    return name

def parse_date(value):

    if pd.isna(value):
        return None

    if isinstance(value,datetime):
        return value

    try:
        return pd.to_datetime(
            value,
            dayfirst=True,
            errors="coerce"
        )
    except:
        return None
def is_public_holiday(date):

    return (
        date.strftime("%Y-%m-%d")
        in PUBLIC_HOLIDAYS
    )        
def add_working_days(date,days):

    current = date
    added = 0

    while added < days:

        current += timedelta(days=1)

        is_sunday = current.weekday() == 6

        if (
            not is_sunday and
            not is_public_holiday(current)
        ):
            added += 1

    return current
def move_next_working_day(date):

    current = date

    while (
        current.weekday()==6
        or
        is_public_holiday(current)
    ):
        current += timedelta(days=1)

    return current
def calculate_lazada_sla(order_date):

    hour = order_date.hour

    if order_date.weekday() == 6:

        sla = add_working_days(
            order_date,
            1
        )

    elif hour < 13:

        sla = order_date

    else:

        sla = add_working_days(
            order_date,
            1
        )

    sla = move_next_working_day(sla)

    sla = sla.replace(
        hour=23,
        minute=59,
        second=59
    )

    return sla
def calculate_shopee_sla(
    order_date,
    payment_method=""
):

    payment_method = str(
        payment_method
    ).lower()

    hour = order_date.hour

    if hour < 12:

        sla = order_date

    else:

        sla = add_working_days(
            order_date,
            1
        )

    sla = move_next_working_day(sla)

    sla = sla.replace(
        hour=23,
        minute=59,
        second=59
    )

    return sla        
def calculate_zalora_sla(order_date):

    sla = add_working_days(
        order_date,
        2
    )

    sla = move_next_working_day(sla)

    sla = sla.replace(
        hour=23,
        minute=59,
        second=59
    )

    return sla    
def read_file(uploaded_file):

    if uploaded_file is None:
        return None

    filename = uploaded_file.name.lower()

    # CSV
    if filename.endswith(".csv"):

        return pd.read_csv(
            uploaded_file,
            dtype=str
        )

    # XLSX
    if filename.endswith(".xlsx"):

        return pd.read_excel(
            uploaded_file,
            dtype=str
        )

    # ZIP
    if filename.endswith(".zip"):

        all_frames = []

        with zipfile.ZipFile(
            uploaded_file
        ) as z:

            for file in z.namelist():

                if file.endswith(".csv"):

                    df = pd.read_csv(
                        z.open(file),
                        dtype=str
                    )

                    all_frames.append(df)

                elif file.endswith(".xlsx"):

                    df = pd.read_excel(
                        z.open(file),
                        dtype=str
                    )

                    all_frames.append(df)

        if all_frames:

            return pd.concat(
                all_frames,
                ignore_index=True
            )

    return None
st.header("📤 Upload Files")

c1,c2 = st.columns(2)

with c1:

    tc_file = st.file_uploader(
        "TC Order Report",
        type=["xlsx","csv","zip"]
    )

    oms_file = st.file_uploader(
        "OMS Report",
        type=["xlsx","csv","zip"]
    )

    holiday_file = st.file_uploader(
        "Holiday File",
        type=["xlsx","csv"]
    )

with c2:

    lazada_file = st.file_uploader(
        "Lazada SG",
        type=["xlsx","csv","zip"]
    )

    shopee_file = st.file_uploader(
        "Shopee SG",
        type=["xlsx","csv","zip"]
    )

    zalora_file = st.file_uploader(
        "Zalora SG",
        type=["xlsx","csv","zip"]
    )            
if holiday_file:

    holiday_df = read_file(
        holiday_file
    )

    first_col = holiday_df.columns[0]

    PUBLIC_HOLIDAYS = pd.to_datetime(
        holiday_df[first_col]
    ).dt.strftime(
        "%Y-%m-%d"
    ).tolist()

    st.success(
        f"{len(PUBLIC_HOLIDAYS)} Holidays Loaded"
    )
run = st.button(
    "🚀 Run Automation",
    use_container_width=True
)
if run:

    tc_df = read_file(tc_file)

    oms_df = read_file(oms_file)

    lazada_df = read_file(lazada_file)

    shopee_df = read_file(shopee_file)

    zalora_df = read_file(zalora_file)

    st.success(
        "All Files Loaded Successfully"
    )
    consolidated = []   
    if lazada_df is not None:

        for _,row in lazada_df.iterrows():

            order = row.get(
                "orderNumber"
            )

            sku = row.get(
                "sellerSku"
            )

            status = row.get(
                "status"
            )

            date = parse_date(
                row.get("createTime")
            )

            if pd.isna(order) or pd.isna(sku):
                continue

            sla = calculate_lazada_sla(
                date
            )

            consolidated.append({

                "order":order,
                "sku":sku,
                "status":status,
                "order_date":date,
                "sla":sla,
                "marketplace":"Lazada SG"

            })
    if shopee_df is not None:

        for _,row in shopee_df.iterrows():

            order = row.get("Order ID")

            sku = row.get(
                "SKU Reference No."
            )

            status = row.get(
                "status"
            )

            date = parse_date(
                row.get(
                    "Order Creation Date"
                )
            )

            payment = row.get(
                "Payment Method"
            )

            sla = calculate_shopee_sla(
                date,
                payment
            )

            consolidated.append({

                "order":order,
                "sku":sku,
                "status":status,
                "order_date":date,
                "sla":sla,
                "marketplace":"Shopee SG"

            })
    if zalora_df is not None:

        for _,row in zalora_df.iterrows():

            order = row.get(
                "Order Number"
            )

            sku = row.get(
                "Seller SKU"
            )

            status = row.get(
                "status"
            )

            date = parse_date(
                row.get(
                    "Created at"
                )
            )

            sla = calculate_zalora_sla(
                date
            )

            consolidated.append({

                "order":order,
                "sku":sku,
                "status":status,
                "order_date":date,
                "sla":sla,
                "marketplace":"Zalora SG"

            })
    mp_df = pd.DataFrame(
        consolidated
    )

    st.subheader(
        "Marketplace Consolidated"
    )

    st.dataframe(
        mp_df,
        use_container_width=True
    )                         
       
