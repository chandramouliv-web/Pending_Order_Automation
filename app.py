from __future__ import annotations

import io
import json
import os
import smtplib
import ssl
import zipfile
from datetime import date, datetime, time, timedelta
from email.message import EmailMessage
from typing import Any
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components


CLEANED_HEADERS = [
    "nickname",
    "order_id",
    "order_number",
    "custom_sku",
    "order_status",
    "ordered_date",
    "payment_status",
    "payment_method",
    "Order X SKU",
    "Marketplace",
    "MP Status",
    "MP SLA",
    "OMS status",
    "SLA Status",
]

MISSING_HEADERS = [
    "Type",
    "Marketplace",
    "Order",
    "SKU",
    "Status",
    "Payment Status",
    "Payment Method",
    "Order Date",
    "Tracking",
]

DEFAULT_TO_EMAILS = "sharon.chua@puma.com, kayla.zhang@puma.com, josegabriel.mendoza@puma.com, gp_puma_sg_ops@ych.com"
DEFAULT_CC_EMAILS = "puma-ecops@graas.ai, sonal.aggarwal@puma.com, ecops-all@graas.ai, harvesters@graas.ai"

SMTP_PRESETS = {
    "Microsoft 365 / Outlook": ("smtp.office365.com", 587, False),
    "Gmail": ("smtp.gmail.com", 587, False),
    "Gmail SSL": ("smtp.gmail.com", 465, True),
    "Custom": ("", 587, False),
}


st.set_page_config(page_title="PUMA SG Pending Orders", page_icon="Box", layout="wide")


def normalize_marketplace(name: Any) -> str:
    value = str(name or "").lower().strip()
    if "shopee" in value:
        return "Shopee SG"
    if "lazada" in value:
        return "Lazada SG"
    if "zalora" in value:
        return "Zalora SG"
    return value or "N/A"


def parse_date(value: Any) -> datetime | None:
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        return value
    if isinstance(value, date):
        return datetime.combine(value, time.min)
    if pd.isna(value):
        return None

    parsed = pd.to_datetime(value, errors="coerce", dayfirst=True)
    if not pd.isna(parsed):
        return parsed.to_pydatetime()

    raw = str(value).strip()
    if not raw:
        return None

    parts = raw.split()
    date_part = parts[0]
    time_part = parts[1] if len(parts) > 1 else "00:00:00"
    date_bits = date_part.replace("-", "/").split("/")
    if len(date_bits) != 3:
        return None

    try:
        first, second, year = [int(bit) for bit in date_bits]
        if year < 100:
            year += 2000
        if first > 12:
            day, month = first, second
        elif second > 12:
            month, day = first, second
        else:
            day, month = first, second
        time_bits = [int(bit) for bit in time_part.split(":")]
        while len(time_bits) < 3:
            time_bits.append(0)
        return datetime(year, month, day, time_bits[0], time_bits[1], time_bits[2])
    except ValueError:
        return None


def format_datetime(value: Any) -> str:
    parsed = parse_date(value)
    return parsed.strftime("%d-%b-%Y %H:%M") if parsed else ""


def format_date(value: Any) -> str:
    parsed = parse_date(value)
    return parsed.strftime("%d-%b-%Y") if parsed else ""


def is_public_holiday(value: datetime, holidays: set[date]) -> bool:
    return value.date() in holidays


def add_working_days(value: datetime, days: int, holidays: set[date]) -> datetime:
    temp = value
    added = 0
    while added < days:
        temp += timedelta(days=1)
        if temp.weekday() != 6 and not is_public_holiday(temp, holidays):
            added += 1
    return temp


def move_to_next_working_day(value: datetime, holidays: set[date]) -> datetime:
    temp = value
    while temp.weekday() == 6 or is_public_holiday(temp, holidays):
        temp += timedelta(days=1)
    return temp


def end_of_day(value: datetime) -> datetime:
    return value.replace(hour=23, minute=59, second=59, microsecond=0)


def calculate_lazada_sla(order_date: datetime, holidays: set[date]) -> datetime:
    if order_date.weekday() == 6 or order_date.hour >= 13:
        sla_date = add_working_days(order_date, 1, holidays)
    else:
        sla_date = order_date
    return end_of_day(move_to_next_working_day(sla_date, holidays))


def calculate_shopee_sla(order_date: datetime, holidays: set[date]) -> datetime:
    sla_date = order_date if order_date.hour < 12 else add_working_days(order_date, 1, holidays)
    return end_of_day(move_to_next_working_day(sla_date, holidays))


def calculate_zalora_sla(order_date: datetime, holidays: set[date]) -> datetime:
    sla_date = add_working_days(order_date, 2, holidays)
    return end_of_day(move_to_next_working_day(sla_date, holidays))


def get_index(columns: list[str], name: str) -> str | None:
    needle = name.lower().strip()
    normalized = {str(col).lower().strip(): col for col in columns}
    if needle in normalized:
        return normalized[needle]
    for col in columns:
        if needle in str(col).lower().strip():
            return col
    return None


def cell(row: pd.Series, column: str | None) -> Any:
    if not column:
        return ""
    value = row.get(column, "")
    return "" if pd.isna(value) else value


def read_one_file(uploaded_file: Any) -> pd.DataFrame:
    if uploaded_file is None:
        return pd.DataFrame()
    name = uploaded_file.name.lower()
    data = uploaded_file.getvalue()
    if name.endswith(".zip"):
        frames: list[pd.DataFrame] = []
        with zipfile.ZipFile(io.BytesIO(data)) as archive:
            for member in archive.namelist():
                lower = member.lower()
                if lower.endswith((".csv", ".xlsx", ".xls")):
                    with archive.open(member) as handle:
                        frames.append(read_bytes(handle.read(), lower))
        return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
    return read_bytes(data, name)


def read_bytes(data: bytes, name: str) -> pd.DataFrame:
    if name.endswith(".csv"):
        return pd.read_csv(io.BytesIO(data), dtype=str, keep_default_na=False)
    if name.endswith(".xls"):
        return pd.read_excel(io.BytesIO(data), dtype=str, keep_default_na=False, engine="xlrd")
    return pd.read_excel(io.BytesIO(data), dtype=str, keep_default_na=False)


def build_marketplace_map(reports: dict[str, pd.DataFrame], holidays: set[date]) -> dict[str, dict[str, Any]]:
    marketplace_rows: list[dict[str, Any]] = []

    def process_sheet(frame: pd.DataFrame, mapping: dict[str, str], marketplace_type: str, marketplace: str) -> None:
        if frame.empty:
            return

        columns = list(frame.columns)
        order_col = get_index(columns, mapping["order"])
        sku_col = get_index(columns, mapping["sku"])
        status_col = get_index(columns, mapping["status"])
        date_col = get_index(columns, mapping["date"])
        payment_col = get_index(columns, "payment method")

        tracking_col = None
        if marketplace_type == "lazada":
            tracking_col = get_index(columns, "trackingCode")
        elif marketplace_type == "shopee":
            tracking_col = get_index(columns, "Tracking Number*") or get_index(columns, "Tracking Number")
        elif marketplace_type == "zalora":
            tracking_col = get_index(columns, "Tracking Code")

        if not all([order_col, sku_col, status_col, date_col]):
            return

        for _, row in frame.iterrows():
            order = cell(row, order_col)
            sku = cell(row, sku_col)
            if not order or not sku:
                continue

            order_date = parse_date(cell(row, date_col))
            if not order_date:
                continue

            if marketplace_type == "lazada":
                sla = calculate_lazada_sla(order_date, holidays)
            elif marketplace_type == "shopee":
                sla = calculate_shopee_sla(order_date, holidays)
            else:
                sla = calculate_zalora_sla(order_date, holidays)

            marketplace_rows.append(
                {
                    "key": f"{str(order).strip()}{str(sku).strip()}",
                    "order": order,
                    "sku": sku,
                    "status": cell(row, status_col),
                    "order_date": order_date,
                    "sla": sla,
                    "marketplace": normalize_marketplace(marketplace),
                    "tracking": cell(row, tracking_col),
                    "payment_method": cell(row, payment_col),
                }
            )

    process_sheet(
        reports["lazada"],
        {"order": "orderNumber", "sku": "sellerSku", "status": "status", "date": "createTime"},
        "lazada",
        "Lazada SG",
    )
    process_sheet(
        reports["shopee"],
        {"order": "Order ID", "sku": "SKU Reference No.", "status": "status", "date": "Order Creation Date"},
        "shopee",
        "Shopee SG",
    )
    process_sheet(
        reports["zalora"],
        {"order": "Order Number", "sku": "Seller SKU", "status": "status", "date": "Created at"},
        "zalora",
        "Zalora SG",
    )

    return {row["key"]: row for row in marketplace_rows}


def build_oms_map(oms: pd.DataFrame) -> dict[str, Any]:
    if oms.empty:
        return {}

    columns = list(oms.columns)
    order_col = get_index(columns, "order")
    ean_col = get_index(columns, "ean")
    status_col = get_index(columns, "line_status")
    if not all([order_col, ean_col, status_col]):
        return {}

    result = {}
    for _, row in oms.iterrows():
        key = f"{str(cell(row, order_col)).strip()}{str(cell(row, ean_col)).strip()}"
        if key:
            result[key] = cell(row, status_col)
    return result


def is_closed_status(status: Any) -> bool:
    value = str(status or "").lower()
    return any(word in value for word in ["delivered", "shipped", "returned", "cancelled", "canceled"])


def should_skip_missing_status(status: Any) -> bool:
    value = str(status or "").lower().strip()
    return any(
        word in value
        for word in ["cancelled", "canceled", "delivered", "return shipped", "returned", "return accepted", "return requested"]
    )


def get_sla_status(sla_date: datetime | None) -> str:
    if not sla_date:
        return ""

    now = datetime.now()
    diff = sla_date - now
    diff_days = (sla_date.date() - now.date()).days

    if diff.total_seconds() < 0:
        return "BREACHED"
    if diff <= timedelta(hours=6):
        return "Critical - Ship ASAP"
    if diff_days == 0:
        return "Need to Ship by Today"
    if diff_days == 1:
        return "Ship by Tomorrow"
    return f"On Track (SLA: {format_date(sla_date)})"


def process_tc(tc: pd.DataFrame, mp_map: dict[str, dict[str, Any]], oms_map: dict[str, Any]) -> pd.DataFrame:
    if tc.empty:
        return pd.DataFrame(columns=CLEANED_HEADERS)

    columns = list(tc.columns)
    idx = {
        "nickname": get_index(columns, "nickname"),
        "order_id": get_index(columns, "order_id"),
        "order_number": get_index(columns, "order_number"),
        "sku": get_index(columns, "custom_sku"),
        "status": get_index(columns, "order_status"),
        "date": get_index(columns, "ordered_date"),
        "sla": get_index(columns, "order_sla"),
        "payment": get_index(columns, "payment_status"),
        "payment_method": get_index(columns, "payment_method"),
    }

    output: list[list[Any]] = []

    for _, row in tc.iterrows():
        order = cell(row, idx["order_number"])
        sku = cell(row, idx["sku"])
        if not order or not sku:
            continue

        key = f"{str(order).strip()}{str(sku).strip()}"
        mp = mp_map.get(key, {})
        mp_status = mp.get("status", "")
        oms_status = oms_map.get(key, "")

        if not str(mp_status).strip():
            continue
        if "canceled" in str(mp_status).lower() and not str(oms_status).strip():
            continue

        tc_status = cell(row, idx["status"])
        payment_status = str(cell(row, idx["payment"])).lower()
        payment_method = str(cell(row, idx["payment_method"])).lower()

        if (
            "unpaid" in str(mp_status).lower()
            and "cod" not in payment_method
            and "pending" in payment_status
            and "new" in str(tc_status).lower()
        ):
            continue

        if not oms_status and ("completed" in payment_status or "cod" in payment_method):
            oms_status = "Not Reflected in OMS"

        if is_closed_status(tc_status) or is_closed_status(mp_status) or is_closed_status(oms_status):
            continue

        sla_date = parse_date(mp.get("sla")) if mp.get("sla") else parse_date(cell(row, idx["sla"]))
        output.append(
            [
                cell(row, idx["nickname"]),
                cell(row, idx["order_id"]),
                order,
                sku,
                tc_status,
                format_datetime(cell(row, idx["date"])),
                cell(row, idx["payment"]),
                cell(row, idx["payment_method"]),
                key,
                mp.get("marketplace", "N/A"),
                mp_status,
                format_datetime(sla_date) if sla_date else "",
                oms_status,
                get_sla_status(sla_date),
            ]
        )

    cleaned = pd.DataFrame(output, columns=CLEANED_HEADERS)
    if cleaned.empty:
        return cleaned

    cleaned["_sort_sla"] = cleaned["MP SLA"].apply(parse_date)
    status_rank = {
        "BREACHED": 0,
        "Critical - Ship ASAP": 1,
        "Need to Ship by Today": 2,
        "Ship by Tomorrow": 3,
    }
    cleaned["_rank"] = cleaned["SLA Status"].map(status_rank).fillna(4)
    cleaned = cleaned.sort_values(["_rank", "_sort_sla"], na_position="last").drop(columns=["_rank", "_sort_sla"])
    return cleaned.reset_index(drop=True)


def build_missing_rows(tc: pd.DataFrame, mp_map: dict[str, dict[str, Any]], oms_map: dict[str, Any]) -> pd.DataFrame:
    if tc.empty:
        return pd.DataFrame(columns=MISSING_HEADERS)

    columns = list(tc.columns)
    idx = {
        "order_number": get_index(columns, "order_number"),
        "sku": get_index(columns, "custom_sku"),
        "status": get_index(columns, "order_status"),
        "payment": get_index(columns, "payment_status"),
        "payment_method": get_index(columns, "payment_method"),
        "date": get_index(columns, "ordered_date"),
    }

    rows: list[list[Any]] = []
    seen: set[str] = set()

    for _, row in tc.iterrows():
        order = cell(row, idx["order_number"])
        sku = cell(row, idx["sku"])
        key = f"{str(order).strip()}{str(sku).strip()}"
        if not key or key in seen or key in oms_map:
            continue

        order_status = cell(row, idx["status"])
        if should_skip_missing_status(order_status):
            continue

        mp = mp_map.get(key, {})
        rows.append(
            [
                "TC not reflected in OMS",
                mp.get("marketplace", "N/A"),
                order,
                sku,
                order_status,
                cell(row, idx["payment"]),
                cell(row, idx["payment_method"]),
                format_datetime(cell(row, idx["date"])),
                mp.get("tracking", ""),
            ]
        )
        seen.add(key)

    for key, mp in mp_map.items():
        if key in seen or key in oms_map or should_skip_missing_status(mp.get("status")):
            continue
        rows.append(
            [
                "Marketplace not reflected in OMS",
                mp.get("marketplace", "N/A"),
                mp.get("order", ""),
                mp.get("sku", ""),
                mp.get("status", ""),
                "",
                "",
                format_datetime(mp.get("order_date")),
                mp.get("tracking", ""),
            ]
        )
        seen.add(key)

    return pd.DataFrame(rows, columns=MISSING_HEADERS)


def build_summary(cleaned: pd.DataFrame) -> tuple[dict[str, int], pd.DataFrame]:
    if cleaned.empty:
        return {"breached": 0, "today": 0, "new_oms": 0, "not_oms": 0, "within": 0}, pd.DataFrame()

    urgent_mask = cleaned["SLA Status"].isin(["Critical - Ship ASAP", "Need to Ship by Today"])
    breached_mask = cleaned["SLA Status"] == "BREACHED"
    not_oms_mask = cleaned["OMS status"].str.lower().str.contains("not reflected", na=False)
    new_oms_mask = cleaned["OMS status"].str.lower().str.contains("new", na=False)
    counts = {
        "breached": int(breached_mask.sum()),
        "today": int(urgent_mask.sum()),
        "new_oms": int(new_oms_mask.sum()),
        "not_oms": int(not_oms_mask.sum()),
        "within": int((~breached_mask & ~urgent_mask).sum()),
    }

    pivot_source = cleaned.copy()
    pivot_source["SLA Date"] = pivot_source["MP SLA"].apply(format_date)
    pivot = pd.pivot_table(
        pivot_source,
        values="Order X SKU",
        index=["Marketplace", "OMS status"],
        columns="SLA Date",
        aggfunc="count",
        fill_value=0,
        margins=True,
        margins_name="Grand Total",
    )
    return counts, pivot


def make_excel(cleaned: pd.DataFrame, missing: pd.DataFrame, pivot: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        cleaned.to_excel(writer, sheet_name="TC_Cleaned", index=False)
        missing.to_excel(writer, sheet_name="Missing_Orders", index=False)
        if not pivot.empty:
            pivot.to_excel(writer, sheet_name="Dashboard")
    return output.getvalue()


def split_emails(value: str) -> list[str]:
    return [email.strip() for email in value.replace(";", ",").split(",") if email.strip()]


def dashboard_cell_color(date_label: Any, oms_status: Any, value: Any) -> str:
    if not value:
        return "#f2f2f2"
    if "not reflected" in str(oms_status).lower():
        return "#ffff00"

    parsed = parse_date(date_label)
    if not parsed:
        return "#f2f2f2"

    today = datetime.now().date()
    if parsed.date() < today:
        return "#ff0000"
    if parsed.date() == today:
        return "#ffc000"
    return "#92d050"


def make_dashboard_html(counts: dict[str, int], pivot: pd.DataFrame) -> str:
    logo_url = "https://images.graas.ai/uploads/merchant/GED/profile-1/profile-1.jpg?time=1778413961"

    if pivot.empty:
        dashboard_table = """
        <table style="border-collapse:collapse;font-family:Consolas,monospace;font-size:11px;">
          <tr><td style="padding:16px;border:1px solid #d9d9d9;">No dashboard rows generated.</td></tr>
        </table>
        """
    else:
        display_pivot = pivot.copy()
        dashboard_table = """
        <table cellpadding="0" cellspacing="0" style="border-collapse:collapse;font-family:Consolas,monospace;font-size:11px;text-align:center;">
          <tr>
            <td colspan="{colspan}" style="background:#111827;color:#ffffff;padding:12px 10px;">
              <table cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;">
                <tr>
                  <td style="width:70px;text-align:left;"><img src="{logo_url}" alt="PUMA" style="height:32px;background:#ffffff;padding:3px;display:block;"></td>
                  <td style="font-size:20px;font-weight:bold;letter-spacing:1px;">📦 Pending Orders - <span style="background:#fde68a;color:#000;padding:1px 4px;">SG</span></td>
                  <td style="width:70px;"></td>
                </tr>
              </table>
            </td>
          </tr>
        """.format(colspan=len(display_pivot.columns) + 3, logo_url=logo_url)

        dashboard_table += """
          <tr style="background:#f2f2f2;font-weight:bold;">
            <td style="border:1px solid #d9d9d9;padding:7px 10px;">Marketplace</td>
            <td style="border:1px solid #d9d9d9;padding:7px 10px;">OMS Status</td>
        """
        for column in display_pivot.columns:
            dashboard_table += f'<td style="border:1px solid #d9d9d9;padding:7px 10px;">{column}</td>'
        dashboard_table += "</tr>"

        if isinstance(display_pivot.index, pd.MultiIndex):
            grouped = {}
            for index_values, row in display_pivot.iterrows():
                marketplace, oms_status = index_values
                grouped.setdefault(marketplace, []).append((oms_status, row))

            for marketplace, rows in grouped.items():
                for row_index, (oms_status, row) in enumerate(rows):
                    dashboard_table += "<tr>"
                    if row_index == 0:
                        dashboard_table += (
                            f'<td rowspan="{len(rows)}" style="border:1px solid #d9d9d9;padding:7px 10px;'
                            f'font-weight:bold;background:#ffffff;">{marketplace}</td>'
                        )
                    dashboard_table += (
                        f'<td style="border:1px solid #d9d9d9;padding:7px 10px;font-weight:bold;">{oms_status}</td>'
                    )
                    for column, value in row.items():
                        display_value = "" if int(value) == 0 else int(value)
                        bg = "#f2f2f2" if column == "Grand Total" else dashboard_cell_color(column, oms_status, display_value)
                        dashboard_table += (
                            f'<td style="border:1px solid #d9d9d9;padding:7px 10px;background:{bg};'
                            f'font-weight:bold;">{display_value}</td>'
                        )
                    dashboard_table += "</tr>"
        else:
            for index_value, row in display_pivot.iterrows():
                dashboard_table += f'<tr><td colspan="2" style="border:1px solid #d9d9d9;padding:7px 10px;font-weight:bold;">{index_value}</td>'
                for value in row:
                    dashboard_table += f'<td style="border:1px solid #d9d9d9;padding:7px 10px;font-weight:bold;">{int(value)}</td>'
                dashboard_table += "</tr>"

        dashboard_table += "</table>"

    summary_table = f"""
    <table cellpadding="0" cellspacing="0" style="border-collapse:collapse;font-family:Consolas,monospace;font-size:12px;font-weight:bold;width:195px;">
      <tr><td colspan="2" style="background:#111827;color:#ffffff;padding:10px;text-align:center;">📊 Order Summary</td></tr>
      <tr><td style="background:#ff0000;color:#ffffff;padding:10px;">🚨 Breached</td><td style="background:#fff1f2;color:#b91c1c;padding:10px;text-align:center;">{counts.get("breached", 0)}</td></tr>
      <tr><td style="background:#ffc000;color:#000000;padding:10px;">📦 Handover Today</td><td style="background:#fff7ed;color:#c2410c;padding:10px;text-align:center;">{counts.get("today", 0)}</td></tr>
      <tr><td style="background:#2563eb;color:#ffffff;padding:10px;">🆕 Order Status at NEW</td><td style="background:#eff6ff;color:#1d4ed8;padding:10px;text-align:center;">{counts.get("new_oms", 0)}</td></tr>
      <tr><td style="background:#92d050;color:#000000;padding:10px;">✅ Within SLA</td><td style="background:#f0fdf4;color:#166534;padding:10px;text-align:center;">{counts.get("within", 0)}</td></tr>
      <tr><td style="background:#ffff00;color:#000000;padding:10px;">⚠️ Not Reflected in OMS</td><td style="background:#fefce8;color:#854d0e;padding:10px;text-align:center;">{counts.get("not_oms", 0)}</td></tr>
    </table>
    """

    return f"""
    <table cellpadding="0" cellspacing="0" style="border-collapse:collapse;">
      <tr>
        <td valign="top">{dashboard_table}</td>
        <td style="width:30px;"></td>
        <td valign="top">{summary_table}</td>
      </tr>
    </table>
    """


def make_details_table_html(cleaned: pd.DataFrame, limit: int = 80) -> str:
    if cleaned.empty:
        return "<p>No pending order details generated.</p>"

    detail_columns = [
        "Marketplace",
        "order_number",
        "custom_sku",
        "MP Status",
        "MP SLA",
        "OMS status",
        "SLA Status",
    ]
    existing_columns = [column for column in detail_columns if column in cleaned.columns]
    detail_rows = cleaned[existing_columns].head(limit)

    html = """
    <table cellpadding="0" cellspacing="0" style="border-collapse:collapse;font-family:Consolas,monospace;font-size:11px;margin-top:18px;">
      <tr>
    """
    for column in existing_columns:
        html += f'<td style="border:1px solid #d9d9d9;background:#1f2937;color:#ffffff;padding:7px 10px;font-weight:bold;">{column}</td>'
    html += "</tr>"

    for _, row in detail_rows.iterrows():
        status = row.get("SLA Status", "")
        if status == "BREACHED":
            bg, color = "#ffe4e6", "#991b1b"
        elif status == "Critical - Ship ASAP":
            bg, color = "#ffedd5", "#7c2d12"
        elif status == "Need to Ship by Today":
            bg, color = "#fff7ed", "#9a3412"
        elif status == "Ship by Tomorrow":
            bg, color = "#e0f2fe", "#075985"
        else:
            bg, color = "#f0fdf4", "#14532d"

        html += "<tr>"
        for column in existing_columns:
            html += (
                f'<td style="border:1px solid #d9d9d9;background:{bg};color:{color};'
                f'padding:7px 10px;font-weight:bold;">{row.get(column, "")}</td>'
            )
        html += "</tr>"

    html += "</table>"
    if len(cleaned) > limit:
        html += f"<p style='font-family:Arial,sans-serif;font-size:12px;'>Showing first {limit} rows. Full details are in the attached workbook.</p>"
    return html


def make_email_html(counts: dict[str, int], pivot: pd.DataFrame, cleaned: pd.DataFrame) -> str:
    dashboard_html = make_dashboard_html(counts, pivot)
    details_html = make_details_table_html(cleaned)
    return f"""
    <html>
      <body style="font-family: Arial, sans-serif; font-size: 13px; color: #17202a;">
        <p>Hi Ops Team,</p>
        <p>Find the attached Pending Orders Report. Kindly ensure all orders are processed and shipped on time to avoid cancellations.</p>
        <p>Below are orders that require immediate attention:</p>
        {dashboard_html}
        <p style="margin-top:18px;font-weight:bold;">Pending Order Details</p>
        {details_html}
        <p>Please prioritize shipping to avoid cancellations.</p>
        <p>Thanks,<br>Graas Team</p>
      </body>
    </html>
    """


def send_email_notification(
    *,
    smtp_host: str,
    smtp_port: int,
    smtp_user: str,
    smtp_password: str,
    sender: str,
    to_emails: list[str],
    cc_emails: list[str],
    subject: str,
    html_body: str,
    attachment_bytes: bytes,
    attachment_name: str,
    use_ssl: bool,
) -> None:
    message = EmailMessage()
    message["Subject"] = subject
    message["From"] = sender
    message["To"] = ", ".join(to_emails)
    if cc_emails:
        message["Cc"] = ", ".join(cc_emails)
    message.set_content("Please view this email in an HTML-compatible client.")
    message.add_alternative(html_body, subtype="html")
    message.add_attachment(
        attachment_bytes,
        maintype="application",
        subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=attachment_name,
    )

    recipients = to_emails + cc_emails
    context = ssl.create_default_context()
    if use_ssl:
        with smtplib.SMTP_SSL(smtp_host, smtp_port, context=context) as smtp:
            if smtp_user:
                smtp.login(smtp_user, smtp_password)
            smtp.send_message(message, to_addrs=recipients)
    else:
        with smtplib.SMTP(smtp_host, smtp_port) as smtp:
            smtp.starttls(context=context)
            if smtp_user:
                smtp.login(smtp_user, smtp_password)
            smtp.send_message(message, to_addrs=recipients)


def make_slack_text(counts: dict[str, int], cleaned: pd.DataFrame, missing: pd.DataFrame) -> str:
    lines = [
        "*PUMA SG Pending Order Report*",
        f"Breached: *{counts['breached']}*",
        f"Ship Today / Critical: *{counts['today']}*",
        f"Order Status at NEW: *{counts.get('new_oms', 0)}*",
        f"Not Reflected in OMS: *{counts['not_oms']}*",
        f"Within SLA: *{counts['within']}*",
        f"Pending rows: *{len(cleaned)}*",
        f"Missing rows: *{len(missing)}*",
    ]

    urgent = cleaned[cleaned["SLA Status"].isin(["BREACHED", "Critical - Ship ASAP", "Need to Ship by Today"])]
    if not urgent.empty:
        lines.append("")
        lines.append("*Top urgent orders*")
        for _, row in urgent.head(10).iterrows():
            lines.append(
                f"- {row['Marketplace']} | {row['order_number']} | {row['custom_sku']} | "
                f"{row['SLA Status']} | OMS: {row['OMS status'] or '-'}"
            )

    if len(urgent) > 10:
        lines.append(f"...and {len(urgent) - 10} more urgent rows in the workbook.")

    return "\n".join(lines)


def send_slack_notification(webhook_url: str, text: str) -> None:
    payload = json.dumps({"text": text}).encode("utf-8")
    request = Request(
        webhook_url,
        data=payload,
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    try:
        with urlopen(request, timeout=20) as response:
            if response.status >= 400:
                raise RuntimeError(f"Slack returned HTTP {response.status}")
    except HTTPError as exc:
        raise RuntimeError(f"Slack returned HTTP {exc.code}: {exc.reason}") from exc
    except URLError as exc:
        raise RuntimeError(f"Could not reach Slack: {exc.reason}") from exc


def row_style(row: pd.Series) -> list[str]:
    status = row.get("SLA Status", "")
    if status == "BREACHED":
        color = "background-color: #ffe4e6; color: #991b1b; font-weight: 700"
    elif status == "Critical - Ship ASAP":
        color = "background-color: #ffedd5; color: #7c2d12; font-weight: 700"
    elif status == "Need to Ship by Today":
        color = "background-color: #fff7ed; color: #9a3412; font-weight: 700"
    elif status == "Ship by Tomorrow":
        color = "background-color: #e0f2fe; color: #075985; font-weight: 700"
    else:
        color = "background-color: #dcfce7; color: #14532d"
    return [color for _ in row]


def render_css() -> None:
    st.markdown(
        """
        <style>
        .block-container { padding-top: 1.4rem; }
        [data-testid="stMetricValue"] { font-size: 2rem; }
        </style>
        """,
        unsafe_allow_html=True,
    )


def main() -> None:
    render_css()

    with st.sidebar:
        st.image("https://images.graas.ai/uploads/merchant/GED/profile-1/profile-1.jpg?time=1778413961", width=92)
        st.title("PUMA SG")
        st.caption("Pending Orders Automation")
        st.divider()
        st.markdown("Upload reports, add holidays, then run the automation.")

    st.title("PUMA SG Pending Order Report")
    st.caption("Streamlit version of the marketplace, OMS, and SLA monitoring workflow.")

    with st.container(border=True):
        st.subheader("Report Uploads")
        col1, col2, col3 = st.columns(3)
        with col1:
            tc_file = st.file_uploader("TC Order Report", type=["csv", "xlsx", "xls", "zip"])
            lazada_file = st.file_uploader("Lazada Report", type=["csv", "xlsx", "xls", "zip"])
        with col2:
            shopee_file = st.file_uploader("Shopee Report", type=["csv", "xlsx", "xls", "zip"])
            zalora_file = st.file_uploader("Zalora Report", type=["csv", "xlsx", "xls", "zip"])
        with col3:
            oms_file = st.file_uploader("OMS Report", type=["csv", "xlsx", "xls", "zip"])
            holidays = st.date_input("Public Holidays", value=[], format="YYYY-MM-DD")

        run = st.button("Run Automation", type="primary", use_container_width=True)

    if run:
        try:
            if tc_file is None:
                st.error("TC Order Report is required.")
                return

            reports = {
                "tc": read_one_file(tc_file),
                "lazada": read_one_file(lazada_file),
                "shopee": read_one_file(shopee_file),
                "zalora": read_one_file(zalora_file),
                "oms": read_one_file(oms_file),
            }
            holiday_set = set(holidays or [])
            mp_map = build_marketplace_map(reports, holiday_set)
            oms_map = build_oms_map(reports["oms"])
            cleaned = process_tc(reports["tc"], mp_map, oms_map)
            missing = build_missing_rows(reports["tc"], mp_map, oms_map)
            counts, pivot = build_summary(cleaned)

            st.session_state["cleaned"] = cleaned
            st.session_state["missing"] = missing
            st.session_state["counts"] = counts
            st.session_state["pivot"] = pivot
            st.success(f"Automation complete. {len(cleaned)} pending rows generated.")
        except Exception as exc:
            st.exception(exc)

    cleaned = st.session_state.get("cleaned", pd.DataFrame(columns=CLEANED_HEADERS))
    missing = st.session_state.get("missing", pd.DataFrame(columns=MISSING_HEADERS))
    counts = st.session_state.get("counts", {"breached": 0, "today": 0, "new_oms": 0, "not_oms": 0, "within": 0})
    pivot = st.session_state.get("pivot", pd.DataFrame())

    metric_cols = st.columns(5)
    metric_cols[0].metric("Breached", counts["breached"])
    metric_cols[1].metric("Ship Today / Critical", counts["today"])
    metric_cols[2].metric("Order Status at NEW", counts.get("new_oms", 0))
    metric_cols[3].metric("Not in OMS", counts["not_oms"])
    metric_cols[4].metric("Within SLA", counts["within"])

    if not cleaned.empty:
        xlsx = make_excel(cleaned, missing, pivot)
        report_filename = f"PUMA_SG_Pending_Order_Report_{datetime.now():%Y-%m-%d}.xlsx"
        export_cols = st.columns([1, 1, 4])
        export_cols[0].download_button(
            "Download CSV",
            cleaned.to_csv(index=False).encode("utf-8-sig"),
            file_name=f"PUMA_SG_Pending_Order_Report_{datetime.now():%Y-%m-%d}.csv",
            mime="text/csv",
            use_container_width=True,
        )
        export_cols[1].download_button(
            "Download Workbook",
            xlsx,
            file_name=report_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

        with st.expander("Share Email and Slack Notification"):
            email_tab, slack_tab = st.tabs(["Email", "Slack"])

            with email_tab:
                st.caption("SMTP credentials can also be set with SMTP_HOST, SMTP_PORT, SMTP_USER, SMTP_PASSWORD, and SMTP_SENDER.")
                email_col1, email_col2 = st.columns(2)
                with email_col1:
                    to_value = st.text_area("To", value=DEFAULT_TO_EMAILS, placeholder="name@example.com, team@example.com")
                    cc_value = st.text_area("CC", value=DEFAULT_CC_EMAILS, placeholder="optional@example.com")
                    subject = st.text_input(
                        "Subject",
                        value=f"PUMA SG Pending Order Report - {datetime.now():%d-%b-%Y}",
                    )
                with email_col2:
                    preset_name = st.selectbox("SMTP preset", list(SMTP_PRESETS), index=0)
                    preset_host, preset_port, preset_ssl = SMTP_PRESETS[preset_name]
                    if preset_name.startswith("Gmail"):
                        st.info(
                            "Gmail SMTP requires a Google App Password when 2-Step Verification is enabled. "
                            "Use the 16-character app password here, not your normal Gmail password."
                        )
                    smtp_host = st.text_input("SMTP host", value=os.getenv("SMTP_HOST", preset_host))
                    smtp_port = st.number_input(
                        "SMTP port",
                        min_value=1,
                        max_value=65535,
                        value=int(os.getenv("SMTP_PORT", str(preset_port))),
                    )
                    smtp_user = st.text_input("SMTP username", value=os.getenv("SMTP_USER", ""))
                    smtp_password = st.text_input(
                        "SMTP password",
                        value=os.getenv("SMTP_PASSWORD", ""),
                        type="password",
                    )
                    sender = st.text_input("Sender email", value=os.getenv("SMTP_SENDER", smtp_user))
                    use_ssl = st.checkbox("Use SSL instead of STARTTLS", value=preset_ssl or smtp_port == 465)

                if st.button("Send Email", type="primary", use_container_width=True):
                    to_emails = split_emails(to_value)
                    cc_emails = split_emails(cc_value)
                    sender_email = sender or smtp_user
                    if not smtp_host or not sender_email or not to_emails:
                        st.error("Please fill SMTP host, sender email, and at least one To recipient. For Microsoft 365, use smtp.office365.com with port 587.")
                    else:
                        try:
                            send_email_notification(
                                smtp_host=smtp_host,
                                smtp_port=int(smtp_port),
                                smtp_user=smtp_user,
                                smtp_password=smtp_password,
                                sender=sender_email,
                                to_emails=to_emails,
                                cc_emails=cc_emails,
                                subject=subject,
                                html_body=make_email_html(counts, pivot, cleaned),
                                attachment_bytes=xlsx,
                                attachment_name=report_filename,
                                use_ssl=use_ssl,
                            )
                            st.success("Email sent.")
                        except Exception as exc:
                            st.error(f"Email failed: {exc}")

            with slack_tab:
                st.caption("Paste a Slack incoming webhook URL, or set SLACK_WEBHOOK_URL before starting Streamlit.")
                slack_webhook = st.text_input(
                    "Slack webhook URL",
                    value=os.getenv("SLACK_WEBHOOK_URL", ""),
                    type="password",
                )
                slack_text = st.text_area(
                    "Message preview",
                    value=make_slack_text(counts, cleaned, missing),
                    height=260,
                )
                if st.button("Send Slack Notification", type="primary", use_container_width=True):
                    if not slack_webhook:
                        st.error("Slack webhook URL is required.")
                    else:
                        try:
                            send_slack_notification(slack_webhook, slack_text)
                            st.success("Slack notification sent.")
                        except Exception as exc:
                            st.error(f"Slack notification failed: {exc}")

    tabs = st.tabs(["Dashboard", "TC Cleaned", "Missing Orders"])

    with tabs[0]:
        st.subheader("Email Dashboard Preview")
        if pivot.empty:
            st.info("Run automation to see the dashboard.")
        else:
            components.html(make_dashboard_html(counts, pivot), height=520, scrolling=True)
            with st.expander("Pending Order Details Preview"):
                st.dataframe(cleaned, use_container_width=True, height=420)

    with tabs[1]:
        st.subheader("TC Cleaned")
        query = st.text_input("Search cleaned orders", placeholder="Search order, SKU, marketplace, status")
        display = cleaned
        if query and not cleaned.empty:
            mask = cleaned.astype(str).apply(lambda col: col.str.contains(query, case=False, na=False)).any(axis=1)
            display = cleaned[mask]
        if display.empty:
            st.info("No cleaned rows to show.")
        else:
            st.dataframe(display.style.apply(row_style, axis=1), use_container_width=True, height=520)

    with tabs[2]:
        st.subheader("Missing Orders")
        if missing.empty:
            st.info("No missing orders to show.")
        else:
            st.dataframe(missing, use_container_width=True, height=520)


if __name__ == "__main__":
    main()
