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


def make_email_html(counts: dict[str, int], pivot: pd.DataFrame) -> str:
    dashboard_html = make_dashboard_html(counts, pivot)
