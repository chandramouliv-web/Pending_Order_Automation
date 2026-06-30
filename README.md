# PUMA SG Pending Orders App

This project now includes a Streamlit app converted from the Google Apps Script upload workflow.

## Run the Streamlit app

Install dependencies:

```powershell
pip install -r requirements.txt
```

Start the app:

```powershell
streamlit run app.py
```

## What it does

- Uploads TC, Lazada, Shopee, Zalora, and OMS reports.
- Supports CSV, XLS, XLSX, and ZIP files containing CSV/XLS/XLSX.
- Applies public holidays to Lazada, Shopee, and Zalora SLA calculations.
- Generates the `TC_Cleaned` pending-order table.
- Generates a missing-orders table.
- Shows the email-style SLA dashboard preview.
- Exports the report as CSV or XLSX.
- Sends the generated workbook by email through SMTP.
- Posts a Slack summary through an incoming webhook.

## Notes

The older standalone browser version is still available in `index.html`, but `app.py` is the Streamlit version to use going forward.

## Notifications

After running automation, open **Share Email and Slack Notification**.

For email, enter SMTP host, port, username, password, sender, recipients, and subject. You can also set these environment variables before starting Streamlit:

- `SMTP_HOST`
- `SMTP_PORT`
- `SMTP_USER`
- `SMTP_PASSWORD`
- `SMTP_SENDER`

For Slack, paste an incoming webhook URL or set:

- `SLACK_WEBHOOK_URL`
