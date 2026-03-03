import pandas as pd
import os
import base64
import requests
from datetime import datetime
from sqlalchemy import text, create_engine
from db_connection import engine
from msal import PublicClientApplication, SerializableTokenCache
import urllib

# === DATABASE CONNECTION ===
def get_engine():
    DRIVER = "ODBC Driver 17 for SQL Server"
    SERVER = "ARAPL-LP-20"        # change if different
    DB = "ARAPL_Configuration"
    param_string = f"DRIVER={{{DRIVER}}};SERVER={SERVER};DATABASE={DB};Trusted_Connection=yes;"
    return create_engine(f"mssql+pyodbc:///?odbc_connect={urllib.parse.quote_plus(param_string)}")

engine = get_engine()


# # === EMAIL CONFIG ===
# CLIENT_ID = "8ad7d1cd-0149-453b-96d0-f1cf3bffd444"
# TENANT_ID = "a3d39074-5bc8-4cdf-8161-af015c12ae47"
# AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
# SCOPES = ["Mail.Send", "User.Read"]

# SENDER_EMAIL = "pratiksha@advancedriskanalytics.com"
# RECIPIENTS = [
#     "pratiksha@advancedriskanalytics.com",
#     "sindhu@advancedriskanalytics.com",
#     "ashish.belambe@advancedriskanalytics.com"
# ]

# EMAIL_BODY = "Attached is the CAT Bond Broker data from the last 45 days."
# ATTACHMENT_NAME = "New CAT Bonds in BrokerData.xlsx"
# ATTACHMENT_PATH = fr"C:\Users\Admin\Downloads\{ATTACHMENT_NAME}"

# === REPORT GENERATION ===
tables = {}
broker_tables = [
    "BrokerAK",
    "BrokerSwissRe",
    "BrokerRBC"
]

for broker_table in broker_tables:
    query = f"""
        SELECT *
        FROM (
            SELECT ROW_NUMBER() OVER (PARTITION BY a.CUSIP ORDER BY a.DateOfIndication desc) AS RankID, a.*
            FROM {broker_table} a
        ) b
        WHERE b.RankID = 1
          AND DATEDIFF(day, b.DateOfIndication, GETDATE()) < 100
        ORDER BY b.DateOfIndication DESC
    """
#  changed----- AND DATEDIFF(dd, DateOfIndication, GETDATE()) < 45   dd--day b.

    df = pd.read_sql_query(text(query), con=engine)
    tables[broker_table] = df

# === Save Excel Report ===
today = datetime.today().strftime("%Y-%m-%d")
output_path = rf"C:\Bonds_Portfolio\Artemis_Outputs\New30Days_CATBonds_update_{today}.xlsx"

with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    for broker_name, df in tables.items():
        df.to_excel(writer, sheet_name=broker_name[:31], index=False)

print(f"✅ Excel file saved successfully:\n{output_path}")        

# # === MSAL DEVICE CODE AUTH (TOKEN CACHE) ===
# cache_path = "msal_token_cache.bin"
# cache = SerializableTokenCache()
# if os.path.exists(cache_path):
#     cache.deserialize(open(cache_path, "r").read())

# app = PublicClientApplication(CLIENT_ID, authority=AUTHORITY, token_cache=cache)
# accounts = app.get_accounts()
# if accounts:
#     print("Using cached credentials...")
#     result = app.acquire_token_silent(SCOPES, account=accounts[0])
# else:
#     print("Requesting login via device code...")
#     flow = app.initiate_device_flow(scopes=SCOPES)
#     if "user_code" not in flow:
#         raise Exception("Device code flow failed.")
#     print(flow["message"])
#     result = app.acquire_token_by_device_flow(flow)

# with open(cache_path, "w") as f:
#     f.write(cache.serialize())

# if not result or "access_token" not in result:
#     error_message = result.get('error_description') if isinstance(result, dict) else "No response received"
#     raise Exception(f"Authentication failed: {error_message}")

# # === PREPARE EMAIL ===
# with open(ATTACHMENT_PATH, "rb") as f:
#     content_bytes = f.read()

# attachment_name = os.path.basename(ATTACHMENT_PATH)
# timestamp = datetime.now().strftime("%d-%b-%Y %H:%M")
# EMAIL_SUBJECT = f"New CAT Bonds in BrokerData in last 45 days – {timestamp}"

# message = {
#     "message": {
#         "subject": EMAIL_SUBJECT,
#         "body": {
#             "contentType": "Text",
#             "content": EMAIL_BODY
#         },
#         "toRecipients": [{"emailAddress": {"address": email}} for email in RECIPIENTS],
#         "attachments": [
#             {
#                 "@odata.type": "#microsoft.graph.fileAttachment",
#                 "name": attachment_name,
#                 "contentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
#                 "contentBytes": base64.b64encode(content_bytes).decode("utf-8")
#             }
#         ]
#     },
#     "saveToSentItems": "true"
# }

# # === SEND EMAIL ===
# GRAPH_API_ENDPOINT = "https://graph.microsoft.com/v1.0/me/sendMail"
# response = requests.post(
#     GRAPH_API_ENDPOINT,
#     headers={
#         "Authorization": f"Bearer {result['access_token']}",
#         "Content-Type": "application/json"
#     },
#     json=message
# )

# if response.status_code == 202:
#     print(" Email sent successfully.")
# else:
#     print(f"Failed to send email: {response.status_code} - {response.text}")

# === CLEANUP ===
# os.remove(ATTACHMENT_PATH)
