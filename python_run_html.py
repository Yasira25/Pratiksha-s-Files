import pandas as pd
import os
import base64
import requests
from datetime import datetime
from textwrap import wrap
import plotly.graph_objs as go
from openpyxl import Workbook
# from msal import PublicClientApplication, SerializableTokenCache
from db_connection import engine

# # === CONFIGURATION ===

# CLIENT_ID = "5b4f8359-9e4f-45cc-8708-5453b1d5c9aa"
# TENANT_ID = "a3d39074-5bc8-4cdf-8161-af015c12ae47"
# AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
# SCOPES = ["Mail.Send", "User.Read"]

# SENDER_EMAIL = "kunal@advancedriskanalytics.com"
# RECIPIENTS = [
#     "ashish.belambe@advancedriskanalytics.com"

# ]
# CC_RECIPIENTS = [
#     "sindhu@advancedriskanalytics.com",
#     "kunal@advancedriskanalytics.com"
# ]

today = datetime.today().strftime("%Y-%m-%d")
ATTACHMENT_PATH = fr"C:\Users\Admin\Downloads\Weekly_SwissRe_Report_{today}.xlsx"

# === FETCH DATA ===
query = f"EXEC ARAPL_Configuration.dbo.sp_getSwissRe_weekly_changes '{today}'"
df = pd.read_sql_query(query, engine)

# === STATUS COUNT ===
status_counts = df["Status"].value_counts()
CouponChanged = status_counts.get("Coupon Changed", 0)
PriceChanged = status_counts.get("Price Changed", 0)
New = status_counts.get("New", 0)
Expired = status_counts.get("Expired", 0)

# === WRITE TO EXCEL ===
wb = Workbook()
ws = wb.active
ws.title = "SwissRe Changes"
ws.append(df.columns.tolist())
for _, row in df.iterrows():
    ws.append(row.tolist())
wb.save(ATTACHMENT_PATH)

# === GENERATE PLOTLY CHART ===
def split_text(text, line_width):
    return "<br>".join(wrap(text, width=line_width))

x_labels = df.apply(lambda row: split_text(f"{row.Bond}     {row.Status}", 15), axis=1)
y_values = df["CurrMidPrice"] - df["PreviousMidPrice"]
fig = go.Figure(data=[go.Bar(x=x_labels, y=y_values, name="Weekly Price Change")])
fig.update_layout(title="Weekly Price Change", height=500)

chart_bytes = fig.to_image(format="png")
encoded_image = base64.b64encode(chart_bytes).decode("utf-8")
img_tag = f'<img src="data:image/png;base64,{encoded_image}" style="max-width:1000px;"><br>'

# # === EMAIL CONTENT ===
# EMAIL_BODY = f"""
# <p>Hi team,</p>
# <p>Please find attached the changes in Swiss Re data for the last week.</p>
# <ul>
# <li><b>Bonds with &gt; 3% price change:</b> {PriceChanged}</li>
# <li><b>Bonds with Coupon change:</b> {CouponChanged}</li>
# <li><b>New Bonds added last week:</b> {New}</li>
# <li><b>Bonds expired/redeemed:</b> {Expired}</li>
# </ul>
# {img_tag}
# <p>Regards,<br>Mandar Belambe</p>
# """

# EMAIL_SUBJECT = f"Weekly SwissRe Changes - {today}"

# # === MSAL DEVICE AUTH ===
# cache_path = "msal_token_cache.bin"
# cache = SerializableTokenCache()

# if os.path.exists(cache_path):
#     cache.deserialize(open(cache_path, "r").read())

# app = PublicClientApplication(CLIENT_ID, authority=AUTHORITY, token_cache=cache)
# accounts = app.get_accounts()
# if accounts:
#     result = app.acquire_token_silent(SCOPES, account=accounts[0])
# else:
#     flow = app.initiate_device_flow(scopes=SCOPES)
#     print(flow["message"])
#     result = app.acquire_token_by_device_flow(flow)

# with open(cache_path, "w") as f:
#     f.write(cache.serialize())

# if "access_token" not in result:
#     raise Exception(f"Auth failed: {result.get('error_description')}")

# # === EMAIL SEND ===
# with open(ATTACHMENT_PATH, "rb") as f:
#     attachment_bytes = f.read()

# message = {
#     "message": {
#         "subject": EMAIL_SUBJECT,
#         "body": {
#             "contentType": "HTML",
#             "content": EMAIL_BODY
#         },
#         "toRecipients": [{"emailAddress": {"address": addr}} for addr in RECIPIENTS],
#         "ccRecipients": [{"emailAddress": {"address": addr}} for addr in CC_RECIPIENTS],
#         "attachments": [{
#             "@odata.type": "#microsoft.graph.fileAttachment",
#             "name": os.path.basename(ATTACHMENT_PATH),
#             "contentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
#             "contentBytes": base64.b64encode(attachment_bytes).decode("utf-8")
#         }]
#     },
#     "saveToSentItems": "true"
# }

# res = requests.post(
#     "https://graph.microsoft.com/v1.0/me/sendMail",
#     headers={
#         "Authorization": f"Bearer {result['access_token']}",
#         "Content-Type": "application/json"
#     },
#     json=message
# )

# if res.status_code == 202:
#     print("Email sent successfully.")
# else:
#     print(f"Email failed: {res.status_code} - {res.text}")

# # === CLEANUP ===
# os.remove(ATTACHMENT_PATH)
