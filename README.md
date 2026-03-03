import requests
from bs4 import BeautifulSoup
import pandas as pd
from datetime import datetime
from sqlalchemy import create_engine, text
import urllib
import os

# === DATABASE CONFIGURATION ===
def get_engine():
    DRIVER = "ODBC Driver 17 for SQL Server"
    SERVER = "ARAPL-LP-20"
    DB = "ARAPL_Configuration"
    param_string = f"DRIVER={{{DRIVER}}};SERVER={SERVER};DATABASE={DB};Trusted_Connection=yes;"
    return create_engine(f"mssql+pyodbc:///?odbc_connect={urllib.parse.quote_plus(param_string)}")

engine = get_engine()

# === SCRAPE ARTEMIS WEBSITE ===
print("🔍 Scraping Artemis CAT Bond losses...")

url = "https://www.artemis.bm/cat-bond-losses/"
r = requests.get(url)
soup = BeautifulSoup(r.text, "lxml")

# Extract table headers
table = soup.find("table", {"id": "tablepress-2"})
headers = [i.text.strip() for i in table.find("thead").find_all("th")]

# Map headers to SQL column names
column_mapping = {
    "Cat bond": "CATBond",
    "Sponsor": "Sponsor",
    "Orig. size": "OrigSize",
    "Cause of loss": "CauseOfLoss",
    "Loss amount": "LossAmount",
    "~Time to payment": "TimeToPayment",
    "Date of loss": "DateOfLoss",
    "url": "URL",
}

# Extract data rows
deals = []
for row in table.find("tbody").find_all("tr"):
    cells = row.find_all("td")
    d = {}
    for i, c in enumerate(cells):
        d[column_mapping[headers[i]]] = c.text.strip()
    d["URL"] = row.find("a")["href"]
    deals.append(d)

df = pd.DataFrame(deals)
print(f"✅ Scraped {len(df)} records from Artemis.")

# === LOAD EXISTING DATA FROM SQL ===
df_existing = pd.read_sql_query("""
    SELECT ID, CATBond, Sponsor, OrigSize, CauseOfLoss, LossAmount, TimeToPayment, DateOfLoss, URL 
    FROM ARAPL_Configuration.dbo.CATBond_ArtemisLossInfo 
    WHERE isDeleted = 0
""", con=engine)
print(f"📂 Loaded {len(df_existing)} existing records from database.")

# === DETECT NEW AND UPDATED RECORDS ===
df_new = df.merge(df_existing, how="left", on=["CATBond"], suffixes=["", "_old"])
df_new = df_new[df_new.ID.isna()][list(column_mapping.values())]

df_updated = df.merge(df_existing, how="left", on=list(column_mapping.values()))
df_updated = df_updated[df_updated.ID.isna()]
df_updated = df_updated[~df_updated.CATBond.isin(df_new.CATBond)]

# === INSERT NEW/UPDATED RECORDS ===
# def insert_without_id(df_insert):
#     if not df_insert.empty:
#         df_insert.to_sql("CATBond_ArtemisLossInfo", con=engine, index=False, if_exists="append")
#         print(f"🆕 Inserted {len(df_insert)} new/updated records into SQL.")
#     else:
#         print("⚪ No new or updated records found for insertion.")

# insert_without_id(df_new)
# insert_without_id(df_updated)


# ====================================
# Insert new data into SQL
# ====================================
def insert_without_id(df_updated):
    if df_updated.empty:
        print("✅ No new records to insert.")
        return

    # Drop ID column if exists (prevents IDENTITY_INSERT error)
    if "ID" in df_updated.columns:
        df_updated = df_updated.drop(columns=["ID"])

    try:
        df_updated.to_sql(
            "CATBond_ArtemisLossInfo",
            con=engine,
            index=False,
            if_exists="append"
        )
        print(f"✅ Inserted {len(df_updated)} new records into SQL.")
    except Exception as e:
        print(f"❌ SQL Insert Error: {e}")

insert_without_id(df_updated)

# === CLEANUP QUERIES ===
with engine.begin() as conn:
    conn.execute(text("""
        UPDATE a SET isDeleted = 1 
        FROM ARAPL_Configuration.dbo.CATBond_ArtemisLossInfo a
        LEFT JOIN (
            SELECT DISTINCT CATBond FROM ARAPL_Configuration.dbo.CATBond_ArtemisLossInfo WHERE isDeleted = 0
        ) b ON a.CATBond = b.CATBond
        WHERE b.CATBond IS NULL AND a.isDeleted = 0;
    """))
    conn.execute(text("""
        UPDATE ARAPL_Configuration.dbo.CATBond_ArtemisLossInfo 
        SET isDeleted = 0 
        WHERE isDeleted IS NULL;
    """))

print("🧾 SQL data cleanup completed.")

# === EXPORT TO EXCEL ===
if not df_new.empty or not df_updated.empty:
    df_combined = pd.concat([df_new, df_updated])
    today = datetime.today().strftime("%Y-%m-%d")
    output_folder = r"Excel_Exports"
    os.makedirs(output_folder, exist_ok=True)

    excel_path = fr"{output_folder}\CATBond_ArtemisLoss_Updates_{today}.xlsx"
    df_combined.to_excel(excel_path, index=False)
    print(f"📊 Exported updated data to Excel: {excel_path}")
else:
    print("⚪ No new data to export to Excel.")

print("🎯 Process completed successfully.")
