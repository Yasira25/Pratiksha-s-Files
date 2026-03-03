import urllib
import pandas as pd
from datetime import datetime
from sqlalchemy import create_engine

# === DATABASE CONNECTION ===
def get_engine():
    DRIVER = "ODBC Driver 17 for SQL Server"
    SERVER = "ARAPL-LP-20"        # change if different
    DB = "ARAPL_Configuration"
    param_string = f"DRIVER={{{DRIVER}}};SERVER={SERVER};DATABASE={DB};Trusted_Connection=yes;"
    return create_engine(f"mssql+pyodbc:///?odbc_connect={urllib.parse.quote_plus(param_string)}")

engine = get_engine()

# === QUERY 1: Active CAT Bonds ===
query1 = """
SELECT
    sponsorlookup.SponsorID,
    sponsorlookup.Sponsor,
    ilsdata.ILSInstrumentID,
    ilsdata.ILSName,
    ilsterms.InceptionDate,
    ilsterms.ExpirationDate,
    ilsterms.PrincipalAmount,
    ilsterms.PercentAnnualSpread
FROM ilsdata
INNER JOIN sponsorlookup ON sponsorlookup.SponsorID = ilsdata.SponsorID
    AND ilsdata.Status IN ('Bound', 'On Hold')
INNER JOIN ILSTerms ON 
    ilsdata.ILSInstrumentID = ilsterms.ILSInstrumentID
    AND ilsterms.InuringSequenceNumber = '1'
INNER JOIN ilsactivitymonitor ON 
    ilsterms.ILSInstrumentID = ilsactivitymonitor.AcctGrpId
    AND ilsterms.AnalysisSID = ilsactivitymonitor.ActivityID
    AND ilsactivitymonitor.IsDefaultAnalysis = 1
    AND ilsactivitymonitor.AnalysisType = 12
    AND ilsactivitymonitor.Vendor = 'AIR'
WHERE
    ilsdata.IsILSDeleted = 0
    AND sponsorlookup.IsSponsorDeleted = 0
    AND ilsdata.Security_or_Reinsurance_Type = 'CAT Bond'
    AND sponsorlookup.Sponsor NOT LIKE ('%demo%')
    AND sponsorlookup.Sponsor NOT LIKE ('%test%')
    AND ilsdata.ILSName NOT LIKE('%test%')
ORDER BY ilsterms.InceptionDate;
"""

# === QUERY 2: Missing CUSIP IDs ===
query2 = """
SELECT * 
FROM ILSData 
WHERE 
    CUSIP_ID IS NULL 
    AND IsILSDeleted = 0
    AND Security_or_Reinsurance_Type = 'CAT Bond'
    AND ILSName NOT LIKE('%test%');
"""

# === QUERY 3: Principal mismatch (SwissRe) ===
query3 = """
SELECT
    sponsorlookup.SponsorID,
    sponsorlookup.Sponsor,
    ilsdata.ILSInstrumentID,
    ilsdata.ILSName,
    ilsterms.InceptionDate,
    ilsterms.ExpirationDate,
    ilsterms.PrincipalAmount,
    ilsterms.PercentAnnualSpread,
    t1.NotionalOutstanding
FROM ilsdata
INNER JOIN sponsorlookup ON sponsorlookup.SponsorID = ilsdata.SponsorID
    AND ilsdata.Status IN ('Bound', 'On Hold')
INNER JOIN ILSTerms ON 
    ilsdata.ILSInstrumentID = ilsterms.ILSInstrumentID
    AND ilsterms.InuringSequenceNumber = '1'
INNER JOIN ilsactivitymonitor ON 
    ilsterms.ILSInstrumentID = ilsactivitymonitor.AcctGrpId
    AND ilsterms.AnalysisSID = ilsactivitymonitor.ActivityID
    AND ilsactivitymonitor.IsDefaultAnalysis = 1
    AND ilsactivitymonitor.AnalysisType = 12
    AND ilsactivitymonitor.Vendor = 'AIR'
INNER JOIN (
    SELECT * FROM BrokerSwissRe
    WHERE DateofIndication = (SELECT MAX(DateofIndication) FROM BrokerSwissRe)
) t1 ON t1.CUSIP = ilsdata.CUSIP_ID
WHERE
    ilsdata.IsILSDeleted = 0
    AND sponsorlookup.IsSponsorDeleted = 0
    AND ilsdata.Security_or_Reinsurance_Type = 'CAT Bond'
    AND sponsorlookup.Sponsor NOT LIKE ('%demo%')
    AND sponsorlookup.Sponsor NOT LIKE ('%test%')
    AND ilsdata.ILSName NOT LIKE('%test%')
    AND t1.NotionalOutstanding <> ilsterms.PrincipalAmount;
"""

# === QUERY 4: Spread mismatch (Bloomberg) ===
query4 = """
SELECT
    sponsorlookup.SponsorID,
    sponsorlookup.Sponsor,
    ilsdata.ILSInstrumentID,
    ilsdata.ILSName,
    ilsterms.InceptionDate,
    ilsterms.ExpirationDate,
    ilsterms.PrincipalAmount,
    ilsterms.PercentAnnualSpread,
    t1.FLT_SPREAD
FROM ilsdata
INNER JOIN sponsorlookup ON sponsorlookup.SponsorID = ilsdata.SponsorID
    AND ilsdata.Status IN ('Bound', 'On Hold')
INNER JOIN ILSTerms ON 
    ilsdata.ILSInstrumentID = ilsterms.ILSInstrumentID
    AND ilsterms.InuringSequenceNumber = '1'
INNER JOIN ilsactivitymonitor ON 
    ilsterms.ILSInstrumentID = ilsactivitymonitor.AcctGrpId
    AND ilsterms.AnalysisSID = ilsactivitymonitor.ActivityID
    AND ilsactivitymonitor.IsDefaultAnalysis = 1
    AND ilsactivitymonitor.AnalysisType = 12
    AND ilsactivitymonitor.Vendor = 'AIR'
INNER JOIN (
    SELECT * FROM BloombergCATBondData
    WHERE DateofIndication = (SELECT MAX(DateofIndication) FROM BloombergCATBondData)
) t1 ON t1.CUSIP = ilsdata.CUSIP_ID
WHERE
    ilsdata.IsILSDeleted = 0
    AND sponsorlookup.IsSponsorDeleted = 0
    AND ilsdata.Security_or_Reinsurance_Type = 'CAT Bond'
    AND sponsorlookup.Sponsor NOT LIKE ('%demo%')
    AND sponsorlookup.Sponsor NOT LIKE ('%test%')
    AND ilsdata.ILSName NOT LIKE('%test%')
    AND t1.FLT_SPREAD <> ilsterms.PercentAnnualSpread;
"""

# === EXECUTE & EXPORT ===
today = datetime.today().strftime("%Y-%m-%d")
output_path = rf"C:\Bonds_Portfolio\Artemis_Outputs\Email_Forwarding\Crown_CATBonds_update_{today}.xlsx"
try:
    with engine.connect() as conn:
        df1 = pd.read_sql(query1, conn)
        df2 = pd.read_sql(query2, conn)
        df3 = pd.read_sql(query3, conn)
        df4 = pd.read_sql(query4, conn)

    # === Save to Excel ===
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df1.to_excel(writer, index=False, sheet_name="ActiveCAT")
        df2.to_excel(writer, index=False, sheet_name="MissingCUSIP")
        df3.to_excel(writer, index=False, sheet_name="PrincipalMismatch")
        df4.to_excel(writer, index=False, sheet_name="SpreadMismatch")

    print(f"✅ Excel file saved successfully:\n{output_path}")

except Exception as e:
    print("❌ Error occurred while exporting:", e)
