import os
import pyodbc
import pandas as pd
import time
from datetime import datetime

# === Start timing ===
start_time = time.time()
print("✅ Query started:", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

# === Establish database connection ===
conn = pyodbc.connect("DSN=YourDSN;DATABASE=YourDataWarehouse;Trusted_Connection=yes")

# === SQL Query ===
sql = """
WITH Prescriptions_Clean AS (
    SELECT *
    FROM [YourDataWarehouse].[dbo].[pharmacy_orders]
    WHERE BRANCH = 'G'
      AND PRINT_FLAG = 'Y'
      AND ODR_CODE <> 'DELETE'
      AND PHA_DATE BETWEEN '1120101' AND '1130630'
),
Prescriptions_Ranked AS (
    SELECT *,
           ROW_NUMBER() OVER (
               PARTITION BY PHA_DATE, PHA_NUM, ODR_SEQ
               ORDER BY PHA_TIME DESC
           ) AS row_num
    FROM Prescriptions_Clean
),
Prescriptions_Final AS (
    SELECT * FROM Prescriptions_Ranked WHERE row_num = 1
),
Visits_Clean AS (
    SELECT *
    FROM [YourDataWarehouse].[dbo].[outpatient_visits]
    WHERE BRANCH = 'G'
      AND OPDER_SW = '0'
      AND CANCEL_FLAG = 'N'
      AND REG_DATE BETWEEN '1120101' AND '1130630'
),
Joined AS (
    SELECT
        p.PHA_NUM,
        p.ODR_CODE,
        v.REG_DATE,
        v.NOON_CODE,
        v.PAT_NO,
        v.PAT_SEQ
    FROM Prescriptions_Final p
    INNER JOIN Visits_Clean v
        ON p.PAT_NO = v.PAT_NO AND p.PAT_SEQ = v.PAT_SEQ
),
PerPatientStats AS (
    SELECT
        REG_DATE,
        NOON_CODE,
        PAT_NO,
        PAT_SEQ,
        COUNT(DISTINCT PHA_NUM) AS PHA_NUM,
        COUNT(ODR_CODE) AS ODR_NUM
    FROM Joined
    GROUP BY REG_DATE, NOON_CODE, PAT_NO, PAT_SEQ
),
FinalStats AS (
    SELECT
        REG_DATE AS [Visit_Date],
        NOON_CODE,
        SUM(PHA_NUM) AS [Prescription_Count],
        SUM(ODR_NUM) AS [Item_Count]
    FROM PerPatientStats
    WHERE NOON_CODE <> '3'
    GROUP BY REG_DATE, NOON_CODE
)
SELECT
    [Visit_Date],
    CASE NOON_CODE
        WHEN '1' THEN 'Morning'
        WHEN '2' THEN 'Afternoon'
        ELSE 'Other'
    END AS [Session],
    [Prescription_Count],
    [Item_Count]
FROM FinalStats
ORDER BY [Visit_Date], [Session]
"""

# === Execute query ===
df = pd.read_sql(sql, conn)
row_count = len(df)

# === Generate filename and export path ===
today_str = datetime.today().strftime("%Y%m%d")
filename = f"outpatient_pharmacy_report_{today_str}.xlsx"
save_dir = os.path.expanduser("~/sql_reports")
os.makedirs(save_dir, exist_ok=True)
output_path = os.path.join(save_dir, filename)

# === Export to Excel ===
df.to_excel(output_path, index=False)

# === Completion summary ===
duration = round(time.time() - start_time, 2)
print(f"\n📊 Query complete — {row_count} rows")
print(f"📄 Excel saved to: {output_path}")
print(f"🕒 Duration: {duration} seconds")
