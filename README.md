# sql-automation-pharmacy-report

This project automates the extraction and summarization of outpatient prescription data from a hospital's MSSQL database using Python.  
It connects via ODBC, runs a multi-step SQL query using CTEs, and exports a daily prescription report to Excel for operations or pharmacy use.

---

## 🧰 Features

- Connects to MSSQL database using DSN and `pyodbc`
- Cleans and deduplicates prescription records using `ROW_NUMBER()`
- Joins patient registration and prescription data
- Aggregates prescription and medication item counts per outpatient session
- Exports output to Excel with dynamic naming (e.g., `outpatient_pharmacy_report_YYYYMMDD.xlsx`)
- Logs total row count and execution time

---

## 🧱 Tech Stack

| Tool        | Purpose                          |
|-------------|----------------------------------|
| Python      | Automation and scripting         |
| pyodbc      | Database connection              |
| pandas      | Data handling and export         |
| ExcelWriter | Save `.xlsx` output              |
| MSSQL + SQL | CTEs, joins, ranking, aggregation|

---

## 📁 Project Structure

---


---

## ▶️ How to Use

1. Ensure you have a valid ODBC DSN configured (e.g., `YourDSN`)
2. Clone this repository
3. Install dependencies:

    ```bash
    pip install pandas pyodbc openpyxl
    ```

4. Run the script:

    ```bash
    python src/generate_report.py
    ```

5. The script will:
    - Query the database  
    - Process and aggregate prescription data  
    - Export the result to Excel in your home directory (`~/sql_reports/`)  
    - Log the duration and row count

---

## 📊 Sample Output

| Visit_Date | Session   | Prescription_Count | Item_Count |
|------------|-----------|--------------------|------------|
| 1120101    | Morning   | 125                | 378        |
| 1120101    | Afternoon | 92                 | 287        |
| ...        | ...       | ...                | ...        |

> All dates follow the Minguo calendar format (`1120101` = `2023-01-01`)

---

## 🗂 Data Schema (Anonymized)

- `pharmacy_orders`: Medication prescription records  
- `outpatient_visits`: Patient registration metadata

> ⚠️ All table and database names in this repo are anonymized for demonstration purposes.
> No actual patient-identifiable data is used or exported.
