# Daily E-Sales Report POS Automation

Python automation para gumawa ng **Daily E-Sales Report** Excel file based sa uploaded `APR.xls` template.

## 1. Template Analysis

Uploaded template: `APR.xls`

Detected layout:

| Row | Purpose |
|---|---|
| Row 1 | Company name |
| Row 2 | Operated by |
| Row 3 | Report title: Daily E-Sales Report |
| Row 4 | Period covered |
| Row 7 | Table headers |
| Row 8 onward | POS terminal summary rows |
| Last row | Grand Total |

Expected report columns:

1. TIN
2. Terminal No.
3. POS Serial No.
4. M.I.N
5. Permit No.
6. Last Invoice No.
7. Net Sales (w/VAT)
8. Zero Rated
9. VAT Exempt
10. Vatable
11. Total Sales (Net of VAT)
12. Output VAT

## 2. Important Note

The Python automation is ready structurally, but the SQL query still needs to be aligned with the exact POS database tables/views.

Edit this file:

```text
sql/query_daily_esales.sql
```

The query must return these aliases exactly:

```text
TIN
TerminalNo
POSSerialNo
MIN
PermitNo
LastInvoiceNo
NetSalesWithVAT
ZeroRated
VATExempt
Vatable
TotalSalesNetOfVAT
OutputVAT
```

## 3. Install Requirements

### Step 1: Install Python

Run:

```text
INSTALL_PYTHON.bat
```

Download Python and make sure checked:

```text
Add python.exe to PATH
```

Then reopen CMD and check:

```cmd
py --version
```

### Step 2: Install ODBC Driver 17

Install Microsoft ODBC Driver 17 for SQL Server if wala pa sa PC.

### Step 3: Install Python packages

Run:

```text
INSTALL_REQUIREMENTS.bat
```

## 4. Setup .env

Copy:

```text
.env.example
```

Rename copy to:

```text
.env
```

Update values:

```env
SQL_SERVER=192.168.254.109
SQL_DATABASE=VQPBOS
SQL_USERNAME=your_sql_username
SQL_PASSWORD=your_sql_password
SQL_DRIVER=ODBC Driver 17 for SQL Server

COMPANY_NAME=STA. MARIA CONSTRUCTION
OPERATED_BY=operated by: JHON MOISES R. BRIAGAS
REPORT_TITLE=Daily E-Sales Report
POS_BRANCH=DAET
```

## 5. Test Connection

Run:

```text
TEST_CONNECTION.bat
```

Expected result:

```text
Connection OK. TestValue: 1
```

## 6. Test Report Generation

Run:

```text
RUN_DAILY_E_SALES_REPORT.bat
```

Default behavior:

- If today is June 2026, it generates May 2026 report.
- It always uses previous month when no month argument is provided.

Manual test for a specific month:

```cmd
py main.py 2026-04
```

Output will be saved to:

```text
output/Daily_E-Sales_Report_BRANCH_YYYY-MM.xlsx
```

Example:

```text
output/Daily_E-Sales_Report_DAET_2026-04.xlsx
```

## 7. Logs

Logs are saved here:

```text
logs/daily_esales_YYYYMMDD.log
```

Check logs kapag may error sa SQL connection, query, or Excel generation.

## 8. Setup Task Scheduler

Run as Administrator:

```text
SETUP_TASK_SCHEDULER.bat
```

This creates a monthly task:

```text
Task Name: Daily E-Sales Report POS Automation
Schedule: Every 1st day of the month at 08:00 AM
```

Since the script defaults to previous month, every 1st day it will generate last month’s report.

## 9. Git Setup

```cmd
git init
git add .
git commit -m "Initial Daily E-Sales POS automation setup"
```

If may GitHub repo na:

```cmd
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/YOUR_REPO_NAME.git
git push -u origin main
```

## 10. Folders

```text
Daily E-Sales Report POS Automation/
├─ main.py
├─ test_connection.py
├─ requirements.txt
├─ .env.example
├─ .gitignore
├─ INSTALL_PYTHON.bat
├─ INSTALL_REQUIREMENTS.bat
├─ TEST_CONNECTION.bat
├─ RUN_DAILY_E_SALES_REPORT.bat
├─ SETUP_TASK_SCHEDULER.bat
├─ sql/
│  └─ query_daily_esales.sql
├─ logs/
└─ output/
```
