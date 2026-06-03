# Daily E-Sales POS Automation

Python automation for generating the **Daily E-Sales Report** from the POS SQL Server database and exporting it to Excel.

The Excel layout is based on the provided `APR.xls` template.

---

## 1. Purpose

This automation will:

1. Connect to the POS SQL Server database.
2. Get sales totals for the selected report period.
3. Use static POS/BIR details from `.env`.
4. Generate an Excel file using the Daily E-Sales format.
5. Save logs in the `logs` folder.
6. Save output reports in the `output` folder.

---

## 2. Project Files

```text
daily_esales_pos_automation/
│
├── main.py
├── test_connection.py
├── requirements.txt
├── .env.example
├── .gitignore
│
├── INSTALL_PYTHON.bat
├── INSTALL_REQUIREMENTS.bat
├── TEST_CONNECTION.bat
├── RUN_DAILY_E_SALES_REPORT.bat
├── SETUP_TASK_SCHEDULER.bat
│
├── sql/
│   └── query_daily_esales.sql
│
├── logs/
│   └── .gitkeep
│
└── output/
    └── .gitkeep