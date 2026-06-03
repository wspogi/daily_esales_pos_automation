/*
============================================================
Daily E-Sales Report Query
============================================================
IMPORTANT:
This query is the ONLY part we still need to align with the
actual POS database tables/views.

The Python script expects these exact column aliases:

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

Parameters supplied by Python:
    ? = report_start_date
    ? = report_end_date_exclusive

Date logic:
    sales date >= report_start_date
    sales date <  report_end_date_exclusive

Replace the sample SELECT below with the correct POS SQL query
or view once we confirm the table names.
============================================================
*/

/*
SAMPLE SHAPE ONLY - palitan natin ito kapag alam na natin exact tables.

SELECT
    MAX(company.TIN) AS TIN,
    terminal.TerminalNo AS TerminalNo,
    terminal.POSSerialNo AS POSSerialNo,
    terminal.MIN AS MIN,
    terminal.PermitNo AS PermitNo,
    MAX(sales.InvoiceNo) AS LastInvoiceNo,
    SUM(sales.NetSalesWithVAT) AS NetSalesWithVAT,
    SUM(sales.ZeroRated) AS ZeroRated,
    SUM(sales.VATExempt) AS VATExempt,
    SUM(sales.Vatable) AS Vatable,
    SUM(sales.TotalSalesNetOfVAT) AS TotalSalesNetOfVAT,
    SUM(sales.OutputVAT) AS OutputVAT
FROM SalesTable sales
LEFT JOIN TerminalTable terminal
    ON sales.TerminalNo = terminal.TerminalNo
CROSS JOIN CompanyTable company
WHERE sales.SalesDate >= ?
  AND sales.SalesDate < ?
GROUP BY
    terminal.TerminalNo,
    terminal.POSSerialNo,
    terminal.MIN,
    terminal.PermitNo
ORDER BY
    terminal.TerminalNo;
*/

-- TEMPORARY TEST QUERY ONLY.
-- This lets you test Excel generation even before final POS query is ready.
-- Delete/replace this block once we plug the real query.
SELECT
    CAST('' AS varchar(50)) AS TIN,
    CAST('01' AS varchar(20)) AS TerminalNo,
    CAST('0712202402086' AS varchar(50)) AS POSSerialNo,
    CAST('24080816245370261' AS varchar(50)) AS MIN,
    CAST('FP082024-067-0461594-00000' AS varchar(50)) AS PermitNo,
    CAST('01001352' AS varchar(50)) AS LastInvoiceNo,
    CAST(456699.45 AS decimal(18,2)) AS NetSalesWithVAT,
    CAST(0 AS decimal(18,2)) AS ZeroRated,
    CAST(0 AS decimal(18,2)) AS VATExempt,
    CAST(407767.38 AS decimal(18,2)) AS Vatable,
    CAST(407767.38 AS decimal(18,2)) AS TotalSalesNetOfVAT,
    CAST(48932.07 AS decimal(18,2)) AS OutputVAT;
