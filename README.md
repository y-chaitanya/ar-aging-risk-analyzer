# Healthcare AR Aging & Write-Off Risk Analyzer

An Excel workbook that takes raw accounts receivable data, ages it, flags accounts at risk of write-off, and summarises the result for a finance manager. The aging and flagging are automated in VBA; the extraction query is written in DB2/AS400 SQL dialect (not executed against a live system).

Built as a portfolio project. Sample data is fictional.

---

## The problem it addresses

In healthcare billing, an invoice that sits too long stops being collectible. Claims have filing deadlines, payers have appeal windows, and balances get harder to recover the older they are. By the time an account is written off, the service has already been delivered and the cost already incurred.

The usual way to spot this is someone sorting a spreadsheet by date once a month. That works until the file gets big, or until the person doing it is on leave.

This workbook does the same job in one click, and shows how much money is sitting past 90 days.

---

## What it does

**Ages every open invoice** — calculates days outstanding from the invoice date.

**Flags write-off risk** — anything 90 days or older with a balance is marked critical and highlighted red. Anything inside 90 days with a balance is active collection. Zero balances are settled.

**Writes a summary** — four figures a manager can read without digging:

| Metric | What it means |
|---|---|
| Total Outstanding AR | Sum of all open balances |
| Critical Write-off Exposure | Balance sitting at 90+ days |
| Active Collections Pipeline | Balance still inside 90 days |
| % of AR over 90 days | The number an AR manager watches month to month |

**Checks its own arithmetic** — critical plus active is compared back to total AR. If they don't tie, the sheet says MISMATCH in red rather than quietly showing a wrong number.

**Stamps the date** — written by the macro when the aging runs, not by a formula on the sheet.

---

## How it runs

1. **Extract** — `AR_Data_Pull.sql` pulls open invoices from the billing system.
2. **Paste** — raw data goes into Sheet1, columns A to E.
3. **Click** — the Run Refresh Pipeline button on the Dashboard ages every row, applies the highlighting, and writes the summary.

---

## Files

```
├── SQL_Queries/
│   └── AR_Data_Pull.sql          DB2 / AS400 extraction query
├── VBA_Source_Code/
│   └── AR_Aging_Analyzer.bas     The macro
├── Workbook/
│   └── AR_Simulation.xlsm        Macro-enabled workbook with sample data
└── Screenshots/
    ├── 1_dashboard.png
    ├── 2_aged_data.png
    ├── 3_vba_code.png
    └── 4_sql_query.png
```

---

## Tools

SQL (IBM DB2 for i syntax) · Excel VBA · Visual Studio Code

---

## Notes on scope, so nothing here is oversold

**This is my first VBA project.** I used AI assistance to get started in a language I hadn't written before, then went through it line by line and fixed what broke. The comments in the file are mine.

Two bugs I hit, both of which shaped the final version:

- A SUMIF with the range fixed at row 5, left over from building on 5 rows of test data. After pasting in 15 rows, exposure showed $36,350 instead of $67,100 and nothing looked wrong on screen. Widening the range fixed it, but a fixed range can go stale again the next time the data grows — so the macro now adds the totals up as it loops, and the tie-out check exists to catch that class of error immediately.
- Days outstanding displaying as `03/02/1900` instead of `62`, because the cell inherited date formatting from the subtraction. Setting the number format before writing the value fixed it.

**The SQL is written for DB2 / AS400 but hasn't been run against a live system.** I wrote it to that syntax — `DAYS(CURRENT_DATE) - DAYS(INV_DATE)`, `LIBRARY.TABLE` naming — from IBM's documentation, because that's the environment it's aimed at. Table and column names would need checking against a real schema.

**The 90-day threshold is a simplification.** Real write-off policy varies by payer and contract, and timely filing limits differ between Medicare, Medi-Cal and commercial plans. A production version would take those from a lookup table rather than one constant.

**Days outstanding is not days past due.** This measures age from the invoice date. Past due would need payment terms or a payer-specific expected payment date, which the sample data doesn't carry.

**The sample data is deliberately stressed.** 70% of AR sitting past 90 days is not a realistic aging profile — the numbers were chosen to exercise the logic, not to look plausible.

---

## What I'd build next

- Power Query connection straight to the database, so nothing is pasted by hand
- Aging buckets (0–30, 31–60, 61–90, 90+) rather than a single 90-day split
- A payer breakdown, since Medicare, Medi-Cal and commercial collect on very different timelines
- Payer-specific thresholds from a lookup table instead of one constant
- A denial-reason field — in healthcare an aged balance is often an unworked denial rather than a slow payer

---

**Chaitanya Yarlagadda** 
