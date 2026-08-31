-- ===================================================================================
-- SCRIPT:      AR_Data_Pull.sql
-- AUTHOR:      Chaitanya Yarlagadda
-- ENVIRONMENT: IBM DB2 for i (AS400)
-- PURPOSE:     Pulls open AR invoices to feed the aging workbook.
--
-- NOTE: Written to DB2 for i syntax from IBM's documentation. I have not run this
--       against a live system, so the table and column names would need checking
--       against the real schema.
-- ===================================================================================

SELECT
    INV_ID          AS Invoice_ID,
    CUST_NAME       AS Customer_Name,
    PAYER_TYPE      AS Payer_Type,        -- Medicare / Medi-Cal / Commercial
    INV_DATE        AS Invoice_Date,
    BALANCE         AS Balance_Due,

    -- DAYS() is the DB2 for i way of getting a day count between two dates.
    -- This is age from the invoice date, not days past due. Past due would need
    -- payment terms, which this table does not carry.
    DAYS(CURRENT_DATE) - DAYS(INV_DATE)   AS Days_Outstanding

FROM FINLIB.AR_MASTER

WHERE BALANCE > 0                          -- open balances only
  AND ACCT_STATUS = 'ACTIVE'               -- leaves out closed and archived accounts

-- Oldest first, so the accounts closest to becoming uncollectible are at the top.
ORDER BY Days_Outstanding DESC;
