---
id: v2-bank-reconciliation
title: Bank Reconciliation
---

# Bank Reconciliation

Bank reconciliation in HITRIX V2 is a focused, manual matching process: a user selects a bank account and a date range, the system shows all unreconciled payment and receipt vouchers for that bank in that period, and the user fills in the reconciliation date row by row as each cheque or RTGS entry is confirmed on the bank statement.

---

## How It Works

```mermaid
flowchart TD
    A([User opens frmBankReConcilEntry]) --> B{Entry type?}
    B -->|Receipt BR| C[Filter tblVoucher\nVtype='BR'\nVACOCODE = bank AcCode]
    B -->|Payment BP| D[Filter tblVoucher\nVtype='BP'\nVACOCODE = bank AcCode]
    C --> E[Grid shows all rows\nwhere VReconDt IS NULL\nfor the date range]
    D --> E
    E --> F[User types reconciliation date\nin column 5 per row]
    F --> G{Save clicked}
    G --> H[For each row with a date:\nUPDATE tblVoucher\nSET VReconDt = date\nWHERE Vno+Vtype+Vyear]
    G --> I[For each row cleared:\nUPDATE tblVoucher\nSET VReconDt = NULL]
    H --> J[Running total\ntxttemp(2) shows\nsum of reconciled amounts]
    I --> J
    J --> K([Form closes])
```

---

## Form: frmBankReConcilEntry

The form has two modes, set via `gSelectedMenu` at load time:

| Mode | VType filtered | Label |
|---|---|---|
| `RE` | `BR` (Bank Receipt) | "Bank Reconciliation Entry - Receipt" |
| else | `BP` (Bank Payment) | "Bank Reconciliation Entry - Payment" |

**Inputs:**
- Bank account (typed/searched from `tblMastAccount` where `AgCode = 90019` — bank group)
- From Date and To Date (masked date fields, default to financial year start/end)

**Save is disabled** until a valid bank account and both dates are set.

---

## Grid Structure

The grid is populated by a query against `tblVoucher`:

```sql
SELECT VNo, AcName, ABS(VAmt), SUBSTR(VNar1, 10),
       VDt, VReconDt, Vtype
FROM tblVoucher, tblMastAccount
WHERE VACCODE = AcCode
  AND Vtype = 'BR'  -- or 'BP'
  AND VACOCODE = <bank_AcCode>
  AND Vdt >= '<from_dt>' AND Vdt <= '<to_dt>'
ORDER BY Vdt, VReConDt, Vno
```

| Col | Content | Editable |
|---|---|---|
| 0 | VNo | No |
| 1 | Party name (AcName) | No |
| 2 | Amount (ABS of VAmt) | No |
| 3 | Narration (from VNar1 after position 10) | No |
| 4 | Voucher date (VDt) | No |
| 5 | **Reconciliation date (VReconDt)** | **Yes** |
| 6 | VType (hidden) | No |

Only column 5 (the reconciliation date) is user-editable. All other columns are read-only. The `VReconDt` field uses the last entered date as a default for the next row (`tmpReConDt` variable carries it forward), so reconciling a sequence of dates is fast.

---

## Save Logic

On save, for every row in the grid:

```vb
If IsDate(.TextMatrix(i, 5)) Then
    UPDATE tblVoucher SET VReconDt = '<date>'
    WHERE Vno = <vno> AND Vtype = '<type>' AND Vyear = <year>
Else
    UPDATE tblVoucher SET VReconDt = NULL
    WHERE Vno = <vno> AND Vtype = '<type>' AND Vyear = <year>
End If
```

No insert or delete — only `VReconDt` is modified. A cleared date field sets the column back to `NULL`, un-reconciling that voucher.

---

## Reconciliation Total

The running total label (`txttemp(2)`) shows the sum of `VAmt` for all rows where `VReconDt` is non-null. This is recalculated every time the user leaves column 5 (`CalGridAmt` sub). It helps the user verify their reconciled total matches the bank statement's closing balance.

```vb
For each row where IsDate(TextMatrix(row, 5)):
    Amt += CDbl(TextMatrix(row, 2))
txttemp(2) = Format(Amt, FStr)
```

---

## What Is Stored

`VReconDt` is the only field written during reconciliation. It lives on `tblVoucher`:

| Field | Description |
|---|---|
| `VReconDt` | `smalldatetime` — the date this voucher line was confirmed against the bank statement. `NULL` = unreconciled. |

This is a direct column on every voucher row. There is no separate reconciliation table or statement import. The bank statement matching is entirely manual — the user compares the paper/PDF bank statement against what is on screen and types the date when it matches.

---

## Reports / Queries That Use VReconDt

Any bank book or cash flow report that filters on `VReconDt IS NULL` shows the outstanding (uncleared) cheques and RTGS entries — i.e., items in HITRIX that have not yet hit the bank account. This is the standard "book vs. bank" difference used to prepare a monthly BRS (Bank Reconciliation Statement).

---

## Key Business Rules

| Rule | Detail |
|---|---|
| One bank at a time | The form always scopes to a single bank account (`AgCode = 90019`) |
| Separate receipt vs. payment passes | BR and BP are reconciled separately via two menu entries that set `gSelectedMenu` |
| No audited check | The reconciliation form has the audited check commented out — reconciliation is allowed even on audited vouchers |
| No VFirm restriction shown | The query does NOT filter by `VFirm` in the form code — all BR/BP entries for the bank account across the date range are shown (this may be an intentional multi-firm BRS view) |
| Date carry-forward | `tmpReConDt` carries the last-entered reconciliation date to the next row, speeding up batch entry for consecutive dates |
| No statement import | There is no bank statement file import or automatic matching — entirely manual |

---

## DhanMan Gaps

| Gap | Detail |
|---|---|
| Bank reconciliation UI | No equivalent in DhanMan — needs a bank-wise uncleared voucher grid with recon date entry |
| `VReconDt` field | DhanMan payment/receipt records have no reconciliation date field |
| Manual-only matching | DhanMan should offer bank statement import (CSV/OFX) with auto-matching by amount + date + reference, with manual override for unmatched items |
| BRS report | A "Book vs. Bank" difference report (uncleared cheques, deposits in transit) needs to be built — currently generated by filtering on `VReconDt IS NULL` |
| Multi-entry pass | DhanMan should allow bulk date entry across multiple voucher lines in a single save, mirroring the grid approach |
