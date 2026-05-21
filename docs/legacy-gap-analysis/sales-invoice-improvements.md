# Sales Invoice Create — Improvements & Upgrade Plan

## Overview

This document tracks what is **missing** and what needs **upgrading** in the Invoice Create feature (`dhanman-app/src/pages/sales/invoice/`). It is a companion to `sales-invoice-create-design.md`.

---

## Missing Features (not built yet)

### M1 — Server-Generated Invoice Number

**Current state:** `invoiceNumber` is hardcoded as `'IN000'` in `create.tsx:35`. User sees a meaningless placeholder.

**What to build:**
- Add or confirm `useGetNextInvoiceNumber(companyId)` SWR hook in `src/api/sales/invoiceHeaderId.ts` (file already exists)
- Call it in `create.tsx` and pass the result as `invoiceNumber` in `initialValues`
- Render `invoiceNumber` as a read-only `TextField` (disabled) in the header row

**Files:** `create.tsx`, `invoiceForm.tsx`, `src/api/sales/invoiceHeaderId.ts`

---

### M2 — Payment Terms Selection

**Current state:** `paymentTerm` is hardcoded as `10` inside `handleCreate` (`invoiceForm.tsx:303`). No UI for it.

**What to build:**
- Add `useGetPaymentTerms(companyId)` hook in `src/api/common/`
- Add `paymentTermId: string` to `InvoiceFormikValues`
- Add a `Select` or `Autocomplete` in the header row beside Due Date
- Replace hardcoded `paymentTerm: 10` in `handleCreate` with `values.paymentTermId`

**Files:** `invoiceForm.tsx`, `InvoiceFormValues.tsx`, `src/api/common/paymentTerms.ts` (new)

---

### M3 — TDS / TCS Fields

**Current state:** `tdsAmount` and `tcsAmount` are hardcoded as `0` passed to `ItemsSummary`. No capture from user.

**What to build:**
- Add `tdsPercent: number` and `tcsPercent: number` to `InvoiceFormikValues` (default `0`)
- Add `TDS` and `TCS` Switch toggles in the switches row alongside existing SO NO / GST / Discount / Fees
- When toggled, show number inputs for TDS % and TCS % in the footer area
- Derive amounts:

```tsx
const tdsAmount = (taxableAmount * toSafeNumber(values.tdsPercent)) / 100;
const tcsAmount = (taxableAmount * toSafeNumber(values.tcsPercent)) / 100;
const grandAmount = Math.round(
  taxableAmount + cgstAmount + sgstAmount + tcsAmount - tdsAmount
);
```

- Pass real values to `ItemsSummary` replacing the hardcoded `0`
- Include `tdsAmount`, `tcsAmount` in `handleCreate` payload

**Files:** `invoiceForm.tsx`, `InvoiceFormValues.tsx`, `components/B2A/ItemsSummary`

---

### M4 — Invoice-Level (Header) Discount

**Current state:** Only line-level percentage discount exists. No way to apply a flat discount across the whole invoice.

**What to build:**
- Add `headerDiscount: number` (percentage, default `0`) to `InvoiceFormikValues`
- Add a `TextField` input in the footer left section beside the Notes field, labelled **Invoice Discount (%)**
- Derive and apply in grand total calculation:

```tsx
const headerDiscountAmount = (taxableAmount * toSafeNumber(values.headerDiscount)) / 100;
const grandAmount = Math.round(
  taxableAmount + cgstAmount + sgstAmount - headerDiscountAmount
);
```

- Add `headerDiscount` as a new prop to `ItemsSummary` and display it between Taxable Amount and GST rows

**Files:** `invoiceForm.tsx`, `InvoiceFormValues.tsx`, `components/B2A/ItemsSummary`

---

### M5 — Print / PDF Export

**Current state:** Preview button exists in the UI (`invoiceForm.tsx:1394`) but is permanently `disabled`. No print or download flow exists.

**What to build:**
- Enable the Preview button only after first save (when `invoiceHeaderId` is set)
- Add a `/sales/invoices/:id/print` route that renders a stripped print layout
- Use `invoiceExcel/invoiceTemplate.tsx` (already exists) as the base template
- Trigger `window.print()` on mount for browser print, or use `react-to-pdf` / `jsPDF` for PDF download

```tsx
<Button
  variant="outlined"
  color="secondary"
  disabled={!invoiceHeaderId}
  onClick={() => window.open(`/sales/invoices/${invoiceHeaderId}/print`, '_blank')}
>
  Preview
</Button>
```

**Files:** `invoiceForm.tsx`, `invoiceExcel/invoiceTemplate.tsx`, `src/routes/modules/sales-routes.tsx`

---

## Needs Upgrading (partially built or incomplete)

### U1 — Save as Draft

**Current state:** Save button is commented out in `invoiceForm.tsx:1397–1399`. The DB table `draft_invoice_headers.sql` already exists — backend is ready.

**What to upgrade:**
- Add `insertDraftInvoice(invoice)` function to `src/api/sales/invoice.ts`
- Uncomment the Save button and assign `submitType: 'saveAsDraft'`
- In `handleCreate`, branch before the existing create/update logic:

```tsx
if (values.submitType === 'saveAsDraft') {
  response = await insertDraftInvoice(newInvoice);
  SuccessToastMessage('invoiceSavedAsDraft');
  setInvoiceHeaderId(response.id);
  hideBackdrop();
  return; // stay on page, do not navigate
}
```

- After save, show a `Draft` status chip via existing `BStatusDisplay` component

**Files:** `invoiceForm.tsx`, `src/api/sales/invoice.ts`

---

### U2 — Approval Workflow Trigger

**Current state:** `InvoiceApprovalButtons.tsx` already exists at `pages/sales/invoice/InvoiceApprovalButtons.tsx`. The DB has `invoice_approval_logs`, `invoice_workflows`, and `invoice_approval_users_account`. It is never rendered from the create/edit form.

**What to upgrade:**
- After `handleCreate` sets `invoiceHeaderId`, conditionally render `InvoiceApprovalButtons` in `MainFooter` when approval is required:

```tsx
{invoiceHeaderId && defaultInvoiceStatus?.requiresApproval && (
  <InvoiceApprovalButtons
    invoiceHeaderId={invoiceHeaderId}
    companyId={company?.id}
  />
)}
```

- No new API work needed — the component already handles the approval call

**Files:** `invoiceForm.tsx`

---

### U3 — isSameState GST Routing Bug (fix before any GST upgrade)

**Current state:** `isSameState` is declared twice — as `useState` at component scope and `const` inside the Formik render prop. The inner declaration shadows the outer, meaning the state-driven warehouse comparison never controls GST routing.

**Why it must be fixed first:** Any upgrade to TDS/TCS (M3) or GST display depends on correct intra/inter state detection. All GST amounts are potentially wrong until this is resolved.

**What to fix:**
- Remove the component-level `useState` for `isSameState` and its `useEffect`
- Keep only the `const isSameState = compareGST(...)` inside the Formik render prop as the single source of truth

**Files:** `invoiceForm.tsx`

---

## Priority Order

| # | ID | Feature | Effort | Dependency |
|---|---|---|---|---|
| 1 | U3 | Fix isSameState GST bug | Low | None — fix first |
| 2 | U1 | Save as Draft | Low | Draft API endpoint |
| 3 | M1 | Invoice number from server | Low | Next-number endpoint |
| 4 | M2 | Payment terms dropdown | Low | Payment terms master API |
| 5 | U2 | Approval workflow trigger | Low | None — component exists |
| 6 | M3 | TDS / TCS fields | Medium | `ItemsSummary` prop update |
| 7 | M4 | Header discount | Medium | `ItemsSummary` prop update |
| 8 | M5 | Print / PDF export | Medium | `invoiceTemplate.tsx` review |

---

## Implementation Prompt

Use the following prompt with Claude Code to implement all improvements:

```
Implement the Invoice Create improvements documented in:
  hitrix-srk/docs/legacy-gap-analysis/sales-invoice-improvements.md

Target files:
  src/pages/sales/invoice/create.tsx
  src/pages/sales/invoice/invoiceForm.tsx
  src/pages/sales/invoice/type/InvoiceFormValues.tsx
  src/pages/sales/invoice/schemas/invoiceSchemas.ts
  src/api/sales/invoice.ts
  src/api/sales/invoiceHeaderId.ts
  components/B2A/ItemsSummary (read before editing)

Work in this exact order:

## U3 — Fix isSameState bug first
Remove the component-level useState for isSameState and its useEffect.
Keep only const isSameState = compareGST(...) inside the Formik render prop.

## U1 — Save as Draft
Add insertDraftInvoice to src/api/sales/invoice.ts.
Uncomment the Save button in invoiceForm.tsx.
Add submitType = 'saveAsDraft' branch in handleCreate before existing logic.
Stay on page after draft save. Show BStatusDisplay with Draft status.

## M1 — Server invoice number
In src/api/sales/invoiceHeaderId.ts confirm or add useGetNextInvoiceNumber(companyId).
In create.tsx call the hook and pass result as invoiceNumber in initialValues.
In invoiceForm.tsx render invoiceNumber as disabled read-only TextField.

## M2 — Payment terms dropdown
Create src/api/common/paymentTerms.ts with useGetPaymentTerms(companyId).
Add paymentTermId to InvoiceFormikValues.
Add Select dropdown in invoiceForm.tsx header row beside Due Date.
Replace hardcoded paymentTerm: 10 in handleCreate with values.paymentTermId.

## U2 — Approval trigger
In invoiceForm.tsx MainFooter, conditionally render InvoiceApprovalButtons
when invoiceHeaderId is set and defaultInvoiceStatus.requiresApproval is true.

## M3 — TDS / TCS fields
Add tdsPercent and tcsPercent to InvoiceFormikValues (default 0).
Add TDS and TCS Switch toggles in the switches row.
Add number inputs shown when toggles are on.
Derive tdsAmount and tcsAmount inside the totals useMemo.
Pass real values to ItemsSummary replacing hardcoded 0.
Include in handleCreate payload.

## M4 — Header discount
Add headerDiscount to InvoiceFormikValues (default 0).
Add TextField input in footer left beside Notes labelled Invoice Discount (%).
Include headerDiscountAmount in grand total derivation.
Pass to ItemsSummary as new prop.

## M5 — Print / PDF
Enable Preview button when invoiceHeaderId is set.
Open /sales/invoices/:id/print in new tab.
Add print route in sales-routes.tsx using invoiceExcel/invoiceTemplate.tsx.

Do NOT:
- Change any existing GST column logic beyond fixing isSameState
- Modify edit.tsx or copy.tsx unless a shared type change requires it
- Change ItemsSummary props in a way that breaks other callers

After each item confirm TypeScript compiles: npx tsc --noEmit
```
