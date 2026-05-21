# Purchase Bill Create — Improvements & Upgrade Plan

## Overview

Companion to `purchase-bill-create-design.md`. Tracks what is missing and what needs upgrading in `dhanman-app/src/pages/purchase/bills/`.

---

## Missing Features

### M1 — Server-Generated Bill Number

**Current state:** `billNumber: ''` in `create.tsx:38`. Empty string shown to user.

**What to build:**
- Add or confirm `useGetNextBillNumber(companyId)` in `src/api/purchase/billHeaderId.ts` (file already exists)
- Call in `create.tsx`, pass result as `billNumber` in `initialValues`
- Render as read-only disabled `TextField` in the header row

**Files:** `create.tsx`, `BillForm.tsx`, `src/api/purchase/billHeaderId.ts`

---

### M2 — Payment Terms Selection

**Current state:** No `paymentTermId` field in `BillFormikValues`. No UI.

**What to build:**
- Add `paymentTermId: string` to `BillFormikValues`
- Add `Select` or `Autocomplete` in header row beside Due Date using `useGetPaymentTerms(companyId)`
- Pass `values.paymentTermId` in `handlerCreate` payload

**Files:** `BillForm.tsx`, `BillFormValues.ts`, `src/api/common/paymentTerms.ts`

---

### M3 — Save as Draft

**Current state:** No draft bill UI or API. Unlike invoice, no draft DB table confirmed — verify with backend before building.

**What to build (once backend confirms):**
- Add `insertDraftBill(bill)` to `src/api/purchase/bill.ts`
- Add Save as Draft button in `MainFooter` with `submitType: 'saveAsDraft'` branch in `handlerCreate`
- Stay on page after draft save; show `Draft` chip via `BStatusDisplay`

**Files:** `BillForm.tsx`, `src/api/purchase/bill.ts`

---

### M4 — Print / PDF Export

**Current state:** Preview button permanently disabled in `BillForm.tsx`.

**What to build:**
- Enable Preview button only when `billHeaderId` is set
- Open `/purchase/bills/:id/print` in new tab
- Add print route using existing bill template or create `billTemplate.tsx`

**Files:** `BillForm.tsx`, `src/routes/modules/purchase-routes.tsx`

---

## Needs Upgrading

### U1 — Complete TDS UI

**Current state:** `tds` (percentage) and `tdsAmount` fields exist in `BillFormikValues` and are included in `handlerCreate`. TDS restoration logic exists for edit mode (`BillForm.tsx:305–317`). But the UI input for entering TDS % is incomplete or missing.

**What to upgrade:**
- Add a `TDS` Switch toggle in the switches row (alongside GST Rates / Discount / Fees)
- When toggled on, show a TDS % `TextField` in the footer
- Derive `tdsAmount = (taxableAmount * tds) / 100` inside the totals useMemo
- Subtract from grand total: `grandAmount = taxableAmount + gstAmount - tdsAmount`
- Already passes to `ItemsSummary` — verify `tdsAmount` prop is wired correctly

**Files:** `BillForm.tsx`

---

### U2 — Approval Workflow Trigger

**Current state:** `BillApprovalButtons.tsx` exists at `pages/purchase/bills/BillApprovalButtons.tsx`. Not rendered from the create/edit form.

**What to upgrade:**
- After `handlerCreate` sets `billHeaderId`, conditionally render `BillApprovalButtons` in `MainFooter`:

```tsx
{billHeaderId && defaultStatus?.requiresApproval && (
  <BillApprovalButtons
    billHeaderId={billHeaderId}
    companyId={company?.id}
  />
)}
```

**Files:** `BillForm.tsx`

---

### U3 — Fix isSameState Shadow Bug

**Current state:** `useState` for `isSameState` at component scope is shadowed by `const isSameState = compareGST(...)` inside Formik render. Same bug as `invoiceForm.tsx`.

**What to fix:**
- Remove component-level `useState` for `isSameState` (`BillForm.tsx:134`) and its `useEffect` (`BillForm.tsx:279–283`)
- Keep only the inline `compareGST` derivation inside the Formik render

**Files:** `BillForm.tsx`

---

### U4 — Remove console.log in Production

**Current state:** `BillForm.tsx:369` logs full bill data on every render.

**Fix:** Delete the line immediately.

```ts
// DELETE THIS LINE
console.log('normalizedInitialValues', normalizedInitialValues);
```

**Files:** `BillForm.tsx`

---

### U5 — Fix Naming Inconsistencies

Three naming issues that cause friction when sharing utilities with the invoice module:

| Issue | Current | Should Be | Files |
|---|---|---|---|
| Line array field | `bill_detail` (snake_case) | `billLines` or `lines` | `BillFormValues.ts`, `BillForm.tsx`, `BillItem.tsx`, `purchaseSchemas.ts` |
| Tax amount casing | `CGSTAmount`, `SGSTAmount`, `IGSTAmount` | `cgstAmount`, `sgstAmount`, `igstAmount` | `BillFormValues.ts`, `BillForm.tsx`, `BillItem.tsx` |
| Settled amount typo | `setteledAmount` | `settledAmount` | `BillFormValues.ts`, `BillForm.tsx` |

---

### U6 — Type Safety

| Location | Issue | Fix |
|---|---|---|
| `BillForm.tsx:83` | `initialValues: any` | Change to `BillFormikValues` |
| `BillForm.tsx:153` | `formikRef: FormikProps<any>` | Change to `FormikProps<BillFormikValues>` |
| `BillFormValues.ts:12` | `selectedProduct: any` | Change to `IProduct \| null` |
| `BillForm.tsx` calculations | All `.reduce()` callbacks typed `any` | Use `BillLineFormikValues` |

---

### U7 — Move Utility Functions Out of Component

Four pure functions are defined inside `BillForm.tsx` and should be extracted:

| Function | Move To |
|---|---|
| `asDate(v)` | `src/utils/dateUtils.ts` |
| `fmt(v)` | `src/utils/dateUtils.ts` |
| `snapToAllowedGst(pct)` | `src/utils/gstUtils.ts` |
| `inferGstPercent(v)` | `src/utils/gstUtils.ts` |

---

## Priority Order

| # | ID | Item | Effort | Dependency |
|---|---|---|---|---|
| 1 | U3 | Fix isSameState shadow bug | Low | None — fix first |
| 2 | U4 | Remove console.log | Low | None |
| 3 | U1 | Complete TDS UI | Low | None — fields already exist |
| 4 | U2 | Approval workflow trigger | Low | `BillApprovalButtons` exists |
| 5 | M1 | Server-generated bill number | Low | Next-number endpoint |
| 6 | M2 | Payment terms dropdown | Low | Payment terms master API |
| 7 | U5 | Fix naming inconsistencies | Medium | Coordinate with API team on field names |
| 8 | U6 | Type safety | Medium | After U5 renames |
| 9 | U7 | Extract utility functions | Low | After U6 |
| 10 | M3 | Save as Draft | Medium | Backend draft table confirmation |
| 11 | M4 | Print / PDF export | Medium | Template review |

---

## Implementation Prompt

```
Implement the Purchase Bill Create improvements documented in:
  hitrix-srk/docs/legacy-gap-analysis/purchase-bill-improvements.md

Target files:
  src/pages/purchase/bills/create.tsx
  src/pages/purchase/bills/BillForm.tsx
  src/pages/purchase/bills/types/BillFormValues.ts
  src/pages/purchase/schemas/purchaseSchemas.ts
  src/api/purchase/bill.ts
  src/api/purchase/billHeaderId.ts

Before starting, read:
  src/pages/sales/invoice/invoiceForm.tsx   ← use as the pattern reference

Work in this exact order:

## U3 — Fix isSameState shadow
Remove useState for isSameState at BillForm.tsx:134 and its useEffect at :279–283.
Keep only const isSameState = compareGST(...) inside the Formik render prop.

## U4 — Remove console.log
Delete BillForm.tsx:369 console.log line.

## U1 — Complete TDS UI
Add TDS Switch toggle in the switches row.
When on, show TDS % TextField in footer.
Derive tdsAmount = (taxableAmount * values.tds) / 100 inside totals useMemo.
Subtract tdsAmount from grandAmount.
Verify tdsAmount is passed correctly to ItemsSummary.

## U2 — Approval trigger
In BillForm.tsx MainFooter, conditionally render BillApprovalButtons
when billHeaderId is set and defaultStatus.requiresApproval is true.

## M1 — Server bill number
In src/api/purchase/billHeaderId.ts confirm or add useGetNextBillNumber(companyId).
In create.tsx call hook and pass result as billNumber in initialValues.
In BillForm.tsx render billNumber as disabled read-only TextField.

## M2 — Payment terms dropdown
Add paymentTermId to BillFormikValues.
Add Select in header row beside Due Date using useGetPaymentTerms(companyId).
Pass values.paymentTermId in handlerCreate payload.

## U6 — Type safety
Change initialValues: any to BillFormikValues in BillForm Props.
Change formikRef to FormikProps<BillFormikValues>.
Change selectedProduct: any to IProduct | null in BillLineFormikValues.

Do NOT rename bill_detail or CGSTAmount yet — coordinate with backend first (U5).
Do NOT touch edit.tsx or copy edit files unless a shared type change requires it.

After each item confirm TypeScript compiles: npx tsc --noEmit
```

---

## Trade Purchase Booking

Screenshot to be provided by developer. Once received, a separate design doc
`purchase-trade-booking-design.md` will be created following the same architecture
as `party-booking-create-design.md`.
