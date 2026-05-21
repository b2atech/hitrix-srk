# Purchase Bill Create — Feature Design & Analysis

## Feature Scope

The bill create feature handles three operating modes from a single form:

| Mode | Entry Point | Behaviour |
|---|---|---|
| Create | `pages/purchase/bills/create.tsx` | Fresh bill with empty initialValues |
| Edit | `pages/purchase/bills/edit.tsx` | Pre-populated form; calls `updateNewBill` on submit |
| Copy | `pages/purchase/bills/editInventoryBill.tsx` / `editExpenseBill.tsx` | Pre-populated; generates new ID on submit |

Two bill types are supported via `InvoiceBillType` enum:

- `InvoiceBillType.Inventory` — requires warehouse source/destination and GST routing
- `InvoiceBillType.Expense` — skips warehouse; quantity auto-set to 1 per line

---

## File Structure

```
src/pages/purchase/bills/
├── create.tsx                      ← thin wrapper: sets initialValues, renders BillForm
├── edit.tsx                        ← edit wrapper
├── editInventoryBill.tsx           ← inventory bill edit entry
├── editExpenseBill.tsx             ← expense bill edit entry
├── BillForm.tsx                    ← main form component
├── BillApprovalButtons.tsx         ← approval action buttons (exists, usage gap)
├── billListColDefs.tsx             ← AG Grid col defs
├── list.tsx / inventoryBillList.tsx / expenseBillList.tsx
├── view.tsx
└── types/
    └── BillFormValues.ts           ← BillFormikValues and BillLineFormikValues types

src/pages/purchase/components/
├── BillItem.tsx                    ← per-row line item component
├── BillField.tsx                   ← editable table cell
└── AlertBillProductDelete.tsx      ← delete confirmation dialog

src/pages/purchase/schemas/
└── purchaseSchemas.ts              ← BillValidationSchema (Yup)

src/api/purchase/
├── bill.ts                         ← insertNewBill, updateNewBill
├── billHeaderId.ts                 ← next bill number (exists)
└── billStatus.ts                   ← useGetDefaultBillStatus
```

---

## Data Model

### BillFormikValues (header)

| Field | Type | Notes |
|---|---|---|
| `id` | `string \| number` | Empty on create; set after first save |
| `billNumber` | `string` | Empty string on create — not server-generated |
| `billDate` | `Date \| null` | Required |
| `dueDate` | `Date \| null` | Required; initial value `null` — crash risk on submit |
| `poNo` | `string` | PO reference; shown via toggle |
| `poDate` | `Date \| string` | PO date alongside poNo |
| `debitAccountId` | `string \| null` | Purchase/expense account |
| `debitAccount` | `IAccount` | Full account object for display |
| `creditAccountId` | `string` | Accounts Payable; from org accounts |
| `vendorInfo` | `VendorInfoForm` | `{ id, name, gstIn }` |
| `sourceWarehouseId` | `string` | Vendor warehouse; hidden for Expense bills |
| `destinationWarehouseId` | `string` | Company warehouse; hidden for Expense bills |
| `bill_detail` | `BillLineFormikValues[]` | Min one line — **uses snake_case (inconsistency)** |
| `currencyId` | `number` | Defaults to INR |
| `discount` | `number` | Aggregate line discount (derived) |
| `fees` | `number` | Aggregate line fees (derived) |
| `taxableAmount` | `number` | Derived |
| `cgstAmount / sgstAmount / igstAmount` | `number` | Derived |
| `tdsAmount` | `number` | TDS deduction amount — field exists, partial UI |
| `tds` | `number` | TDS percentage — field exists, partial UI |
| `gstNumber` | `number` | Header-level GST % for expense bills |
| `roundOff` | `number` | Derived |
| `totalAmount` | `number` | Derived grand total |
| `setteledAmount` | `number` | **Typo — should be `settledAmount`** |
| `note` | `string` | Max 500 characters |
| `submitType` | `string` | `'createAndSend'` etc. |

### BillLineFormikValues (per line)

| Field | Type | Notes |
|---|---|---|
| `id` | `string` | UUID row key |
| `productId` | `string` | Product ID |
| `selectedProduct` | `any` | **Untyped — should be `IProduct \| null`** |
| `description` | `string` | Product description |
| `quantity` | `number` | Auto-set to `1` for Expense bills |
| `price` | `number` | Unit price |
| `discount` | `number` | Percentage discount |
| `fees` | `number` | Additional fees |
| `cgst / sgst / igst` | `number` | Tax rates from product master |
| `taxableAmount` | `number` | Derived per line |
| `CGSTAmount / SGSTAmount / IGSTAmount` | `number` | **PascalCase — inconsistent with invoice's camelCase** |
| `serialNumber` | `number` | Display order |
| `rateVisibility / discountVisibility / feesVisibility` | `boolean` | Column toggle flags |

---

## Component Architecture

```
CreateBill (create.tsx)
└── BillForm (BillForm.tsx)
    ├── PageHeader
    ├── MainCard
    │   └── Formik (innerRef={formikRef}, enableReinitialize={isInEditMode})
    │       ├── Vendor Autocomplete (MultilineTooltipAutocomplete)
    │       ├── BDate (bill date)
    │       ├── BDate (due date)
    │       ├── Purchase Account Autocomplete
    │       ├── Currency Select
    │       ├── WareHouse (source/vendor) [Inventory only]
    │       ├── WareHouse (destination/company) [Inventory only]
    │       ├── PO No + PO Date [toggle]
    │       ├── GST % field [Expense only]
    │       ├── Switches: PO NO | GST Rates | Discount | Fees
    │       ├── FieldArray → BillItem rows (name: 'bill_detail')
    │       │   └── BillItem (per row)
    │       ├── Notes TextField
    │       ├── TDS % input + TDS Amount (partial — exists but incomplete)
    │       ├── ItemsSummary (totals panel)
    │       └── MainFooter
    │           ├── Preview Button [disabled]
    │           └── Create & Send / Update & Send / Copy & Create
    └── Dialog: vendor-change confirmation
```

### Key Hooks Used

| Hook | Source | Purpose |
|---|---|---|
| `useGetVendorNames` | `api/purchase/vendor` | Vendor dropdown |
| `useGetProductByVendorId` | `api/inventory/product` | Vendor-specific products |
| `useGetAllProducts` | `api/inventory/product` | All products |
| `useGetDefaultBillStatus` | `api/purchase/billStatus` | Initial status for new bills |
| `useGetCurrencies` | `api/common/currencies` | Currency dropdown |
| `useGetOrganizationAccounts` | `api/common/organization` | AP account ID |
| `useGetPurchaseAccounts` | `api/common/accounts` | Purchase account dropdown |
| `useBackdrop` | `hooks/useBackdropLoader` | Loading overlay |
| `useAuth` | `hooks/useAuth` | User for activity log |
| `useConfig` | `hooks/useConfig` | company context |

---

## Known Bugs

### B1 — `isSameState` useState shadow (HIGH)

**File:** `BillForm.tsx:134`, then again derived inside Formik render prop

Same bug as in `invoiceForm.tsx`. The component-level `useState` for `isSameState` (set by the warehouse `useEffect`) is shadowed by a `const isSameState = compareGST(...)` computed inside the Formik render. GST routing may be incorrect in some warehouse combinations.

**Fix:** Remove the component-level `useState` and its `useEffect`. Keep only the inline `compareGST` inside the Formik render, or extract to a `useMemo`.

---

### B2 — `dueDate` null crash-risk on submit (HIGH)

**File:** `BillForm.tsx` `handlerCreate`

`dueDate` initial value is `null` in `create.tsx:44`. The `fmt()` helper guards against this (`fmt(values?.dueDate) ?? billDateStr`) which is better than invoice — but the fallback silently uses `billDate` instead of failing validation. The Yup schema should require `dueDate`.

**Fix:** Mark `dueDate` required in `BillValidationSchema`. Remove the silent fallback.

---

### B3 — `console.log` debug statement in production (MEDIUM)

**File:** `BillForm.tsx:369`

```ts
console.log('normalizedInitialValues', normalizedInitialValues);
```

This logs full bill data on every render in production. Must be removed.

---

### B4 — `bill_detail` snake_case field name (MEDIUM)

**File:** `BillFormValues.ts:48`, `BillForm.tsx` throughout

The line array field is named `bill_detail` (snake_case) while every other Formik field and the invoice equivalent uses camelCase (`lines`). This breaks naming consistency and causes friction when sharing utilities between bill and invoice.

**Fix:** Rename to `billLines` or `lines` — update all references in `BillForm.tsx`, `BillItem.tsx`, `purchaseSchemas.ts`, and API mapping.

---

### B5 — `setteledAmount` typo (LOW)

**File:** `BillFormValues.ts:60`, `BillForm.tsx`

`setteledAmount` should be `settledAmount`. Propagates into API payloads and any backend field mapping.

---

### B6 — `initialValues: any` in BillForm Props (MEDIUM)

**File:** `BillForm.tsx:83`

Same as invoice — `initialValues` is typed `any`. Should be `BillFormikValues`.

---

### B7 — `CGSTAmount / SGSTAmount / IGSTAmount` PascalCase (LOW)

**File:** `BillFormValues.ts:22–27`

Tax amount fields use PascalCase (`CGSTAmount`) while invoice uses camelCase (`cgstAmount`). Inconsistent casing across the two modules creates confusion when sharing summary or reporting components.

---

## Optimization Recommendations

### O1 — Extract tax calculations to useMemo

`BillForm.tsx` has the same inline `.reduce()` tax calculations as `invoiceForm.tsx`. Wrap in a single `useMemo([values.bill_detail, isSameState])` to stop recalculating on every keystroke.

### O2 — Move `asDate`, `fmt`, `snapToAllowedGst`, `inferGstPercent` out of component

These are pure utility functions defined inside the component. Move to `src/utils/dateUtils.ts` and `src/utils/gstUtils.ts` respectively.

### O3 — Remove `normalizedInitialValues` `useMemo` console.log

Remove `BillForm.tsx:369` debug log immediately.

### O4 — Type `selectedProduct` strictly

Change `selectedProduct: any` in `BillLineFormikValues` to `selectedProduct: IProduct | null`.

### O5 — Remove or implement the Preview button

Same as invoice — Preview button is permanently disabled. Wire or remove.

---

## Feature Gaps vs Full Bill Workflow

| Gap | Status | Notes |
|---|---|---|
| Server-generated bill number | Missing — empty string on create | Same gap as invoice |
| Payment terms selection | Missing — no field in form or type | Same gap as invoice |
| Save as Draft | Missing — no draft bill API or UI | Unlike invoice, no draft DB table confirmed |
| TDS capture UI | Partial — fields exist in type and handleCreate, but UI is incomplete | `tds` and `tdsAmount` exist in `BillFormikValues` |
| Print / PDF export | Missing — Preview disabled | Same gap as invoice |
| Approval workflow from create screen | Partial — `BillApprovalButtons.tsx` exists but not rendered from create form | Same gap as invoice |
| Expense bill GST header field | Partial — `gstNumber` field exists but inference logic is complex and fragile | Only for Expense type |

---

## Bill vs Invoice — Differences Summary

| Aspect | Invoice Create | Bill Create |
|---|---|---|
| Party | Customer (`customerInfo`) | Vendor (`vendorInfo`) |
| Line array field name | `lines` (camelCase) | `bill_detail` (snake_case) — inconsistency |
| Tax amount casing | `cgstAmount` (camelCase) | `CGSTAmount` (PascalCase) — inconsistency |
| TDS | Missing entirely | Partial — fields exist, UI incomplete |
| GST header field | Not present | `gstNumber` for Expense bills |
| Vendor autocomplete | `LoadingAutocomplete` | `MultilineTooltipAutocomplete` — different component |
| Data normalization | None | `normalizedInitialValues` useMemo — handles edit hydration |
| `isSameState` bug | Present (shadow) | Present (shadow) — same bug |
| Draft support | DB exists (`draft_invoice_headers`) | No draft table confirmed |
| Bill number on create | `'IN000'` hardcoded | Empty string `''` |

---

## Open Questions

- Does the API auto-generate `billNumber`, or is it always user-entered?
- Is the `gstNumber` field (Expense bills) the overall GST % applied to the whole bill, or per-line?
- Should `poNo` link to an actual Purchase Order record, or remain a free-text reference?
- Is `setteledAmount` ever set from the UI, or only from the API response?
- Is there a `draft_purchase_bills` table in the backend?
