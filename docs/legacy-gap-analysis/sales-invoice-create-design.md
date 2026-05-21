# Sales Invoice Create — Feature Design & Analysis

## Purpose

This document records the design, current implementation state, known bugs, and optimization recommendations for the Sales Invoice Create screen in `dhanman-app`. It also serves as the reference baseline for the upcoming Sales Order create screen design.

---

## Feature Scope

The invoice create feature handles three operating modes from a single form:

| Mode | Entry Point | Behaviour |
|---|---|---|
| Create | `pages/sales/invoice/create.tsx` | Fresh invoice with empty initialValues |
| Edit | `pages/sales/invoice/edit.tsx` | Pre-populated form; calls `updateInvoice` on submit |
| Copy | `pages/sales/invoice/copy.tsx` | Pre-populated form; generates a new ID on submit |

Two invoice types are supported via the `InvoiceBillType` enum:

- `InvoiceBillType.Product` — requires warehouse source/destination and GST routing
- `InvoiceBillType.Service` — skips warehouse fields; uses a fixed default warehouse ID

---

## File Structure

```
src/pages/sales/invoice/
├── create.tsx                     ← thin wrapper: sets initialValues, renders InvoiceForm
├── invoiceForm.tsx                ← main form component (~1427 lines)
├── invoiceItem.tsx                ← single line-item row
├── component/
│   └── VirtualizedListbox.tsx    ← virtualized customer dropdown list
├── type/
│   ├── InvoiceFormValues.tsx     ← InvoiceFormikValues and InvoiceLineFormikValues types
│   ├── InvoiceEditFormikValues.tsx
│   └── invoiceUtils.tsx
└── schemas/
    └── invoiceSchemas.ts          ← Yup validation schema
```

Related API layer:

```
src/api/sales/
├── invoice.ts         ← insertNewInvoice, updateInvoice
├── invoiceStatus.ts   ← useGetDefaultInvoiceStatus
└── customer.ts        ← useGetCustomerNames
```

---

## Data Model

### InvoiceFormikValues (header)

| Field | Type | Notes |
|---|---|---|
| `id` | `string \| number` | Empty on create; set after first save |
| `invoiceNumber` | `string` | Hardcoded `'IN000'` on create — not server-generated |
| `invoiceVoucher` | `string` | Hardcoded `'INV0023'` on create |
| `invoiceDate` | `Date \| null` | Required |
| `dueDate` | `Date \| null` | Required; initial value is `null` — see bug below |
| `soNo` | `string \| number` | Optional; toggled visible via SO NO switch |
| `soDate` | `Date \| string` | Optional |
| `creditAccountId` | `string` | Sales account; auto-filled from customer default |
| `debitAccountId` | `string` | Accounts Receivable; read from org accounts |
| `sourceWarehouseId` | `string` | Company warehouse; hidden for Service invoices |
| `destinationWarehouseId` | `string` | Customer warehouse; hidden for Service invoices |
| `customerInfo` | `CustomerInfoForm` | `{ id, name, gstIn }` |
| `currencyId` | `number` | Defaults to INR |
| `lines` | `InvoiceLineFormikValues[]` | Minimum one line required |
| `discount` | `number` | Aggregate line discount in rupees (derived, not entered) |
| `fees` | `number` | Aggregate line fees in rupees (derived, not entered) |
| `taxableAmount` | `number` | Derived |
| `cgstAmount / sgstAmount / igstAmount` | `number` | Derived; CGST+SGST for intra-state, IGST for inter-state |
| `roundOff` | `number` | Derived: `Math.round(total) - total` |
| `totalAmount` | `number` | Derived grand total |
| `note` | `string` | Max 500 characters |
| `type` | `InvoiceBillType` | `Product` or `Service` |

### InvoiceLineFormikValues (per line)

| Field | Type | Notes |
|---|---|---|
| `id` | `string` | UUID; stable row key |
| `selectedProduct` | `any` | Full product object; `null` for empty rows |
| `quantity` | `number \| string` | String during input, converted on submit |
| `price` | `number` | Unit price |
| `discount` | `number` | Percentage discount per line |
| `fees` | `number` | Additional fees per line |
| `cgst / sgst / igst` | `number` | Tax rates from product master |
| `taxableAmount` | `number` | Derived per line |
| `serialNumber` | `number` | Display order |

---

## Component Architecture

```
CreateInvoice (create.tsx)
└── InvoiceForm (invoiceForm.tsx)
    ├── PageHeader
    ├── MainCard
    │   └── Formik
    │       ├── Customer Autocomplete (VirtualizedListbox)
    │       ├── BDate (invoice date)
    │       ├── BDate (due date)
    │       ├── Sales Account Autocomplete
    │       ├── Currency Select
    │       ├── WareHouse (source) [Product only]
    │       ├── WareHouse (destination) [Product only]
    │       ├── SO No + SO Date [toggle]
    │       ├── Switches: SO NO | GST Rates | Discount | Fees
    │       ├── FieldArray → InvoiceItem rows
    │       │   └── InvoiceItem (per row)
    │       ├── Notes TextField
    │       ├── ItemsSummary (totals panel)
    │       └── MainFooter
    │           ├── Preview Button [disabled]
    │           └── Create & Send / Update & Send / Copy & Create Button
    └── Dialog: customer-change confirmation
```

### Key Hooks Used

| Hook | Source | Purpose |
|---|---|---|
| `useGetCustomerNames` | `api/sales/customer` | Customer dropdown list |
| `useGetPrimaryOwners` | `api/community/unit` | Merged with customers for apartment mode |
| `useGetDefaultInvoiceStatus` | `api/sales/invoiceStatus` | Initial status ID for new invoices |
| `useGetCurrencies` | `api/common/currencies` | Currency dropdown |
| `useGetOrganizationAccounts` | `api/common/organization` | AR account ID |
| `useGetAllProducts` | `api/inventory/product` | Product dropdown (all) |
| `useGetProductByCustomerId` | `api/inventory/product` | Product dropdown (customer-specific) |
| `useDateFilter` | `hooks/useDateFilter` | fromDate/toDate for update API call |
| `useBackdrop` | `hooks/useBackdropLoader` | Loading overlay |
| `useAuth` | `hooks/useAuth` | User for activity log |

---

## Business Logic

### GST Routing (Intra vs Inter State)

GST type is determined by comparing the first two characters of the GSTIN of the source warehouse and destination warehouse. If they match → CGST + SGST apply. If they differ → IGST applies.

The comparison utility is `compareGST` from `components/B2A/BcompareGSTColumns`.

On submit, the correct amounts are zeroed out:
```
cgstAmount: isSameState ? cTaxAmount : 0
igstAmount: !isSameState ? iTaxAmount : 0
```

### Single-Product Customer Auto-fill

When a customer is selected and `useGetProductByCustomerId` returns exactly one product, `invoiceForm.tsx` auto-seeds the first line item with that product's details (name, price, tax rates). This removes a manual step for customers who always buy the same item.

### Customer Change Guard

If any line already has a selected product and the user changes the customer, a confirmation dialog appears to warn that lines will be cleared. This prevents stale product-customer associations.

### Tax Calculation (Inline — currently unextracted)

Six `.reduce()` calls inside the Formik render prop recompute on every render:
- `subtotal`
- `cgstAmount`
- `sgstAmount`
- `igstAmount`
- `feeValue`
- `discountAmount`

These drive `taxableAmount`, `roundingAmount`, `grandAmount`, and `roundingOff`.

---

## Known Bugs

### B1 — `isSameState` variable shadowing (HIGH)

**File:** `invoiceForm.tsx:461`

A `const isSameState` is declared inside the Formik render prop, shadowing the `useState` at component scope. The `useState` value (updated by the `useEffect` watching warehouse changes) is never used to control GST routing — the render-time comparison always wins.

**Fix:** Remove the component-level `useState` for `isSameState`. Derive it only once inside the render prop from warehouse GSTINs, or extract to a `useMemo`.

---

### B2 — `soNo` / `soNumber` field name mismatch (HIGH)

**File:** `create.tsx:43` initialises `soNo: ''`; `invoiceForm.tsx:961` reads `values.soNumber`

The SO No text field binds to `values.soNumber` which is never initialised. Input is silently lost on submit.

**Fix:** Rename the field reference in `invoiceForm.tsx:961` from `values.soNumber` to `values.soNo`, and ensure `handleChange` on the `TextField` uses `name="soNo"`.

---

### B3 — Null `dueDate` crash-risk on submit (HIGH)

**File:** `invoiceForm.tsx:316`

```ts
dueDate: format(new Date(values?.dueDate), 'yyyy-MM-dd'),
```

When `dueDate` is `null` (its create-mode initial value), `new Date(null)` evaluates to epoch (`1970-01-01`). The invoice is saved with an incorrect due date rather than raising a validation error.

**Fix:** The Yup schema should mark `dueDate` required. In `handleCreate`, guard with `values?.dueDate ? format(...) : null` and let the API reject a null if required.

---

### B4 — Direct Formik value mutation (MEDIUM)

**File:** `invoiceForm.tsx:566–573`

```ts
values.discount = discountAmount;
values.fees = feeValue;
values.taxableAmount = taxableAmount;
// etc.
```

Formik values are mutated directly inside the render function. This bypasses Formik's change tracking and can cause stale-render issues.

**Fix:** Move these derived values into a `useMemo`. Write them to Formik only in `handleCreate` at submit time, not on every render.

---

### B5 — `invoiceNumber` hardcoded on create (MEDIUM)

**File:** `create.tsx:35–36`

```ts
invoiceNumber: 'IN000',
invoiceVoucher: 'INV0023',
```

Invoice number and voucher are hardcoded strings, not server-generated. The API likely assigns the real number on insert, but the UI shows a meaningless placeholder.

**Fix:** Leave `invoiceNumber` blank or fetch an auto-number from the server (if an endpoint exists), then display it as a read-only field after creation.

---

### B6 — `paymentTerm` hardcoded (LOW)

**File:** `invoiceForm.tsx:303`

```ts
paymentTerm: 10,
```

Payment term is hardcoded to 10 with no UI control. Should be a dropdown populated from a payment terms master.

---

## Optimization Recommendations

### O1 — Extract tax calculation to a hook

Create `hooks/useInvoiceTotals.ts`:

```ts
export function useInvoiceTotals(lines: InvoiceLineFormikValues[], isSameState: boolean) {
  return useMemo(() => {
    // subtotal, cgstAmount, sgstAmount, igstAmount, feeValue, discountAmount
  }, [lines, isSameState]);
}
```

This eliminates six redundant `.reduce()` calls per keystroke.

---

### O2 — Extract customer+owner merge to a hook

Create `hooks/useCustomersWithOwners.ts`:

```ts
export function useCustomersWithOwners(companyId: string) {
  const { customerNames, customerNamesLoading } = useGetCustomerNames(companyId, false);
  const { primaryOwners, primaryOwnersLoading } = useGetPrimaryOwners(companyId);
  const merged = useMemo(() => { /* merge logic */ }, [customerNames, primaryOwners]);
  return { customers: merged, loading: customerNamesLoading || primaryOwnersLoading };
}
```

---

### O3 — Type `initialValues` strictly

Change `invoiceForm.tsx:89`:
```ts
// Before
initialValues: any;
// After
initialValues: InvoiceFormikValues;
```

Also type `formikRef` as `useRef<FormikProps<InvoiceFormikValues>>(null)`.

---

### O4 — Move `toSafeNumber` and `isNullOrUndefined` out of the component

Both are pure utilities with no dependency on component state. Move them to `utils/numberUtils.ts`.

---

### O5 — Remove or implement the Preview button

`invoiceForm.tsx:1394` renders a permanently disabled `Preview` button. Either wire it to a print/PDF view or remove it.

---

### O6 — Enable Save as Draft

`invoiceForm.tsx:1397–1399` has the Save button commented out. The database schema already has `draft_invoice_headers.sql`. Uncomment and wire to a separate API call that saves without finalising the status.

---

## Feature Gaps vs Full Invoice Workflow

| Gap | Status | Priority |
|---|---|---|
| Server-generated invoice number displayed in UI | Missing | High |
| Payment terms selection | Missing | Medium |
| Save as Draft | Partial (DB exists, UI commented out) | Medium |
| TDS / TCS capture per invoice | Missing | Medium |
| Invoice-level (header) discount field | Missing | Low |
| Print / PDF export | Missing (Preview disabled) | Medium |
| Approval workflow trigger from create screen | Missing | Low |

---

## Sales Order Screen — Design Baseline

When the Sales Order create screen is designed, the following patterns from invoice create apply directly:

### Reusable Components

| Component | Path | Reuse Notes |
|---|---|---|
| `InvoiceItem` row | `pages/sales/invoice/invoiceItem.tsx` | Rename props; same line structure |
| `ItemsSummary` panel | `components/B2A/ItemsSummary` | Identical totals panel |
| `WareHouse` picker | `components/finance/Warehouses` | Same source/destination pattern |
| `BDate` | `components/B2A/BDate` | Same date picker |
| `BInputLabel` | `components/B2A/BInputLabel` | Same label component |
| `LoadingAutocomplete` | `components/B2A/LoadingAutoComplete` | Same loading autocomplete |
| `VirtualizedListbox` | `pages/sales/invoice/component/VirtualizedListbox` | Same virtualized dropdown |

### Sales Order Additional Fields (not in invoice)

| Field | Type | Notes |
|---|---|---|
| `soStatus` | Enum / Dropdown | Draft / Confirmed / Fulfilled / Cancelled |
| `deliveryDate` | Date | Separate from payment due date |
| `invoiceReference` | String (read-only) | Populated when SO is converted to invoice |
| `approvedBy` | User select | Optional approval field |
| `poReference` | String | Customer's purchase order number |

### Key Difference from Invoice

A Sales Order is a commitment, not a financial transaction. It should not post to accounts or trigger GST at creation. GST and accounting entries happen when the SO is converted to an invoice. The `creditAccountId` and `debitAccountId` fields are therefore not needed on the SO form.

---

## Open Questions

- Does the API auto-generate invoice numbers, or is `IN000` a placeholder the user overwrites?
- Is there a payment terms master endpoint to populate the hardcoded `paymentTerm: 10`?
- Should SO No on the invoice link back to a Sales Order record, or is it a free-text reference field?
- What triggers the transition from Draft to Confirmed status — explicit button, or any save?

---

## Fix Prompt

Use the following prompt with Claude Code (or any AI agent) to apply all bugs fixes and optimizations documented above:

```
Fix the Sales Invoice Create form in dhanman-app based on the design doc at:
  hitrix-srk/docs/legacy-gap-analysis/sales-invoice-create-design.md

Target files:
  src/pages/sales/invoice/create.tsx
  src/pages/sales/invoice/invoiceForm.tsx
  src/pages/sales/invoice/type/InvoiceFormValues.tsx
  src/pages/sales/invoice/schemas/invoiceSchemas.ts

Apply these fixes in order:

## B1 — Remove isSameState useState shadow
In invoiceForm.tsx, remove the component-level useState for isSameState and the
useEffect that sets it. The const isSameState = compareGST(...) inside the Formik
render prop is the correct single source of truth. Remove all references to the outer
state variable.

## B2 — Fix soNo / soNumber field name mismatch
In invoiceForm.tsx around line 961, the TextField for SO No uses name="soNumber" and
reads values.soNumber. Change both to soNo to match the initialValues in create.tsx
and the InvoiceFormikValues type.

## B3 — Guard null dueDate on submit
In invoiceForm.tsx in handleCreate, replace:
  dueDate: format(new Date(values?.dueDate), 'yyyy-MM-dd')
with:
  dueDate: values?.dueDate ? format(new Date(values.dueDate), 'yyyy-MM-dd') : ''
Also confirm invoiceSchemas.ts marks dueDate as required.

## B4 — Stop mutating Formik values directly
In invoiceForm.tsx lines 566–573, remove all direct assignments like:
  values.discount = discountAmount
  values.fees = feeValue
  values.taxableAmount = taxableAmount
  values.cgstAmount = cgstAmount
  values.sgstAmount = sgstAmount
  values.igstAmount = igstAmount
  values.roundOff = roundingOff
  values.totalAmount = grandAmount
These derived values should only be written to the newInvoice object inside handleCreate
at submit time. Do not write them to Formik state on every render.

## O1 — Extract tax calculations to useMemo
Wrap the six reduce calculations (subtotal, cgstAmount, sgstAmount, igstAmount,
feeValue, discountAmount) in a single useMemo inside the Formik render prop, keyed on
[values.lines, isSameState]. Do not move them out of the render prop yet — just memo them.

## O3 — Fix type safety
In invoiceForm.tsx Props interface, change initialValues: any to
initialValues: InvoiceFormikValues.
Change formikRef from useRef<FormikProps<any>>(null) to
useRef<FormikProps<InvoiceFormikValues>>(null).
In InvoiceFormValues.tsx, change selectedProduct?: any to selectedProduct?: IProduct | null
using the existing product type from src/types.

## O4 — Move toSafeNumber and isNullOrUndefined out of the component
Move toSafeNumber and isNullOrUndefined to src/utils/numberUtils.ts as named exports.
Import them into invoiceForm.tsx. Do not change their logic.

Do NOT:
- Change any UI layout or visual behaviour
- Remove any feature flags or toggles (soNoSoDate, showGSTRates, discount, feesVisible)
- Touch edit.tsx, copy.tsx, or invoiceItem.tsx
- Modify the API layer

After each fix, confirm TypeScript still compiles with: npx tsc --noEmit
```
