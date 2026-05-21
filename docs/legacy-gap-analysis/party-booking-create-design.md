# Party Booking Create — Feature Design

## Screen Reference

**Screenshot:** `Images/CratePartyBooking.jpeg`
**Legacy VB form:** `hitrix-srk/src/Access/frmBookingParty.frm` — caption: "Party Booking Entry"
**Dhanman route:** `Sales → Party Bookings → Create Party Booking`
**Page title:** Create Party Booking
**Breadcrumb:** Home / Party Bookings / Create Party Booking

---

## Business Context

A Party Booking is a **pre-sales commitment** in the yarn/textile trading domain. A customer (party) places a forward booking for a specific yarn count, quantity, and agreed rate before dispatch or invoicing. The booking links to:

- the selling party (customer) and their broker
- the mill supplying the goods
- specific yarn counts, bag quantities, standard packing weights, and rates

The booking later becomes the basis for a Party Invoice and Gate Pass dispatch.

**Legacy equivalent:** `frmBookingParty.frm` in HITRIX.
**Related flow:** `Sales_Purchase_Inventory_Broker_Flow.md` — Section 1.1 Sales Order Creation.

---

## Architecture: Mirror of Invoice Create

This screen must follow **the exact same architecture as `invoiceForm.tsx`**. Every pattern, hook, state variable shape, and component choice is mirrored. The table below maps each invoice element to its party booking equivalent.

| Invoice Create | Party Booking Create | Notes |
|---|---|---|
| `create.tsx` | `create.tsx` | Same thin wrapper pattern |
| `invoiceForm.tsx` | `partyBookingForm.tsx` | Same monolithic form component |
| `InvoiceItem.tsx` | `PartyBookingItem.tsx` | Same per-row component |
| `createEmptyInvoiceLine()` | `createEmptyBookingLine()` | Same empty-row factory |
| `InvoiceFormikValues` | `PartyBookingFormikValues` | Same header type shape |
| `InvoiceLineFormikValues` | `PartyBookingLineFormikValues` | Same line type shape |
| `invoiceValidationSchema` | `partyBookingValidationSchema` | Same Yup schema pattern |
| `insertNewInvoice` / `updateInvoice` | `insertPartyBooking` / `updatePartyBooking` | Same API wrapper pattern |
| `useGetCustomerNames` | `useGetPartyNames` (same customer API) | Party = Customer |
| `useGetAllProducts` | `useGetAllCounts` (products API filtered) | Count = Yarn Item |
| `useGetProductByCustomerId` | `useGetCountsByPartyId` | Optional: party-specific items |
| `selectedCustomer` | `selectedParty` | Same selection state |
| `isCustomerSelected` | `isPartySelected` | Same guard flag |
| `isGetAllCustomers` checkbox | `isGetAllParties` checkbox | Same "show all" toggle |
| `customersWithPrimaryOwner` | `partiesWithBroker` | Same merge-and-enrich pattern |
| `tempCustomer` + `showCustomerChangeDialog` | `tempParty` + `showPartyChangeDialog` | Same change-guard dialog |
| `invoiceHeaderId` | `bookingHeaderId` | Same post-save ID tracking |
| `deleteIndex` + `isDeleteDialogOpen` | `deleteIndex` + `isDeleteDialogOpen` | Same delete confirmation pattern |
| `handleCreate` async fn | `handleCreate` async fn | Same submit handler shape |
| `soNoSoDate` switch | Not a switch — `refNo`/`refDate` always visible | Visible in screenshot |
| `showGSTRates` switch | No GST columns — `isRateExcludingGST` checkbox instead | GST set once at header |
| `discount` switch | No discount column (not in Party Booking) | Removed |
| `feesVisible` switch | No fees column (not in Party Booking) | Removed |
| `ItemsSummary` component | New `PartyBookingSummary` inline block | Bag / Weight / Amount |
| `WareHouse` component | Not used | No warehouse in booking |
| `useGetOrganizationAccounts` | Not used | No accounting at booking stage |
| `useGetCurrencies` | Not used | No multi-currency in booking |
| `compareGST` / `isSameState` | Not used | GST routing not on booking form |
| `useDateFilter(finYear)` | `useDateFilter(finYear)` | Same — used for edit/update API |
| `useBackdrop` | `useBackdrop` | Same loading overlay |
| `useAuth` + `insertUserActivity` | `useAuth` + `insertUserActivity` | Same audit trail |
| `SuccessToastMessage` / `ToastMessage` | `SuccessToastMessage` / `ToastMessage` | Same toast pattern |
| `useConfig` | `useConfig` | Same company/finYear |
| `PageHeader` | `PageHeader` | Same |
| `MainCard` | `MainCard` | Same |
| `MainFooter` | `MainFooter` | Same |
| `BDate` | `BDate` | Same date picker |
| `BInputLabel` | `BInputLabel` | Same label |
| `LoadingAutocomplete` | `LoadingAutocomplete` | Same loading autocomplete |
| `VirtualizedListbox` | `VirtualizedListbox` | Same — for Party dropdown |
| `RenderMultipleOption` + `customPaperComponent` | `RenderMultipleOption` + `customPaperComponent` | Same dropdown rendering |
| `searchByFields` | `searchByFields` | Same filter |
| `AlertProductDelete` | `AlertProductDelete` | Same delete dialog |

---

## File Structure

```
src/pages/sales/partyBooking/
├── create.tsx                          ← thin wrapper (same as invoice/create.tsx)
├── edit.tsx                            ← edit wrapper
├── list.tsx                            ← list with AG Grid
├── view.tsx                            ← read-only detail
├── partyBookingForm.tsx                ← main form (mirrors invoiceForm.tsx)
├── PartyBookingItem.tsx                ← line-item row (mirrors InvoiceItem.tsx)
├── partyBookingColDefs.tsx             ← AG Grid col defs for list
├── type/
│   ├── PartyBookingFormValues.tsx      ← types (mirrors InvoiceFormValues.tsx)
│   └── partyBookingUtils.tsx           ← createEmptyBookingLine helper
└── schemas/
    └── partyBookingSchemas.ts          ← Yup schema (mirrors invoiceSchemas.ts)

src/api/sales/
└── partyBooking.ts                     ← API wrapper (mirrors invoice.ts)
```

---

## Data Model

### PartyBookingFormikValues (header)

Mirrors `InvoiceFormikValues`. Fields that have no invoice equivalent are marked **NEW**.

| Field | Type | Invoice Equivalent | Notes |
|---|---|---|---|
| `id` | `string` | `id` | Empty on create |
| `bookingNo` | `number` | `invoiceNumber` | Auto-incremented; read-only after save |
| `bookingDate` | `Date \| null` | `invoiceDate` | Required |
| `partyGroupId` | `string` | — | **NEW**: Party group master |
| `partyInfo` | `PartyInfoForm` | `customerInfo` | `{ id, name, gstIn }` |
| `brokerId` | `string` | — | **NEW**: Broker master |
| `millId` | `string` | — | **NEW**: Mill master |
| `ratePer` | `'IncludingGST' \| 'ExcludingGST'` | — | **NEW**: Header-level GST mode |
| `refNo` | `number` | `soNo` | Reference number |
| `refDate` | `Date \| null` | `soDate` | Reference date |
| `printTicketName` | `boolean` | — | **NEW**: Print option |
| `printExport` | `boolean` | — | **NEW**: Export print option |
| `isRateExcludingGST` | `boolean` | — | **NEW**: Override GST mode |
| `deliveryPlace` | `string` | — | **NEW**: Delivery location |
| `side` | `number` | — | **NEW**: Side / lot number |
| `paymentTerms` | `string` | `note` | Required free-text (invoice has textarea) |
| `remark` | `string` | `note` | Internal note |
| `lines` | `PartyBookingLineFormikValues[]` | `lines` | Min one line |
| `totalBag` | `number` | — | Derived |
| `totalWeight` | `number` | — | Derived |
| `totalAmount` | `number` | `totalAmount` | Derived |
| `type` | `BookingType` | `InvoiceBillType` | Booking type enum |

### PartyBookingLineFormikValues (per line)

Mirrors `InvoiceLineFormikValues`.

| Field | Type | Invoice Equivalent | Notes |
|---|---|---|---|
| `id` | `string` | `id` | UUID row key |
| `countId` | `string` | `productId` | Yarn item ID |
| `selectedCount` | `IYarnCount \| null` | `selectedProduct` | Full object |
| `bag` | `number \| string` | `quantity` | Physical bag count |
| `stdPacking` | `number \| string` | — | **NEW**: Kg per bag; defaults from item master |
| `weight` | `number` | `taxableAmount` | Derived: `bag × stdPacking` |
| `rate` | `number` | `price` | Rate per unit |
| `rtPer` | `string` | — | **NEW**: Rate unit (e.g., `1 KG`) |
| `amount` | `number` | row total | Derived: `weight × rate` |
| `isComplete` | `number` | — | **NEW**: Fulfilment flag (0 = open, 1 = done) |
| `serialNumber` | `number` | `serialNumber` | Display order |

---

## Component Internals: partyBookingForm.tsx

### Props (mirrors InvoiceForm Props)

```tsx
interface Props {
  initialValues: PartyBookingFormikValues;
  isInEditMode: boolean;
  isInCopyMode?: boolean;
}
```

### State Variables (mirrors invoiceForm.tsx state block)

```tsx
// Party selection — mirrors customer selection block
const [isGetAllParties, setIsGetAllParties] = useState(false);
const [isPartySelected, setIsPartySelected] = useState(false);
const [selectedParty, setSelectedParty] = useState<IParty | null>(null);
const [partiesWithBroker, setPartiesWithBroker] = useState<IParty[]>([]);
const [tempParty, setTempParty] = useState<IParty | null>(null);
const [showPartyChangeDialog, setShowPartyChangeDialog] = useState(false);

// Delete dialog — identical to invoice
const [deleteIndex, setDeleteIndex] = useState<number | null>(null);
const [isDeleteDialogOpen, setIsDeleteDialogOpen] = useState(false);

// Post-save ID tracking — mirrors invoiceHeaderId
const [bookingHeaderId, setBookingHeaderId] = useState<string | null>('');
```

### Hooks (mirrors invoiceForm.tsx hook block)

```tsx
const { user } = useAuth();
const { showBackdrop, hideBackdrop } = useBackdrop();
const { company, finYear } = useConfig();
const navigate = useNavigate();
const { fromDate, toDate } = useDateFilter(finYear);

// Party list — mirrors useGetCustomerNames
const { partyNames, partyNamesLoading } = useGetPartyNames(company?.id, isGetAllParties);

// Count (yarn item) list — mirrors useGetAllProducts
const { counts, countsLoading } = useGetAllCounts(company?.id);

// Default booking status — mirrors useGetDefaultInvoiceStatus
const { defaultBookingStatus } = useGetDefaultBookingStatus(company?.id);
```

### createEmptyBookingLine (mirrors createEmptyInvoiceLine)

```tsx
const createEmptyBookingLine = (serialNumber: number = 1) => ({
  id: UIDV4(),
  selectedCount: null,
  countId: '',
  bag: '',
  stdPacking: '',
  weight: 0,
  rate: '',
  rtPer: '1 KG',
  amount: 0,
  isComplete: 0,
  serialNumber,
});
```

### Derived Totals (mirrors inline reduce calculations)

Wrap in `useMemo` keyed on `[values.lines]` — same pattern as recommended O1 fix in invoice:

```tsx
const { totalBag, totalWeight, totalAmount } = useMemo(() => {
  const activeLine = (line: any) => line.selectedCount !== null;
  return {
    totalBag:    values.lines.filter(activeLine).reduce((s, l) => s + toSafeNumber(l.bag), 0),
    totalWeight: values.lines.filter(activeLine).reduce((s, l) => s + toSafeNumber(l.weight), 0),
    totalAmount: values.lines.filter(activeLine).reduce((s, l) => s + toSafeNumber(l.amount), 0),
  };
}, [values.lines]);
```

### handleCreate (mirrors invoice handleCreate shape)

```tsx
const handleCreate = async (values: PartyBookingFormikValues) => {
  try {
    showBackdrop();
    const newBooking: IPartyBookingHeader = {
      bookingHeaderId: isInCopyMode ? '' : bookingHeaderId || '',
      companyId: company?.id,
      partyId: values.partyInfo?.id,
      brokerId: values.brokerId,
      millId: values.millId,
      partyGroupId: values.partyGroupId,
      bookingDate: format(new Date(values.bookingDate!), 'yyyy-MM-dd'),
      ratePer: values.ratePer,
      refNo: values.refNo,
      refDate: values.refDate ? format(new Date(values.refDate), 'yyyy-MM-dd') : '',
      printTicketName: values.printTicketName,
      printExport: values.printExport,
      isRateExcludingGST: values.isRateExcludingGST,
      deliveryPlace: values.deliveryPlace,
      side: values.side,
      paymentTerms: values.paymentTerms,
      remark: values.remark,
      bookingStatusId: defaultBookingStatus?.initialStatus,
      totalBag,
      totalWeight,
      totalAmount,
      lines: values.lines
        .filter((line) => line.selectedCount)
        .map((line, index) => ({
          id: isInCopyMode ? UIDV4() : line.id || UIDV4(),
          countId: line.selectedCount?.id,
          bag: toSafeNumber(line.bag),
          stdPacking: toSafeNumber(line.stdPacking),
          weight: toSafeNumber(line.bag) * toSafeNumber(line.stdPacking),
          rate: toSafeNumber(line.rate),
          rtPer: line.rtPer,
          amount: toSafeNumber(line.bag) * toSafeNumber(line.stdPacking) * toSafeNumber(line.rate),
          isComplete: line.isComplete,
          serialNumber: index + 1,
        })),
    };

    let response = null;
    if (isInEditMode && !isInCopyMode) {
      response = await updatePartyBooking(newBooking, company.id, finYear.financeYearId, fromDate, toDate);
    } else {
      response = await insertPartyBooking(newBooking);
    }

    if (response !== null) {
      SuccessToastMessage('bookingConfirmed');
      insertUserActivity(notificationValues);
      setBookingHeaderId(response.id || newBooking.bookingHeaderId);
      hideBackdrop();
      navigate('/sales/party-bookings/list');
    }
  } catch (error) {
    hideBackdrop();
    console.error('Error creating/updating booking:', error);
  }
};
```

---

## Form Layout (mirrors invoiceForm.tsx JSX structure)

```
<PageHeader showFilters={false} />
<MainCard>
  <Formik innerRef={formikRef} ...>
    {({ ... }) => (
      <Form onSubmit={handleSubmit}>
        <Grid container spacing={2}>

          {/* Row 1: Party Group | Party | [Show All checkbox] | Broker | Mill */}
          {/* Same layout as invoice: Customer | [checkbox] | Sales Account | Currency */}
          <Grid size={{ xs:12, sm:6, md:3 }}>   ← Party (required, VirtualizedListbox)
          <Grid mt={2.5} size={{ xs:12, sm:1, md:1 }}>  ← Show All Parties checkbox
          <Grid size={{ xs:12, sm:4, md:2 }}>   ← Broker (Autocomplete)
          <Grid size={{ xs:12, sm:4, md:2 }}>   ← Mill (Autocomplete)
          <Grid size={{ xs:12, sm:4, md:2 }}>   ← Party Group (Autocomplete)

          {/* Row 2: Booking No + Date | Rate Per | Ref No + Ref Date | [Back button] */}
          {/* Mirrors: Invoice Number | Invoice Date | Due Date | [Back button] */}
          <Grid size={{ xs:12, sm:4, md:1.8 }}>  ← Booking No (TextField number)
          <Grid size={{ xs:12, sm:4, md:1.8 }}>  ← Booking Date (BDate)
          <Grid size={{ xs:12, sm:4, md:2 }}>    ← Rate Per (Select dropdown)
          <Grid size={{ xs:12, sm:4, md:1.5 }}>  ← Ref No (TextField number)
          <Grid size={{ xs:12, sm:4, md:1.8 }}>  ← Ref Date (BDate)
          <Grid size={{ xs:12, sm:6, md:3 }}>    ← Back button (top right)

          {/* Row 3: Checkboxes (replaces invoice Switches row) */}
          {/* Mirrors the Switches row: SO NO | GST Rates | Discount | Fees */}
          <Grid size={12}>
            <FormControlLabel control={<Checkbox>} label="Print Ticket Name?" />
            <FormControlLabel control={<Checkbox>} label="Print Export?" />
            <FormControlLabel control={<Checkbox>} label="Is Rate Excluding GST?" />

          {/* FieldArray — identical structure to invoice */}
          <Grid size={12}>
            <FieldArray name="lines" render={({ remove, push }) => (
              <TableContainer>
                <Table size="small">
                  <TableHead>
                    <TableRow>
                      <TableCell>#</TableCell>
                      <TableCell>Count</TableCell>    ← replaces Name
                      <TableCell>Bag</TableCell>       ← replaces Qty
                      <TableCell>Std Packing</TableCell> ← NEW
                      <TableCell>Weight</TableCell>    ← replaces Taxable Amt (derived)
                      <TableCell>Rate</TableCell>      ← same
                      <TableCell>Rt Per</TableCell>    ← NEW
                      <TableCell>Amount</TableCell>    ← replaces Row Total
                      <TableCell>Is Complete(Y)</TableCell> ← NEW
                      <TableCell>Actions</TableCell>
                    </TableRow>
                  </TableHead>
                  <TableBody>
                    {values.lines.map((item, index) => (
                      <TableRow key={item.id}>
                        <TableCell>
                          {index + 1}
                          <Button color="error" onClick={() => handleDeleteAction(index)}>
                            <DeleteOutlined />
                          </Button>
                          {isDeleteDialogOpen && <AlertProductDelete ... />}
                        </TableCell>
                        <PartyBookingItem key={item.id} index={index} ... />
                      </TableRow>
                    ))}
                  </TableBody>
                </Table>
              </TableContainer>
            )} />

          {/* Footer: Notes left + Summary right — mirrors invoice footer */}
          <Grid container justifyContent="space-between">
            <Grid size={{ xs:6, sm:6, md:8 }}>  ← Left: Delivery Place, Side, Payment Terms, Remark
            <Grid size={{ xs:6, sm:6, md:4 }}>  ← Right: PartyBookingSummary (Bag/Weight/Amount)

        </Grid>

        {/* Party Change Dialog — identical to customer change dialog in invoice */}
        <Dialog open={showPartyChangeDialog} ...>
          <DialogTitle>Confirm Party Change</DialogTitle>
          <DialogContent>Changing the party will clear selected counts. Continue?</DialogContent>
          <DialogActions>
            <Button onClick={() => setShowPartyChangeDialog(false)}>No</Button>
            <Button color="error" variant="contained" onClick={confirmPartyChange}>Yes</Button>
          </DialogActions>
        </Dialog>

        {/* MainFooter — identical to invoice */}
        <MainFooter>
          <Button variant="outlined" onClick={handleBackToList}>Cancel</Button>
          <Button color="primary" variant="contained" type="submit" disabled={isSubmitting}>
            {isInCopyMode ? 'Copy & Create' : isInEditMode ? 'Update Booking' : 'Confirm Booking'}
          </Button>
        </MainFooter>
      </Form>
    )}
  </Formik>
</MainCard>
```

---

## PartyBookingItem.tsx (mirrors InvoiceItem.tsx)

The row component renders these cells in order:

| Cell | Component | Derivation |
|---|---|---|
| Count | `Autocomplete` with `VirtualizedListbox` | User selects; auto-fills `stdPacking` from item master |
| Bag | `SalesOrderField` (type=number) | User enters |
| Std Packing | `SalesOrderField` (type=number) | Auto-filled from Count, editable |
| Weight | Read-only `Typography` | `bag × stdPacking` — recalculates on Bag/StdPacking change |
| Rate | `SalesOrderField` (type=number) | User enters |
| Rt Per | `Select` with options `['1 KG', '1 BAG', '1 MT']` | User selects |
| Amount | Read-only `Typography` | `weight × rate` |
| Is Complete | `SalesOrderField` (type=number, 0 or 1) | User enters |
| Actions | Delete `Button` + Copy row `Button` | Same as invoice |

On pressing Tab/Enter in the **last row's** Amount cell → `handleAddItem` pushes a new empty line (same as `handleAddItem` in invoice).

---

## PartyBookingSummary (replaces ItemsSummary)

A simple inline block placed in the footer right column — no separate component file needed initially:

```tsx
<Stack spacing={1} alignItems="flex-end" sx={{ pr: 2 }}>
  <Typography>Total Bag: <strong>{totalBag}</strong></Typography>
  <Typography>Total Weight: <strong>{totalWeight}</strong></Typography>
  <Divider sx={{ width: '100%' }} />
  <Typography variant="h6">
    Total Amount: <strong>{selectedCurrencySymbol} {addCommas(totalAmount)}</strong>
  </Typography>
</Stack>
```

---

## create.tsx (thin wrapper — identical pattern)

```tsx
const CreatePartyBooking = () => {
  const { company } = useConfig();
  const initialValues: PartyBookingFormikValues = {
    id: '',
    bookingNo: 0,
    bookingDate: new Date(),
    partyGroupId: '',
    partyInfo: { id: '', name: '', gstIn: '' },
    brokerId: '',
    millId: '',
    ratePer: 'IncludingGST',
    refNo: 0,
    refDate: new Date(),
    printTicketName: false,
    printExport: false,
    isRateExcludingGST: false,
    deliveryPlace: '',
    side: 0,
    paymentTerms: '',
    remark: '',
    lines: [createEmptyBookingLine(1)],
    totalBag: 0,
    totalWeight: 0,
    totalAmount: 0,
    type: BookingType.Party,
  };

  return (
    <PartyBookingForm
      initialValues={initialValues}
      isInEditMode={false}
    />
  );
};
```

---

## API Wrapper: partyBooking.ts (mirrors invoice.ts)

```ts
// insertPartyBooking(booking)     → POST /party-bookings
// updatePartyBooking(booking, ...) → PUT  /party-bookings/:id
// useGetPartyBookings(companyId)  → SWR GET /party-bookings?companyId=...
// useGetPartyBookingById(id)      → SWR GET /party-bookings/:id
```

Follow the exact Axios + SWR fetcher pattern in `src/api/sales/invoice.ts`.

---

## Routes and Menu

In `src/routes/modules/sales-routes.tsx` (lazy load, same as invoice routes):
```tsx
{ path: 'party-bookings/list',     element: <PartyBookingList /> }
{ path: 'party-bookings/create',   element: <CreatePartyBooking /> }
{ path: 'party-bookings/:id',      element: <ViewPartyBooking /> }
{ path: 'party-bookings/:id/edit', element: <EditPartyBooking /> }
```

In `src/menu-items/sales.tsx` add under Sales group (same permission pattern):
```tsx
{ id: 'party-bookings', title: 'Party Bookings', url: '/sales/party-bookings/list',
  permissions: [DhanmanPermissions.PartyBooking_View] }
```

---

## Open Questions

- Does the API auto-generate `bookingNo`, or is it user-entered like invoice number?
- Is `Party` the same master as `Customer` (same API) or a separate entity?
- Should `stdPacking` auto-fill from the item/count master when Count is selected?
- What statuses does a booking use (Draft / Confirmed / Fulfilled / Cancelled)?
- Is one booking always converted to exactly one Party Invoice, or can it split?
- Does `isComplete` on a line update automatically from Gate Pass, or is it manual?

---

## Implementation Prompt

Use the following prompt with Claude Code to build the Party Booking Create screen:

```
Build the Party Booking Create screen in dhanman-app.

Architecture reference:
  hitrix-srk/docs/legacy-gap-analysis/party-booking-create-design.md

Screenshot reference:
  Images/CratePartyBooking.jpeg

This screen must follow the EXACT SAME architecture as the Invoice Create screen.
Before writing any code, read:
  src/pages/sales/invoice/create.tsx
  src/pages/sales/invoice/invoiceForm.tsx
  src/pages/sales/invoice/type/InvoiceFormValues.tsx
  src/pages/sales/invoice/schemas/invoiceSchemas.ts
  src/api/sales/invoice.ts

Then build each file below using those as the direct pattern source.

## Step 1 — Types
Create src/pages/sales/partyBooking/type/PartyBookingFormValues.tsx
  Mirror InvoiceFormValues.tsx exactly. Replace:
    customerInfo → partyInfo (same shape: id, name, gstIn)
    invoiceNumber → bookingNo (number)
    invoiceDate → bookingDate
    dueDate → refDate
    soNo → refNo (number)
    soDate → refDate
    creditAccountId / debitAccountId → remove (no accounting)
    sourceWarehouseId / destinationWarehouseId → remove
    currencyId → remove
    cgstAmount / sgstAmount / igstAmount → remove
    roundOff → remove
    note → split into paymentTerms (required) and remark
  Add new header fields: partyGroupId, brokerId, millId, ratePer, side,
    deliveryPlace, printTicketName, printExport, isRateExcludingGST,
    totalBag, totalWeight
  For line type, replace:
    quantity → bag
    price → rate
    productId / selectedProduct → countId / selectedCount
    description → remove
    discount / fees / cgst / sgst / igst / cgstAmount / sgstAmount / igstAmount → remove
    taxableAmount → weight (derived: bag × stdPacking)
  Add new line fields: stdPacking, rtPer, amount (derived: weight × rate), isComplete

Create src/pages/sales/partyBooking/type/partyBookingUtils.tsx with:
  createEmptyBookingLine(serialNumber = 1) — same shape as createEmptyInvoiceLine

## Step 2 — Validation schema
Create src/pages/sales/partyBooking/schemas/partyBookingSchemas.ts
  Mirror invoiceSchemas.ts. Required: partyId (inside partyInfo.id), bookingDate,
  ratePer, refNo, refDate, side, paymentTerms.
  Lines: min 1 item where selectedCount is not null, bag > 0, stdPacking > 0, rate > 0.

## Step 3 — API wrapper
Create src/api/sales/partyBooking.ts
  Mirror src/api/sales/invoice.ts exactly.
  Functions: insertPartyBooking, updatePartyBooking, useGetPartyBookings,
  useGetPartyBookingById.
  Use placeholder endpoint paths for now (e.g., '/party-bookings').

## Step 4 — PartyBookingItem row component
Create src/pages/sales/partyBooking/PartyBookingItem.tsx
  Mirror InvoiceItem.tsx structure but with these columns:
    Count (Autocomplete, same pattern as product autocomplete in InvoiceItem)
    Bag (SalesOrderField type=number — import from pages/sales/components/SaleOrder.tsx)
    Std Packing (SalesOrderField type=number)
    Weight (read-only Typography: toSafeNumber(bag) * toSafeNumber(stdPacking))
    Rate (SalesOrderField type=number)
    Rt Per (Select with options ['1 KG', '1 BAG', '1 MT'])
    Amount (read-only Typography: weight * rate)
    Is Complete (SalesOrderField type=number)
  handleAddItem on last row tab — same logic as InvoiceItem.

## Step 5 — Main form component
Create src/pages/sales/partyBooking/partyBookingForm.tsx
  Mirror invoiceForm.tsx section by section:

  Props: { initialValues: PartyBookingFormikValues, isInEditMode: boolean, isInCopyMode?: boolean }

  State block (mirror invoiceForm.tsx state block):
    isGetAllParties / setIsGetAllParties (bool) — same as isGetAllCustomers
    isPartySelected / setIsPartySelected (bool) — same as isCustomerSelected
    selectedParty / setSelectedParty — same as selectedCustomer
    tempParty / setTempParty — same as tempCustomer
    showPartyChangeDialog / setShowPartyChangeDialog — same as showCustomerChangeDialog
    deleteIndex / setDeleteIndex — identical
    isDeleteDialogOpen / setIsDeleteDialogOpen — identical
    bookingHeaderId / setBookingHeaderId — same as invoiceHeaderId

  Hooks block (mirror invoiceForm.tsx hooks):
    useAuth, useBackdrop, useConfig, useNavigate, useDateFilter — identical
    useGetPartyNames(company?.id, isGetAllParties) — mirrors useGetCustomerNames
    useGetAllCounts(company?.id) — mirrors useGetAllProducts
    useGetDefaultBookingStatus(company?.id) — mirrors useGetDefaultInvoiceStatus

  createEmptyBookingLine — same placement as createEmptyInvoiceLine in invoiceForm

  toSafeNumber — same utility (copy from invoiceForm until extracted)

  Derived totals useMemo([values.lines]):
    totalBag = sum of line.bag for lines with selectedCount
    totalWeight = sum of (bag × stdPacking) for lines with selectedCount
    totalAmount = sum of (weight × rate) for lines with selectedCount
    Do NOT mutate values directly. Only use in display and submit payload.

  handleCreate — same shape as invoice handleCreate (showBackdrop, build payload,
    call insertPartyBooking or updatePartyBooking, SuccessToastMessage,
    insertUserActivity, setBookingHeaderId, hideBackdrop, navigate)

  handlePartySelected — same as handleCustomerSelected

  JSX structure (mirror invoiceForm.tsx JSX section by section):
    <PageHeader showFilters={false} />
    <MainCard>
      <Formik innerRef={formikRef} enableReinitialize={isInEditMode} ...>
        Row 1: Party (VirtualizedListbox, same onChange party-change guard) |
                Show All Parties checkbox | Broker Autocomplete |
                Mill Autocomplete | Party Group Autocomplete
        Row 2: Booking No (TextField number) + Booking Date (BDate) |
                Rate Per (Select: IncludingGST/ExcludingGST) |
                Ref No (TextField number) + Ref Date (BDate) |
                Back button (top right, same position as invoice)
        Row 3: Three Checkbox controls (Print Ticket Name, Print Export,
                Is Rate Excluding GST) — replaces the Switch row in invoice
        FieldArray name="lines" — identical structure to invoice FieldArray:
          TableContainer > Table > TableHead + TableBody
          Each TableRow: serial+delete button cell | <PartyBookingItem> cells
          handleAddItem / handleDeleteAction / handleDeleteDialogClose — copy exactly
        Footer row:
          Left Grid (md=8): Delivery Place TextField | Side * TextField |
            Payment Terms * TextField multiline | Remark TextField multiline
          Right Grid (md=4): totalBag / totalWeight / totalAmount summary block
        Party change Dialog — copy customer change dialog, change text to party
        MainFooter: Cancel button | Confirm Booking / Update Booking / Copy & Create

## Step 6 — create.tsx
Create src/pages/sales/partyBooking/create.tsx
  Mirror invoice/create.tsx exactly. Set initialValues with createEmptyBookingLine(1).
  Render <PartyBookingForm initialValues={initialValues} isInEditMode={false} />

## Step 7 — Routes and menu
In src/routes/modules/sales-routes.tsx add lazy routes for list/create/view/edit.
In src/menu-items/sales.tsx add Party Bookings menu item with same permission pattern.

## Constraints — these must all hold
- Every pattern from invoiceForm.tsx applies unchanged unless explicitly replaced above
- Use SalesOrderField (src/pages/sales/components/SaleOrder.tsx) for editable number cells
- Use VirtualizedListbox for the Party dropdown (same as customer dropdown)
- Use BDate for all date fields
- Use BInputLabel for all labels
- Use MainCard, MainFooter, PageHeader exactly as in invoiceForm.tsx
- No direct mutation of Formik values (no values.totalAmount = x pattern)
- No inline styles — sx prop only
- No any types in Props or derived values — use PartyBookingFormikValues throughout
- toSafeNumber used for all numeric reads (same as invoice)

After each step confirm TypeScript still compiles: npx tsc --noEmit
```
