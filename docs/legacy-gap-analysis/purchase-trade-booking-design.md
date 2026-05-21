# Trade Purchase Booking Create — Feature Design

## Screen Reference

**Screenshot:** _Pending — developer to provide_
**Legacy VB form:** `hitrix-srk/src/ABabu_Cd_15-05-2026/frmBookingTradePurch.frm` — caption: "Purchase Entry"
**Also see:** `hitrix-srk/src/Access/frmBookingTrade.frm`
**Dhanman route:** `Purchase → Trade Bookings → Create Trade Booking`
**Page title:** Create Trade Purchase Booking
**Breadcrumb:** Home / Trade Bookings / Create Trade Purchase Booking

> **Note:** Visual layout and field details will be updated once the developer provides the design screenshot.
> Architecture below is pre-filled based on the `party-booking-create-design.md` mirror pattern and HITRIX legacy form evidence.

---

## Business Context

A Trade Purchase Booking is the **purchase-side equivalent of Party Booking**. Where Party Booking records a customer's forward commitment to buy yarn, Trade Purchase Booking records the company's forward commitment to **buy from a mill** at an agreed rate.

The booking links to:
- the mill supplying the goods
- the broker facilitating the trade
- the party (internal or external) on whose behalf the purchase is made
- specific yarn counts, bag quantities, standard packing weights, and agreed rates

It becomes the basis for a Trade Purchase Bill (goods receipt) and inventory inward entry.

**Legacy equivalent:** `frmBookingTradePurch.frm` — "Purchase Entry" in HITRIX.
**Related flow:** `Sales_Purchase_Inventory_Broker_Flow.md` — Purchase side.
**Mirror of:** `party-booking-create-design.md` (sales side).

---

## Architecture: Mirror of Bill Create + Party Booking Pattern

This screen follows the **same architecture as `BillForm.tsx`** for the form shell, and the **same line-item pattern as `partyBookingForm.tsx`** for the table columns (Count / Bag / Std Packing / Weight / Rate / Rt Per / Amount).

| Party Booking (Sales) | Trade Purchase Booking (Purchase) | Notes |
|---|---|---|
| `partyBookingForm.tsx` | `tradePurchaseBookingForm.tsx` | Same monolithic form pattern |
| `PartyBookingItem.tsx` | `TradePurchaseBookingItem.tsx` | Same per-row component |
| `createEmptyBookingLine()` | `createEmptyTradeLine()` | Same empty-row factory |
| `useGetPartyNames` | `useGetVendorNames` / `useGetMillNames` | Mill = supplier |
| Party (customer) | Mill (vendor/supplier) | Direction reversal |
| Broker | Broker | Same — common to both |
| `insertPartyBooking` | `insertTradePurchaseBooking` | Mirror API call |
| `BillForm.tsx` shell | Same shell | Reuse BillForm patterns |

---

## Data Model (Pre-filled — update from screenshot)

### TradePurchaseBookingFormikValues (header)

| Field | Type | Required | Notes |
|---|---|---|---|
| `id` | `string` | — | Empty on create |
| `bookingNo` | `number` | Yes | Auto-incremented; read-only after save |
| `bookingDate` | `Date \| null` | Yes | Booking date |
| `millId` | `string` | Yes | Mill / supplier (purchase-side Party) |
| `brokerId` | `string` | No | Broker master |
| `partyId` | `string` | No | Party on whose behalf purchased |
| `partyGroupId` | `string` | No | Party group |
| `ratePer` | `'IncludingGST' \| 'ExcludingGST'` | Yes | Header GST mode |
| `refNo` | `number` | Yes | Reference number |
| `refDate` | `Date \| null` | Yes | Reference date |
| `printMillCopy` | `boolean` | No | Print — Mill Copy flag |
| `printOfficeCopy` | `boolean` | No | Print — Office Copy flag |
| `deliveryPlace` | `string` | No | Delivery location |
| `side` | `number` | Yes | Side / lot identifier |
| `paymentTerms` | `string` | Yes | Free-text |
| `remark` | `string` | No | Internal note |
| `lines` | `TradePurchaseBookingLineFormikValues[]` | Yes | Min one line |
| `totalBag` | `number` | — | Derived |
| `totalWeight` | `number` | — | Derived |
| `totalAmount` | `number` | — | Derived |

### TradePurchaseBookingLineFormikValues (per line)

Same structure as `PartyBookingLineFormikValues` — direction is purchase not sale:

| Field | Type | Notes |
|---|---|---|
| `id` | `string` | UUID row key |
| `countId` | `string` | Yarn item ID |
| `selectedCount` | `IYarnCount \| null` | Full object |
| `bag` | `number \| string` | Physical bag count |
| `stdPacking` | `number \| string` | Kg per bag; defaults from item master |
| `weight` | `number` | Derived: `bag × stdPacking` |
| `rate` | `number` | Purchase rate |
| `rtPer` | `string` | Rate unit (e.g., `1 KG`) |
| `amount` | `number` | Derived: `weight × rate` |
| `isComplete` | `number` | Fulfilment flag (0 = open, 1 = done) |
| `serialNumber` | `number` | Display order |

---

## Component Architecture

```
CreateTradePurchaseBooking (pages/purchase/tradeBooking/create.tsx)
└── TradePurchaseBookingForm (pages/purchase/tradeBooking/tradePurchaseBookingForm.tsx)
    ├── PageHeader
    ├── MainCard
    │   └── Formik (innerRef={formikRef})
    │       ├── Row 1: Party Group | Mill * | [Show All checkbox] | Broker | Party
    │       ├── Row 2: Booking No + Date | Rate Per | Ref No + Ref Date | Back button
    │       ├── Row 3: Checkboxes (Mill Copy, Office Copy, Is Rate Excluding GST)
    │       ├── FieldArray → TradePurchaseBookingItem rows
    │       │   └── TradePurchaseBookingItem (per row)
    │       ├── Footer left: Delivery Place | Side * | Payment Terms * | Remark
    │       └── Footer right: Total Bag / Total Weight / Total Amount
    └── MainFooter
        ├── Cancel → navigate to list
        └── Confirm Booking button
```

---

## File Structure to Create

```
src/pages/purchase/tradeBooking/
├── create.tsx
├── edit.tsx
├── list.tsx
├── view.tsx
├── tradePurchaseBookingForm.tsx
├── TradePurchaseBookingItem.tsx
├── tradeBookingColDefs.tsx
├── type/
│   ├── TradePurchaseBookingFormValues.tsx
│   └── tradeBookingUtils.tsx              ← createEmptyTradeLine()
└── schemas/
    └── tradeBookingSchemas.ts

src/api/purchase/
└── tradeBooking.ts                        ← insertTradePurchaseBooking,
                                              updateTradePurchaseBooking,
                                              useGetTradeBookings,
                                              useGetTradeBookingById
```

---

## Reused Components

| Component | Source | Usage |
|---|---|---|
| `SalesOrderField` | `pages/sales/components/SaleOrder.tsx` | Bag, Std Packing, Rate cells |
| `BDate` | `components/B2A/BDate` | Booking Date, Ref Date |
| `BInputLabel` | `components/B2A/BInputLabel` | All labels |
| `LoadingAutocomplete` | `components/B2A/LoadingAutoComplete` | Mill, Broker, Count dropdowns |
| `VirtualizedListbox` | `pages/sales/invoice/component/VirtualizedListbox` | Mill dropdown |
| `AlertBillProductDelete` | `pages/purchase/components/AlertBillProductDelete` | Delete line confirmation |
| `PageHeader` | `layout/CommonLayout/PageHeader` | Page header |
| `MainCard` | `components/MainCard` | Form container |
| `MainFooter` | `components/MainFooter` | Footer action bar |

---

## Key Differences vs Party Booking

| Aspect | Party Booking (Sales) | Trade Purchase Booking (Purchase) |
|---|---|---|
| Direction | Customer commits to buy | Company commits to buy from mill |
| Primary party | Party (customer) | Mill (supplier) |
| Print options | Print Ticket Name / Print Export | Mill Copy / Office Copy |
| API target | Sales booking endpoint | Purchase booking endpoint |
| Post-booking document | Party Invoice | Trade Purchase Bill |
| Accounting entry | None at booking stage | None at booking stage |

---

## Derived Totals Logic (same as Party Booking)

```tsx
const { totalBag, totalWeight, totalAmount } = useMemo(() => {
  const active = (l: any) => l.selectedCount !== null;
  return {
    totalBag:    values.lines.filter(active).reduce((s, l) => s + toSafeNumber(l.bag), 0),
    totalWeight: values.lines.filter(active).reduce((s, l) => s + toSafeNumber(l.weight), 0),
    totalAmount: values.lines.filter(active).reduce((s, l) => s + toSafeNumber(l.amount), 0),
  };
}, [values.lines]);
```

---

## Routes and Menu

In `src/routes/modules/purchase-routes.tsx`:
```tsx
{ path: 'trade-bookings/list',      element: <TradeBookingList /> }
{ path: 'trade-bookings/create',    element: <CreateTradePurchaseBooking /> }
{ path: 'trade-bookings/:id',       element: <ViewTradePurchaseBooking /> }
{ path: 'trade-bookings/:id/edit',  element: <EditTradePurchaseBooking /> }
```

In `src/menu-items/purchase.tsx`:
```tsx
{ id: 'trade-bookings', title: 'Trade Bookings', url: '/purchase/trade-bookings/list',
  permissions: [DhanmanPermissions.TradeBooking_View] }
```

---

## Open Questions

- Does the screenshot show Mill as the primary required field (equivalent to Party on the sales side)?
- Are print options "Mill Copy" and "Office Copy" checkboxes — confirmed from `frmBookingTradePurch.frm`?
- Is the line table identical to Party Booking (Count / Bag / Std Packing / Weight / Rate / Rt Per / Amount / Is Complete)?
- Does `isComplete` update automatically from Purchase Bill inward, or manually?
- Is booking number shared with Party Booking sequence, or a separate series?

---

## Implementation Prompt

_Will be completed once the developer provides the design screenshot and open questions above are answered._

```
Build the Trade Purchase Booking Create screen in dhanman-app.

Design reference:
  hitrix-srk/docs/legacy-gap-analysis/purchase-trade-booking-design.md

This screen mirrors party-booking-create-design.md on the purchase side.
Before writing any code, read:
  src/pages/sales/partyBooking/partyBookingForm.tsx    ← sales-side mirror
  src/pages/purchase/bills/BillForm.tsx                ← purchase shell pattern
  src/pages/purchase/bills/types/BillFormValues.ts

Then build each file under src/pages/purchase/tradeBooking/ using those as
the direct pattern source.

Follow the same 7-step build order as party-booking-create-design.md:
  Step 1 — Types
  Step 2 — Validation schema
  Step 3 — API wrapper (src/api/purchase/tradeBooking.ts)
  Step 4 — TradePurchaseBookingItem row component
  Step 5 — tradePurchaseBookingForm.tsx (main form)
  Step 6 — create.tsx (thin wrapper)
  Step 7 — Routes and menu

Key differences from Party Booking:
  - Primary party is Mill (vendor/supplier), not Party (customer)
  - Use useGetVendorNames or useGetMillNames instead of useGetPartyNames
  - Print checkboxes: Mill Copy + Office Copy (not Print Ticket Name + Print Export)
  - API target: src/api/purchase/tradeBooking.ts (not sales)
  - AlertBillProductDelete instead of AlertProductDelete for line delete
  - Navigate to /purchase/trade-bookings/list on submit

After each step confirm TypeScript compiles: npx tsc --noEmit
```
