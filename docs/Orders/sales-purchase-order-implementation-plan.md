# Sales Order And Purchase Order Implementation Plan

## Purpose

Implement Sales Order and Purchase Order in Dhanman based on the Hitrix booking/order screens and database structure.

Hitrix treats these records as a pre-transaction booking layer. No ledger posting should happen at this stage. Sales Orders can later convert to Sales Invoices, and Purchase Orders can later convert to Bills.

## Hitrix References

- Booking/order documentation: `hitrix-srk/docs/02-booking.md`
- Sales order screen: `hitrix-srk/docs/Orders/Sales order.jpeg`
- Purchase order screen: `hitrix-srk/docs/Orders/Purchase order.jpeg`
- Hitrix database script: `hitrix-srk/docs/Orders/script.sql`

## Hitrix Source Tables

The main Hitrix tables used for Sales Order and Purchase Order are:

| Hitrix Table | Purpose | Dhanman Mapping |
|---|---|---|
| `tblBooking` | Header/order-level booking data | Order header table |
| `tblBooKingSub` | Item/detail-level booking data | Order detail table |

## Hitrix Screen Mapping

| Hitrix Screen | Business Meaning | New Dhanman Feature |
|---|---|---|
| Party Booking | Sales booking/order from customer | Sales Order |
| Trade Purchase Booking | Purchase booking/order from supplier/mill | Purchase Order |

## Header Field Mapping

| Hitrix Field | Meaning | Suggested Dhanman Field |
|---|---|---|
| `VNo` | Booking/order number | `order_number` |
| `VType` | Booking type | `order_type` |
| `Vdt` | Order date | `order_date` |
| `VYear` | Financial year | `fin_year` |
| `VFirm` | Firm/company code | `company_id` / legacy reference |
| `BkParty` | Customer/party | `customer_id` |
| `BkMillCode` | Mill/vendor | `vendor_id` or `mill_id` |
| `BkBroker` | Broker | `broker_id` |
| `BkRefNo` | Reference/contract number | `reference_number` / `contract_number` |
| `BkRefDt` | Reference/contract date | `reference_date` / `contract_date` |
| `BkSubAmt` | Sub amount | `subtotal_amount` |
| `BkBillAmt` | Bill/order amount | `total_amount` |
| `BkRoff` | Round off | `round_off` |
| `DueDays` | Payment term days | `payment_term` |
| `DueDate` | Due date | `due_date` |
| `BkItCode` | Item/product code | `product_id` if header-level item is needed |
| `BkBag` | Total bag | `total_bag` |
| `BkWt` | Total weight | `total_weight` |
| `BkRt` | Rate | `rate` |
| `BkRtPer` | Rate per | `rate_per` |
| `BkRateType` | Rate type | `rate_type` |
| `BkDeleAdd` | Delivery address id | `delivery_address_id` |
| `BkPartyGroup` | Party group | `party_group_id` or legacy reference |
| `BkExtraCount` | Extra count | `extra_count` |
| `BkNote` | Remarks | `note` |
| `BkCbag` | Completed bag | `completed_bag` |
| `BkCWt` | Completed weight | `completed_weight` |
| `BkCDate` | Completed date | `completed_date` |
| `BkShDate` | Schedule/ship date | `schedule_date` |
| `BkShBag` | Scheduled/shipped bag | `scheduled_bag` |
| `BkShWt` | Scheduled/shipped weight | `scheduled_weight` |
| `BkIsDirectPayment` | Direct payment flag | `is_direct_payment` |
| `BkIsFrghtInsu` | Freight insurance flag | `is_freight_insurance` |
| `BkIsInsu` | Insurance flag | `is_insurance` |
| `BkMailSendDt` | Broker mail/message date | `broker_message_sent_on` |
| `BkmailSendPartryDt` | Party mail/message date | `party_message_sent_on` |

## Detail Field Mapping

| Hitrix Field | Meaning | Suggested Dhanman Field |
|---|---|---|
| `VNo` | Booking/order number | Header relation |
| `VType` | Booking type | Header relation |
| `Vdt` | Order date | Header relation |
| `VYear` | Financial year | Header relation |
| `VFirm` | Firm/company code | Header relation |
| `BkSubItSrNo` | Line serial number | `serial_number` |
| `BkSubBag` | Bag | `bag` |
| `BkSubWt` | Weight | `weight` |
| `BkSubDDate` | Delivery date | `delivery_date` |
| `BkSubItCode` | Item/product code | `product_id` |
| `BkSubDespDt` | Dispatch date | `dispatch_date` |
| `BkSubDespBag` | Dispatched bag | `dispatch_bag` |
| `BkSubDespWt` | Dispatched weight | `dispatch_weight` |
| `BkSubTrnpt` | Transport | `transport_id` |
| `BkSubDespVno` | Dispatch voucher number | `dispatch_voucher_number` |
| `BkSubDespVyear` | Dispatch financial year | `dispatch_fin_year` |

## Tables To Create

### Sales Module

Create these tables in the Sales service/database:

```text
sales_order_headers
sales_order_details
```

### Purchase Module

Create these tables in the Purchase service/database:

```text
purchase_order_headers
purchase_order_details
```

## Suggested Sales Order Header Columns

```text
id
company_id
customer_id
broker_id
mill_id
order_number
order_date
fin_year
reference_number
reference_date
payment_term
due_date
rate_type
rate_per
total_bag
total_weight
subtotal_amount
round_off
total_amount
delivery_address_id
note
status_id
is_rate_excluding_gst
is_export
print_ticket_name
created_on_utc
modified_on_utc
deleted_on_utc
is_deleted
created_by
modified_by
```

## Suggested Sales Order Detail Columns

```text
id
sales_order_header_id
serial_number
product_id
count_name
bag
standard_packing
weight
rate
rate_per
amount
delivery_date
dispatch_date
dispatch_bag
dispatch_weight
transport_id
dispatch_voucher_number
dispatch_fin_year
is_complete
created_on_utc
modified_on_utc
deleted_on_utc
is_deleted
```

## Suggested Purchase Order Header Columns

```text
id
company_id
vendor_id
broker_id
order_number
order_date
fin_year
contract_number
contract_date
rate_type
rate_per
total_bag
total_weight
subtotal_amount
round_off
total_amount
delivery_address_id
payment_terms
note
status_id
created_on_utc
modified_on_utc
deleted_on_utc
is_deleted
created_by
modified_by
```

## Suggested Purchase Order Detail Columns

```text
id
purchase_order_header_id
serial_number
product_id
count_name
bag
standard_packing
weight
rate
rate_per
amount
delivery_date
received_bag
received_weight
is_complete
created_on_utc
modified_on_utc
deleted_on_utc
is_deleted
```

## Broker Reference

Both Sales Order and Purchase Order should store `broker_id`.

Broker should come from the common broker table:

```csharp
public class Broker : Entity, IAuditableEntity, ISoftDeletableEntity
{
    public string Code { get; private set; } = string.Empty;
    public string Name { get; private set; } = string.Empty;
    public string? DisplayName { get; private set; }
    public Guid? AddressId { get; private set; }
    public string? TaxIdPan { get; private set; }
    public string? TaxIdGstin { get; private set; }
    public bool IsActive { get; private set; }
}
```

Recommended approach:

- Store only `broker_id` in order header tables.
- Do not duplicate broker name, PAN, GSTIN, or address fields.
- If the Sales/Purchase database has access to `brokers`, add a foreign key.
- If broker is owned by a separate common service/database, keep `broker_id` as UUID and validate through application/service logic.

Possible database column:

```sql
broker_id uuid NULL
```

If FK is possible:

```sql
CONSTRAINT fk_sales_order_headers_broker
    FOREIGN KEY (broker_id)
    REFERENCES public.brokers (id)
```

## Field Clarifications

### `delivery_address_id`

Use an id only. Do not save delivery address text on the order header.

Hitrix field:

```text
BkDeleAdd
```

Hitrix meaning:

```text
Delivery address code, linked to tblMastDeleAdd
```

Dhanman recommendation:

```text
delivery_address_id uuid NULL
```

This should reference the Dhanman/common delivery address table once multi-delivery-address support is available. If that table does not exist yet, create it in Common rather than storing plain text in Sales/Purchase.

### `delivery_schedule`

Do not create this as a lookup table.

In Hitrix, delivery schedule is not a master/lookup. It is transactional schedule data. In the GST booking flow, each `tblBooking` row is one schedule line and uses:

```text
BkSrNo
BkShDate
BkShBag
BkShWt
```

So in Dhanman, represent schedule through detail rows:

```text
serial_number
delivery_date
bag
weight
```

If the UI needs multiple delivery schedules for the same product, store each schedule as a separate detail row, or add a child table later:

```text
sales_order_schedules
purchase_order_schedules
```

For the first implementation, keeping schedule fields in order details is enough.

### `note`

Free-text remarks/narration from Hitrix:

```text
BkNote
```

Suggested column:

```text
note text NOT NULL DEFAULT ''
```

Use this for user-entered remarks from the screen.

Hitrix also has spare narration fields:

```text
BkTmpChr1
BkTmpChr2
BkTmpChr3
```

Do not add these unless the UI/business needs extra narration lines. If needed, use clear names like:

```text
narration_1
narration_2
```

### `status_id`

This should be a lookup/status id for order lifecycle.

Recommended statuses:

| Id | Name | Meaning |
|---|---|---|
| 1 | Pending | Created but not fulfilled |
| 2 | Partially Completed | Some quantity dispatched/received |
| 3 | Completed | Fully dispatched/received or converted |
| 4 | Cancelled | Cancelled before fulfillment |

Suggested columns:

```text
status_id integer NOT NULL DEFAULT 1
```

Suggested lookup tables:

```text
sales_order_statuses
purchase_order_statuses
```

Separate tables are better than one shared table if Sales and Purchase workflows may differ later.

### `is_rate_excluding_gst`

Boolean.

From Sales Order screen:

```text
Is Rate Excluding GST ?
```

Suggested column:

```text
is_rate_excluding_gst boolean NOT NULL DEFAULT false
```

Meaning:

- `false`: rate includes GST.
- `true`: rate excludes GST and tax should be calculated separately.

### `is_export`

Boolean.

From Sales Order screen:

```text
Print Export ?
```

Suggested column:

```text
is_export boolean NOT NULL DEFAULT false
```

Meaning:

- `false`: domestic order.
- `true`: export order/print format requirement.

This is sales-side only unless purchase has a confirmed export/import use case.

### `print_ticket_name`

Boolean.

From Sales Order screen:

```text
Print Ticket Name ?
```

Suggested column:

```text
print_ticket_name boolean NOT NULL DEFAULT false
```

Meaning:

- `false`: do not print ticket/count name in the order print.
- `true`: include ticket/count name in the print format.

This appears to be a print/display preference, not a business calculation field.

### `side`

`side` appears on the Sales Order screen as a small numeric field.

Current conclusion:

```text
Do not include this in the first implementation.
```

Reason:

- The Hitrix screen shows `Side`.
- The available `tblBooking` SQL does not have a clear column named `Side`.
- It may be stored in one of the temporary fields like `BkTmpNo1`, `BkTmpNo2`, or `BkTmpNo3`.
- Its business meaning is not confirmed from the available docs.

Recommendation:

- Exclude `side` from the first database design.
- Add it later only if the business confirms what it means.
- If required later, add it as nullable:

```text
side integer NULL
```

Do not create a lookup table for `side` unless fixed values and meanings are confirmed.

## Migration-Only Legacy Fields

These fields are not required for normal Dhanman business logic.

Use them only if old Hitrix booking/order records are migrated into Dhanman and users need traceability back to the old system.

Hitrix identifies records using this composite key:

```text
VNo + VType + VYear + VFirm
```

Dhanman should use:

```text
id uuid
company_id uuid
order_number
fin_year
```

The legacy fields should be nullable and should not drive new Dhanman order logic.

### `legacy_v_no`

Stores the original Hitrix voucher/order number.

Hitrix field:

```text
VNo
```

Suggested column:

```text
legacy_v_no integer NULL
```

Use this only for migration/audit/reference. New Dhanman records should use:

```text
order_number
```

### `legacy_v_type`

Stores the original Hitrix voucher type.

Hitrix field:

```text
VType
```

Examples from Hitrix notes:

| VType | Meaning |
|---|---|
| BK | Trade Booking GST |
| BT | Legacy Trade Booking |
| BM | Mill Bill Booking |
| BD | Mill Bill Despatch Booking |
| BP | Booking Party |
| BI | Booking Inward |
| BS | Booking SIT |

Suggested column:

```text
legacy_v_type text NULL
```

Use this for imported Hitrix records and traceability.

### `legacy_v_firm`

Stores the original Hitrix firm/company code.

Hitrix field:

```text
VFirm
```

Suggested column:

```text
legacy_v_firm text NULL
```

In Dhanman, the real company relation should be:

```text
company_id uuid NOT NULL
```

`legacy_v_firm` is only for migration traceability and reconciliation with old Hitrix data.

### Optional Legacy Columns

Add these only if Hitrix data migration is in scope:

```text
legacy_v_no integer NULL
legacy_v_type text NULL
legacy_v_firm text NULL
legacy_v_year integer NULL
```

Recommended unique migration reference index:

```sql
CREATE UNIQUE INDEX ux_sales_order_headers_legacy_reference
ON public.sales_order_headers (legacy_v_no, legacy_v_type, legacy_v_year, legacy_v_firm)
WHERE legacy_v_no IS NOT NULL;
```

For purchase:

```sql
CREATE UNIQUE INDEX ux_purchase_order_headers_legacy_reference
ON public.purchase_order_headers (legacy_v_no, legacy_v_type, legacy_v_year, legacy_v_firm)
WHERE legacy_v_no IS NOT NULL;
```

## Status Recommendation

Add order status support. Suggested statuses:

| Id | Name |
|---|---|
| 1 | Pending |
| 2 | Partially Completed |
| 3 | Completed |
| 4 | Cancelled |

Use either a shared `order_statuses` table or separate tables:

```text
sales_order_statuses
purchase_order_statuses
```

Separate status tables are safer if Sales and Purchase workflows may diverge later.

## Sales Backend Implementation Plan

1. Add domain entities:
   - `SalesOrderHeader`
   - `SalesOrderDetail`

2. Add EF configurations:
   - `SalesOrderHeaderConfiguration`
   - `SalesOrderDetailConfiguration`

3. Add `DbSet`s in Sales `ApplicationDbContext`.

4. Add repositories:
   - `ISalesOrderHeaderRepository`
   - `ISalesOrderDetailRepository`
   - `SalesOrderHeaderRepository`
   - `SalesOrderDetailRepository`

5. Add contracts:
   - `CreateSalesOrderRequest`
   - `UpdateSalesOrderRequest`
   - `SalesOrderResponse`
   - `SalesOrderListResponse`
   - `SalesOrderLineRequest`
   - `SalesOrderLineResponse`

6. Add commands:
   - `CreateSalesOrderCommand`
   - `UpdateSalesOrderCommand`
   - `DeleteSalesOrderCommand`

7. Add queries:
   - `GetAllSalesOrdersQuery`
   - `GetSalesOrderByIdQuery`
   - `GetSalesOrdersByCustomerIdQuery`
   - `GetSalesOrderDetailsByHeaderIdQuery`
   - `GetPendingSalesOrdersQuery`

8. Add controller:
   - `SalesOrdersController`

9. Add API routes in Sales `ApiRoutes.cs`.

10. Add EF migration for:
    - `sales_order_headers`
    - `sales_order_details`

## Purchase Backend Implementation Plan

1. Add domain entities:
   - `PurchaseOrderHeader`
   - `PurchaseOrderDetail`

2. Add EF configurations:
   - `PurchaseOrderHeaderConfiguration`
   - `PurchaseOrderDetailConfiguration`

3. Add `DbSet`s in Purchase `ApplicationDbContext`.

4. Add repositories:
   - `IPurchaseOrderHeaderRepository`
   - `IPurchaseOrderDetailRepository`
   - `PurchaseOrderHeaderRepository`
   - `PurchaseOrderDetailRepository`

5. Add contracts:
   - `CreatePurchaseOrderRequest`
   - `UpdatePurchaseOrderRequest`
   - `PurchaseOrderResponse`
   - `PurchaseOrderListResponse`
   - `PurchaseOrderLineRequest`
   - `PurchaseOrderLineResponse`

6. Add commands:
   - `CreatePurchaseOrderCommand`
   - `UpdatePurchaseOrderCommand`
   - `DeletePurchaseOrderCommand`

7. Add queries:
   - `GetAllPurchaseOrdersQuery`
   - `GetPurchaseOrderByIdQuery`
   - `GetPurchaseOrdersByVendorIdQuery`
   - `GetPurchaseOrderDetailsByHeaderIdQuery`
   - `GetPendingPurchaseOrdersQuery`

8. Add controller:
   - `PurchaseOrdersController`

9. Add API routes in Purchase `ApiRoutes.cs`.

10. Add EF migration for:
    - `purchase_order_headers`
    - `purchase_order_details`

## Suggested Sales API Routes

```csharp
public static class SalesOrders
{
    public const string GetAllSalesOrders =
        apiVersion + "companies/{companyId:guid}/sales-orders/finYear/{finYearId:int}";

    public const string CreateSalesOrder =
        apiVersion + "sales-orders";

    public const string UpdateSalesOrder =
        apiVersion + "sales-orders";

    public const string GetSalesOrderById =
        apiVersion + "sales-orders/{id:guid}";

    public const string DeleteSalesOrder =
        apiVersion + "sales-orders/{id:guid}";

    public const string GetSalesOrdersByCustomerId =
        apiVersion + "customers/{customerId:guid}/sales-orders/finYear/{finYearId:int}";

    public const string GetSalesOrderDetails =
        apiVersion + "sales-orders/{salesOrderId:guid}/details";
}
```

## Suggested Purchase API Routes

```csharp
public static class PurchaseOrders
{
    public const string GetAllPurchaseOrders =
        apiVersion + "companies/{companyId:guid}/purchase-orders/finYear/{finYearId:int}";

    public const string CreatePurchaseOrder =
        apiVersion + "purchase-orders";

    public const string UpdatePurchaseOrder =
        apiVersion + "purchase-orders";

    public const string GetPurchaseOrderById =
        apiVersion + "purchase-orders/{id:guid}";

    public const string DeletePurchaseOrder =
        apiVersion + "purchase-orders/{id:guid}";

    public const string GetPurchaseOrdersByVendorId =
        apiVersion + "vendors/{vendorId:guid}/purchase-orders/finYear/{finYearId:int}";

    public const string GetPurchaseOrderDetails =
        apiVersion + "purchase-orders/{purchaseOrderId:guid}/details";
}
```

## Suggested Controllers

Create these controllers:

```text
dhanman-sales/src/Dhanman.Sales.Api/Controllers/SalesOrdersController.cs
dhanman-purchase/src/Dhanman.Purchase.Api/Controllers/PurchaseOrdersController.cs
```

Each controller should support:

- Create
- Update
- Delete
- Get by id
- Get all by company and financial year
- Get by customer/vendor
- Get details by header id

## Business Rules

1. Orders are pre-transaction records.
2. Orders should not post to ledger.
3. Orders should not affect customer/vendor outstanding directly.
4. Ordered quantity should be tracked separately from dispatched/received quantity.
5. Duplicate prevention should check same party/vendor, date, product, and broker.
6. Completed or cancelled orders should not be editable unless business explicitly allows reopening.
7. Sales Order should later convert to Invoice.
8. Purchase Order should later convert to Bill.

## Conversion APIs For Later Phase

After CRUD is stable, add:

```text
sales-orders/{id}/convert-to-invoice
purchase-orders/{id}/convert-to-bill
```

Conversion should carry forward:

- Order number
- Order date
- Customer/vendor
- Broker
- Product lines
- Bags
- Weight
- Rate
- Amount
- Delivery information
- Reference/contract number

## Implementation Sequence

Recommended order:

1. Create database tables and EF entities.
2. Add repositories and DbContext registration.
3. Add request/response contracts.
4. Add create/get/update/delete commands and queries.
5. Add controllers and API routes.
6. Add migrations.
7. Add validations.
8. Add order status transitions.
9. Add conversion to invoice/bill.
10. Add reports and pending order views.

## Important Recommendation

Do not implement Sales Order inside `invoice_headers`.

Do not implement Purchase Order inside `bill_headers`.

The existing `invoice_headers` and `bill_headers` tables are accounting transaction tables. Hitrix booking/order records are pre-transaction records, so separate order tables are the correct design.
