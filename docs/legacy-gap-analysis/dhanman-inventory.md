# dhanman-inventory Legacy Comparison

## Project Summary

`dhanman-inventory` is the primary inventory backend service. It contains APIs, application features, domain entities, persistence objects, and database scripts for inventory masters, transactions, and reporting surfaces.

## Current Dhanman Scope

- API controllers in `src/Dhanman.Inventory.Api/Controllers`
- application features in `src/Core/Dhanman.Inventory.Application/Features`
- domain entities in `src/Core/Dhanman.Inventory.Domain/Entities`
- persistence scripts in `src/Dhanman.Inventory.Persistence/DBObjects/Table`

## Legacy VB Scope Placeholder

Legacy VB scope is not available in this workspace.

Expected legacy evidence to capture later:

- inventory master tables
- transaction entry forms
- posting logic
- stock movement reports
- approval or audit routines

## Feature Comparison Matrix

| Area | Feature / Object | Legacy VB | Dhanman | Gap | Priority | Evidence | Notes |
| --- | --- | --- | --- | --- | --- | --- | --- |
| Masters | Products, categories, units, tax categories | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `ProductController.cs`, `CategoriesController.cs`, `UnitController.cs`, `ProductTaxCategoryController.cs` | Core master APIs exist. |
| Masters | Warehouses and warehouse locations | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `WarehousesController.cs`, `WarehouseLocationsController.cs`, `Entities/Warehouses`, `Entities/WarehouseLocations` | Warehouse ownership is clearly implemented in this service. |
| Transactions | Goods receipt and goods issue | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `GoodsReceiptController.cs`, `GoodsIssueController.cs`, `Entities/GoodsReceiptHeaders`, `Entities/GoodsIssueHeaders` | Core transaction flows are present. |
| Transactions | Material requests | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `MaterialController.cs`, `Entities/MaterialRequestHeaders`, `Entities/MaterialRequestWorkflows` | Request and workflow entities exist. |
| Transactions | Stock transfer, stock adjustment, stock count | To be verified from VB form/report/table | `Partial` | Cannot assess until legacy evidence is added | `High` | `Entities/StockTransferHeader`, `Entities/StockAdjustmentHeader`, `Entities/StockCountHeaders` | Data model exists; parity should be checked per end-to-end workflow. |
| Reporting | Inventory balance and inventory ledger | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `InventoryBalancesController.cs`, `InventoryLedgerController.cs`, `Entities/InventoryBalances`, `Entities/StockLedgers` | Reporting and ledger surfaces exist. |
| Integration | Purchase/sales sync events | To be verified from VB form/report/table | `Partial` | Cannot assess until legacy evidence is added | `Medium` | `dhanman-shared-contracts/src/Dhanman.Shared.Contracts/Events/PurchaseReceiptPostedEvent.cs`, `SalesInvoicePostedEvent.cs`, `dhanman-inventory/inventory-module-gap-analysis.md` | Important for cross-service parity tracking. |

## Tables And Entities Matrix

| Area | Feature / Object | Legacy VB | Dhanman | Gap | Priority | Evidence | Notes |
| --- | --- | --- | --- | --- | --- | --- | --- |
| Masters | Product and product setup tables | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `products.sql`, `product_group.sql`, `product_reorder_settings.sql`, `Entities/Products`, `Entities/ProductGroups`, `Entities/ProductReorderSettings` | Strong evidence of product master design. |
| Masters | Product price guardrails (`MinSellingPrice`, `MaxPurchasePrice`) | To be verified from VB form/report/table | `Missing` | Product-level floor/ceiling rate controls are not visible in the current inventory model | `High` | `products.sql`, `ProductConfiguration.cs` | Recommended improvement to prevent accidental low sales rates and high purchase rates by end users. |
| Masters | Warehouse model | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `Entities/Warehouses`, `Entities/WarehouseLocations`, `product_warehouse_quantities.sql` | Includes warehouse quantities and locations. |
| Transactions | Goods receipt and goods issue entities | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `Entities/GoodsReceiptHeaders`, `Entities/GoodsReceiptDetails`, `Entities/GoodsIssueHeaders`, `Entities/GoodsIssueDetails` | Core document entities are present. |
| Controls | Stock movement and reservation tables | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `stock_ledgers.sql`, `stock_reservations.sql`, `stock_movement_types.sql`, `Entities/StockLedgers`, `Entities/StockReservations` | Good parity anchor for later VB table mapping. |
| Controls | Transfer, count, and adjustment tables | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `stock_transfer_headers.sql`, `stock_adjustment_headers.sql`, `Entities/StockTransferHeader`, `Entities/StockAdjustmentHeader`, `Entities/StockCountHeaders` | Data structures are visible even where workflows may still be partial. |
| Security | Users and organizations | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `Low` | `users.sql`, `organizations.sql`, `Entities/Users`, `Entities/Organizations` | Supporting entities for workflow ownership and permissions. |

## All Tables

All SQL table/object scripts currently visible under `dhanman-inventory/src/Dhanman.Inventory.Persistence/DBObjects/Table`:

- `__EFMigrationsHistory.sql`
- `categories.sql`
- `companies.sql`
- `company_categories.sql`
- `company_products.sql`
- `customer_products.sql`
- `default_apartment_products.sql`
- `gate_pass_logs.sql`
- `order_details.sql`
- `orders.sql`
- `organizations.sql`
- `product_group.sql`
- `product_reorder_settings.sql`
- `product_tax_categories.sql`
- `product_tax_category_details.sql`
- `product_warehouse_quantities.sql`
- `products.sql`
- `roles.sql`
- `schema_versions.sql`
- `stock_adjustment_details.sql`
- `stock_adjustment_headers.sql`
- `stock_adjustment_statuses.sql`
- `stock_ledgers.sql`
- `stock_movement_types.sql`
- `stock_reservation_statuses.sql`
- `stock_reservations.sql`
- `stock_transfer_details.sql`
- `stock_transfer_headers.sql`
- `stock_transfer_statuses.sql`
- `units.sql`
- `users.sql`
- `vendor_products.sql`

## APIs / Controllers / Services Matrix

| Area | Feature / Object | Legacy VB | Dhanman | Gap | Priority | Evidence | Notes |
| --- | --- | --- | --- | --- | --- | --- | --- |
| APIs | Product and reorder APIs | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `ProductController.cs`, `ProductReorderSettingsController.cs` | Covers products and reorder settings. |
| APIs | Receipt type and unit conversion APIs | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `Medium` | `ReceiptTypesController.cs`, `UnitConversionController.cs` | Useful for detailed parity review. |
| APIs | Warehouse APIs | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `WarehousesController.cs`, `WarehouseLocationsController.cs` | Important for sync and ownership analysis. |
| APIs | Transaction APIs | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `GoodsReceiptController.cs`, `GoodsIssueController.cs`, `GatePassController.cs`, `MaterialController.cs` | Main inventory transaction entry points. |
| APIs | Reporting APIs | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `Medium` | `InventoryBalancesController.cs`, `InventoryLedgerController.cs`, `OrderController.cs` | Reporting and order-support surfaces are visible. |
| Services | Messaging and sync support | To be verified from VB form/report/table | `Partial` | Cannot assess until legacy evidence is added | `Medium` | `src/Dhanman.Inventory.Api/Extensions`, `inventory-module-gap-analysis.md` | Warehouse sync and event surfaces exist, but full runtime parity needs later review. |

## Screens / Forms / Pages Matrix

| Area | Feature / Object | Legacy VB | Dhanman | Gap | Priority | Evidence | Notes |
| --- | --- | --- | --- | --- | --- | --- | --- |
| Backend workflow | Master maintenance flow | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `ProductController.cs`, `CategoriesController.cs`, `UnitController.cs`, `WarehousesController.cs` | API support exists for the frontend or external clients. |
| Backend workflow | Stock posting document flow | To be verified from VB form/report/table | `Partial` | Cannot assess until legacy evidence is added | `High` | `GoodsReceiptController.cs`, `GoodsIssueController.cs`, `Entities/StockLedgers` | Use later to compare with VB post/approve routines. |
| Backend workflow | Stock control workflow | To be verified from VB form/report/table | `Partial` | Cannot assess until legacy evidence is added | `High` | `Entities/StockTransferHeader`, `Entities/StockAdjustmentHeader`, `Entities/StockCountHeaders` | Workflow parity should be checked against UI and posting completeness. |

## Reports / Export / Integration Matrix

| Area | Feature / Object | Legacy VB | Dhanman | Gap | Priority | Evidence | Notes |
| --- | --- | --- | --- | --- | --- | --- | --- |
| Reports | Inventory balance | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `InventoryBalancesController.cs`, `Entities/InventoryBalances` | Key parity candidate for stock reporting. |
| Reports | Inventory ledger / stock movement | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `InventoryLedgerController.cs`, `stock_ledgers.sql` | Important for stock audit comparison. |
| Integration | Shared warehouse events | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `Medium` | `dhanman-shared-contracts/src/Dhanman.Shared.Contracts/Events/WarehouseCreatedEvent.cs`, `WarehouseUpdatedEvent.cs`, `WarehouseDeactivatedEvent.cs` | Relevant if VB used shared warehouse masters. |
| Integration | Purchase and sales posting events | To be verified from VB form/report/table | `Partial` | Cannot assess until legacy evidence is added | `Medium` | `PurchaseReceiptPostedEvent.cs`, `SalesInvoicePostedEvent.cs`, `inventory-module-gap-analysis.md` | Helps compare legacy cross-module posting behavior. |

## Gaps And Recommendations

- Map legacy VB stock tables first against `stock_ledgers`, `stock_reservations`, and quantity tables because those usually drive the hardest migration questions.
- Add product-level price guardrails with `MinSellingPrice` and `MaxPurchasePrice`, then enforce them in sales and purchase entry flows to reduce operator mistakes.
- Track whether VB used post/approve/cancel states for receipt, issue, adjustment, transfer, and count separately.
- Note every legacy report that depended on stock ledger or inventory balance logic; those often reveal hidden business rules.
- If VB had direct purchase or sales posting into stock, capture that clearly because Dhanman uses cross-service integration patterns.

## Evidence From Codebase

- `dhanman-inventory/src/Dhanman.Inventory.Api/Controllers`
- `dhanman-inventory/src/Core/Dhanman.Inventory.Domain/Entities`
- `dhanman-inventory/src/Dhanman.Inventory.Persistence/DBObjects/Table`
- `dhanman-inventory/inventory-module-gap-analysis.md`
- `dhanman-inventory/docs/warehouse-address-refactor-gpt-review.md`

## Open Legacy Questions

- Which VB tables represented stock ledger, on-hand balance, and reservation logic?
- Did VB enforce any item-wise minimum sales rate or maximum purchase rate, or was it handled only by user training and approvals?
- Did VB store warehouse locations separately from warehouses?
- Which VB transactions were draft-only versus posted-to-stock?
- Did VB have stock transfer, stock count, and adjustment workflows with approvals?
