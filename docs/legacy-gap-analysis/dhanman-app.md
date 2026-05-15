# dhanman-app Legacy Comparison

## Project Summary

`dhanman-app` is the primary frontend application. It is the shared UI shell for multiple Dhanman business modules including inventory, sales, purchase, payroll, payment, company, finance, and community flows.

## Current Dhanman Scope

- module pages under `src/pages`
- API wrappers under `src/api`
- module route definitions under `src/routes/modules`
- menu and permission wiring under `src/menu-items`

## Legacy VB Scope Placeholder

Legacy VB scope is not yet available in this workspace.

Expected legacy evidence to capture later:

- legacy desktop menu path by module
- forms for masters, transactions, utilities, and reports
- report/export screens
- lookup dialogs and approval dialogs

## Feature Comparison Matrix

| Area | Feature / Object | Legacy VB | Dhanman | Gap | Priority | Evidence | Notes |
| --- | --- | --- | --- | --- | --- | --- | --- |
| Inventory UI | Product list/create/view/edit | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `dhanman-app/src/pages/inventory/product`, `dhanman-app/src/routes/modules/inventory-routes.tsx` | Product routes and pages are wired. |
| Frontend shell | Multi-module page structure | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `dhanman-app/src/pages/company`, `dhanman-app/src/pages/community`, `dhanman-app/src/pages/payment`, `dhanman-app/src/pages/payroll`, `dhanman-app/src/pages/purchase`, `dhanman-app/src/pages/sales` | Frontend is organized around multiple business modules, not inventory only. |
| Frontend shell | Shared dashboard and navigation structure | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `Medium` | `dhanman-app/src/pages/dashboard`, `dhanman-app/src/menu-items` | Useful for comparing legacy menu layout and module entry points. |
| Inventory UI | Categories and units management | To be verified from VB form/report/table | `Partial` | Cannot assess until legacy evidence is added | `High` | `dhanman-app/src/pages/inventory/category`, `dhanman-app/src/pages/inventory/unit`, `dhanman-app/src/api/inventory/categoryUnitMappers.ts` | UI exists, but earlier analysis marks some screens as mock/partial. |
| Inventory UI | Company/vendor/customer warehouses | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `dhanman-app/src/pages/inventory/warehouse`, `dhanman-app/src/api/inventory/warehouse.ts` | Separate list/create/edit/view flows exist. |
| Inventory UI | Warehouse locations | To be verified from VB form/report/table | `Partial` | Cannot assess until legacy evidence is added | `Medium` | `dhanman-app/src/pages/inventory/warehouseLocation`, `dhanman-app/src/api/inventory/warehouseLocation.ts` | Present in routes and pages, completeness should be checked against backend. |
| Inventory UI | Goods receipt and goods issue | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `dhanman-app/src/pages/inventory/goodsReceipt`, `dhanman-app/src/pages/inventory/goodsIssue` | Core transaction pages are present. |
| Inventory UI | Stock transfer, adjustment, count, gate pass | To be verified from VB form/report/table | `Partial` | Cannot assess until legacy evidence is added | `High` | `dhanman-app/src/pages/inventory/stockTransfer`, `dhanman-app/src/pages/inventory/stockAdjustment`, `dhanman-app/src/pages/inventory/stockCount`, `dhanman-app/src/pages/inventory/gatePass` | List screens exist, several create/edit routes still point to `ComingSoon`. |
| Inventory UI | Reports: inventory balance, stock ledger, inventory ledger | To be verified from VB form/report/table | `Partial` | Cannot assess until legacy evidence is added | `High` | `dhanman-app/src/pages/inventory/inventoryBalance`, `dhanman-app/src/pages/inventory/stockLedger`, `dhanman-app/src/pages/inventory/inventoryLedger` | Reporting entry points exist; some screens were previously noted as mock/partial. |

## Tables And Entities Matrix

| Area | Feature / Object | Legacy VB | Dhanman | Gap | Priority | Evidence | Notes |
| --- | --- | --- | --- | --- | --- | --- | --- |
| Frontend data model | Inventory API mappers | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `Medium` | `dhanman-app/src/api/inventory/categoryUnitMappers.ts`, `dhanman-app/src/api/inventory/warehouseLocationMappers.ts` | Mapper layer exists for some modules. |
| Frontend data model | Multi-module API folder structure | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `Medium` | `dhanman-app/src/api/common`, `dhanman-app/src/api/community`, `dhanman-app/src/api/finance`, `dhanman-app/src/api/payroll`, `dhanman-app/src/api/purchase`, `dhanman-app/src/api/sales` | Frontend API surface spans several business domains. |
| Frontend data model | Contract tests for API mappers | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `Low` | `dhanman-app/src/api/inventory/categoryUnitMappers.contract.test.ts`, `dhanman-app/src/api/inventory/warehouseLocationMappers.contract.test.ts` | Evidence of mapper verification exists. |
| Frontend data model | Shared page-level types by inventory module | To be verified from VB form/report/table | `Unknown` | Cannot assess until legacy evidence is added | `Medium` | `dhanman-app/src/pages/inventory` | Type consistency should be checked later by screen. |

## APIs / Controllers / Services Matrix

| Area | Feature / Object | Legacy VB | Dhanman | Gap | Priority | Evidence | Notes |
| --- | --- | --- | --- | --- | --- | --- | --- |
| Frontend integration | Inventory product API wrapper | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `dhanman-app/src/api/inventory/product.ts` | API wrapper exists for product flows. |
| Frontend integration | Shared and finance API wrappers | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `Medium` | `dhanman-app/src/api/common`, `dhanman-app/src/api/finance`, `dhanman-app/src/api/razorpay` | Current frontend talks to common, finance, and payment-oriented APIs. |
| Frontend integration | Sales, purchase, payroll, and community APIs | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `Medium` | `dhanman-app/src/api/sales`, `dhanman-app/src/api/purchase`, `dhanman-app/src/api/payroll`, `dhanman-app/src/api/community` | Use this to map legacy module-level frontend integrations later. |
| Frontend integration | Goods receipt and goods issue APIs | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `dhanman-app/src/api/inventory/goodsReceipts.ts`, `dhanman-app/src/api/inventory/goodsIssue.ts` | Core transaction APIs are wired. |
| Frontend integration | Material request, reorder, warehouse APIs | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `dhanman-app/src/api/inventory/materialRequest.ts`, `dhanman-app/src/api/inventory/reorderSettings.ts`, `dhanman-app/src/api/inventory/warehouse.ts` | Broad inventory API coverage is visible. |
| Frontend integration | Gate pass API | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `Medium` | `dhanman-app/src/api/inventory/gatePass.ts` | Backend wrapper exists, UI completeness still needs parity review. |
| Frontend integration | Stock transfer and stock count APIs | To be verified from VB form/report/table | `Missing` | Cannot assess until legacy evidence is added | `High` | `dhanman-app/src/api/inventory` | No dedicated `stockTransfer.ts` or `stockCount.ts` wrapper was found. |

## Screens / Forms / Pages Matrix

| Area | Feature / Object | Legacy VB | Dhanman | Gap | Priority | Evidence | Notes |
| --- | --- | --- | --- | --- | --- | --- | --- |
| Routes | Inventory master routes | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `dhanman-app/src/routes/modules/inventory-routes.tsx` | Products, categories, units, receipt types, warehouses, and locations are routed. |
| Routes | Multi-module route organization | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `dhanman-app/src/routes/modules`, `dhanman-app/src/pages/company`, `dhanman-app/src/pages/purchase`, `dhanman-app/src/pages/sales`, `dhanman-app/src/pages/payroll`, `dhanman-app/src/pages/community` | Frontend routes are partitioned by service/module. |
| Routes | Inventory transaction routes | To be verified from VB form/report/table | `Partial` | Cannot assess until legacy evidence is added | `High` | `dhanman-app/src/routes/modules/inventory-routes.tsx` | Goods receipt and issue are wired; several transaction create/edit routes are `ComingSoon`. |
| Navigation | Inventory menu grouping | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `Medium` | `dhanman-app/src/menu-items/inventory.tsx` | Masters, Transactions, and Reports groupings are clearly defined. |
| Navigation | Shared module menu structure | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `Medium` | `dhanman-app/src/menu-items` | Use later to compare the Dhanman navigation tree with legacy desktop menus. |
| Permissions | Inventory permission-gated navigation | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `Medium` | `dhanman-app/src/menu-items/inventory.tsx` | Menu items use `DhanmanPermissions`. |

## Reports / Export / Integration Matrix

| Area | Feature / Object | Legacy VB | Dhanman | Gap | Priority | Evidence | Notes |
| --- | --- | --- | --- | --- | --- | --- | --- |
| Reports UI | Inventory balance report screen | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `dhanman-app/src/pages/inventory/inventoryBalance/list.tsx` | Exists as a report entry screen. |
| Reports UI | Module-specific reporting surfaces outside inventory | To be verified from VB form/report/table | `Unknown` | Cannot assess until legacy evidence is added | `Medium` | `dhanman-app/src/pages`, `dhanman-app/src/api` | Needs later service-by-service page review beyond inventory. |
| Reports UI | Stock ledger report screen | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `dhanman-app/src/pages/inventory/stockLedger/list.tsx` | Exists as a report entry screen. |
| Reports UI | Inventory ledger report screen | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `Medium` | `dhanman-app/src/pages/inventory/inventoryLedger/inventoryLedgers.tsx` | Exists but should be validated against business parity later. |
| Integration | API-backed vs mock-backed screens | To be verified from VB form/report/table | `Partial` | Cannot assess until legacy evidence is added | `High` | `dhanman-app/src/pages/inventory`, `dhanman-inventory/inventory-module-gap-analysis.md` | Existing analysis notes a mix of API-backed and mock screens. |

## Gaps And Recommendations

- Use this file as the frontend entry point for all legacy UI comparison, then drill into the backend service files for deeper business rules.
- Map the VB menu tree to `src/pages`, `src/routes/modules`, and `src/menu-items` before comparing individual screens.
- Track `ComingSoon` routes separately because they are likely to become parity gaps.
- Add VB report names beside the three current Dhanman report screens first; this will quickly show reporting gaps.
- If the legacy UI has batch operations, keyboard shortcuts, or print flows, record them in `Notes` because those are easy to miss in a page-only review.

## Evidence From Codebase

- `dhanman-app/src/pages/inventory`
- `dhanman-app/src/pages`
- `dhanman-app/src/api`
- `dhanman-app/src/api/inventory`
- `dhanman-app/src/routes/modules`
- `dhanman-app/src/routes/modules/inventory-routes.tsx`
- `dhanman-app/src/menu-items`
- `dhanman-app/src/menu-items/inventory.tsx`
- `dhanman-inventory/inventory-module-gap-analysis.md`

## Open Legacy Questions

- Which VB desktop menus map to company, community, inventory, sales, purchase, payroll, and payment areas?
- Which VB menus map to Products, Categories, Units, and Receipt Types?
- Did the VB system separate company, vendor, and customer warehouses the same way?
- Which VB reports existed for stock balance, stock ledger, and inventory ledger?
- Did VB support create/edit flows for stock transfer, stock adjustment, stock count, and gate pass?
