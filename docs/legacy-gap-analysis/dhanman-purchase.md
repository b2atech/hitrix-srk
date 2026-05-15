# dhanman-purchase Legacy Comparison

## Project Summary

`dhanman-purchase` is the purchase backend service. It contains vendor, bill, transaction, gate pass, warehouse, and related admin surfaces for procurement-side operations.

## Current Dhanman Scope

- API controllers in `src/Dhanman.Purchase.Api/Controllers`
- domain entities in `src/Dhanman.Purchase.Domain`
- persistence and infrastructure under `src/Dhanman.Purchase.Persistence` and `src/Infrastructure`

## Legacy VB Scope Placeholder

Legacy VB scope is not available in this workspace.

Expected legacy evidence to capture later:

- vendor master forms
- purchase bill entry and posting
- purchase payment or money movement flows
- gate pass or goods inward support
- purchase reporting and approvals

## Feature Comparison Matrix

| Area | Feature / Object | Legacy VB | Dhanman | Gap | Priority | Evidence | Notes |
| --- | --- | --- | --- | --- | --- | --- | --- |
| Purchase | Vendors and vendor notes | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `VendorsController.cs`, `VendorNoteController.cs` | Vendor-facing operational surface is present. |
| Purchase | Bills and ledger support | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `BillsController.cs`, `LedgerController.cs`, `TransactionController.cs` | Core procurement transaction flow is visible. |
| Purchase | Money and payment operations | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `Medium` | `MoneyController.cs`, `TransactionController.cs` | Useful parity area for settlement workflows. |
| Purchase | Gate pass and warehouse references | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `Medium` | `GatePassController.cs`, `WareHouseController.cs` | Cross-service warehouse consistency matters here too. |
| Purchase | Cache and admin support | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `Low` | `AdminCacheController.cs`, `PermissionsController.cs`, `MasterController.cs` | Support functions, not direct business parity drivers. |

## Tables And Entities Matrix

| Area | Feature / Object | Legacy VB | Dhanman | Gap | Priority | Evidence | Notes |
| --- | --- | --- | --- | --- | --- | --- | --- |
| Masters | Vendor and warehouse master tables | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `vendors.sql`, `vendor_categories.sql`, `vendor_contacts.sql`, `vendor_bank_accounts.sql`, `warehouses.sql` | Purchase has concrete vendor and warehouse master tables. |
| Transactions | Bill header/detail/payment tables | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `bill_headers.sql`, `bill_details.sql`, `bill_payment_headers.sql`, `bill_payment_details.sql`, `payment_orders.sql` | Main procurement transaction tables are explicit in persistence scripts. |
| Workflow | Bill approval and status tables | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `bill_statuses.sql`, `bill_workflow.sql`, `bill_approval_logs.sql`, `bill_approval_issue_logs.sql`, `bill_account_approval_levels.sql` | Strong evidence of status and approval workflow support. |
| Operations | Gate pass and schedule execution tables | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `Medium` | `gate_pass_headers.sql`, `gate_pass_details.sql`, `gate_pass_statuses.sql`, `recurring_bill_schedules.sql`, `schedule_execution_logs.sql` | Includes physical operations and recurring/batch billing support. |
| Projection | Advance and settlement tables | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `Medium` | `vendor_advances.sql`, `vendor_advance_refunds.sql`, `vendor_advance_settlements.sql`, `payment_statuses.sql` | Useful for comparing legacy vendor advance handling. |

## All Tables

All SQL table/object scripts currently visible under `dhanman-purchase/src/Dhanman.Purchase.Persistence/DBObjects/Table`:

- `__EFMigrationsHistory.sql`
- `addresses.sql`
- `bank_accounts.sql`
- `banks.sql`
- `bill_account_approval_levels.sql`
- `bill_approval_issue_logs.sql`
- `bill_approval_logs.sql`
- `bill_approval_user_account.sql`
- `bill_approval_user_company.sql`
- `bill_details.sql`
- `bill_header_ids.sql`
- `bill_headers.sql`
- `bill_payment_details.sql`
- `bill_payment_header_ids.sql`
- `bill_payment_headers.sql`
- `bill_status_company_configs.sql`
- `bill_statuses.sql`
- `bill_types.sql`
- `bill_workflow.sql`
- `cities.sql`
- `companies.sql`
- `contacts.sql`
- `countries.sql`
- `draft_bill_details.sql`
- `draft_bill_header_ids.sql`
- `draft_bill_headers.sql`
- `gate_pass_details.sql`
- `gate_pass_headers.sql`
- `gate_pass_statuses.sql`
- `organizations.sql`
- `payment_orders.sql`
- `payment_statuses.sql`
- `procedure_logs.sql`
- `recurrence_bill_schedule_details.sql`
- `recurring_bill_schedules.sql`
- `roles.sql`
- `schedule_execution_logs.sql`
- `schema_versions.sql`
- `states.sql`
- `tds_statuses.sql`
- `temp_bill_next_status.sql`
- `upis.sql`
- `user_roles.sql`
- `users.sql`
- `v_organization_id.sql`
- `vendor_advance_ids.sql`
- `vendor_advance_refunds.sql`
- `vendor_advance_settlement_ids.sql`
- `vendor_advance_settlements.sql`
- `vendor_advances.sql`
- `vendor_bank_accounts.sql`
- `vendor_categories.sql`
- `vendor_contacts.sql`
- `vendor_default_accounts.sql`
- `vendor_note_details.sql`
- `vendor_note_headers.sql`
- `vendor_note_statuses.sql`
- `vendor_note_work_flows.sql`
- `vendor_upis.sql`
- `vendors.sql`
- `warehouses.sql`

## APIs / Controllers / Services Matrix

| Area | Feature / Object | Legacy VB | Dhanman | Gap | Priority | Evidence | Notes |
| --- | --- | --- | --- | --- | --- | --- | --- |
| APIs | Vendor management APIs | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `VendorsController.cs`, `VendorNoteController.cs` | Vendor administration is directly exposed. |
| APIs | Purchase bill and ledger APIs | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `BillsController.cs`, `LedgerController.cs` | Key parity area for procurement accounting. |
| APIs | Money, transaction, and gate pass APIs | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `Medium` | `MoneyController.cs`, `TransactionController.cs`, `GatePassController.cs` | Operational purchase APIs are present. |
| APIs | Warehouse and org support APIs | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `Medium` | `WareHouseController.cs`, `OrganizationController.cs`, `MasterController.cs` | Supporting configuration surfaces. |

## Screens / Forms / Pages Matrix

| Area | Feature / Object | Legacy VB | Dhanman | Gap | Priority | Evidence | Notes |
| --- | --- | --- | --- | --- | --- | --- | --- |
| Business workflow | Vendor setup and maintenance | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `VendorsController.cs`, `AddressController.cs` | Compare later with VB vendor onboarding forms. |
| Business workflow | Bill entry and ledger impact | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `BillsController.cs`, `LedgerController.cs` | Likely the main legacy parity workflow. |
| Business workflow | Goods movement support | To be verified from VB form/report/table | `Partial` | Cannot assess until legacy evidence is added | `Medium` | `GatePassController.cs`, `WareHouseController.cs`, `dhanman-common/inventory-module-gap-analysis.md` | Compare with legacy inward/gate pass behavior and inventory integration. |

## Reports / Export / Integration Matrix

| Area | Feature / Object | Legacy VB | Dhanman | Gap | Priority | Evidence | Notes |
| --- | --- | --- | --- | --- | --- | --- | --- |
| Reports | Vendor and bill reporting | To be verified from VB form/report/table | `Partial` | Cannot assess until legacy evidence is added | `High` | `BillsController.cs`, `VendorsController.cs` | Specific report endpoints are not obvious from controller names alone. |
| Reports | Ledger and transaction reporting | To be verified from VB form/report/table | `Partial` | Cannot assess until legacy evidence is added | `Medium` | `LedgerController.cs`, `TransactionController.cs` | Add exact VB reports later. |
| Integration | Inventory warehouse and receipt linkage | To be verified from VB form/report/table | `Partial` | Cannot assess until legacy evidence is added | `High` | `WareHouseController.cs`, `dhanman-common/inventory-module-gap-analysis.md`, `inventory-module-remaining-gaps.md` | Important if legacy purchase posting directly created stock receipts. |

## Gaps And Recommendations

- Capture the full purchase bill lifecycle from VB before comparing individual tables; approval and posting rules matter more than surface names.
- Mark whether VB purchase receipts updated stock directly or through a separate inventory workflow.
- Add vendor payment and money movement parity rows once the legacy evidence is available.
- Keep warehouse and gate pass notes tied to the inventory comparison to avoid conflicting conclusions.

## Evidence From Codebase

- `dhanman-purchase/src/Dhanman.Purchase.Api/Controllers`
- `dhanman-purchase/src/Dhanman.Purchase.Domain`
- `dhanman-purchase/src/Dhanman.Purchase.Persistence/DBObjects/Table`
- `dhanman-common/inventory-module-gap-analysis.md`
- `dhanman-common/inventory-module-remaining-gaps.md`

## Open Legacy Questions

- Which VB forms handled vendors, bills, and purchase ledger updates?
- Did the legacy system have separate inward, receipt, or gate pass workflows?
- Which purchase reports were used daily by operations and finance teams?
- Was the warehouse master maintained inside purchase or synchronized from a shared module?
