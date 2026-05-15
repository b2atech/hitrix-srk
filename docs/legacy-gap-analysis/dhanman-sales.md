# dhanman-sales Legacy Comparison

## Project Summary

`dhanman-sales` is the sales backend service. It contains customer-facing business entities and APIs for invoicing, payments, customer notes, gate passes, warehouse references, and supporting masters.

## Current Dhanman Scope

- API controllers in `src/Dhanman.Sales.Api/Controllers`
- domain entities in `src/Core/Dhanman.Sales.Domain/Entities`
- persistence and infrastructure under `src/Dhanman.Sales.Persistence` and `src/Infrastructure`

## Legacy VB Scope Placeholder

Legacy VB scope is not available in this workspace.

Expected legacy evidence to capture later:

- customer master forms
- invoice entry and approval screens
- payment receipt screens
- defaulter and penalty processes
- sales reports and gate pass reports

## Feature Comparison Matrix

| Area | Feature / Object | Legacy VB | Dhanman | Gap | Priority | Evidence | Notes |
| --- | --- | --- | --- | --- | --- | --- | --- |
| Sales | Customers and customer notes | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `CustomersController.cs`, `CustomerNoteController.cs`, `Entities/Customers`, `Entities/CustomerNoteHeaders` | Strong customer and note management evidence. |
| Sales | Invoices and grouped invoices | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `InvoicesController.cs`, `Entities/InvoiceHeaders`, `Entities/GroupedInvoices` | Core billing workflow is present. |
| Sales | Payments and ledger support | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `PaymentsController.cs`, `LedgerController.cs`, `Entities/InvoicePaymentHeaders` | Important for collections parity. |
| Sales | Defaulters and penalty workflows | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `Medium` | `DefaultersController.cs`, `Entities/Defaulters`, `Entities/PenaltyConfigs` | Useful for later business-rule comparison. |
| Sales | Gate pass and warehouse references | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `Medium` | `GatePassController.cs`, `WareHouseController.cs`, `Entities/GatePassHeaders`, `Entities/WareHouses` | Cross-service warehouse parity matters here. |
| Sales | Batch and recurring schedule support | To be verified from VB form/report/table | `Partial` | Cannot assess until legacy evidence is added | `Medium` | `BatchController.cs`, `Entities/RecurringSalesSchedules`, `Entities/ScheduleExecutionLogs` | Needs later parity review against legacy operational jobs. |

## Tables And Entities Matrix

| Area | Feature / Object | Legacy VB | Dhanman | Gap | Priority | Evidence | Notes |
| --- | --- | --- | --- | --- | --- | --- | --- |
| Masters | Customers and contacts | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `Entities/Customers`, `Entities/CustomerContacts`, `Entities/CustomerBankAccounts` | Customer master model is broad. |
| Transactions | Invoice and payment tables | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `Entities/InvoiceHeaders`, `Entities/InvoiceDetails`, `Entities/InvoicePaymentHeaders`, `Entities/InvoicePaymentDetails` | Core sales transaction evidence. |
| Control | Statuses, workflows, approvals | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `Entities/InvoiceStatuses`, `Entities/InvoiceWorkflows`, `Entities/InvoiceApprovals`, `Entities/InvoiceApprovalLogs` | Important for matching legacy approval rules. |
| Operations | Gate pass and defaulter tables | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `Medium` | `Entities/GatePassHeaders`, `Entities/GatePassDetails`, `Entities/Defaulters` | Supporting operations beyond invoicing. |
| Projection | Warehouse and organization references | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `Medium` | `Entities/WareHouses`, `Entities/Organizations`, `Entities/Companies` | Useful for cross-service alignment. |
| Physical tables | Sales DB table scripts | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `invoice_headers.sql`, `invoice_details.sql`, `invoice_payment_headers.sql`, `customers.sql`, `defaulter_configurations.sql`, `warehouses.sql` | Concrete SQL scripts exist for billing, customer, defaulter, and warehouse areas. |

## All Tables

All SQL table/object scripts currently visible under `dhanman-sales/src/Dhanman.Sales.Persistence/DBObjects/Table`:

- `__EFMigrationsHistory.sql`
- `addresses.sql`
- `audit_queue.sql`
- `audit_queue_invoice_details.sql`
- `audit_queue_invoice_headers.sql`
- `audit_queue_invoice_payment_details.sql`
- `audit_queue_invoice_payment_headers.sql`
- `bank_accounts.sql`
- `banks.sql`
- `batch_schedules.sql`
- `cities.sql`
- `collection_policies.sql`
- `companies.sql`
- `company_defaultor_configurations.sql`
- `contacts.sql`
- `countries.sql`
- `customer_bank_accounts.sql`
- `customer_contacts.sql`
- `customer_default_accounts.sql`
- `customer_note_details.sql`
- `customer_note_headers.sql`
- `customer_note_statuses.sql`
- `customer_note_work_flows.sql`
- `customer_upis.sql`
- `customers.sql`
- `customers_audit.sql`
- `defaulter_configurations.sql`
- `delinquency_action_types.sql`
- `delinquency_actions.sql`
- `delinquency_cycles.sql`
- `delinquency_discussion_messages.sql`
- `delinquency_resolution_states.sql`
- `delinquency_resolution_window.sql`
- `delinquency_snapshot_invoices.sql`
- `delinquency_snapshots.sql`
- `draft_invoice_details.sql`
- `draft_invoice_header_ids.sql`
- `draft_invoice_headers.sql`
- `gate_pass_details.sql`
- `gate_pass_headers.sql`
- `gate_pass_statuses.sql`
- `group_invoice_details.sql`
- `group_invoice_header_ids.sql`
- `group_invoice_headers.sql`
- `group_invoice_template_customers.sql`
- `group_invoice_templates.sql`
- `invoice_account_approval_levels.sql`
- `invoice_approval_issue_log.sql`
- `invoice_approval_logs.sql`
- `invoice_approval_user_company.sql`
- `invoice_approval_users_account.sql`
- `invoice_details.sql`
- `invoice_header_ids.sql`
- `invoice_headers.sql`
- `invoice_payment_details.sql`
- `invoice_payment_header_ids.sql`
- `invoice_payment_headers.sql`
- `invoice_penalties.sql`
- `invoice_status_company_configs.sql`
- `invoice_statuses.sql`
- `invoice_voucher_ids.sql`
- `invoice_workflow.sql`
- `organizations.sql`
- `payment_statuses.sql`
- `penalty_configs.sql`
- `penalty_frequencies.sql`
- `penalty_processing_logs.sql`
- `record_exists.sql`
- `recurrence_schedule_details.sql`
- `recurring_sales_schedules.sql`
- `roles.sql`
- `schedule_execution_logs.sql`
- `schema_versions.sql`
- `states.sql`
- `temp_invoice_next_status.sql`
- `upis.sql`
- `user_roles.sql`
- `users.sql`
- `v_contact_number.sql`
- `v_created_by.sql`
- `v_invoice_status_id.sql`
- `v_user_id.sql`
- `warehouses.sql`

## APIs / Controllers / Services Matrix

| Area | Feature / Object | Legacy VB | Dhanman | Gap | Priority | Evidence | Notes |
| --- | --- | --- | --- | --- | --- | --- | --- |
| APIs | Customer-facing sales APIs | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `CustomersController.cs`, `CustomerNoteController.cs`, `PaymentsController.cs` | Covers customer operations. |
| APIs | Billing and ledger APIs | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `InvoicesController.cs`, `LedgerController.cs` | Core sales financial flows. |
| APIs | Warehouse and gate pass APIs | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `Medium` | `GatePassController.cs`, `WareHouseController.cs` | Warehouse references must stay aligned with inventory. |
| APIs | Admin and support APIs | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `Medium` | `PermissionsController.cs`, `UsersController.cs`, `MasterController.cs`, `OrganizationController.cs` | Supporting admin APIs are present. |

## Screens / Forms / Pages Matrix

| Area | Feature / Object | Legacy VB | Dhanman | Gap | Priority | Evidence | Notes |
| --- | --- | --- | --- | --- | --- | --- | --- |
| Business workflow | Customer onboarding and maintenance | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `CustomersController.cs`, `AddressController.cs` | Use later to compare VB customer forms. |
| Business workflow | Invoice creation and follow-up | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `InvoicesController.cs`, `Entities/InvoiceWorkflows` | Likely one of the main parity tracks. |
| Business workflow | Collections and penalties | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `Medium` | `PaymentsController.cs`, `DefaultersController.cs` | Capture legacy dunning flow details later. |

## Reports / Export / Integration Matrix

| Area | Feature / Object | Legacy VB | Dhanman | Gap | Priority | Evidence | Notes |
| --- | --- | --- | --- | --- | --- | --- | --- |
| Reports | Invoice and customer account reports | To be verified from VB form/report/table | `Partial` | Cannot assess until legacy evidence is added | `High` | `LedgerController.cs`, `InvoicesController.cs` | Specific report endpoints are not yet obvious from surface evidence alone. |
| Reports | Defaulter and penalty reporting | To be verified from VB form/report/table | `Partial` | Cannot assess until legacy evidence is added | `Medium` | `DefaultersController.cs`, `Entities/PenaltyProcessingLogs` | Good candidate for later legacy mapping. |
| Integration | Inventory warehouse and stock linkage | To be verified from VB form/report/table | `Partial` | Cannot assess until legacy evidence is added | `Medium` | `WareHouseController.cs`, `dhanman-common/inventory-module-gap-analysis.md` | Cross-module warehouse consistency matters for migration. |

## Gaps And Recommendations

- Fill invoice statuses, approval levels, and penalty rules from VB first because they often hide business-critical differences.
- Track whether legacy sales flows posted stock directly or only referenced warehouse and gate pass data.
- Add a row per major legacy sales report once the VB report catalog is available.
- Record whether batch or recurring invoice operations existed in VB and how they were triggered.

## Evidence From Codebase

- `dhanman-sales/src/Dhanman.Sales.Api/Controllers`
- `dhanman-sales/src/Core/Dhanman.Sales.Domain/Entities`
- `dhanman-sales/src/Dhanman.Sales.Persistence/DBObjects/Table`
- `dhanman-common/inventory-module-gap-analysis.md`

## Open Legacy Questions

- Which VB forms created invoices, grouped invoices, and customer notes?
- Did VB include defaulter, penalty, and recurring billing workflows?
- Which reports were available for invoices, receipts, and customer ledgers?
- Did the legacy sales system maintain its own warehouse master or consume a shared one?
