# dhanman-common Legacy Comparison

## Project Summary

`dhanman-common` is a broad shared backend service covering authentication, company and organization setup, finance, users, documents, notifications, token relay, and warehouse projections.

## Current Dhanman Scope

- API controllers in `src/Dhanman.Common.Api/Controllers`
- domain entities in `src/Dhanman.Common.Domain/Entities`
- application and persistence layers under `src/Dhanman.Common.Application*` and `src/Dhanman.Common.Persistence`

## Legacy VB Scope Placeholder

Legacy VB scope is not available in this workspace.

Expected legacy evidence to capture later:

- shared masters used across modules
- login and user setup
- company or organization setup screens
- finance setup and reporting
- notification, document, and scheduler utilities

## Feature Comparison Matrix

| Area | Feature / Object | Legacy VB | Dhanman | Gap | Priority | Evidence | Notes |
| --- | --- | --- | --- | --- | --- | --- | --- |
| Security | Authentication, roles, permissions, tokens | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `AuthenticationController.cs`, `RolesController.cs`, `PermissionsController.cs`, `TokenController.cs` | Strong security and authorization surface. |
| Company setup | Company, organization, preferences | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `CompanyController.cs`, `CompanyPreferencesController.cs`, `OrganizationController.cs` | Core setup flows exist. |
| Finance | Accounts, journals, ledgers, transactions | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `AccountsController.cs`, `JournalVoucherController.cs`, `GeneralLedgersController.cs`, `TransactionController.cs` | Large finance capability footprint. |
| Finance | Budgeting and bank reconciliation | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `Medium` | `BudgetController.cs`, `BankReconciliationController.cs`, `BankStatementController.cs` | Useful for cross-module parity review. |
| Utility | Notifications, communication, documents, scheduler | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `Medium` | `NotificationController.cs`, `CommunicationController.cs`, `DocumentController.cs`, `SchedulerController.cs` | Shared utility capabilities are visible. |
| Integration | Warehouse projection and token relay | To be verified from VB form/report/table | `Partial` | Cannot assess until legacy evidence is added | `Medium` | `WareHouseController.cs`, `TokenRelayController.cs`, `inventory-module-remaining-gaps.md` | Important for cross-service sync parity. |

## Tables And Entities Matrix

| Area | Feature / Object | Legacy VB | Dhanman | Gap | Priority | Evidence | Notes |
| --- | --- | --- | --- | --- | --- | --- | --- |
| Security | Users, roles, permissions | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `Entities/Users`, `Entities/Roles`, `Entities/Permissions`, `Entities/UserRoles`, `Entities/RolePermissions` | Core identity model is well represented. |
| Company setup | Companies, organizations, contacts | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `Entities/Companies`, `Entities/Organizations`, `Entities/Contacts`, `Entities/CompanyUsers` | Useful anchor for legacy master data mapping. |
| Finance | Accounts and ledgers | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `Entities/Accounts`, `Entities/ChartOfAccounts`, `Entities/GeneralLedgers`, `Entities/JournalVoucherHeaders` | Finance model appears mature. |
| Finance | Budget and reconciliation | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `Medium` | `Entities/Budgets`, `Entities/BudgetLines`, `Entities/BankReconciliations`, `Entities/BankStatements` | Good parity candidate for finance admin flows. |
| Utility | Notifications, email rules, document metadata | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `Medium` | `Entities/Notifications`, `Entities/EmailNotificationRules`, `Entities/DocumentMetaDatas` | Important when comparing operational support features. |
| Projection | Warehouses and customers/vendors | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `Medium` | `Entities/WareHouses`, `Entities/Customers`, `Entities/Vendors` | Helps track how common data is shared across services. |
| Physical tables | Common DB table scripts | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `account_categories.sql`, `chart_of_accounts.sql`, `general_ledgers.sql`, `transaction_headers.sql`, `company_preferences.sql`, `document_meta_datas.sql`, `warehouses.sql` | Concrete SQL scripts are available for the main common service areas. |

## All Tables

All SQL table/object scripts currently visible under `dhanman-common/src/Dhanman.Common.Persistence/DBObjects/Table`:

- `__EFMigrationsHistory.sql`
- `account_categories.sql`
- `account_groups.sql`
- `account_opening_balances.sql`
- `account_types.sql`
- `addresses.sql`
- `audit_queries.sql`
- `audit_query_responses.sql`
- `audit_query_statuses.sql`
- `bank_accounts.sql`
- `bank_reconciliation_links.sql`
- `bank_reconciliation_stagings.sql`
- `bank_reconciliations.sql`
- `bank_statements.sql`
- `bank_transfer_ids.sql`
- `bank_transfers.sql`
- `banks.sql`
- `budget_lines.sql`
- `budget_statuses.sql`
- `budget_workflow.sql`
- `budgets.sql`
- `chart_of_accounts.sql`
- `cities.sql`
- `companies.sql`
- `company_bank_accounts.sql`
- `company_contacts.sql`
- `company_finance_year.sql`
- `company_preferences.sql`
- `company_upis.sql`
- `company_users.sql`
- `contacts.sql`
- `countries.sql`
- `currencies.sql`
- `customers.sql`
- `default_organization_users.sql`
- `departments.sql`
- `description_templates.sql`
- `designations.sql`
- `document_meta_datas.sql`
- `dummy_log_table.sql`
- `dummy_test_table.sql`
- `email_templates.sql`
- `employees.sql`
- `entry_sources.sql`
- `finance_year.sql`
- `fixed_deposits.sql`
- `general_ledgers.sql`
- `images.sql`
- `journal_entries.sql`
- `journal_entries_2022_2023.sql`
- `journal_entries_2023_2024.sql`
- `journal_entries_2024_2025.sql`
- `journal_entries_2025_2026.sql`
- `journal_narrations.sql`
- `journal_voucher_details.sql`
- `journal_voucher_header_id.sql`
- `journal_voucher_headers.sql`
- `logs.sql`
- `messages.sql`
- `notification_types.sql`
- `notifications.sql`
- `opening_balances.sql`
- `organization_accounts.sql`
- `organization_types.sql`
- `organization_users.sql`
- `organizations.sql`
- `payment_references.sql`
- `permission_groups.sql`
- `permission_template_mappings.sql`
- `permission_templates.sql`
- `permissions.sql`
- `reconciliation_statuses.sql`
- `roles.sql`
- `schema_versions.sql`
- `states.sql`
- `temp_organization_logs.sql`
- `temp_paymnet_data.sql`
- `transaction_header_2022_2023.sql`
- `transaction_header_2023_2024.sql`
- `transaction_header_2024_2025.sql`
- `transaction_header_2025_2026.sql`
- `transaction_headers.sql`
- `transaction_headers_19_05_25.sql`
- `transaction_headers_bk_20_5_25.sql`
- `transaction_source_types.sql`
- `upis.sql`
- `user_deletion_request_statuses.sql`
- `user_deletion_requests.sql`
- `user_global_permissions.sql`
- `user_permissions.sql`
- `user_roles.sql`
- `users.sql`
- `users_bk.sql`
- `v_finance_year_id.sql`
- `v_is_second_last_leaf.sql`
- `v_permission_id.sql`
- `vendors.sql`
- `warehouses.sql`

## APIs / Controllers / Services Matrix

| Area | Feature / Object | Legacy VB | Dhanman | Gap | Priority | Evidence | Notes |
| --- | --- | --- | --- | --- | --- | --- | --- |
| APIs | Security and user APIs | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `AuthenticationController.cs`, `UsersController.cs`, `PublicUserController.cs`, `TestTokenController.cs` | Broad user and auth entry points. |
| APIs | Company and organization APIs | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `CompanyController.cs`, `OrganizationController.cs`, `AppContextController.cs` | Important setup and context surfaces. |
| APIs | Finance processing APIs | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `AccountsController.cs`, `TransactionController.cs`, `FinanceReportsController.cs`, `JournalVoucherController.cs` | Likely heavy parity area if VB included accounting functions. |
| APIs | Communication and notification APIs | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `Medium` | `CommunicationController.cs`, `NotificationController.cs`, `DocumentController.cs` | Shared support services. |
| APIs | Warehouse and scheduler APIs | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `Medium` | `WareHouseController.cs`, `SchedulerController.cs`, `TokenRelayController.cs` | Useful for operational and integration comparisons. |

## Screens / Forms / Pages Matrix

| Area | Feature / Object | Legacy VB | Dhanman | Gap | Priority | Evidence | Notes |
| --- | --- | --- | --- | --- | --- | --- | --- |
| Shared admin | User and role maintenance | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `UsersController.cs`, `RolesController.cs`, `PermissionsController.cs` | Compare with VB admin forms later. |
| Shared admin | Company and organization setup | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `CompanyController.cs`, `OrganizationController.cs`, `CompanyPreferencesController.cs` | Important migration baseline. |
| Shared admin | Finance admin and reporting | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `AccountsController.cs`, `FinanceReportsController.cs`, `BudgetController.cs` | May represent large legacy scope. |

## Reports / Export / Integration Matrix

| Area | Feature / Object | Legacy VB | Dhanman | Gap | Priority | Evidence | Notes |
| --- | --- | --- | --- | --- | --- | --- | --- |
| Reports | Finance reports | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `FinanceReportsController.cs`, `Entities/FinanceReports` | Add specific legacy report names later. |
| Reports | Bank reconciliation outputs | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `Medium` | `BankReconciliationController.cs`, `BankStatementController.cs` | Important if legacy finance ops relied on them. |
| Integration | Token relay and scheduler hooks | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `Medium` | `TokenRelayController.cs`, `SchedulerController.cs`, `TASK-TRACKING.md` | Cross-service operational feature set. |
| Integration | Warehouse sync projection | To be verified from VB form/report/table | `Partial` | Cannot assess until legacy evidence is added | `Medium` | `WareHouseController.cs`, `inventory-module-remaining-gaps.md` | Remaining runtime verification is already documented. |

## Gaps And Recommendations

- Split legacy comparison into admin, finance, and utility tracks when your team fills this file because the service covers many domains.
- Map VB finance reports carefully; they are likely broader than what the current repo evidence alone makes obvious.
- Use the warehouse projection entries to track whether shared masters were centrally owned in VB or duplicated per module.
- If VB had scheduler or batch jobs, add them in `Notes` because those often drive hidden parity gaps.

## Evidence From Codebase

- `dhanman-common/src/Dhanman.Common.Api/Controllers`
- `dhanman-common/src/Dhanman.Common.Domain/Entities`
- `dhanman-common/src/Dhanman.Common.Persistence/DBObjects/Table`
- `dhanman-common/inventory-module-gap-analysis.md`
- `dhanman-common/inventory-module-remaining-gaps.md`
- `dhanman-common/TASK-TRACKING.md`

## Open Legacy Questions

- Which VB modules owned users, permissions, and company setup?
- Did the VB system centralize finance and reporting in one shared project?
- Which VB warehouse records were shared across inventory, sales, and purchase?
- Were documents, notifications, and scheduler jobs part of the same legacy application or separate tools?
