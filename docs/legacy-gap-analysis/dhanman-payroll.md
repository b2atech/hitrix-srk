# dhanman-payroll Legacy Comparison

## Project Summary

`dhanman-payroll` is the payroll backend service. It includes employee, payroll component, attendance, leave, timesheet, payroll transaction, and organization support models.

## Current Dhanman Scope

- API controllers in `src/Dhanman.Payroll.Api/Controllers`
- domain entities in `src/Core/Dhanman.Payroll.Domain/Entities`
- persistence and infrastructure under `src/Dhanman.Payroll.Persistence` and `src/Infrastructure`

## Legacy VB Scope Placeholder

Legacy VB scope is not available in this workspace.

Expected legacy evidence to capture later:

- employee master setup
- payroll component configuration
- leave and attendance flows
- timesheet or project hour tracking
- payroll run and payment reporting

## Feature Comparison Matrix

| Area | Feature / Object | Legacy VB | Dhanman | Gap | Priority | Evidence | Notes |
| --- | --- | --- | --- | --- | --- | --- | --- |
| HR setup | Employees, addresses, companies | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `EmployeeController.cs`, `AddressController.cs`, `CompanyController.cs`, `Entities/Employees` | Core HR master support exists. |
| Payroll setup | Payroll components and component types | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `PayrollComponentController.cs`, `ComponentTypeController.cs`, `Entities/PayrollComponents`, `Entities/ComponentTypes` | Strong payroll configuration evidence. |
| Attendance | Timesheet and attendance support | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `TimesheetController.cs`, `Entities/Attendance`, `Entities/Timesheets`, `Entities/TimeSheetLines` | Important operational track. |
| Leave | Leave setup and workflows | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `Medium` | `Entities/LeaveApplications`, `Entities/LeaveEntitlements`, `Entities/LeaveWorkFlows` | Use later for deeper parity review. |
| Payroll processing | Payroll transactions and payments | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `Entities/PayrollTransactions`, `Entities/PayrollTransactionDetails`, `Entities/Payments` | Important for run-cycle comparison. |

## Tables And Entities Matrix

| Area | Feature / Object | Legacy VB | Dhanman | Gap | Priority | Evidence | Notes |
| --- | --- | --- | --- | --- | --- | --- | --- |
| Masters | Employee and organization entities | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `Entities/Employees`, `Entities/Companies`, `Entities/Departments`, `Entities/Designations` | Good anchor for HR migration mapping. |
| Payroll config | Payroll components and config entities | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `Entities/PayrollComponents`, `Entities/PayrollConfigs`, `Entities/PayrollConfigurationHeaders`, `Entities/PayrollConfigurationDetails` | Strong configuration layer evidence. |
| Operations | Attendance, leave, and holiday entities | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `Medium` | `Entities/Attendance`, `Entities/LeaveApplications`, `Entities/Holidays`, `Entities/Workdays` | Daily operational data model is visible. |
| Processing | Payroll run and payment entities | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `Entities/PayrollTransactions`, `Entities/PaymentHeaders`, `Entities/PaymentDetails` | Key run and settlement model. |
| Time tracking | Project and timesheet entities | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `Medium` | `Entities/Projects`, `Entities/SubProjects`, `Entities/Tasks`, `Entities/Timesheets` | Helps compare job-costing or project-hour legacy workflows. |
| Physical tables | Payroll DB table scripts | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `employees.sql`, `payroll_components.sql`, `payroll_transaction_headers.sql`, `timesheets.sql`, `leave_applications.sql`, `payments.sql` | Concrete SQL scripts exist for core payroll operations. |

## All Tables

All SQL table/object scripts currently visible under `dhanman-payroll/src/Dhanman.Payroll.Persistence/DBObjects/Table`:

- `__EFMigrationsHistory.sql`
- `addresses.sql`
- `advance_payments.sql`
- `attendance.sql`
- `bank_accounts.sql`
- `banks.sql`
- `benefits.sql`
- `cities.sql`
- `companies.sql`
- `component_types.sql`
- `countries.sql`
- `deductions.sql`
- `departments.sql`
- `designations.sql`
- `employee_bank_accounts.sql`
- `employee_leave_balances.sql`
- `employee_upis.sql`
- `employees.sql`
- `holidays.sql`
- `leave_applications.sql`
- `leave_entitlements.sql`
- `leave_records.sql`
- `leave_statuses.sql`
- `leave_types.sql`
- `leave_workflow.sql`
- `loans.sql`
- `organizations.sql`
- `payment_details.sql`
- `payment_header_ids.sql`
- `payment_headers.sql`
- `payments.sql`
- `payroll_components.sql`
- `payroll_configuration_details.sql`
- `payroll_configuration_headers.sql`
- `payroll_transaction_details.sql`
- `payroll_transaction_header_ids.sql`
- `payroll_transaction_headers.sql`
- `projects.sql`
- `states.sql`
- `statuses.sql`
- `sub_projects.sql`
- `tasks.sql`
- `timesheet_hours.sql`
- `timesheet_lines.sql`
- `timesheet_statuses.sql`
- `timesheets.sql`
- `upis.sql`
- `users.sql`
- `workdays.sql`

## APIs / Controllers / Services Matrix

| Area | Feature / Object | Legacy VB | Dhanman | Gap | Priority | Evidence | Notes |
| --- | --- | --- | --- | --- | --- | --- | --- |
| APIs | Employee and company APIs | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `EmployeeController.cs`, `CompanyController.cs`, `OrganizationController.cs` | Core admin surfaces. |
| APIs | Payroll configuration APIs | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `PayrollComponentController.cs`, `ComponentTypeController.cs`, `MasterController.cs` | Supports payroll setup. |
| APIs | Timesheet and user APIs | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `Medium` | `TimesheetController.cs`, `UsersController.cs`, `PermissionsController.cs` | Supporting operational and admin APIs. |

## Screens / Forms / Pages Matrix

| Area | Feature / Object | Legacy VB | Dhanman | Gap | Priority | Evidence | Notes |
| --- | --- | --- | --- | --- | --- | --- | --- |
| Business workflow | Employee onboarding and maintenance | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `EmployeeController.cs`, `AddressController.cs` | Compare with VB employee forms later. |
| Business workflow | Payroll setup and calculation | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `PayrollComponentController.cs`, `Entities/PayrollConfigs` | Likely major parity area. |
| Business workflow | Attendance and timesheet capture | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `Medium` | `TimesheetController.cs`, `Entities/Attendance` | Confirm whether VB used timesheets, attendance, or both. |

## Reports / Export / Integration Matrix

| Area | Feature / Object | Legacy VB | Dhanman | Gap | Priority | Evidence | Notes |
| --- | --- | --- | --- | --- | --- | --- | --- |
| Reports | Payroll run and payment reporting | To be verified from VB form/report/table | `Partial` | Cannot assess until legacy evidence is added | `High` | `Entities/PayrollTransactions`, `Entities/Payments` | Exact report APIs are not yet obvious from the current surface evidence. |
| Reports | Leave and attendance reporting | To be verified from VB form/report/table | `Partial` | Cannot assess until legacy evidence is added | `Medium` | `Entities/LeaveApplications`, `Entities/Attendance`, `Entities/Timesheets` | Add legacy report names later. |

## Gaps And Recommendations

- Fill the legacy payroll process in run order: employee setup, payroll config, attendance/leave, payroll run, payment, reports.
- Track exact earning, deduction, and leave-status codes used in VB because those usually affect migration logic.
- Add a note for whether payroll was project-based, attendance-based, or fixed-salary-only in the old system.
- Capture statutory or compliance reports separately if they existed in VB.

## Evidence From Codebase

- `dhanman-payroll/src/Dhanman.Payroll.Api/Controllers`
- `dhanman-payroll/src/Core/Dhanman.Payroll.Domain/Entities`
- `dhanman-payroll/src/Dhanman.Payroll.Persistence/DBObjects/Table`

## Open Legacy Questions

- Which VB forms maintained employees, payroll components, leave, and attendance?
- How were payroll runs approved and finalized in the old system?
- Which payroll reports were mandatory for business and compliance use?
