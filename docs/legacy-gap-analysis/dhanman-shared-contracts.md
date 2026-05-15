# dhanman-shared-contracts Legacy Comparison

## Project Summary

`dhanman-shared-contracts` is the shared message and contract library used across Dhanman services. It defines events, commands, and routing helpers for cross-service communication.

## Current Dhanman Scope

- events in `src/Dhanman.Shared.Contracts/Events`
- commands in `src/Dhanman.Shared.Contracts/Commands`
- routing helpers in `src/Dhanman.Shared.Contracts/Routing`

## Legacy VB Scope Placeholder

Legacy VB scope is not available in this workspace.

Expected legacy evidence to capture later:

- shared DTOs or COM contracts
- integration message formats
- batch scheduler command inputs
- cross-module interface definitions

## Feature Comparison Matrix

| Area | Feature / Object | Legacy VB | Dhanman | Gap | Priority | Evidence | Notes |
| --- | --- | --- | --- | --- | --- | --- | --- |
| Events | Organization, company, user bootstrap events | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `Medium` | `BasicOrganizationCreatedEvent.cs`, `BasicCompanyCreatedEvent.cs`, `UserCreatedEvent.cs`, `OrganizationInitializedEvent.cs` | Shared lifecycle events are clearly modeled. |
| Events | Sales and purchase posting events | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `PurchaseReceiptPostedEvent.cs`, `SalesInvoicePostedEvent.cs`, `TransactionPostedEvent.cs` | Important cross-module parity anchor. |
| Events | Warehouse sync events | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `WarehouseCreatedEvent.cs`, `WarehouseUpdatedEvent.cs`, `WarehouseDeactivatedEvent.cs` | Strong evidence of inventory-owned master sync. |
| Commands | Scheduler and business-entity commands | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `Medium` | `InitializeOrganizationCommand.cs`, `RunCommonSchedulerCommand.cs`, `CreateBusinessEntityCommand.cs`, `CreateUserCommand.cs` | Covers setup and scheduled processing. |
| Routing | Shared routing keys | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `Medium` | `RoutingKeys.cs` | Key for comparing cross-module integration patterns. |

## Tables And Entities Matrix

| Area | Feature / Object | Legacy VB | Dhanman | Gap | Priority | Evidence | Notes |
| --- | --- | --- | --- | --- | --- | --- | --- |
| DTOs | Shared event contract files | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `BasicCompanyCreatedEvent.cs`, `PurchaseReceiptPostedEvent.cs`, `SalesInvoicePostedEvent.cs`, `WarehouseCreatedEvent.cs` | Contract library is file-based DTO storage rather than database-table based. |
| DTOs | Shared command contract files | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `Medium` | `CreateUserCommand.cs`, `InitializeOrganizationCommand.cs`, `RunInventorySchedulerCommand.cs`, `RunSalesSchedulerCommand.cs` | Shared command set is explicit and concrete. |
| Common objects | DTO and request object files | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `Medium` | `Dtos/BasicCompanyDto.cs`, `Dtos/TransactionDto.cs`, `Common/Requests/TransactionHeaderRequest.cs`, `Common/Requests/InvoiceNotificationRequest.cs` | These are the closest equivalents to shared interface objects if VB had common contract files. |
| Routing | Routing key definitions | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `Medium` | `src/Dhanman.Shared.Contracts/Routing/RoutingKeys.cs`, `Common/Constants/EventRoutingKeys.cs` | Compare with any legacy MQ, file-drop, or COM routing conventions. |

## All Tables

No database table scripts are expected in `dhanman-shared-contracts`.

This project is contract/object-based. The closest comparison inventory is:

- `Events/*.cs`
- `Commands/*.cs`
- `Dtos/*.cs`
- `Common/Requests/*.cs`
- `Routing/RoutingKeys.cs`
- `Common/Constants/EventRoutingKeys.cs`

## APIs / Controllers / Services Matrix

| Area | Feature / Object | Legacy VB | Dhanman | Gap | Priority | Evidence | Notes |
| --- | --- | --- | --- | --- | --- | --- | --- |
| Contracts | Event-driven service integration | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `Events/PurchaseReceiptPostedEvent.cs`, `Events/SalesInvoicePostedEvent.cs`, `Events/WarehouseCreatedEvent.cs` | Core integration style is event-based. |
| Contracts | Command-driven scheduler operations | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `Medium` | `Commands/RunCommonSchedulerCommand.cs`, `RunInventorySchedulerCommand.cs`, `RunSalesSchedulerCommand.cs`, `RunPurchaseSchedulerCommand.cs`, `RunPayrollSchedulerCommand.cs` | Shared scheduled-processing contract set. |
| Contracts | Payment processing commands | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `Medium` | `Commands/PurchasePaymentProcessingCommand.cs`, `Commands/SalesPaymentProcessingCommand.cs` | Useful if VB had batch or queue-driven payment jobs. |

## Screens / Forms / Pages Matrix

| Area | Feature / Object | Legacy VB | Dhanman | Gap | Priority | Evidence | Notes |
| --- | --- | --- | --- | --- | --- | --- | --- |
| Integration workflow | Shared bootstrap and provisioning | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `Medium` | `InitializeOrganizationCommand.cs`, `BasicCompanyCreatedEvent.cs`, `UserCustomerMappingEvent.cs` | Compare to legacy setup sequencing. |
| Integration workflow | Inventory warehouse sync | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `WarehouseCreatedEvent.cs`, `WarehouseUpdatedEvent.cs`, `WarehouseDeactivatedEvent.cs` | Important if legacy used shared warehouse masters. |
| Integration workflow | Sales and purchase stock-impact messages | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `PurchaseReceiptPostedEvent.cs`, `SalesInvoicePostedEvent.cs` | Helps compare old cross-module posting behavior. |

## Reports / Export / Integration Matrix

| Area | Feature / Object | Legacy VB | Dhanman | Gap | Priority | Evidence | Notes |
| --- | --- | --- | --- | --- | --- | --- | --- |
| Integration | Routing standardization | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `Medium` | `RoutingKeys.cs` | Use this to compare shared channel naming or transport patterns. |
| Integration | Event versioning and shared payloads | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `Events/*.cs` | Important if legacy had file-based or tightly coupled interfaces instead of shared contracts. |

## Gaps And Recommendations

- Document how the legacy system exchanged data between modules before assessing field-level parity.
- If VB used database sharing instead of messages, note that clearly because it changes migration planning.
- Map legacy scheduler or batch-job inputs against the existing shared command set.
- Track versioning and routing differences separately from business-feature gaps.

## Evidence From Codebase

- `dhanman-shared-contracts/src/Dhanman.Shared.Contracts/Events`
- `dhanman-shared-contracts/src/Dhanman.Shared.Contracts/Commands`
- `dhanman-shared-contracts/src/Dhanman.Shared.Contracts/Routing/RoutingKeys.cs`

## Open Legacy Questions

- Did the VB landscape use shared DTOs, shared database tables, COM contracts, or message queues?
- Which cross-module integrations were synchronous versus batch-driven?
- Were scheduler and payment-processing jobs centrally defined or module-specific?
