# dhanman-payment Legacy Comparison

## Project Summary

`dhanman-payment` is the payment backend service. It appears focused on payment collection, mobile payment handling, payouts, permissions, and webhook processing.

## Current Dhanman Scope

- API controllers in `src/Dhanman.Payment.Api/Controllers`
- domain entities in `src/Core/Dhanman.Payment.Domain/Entities`
- persistence and infrastructure under `src/Dhanman.Payment.Persistence` and `src/Infrastructure`

## Legacy VB Scope Placeholder

Legacy VB scope is not available in this workspace.

Expected legacy evidence to capture later:

- collection entry screens
- payment initiation and response handling
- payout workflows
- payment status and reconciliation screens

## Feature Comparison Matrix

| Area | Feature / Object | Legacy VB | Dhanman | Gap | Priority | Evidence | Notes |
| --- | --- | --- | --- | --- | --- | --- | --- |
| Payments | Collection flows | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `CollectController.cs` | Primary collections API is visible. |
| Payments | Mobile payment handling | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `MobilePaymentsController.cs` | Indicates mobile payment workflow support. |
| Payments | Payout processing | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `Medium` | `PayoutController.cs` | Capture legacy outbound payment handling later. |
| Integration | Webhook processing | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `WebhookController.cs` | Useful for comparing current gateway-driven workflows to VB/manual flows. |
| Security | User and permission support | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `Low` | `UsersController.cs`, `PermissionsController.cs` | Supporting admin capability. |

## Tables And Entities Matrix

| Area | Feature / Object | Legacy VB | Dhanman | Gap | Priority | Evidence | Notes |
| --- | --- | --- | --- | --- | --- | --- | --- |
| Payments | Payment initiation and result entities | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `Entities/PaymentInitiationHeaders`, `Entities/PaymentInitiationDetails`, `Entities/PaymentResults` | Good baseline for transaction-state mapping. |
| Payments | Payment transaction and status entities | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `Entities/PaymentTransactions`, `Entities/PaymentStatuses`, `Entities/TransactionTypes` | Core payment tracking model. |
| Storage | Physical DB object scripts | To be verified from VB form/report/table | `Partial` | Cannot assess until legacy evidence is added | `Medium` | `dhanman-payment/src/Dhanman.Payment.Persistence/DBObjects` | `DBObjects` exists, but no `Table` folder was visible during this evidence pass. |
| Projection | Users and warehouse references | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `Low` | `Entities/Users`, `Entities/WareHouses` | Supporting references only. |

## All DB Objects

All SQL/object scripts currently visible under `dhanman-payment/src/Dhanman.Payment.Persistence/DBObjects`:

- `invoice_approval_level_view.sql`
- `vw_invoice_approval_permissions.sql`

No `DBObjects/Table` folder was visible in the current repo path during this pass.

## APIs / Controllers / Services Matrix

| Area | Feature / Object | Legacy VB | Dhanman | Gap | Priority | Evidence | Notes |
| --- | --- | --- | --- | --- | --- | --- | --- |
| APIs | Collection APIs | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `CollectController.cs` | Compare with legacy cash/UPI/card collection forms later. |
| APIs | Mobile payment and webhook APIs | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `MobilePaymentsController.cs`, `WebhookController.cs` | Integration-driven payment surface. |
| APIs | Payout and admin APIs | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `Medium` | `PayoutController.cs`, `UsersController.cs`, `PermissionsController.cs` | Supporting operational APIs. |

## Screens / Forms / Pages Matrix

| Area | Feature / Object | Legacy VB | Dhanman | Gap | Priority | Evidence | Notes |
| --- | --- | --- | --- | --- | --- | --- | --- |
| Business workflow | Payment initiation and collection | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `CollectController.cs`, `Entities/PaymentInitiationHeaders` | Start VB comparison with payment initiation rules and statuses. |
| Business workflow | Gateway callback processing | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `Medium` | `WebhookController.cs`, `Entities/PaymentResults` | Important if VB used manual response imports instead. |
| Business workflow | Payout execution | To be verified from VB form/report/table | `Partial` | Cannot assess until legacy evidence is added | `Medium` | `PayoutController.cs` | Needs detailed legacy evidence to assess parity. |

## Reports / Export / Integration Matrix

| Area | Feature / Object | Legacy VB | Dhanman | Gap | Priority | Evidence | Notes |
| --- | --- | --- | --- | --- | --- | --- | --- |
| Reports | Payment status and transaction reporting | To be verified from VB form/report/table | `Partial` | Cannot assess until legacy evidence is added | `High` | `Entities/PaymentTransactions`, `Entities/PaymentStatuses` | Explicit report endpoints are not yet obvious from the current surface evidence. |
| Integration | Mobile gateway integrations | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `MobilePaymentsController.cs`, `WebhookController.cs` | Useful migration comparison area if VB was batch/manual. |

## Gaps And Recommendations

- Capture whether VB collections were manual, bank-file based, or gateway-based before trying to compare field-level tables.
- Use the payment status entities to map old lifecycle states to current Dhanman states.
- Record every external payment provider used by the legacy system in `Notes`.
- If the legacy process had no webhook model, track which reconciliation or import step replaced it.

## Evidence From Codebase

- `dhanman-payment/src/Dhanman.Payment.Api/Controllers`
- `dhanman-payment/src/Core/Dhanman.Payment.Domain/Entities`
- `dhanman-payment/src/Dhanman.Payment.Persistence`

## Open Legacy Questions

- Which VB forms captured collections, mobile payments, and payouts?
- Did the legacy system use real-time gateway callbacks or manual reconciliation?
- Which reports existed for payment success, failure, retry, and settlement?
