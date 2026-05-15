# Dhanman Legacy Gap Analysis Pack

## Purpose

This documentation pack is used to compare the current Dhanman codebase with the legacy Visual Basic system on a service-by-service basis.

The VB source code is not available in this workspace, so this first version documents only the Dhanman side from repo evidence and leaves structured placeholders for the legacy review team to complete later.

## Comparison Method

Each service file uses the same comparison structure:

1. Project summary
2. Current Dhanman scope
3. Legacy VB scope placeholder
4. Feature comparison matrix
5. Tables and entities matrix
6. APIs/controllers/services matrix
7. Screens/forms/pages matrix
8. Reports/export/integration matrix
9. Gaps and recommendations
10. Evidence from codebase
11. Open legacy questions

### Status Meanings

| Status | Meaning |
| --- | --- |
| `Present` | Clear implementation evidence exists in the current Dhanman repo. |
| `Partial` | Some implementation exists, but workflow/UI/reporting/integration is incomplete. |
| `Missing` | No clear implementation evidence was found in the current Dhanman repo. |
| `Unknown` | Cannot be confirmed from the current repo evidence available. |

### Priority Meanings

| Priority | Meaning |
| --- | --- |
| `High` | Important for business parity, day-to-day operation, or migration readiness. |
| `Medium` | Valuable for service completeness, productivity, or rollout quality. |
| `Low` | Useful later, but not a blocker for initial comparison or parity tracking. |

## How To Fill The Legacy VB Columns Later

When the legacy system is reviewed, update the `Legacy VB` column with concrete evidence only:

- VB module or project name
- form or screen name
- report name
- table name
- stored procedure or query object
- menu path
- screenshot reference
- sample workflow notes

Use explicit placeholder wording until evidence is collected:

- `To be verified from VB form/report/table`
- `Cannot assess until legacy evidence is added`
- `Capture VB menu path, table name, report name, or screenshot`

Do not mark a legacy feature as present without a source.

## Service Files

- [dhanman-app](./dhanman-app.md)
- [dhanman-common](./dhanman-common.md)
- [dhanman-inventory](./dhanman-inventory.md)
- [dhanman-sales](./dhanman-sales.md)
- [dhanman-purchase](./dhanman-purchase.md)
- [dhanman-payment](./dhanman-payment.md)
- [dhanman-payroll](./dhanman-payroll.md)
- [dhanman-myhome](./dhanman-myhome.md)
- [dhanman-documents](./dhanman-documents.md)
- [dhanman-shared-contracts](./dhanman-shared-contracts.md)

## Usage Notes

- Start with the service that maps most closely to the legacy business area being reviewed.
- Fill the `Legacy VB` column first, then update `Gap`, `Priority`, and `Notes`.
- Keep `Evidence` paths accurate and repo-relative so engineers can jump directly to the current implementation.
- If a service depends on another one, note the dependency in `Notes` rather than duplicating the same analysis in multiple files.
