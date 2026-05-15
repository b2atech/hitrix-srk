# DhanMan Context Document for AI / Coding Agents

> **Purpose:**  
> This document gives any AI agent, developer, or reviewer enough context to understand what **DhanMan** is, how it is structured, what technology stack it uses, what architectural patterns are expected, and how future work should be implemented without breaking existing functionality.

---

## 1. What is DhanMan?

**DhanMan** is a multi-module, multi-service platform being built under **B2A Technologies Private Limited**.

DhanMan started as a financial management system for apartment complexes and has evolved into a broader **apartment / housing society management platform** with a strong focus on:

- Financial transparency
- Apartment governance
- Resident participation
- Maintenance tracking
- Visitor and gate management
- Notices and communication
- Amenity booking
- Approval workflows
- Accounting and reporting
- Document management
- Community decision tracking

The product vision is not just operational management. The long-term vision is to **digitize governance in housing societies**.

DhanMan is intended to help housing societies become more structured, accountable, transparent, and institution-like over time.

---

## 2. Product Vision

DhanMan is not only a day-to-day apartment operations tool. It aims to become a platform that preserves the institutional memory of a housing society.

The broader goals are:

1. **Transparency**
   - Make financial, operational, and governance-related information visible to the right people.
   - Reduce ambiguity in society operations.

2. **Accountability**
   - Track who made decisions, who approved them, and when.
   - Preserve audit trails.

3. **Participation**
   - Allow residents, committee members, facility managers, guards, and vendors to participate in workflows according to their roles.

4. **Governance**
   - Support structured decision-making through meetings, polls, approvals, and records.
   - Avoid informal WhatsApp-only decision-making where context is lost.

5. **Continuity**
   - Preserve community memory even when committee members change.
   - Maintain historical records of decisions, payments, approvals, and documents.

6. **Operational Efficiency**
   - Reduce manual follow-ups for payments, visitors, notices, requests, documents, and approvals.

---

## 3. Company and Product Identity

| Item | Details |
|---|---|
| Company | B2A Technologies Private Limited |
| Product | DhanMan |
| Domain | dhanman.com |
| Main customer segment | Apartment complexes / housing societies |
| Current customer geography | India |
| Hosting | OVH Cloud |
| Business model | Subscription-based model for apartment complexes |

---

## 4. Main Functional Areas

DhanMan currently includes or is planned to include the following major functional areas.

### 4.1 Financial Management

This is one of the core modules.

It includes:

- Apartment maintenance invoicing
- Customer invoices
- Vendor bills
- Payments
- Receipts
- Bank reconciliation
- Bank transfers
- Manual journal vouchers
- Trial balance
- Balance sheet
- Profit and loss reports
- Customer ledgers
- Vendor ledgers
- Employee ledgers
- Comparative income and expense reports
- Actual income vs actual expense tracking
- Monthly income / expense overview

Financial management is designed to work for apartment complexes first, but the architecture is modular enough to support other business entities.

---

### 4.2 Apartment / Community Management

The apartment management functionality includes:

- Apartment setup
- Company setup under organization
- Unit / flat management
- Resident management
- Member requests
- Resident requests
- Notices
- Maintenance-related communication
- Amenity booking
- Governance workflows
- Decision records

---

### 4.3 Gate and Visitor Management

DhanMan includes gate and visitor management for apartment complexes.

Features include:

- Visitor entry
- Visitor exit
- Visitor types
- Gate management
- Visitor approval by residents
- Guard-based verification
- Public visitor self-entry using QR code
- Visitor status tracking
- SignalR-based real-time visitor status updates
- Push notifications using FCM tokens

A key flow is the **Public Visitor QR Self-Entry Flow**, where a visitor can scan a QR code and submit entry details without being a registered system user.

---

### 4.4 Approval Workflows

DhanMan includes approval configuration for Sales and Purchase workflows.

Approval can exist at:

- Company level
- Account level

Approval flows may include:

- Basic approval
- Advanced approval
- Multiple approval levels
- Committee member assignment
- Level 1, Level 2, Level 3 approvers
- Payment-level approval
- Visual flow representation using React Flow or Material UI Stepper

Important design expectation:

- Keep approval configuration flexible.
- Avoid hardcoding only one workflow.
- Support future extension without rewriting the module.

---

### 4.5 Meetings and Minutes of Meeting

DhanMan includes meeting-related functionality.

Current / planned concepts include:

- Meeting events
- Recurring meetings
- Meeting occurrences
- Agenda items
- Participants
- Action items
- Meeting notes
- Minutes of meeting
- Meeting-based decisions

Important design direction:

- Meeting occurrence data should be separated from recurring event metadata.
- Recurrence should be managed through RRULE-style logic where applicable.
- Meeting decisions should eventually connect to a generic community decision module.

---

### 4.6 Polls and Community Decisions

DhanMan is expected to support a world-class polling and decision system.

The purpose is to convert informal discussions into structured decisions.

Potential sources of decisions:

- Polls
- Meetings
- Committee approvals
- Resident voting
- Formal notices
- Maintenance proposals
- Financial approvals

Important design direction:

- A decision should be generic and not tightly coupled only to meetings.
- Decisions should be discoverable, auditable, and expandable.
- The system should preserve decision context over time.

---

### 4.7 Document Management

DhanMan includes a separate document service.

Important points:

- Service name: `Dhanman.Document`
- Storage: MinIO
- Architecture: simplified clean architecture
- Does not use MediatR
- Uses .NET Core
- Supports Swagger, authentication, authorization, health checks, exception middleware, Prometheus metrics, response compression, CORS, API versioning, and Serilog.
- Apartment documents may be stored in apartment-specific buckets.
- Visitor photos / guard-uploaded photos may be categorized separately, such as under avatars.

---

### 4.8 Resource / Consumables Tracking

DhanMan includes tracking of resources such as water and fuel.

Examples:

- Water tanker entries
- Water meter readings
- Capacity tracking
- Actual received quantity
- Trend reporting

Important known issue / direction:

- UI may capture water meter readings in KL with decimals.
- Backend often stores actual received values in liters.
- Avoid unit confusion.
- Standardize conversion rules.
- Meter readings should support decimals.
- UI should clearly display units.

---

### 4.9 Bank Reconciliation

DhanMan includes bank reconciliation workflows.

Features / design ideas include:

- Bank statement import
- CSV / XLSX parsing
- Vendor / customer prediction
- Payment matching
- Fuzzy matching
- Matched transactions
- Bank-only transactions
- System-only payments
- Reconciliation proposals
- Manual correction flow
- Learning from corrections
- Summary widgets for matched / unmatched records

Important design direction:

- Keep matching explainable.
- Avoid showing misleading party proposals for unmatched transactions.
- Corrections should be useful for future retraining or rule improvement.

---

## 5. High-Level Architecture

DhanMan uses a multi-service architecture.

The system is split into multiple backend services, each having responsibility for a domain area.

Common services include:

| Service | Purpose |
|---|---|
| `dhanman-common` | Shared/common functionality, accounting, organization/company/user/accounting-related features |
| `dhanman-community` / `dhanman-myhome` | Apartment/community management, visitor management, gates, meetings, events |
| `dhanman-sales` | Sales invoices, customer payments, recurring invoices |
| `dhanman-purchase` | Vendor bills, bill payments, purchase workflows |
| `dhanman-inventory` | Warehouses, inventory, product-related features |
| `dhanman-payroll` | Payroll, salary posting, employee payments |
| `dhanman-document` | Document upload/storage service using MinIO |

The services communicate through:

- HTTP APIs
- RabbitMQ / MassTransit for async events and data replication
- Direct service-to-service calls where strict consistency is required
- Database-level functions / stored procedures where needed

---

## 6. DLL Names for Services

When deploying .NET services, the DLL names are:

| Service | DLL |
|---|---|
| `dhanman-common` | `Dhanman.Common.Api.dll` |
| `dhanman-community` | `Dhanman.MyHome.Api.dll` |
| `dhanman-inventory` | `Dhanman.Inventory.Api.dll` |
| `dhanman-payroll` | `Dhanman.Payroll.Api.dll` |
| `dhanman-purchase` | `Dhanman.Purchase.Api.dll` |
| `dhanman-sales` | `Dhanman.Sales.Api.dll` |

---

## 7. Technology Stack

### 7.1 Backend

The backend primarily uses:

- .NET 9
- ASP.NET Core
- EF Core
- Npgsql
- PostgreSQL 17
- CQRS
- MediatR
- Carter for endpoints in some services
- FluentValidation
- Mapster
- Result pattern
- Repository pattern
- Serilog
- Swagger / OpenAPI
- Prometheus metrics
- Health checks
- Response compression
- CORS
- API versioning

Some modules use stored procedures and PostgreSQL functions heavily.

---

### 7.2 Frontend

The web frontend primarily uses:

- React
- TypeScript
- Vite
- Material UI
- Formik
- Yup
- React Intl
- TanStack React Table
- ApexCharts
- React Flow
- FullCalendar

Important frontend preferences:

- Use Material UI.
- Keep screens clean, minimal, and practical.
- Follow Single Responsibility Principle.
- Keep generic components generic.
- Push page-specific logic to page/module level.
- Preserve existing structure and naming conventions unless explicitly asked to refactor.

---

### 7.3 Mobile

The mobile app is being built using:

- Flutter
- Android
- iOS

The mobile app focuses on apartment management features such as:

- Gate management
- Finance
- Ticketing
- Visitor workflows
- Notifications
- Resident interactions

---

### 7.4 Infrastructure

DhanMan is hosted primarily on:

- OVH Cloud

Infrastructure and operations include:

- Ubuntu servers
- NGINX reverse proxy
- systemd services
- Docker where needed
- PostgreSQL
- RabbitMQ
- Redis
- Vault
- Grafana
- Loki
- Promtail
- Prometheus
- MinIO
- Jenkins
- GitHub Actions for some workflows
- Backblaze B2 for server audit backups

---

## 8. Environments

DhanMan typically uses:

- Local
- QA
- Production

Environment-specific configuration is handled through:

- `appsettings.Local.json`
- `appsettings.QA.json`
- `appsettings.Production.json`
- `.env` files
- Jenkins credentials
- Vault secrets

Important expectation:

- Static configuration should remain in templates.
- Secrets should be replaced dynamically using CI/CD or Vault.
- Do not hardcode secrets in source code.
- Avoid committing actual secret values.

---

## 9. Hosting and Deployment

### 9.1 Hosting

Current hosting is on OVH.

The app does not currently depend on Azure or AWS for hosting.

### 9.2 Reverse Proxy

NGINX is used as a reverse proxy.

Domain convention:

- Main domain: `dhanman.com`
- QA services use subdomains like:
  - `qa.common.dhanman.com`
  - `qa.community.dhanman.com`
  - `qa.sales.dhanman.com`
  - `qa.purchase.dhanman.com`
  - `qa.inventory.dhanman.com`
  - `qa.payroll.dhanman.com`

### 9.3 Services

Backend services are generally hosted under:

- `/var/www/qa/<service-name>`
- `/var/www/prod/<service-name>`

Each service is run using systemd.

Example systemd pattern:

```ini
WorkingDirectory=/var/www/qa/dhanman-purchase
ExecStart=/usr/bin/dotnet /var/www/qa/dhanman-purchase/Dhanman.Purchase.Api.dll
EnvironmentFile=/var/www/qa/dhanman-purchase/.env
```

---

## 10. Authentication and Authorization

DhanMan uses:

- Auth0 for authentication
- Custom authorization system for permissions

Important concepts:

- Users are authenticated using Auth0 tokens.
- Application-level permissions are managed by DhanMan.
- `[RequiresPermissions]` attributes are used in backend controllers/endpoints.
- Permission checks are important and should not be bypassed.
- Some service-to-service flows may require machine-to-machine tokens.
- There has been discussion around token handling for scheduled/background tasks.

Important design expectation:

- Authentication and authorization must remain cleanly separated.
- Do not mix Auth0 identity logic with internal permission logic carelessly.
- Machine-to-machine flows should be explicit and secure.
- Avoid storing or reusing tokens in unsafe ways.

---

## 11. Organization, Company, and User Model

The system generally follows this hierarchy:

```text
Organization
  └── Company
        └── Users associated with company
```

Important tables / concepts:

- `OrganizationUser`
- `CompanyUser`
- Users are often associated with companies.
- Permissions are often evaluated at organization level.
- The user mostly uses `CompanyUser` in practice.
- There has been a need to check whether `OrganizationUser` is still actively used by stored procedures, functions, foreign keys, or application logic.

Important expectation:

- Be careful when changing user/organization/company logic.
- These concepts affect authorization, data access, and multi-tenant isolation.

---

## 12. Database

### 12.1 Database Engine

DhanMan uses:

- PostgreSQL 17

### 12.2 Database Design Style

The project uses a combination of:

- EF Core entity mapping
- PostgreSQL functions
- PostgreSQL stored procedures
- SQL reporting functions
- CQRS command/query handlers
- Repository pattern

### 12.3 Multi-Database / Multi-Service Model

Each service may have its own database.

Examples:

- Common DB
- Community DB
- Sales DB
- Purchase DB
- Inventory DB
- Payroll DB
- Document DB or storage metadata DB

Data consistency across services may be handled through:

- RabbitMQ events
- Service-to-service API calls
- Reconciliation jobs
- Manual sync scripts
- Database comparison tools

---

## 13. Backend Coding Pattern

Most backend features follow a clean CQRS-style flow.

Typical flow:

```text
API Endpoint / Controller
  -> Request DTO
  -> Command / Query
  -> Validator
  -> Handler
  -> Repository / DbContext / SQL Function
  -> Response DTO
  -> Result.Match(...)
```

Typical libraries / patterns:

- MediatR
- FluentValidation
- Result pattern
- Mapster
- Carter endpoints in some services
- EF Core
- Repository abstractions
- Domain entities

Expected style:

- Keep commands focused.
- Keep handlers readable.
- Avoid putting too much business logic in controllers.
- Validate early.
- Use domain constants/enums where possible.
- Preserve existing conventions in each service.
- Do not introduce a completely different architecture unless explicitly requested.

---

## 14. Result Pattern

The project commonly uses a Result pattern similar to:

```csharp
Result.Success(request)
    .Ensure(...)
    .Map(...)
    .Bind(...)
    .Match(Ok, BadRequest);
```

Important expectation:

- Prefer existing result/error handling style.
- Use existing error constants such as `Errors.General.BadRequest` where applicable.
- Do not throw exceptions for expected validation failures.
- Use exceptions only for unexpected failures.

---

## 15. API Style

Common API style:

- REST-style endpoints
- `/api/v1/...`
- `apartmentId`, `companyId`, `organizationId`, etc. often appear in routes or request payloads
- Controllers/endpoints construct commands or queries
- Permission attributes are applied at endpoint/controller level

Examples of endpoint areas:

```text
GET    /api/v1/visitor-types
GET    /api/v1/apartments/{apartmentId}/gates
GET    /api/v1/apartments/{apartmentId}/gate-names
POST   /api/v1/visitors/public/generate-qr-token
POST   /api/v1/visitors/public/self-entry
GET    /api/v1/events
GET    /api/v1/apartments/{apartmentId}/resident-requests
```

Important expectation:

- Keep backward compatibility where possible.
- Avoid breaking mobile or frontend flows.
- Add new endpoints when needed instead of overloading unrelated ones.
- Use clear route names.

---

## 16. Frontend Coding Pattern

Frontend uses React + TypeScript + Material UI.

Expected frontend style:

- Functional components
- Hooks
- TypeScript interfaces
- Modular API wrappers
- Formik for forms where already used
- Yup for validation where already used
- Material UI components
- Reusable table/action components
- Avoid deeply mixing page logic into generic components

Important SRP principle:

```text
Generic component:
  Should only render generic UI.

Page/module:
  Should decide permissions, visibility, row-specific conditions, status-specific behavior.
```

Example:

- `BReactTableActionButtons` should remain generic.
- Page-level modules like `InvoiceList` should decide which actions are allowed or visible.
- Overflow menus should use a reusable `ActionMenu`.

---

## 17. Logging

### 17.1 Application Logging

DhanMan uses:

- Serilog
- File-based logs
- Console logs
- Seq in some contexts
- Loki / Promtail / Grafana for centralized logs

### 17.2 Log Paths

Logs are centralized by environment.

Examples:

```text
/var/www/qa/logs/
/var/www/prod/logs/
```

Log file naming pattern may include service names:

```text
dhanman-common-20250411.log
dhanman-purchase-20250411.log
```

### 17.3 Loki / Promtail

Promtail extracts:

- Environment label: QA / PROD
- Service label from filename

Regex pattern example:

```regex
.*/([^/]+)-\d{8}\.log
```

### 17.4 Logging Expectations

When implementing features:

- Add useful logs around critical operations.
- Do not log secrets.
- Do not log passwords, tokens, or sensitive payloads.
- Log IDs, statuses, counts, and failure reasons.
- Keep logs structured and searchable.
- Avoid noisy logs in hot paths unless controlled by switches.

---

## 18. Monitoring

DhanMan uses or plans to use:

- Grafana
- Loki
- Promtail
- Prometheus
- Uptime Kuma
- Health checks
- Service dashboards

Important dashboard needs:

- Service health
- Error count
- Recent error logs
- Environment filter
- Service filter
- SQL / EF query logs where useful
- Redis diagnostics
- RabbitMQ status
- API latency
- Failed jobs
- Database connection pressure

---

## 19. Redis and Caching

Redis is used for caching.

Known design direction:

- Cache should sit between multiple databases and multiple microservices.
- There should be good diagnostics for what is cached and what is not.
- Manual cache invalidation should be possible.
- Cache expiration should be configurable.
- Default cache TTL should exist.
- Custom TTL should be possible per cache area.
- Logging switches should exist for cache hit/miss diagnostics.

Suggested cache configuration pattern:

```json
{
  "Cache": {
    "Default": {
      "LocalExpirationMinutes": 5,
      "DistributedExpirationMinutes": 15
    },
    "Vendors": {
      "LocalExpirationMinutes": 10,
      "DistributedExpirationMinutes": 30
    }
  }
}
```

Expected behavior:

- Use default settings unless custom settings are provided.
- Avoid requiring every API/cache to define its own TTL.
- Add custom cache config only when needed.

Important operational note:

- Redis password mismatches can cause `AuthenticationFailure`.
- Always verify actual Redis `requirepass` and service `.env` connection string match.

---

## 20. RabbitMQ and Messaging

DhanMan uses RabbitMQ with MassTransit for:

- Cross-service communication
- Data replication
- Event publishing
- Event consumption
- Eventually consistent synchronization

Examples:

- Warehouse created in Inventory, then published to Sales/Purchase/Common.
- Customer/vendor/user records may be replicated across relevant services.
- Some operations are tightly coupled and require downstream services to be available.
- RabbitMQ helps preserve packets/messages when a service is temporarily down.

Important expectations:

- Keep message contracts stable.
- Use shared messaging abstractions where already available.
- Avoid duplicating publisher/consumer boilerplate unnecessarily.
- Add idempotency to consumers where possible.
- Add reconciliation jobs for eventual consistency.

---

## 21. Shared RabbitMQ Library

A shared RabbitMQ messaging library exists under:

```text
Messaging/RabbitMQ/
```

Common folders include:

```text
Abstractions
Configurations
DependencyInjection
Extensions
Models
Services
```

Known classes / concepts:

- `CommandPublisher`
- `EventPublisher`
- Envelope models
- Options/configuration classes

Expected direction:

- Standardize and reuse this cross-cutting messaging library across services.
- Avoid each service inventing its own RabbitMQ implementation.

---

## 22. Vault and Secrets

DhanMan uses self-hosted HashiCorp Vault.

Important details:

- Separate QA and PROD Vaults
- AppRole auth enabled
- File storage under `/opt/vault/data`
- Config path: `/etc/vault.hcl`
- UI enabled
- TLS may be disabled for local/internal use depending on setup

Expected secret management:

- Do not commit secrets.
- Use Vault or Jenkins credentials for sensitive values.
- Keep environment-specific secrets separate.
- Maintain backup/export strategy for Vault.
- Rotate secrets carefully.

---

## 23. Jenkins and CI/CD

Jenkins is used for deployment pipelines.

It stores credentials for:

- Database hosts
- Database users
- Database passwords
- SSH keys
- Vault values
- Firebase keys
- Git credentials

Known Git credential:

```text
bharat-mane-git-personal-token
```

Important expectation:

- Do not expose Jenkins credentials in logs.
- Use Jenkins credential bindings.
- Keep deployment scripts repeatable.
- Ensure `.env` files are created/updated correctly during deployment.
- Restart systemd services after deployment.
- Validate service health after deployment.

---

## 24. GitHub Actions

GitHub Actions is used in some workflows.

Example:

- Nightly LightBDD end-to-end tests
- Reports generated and uploaded to OVH-hosted BDD site
- Build output redirected to a consistent `TestReports` folder

---

## 25. Server Audit

A custom server audit script exists:

```text
dr-audit-extended.sh
```

Purpose:

- Collect system state
- Services
- Docker
- NGINX
- SSL
- Promtail
- Grafana
- Other infrastructure details
- Save into timestamped folder
- Upload to Backblaze B2

Purpose of audit:

- Understand current production setup
- Prepare migration to Ansible-managed infrastructure

---

## 26. Ansible Direction

The user has WSL installed and is preparing to use Ansible from a local Windows machine via WSL.

Expected direction:

- Move infrastructure from manual setup to Ansible-managed setup.
- Keep environment reproducible.
- Standardize server configuration.
- Avoid snowflake servers.

---

## 27. Important Domain Concepts

### 27.1 Organization

A top-level tenant-like structure.

### 27.2 Company

A company belongs to an organization.

In apartment context, a company may represent an apartment association or accounting entity.

### 27.3 Apartment

Apartment/community entity where units, residents, gates, visitors, notices, resources, and meetings are managed.

### 27.4 Unit

A flat/unit in an apartment.

Example naming:

```text
A 511
```

### 27.5 User

A person in the system.

Users may be:

- Residents
- Admins
- Facility managers
- Committee members
- Guards
- Employees
- Vendors/service providers

### 27.6 Visitor

A person entering the apartment temporarily.

### 27.7 Gate

Entry/exit point for visitor movement.

Gate may include:

- Latitude
- Longitude
- Allowed radius meters
- Gate type
- Allowed user categories
- In/out usage flags

### 27.8 Approval Configuration

Defines who approves what and at which level.

### 27.9 Decision

A generic governance artifact representing a community decision.

---

## 28. Public Visitor QR Self-Entry Flow

This is an important recurring feature.

### 28.1 Purpose

Allow visitors to scan a QR code and submit their entry without logging in.

### 28.2 Important Security Principle

The visitor is not a registered system user.

Therefore:

- Public endpoint must be secure.
- Token must be signed or validated.
- Token must expire.
- Token usage count must be controlled.
- Do not trust client-sent gate/geofence data.

### 28.3 Token Concepts

Important fields:

```text
TokenHash
Nonce
ExpiresOnUtc
MaxUsageCount
UsedCount
AppliesToAllUnits
AllowedUnitIds
DefaultVisitorTypeId
IsActive
GateId
```

### 28.4 Special AllowedUnitIds Rule

If:

```text
AllowedUnitIds contains 0
AppliesToAllUnits = true
```

Then backend should interpret this as:

```text
All units under the selected apartment
```

Backend should fetch the units server-side.

### 28.5 Public Endpoints

Examples:

```text
POST /api/v1/visitors/public/generate-qr-token
POST /api/v1/visitors/public/self-entry
```

### 28.6 Frontend Paths

Known frontend paths:

```text
src/pages/public/visitorEntry.tsx
src/api/visitor/publicSelfEntry.ts
src/pages/organization/visitorQr/create.tsx
```

### 28.7 Geofencing

Gate-linked geofencing is expected.

Important rule:

- Server must fetch authoritative geofence data using `GateId`.
- Do not trust client-provided latitude/longitude/radius for gate configuration.
- Client may send current location and accuracy.
- Server can enforce or bypass geofence depending on accuracy and config.

Known issue:

- Mobile sometimes sends very poor accuracy such as `100000m`.
- This can trigger errors like `PublicEntry.LowAccuracy`.
- The system should be careful not to break existing flows if location is unavailable.

---

## 29. SignalR Visitor Status Flow

The visitor self-entry flow should support real-time status updates.

Concept:

1. Visitor submits self-entry.
2. System creates pending approval.
3. Resident receives approval request.
4. Resident approves/rejects.
5. Waiting visitor UI receives real-time status update.

Possible hub path:

```text
/hubs/visitor-status
```

Known concepts:

- `PublicVisitorStatusHub`
- `IVisitorStatusBroadcaster`
- `VisitorCreatedPendingApprovalEvent`
- `VisitorApprovedByUnitEvent`

Known implementation risks:

- DI registration missing for broadcaster.
- `IHubContext<PublicVisitorStatusHub>` not registered/resolved.
- Lifetime mismatch in event handlers.
- Background handlers should use proper scoped dependencies.

---

## 30. Notifications and FCM

DhanMan uses Firebase Cloud Messaging for push notifications.

Important flows:

- Resident notification for visitor approval
- Facility manager notification
- Role-based notification
- Unit-resident notification

Known DB/function concepts:

```text
unit_resident_fcm_tokens(p_unit_id integer)
get_fcm_tokens_by_roles(...)
```

Important expectations:

- Ensure tokens are not deleted.
- Ensure user/role mapping is correct.
- Ensure recipient resolution is logged.
- Avoid sending duplicate notifications.

---

## 31. Accounting and Finance Details

DhanMan has strong accounting requirements.

Important accounting entities/concepts:

- Chart of accounts
- Customer accounts
- Vendor accounts
- Employee accounts
- Credit account
- Debit account
- Transaction headers
- Journal entries
- Entry types
- Entry sources
- Invoice headers
- Bill headers
- Payments
- Bank transfers
- Manual journal vouchers

Important design direction:

- Financial posting must be consistent.
- Avoid partial posting.
- Create transaction headers and journal entries together where required.
- Use existing helper patterns from Purchase and Sales modules.
- Refactor Sales logic to match clean Purchase module patterns where appropriate.

---

## 32. Bank Transfer Design

Bank transfer feature uses:

- `bank_transfers`
- `transaction_headers`
- `journal_entries`

Important direction:

- `bank_transfers` should be the source document.
- One bank transfer should create:
  - One `transaction_header`
  - Two `journal_entries`
- There is no separate `bank_transfer_detail` table.
- Use domain entities correctly.
- Be careful with ID types:
  - Some entities use `long`
  - Some use `Guid`
- Avoid domain vs contract model mismatches.

---

## 33. Invoice and Bill Patterns

### 33.1 Sales

Sales module includes:

- Invoice creation
- Invoice payments
- Bulk invoice payments
- Recurring invoices
- Customer account defaults

### 33.2 Purchase

Purchase module includes:

- Bill creation
- Bill payments
- Vendor account defaults
- Purchase approval flow

### 33.3 Refactoring Direction

User prefers Sales payment and invoice handlers to follow the same clean structure as Purchase handlers.

Common helper direction:

```text
TransactionHelper
JournalEntryHelper
```

Expected outcome:

- Less duplicated code
- Clear financial posting logic
- Safer journal entry generation
- Consistent transaction creation

---

## 34. Recurring Invoice Deletion

Known DB function:

```sql
delete_bulk_recurring_schedules(p_schedule_ids uuid[])
```

Expected return shape:

```text
id
invoice_id
status
```

Important implementation issue:

- EF `FromSqlRaw` requires returned column names to match the mapped result class.
- If expected column `id` is missing, EF throws:
  - `The required column 'id' was not present...`

Expected fix:

- Ensure SQL function return columns exactly match EF projection/result model.
- Avoid shape mismatches.

---

## 35. Reporting

DhanMan includes reporting across accounting and appointments/medical-style modules in other projects.

For DhanMan accounting reports:

- Profit and loss
- Balance sheet
- Trial balance
- Customer ledger
- Vendor ledger
- Employee ledger
- Comparative accounts overview
- Income vs expense
- Actual income vs expense

Important expectation:

- Reports should be correct first, optimized second.
- Use SQL functions where suitable.
- Avoid hardcoded financial year assumptions.
- Use company ID and financial year ID correctly.
- Validate with real sample data.

---

## 36. Database Comparison Tooling

User has a schema comparison workflow.

Existing tool:

```text
SchemaComparisonRunner
```

It handles comparison for:

- Tables
- Functions
- Procedures
- Sequences

Expected enhancements:

- Generate actual SQL create scripts for missing:
  - Indexes
  - Primary keys
  - Foreign keys
  - Columns

Important expectation:

- Tool should not only detect mismatches.
- It should generate usable sync scripts.

---

## 37. db-tula / Missing Permission Analyzer

The user wants reusable CLI tools similar to `db-tula`.

One tool idea:

```text
MissingPermissionAnalyzer
```

Purpose:

- Scan ASP.NET Core controllers.
- Detect methods missing `[RequiresPermissions]`.
- Help enforce authorization consistency.

Expected direction:

- Build as reusable CLI.
- Use it across services.
- Support automated checks.

---

## 38. Testing

Testing approaches used:

- xUnit
- EF Core InMemory provider
- Mocked DbSets in older tests
- LightBDD for end-to-end tests
- GitHub Actions nightly test runs

Important testing preference:

- Prefer EF Core InMemory provider over complex mocked DbSet setups where appropriate.
- Extract common mock data into constants/shared test fixtures.
- Keep tests readable and close to real query behavior.

---

## 39. Common Backend Errors and Lessons

### 39.1 NpgsqlConnection ObjectDisposedException

Common cause:

- DbContext lifetime misuse
- Using scoped DbContext in singleton/background service
- Using DbContext after request scope is disposed

Expected fix:

- Keep DbContext scoped.
- Use `IServiceScopeFactory` in background/event handlers.
- Do not capture scoped services in singletons.

---

### 39.2 A Second Operation Started on This Context

Common cause:

- Multiple concurrent EF operations on same DbContext instance.
- Missing `await`.
- Parallel operations using same context.

Expected fix:

- Await each EF operation.
- Use separate scopes/contexts for parallel work.
- Avoid reusing DbContext across threads.

---

### 39.3 Too Many PostgreSQL Clients

Common cause:

- Connection leaks
- Excessive pooling
- Services opening too many DB connections
- Missing disposal
- Too many parallel operations

Expected fix:

- Review connection pooling.
- Dispose connections correctly.
- Optimize DB calls.
- Add monitoring for active connections.

---

### 39.4 Foreign Key Violations

Common examples:

- User ID missing in users table
- Event type ID missing
- Vendor/product relation missing
- Warehouse created_by FK issue

Expected fix:

- Add orphan detection SQL.
- Validate referenced rows before insert.
- Use `NOT VALID` then `VALIDATE CONSTRAINT` carefully when retrofitting FKs.

---

### 39.5 SQL Function Return Shape Mismatch

Common cause:

- PostgreSQL function returns columns that do not match EF result model.

Expected fix:

- Ensure exact column names.
- Ensure exact types.
- Use explicit aliases.
- Test function independently in SQL before using from EF.

---

## 40. Frontend UI / UX Preferences

User prefers:

- Simple
- Minimal
- Clean
- Practical
- Professional
- Not over-engineered
- Material UI-based
- Consistent with existing design

Important UI direction:

- Use clear labels.
- Avoid clutter.
- Keep forms understandable.
- Use summary widgets where helpful.
- Use tables with total rows where needed.
- Use tabs when comparing categories.
- Use steppers/flows for approval where it improves clarity.

---

## 41. Frontend Component Design Rules

Important rules:

1. **Follow SRP**
   - Generic components should remain generic.
   - Business rules should be handled at page/module level.

2. **Do not break existing structure**
   - Preserve naming conventions.
   - Preserve folder structure unless asked.

3. **Action buttons**
   - Use reusable action menu.
   - Decide visible actions at page level.
   - Keep permissions separate from row-status conditions.

4. **Forms**
   - Use Formik and Yup where existing form uses them.
   - Keep validation explicit.
   - Avoid hidden business logic inside UI controls.

5. **Tables**
   - Use TanStack React Table where already used.
   - Add total rows where requested.
   - Keep columns modular.

---

## 42. Known Frontend Files / Areas

Examples from previous work:

```text
src/pages/public/visitorEntry.tsx
src/api/visitor/publicSelfEntry.ts
src/pages/organization/visitorQr/create.tsx
```

Invoice/action menu areas:

```text
BReactTableActionButtons.tsx
commonActionButtons.tsx
invoiceActionButtons.tsx
ActionMenu.tsx
list.tsx
```

Bank import/reconciliation areas:

```text
BankImport
useFileParser
useReconciliation
```

Approval areas:

```text
ApprovalConfiguration
ApprovalFlow
```

---

## 43. Mobile App Direction

DhanMan mobile app should focus on:

- Resident convenience
- Gate workflows
- Visitor approval
- Finance visibility
- Ticketing
- Notifications
- Apartment transparency

Important technical expectations:

- Watch bootstrap performance.
- Track API call count and loading time.
- Consider React Query-like caching patterns for Flutter if useful.
- Avoid slow startup due to unnecessary API calls.

---

## 44. Document Service Direction

The document service should be simpler than other services.

Important rule:

- It does not use MediatR.

Expected structure:

- Controllers/endpoints
- Services
- Storage abstraction
- MinIO implementation
- Metadata persistence
- Auth and authorization
- Health checks
- Metrics
- Swagger

Important storage concept:

- Apartment documents can be grouped by apartment bucket/folder.
- Visitor photos and avatars may be categorized separately.

---

## 45. MinIO

MinIO is used for object storage.

Expected use cases:

- Apartment documents
- Visitor photos
- Uploaded files
- Avatars
- Possibly reports or generated documents

Important expectations:

- Use bucket naming carefully.
- Avoid exposing raw internal bucket paths directly.
- Use presigned URLs where appropriate.
- Enforce authorization before file access.
- Store metadata in DB.

---

## 46. Auth0 Custom Database / Webtask Issues

There have been Auth0 timeout issues such as:

```json
{
  "error": {
    "message": "Request to Webtask exceeded allowed execution time"
  }
}
```

Known context:

- Auth0 custom database login script connects to PostgreSQL.
- Long-running DB connection/query can cause timeout.
- Need optimized DB connection and query.
- Avoid excessive logging or slow operations in Auth0 script.
- Use proper indexing on email/user_name.
- Ensure DB host is reachable quickly.

Important expectation:

- Auth0 login scripts must be extremely fast.
- Avoid heavy logic inside Auth0 Webtask.
- Use connection timeout safeguards.

---

## 47. Redis Operational Issue Example

Known Redis problem:

- Application `.env` had one password.
- Actual Redis server accepted another password.

Example actual successful test:

```bash
redis-cli -a 'B@dhi$1234' ping
# PONG
```

Important lesson:

- Always verify actual Redis server password using `redis-cli`.
- Match service `.env` connection string.
- Restart service after `.env` change.
- Check process environment using `/proc/<pid>/environ`.

---

## 48. Git and Repository Preferences

User often asks for:

- PR name
- Commit message
- Copilot prompt
- Codex prompt
- One-paragraph implementation prompt
- Git-style naming

Expected commit message style:

```text
fix(scope): concise description
feat(scope): concise description
refactor(scope): concise description
chore(scope): concise description
```

Example PR naming style:

```text
Fix Redis authentication configuration for QA purchase service
Refactor invoice payment posting to use transaction helpers
Add visitor QR gate-linked geofence validation
```

---

## 49. Prompting Style for Other Agents

User often wants prompts to give to:

- Copilot
- Codex
- Claude
- Other coding agents

Preferred prompt style:

- Usually one paragraph
- Direct
- Include exact files/paths if known
- Include constraints
- Mention “without breaking existing functionality”
- Mention current architecture/patterns
- Mention expected output
- Avoid vague instructions

Good prompt template:

```text
Analyze the existing implementation in [module/files] and fix [issue] without breaking existing functionality. Preserve the current architecture, naming conventions, folder structure, CQRS/MediatR/Result pattern, and existing API contracts. First identify the root cause, then apply the smallest safe production-ready change. Add meaningful logs and validation where needed, avoid hardcoding, and ensure the fix works for QA and production environments. Update or add tests if the project has nearby test coverage.
```

---

## 50. Coding Agent Instructions for DhanMan

Any AI/coding agent working on DhanMan should follow these rules.

### 50.1 Do Not Break Existing Functionality

This is the most important rule.

Before changing code:

- Understand current flow.
- Check existing callers.
- Preserve API contracts.
- Avoid unnecessary refactors.
- Keep backward compatibility.
- Add new logic in a safe way.

---

### 50.2 Follow Existing Architecture

Do not introduce a different architecture unless explicitly requested.

Backend should generally follow:

```text
Controller / Endpoint
Command / Query
Validator
Handler
Repository / DbContext
Response DTO
Result pattern
```

Frontend should generally follow:

```text
Page
Hooks
API wrapper
Reusable components
Material UI
TypeScript interfaces
```

---

### 50.3 Keep Generic Components Generic

Do not put page-specific logic into shared components.

Example:

Bad:

```text
BReactTableActionButtons knows invoice status and invoice permissions internally.
```

Good:

```text
InvoiceList prepares visible actions and passes them to generic action button/menu component.
```

---

### 50.4 Prefer Small, Safe Changes

When fixing bugs:

- Do root cause analysis.
- Make the smallest safe change.
- Add logs if useful.
- Do not rewrite unrelated parts.
- Do not rename files unnecessarily.
- Do not change public contracts unless required.

---

### 50.5 Use Production-Ready Code

Production-ready means:

- Handles nulls safely
- Validates inputs
- Logs meaningful context
- Avoids secrets in logs
- Uses cancellation tokens where appropriate
- Avoids race conditions
- Handles concurrency
- Does not hardcode environment-specific values
- Works in QA and production

---

### 50.6 Respect Multi-Tenant Boundaries

Always consider:

- OrganizationId
- CompanyId
- ApartmentId
- UserId
- Role/permission scope

Do not fetch or modify data without tenant scoping.

---

### 50.7 Be Careful With Financial Logic

Financial changes require extra care.

Before changing:

- Understand debit/credit behavior.
- Understand transaction header generation.
- Understand journal entries.
- Understand source document.
- Avoid partial posting.
- Use existing helper patterns.

---

### 50.8 Be Careful With Public Endpoints

Public endpoints like visitor self-entry do not have normal user authentication.

Therefore:

- Validate signed tokens.
- Check expiration.
- Check usage count.
- Avoid trusting client-provided IDs blindly.
- Do not expose sensitive internal data.
- Add rate limiting if needed.

---

## 51. Important Security Expectations

Security rules:

- Never commit secrets.
- Never log passwords.
- Never log tokens.
- Validate all public inputs.
- Enforce authorization.
- Avoid exposing internal IDs unnecessarily.
- Use HTTPS.
- Use proper CORS configuration.
- Use secure storage for files.
- Use Vault/Jenkins credentials for secrets.
- Use least privilege for DB users.
- Be careful with public QR/token flows.

---

## 52. Performance Expectations

Important performance areas:

- Startup time
- API latency
- DB query performance
- Redis caching
- Flutter bootstrap time
- Frontend bundle size
- Large dropdowns/autocompletes
- Bank import processing
- Reporting queries
- EF Core query optimization

Preferred optimization approach:

1. Measure first.
2. Identify bottleneck.
3. Optimize only relevant path.
4. Avoid premature over-engineering.
5. Preserve readability.

---

## 53. Database Performance Preferences

User prefers optimized SQL where:

- Joins are applied only when needed.
- Filters are applied only when relevant parameters are passed.
- Avoid `FIND_IN_SET`.
- Avoid unnecessary null checks.
- Use indexes where needed.
- Use summary/reporting tables where appropriate.
- Avoid dynamic SQL unless justified.

---

## 54. Reporting Query Pattern

For reporting modules, the user has used optimized summary tables such as:

```text
appointmentbillingsummaryrpt
```

Expected direction:

- Use consolidated summary tables when available.
- Avoid repeating expensive joins.
- Keep report filters dynamic but efficient.

---

## 55. Known User Preferences

The user prefers:

- Direct answers
- Blunt mentor-like tone
- Best practices
- Industry-standard implementation
- Detailed reasoning when needed
- Ready-to-copy code
- Exact commands
- Stepwise debugging
- Practical solutions
- Minimal unnecessary theory
- Prompts that can be directly given to Copilot/Codex/Claude

The user dislikes:

- Vague answers
- Overengineering without need
- Breaking existing structure
- Unnecessary follow-up questions
- Generic advice without actionable steps
- Changing naming conventions without reason
- Hidden assumptions

---

## 56. Example Dummy Apartment

A dummy apartment used for setup/testing:

```text
Tornites
Location: Bangalore
```

Example dummy company/apartment naming:

```text
Tornites Apartment Association
Tornites 1
```

---

## 57. Common QA Server Context

QA server is typically referred to as:

```text
dhanman-qa
```

Known QA public IP from previous diagnostics:

```text
54.37.159.71
```

Common QA paths:

```text
/var/www/qa/dhanman-common
/var/www/qa/dhanman-community
/var/www/qa/dhanman-purchase
/var/www/qa/dhanman-sales
/var/www/qa/dhanman-inventory
/var/www/qa/dhanman-payroll
/var/www/qa/logs
```

Common service check commands:

```bash
sudo systemctl status dhanman-purchase-qa --no-pager
sudo journalctl -u dhanman-purchase-qa.service -n 100 --no-pager
systemctl cat dhanman-purchase-qa
```

---

## 58. Useful Linux Diagnostic Commands

### Check service environment

```bash
sudo cat /proc/$(systemctl show -p MainPID dhanman-purchase-qa | cut -d= -f2)/environ | tr '\0' '\n' | grep -i redis
```

### Check systemd service config

```bash
systemctl cat dhanman-purchase-qa
```

### Check service status

```bash
sudo systemctl status dhanman-purchase-qa --no-pager
```

### Check logs

```bash
sudo journalctl -u dhanman-purchase-qa.service -n 100 --no-pager
```

### Test Redis password

```bash
redis-cli -a '<password>' ping
```

---

## 59. Agent Checklist Before Making Backend Changes

Before changing backend code, an agent should check:

- Which service owns the feature?
- Which database owns the data?
- Is this command/query already implemented elsewhere?
- Is there an existing helper pattern?
- Is this endpoint protected by permissions?
- Does this require organization/company/apartment scoping?
- Does this affect financial posting?
- Does this publish RabbitMQ events?
- Does this require cache invalidation?
- Does this require frontend/mobile contract changes?
- Are there existing tests?
- Are there stored procedures/functions involved?
- Is there a migration needed?

---

## 60. Agent Checklist Before Making Frontend Changes

Before changing frontend code, an agent should check:

- Which page owns the behavior?
- Is there an existing API wrapper?
- Is there a reusable component already?
- Is this a generic component or page-specific logic?
- Are permissions handled at page level?
- Are row-status conditions handled separately?
- Does this affect mobile/responsive layout?
- Does this use Formik/Yup?
- Does this use Material UI?
- Are there existing TypeScript interfaces?
- Does the change preserve current UX?

---

## 61. Agent Checklist Before Making Infra Changes

Before changing infra, an agent should check:

- QA or production?
- Which service?
- Which systemd file?
- Which `.env` file?
- Which NGINX config?
- Which Vault path?
- Which Jenkins credentials?
- Is Redis/RabbitMQ/PostgreSQL involved?
- Does the service need restart?
- Are logs visible after restart?
- Is health check passing?
- Is SSL affected?
- Are secrets exposed?

---

## 62. Suggested Agent Operating Mode

For any DhanMan task, an agent should work like this:

1. Read relevant files first.
2. Understand current implementation.
3. Identify root cause or required change.
4. Propose minimal safe approach.
5. Implement without breaking existing behavior.
6. Add logs/tests only where useful.
7. Validate build.
8. Provide summary:
   - Files changed
   - What changed
   - Why changed
   - How to test
   - Risks/notes

---

## 63. Common “Do Not Do” List

Do not:

- Hardcode secrets.
- Change public API response shape casually.
- Rename existing files without reason.
- Move business logic into generic UI components.
- Add direct DB calls in controllers.
- Skip permission checks.
- Trust public request payloads blindly.
- Mix tenant data.
- Log sensitive values.
- Overengineer small fixes.
- Rewrite working modules unnecessarily.
- Ignore existing patterns.
- Add new libraries without need.
- Break QA/prod environment config.
- Assume one service has all data.
- Use the same DbContext concurrently.
- Capture scoped services in singleton/background workers.

---

## 64. Ideal Definition of Done

A DhanMan change is done when:

- Code builds.
- Existing functionality is preserved.
- New behavior works in local/QA.
- API contracts are respected.
- Logs are meaningful.
- Secrets are safe.
- Permissions are correct.
- Multi-tenant scoping is correct.
- Database changes are documented.
- Frontend and backend contracts match.
- Tests are added or updated where practical.
- Deployment impact is clear.

---

## 65. One-Paragraph Context for Quick Agent Use

DhanMan is a multi-service apartment and financial management platform built under B2A Technologies Private Limited, hosted on OVH, using .NET 9, PostgreSQL 17, React TypeScript with Material UI, Flutter mobile, RabbitMQ/MassTransit, Redis, Vault, Grafana/Loki/Promtail, Jenkins, and MinIO. It includes services such as Common, Community/MyHome, Sales, Purchase, Inventory, Payroll, and Document. The product manages apartment finance, accounting, invoices, bills, payments, bank reconciliation, visitors, gates, QR self-entry, approvals, meetings, notices, documents, resources, and governance decisions. Backend work should follow the existing CQRS/MediatR/Result pattern with validators, handlers, repositories/DbContext, PostgreSQL functions where used, tenant scoping by organization/company/apartment, and strict permission checks. Frontend work should preserve existing React TypeScript and Material UI structure, follow SRP, keep generic components generic, and place permission/status/page-specific logic at page level. All changes must be production-ready, backward compatible, secure, and should not break existing QA or production behavior.

---

## 66. Final Guidance for Any Agent

When working on DhanMan, assume the system is already live and used by real users.

Therefore:

- Prefer safe incremental changes.
- Preserve existing functionality.
- Respect architecture.
- Respect tenant boundaries.
- Respect financial correctness.
- Respect security.
- Add observability where needed.
- Keep the implementation clean and maintainable.

The best output is not just code that works once.  
The best output is code that fits the existing DhanMan ecosystem and can be safely maintained over time.
