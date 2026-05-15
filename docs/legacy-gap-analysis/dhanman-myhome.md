# dhanman-myhome Legacy Comparison

## Project Summary

`dhanman-myhome` is the community and resident management backend service. It includes apartments, buildings, residents, visitors, tickets, notices, events, documents, service providers, gates, passes, and related operational workflows.

## Current Dhanman Scope

- API controllers in `src/Dhanman.Community.Api/Controllers`
- domain entities in `src/Core/Dhanman.Community.Domain/Entities`
- persistence and infrastructure under community and myhome persistence projects

## Legacy VB Scope Placeholder

Legacy VB scope is not available in this workspace.

Expected legacy evidence to capture later:

- apartment and resident masters
- visitor and gate pass workflows
- complaint and ticket tracking
- notices, calendar, and community documents
- service provider and security operations

## Feature Comparison Matrix

| Area | Feature / Object | Legacy VB | Dhanman | Gap | Priority | Evidence | Notes |
| --- | --- | --- | --- | --- | --- | --- | --- |
| Community | Apartments, buildings, floors, units | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `ApartmentsController.cs`, `BuildingsController.cs`, `FloorsController.cs`, `UnitsController.cs` | Core property structure is present. |
| Residents | Residents, categories, directory, public context | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `ResidentsController.cs`, `CategoriesController.cs`, `CommunityDirectoryController.cs`, `PublicController.cs` | Resident-facing scope is broad. |
| Visitors and access | Visitors, gate passes, gates, security | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `VisitorsController.cs`, `GatePassesController.cs`, `GatesController.cs`, `SecurityGuardController.cs` | Strong visitor and access-control footprint. |
| Operations | Complaints, tickets, support teams | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `ComplaintsController.cs`, `TicketsController.cs`, `SupportTeamsController.cs` | Likely a major parity area if VB had facility management. |
| Community life | Notices, events, calendar, polls | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `Medium` | `NoticesController.cs`, `EventsController.cs`, `CommunityCalenderController.cs`, `PollController.cs` | Strong community feature set. |
| Service operations | Service providers, delivery companies, resources | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `Medium` | `ServiceProvidersController.cs`, `ServiceProviderCompanyController.cs`, `DeliveryCompaniesController.cs`, `ResourcesController.cs` | Important for operations comparison. |

## Tables And Entities Matrix

| Area | Feature / Object | Legacy VB | Dhanman | Gap | Priority | Evidence | Notes |
| --- | --- | --- | --- | --- | --- | --- | --- |
| Property model | Apartments, buildings, floors, units | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `Entities/Apartments`, `Entities/Buildings`, `Entities/Floors`, `Entities/Units` | Main community structure model. |
| Resident model | Residents, resident units, resident requests | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `Entities/Residents`, `Entities/ResidentUnits`, `Entities/ResidentRequests` | Good baseline for member management parity. |
| Visitor model | Visitors, approvals, logs, vehicles | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `Entities/Visitors`, `Entities/VisitorApprovals`, `Entities/VisitorLogs`, `Entities/VisitorVehicles` | Strong visitor lifecycle evidence. |
| Support model | Tickets, statuses, workflows, comments | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `Entities/Tickets`, `Entities/TicketStatuses`, `Entities/TicketWorkflows`, `Entities/TicketComments` | Important for service desk parity. |
| Community content | Notices, documents, polls, events | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `Medium` | `Entities/Notices`, `Entities/CommunityDocuments`, `Entities/Polls`, `Entities/Events` | Shows broad community engagement scope. |
| Physical tables | Community DB table scripts | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `apartments.sql`, `residents.sql`, `visitors.sql`, `tickets.sql`, `notices.sql`, `service_providers.sql`, `gates.sql` | Concrete SQL scripts exist for the main community operational areas. |

## All Tables

All SQL table/object scripts currently visible under `dhanman-myhome/src/Dhanman.MyHome.Persistence/DBObjects/Table`:

- `__EFMigrationsHistory.sql`
- `addresses.sql`
- `apartment_requests.sql`
- `apartment_types.sql`
- `apartments.sql`
- `banks.sql`
- `building_types.sql`
- `buildings.sql`
- `calendars.sql`
- `categories.sql`
- `cities.sql`
- `committee_members.sql`
- `community_calenders.sql`
- `companies.sql`
- `complaint.sql`
- `countries.sql`
- `delivery_companies.sql`
- `delivery_company_categories.sql`
- `delivery_personnel.sql`
- `emergency_contacts.sql`
- `event_occurrence_status.sql`
- `event_occurrences.sql`
- `event_participants.sql`
- `event_status.sql`
- `event_types.sql`
- `events.sql`
- `facility_bookings  .sql`
- `floors.sql`
- `gate_types.sql`
- `gates.sql`
- `guest_approvals.sql`
- `has_default.sql`
- `identity_types.sql`
- `meeting_action_items.sql`
- `meeting_agenda_items.sql`
- `meeting_notes.sql`
- `meeting_participants.sql`
- `member_additional_details.sql`
- `multi_unit_visits.sql`
- `notice_categories.sql`
- `notice_priorities.sql`
- `notices.sql`
- `occupancy_types.sql`
- `occupant_types.sql`
- `organizations.sql`
- `pins.sql`
- `portfolios.sql`
- `possible_visitors.sql`
- `pre_approved_entries.sql`
- `pre_approved_schedule_rules.sql`
- `push_notifications.sql`
- `raise_tickets.sql`
- `resident_request_statuses.sql`
- `resident_requests.sql`
- `resident_tokens.sql`
- `resident_types.sql`
- `resident_units.sql`
- `residents.sql`
- `roles.sql`
- `schema_versions.sql`
- `service_provider_addresses.sql`
- `service_provider_apartments.sql`
- `service_provider_logs.sql`
- `service_provider_logs_2022_2023.sql`
- `service_provider_logs_2023_2024.sql`
- `service_provider_logs_2024_2025.sql`
- `service_provider_logs_2025_2026.sql`
- `service_provider_sub_types.sql`
- `service_provider_ticket_categories.sql`
- `service_provider_types.sql`
- `service_provider_verifications.sql`
- `service_providers.sql`
- `states.sql`
- `sub_categories.sql`
- `ticket_categories.sql`
- `ticket_comments.sql`
- `ticket_fors.sql`
- `ticket_logs.sql`
- `ticket_number_sequences.sql`
- `ticket_priorities.sql`
- `ticket_service_provider_otps.sql`
- `ticket_statuses.sql`
- `ticket_workflow.sql`
- `tickets.sql`
- `unit_service_providers.sql`
- `unit_statuses.sql`
- `unit_types.sql`
- `unit_vehicle_limits.sql`
- `units.sql`
- `user_device_tokens.sql`
- `user_fcm_tokens.sql`
- `user_roles.sql`
- `users.sql`
- `UserToPaidModulesMapping.sql`
- `UserToRoleMapping.sql`
- `vehicle_types.sql`
- `vehicles.sql`
- `verification_types.sql`
- `visit_types.sql`
- `visitor_apartments.sql`
- `visitor_approvals.sql`
- `visitor_delivery_company.sql`
- `visitor_details.sql`
- `visitor_emails.sql`
- `visitor_identities.sql`
- `visitor_logs.sql`
- `visitor_phones.sql`
- `visitor_statuses.sql`
- `visitor_types.sql`
- `visitor_vehicles.sql`
- `visitors.sql`
- `water_tanker_deliveries.sql`

## APIs / Controllers / Services Matrix

| Area | Feature / Object | Legacy VB | Dhanman | Gap | Priority | Evidence | Notes |
| --- | --- | --- | --- | --- | --- | --- | --- |
| APIs | Property and resident APIs | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `ApartmentsController.cs`, `BuildingsController.cs`, `ResidentsController.cs`, `UnitsController.cs` | Core master and occupancy APIs. |
| APIs | Visitor and access-control APIs | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `VisitorsController.cs`, `VisitorsPublicController.cs`, `GatePassesController.cs`, `GatesController.cs` | Strong entry-management footprint. |
| APIs | Complaint and service APIs | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `ComplaintsController.cs`, `TicketsController.cs`, `ServiceProvidersController.cs`, `SupportTeamsController.cs` | Important for FM-style comparisons. |
| APIs | Community content APIs | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `Medium` | `NoticesController.cs`, `EventsController.cs`, `CommunityDocumentController.cs`, `CommunityCalenderController.cs` | Good candidate for later report and workflow review. |

## Screens / Forms / Pages Matrix

| Area | Feature / Object | Legacy VB | Dhanman | Gap | Priority | Evidence | Notes |
| --- | --- | --- | --- | --- | --- | --- | --- |
| Business workflow | Resident and unit management | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `ResidentsController.cs`, `UnitsController.cs`, `ApartmentsController.cs` | Start property/resident parity here. |
| Business workflow | Visitor approval and gate pass handling | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `VisitorsController.cs`, `GatePassesController.cs`, `SecurityGuardController.cs` | Likely one of the most operationally visible flows. |
| Business workflow | Complaint and support resolution | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `ComplaintsController.cs`, `TicketsController.cs`, `DecisionController.cs` | Capture SLA and escalation differences later. |

## Reports / Export / Integration Matrix

| Area | Feature / Object | Legacy VB | Dhanman | Gap | Priority | Evidence | Notes |
| --- | --- | --- | --- | --- | --- | --- | --- |
| Reports | Visitor, gate pass, and security reporting | To be verified from VB form/report/table | `Partial` | Cannot assess until legacy evidence is added | `High` | `VisitorLogs`, `GatePassesController.cs`, `SecurityGuardController.cs` | Exact report endpoints are not yet obvious from the current surface evidence. |
| Reports | Complaints and ticket reporting | To be verified from VB form/report/table | `Partial` | Cannot assess until legacy evidence is added | `High` | `TicketsController.cs`, `Entities/TicketSlaTrackings`, `Entities/TicketEscalationLogs` | Important for service-quality parity. |
| Integration | Community documents and public flows | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `Medium` | `CommunityDocumentController.cs`, `PublicController.cs`, `VisitorsPublicController.cs` | Public/resident-facing operations may differ from legacy significantly. |

## Gaps And Recommendations

- Break the legacy review into property, resident, visitor, and support domains to keep this large service manageable.
- Focus first on visitor/gate pass and complaints/tickets if those are business-critical operational areas.
- Record any kiosk, guard, QR, or public approval flows from VB because they may not map one-to-one.
- Add legacy report names for security logs, resident statements, complaints, and service-provider operations as soon as they are known.

## Evidence From Codebase

- `dhanman-myhome/src/Dhanman.Community.Api/Controllers`
- `dhanman-myhome/src/Core/Dhanman.Community.Domain/Entities`
- `dhanman-myhome/src/Dhanman.MyHome.Persistence/DBObjects/Table`

## Open Legacy Questions

- Which VB forms covered apartments, residents, and visitors?
- Did VB support gate passes, approvals, or public/self-service links?
- Which complaints, ticket, and resident communication reports existed?
