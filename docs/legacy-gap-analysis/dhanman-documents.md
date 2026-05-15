# dhanman-documents Legacy Comparison

## Project Summary

`dhanman-documents` is a focused document service centered on upload, download, access control, cleanup, and MinIO-backed storage.

## Current Dhanman Scope

- API controller in `src/Dhanman.Document/Controllers`
- service interfaces and implementations in `src/Dhanman.Document/Interfaces` and `src/Dhanman.Document/Services`
- models and configuration in `src/Dhanman.Document/Models` and `src/Dhanman.Document/Configuration`

## Legacy VB Scope Placeholder

Legacy VB scope is not available in this workspace.

Expected legacy evidence to capture later:

- document upload forms
- file storage locations
- access control rules
- attachment download or preview flows
- cleanup or archival routines

## Feature Comparison Matrix

| Area | Feature / Object | Legacy VB | Dhanman | Gap | Priority | Evidence | Notes |
| --- | --- | --- | --- | --- | --- | --- | --- |
| Documents | Upload and download APIs | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `Controllers/DocumentController.cs`, `Models/UploadRequest.cs`, `Models/DownloadRequest.cs` | Primary document-service capability. |
| Security | Access control and auth integration | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `Services/AccessControlService.cs`, `Extensions/Auth0Extensions.cs` | Important for secure document handling. |
| Storage | MinIO-backed document storage | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `Services/MinioDocumentService.cs`, `Services/MinioBucketManager.cs`, `Configuration/MinioSettings.cs` | Storage implementation is explicit. |
| Operations | Cleanup background processing | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `Medium` | `Services/CleanupBackgroundService.cs` | Useful parity item for retention/archive behavior. |
| Models | Batch download URL flows | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `Medium` | `Models/DownloadUrlBatchRequest.cs`, `Models/DownloadUrlsRequest.cs`, `Models/DownloadUrlResponse.cs` | Suggests signed URL or batch access support. |

## Tables And Entities Matrix

| Area | Feature / Object | Legacy VB | Dhanman | Gap | Priority | Evidence | Notes |
| --- | --- | --- | --- | --- | --- | --- | --- |
| Request models | Upload and download models | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `Models/UploadRequest.cs`, `Models/DownloadRequest.cs`, `Models/UploadUrlResponse.cs`, `Models/DownloadUrlResponse.cs` | Model layer is visible even without DB entities. |
| Domain models | Apartment and visitor document types | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `Medium` | `Models/ApartmentDocument.cs`, `Models/VisitorAvatar.cs` | Indicates document types for different business contexts. |
| Service abstractions | Document service interface | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `Medium` | `Interfaces/IDocumentService.cs` | Good comparison point if VB had shared attachment utilities. |
| Storage objects | Physical document tables/scripts | To be verified from VB form/report/table | `Missing` | Cannot assess until legacy evidence is added | `Medium` | `dhanman-documents/src/Dhanman.Document` | No persistence table scripts or ORM entity tables were visible in this document service during the current repo scan. |

## All Tables

No SQL table script inventory was visible for `dhanman-documents` during this repo scan.

Visible storage/service objects are code-first/service-first items such as:

- `Controllers/DocumentController.cs`
- `Interfaces/IDocumentService.cs`
- `Services/AccessControlService.cs`
- `Services/CleanupBackgroundService.cs`
- `Services/MinioBucketManager.cs`
- `Services/MinioDocumentService.cs`
- `Models/ApartmentDocument.cs`
- `Models/VisitorAvatar.cs`
- `Models/UploadRequest.cs`
- `Models/DownloadRequest.cs`

## APIs / Controllers / Services Matrix

| Area | Feature / Object | Legacy VB | Dhanman | Gap | Priority | Evidence | Notes |
| --- | --- | --- | --- | --- | --- | --- | --- |
| APIs | Document controller | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `Controllers/DocumentController.cs` | Main entry point for document operations. |
| Services | MinIO storage service | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `Services/MinioDocumentService.cs`, `Services/MinioBucketManager.cs` | Compare later with legacy file-share or DB-blob model. |
| Services | Access control and cleanup services | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `Medium` | `Services/AccessControlService.cs`, `Services/CleanupBackgroundService.cs` | Important for governance and retention. |

## Screens / Forms / Pages Matrix

| Area | Feature / Object | Legacy VB | Dhanman | Gap | Priority | Evidence | Notes |
| --- | --- | --- | --- | --- | --- | --- | --- |
| Business workflow | Upload flow | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `DocumentController.cs`, `UploadRequest.cs` | Compare with legacy attachment capture flow later. |
| Business workflow | Download and access validation | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `DocumentController.cs`, `AccessControlService.cs`, `DownloadRequest.cs` | Important parity area for secure downloads. |
| Business workflow | Background cleanup and expiry | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `Medium` | `CleanupBackgroundService.cs` | Track legacy archival/cleanup rules later. |

## Reports / Export / Integration Matrix

| Area | Feature / Object | Legacy VB | Dhanman | Gap | Priority | Evidence | Notes |
| --- | --- | --- | --- | --- | --- | --- | --- |
| Integration | External object storage integration | To be verified from VB form/report/table | `Present` | Cannot assess until legacy evidence is added | `High` | `MinioDocumentService.cs`, `VaultMinioConfigurationExtensions.cs` | Important difference if legacy stored documents locally or in DB. |
| Reporting | Document audit or activity reporting | To be verified from VB form/report/table | `Unknown` | Cannot assess until legacy evidence is added | `Medium` | `src/Dhanman.Document` | No explicit reporting surface was obvious from the current evidence pass. |

## Gaps And Recommendations

- Record whether VB stored documents in file shares, database blobs, or third-party systems.
- If the legacy document service used physical document tables, add them explicitly because the current Dhanman service appears storage-service oriented rather than table-centric.
- Compare access rules closely; document permissions often differ more than upload screens do.
- Note retention, cleanup, and URL-expiry behavior from legacy because those policies are easy to miss.
- If the VB system generated attachment reports, add them in the reporting matrix once known.

## Evidence From Codebase

- `dhanman-documents/src/Dhanman.Document/Controllers/DocumentController.cs`
- `dhanman-documents/src/Dhanman.Document/Interfaces/IDocumentService.cs`
- `dhanman-documents/src/Dhanman.Document/Services`
- `dhanman-documents/src/Dhanman.Document/Models`

## Open Legacy Questions

- How were documents uploaded, stored, and downloaded in the VB system?
- Did the legacy app enforce role-based access or only module-level access?
- Were old files archived, versioned, or cleaned automatically?
