# Domain 10: Administration & Security

## Purpose
Manage system users, access control, password management, data backup/restore, and audit mechanisms. Ensures data integrity and controlled access across a multi-user, multi-company deployment.

---

## Key Features

### 10.1 User Management (`frmMastUser`, `frmAdmUser`)
- User account creation with username and password
- Admin flag: Admin users bypass all menu restrictions
- Account active / inactive status
- Default password enforcement on first login
- User listed in setup DB for multi-firm access

### 10.2 User Privileges (`frmAdmUserPrevilage`, `frmUserPrevilage`)
- Menu-level access control per user
- Permissions per form:
  - **Add:** Create new records
  - **Modify:** Edit existing records
  - **Delete:** Remove records
  - **View:** Read-only access
- Privilege matrix: User × Form × Permission type
- Admin user: Global bypass of all restrictions
- Role assignment not explicit — per-form control

### 10.3 Password Management (`frmUtlChangePassword`, `frmUtlCheckPassword`)
- Self-service password change
- Password validation before high-risk operations (delete, audit modify)
- Minimum/maximum length enforced (10–15 characters)
- Password stored in database (legacy — no encryption described)

### 10.4 Data Backup (`frmUtlBackupData`)
- Full database backup to external drive / network path
- Manual trigger (no scheduling)
- Backup of MDB files per company
- Status confirmation after backup

### 10.5 Data Restore (`frmUtlRestoreData`)
- Restore from backup files
- Overwrites current data (destructive — confirmation required)
- Supports point-in-time restore via backup files

### 10.6 Audit & Transaction Protection
- **Audit lock (`VIsAudited`):** Once a transaction is marked audited, no modification or deletion is possible without password
- **Check Pass (SMS/WhatsApp alerts):** Admin/owner notified in real-time when any transaction is modified or deleted
- **User tracking:** `vUserName` field on every transaction records the creating user
- **Fiscal year lock:** Cannot post to a closed/locked fiscal year

### 10.7 Utility Functions
| Utility | Description |
|---|---|
| `frmUtlAccountTransfer` | Transfer account balances between firms/years |
| `frmUtlAddAccounts` | Bulk account creation utility |
| `frmUtlCreationToFirm` | Copy master setup from one company to another |
| `frmUtlLastYrOutstanding` | Carry forward sales outstanding from previous year |
| `frmUtlLastYrOutstandingPurch` | Carry forward purchase outstanding from previous year |
| `frmUtlSelectYear` | Switch between fiscal years |
| `frmSelCompany` | Switch between companies (multi-company) |
| `frmSelectYear` | Alternate year selection form |
| `frmSecurity` | Security settings management |

### 10.8 Multi-Company / Multi-Firm Support
- Single HITRIX installation serves multiple trading firms
- Each firm has its own MDB database file
- Central Setup DB: Cross-firm user table
- Company selection at login (`frmSelCompany`)
- Year selection after company selection (`frmSelectYear`)

---

## Key Data Captured

| Field | Description |
|---|---|
| `UserName`, `Password` | Login credentials |
| `IsAdminUser` | Admin privilege flag |
| `IsActive` | Account active/inactive |
| `MenuName`, `MenuIndex` | Form/menu identifier |
| `Add`, `Modify`, `Delete`, `View` | Permission flags (1 = allowed, 0 = denied) |
| `vUserName` | User who created the transaction (on all records) |
| `VIsAudited` | Audit-locked flag (on all transaction records) |
| `CompCode`, `CompYear` | Company and year identifiers (multi-company scope) |

---

## Business Rules & Logic

- Admin user (`gIsAdminUser = True`): Bypasses all menu privilege checks at runtime
- Regular user: Permissions checked against `tblUserPrevilage` table on every menu action
- Privilege scope: Currently company-wide (not company-specific per user)
- Audit lock: Once set, `VIsAudited = True` prevents modification/deletion without unlock password
- Check pass: Any delete or modify on a protected transaction triggers real-time SMS/WhatsApp to admin
- Password: 10–15 character limit; stored in DB (legacy plain text or simple encoding)
- Year-end carryforward: Utilities carry outstanding balances to new fiscal year
- Company copy: `frmUtlCreationToFirm` copies masters (accounts, narrations, items) from source to new firm

---

## Integration Points

| Integrates With | How |
|---|---|
| All Modules | User privileges checked on every menu action |
| All Transactions | `vUserName` captured on every record |
| Finance | Audit lock prevents modification post-audit |
| Communication | Check Pass SMS/WhatsApp alerts sent on sensitive operations |
| Data Management | Backup/restore protects all company MDB files |
| Multi-Company | Company and year selection at login drives which DB is active |
