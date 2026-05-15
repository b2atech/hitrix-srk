# Domain 1: Masters & Configuration

## Purpose
Maintain the foundational data required for all business transactions — company setup, accounts/ledgers, product classifications, user management, and system settings. All transaction modules depend on data defined here.

---

## Key Features

### 1.1 Company Master (`frmMastCompany`)
- Multi-company support within one installation
- Fiscal year management (Indian standard: April 1 – March 31)
- GST/TAN/PAN/CST/BST registration numbers
- Bank details with RTGS/IFSC codes
- TDS applicability flag
- Address, phone, email for printed documents

### 1.2 Account Master / Chart of Accounts (`frmMastAccount`)
- Debit/Credit opening balances at fiscal year start
- Interest percentage setup (for inter-company accounts)
- Broker commission linkage (`AcBrkCode`)
- Account grouping (tied to `frmMastGroup`)
- Mobile numbers for WhatsApp/SMS integration
- Account code range configurable per company

### 1.3 Account Group Master (`frmMastGroup`)
- Two-level hierarchy: Primary Group → Detail Account
- Fixed groups (system reserved) vs. dynamic user-defined groups
- Classification for Balance Sheet and P&L separation:
  - Liability (TopGroup 90001)
  - Asset (TopGroup 90002)
  - P&L (TopGroup 90003)
  - Stock (TopGroup 90018)
- Debit balance reversal flag (`AgCode1`)

### 1.4 Company Group Master (`frmMastCompGroup`)
- Party/weaver classification and grouping
- Used for filtering in reports (group-wise outstanding, etc.)

### 1.5 Item / Commodity Master (`frmMastItem`, `frmMastItemOther`)
- Commodity/product definitions
- Item code and description
- Unit of measure linkage
- HSN code for GST compliance
- Non-standard/other items in a separate master

### 1.6 Narration Master (`frmMastNarration`, `frmMastNarrationTax`, `frmMastNarrationTaxGST`)
- Transaction narration text library
- Narration types: Sales, Purchase, Journal
- Tax narration master: Maps narration to tax rate + liability account
  - Pre-GST version (VAT era)
  - GST version (CGST/SGST/IGST accounts)

### 1.7 User Master (`frmMastUser`, `frmAdmUser`)
- User creation with username and password
- Admin vs Regular user classification
- Login control (active/inactive)
- Default password change enforcement

### 1.8 Other Reference Masters
| Master | Purpose |
|---|---|
| Count Master | Unit of measure definitions (metres, kg, bags) |
| Godown Master | Storage location / warehouse definitions |
| Transport Master | Transportation provider definitions |
| Sizer/Consignee Master | Third-party warehouse handlers |
| Mill Expense Description | Expense categories for mill processing charges |

### 1.9 System Settings (`frmSetting`)
Default account code assignments for automatic posting:

| Setting Area | Description |
|---|---|
| Purchase Default Accounts | Trade purchase, taxable purchase, SIT, consignment, depot |
| Sales Default Accounts | Trade sale, taxable sale, SIT, consignment, depot |
| TDS Account | TDS payable account code |
| Service Tax Account | Legacy service tax account |
| Brokerage Account | Commission/brokerage payable |
| Bank Commission Account | Bank charges account |
| Late Payment Interest Account | Interest income account |
| Bill Debit Note Account | Debit note contra account |
| TCS Pay/Receive Accounts | TCS liability and receivable accounts |
| VAT Voucher Generation Flag | Whether to auto-generate VAT vouchers |
| Mill Bill Debit Note Config | Mill bill adjustment configuration |

---

## Key Data Captured

| Field | Description |
|---|---|
| `CompCode`, `CompName` | Company identifier and name |
| `CompYear` | Fiscal year (4-char, e.g., "2526") |
| `GSTIN`, `PAN`, `TAN` | Tax registration numbers |
| `CST`, `BST` | Pre-GST state/central tax registrations |
| `AcCode`, `AcName` | Account code and name |
| `AgCode`, `AgCode1` | Primary group, debit reversal group |
| `AcBalance`, `BalanceCD` | Opening balance and Dr/Cr direction |
| `AcIntPerc` | Interest percentage on account |
| `AcBrkCode` | Linked broker account code |
| `NarrCode`, `NarrType` | Narration code and type |
| `MastTaxPurAcCode` | Purchase tax liability account |
| `MastTaxSAAcCode` | Sales tax liability account |
| `GSTRate`, `VATRate` | Tax percentages |
| `IsFixGroup` | Fixed group flag (system-reserved) |
| `HSNCode` | HSN code for GST classification |

---

## Business Rules & Logic

- Account code range is configurable per company (start/end codes enforced)
- Opening balances must be set at fiscal year start before transactions begin
- Narration groups separate sales and purchase tax narrations
- Fixed account groups cannot be renamed or deleted by users
- Interest percentage at account level used by Late Payment Interest module
- System enforces mandatory default account assignments in Settings before transactions can be posted
- Admin user bypasses all menu/privilege restrictions
- Broker account linkage (`AcBrkCode`) enables automatic commission calculations
- HSN codes in item master propagate to GST invoices automatically

---

## Integration Points

| Integrates With | How |
|---|---|
| All Transaction Modules | Account master used in every voucher |
| Tax Compliance | Tax narration masters provide rates and accounts |
| Finance / Accounting | Default accounts from Settings enable auto-posting |
| Reporting | Group hierarchy drives P&L and Balance Sheet structure |
| Communication | Mobile numbers from account master used for SMS/WhatsApp |
| User Privileges | User master feeds access control module |
