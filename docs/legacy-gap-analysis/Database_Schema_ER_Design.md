# Database Schema & Entity Relationship (ER) Diagram

**Purpose:** Define complete database schema for Dhanman commodity trading system; clarify Broker table design and relationships.

**Document Date:** 2026-05-16 | **Related Documents:** HITRIX_to_Dhanman_Mapping.md, Sales_Purchase_Inventory_Broker_Flow.md

---

## 1. BROKER TABLE DESIGN DECISION

### 1.1 Should Broker be a Separate Table?

**Answer: YES** ✅

**Reasoning:**

| Aspect | Account Flag Approach | Separate Broker Table |
|--------|---|---|
| **Data Model** | Brokers stored in Account table with type='BROKER' | Dedicated Broker entity with unique attributes |
| **Unique Fields** | Limited to Account fields; commission hardcoded | Broker-specific: sales commission %, purchase commission %, settlement terms, margin %, monthly fee |
| **AR/AP Handling** | Broker must have Account for AR/AP settlement | Broker links to Account only for payment; keeps logic separate |
| **Scalability** | Limits future broker-specific features | Easily extensible (broker ratings, performance metrics, etc.) |
| **Query Performance** | Slower: filter Account by type='BROKER' + tax rates | Faster: join Broker table directly |
| **Business Logic** | Broker commissions hardcoded in Invoice logic | Centralized in Broker service; clean separation |
| **Reporting** | Mixed with customer/vendor reports | Separate broker reporting suite |
| **Compliance** | Difficult to track broker-specific audits | Easy to audit broker commission calculations |

**Recommendation:** Create separate **Broker** table + link to Account (for AR/AP settlement)

---

## 2. COMPLETE DATABASE SCHEMA

### 2.1 Core Master Tables

#### **Table: companies**
```sql
CREATE TABLE companies (
  id BIGINT PRIMARY KEY AUTO_INCREMENT,
  code VARCHAR(10) UNIQUE NOT NULL,           -- Company code (e.g., "ABC")
  name VARCHAR(255) NOT NULL,
  gstin VARCHAR(15) UNIQUE,                    -- GST registration number
  pan VARCHAR(10),                             -- PAN
  currency_code VARCHAR(3) DEFAULT 'INR',      -- Currency
  financial_year_start_date DATE,              -- FY start (e.g., 2026-04-01)
  financial_year_end_date DATE,                -- FY end (e.g., 2027-03-31)
  address TEXT,
  city VARCHAR(100),
  state VARCHAR(100),
  postal_code VARCHAR(10),
  phone VARCHAR(20),
  email VARCHAR(100),
  is_active BOOLEAN DEFAULT TRUE,
  created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
  updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  
  INDEX idx_code (code),
  INDEX idx_gstin (gstin),
  INDEX idx_active (is_active)
);
```

#### **Table: accounts** (Customers, Suppliers, Employees)
```sql
CREATE TABLE accounts (
  id BIGINT PRIMARY KEY AUTO_INCREMENT,
  company_id BIGINT NOT NULL,
  code VARCHAR(20) UNIQUE NOT NULL,            -- Account code (e.g., "C001")
  name VARCHAR(255) NOT NULL,
  account_type ENUM('CUSTOMER', 'SUPPLIER', 'EMPLOYEE', 'INTERNAL', 'BANK') NOT NULL,
  
  -- Contact Details
  gstin VARCHAR(15),                           -- GST registration number
  pan VARCHAR(10),
  contact_person VARCHAR(100),
  phone VARCHAR(20),
  email VARCHAR(100),
  address TEXT,
  city VARCHAR(100),
  state VARCHAR(100),
  postal_code VARCHAR(10),
  
  -- Business Rules
  credit_limit DECIMAL(15,2) DEFAULT 0,        -- Max outstanding allowed
  credit_days INT DEFAULT 0,                   -- Payment term (days)
  interest_rate DECIMAL(5,2) DEFAULT 0,        -- Interest % p.a. on late payment
  tds_percentage DECIMAL(5,2) DEFAULT 0,       -- TDS % (for suppliers)
  gst_status ENUM('REGISTERED', 'UNREGISTERED', 'EXEMPT', 'COMPOSITION') DEFAULT 'REGISTERED',
  
  -- Account Flags
  is_consignment BOOLEAN DEFAULT FALSE,        -- Consignment supplier/customer
  is_depot BOOLEAN DEFAULT FALSE,              -- Depot location
  is_pay_to_mill BOOLEAN DEFAULT FALSE,        -- Mill payment routing
  
  -- GL Account Mapping
  gl_account_id BIGINT,                        -- Link to GL account
  
  is_active BOOLEAN DEFAULT TRUE,
  created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
  updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  
  FOREIGN KEY (company_id) REFERENCES companies(id),
  FOREIGN KEY (gl_account_id) REFERENCES gl_accounts(id),
  INDEX idx_company_id (company_id),
  INDEX idx_account_type (account_type),
  INDEX idx_gstin (gstin),
  INDEX idx_gst_status (gst_status),
  INDEX idx_active (is_active)
);
```

#### **Table: brokers** (NEW — Separate Broker Entity)
```sql
CREATE TABLE brokers (
  id BIGINT PRIMARY KEY AUTO_INCREMENT,
  company_id BIGINT NOT NULL,
  code VARCHAR(20) UNIQUE NOT NULL,            -- Broker code (e.g., "BRK001")
  name VARCHAR(255) NOT NULL,
  
  -- Commission Structure
  sales_commission_percent DECIMAL(5,2) DEFAULT 0,   -- Commission % on sales
  purchase_commission_percent DECIMAL(5,2) DEFAULT 0, -- Commission % on purchases
  commission_type ENUM('PERCENT', 'FIXED_AMOUNT') DEFAULT 'PERCENT',
  commission_amount DECIMAL(15,2) DEFAULT 0,  -- Fixed commission (if type=FIXED_AMOUNT)
  
  -- Settlement Terms
  settlement_frequency ENUM('DAILY', 'WEEKLY', 'MONTHLY', 'QUARTERLY') DEFAULT 'MONTHLY',
  settlement_day_of_month INT,                 -- Day of month for settlement (e.g., 15)
  payment_terms VARCHAR(100),                  -- E.g., "NET 30" or "2/10 NET 30"
  
  -- Financial Terms
  margin_percentage DECIMAL(5,2) DEFAULT 0,    -- Security margin required
  margin_amount DECIMAL(15,2) DEFAULT 0,       -- Fixed margin (if applicable)
  monthly_retainer DECIMAL(15,2) DEFAULT 0,    -- Monthly fee (if applicable)
  
  -- Contact & Banking
  contact_person VARCHAR(100),
  phone VARCHAR(20),
  email VARCHAR(100),
  address TEXT,
  city VARCHAR(100),
  state VARCHAR(100),
  postal_code VARCHAR(10),
  
  -- Banking Details (for settlement payments)
  bank_account_id BIGINT,                      -- Link to bank account (for remittance)
  ar_account_id BIGINT,                        -- Link to AR Account (for settlement tracking)
  
  -- Performance Metrics
  total_commissions_earned DECIMAL(15,2) DEFAULT 0,   -- YTD commission
  total_commissions_paid DECIMAL(15,2) DEFAULT 0,     -- YTD payments made
  outstanding_commission DECIMAL(15,2) DEFAULT 0,     -- Current outstanding
  
  is_active BOOLEAN DEFAULT TRUE,
  created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
  updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  
  FOREIGN KEY (company_id) REFERENCES companies(id),
  FOREIGN KEY (bank_account_id) REFERENCES bank_accounts(id),
  FOREIGN KEY (ar_account_id) REFERENCES accounts(id),
  INDEX idx_company_id (company_id),
  INDEX idx_code (code),
  INDEX idx_active (is_active)
);
```

#### **Table: products** (Items/Commodities)
```sql
CREATE TABLE products (
  id BIGINT PRIMARY KEY AUTO_INCREMENT,
  company_id BIGINT NOT NULL,
  code VARCHAR(20) UNIQUE NOT NULL,            -- Product code
  name VARCHAR(255) NOT NULL,
  description TEXT,
  
  -- UOM & Conversion
  primary_unit VARCHAR(10) NOT NULL,           -- MT, BAGS, QUINTAL, etc.
  secondary_unit VARCHAR(10),                  -- For conversion
  conversion_factor DECIMAL(10,2),             -- secondary to primary
  
  -- Classification
  hsn_code VARCHAR(8),                         -- HSN code for GST
  tax_rate DECIMAL(5,2) DEFAULT 0,             -- Default tax rate %
  
  -- Pricing
  base_purchase_price DECIMAL(15,2) DEFAULT 0, -- Reference purchase price
  base_selling_price DECIMAL(15,2) DEFAULT 0,  -- Reference selling price
  min_selling_price DECIMAL(15,2) DEFAULT 0,   -- Minimum allowed selling price
  max_purchase_price DECIMAL(15,2) DEFAULT 0,  -- Maximum allowed purchase price
  
  -- Inventory
  reorder_level DECIMAL(12,2) DEFAULT 0,       -- Minimum stock alert
  max_stock_level DECIMAL(12,2) DEFAULT 0,     -- Maximum stock capacity
  
  is_active BOOLEAN DEFAULT TRUE,
  created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
  updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  
  FOREIGN KEY (company_id) REFERENCES companies(id),
  INDEX idx_company_id (company_id),
  INDEX idx_code (code),
  INDEX idx_hsn (hsn_code),
  INDEX idx_active (is_active)
);
```

#### **Table: tax_rules**
```sql
CREATE TABLE tax_rules (
  id BIGINT PRIMARY KEY AUTO_INCREMENT,
  company_id BIGINT NOT NULL,
  code VARCHAR(20) NOT NULL,                   -- Tax code (T001, T002, etc.)
  description VARCHAR(255),
  
  -- Tax Rates
  cgst_rate DECIMAL(5,2) DEFAULT 0,            -- Central GST %
  sgst_rate DECIMAL(5,2) DEFAULT 0,            -- State GST %
  igst_rate DECIMAL(5,2) DEFAULT 0,            -- Integrated GST %
  
  -- Applicability
  applicable_products VARCHAR(1000),           -- JSON array of product IDs or HSN codes
  applicable_from DATE,
  applicable_to DATE,
  
  -- Special Rules
  is_rcm_applicable BOOLEAN DEFAULT FALSE,     -- Reverse Charge Mechanism
  is_tcs_applicable BOOLEAN DEFAULT FALSE,     -- Tax Collected at Source
  tcs_rate DECIMAL(5,2) DEFAULT 0,             -- TCS rate %
  
  is_active BOOLEAN DEFAULT TRUE,
  created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
  updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  
  FOREIGN KEY (company_id) REFERENCES companies(id),
  INDEX idx_company_id (company_id),
  INDEX idx_code (code),
  UNIQUE KEY uq_company_code (company_id, code)
);
```

#### **Table: warehouses** (Godowns)
```sql
CREATE TABLE warehouses (
  id BIGINT PRIMARY KEY AUTO_INCREMENT,
  company_id BIGINT NOT NULL,
  code VARCHAR(20) NOT NULL,                   -- Godown code
  name VARCHAR(255) NOT NULL,
  warehouse_type ENUM('OWN', 'LEASED', 'CONSIGNMENT', 'DEPOT') DEFAULT 'OWN',
  
  address TEXT,
  city VARCHAR(100),
  state VARCHAR(100),
  postal_code VARCHAR(10),
  
  -- Capacity
  capacity_units DECIMAL(12,2),                -- Max capacity (in primary unit)
  primary_unit VARCHAR(10),                    -- Unit for capacity (MT, etc.)
  
  -- Ownership (if consignment/depot)
  owner_account_id BIGINT,                     -- Link to owner Account (if not ours)
  
  is_active BOOLEAN DEFAULT TRUE,
  created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
  updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  
  FOREIGN KEY (company_id) REFERENCES companies(id),
  FOREIGN KEY (owner_account_id) REFERENCES accounts(id),
  INDEX idx_company_id (company_id),
  INDEX idx_warehouse_type (warehouse_type),
  UNIQUE KEY uq_company_code (company_id, code)
);
```

### 2.2 GL Chart of Accounts

#### **Table: gl_accounts**
```sql
CREATE TABLE gl_accounts (
  id BIGINT PRIMARY KEY AUTO_INCREMENT,
  company_id BIGINT NOT NULL,
  code VARCHAR(20) NOT NULL,                   -- GL code (e.g., "1100", "5101")
  name VARCHAR(255) NOT NULL,
  account_type ENUM('ASSET', 'LIABILITY', 'EQUITY', 'REVENUE', 'EXPENSE') NOT NULL,
  
  -- Hierarchy
  parent_account_id BIGINT,                    -- For sub-accounts
  level INT DEFAULT 1,                         -- Hierarchy level
  
  -- GL Account Groups
  account_group VARCHAR(100),                  -- ASSETS, LIABILITIES, etc.
  account_subgroup VARCHAR(100),               -- Debtors, Creditors, etc.
  
  -- GL Mapping (for auto-posting)
  is_summary BOOLEAN DEFAULT FALSE,            -- Summary account (no direct posting)
  allow_manual_posting BOOLEAN DEFAULT TRUE,   -- Can be posted manually
  
  is_active BOOLEAN DEFAULT TRUE,
  created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
  updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  
  FOREIGN KEY (company_id) REFERENCES companies(id),
  FOREIGN KEY (parent_account_id) REFERENCES gl_accounts(id),
  INDEX idx_company_id (company_id),
  INDEX idx_code (code),
  INDEX idx_account_type (account_type),
  UNIQUE KEY uq_company_code (company_id, code)
);
```

#### **Common GL Accounts Structure:**
```
ASSETS (1000-1999)
  ├─ Current Assets (1100-1199)
  │  ├─ Cash (1100)
  │  ├─ Bank Accounts (1110)
  │  ├─ Accounts Receivable (1200)
  │  ├─ Inventory (1300)
  │  └─ TDS Receivable (1350)
  └─ Other Assets (1900)

LIABILITIES (2000-2999)
  ├─ Current Liabilities (2100-2199)
  │  ├─ Accounts Payable (2100)
  │  ├─ Tax Payable (2200)
  │  │  ├─ CGST Payable (2210)
  │  │  ├─ SGST Payable (2220)
  │  │  └─ IGST Payable (2230)
  │  ├─ RCM Payable (2250)
  │  └─ TCS Payable (2260)
  └─ Other Liabilities (2900)

EQUITY (3000-3999)
  ├─ Capital (3100)
  └─ Retained Earnings (3200)

REVENUE (4000-4999)
  ├─ Sales Revenue (4100)
  ├─ Sales Returns (4110)
  └─ Other Income (4900)

EXPENSE (5000-5999)
  ├─ Cost of Goods Sold (5100)
  ├─ Commission Expense (5200)
  ├─ Interest Expense (5300)
  ├─ Tax Expense (5400)
  └─ Other Expenses (5900)
```

### 2.3 Transaction Tables

#### **Table: orders** (Sales Orders & Purchase Orders — Unified)
```sql
CREATE TABLE orders (
  id BIGINT PRIMARY KEY AUTO_INCREMENT,
  company_id BIGINT NOT NULL,
  order_type ENUM('SALES', 'PURCHASE', 'BOOKING') NOT NULL,
  
  -- Order Details
  order_number VARCHAR(20) NOT NULL,           -- Unique order number
  order_date DATE NOT NULL,
  reference_number VARCHAR(100),               -- Customer/Supplier PO number
  
  -- Party Details
  account_id BIGINT NOT NULL,                  -- Customer (SALES) or Supplier (PURCHASE)
  broker_id BIGINT,                            -- Broker (if applicable)
  
  -- Warehouse
  warehouse_id BIGINT,                         -- Destination (PURCHASE) or Source (SALES)
  
  -- Commercial Terms
  currency_code VARCHAR(3) DEFAULT 'INR',
  base_amount DECIMAL(15,2) DEFAULT 0,         -- Before tax
  tax_amount DECIMAL(15,2) DEFAULT 0,          -- Total tax
  total_amount DECIMAL(15,2) DEFAULT 0,        -- Base + Tax
  
  -- Broker Commission
  broker_commission_percent DECIMAL(5,2) DEFAULT 0,
  broker_commission_amount DECIMAL(15,2) DEFAULT 0,
  
  -- Special Flags
  is_rcm_applicable BOOLEAN DEFAULT FALSE,
  rcm_amount DECIMAL(15,2) DEFAULT 0,
  
  is_tcs_applicable BOOLEAN DEFAULT FALSE,
  tcs_amount DECIMAL(15,2) DEFAULT 0,
  
  -- Status
  status ENUM('DRAFT', 'CONFIRMED', 'PARTIAL_SHIPPED', 'SHIPPED', 'DELIVERED', 'INVOICED', 'CANCELLED') DEFAULT 'DRAFT',
  
  -- Dates
  delivery_date DATE,
  required_by_date DATE,
  
  -- Notes
  notes TEXT,
  terms_and_conditions TEXT,
  
  -- Audit
  created_by BIGINT,
  created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
  updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  
  FOREIGN KEY (company_id) REFERENCES companies(id),
  FOREIGN KEY (account_id) REFERENCES accounts(id),
  FOREIGN KEY (broker_id) REFERENCES brokers(id),
  FOREIGN KEY (warehouse_id) REFERENCES warehouses(id),
  FOREIGN KEY (created_by) REFERENCES users(id),
  INDEX idx_company_id (company_id),
  INDEX idx_order_type (order_type),
  INDEX idx_account_id (account_id),
  INDEX idx_broker_id (broker_id),
  INDEX idx_status (status),
  INDEX idx_order_date (order_date),
  UNIQUE KEY uq_company_number (company_id, order_number)
);
```

#### **Table: order_items** (Line Items)
```sql
CREATE TABLE order_items (
  id BIGINT PRIMARY KEY AUTO_INCREMENT,
  order_id BIGINT NOT NULL,
  product_id BIGINT NOT NULL,
  
  -- Quantity & Pricing
  quantity DECIMAL(12,2) NOT NULL,
  unit VARCHAR(10) NOT NULL,                   -- MT, BAGS, etc.
  rate DECIMAL(15,2) NOT NULL,                 -- Price per unit
  line_amount DECIMAL(15,2) DEFAULT 0,         -- qty × rate
  
  -- Tax
  tax_code_id BIGINT,                          -- Link to TaxRule
  tax_rate DECIMAL(5,2) DEFAULT 0,
  tax_amount DECIMAL(15,2) DEFAULT 0,
  line_total DECIMAL(15,2) DEFAULT 0,          -- line_amount + tax
  
  -- Warehouse (for multi-warehouse orders)
  source_warehouse_id BIGINT,                  -- SALES: where goods taken from
  
  -- GST Details
  hsn_code VARCHAR(8),
  cgst_amount DECIMAL(15,2) DEFAULT 0,
  sgst_amount DECIMAL(15,2) DEFAULT 0,
  igst_amount DECIMAL(15,2) DEFAULT 0,
  
  -- Tracking
  gst_transaction_id VARCHAR(50),              -- For GST matching
  
  -- Status
  status ENUM('PENDING', 'ALLOCATED', 'PACKED', 'SHIPPED', 'DELIVERED', 'INVOICED', 'CANCELLED') DEFAULT 'PENDING',
  
  created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
  updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  
  FOREIGN KEY (order_id) REFERENCES orders(id) ON DELETE CASCADE,
  FOREIGN KEY (product_id) REFERENCES products(id),
  FOREIGN KEY (tax_code_id) REFERENCES tax_rules(id),
  FOREIGN KEY (source_warehouse_id) REFERENCES warehouses(id),
  INDEX idx_order_id (order_id),
  INDEX idx_product_id (product_id),
  INDEX idx_status (status)
);
```

#### **Table: goods_issues** (Sales: Gate Pass/Dispatch)
```sql
CREATE TABLE goods_issues (
  id BIGINT PRIMARY KEY AUTO_INCREMENT,
  company_id BIGINT NOT NULL,
  order_id BIGINT NOT NULL,                    -- Link to SALES order
  
  issue_number VARCHAR(20) NOT NULL,           -- Gate pass number
  issue_date DATE NOT NULL,
  warehouse_id BIGINT NOT NULL,                -- Warehouse from which goods issued
  
  status ENUM('DRAFT', 'PENDING', 'PARTIAL', 'COMPLETED', 'CANCELLED') DEFAULT 'DRAFT',
  
  -- Tracking
  vehicle_number VARCHAR(20),
  driver_name VARCHAR(100),
  carrier_id BIGINT,                           -- Link to Carrier Account
  
  notes TEXT,
  
  created_by BIGINT,
  created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
  updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  
  FOREIGN KEY (company_id) REFERENCES companies(id),
  FOREIGN KEY (order_id) REFERENCES orders(id),
  FOREIGN KEY (warehouse_id) REFERENCES warehouses(id),
  FOREIGN KEY (carrier_id) REFERENCES accounts(id),
  FOREIGN KEY (created_by) REFERENCES users(id),
  INDEX idx_company_id (company_id),
  INDEX idx_order_id (order_id),
  INDEX idx_status (status),
  UNIQUE KEY uq_company_number (company_id, issue_number)
);
```

#### **Table: goods_issue_items** (Line items for dispatch)
```sql
CREATE TABLE goods_issue_items (
  id BIGINT PRIMARY KEY AUTO_INCREMENT,
  goods_issue_id BIGINT NOT NULL,
  order_item_id BIGINT NOT NULL,
  
  -- Quantity dispatched
  quantity_issued DECIMAL(12,2) NOT NULL,
  unit VARCHAR(10) NOT NULL,
  
  -- Bag/Container Tracking (optional for Wave 1; full in Wave 2)
  bag_ids VARCHAR(1000),                       -- Comma-separated bag IDs (if tracking enabled)
  
  created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
  updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  
  FOREIGN KEY (goods_issue_id) REFERENCES goods_issues(id) ON DELETE CASCADE,
  FOREIGN KEY (order_item_id) REFERENCES order_items(id),
  INDEX idx_goods_issue_id (goods_issue_id),
  INDEX idx_order_item_id (order_item_id)
);
```

#### **Table: goods_receipts** (Purchase: Goods Received Note)
```sql
CREATE TABLE goods_receipts (
  id BIGINT PRIMARY KEY AUTO_INCREMENT,
  company_id BIGINT NOT NULL,
  order_id BIGINT NOT NULL,                    -- Link to PURCHASE order
  
  receipt_number VARCHAR(20) NOT NULL,         -- GRN number
  receipt_date DATE NOT NULL,
  warehouse_id BIGINT NOT NULL,                -- Warehouse where goods received
  
  status ENUM('DRAFT', 'PENDING_QC', 'QC_PASS', 'QC_REJECT', 'RECEIVED', 'CANCELLED') DEFAULT 'DRAFT',
  
  -- Supplier Invoice Reference
  supplier_invoice_number VARCHAR(50),
  supplier_invoice_date DATE,
  
  notes TEXT,
  
  created_by BIGINT,
  created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
  updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  
  FOREIGN KEY (company_id) REFERENCES companies(id),
  FOREIGN KEY (order_id) REFERENCES orders(id),
  FOREIGN KEY (warehouse_id) REFERENCES warehouses(id),
  FOREIGN KEY (created_by) REFERENCES users(id),
  INDEX idx_company_id (company_id),
  INDEX idx_order_id (order_id),
  INDEX idx_status (status),
  UNIQUE KEY uq_company_number (company_id, receipt_number)
);
```

#### **Table: goods_receipt_items** (Line items for receipt)
```sql
CREATE TABLE goods_receipt_items (
  id BIGINT PRIMARY KEY AUTO_INCREMENT,
  goods_receipt_id BIGINT NOT NULL,
  order_item_id BIGINT NOT NULL,
  
  -- Quantity received (may differ from PO qty)
  quantity_received DECIMAL(12,2) NOT NULL,
  unit VARCHAR(10) NOT NULL,
  
  -- Quality Control
  qc_status ENUM('PENDING', 'PASS', 'REJECT', 'HOLD') DEFAULT 'PENDING',
  qc_notes TEXT,
  qc_checked_by BIGINT,
  qc_checked_date DATETIME,
  
  created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
  updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  
  FOREIGN KEY (goods_receipt_id) REFERENCES goods_receipts(id) ON DELETE CASCADE,
  FOREIGN KEY (order_item_id) REFERENCES order_items(id),
  FOREIGN KEY (qc_checked_by) REFERENCES users(id),
  INDEX idx_goods_receipt_id (goods_receipt_id),
  INDEX idx_qc_status (qc_status)
);
```

#### **Table: invoices** (Sales & Purchase Invoices — Unified)
```sql
CREATE TABLE invoices (
  id BIGINT PRIMARY KEY AUTO_INCREMENT,
  company_id BIGINT NOT NULL,
  order_id BIGINT NOT NULL,                    -- Link to order
  
  invoice_type ENUM('SALES', 'PURCHASE', 'DEBIT_NOTE', 'CREDIT_NOTE') NOT NULL,
  invoice_number VARCHAR(20) NOT NULL,
  invoice_date DATE NOT NULL,
  due_date DATE,
  
  -- Party
  account_id BIGINT NOT NULL,                  -- Customer (SALES) or Supplier (PURCHASE)
  broker_id BIGINT,                            -- Broker (if applicable)
  
  -- Amounts
  base_amount DECIMAL(15,2) DEFAULT 0,
  discount_amount DECIMAL(15,2) DEFAULT 0,
  tax_amount DECIMAL(15,2) DEFAULT 0,
  total_amount DECIMAL(15,2) DEFAULT 0,
  
  -- Tax Breakdown
  cgst_amount DECIMAL(15,2) DEFAULT 0,
  sgst_amount DECIMAL(15,2) DEFAULT 0,
  igst_amount DECIMAL(15,2) DEFAULT 0,
  
  -- Special Tax
  is_rcm_applicable BOOLEAN DEFAULT FALSE,
  rcm_amount DECIMAL(15,2) DEFAULT 0,
  
  is_tcs_applicable BOOLEAN DEFAULT FALSE,
  tcs_amount DECIMAL(15,2) DEFAULT 0,
  
  tds_percentage DECIMAL(5,2) DEFAULT 0,
  tds_amount DECIMAL(15,2) DEFAULT 0,
  
  -- Broker Commission
  broker_commission_percent DECIMAL(5,2) DEFAULT 0,
  broker_commission_amount DECIMAL(15,2) DEFAULT 0,
  
  -- E-Invoice (IRN)
  irn_number VARCHAR(16),                      -- Invoice Registration Number
  irn_qr_code LONGTEXT,                        -- QR Code (base64)
  irn_ack_number VARCHAR(16),
  irn_generated_at TIMESTAMP,
  irn_status ENUM('PENDING', 'SUCCESS', 'FAILED', 'CANCELLED') DEFAULT 'PENDING',
  
  -- Status
  status ENUM('DRAFT', 'SENT', 'PARTIALLY_PAID', 'PAID', 'CANCELLED', 'WRITEOFF') DEFAULT 'DRAFT',
  
  -- GL Reference
  journal_entry_id BIGINT,                     -- Link to GL journal entry
  
  notes TEXT,
  
  created_by BIGINT,
  created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
  updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  
  FOREIGN KEY (company_id) REFERENCES companies(id),
  FOREIGN KEY (order_id) REFERENCES orders(id),
  FOREIGN KEY (account_id) REFERENCES accounts(id),
  FOREIGN KEY (broker_id) REFERENCES brokers(id),
  FOREIGN KEY (journal_entry_id) REFERENCES journal_entries(id),
  FOREIGN KEY (created_by) REFERENCES users(id),
  INDEX idx_company_id (company_id),
  INDEX idx_invoice_type (invoice_type),
  INDEX idx_account_id (account_id),
  INDEX idx_broker_id (broker_id),
  INDEX idx_status (status),
  INDEX idx_irn_number (irn_number),
  INDEX idx_invoice_date (invoice_date),
  UNIQUE KEY uq_company_number (company_id, invoice_number)
);
```

#### **Table: invoice_items** (Line items for invoice)
```sql
CREATE TABLE invoice_items (
  id BIGINT PRIMARY KEY AUTO_INCREMENT,
  invoice_id BIGINT NOT NULL,
  order_item_id BIGINT NOT NULL,
  product_id BIGINT NOT NULL,
  
  -- Quantity & Pricing
  quantity DECIMAL(12,2) NOT NULL,
  unit VARCHAR(10) NOT NULL,
  rate DECIMAL(15,2) NOT NULL,
  line_amount DECIMAL(15,2) DEFAULT 0,
  
  -- Tax
  tax_code_id BIGINT,
  tax_rate DECIMAL(5,2) DEFAULT 0,
  cgst_amount DECIMAL(15,2) DEFAULT 0,
  sgst_amount DECIMAL(15,2) DEFAULT 0,
  igst_amount DECIMAL(15,2) DEFAULT 0,
  tax_amount DECIMAL(15,2) DEFAULT 0,
  
  line_total DECIMAL(15,2) DEFAULT 0,          -- line_amount + tax
  
  -- HSN for GST
  hsn_code VARCHAR(8),
  
  -- GL Posting Reference
  gl_account_id BIGINT,                        -- GL account for this item (Revenue or COGS)
  
  created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
  updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  
  FOREIGN KEY (invoice_id) REFERENCES invoices(id) ON DELETE CASCADE,
  FOREIGN KEY (order_item_id) REFERENCES order_items(id),
  FOREIGN KEY (product_id) REFERENCES products(id),
  FOREIGN KEY (tax_code_id) REFERENCES tax_rules(id),
  FOREIGN KEY (gl_account_id) REFERENCES gl_accounts(id),
  INDEX idx_invoice_id (invoice_id),
  INDEX idx_product_id (product_id)
);
```

### 2.4 Finance Tables

#### **Table: journal_entries**
```sql
CREATE TABLE journal_entries (
  id BIGINT PRIMARY KEY AUTO_INCREMENT,
  company_id BIGINT NOT NULL,
  
  journal_number VARCHAR(20) NOT NULL,         -- JV number
  entry_date DATE NOT NULL,
  reference_type VARCHAR(50),                  -- 'INVOICE', 'PAYMENT', 'MANUAL'
  reference_id VARCHAR(100),                   -- Invoice ID, Payment ID, etc.
  
  description VARCHAR(500),
  
  total_debit DECIMAL(15,2) DEFAULT 0,
  total_credit DECIMAL(15,2) DEFAULT 0,
  
  status ENUM('DRAFT', 'POSTED', 'REVERSED', 'CANCELLED') DEFAULT 'DRAFT',
  
  created_by BIGINT,
  posted_by BIGINT,
  posted_date DATETIME,
  
  created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
  updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  
  FOREIGN KEY (company_id) REFERENCES companies(id),
  FOREIGN KEY (created_by) REFERENCES users(id),
  FOREIGN KEY (posted_by) REFERENCES users(id),
  INDEX idx_company_id (company_id),
  INDEX idx_entry_date (entry_date),
  INDEX idx_reference_type (reference_type),
  INDEX idx_status (status),
  UNIQUE KEY uq_company_number (company_id, journal_number)
);
```

#### **Table: journal_lines** (GL Posting Lines)
```sql
CREATE TABLE journal_lines (
  id BIGINT PRIMARY KEY AUTO_INCREMENT,
  journal_entry_id BIGINT NOT NULL,
  
  gl_account_id BIGINT NOT NULL,
  
  debit_amount DECIMAL(15,2) DEFAULT 0,
  credit_amount DECIMAL(15,2) DEFAULT 0,
  
  description VARCHAR(500),
  
  -- Optional Reference
  account_id BIGINT,                           -- Link to customer/supplier account (for AR/AP matching)
  
  created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
  
  FOREIGN KEY (journal_entry_id) REFERENCES journal_entries(id) ON DELETE CASCADE,
  FOREIGN KEY (gl_account_id) REFERENCES gl_accounts(id),
  FOREIGN KEY (account_id) REFERENCES accounts(id),
  INDEX idx_journal_entry_id (journal_entry_id),
  INDEX idx_gl_account_id (gl_account_id),
  INDEX idx_account_id (account_id)
);
```

#### **Table: payments** (Bank & Cash Payments)
```sql
CREATE TABLE payments (
  id BIGINT PRIMARY KEY AUTO_INCREMENT,
  company_id BIGINT NOT NULL,
  
  payment_number VARCHAR(20) NOT NULL,
  payment_date DATE NOT NULL,
  payment_type ENUM('BANK_TRANSFER', 'CHEQUE', 'CASH', 'DD', 'NEFT', 'RTGS', 'DEMAND_DRAFT') NOT NULL,
  
  -- Party
  account_id BIGINT NOT NULL,                  -- Account being paid to/from
  
  -- Amount
  payment_amount DECIMAL(15,2) NOT NULL,
  
  -- Bank Details
  bank_account_id BIGINT,                      -- Payment from which bank account
  cheque_number VARCHAR(20),                   -- If cheque payment
  cheque_date DATE,
  
  -- TDS Deduction
  tds_deducted DECIMAL(15,2) DEFAULT 0,        -- TDS amount deducted
  net_payment DECIMAL(15,2) DEFAULT 0,         -- payment_amount - tds_deducted
  
  -- Reference
  invoice_ids VARCHAR(1000),                   -- Comma-separated invoice IDs paid against
  
  status ENUM('DRAFT', 'AUTHORIZED', 'PROCESSED', 'CLEARED', 'CANCELLED') DEFAULT 'DRAFT',
  
  -- GL Reference
  journal_entry_id BIGINT,
  
  created_by BIGINT,
  created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
  updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  
  FOREIGN KEY (company_id) REFERENCES companies(id),
  FOREIGN KEY (account_id) REFERENCES accounts(id),
  FOREIGN KEY (bank_account_id) REFERENCES bank_accounts(id),
  FOREIGN KEY (journal_entry_id) REFERENCES journal_entries(id),
  FOREIGN KEY (created_by) REFERENCES users(id),
  INDEX idx_company_id (company_id),
  INDEX idx_account_id (account_id),
  INDEX idx_payment_date (payment_date),
  INDEX idx_status (status),
  UNIQUE KEY uq_company_number (company_id, payment_number)
);
```

### 2.5 Inventory Tables

#### **Table: stock_balance** (Current Stock Levels)
```sql
CREATE TABLE stock_balance (
  id BIGINT PRIMARY KEY AUTO_INCREMENT,
  company_id BIGINT NOT NULL,
  product_id BIGINT NOT NULL,
  warehouse_id BIGINT NOT NULL,
  
  -- Stock Quantities
  opening_balance DECIMAL(12,2) DEFAULT 0,
  inward_quantity DECIMAL(12,2) DEFAULT 0,     -- Goods received
  outward_quantity DECIMAL(12,2) DEFAULT 0,    -- Goods issued
  reserved_quantity DECIMAL(12,2) DEFAULT 0,   -- Reserved for sales orders
  closing_balance DECIMAL(12,2) DEFAULT 0,     -- Available = Closing - Reserved
  
  -- Unit
  unit VARCHAR(10) NOT NULL,
  
  -- Cost (weighted average)
  cost_per_unit DECIMAL(15,2) DEFAULT 0,
  total_cost DECIMAL(15,2) DEFAULT 0,          -- closing_balance × cost_per_unit
  
  last_transaction_date DATE,
  
  created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
  updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  
  FOREIGN KEY (company_id) REFERENCES companies(id),
  FOREIGN KEY (product_id) REFERENCES products(id),
  FOREIGN KEY (warehouse_id) REFERENCES warehouses(id),
  INDEX idx_company_id (company_id),
  INDEX idx_product_id (product_id),
  INDEX idx_warehouse_id (warehouse_id),
  UNIQUE KEY uq_product_warehouse (product_id, warehouse_id)
);
```

#### **Table: stock_ledger** (Transaction History)
```sql
CREATE TABLE stock_ledger (
  id BIGINT PRIMARY KEY AUTO_INCREMENT,
  company_id BIGINT NOT NULL,
  product_id BIGINT NOT NULL,
  warehouse_id BIGINT NOT NULL,
  
  transaction_type ENUM('OPENING', 'INWARD_GOODS_RECEIPT', 'OUTWARD_GOODS_ISSUE', 'INTERNAL_TRANSFER', 'ADJUSTMENT', 'RETURN') NOT NULL,
  reference_type VARCHAR(50),                  -- 'GOODS_RECEIPT', 'GOODS_ISSUE', 'INVOICE'
  reference_id BIGINT,
  
  quantity_in DECIMAL(12,2) DEFAULT 0,
  quantity_out DECIMAL(12,2) DEFAULT 0,
  unit VARCHAR(10) NOT NULL,
  
  running_balance DECIMAL(12,2) DEFAULT 0,
  cost_per_unit DECIMAL(15,2) DEFAULT 0,
  
  notes TEXT,
  
  created_by BIGINT,
  transaction_date DATE,
  created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
  
  FOREIGN KEY (company_id) REFERENCES companies(id),
  FOREIGN KEY (product_id) REFERENCES products(id),
  FOREIGN KEY (warehouse_id) REFERENCES warehouses(id),
  FOREIGN KEY (created_by) REFERENCES users(id),
  INDEX idx_company_id (company_id),
  INDEX idx_product_id (product_id),
  INDEX idx_warehouse_id (warehouse_id),
  INDEX idx_transaction_type (transaction_type),
  INDEX idx_transaction_date (transaction_date)
);
```

---

## 3. ENTITY RELATIONSHIP DIAGRAM (ER DIAGRAM)

### 3.1 Text-Based ER Overview

```
┌────────────────────────────────────────────────────────────────────────────┐
│                         DHANMAN TRADING SYSTEM ER DIAGRAM                  │
│                          (Simplified View)                                  │
└────────────────────────────────────────────────────────────────────────────┘


                              ┌─────────────────┐
                              │   companies     │
                              │  (Company Setup)│
                              └────────┬────────┘
                                       │
                 ┌─────────┬──────────┬─┬──────┬──────────┬─────────┐
                 │         │          │ │      │          │         │
                 │         │          │ │      │          │         │
                 ▼         ▼          ▼ ▼      ▼          ▼         ▼
          ┌──────────┐ ┌────────┐ ┌─────────┐ ┌───────┐ ┌──────┐ ┌─────┐
          │ accounts │ │ brokers│ │ products│ │ users │ │ taxes│ │warehouses│
          │(Parties) │ │        │ │         │ │       │ │rules │ │(Godowns) │
          └────┬─────┘ └─┬──────┘ └────┬────┘ │       │ │      │ └────┬────┘
               │         │             │      │       │ │      │      │
               │         │             │      │       │ │      │      │
         ┌─────┴────┬────┴─────┬───────┴──┐   └───────┘ └──────┘      │
         │          │          │          │                           │
         │          │          │          │                           │
         ▼          ▼          ▼          ▼                           │
    ┌──────────┐ ┌──────────┐ ┌────────────┐                         │
    │ orders   │◄┤gl_accounts◄┤tax_rules(FK)                         │
    │(Sales &  │ │(Chart of  │ └────────────┘                         │
    │ Purchase)│ │ Accounts) │                                        │
    └────┬─────┘ └──────────┘                                        │
         │                                                            │
    ┌────┴──────────┐                                                │
    │ order_items   │◄──────────────────────────────────────────────┤
    │               │                                                │
    └────┬──────────┘                                                │
         │                                                            │
    ┌────┼──────────────────┬─────────────────┐                     │
    │    │                  │                 │                     │
    ▼    ▼                  ▼                 ▼                     ▼
┌────────────┐ ┌──────────────────┐ ┌────────────────┐ ┌─────────────────┐
│ invoices   │ │ goods_issues     │ │ goods_receipts │ │ stock_balance   │◄──┤
│(Sales &    │ │(Sales Dispatch)  │ │(Purchase GRN)  │ │(Current Stock)  │   │
│Purchase)   │ │                  │ │                │ └──────┬──────────┘   │
└────┬───────┘ └──────┬───────────┘ └────────┬───────┘        │              │
     │                │                      │                │              │
     │                │                      │                │              │
┌────┴──────┐ ┌──────┴────────┐ ┌───────────┴───────┐ ┌──────┴───────┐    │
│ invoice_  │ │goods_issue_   │ │goods_receipt_    │ │stock_ledger  │    │
│items      │ │items          │ │items             │ │(Transaction  │    │
│(Line Items)│ │(Line Items)   │ │(Line Items)      │ │ History)     │    │
└───────────┘ └───────────────┘ └──────────────────┘ └──────────────┘    │
                                                                           │
     ┌──────────────────────────────────────────────────────────────────┘
     │
     │    (REFERENCES)
     ▼
┌──────────────────┐
│ journal_entries  │
│ & journal_lines  │
│(GL Posting)      │
└──────────┬───────┘
           │
     ┌─────┴──────────┐
     │                │
     ▼                ▼
┌──────────┐   ┌──────────────┐
│payments  │   │bank_accounts │
│          │   │              │
└──────────┘   └──────────────┘
```

### 3.2 Detailed ER Relationships

#### **Core Sales Flow Relationships**

```
┌─────────────┐
│  companies  │
│             │
└──────┬──────┘
       │ 1:N
       │
       ▼
┌──────────────────────────────────────────────────────────────┐
│                    SALES FLOW                                │
└──────────────────────────────────────────────────────────────┘
       │
       ├─ 1:N ──→ ┌─────────────┐    (Customer)
       │          │  accounts   │◄───────────┐
       │          │  (CUSTOMER) │            │
       │          └─────────────┘        PK: id
       │               │
       │          (FK: id)
       │               │
       ├─ 1:N ──→ ┌─────────────┐    (Broker)
       │          │  brokers    │
       │          │             │
       │          └─────────────┘
       │               │
       │          (FK: id)
       │               │
       ├─ 1:N ──→ ┌──────────────────┐
       │          │ orders (type=    │
       │          │  SALES)          │
       │          └────────┬─────────┘
       │                   │
       │           1:N (FK: order_id)
       │                   │
       ├─────────────→ ┌──────────────────┐
       │               │  order_items     │
       │               │  (qty, rate)     │
       │               └────────┬─────────┘
       │                        │
       │            1:1 (FK: order_item_id)
       │                        │
       ├─ 1:N ──→ ┌────────────────┐
       │          │ goods_issues   │───── Warehouse
       │          │ (Gate Pass)    │      (source_warehouse_id)
       │          └────────┬───────┘
       │                   │
       │           1:N (FK: goods_issue_id)
       │                   │
       │          ┌────────▼──────────┐
       │          │goods_issue_items  │
       │          └───────────────────┘
       │
       ├─ 1:N ──→ ┌──────────────────┐
       │          │  invoices        │
       │          │ (type=SALES)     │
       │          │ ┌─ irn_number    │
       │          │ ┌─ status        │
       │          │ └─ total_amount  │
       │          └────────┬─────────┘
       │                   │
       │           1:N (FK: invoice_id)
       │                   │
       │          ┌────────▼──────────┐
       │          │ invoice_items     │
       │          │ (qty, rate, tax)  │
       │          └───────────────────┘
       │
       ├─ 1:N ──→ ┌──────────────────┐
       │          │ journal_entries  │
       │          │ (GL Posting)     │
       │          │ reference_id=    │
       │          │   invoice.id     │
       │          └────────┬─────────┘
       │                   │
       │           1:N (FK: journal_entry_id)
       │                   │
       │          ┌────────▼──────────┐
       │          │ journal_lines     │
       │          │ ┌─ GL Account     │
       │          │ ┌─ Debit Amount   │
       │          │ └─ Credit Amount  │
       │          └───────────────────┘
```

#### **Core Purchase Flow Relationships**

```
┌─────────────┐
│  companies  │
│             │
└──────┬──────┘
       │ 1:N
       │
       ▼
┌──────────────────────────────────────────────────────────────┐
│                   PURCHASE FLOW                              │
└──────────────────────────────────────────────────────────────┘
       │
       ├─ 1:N ──→ ┌─────────────┐    (Supplier)
       │          │  accounts   │◄─────────────┐
       │          │ (SUPPLIER)  │              │
       │          └─────────────┘          PK: id
       │               │
       │          (FK: id)
       │               │
       ├─ 1:N ──→ ┌─────────────┐    (Broker)
       │          │  brokers    │
       │          │             │
       │          └─────────────┘
       │               │
       │          (FK: id)
       │               │
       ├─ 1:N ──→ ┌──────────────────┐
       │          │ orders (type=    │
       │          │  PURCHASE)       │
       │          └────────┬─────────┘
       │                   │
       │           1:N (FK: order_id)
       │                   │
       ├─────────────→ ┌──────────────────┐
       │               │  order_items     │
       │               │  (qty, rate)     │
       │               └────────┬─────────┘
       │                        │
       │            1:1 (FK: order_item_id)
       │                        │
       ├─ 1:N ──→ ┌────────────────┐
       │          │ goods_receipts │───── Warehouse
       │          │ (GRN)          │      (destination_warehouse_id)
       │          └────────┬───────┘
       │                   │
       │           1:N (FK: goods_receipt_id)
       │                   │
       │          ┌────────▼──────────────┐
       │          │goods_receipt_items    │
       │          │(qty received, QC)     │
       │          └───────────────────────┘
       │
       ├─ 1:N ──→ ┌──────────────────┐
       │          │  invoices        │
       │          │ (type=PURCHASE)  │
       │          │ ┌─ tds_amount    │
       │          │ ┌─ is_rcm_        │
       │          │   applicable     │
       │          │ └─ total_amount  │
       │          └────────┬─────────┘
       │                   │
       │           1:N (FK: invoice_id)
       │                   │
       │          ┌────────▼──────────┐
       │          │ invoice_items     │
       │          │ (qty, rate, tax)  │
       │          └───────────────────┘
       │
       ├─ 1:N ──→ ┌──────────────────┐
       │          │ journal_entries  │
       │          │ (GL Posting)     │
       │          │ reference_id=    │
       │          │   invoice.id     │
       │          └────────┬─────────┘
       │                   │
       │           1:N (FK: journal_entry_id)
       │                   │
       │          ┌────────▼──────────┐
       │          │ journal_lines     │
       │          │ ┌─ GL Account     │
       │          │ ┌─ Debit Amount   │
       │          │ └─ Credit Amount  │
       │          └───────────────────┘
       │
       ├─ 1:N ──→ ┌──────────────────┐
       │          │  payments        │
       │          │(bank transfers)  │
       │          │┌─ tds_deducted   │
       │          │┌─ net_payment    │
       │          │ └─ status        │
       │          └────────┬─────────┘
       │                   │
       │           1:N (FK: payment id)
       │                   │
       │          ┌────────▼──────────┐
       │          │ journal_entries  │
       │          │ (GL: Bank Debit, │
       │          │  AP/TDS Credit)  │
       │          └───────────────────┘
```

#### **Inventory Flow Relationships**

```
┌─────────────┐
│  companies  │
└──────┬──────┘
       │ 1:N
       ▼
   ┌────────────┐         ┌────────────┐
   │ products   │◄────1:N─┤ warehouses │
   │            │         │ (Godowns)  │
   └────┬───────┘         └────────────┘
        │
        │ 1:N
        ▼
    ┌──────────────────────────────────────┐
    │     stock_balance                    │
    │  (Product per Warehouse)             │
    │  ┌─ opening_balance                  │
    │  ┌─ inward_quantity                  │
    │  ┌─ outward_quantity                 │
    │  ┌─ reserved_quantity (for Sales)    │
    │  ├─ closing_balance                  │
    │  └─ cost_per_unit                    │
    └────────┬─────────────────────────────┘
             │
             │ 1:N (FK: product_id, warehouse_id)
             │
             ▼
    ┌──────────────────────────────────────┐
    │     stock_ledger                     │
    │  (Transaction History)               │
    │  ┌─ transaction_type                 │
    │  │   OPENING                         │
    │  │   INWARD_GOODS_RECEIPT            │
    │  │   OUTWARD_GOODS_ISSUE             │
    │  │   INTERNAL_TRANSFER               │
    │  ├─ reference_id (GRN or Dispatch)   │
    │  ├─ quantity_in                      │
    │  ├─ quantity_out                     │
    │  └─ running_balance                  │
    └──────────────────────────────────────┘
```

#### **Broker Relationship (Central Hub)**

```
┌──────────────────────────────────────────────────────────────┐
│                     BROKER RELATIONSHIPS                     │
│                    (Central Hub)                             │
└──────┬───────────────────────────────────────────────────────┘
       │
       ▼
   ┌──────────────────────────────────────┐
   │          brokers                     │
   │  ┌─ code                             │
   │  ├─ sales_commission_percent         │
   │  ├─ purchase_commission_percent      │
   │  ├─ settlement_frequency             │
   │  ├─ margin_percentage                │
   │  ├─ total_commissions_earned         │
   │  ├─ total_commissions_paid           │
   │  ├─ outstanding_commission           │
   │  ├─ ar_account_id (FK→accounts)      │
   │  └─ bank_account_id (FK→bank_accounts)
   └────────────┬────────────────────────┘
                │
        ┌───────┼────────┐
        │       │        │
        │       │        │
   (1:N on     │        │
   orders)     │        │
        │      │        │
        ▼      │        │
    ┌────────────┐      │
    │   orders   │      │
    │ (broker_id)├───┐  │
    └────────────┘   │  │
         │           │  │
         │   ┌──────────────────┐
         │   │   invoices       │
         │   │ (broker_id)      │
         │   │ ┌ broker_commission│
         │   │   _percent       │
         │   │ ├ broker_commission│
         │   │   _amount        │
         │   └──────────────────┘
         │           │
         │   ┌───────┴────┐
         │   │            │
         │   │     GL Entries:
         │   │     Debit: Commission Expense (5200)
         │   │     Credit: AP—Broker (2100)
         │   │
         │   └──────────────┐
         │                  │
         ▼                  │
    ┌─────────────┐     (settlement)
    │  AR Account │         │
    │  (for broker)         │
    │ (links to   │         │
    │  broker.ar_ │         │
    │ account_id) │         │
    └─────────────┘         │
                            │
                   (Monthly Settlement)
                            │
                            ▼
                   ┌──────────────┐
                   │   payments   │
                   │ (to broker)  │
                   │account_id= │
                   │broker.ar_   │
                   │account_id   │
                   └──────────────┘
```

---

## 4. DATABASE MIGRATION STRATEGY

### 4.1 Migration Sequence (Wave 1)

**Phase 1: Create Core Master Tables (Week 1)**
```sql
-- Core setup
1. companies
2. gl_accounts (chart of accounts)
3. accounts (with new fields: credit_limit, interest_rate, tds_%, account_type, flags)
4. brokers (NEW table)
5. products (with new fields: min_selling_price, max_purchase_price)
6. warehouses
7. tax_rules
8. users
```

**Phase 2: Create Transaction Tables (Week 2)**
```sql
9. orders (SALES & PURCHASE unified)
10. order_items
11. goods_issues (Gate Pass)
12. goods_issue_items
13. goods_receipts (GRN)
14. goods_receipt_items
15. invoices (with IRN, broker, RCM, TCS, TDS fields)
16. invoice_items
```

**Phase 3: Create Finance Tables (Week 3)**
```sql
17. journal_entries
18. journal_lines
19. gl_accounts (existing, but add indexes)
20. payments
21. bank_accounts
```

**Phase 4: Create Inventory Tables (Week 3)**
```sql
22. stock_balance
23. stock_ledger
```

### 4.2 Indexes & Performance

**Critical Indexes for Wave 1:**
```sql
-- Frequently searched by Account
CREATE INDEX idx_accounts_company_type ON accounts(company_id, account_type);
CREATE INDEX idx_accounts_gstin ON accounts(gstin);

-- Order lookup
CREATE INDEX idx_orders_company_status ON orders(company_id, status);
CREATE INDEX idx_orders_account_date ON orders(account_id, order_date);

-- Invoice tracking
CREATE INDEX idx_invoices_company_status ON invoices(company_id, status);
CREATE INDEX idx_invoices_irn ON invoices(irn_number);
CREATE INDEX idx_invoices_account ON invoices(account_id);

-- GL posting
CREATE INDEX idx_journal_company_date ON journal_entries(company_id, entry_date);
CREATE INDEX idx_journal_lines_gl_account ON journal_lines(gl_account_id);

-- Stock queries
CREATE INDEX idx_stock_balance_product_warehouse ON stock_balance(product_id, warehouse_id);
CREATE INDEX idx_stock_ledger_product_date ON stock_ledger(product_id, transaction_date);
```

---

## 5. ANSWER: BROKER TABLE DESIGN

### ✅ YES — Create Separate Broker Table

**Why:**

1. **Unique Attributes**: Brokers have commission rates, settlement terms, margin requirements that don't fit Account entity
2. **Clean Separation**: Account = party for AR/AP; Broker = business entity with specific rules
3. **Scalability**: Easy to add broker-specific features (ratings, performance metrics, compliance)
4. **Query Efficiency**: Faster broker lookup without filtering by account type
5. **Business Logic**: Centralized broker commission calculations (not scattered in Invoice logic)
6. **Reporting**: Separate broker reports (commission summary, settlement status)

### ✅ Link to Account (For AR/AP Settlement)

Broker table has FK to Account (`ar_account_id`), which means:
- Broker settlement creates AR entries (commission owed to broker)
- Separate payment tracking for broker commissions
- GL entries for Commission Expense vs. AP to Broker

### ✅ Use Broker ID Where Needed

```
Orders: order.broker_id (FK→brokers.id)
Invoices: invoice.broker_id (FK→brokers.id)
Payments: payments still link to account_id, but
          for broker settlement, account_id = broker.ar_account_id
```

---

## 6. SUMMARY: COMPLETE ER SCHEMA

| Component | Count | Status |
|-----------|-------|--------|
| **Master Tables** | 7 | ✅ Defined |
| **GL Accounts** | 1 | ✅ Defined |
| **Transaction Tables** | 8 | ✅ Defined |
| **Finance Tables** | 3 | ✅ Defined |
| **Inventory Tables** | 2 | ✅ Defined |
| **Total Tables** | 21 | ✅ Ready for implementation |
| **Total Indexes** | 30+ | ✅ Recommended |

**Next Steps:**
1. ✅ Review schema with database team
2. ✅ Create migration scripts (SQL files per phase)
3. ✅ Test on dev environment
4. ✅ Load HITRIX data for integration testing
5. ✅ Optimize queries based on actual data volume

---

**Broker Table is SEPARATE & ESSENTIAL for proper commodity trading workflow.** 🎯
