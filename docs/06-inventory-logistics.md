# Domain 6: Inventory & Logistics

## Purpose
Track physical goods movement at the bag/cartoon level — from inward receipt through storage to outward despatch. Provides documentary control via gate passes and delivery confirmations. This is critical in textile trading where weight reconciliation and physical custody tracking are legally and commercially important.

---

## Key Features

### 6.1 Bag Entry / Goods Inward (`frmBagEntry`)
- Receipt of goods in bags or cartons
- Sequential bag number assignment per lot/item
- Individual bag weight capture
- Total weight vs balance weight reconciliation
- Goods lot / batch identification
- Supplier reference and lorry receipt number
- Inward date and time recording
- Links to purchase bill reference

### 6.2 Gate Pass Entry — Inward (`frmGatePass`, `frmGatePassNew`)
- Documentary control for goods entering godown
- Party and goods description
- Bag/cartoon list attached
- Vehicle/transport details (truck number, driver)
- Gate entry time
- Printed gate pass (PDF format)
- WhatsApp delivery to party/broker

### 6.3 Gate Pass Entry — Outward (`frmGatePassNew`, `xxfrmGatePassNew`)
- Documentary control for goods leaving godown
- Delivery party and destination
- Bag numbers / total bags and weight
- Vehicle number and transporter
- Gate exit time
- Printed gate pass for truck/driver
- Linked to sales bill or despatch order

### 6.4 GP Delivery Entry (`frmGPDeliveryEntry`, `frmGPDeliveryEntryNew`)
- Confirm delivery of goods to final recipient
- Delivery party and receiver name
- Delivery date and time
- Condition of goods (if noted)
- Signature / acknowledgment reference
- Links gate pass to final delivery confirmation

### 6.5 Item / Commodity Master (`frmItemMast`)
- Item code, description
- Unit of measure
- HSN code
- Category classification

### 6.6 Form Entry (`frmFormEntry`)
- Entry and tracking of CST statutory forms (C, E1, E2, F)
- Form number, type, and dates
- Linked to specific purchase bills
- Validity / expiry tracking

---

## Key Data Captured

| Field | Description |
|---|---|
| `BagNo`, `BagDt` | Bag sequential number and inward date |
| `ItemCode`, `ItemName` | Commodity / product identifier |
| `BagWt` | Individual bag weight |
| `TotWt`, `BalWt` | Total weight, balance (unused) weight |
| `SupplierRef`, `LotNo` | Supplier reference, lot/batch number |
| `LRNo`, `LRDt` | Lorry receipt number and date |
| `GPNo`, `GPDt` | Gate pass number and date |
| `GPType` | Inward / Outward |
| `GPParty` | Consignee / party for gate pass |
| `TransportMode` | Vehicle type / transporter |
| `VehicleNo` | Truck/vehicle number |
| `DeliveryParty` | Final delivery recipient |
| `DeliveryDt` | Delivery confirmation date |
| `IsGPOutward` | Direction flag: True = Outward |
| `GodownCode` | Storage location |

---

## Business Rules & Logic

- **Bag numbers:** Sequential per item per inward lot; unique within company
- **Weight reconciliation:** Balance weight = Total weight − Used weight (for partial deliveries)
- **Gate passes:** Mandatory for all goods movement (audit and legal trail)
- **Multiple bags per gate pass:** One gate pass covers a list of bags / cartons
- **Delivery tracking:** Confirms receiver identity and time; closes the despatch loop
- **Stock valuation:** Cost per bag calculated from purchase rate and total weight
- **Item shortfall:** Tracks losses, damaged goods, and quality issues at inward stage
- **SIT goods:** Goods under SIT arrangement tracked separately until transferred to company custody
- **Gate pass printing:** Pre-printed format required for truck/driver and warehouse records
- **WhatsApp gate pass:** PDF gate pass sent to consignee/party for advance information

---

## Reports Available

| Report | Description |
|---|---|
| `rptInwardSummCountwise.rpt` | Inward summary — unit/count wise |
| `rptInwardSummMillwise.rpt` | Inward summary — mill wise |
| `rptInwardSummTotal.rpt` | Inward total summary |
| Gate Pass Register (screen) | All gate passes issued — inward and outward |
| Stock Balance (frmRptStockReports) | Current inventory quantities by item |
| Goods in Transit | SIT goods not yet received into godown |
| Delivery Summary | All deliveries with receiver details |

---

## Integration Points

| Integrates With | How |
|---|---|
| Purchase | Goods inward linked to purchase bill by bill reference |
| Sales | Outward gate pass linked to sales bill for despatch |
| Tax Compliance | Form entry (C/E1/E2/F) linked to specific bills |
| Finance / Accounting | Inward creates stock asset entry; outward reduces stock |
| Communication | Gate pass PDF sent to party/consignee via WhatsApp |
| Reporting | Inward/outward summaries and stock balance reports |
