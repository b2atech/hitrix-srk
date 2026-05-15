---
id: business-flows
title: Business Process Flows
---

# Business Process Flows

> Derived from HITRIX VB6 form logic and Access DB structure. Voucher types in parentheses are the `VType` codes written to `tblVoucher`.

---

## Sales Cycle

```mermaid
flowchart TD
    A([Customer Order]) --> B{Sale Type?}
    B -->|Trade| C[Trade Sale Entry\nVType = SY]
    B -->|Consignment| D[Consignment Sale\nVType = SO]
    B -->|Depot| E[Depot Sale\nVType = SD]
    B -->|SIT| F[SIT Sale\nVType = ST]
    B -->|Mill Bill| G[Mill Bill\nVType = SM]

    C & D & E & F & G --> H[tblVoucher\nHeader created]
    H --> I[tblvousub\nItem lines — qty/wt/rate]
    H --> J[tblAddLess\nAdd/Less charges + GST]
    H --> K[tblOutStanding\nOpen receivable created]

    K --> L{GST applicable?}
    L -->|Yes| M[tblAddLess\nCGST/SGST/IGST amounts]
    L -->|IRN required| N[E-Invoice API call\ntbleInvoiceIRN — IRN + QR stored]

    K --> O{Confirmation needed?}
    O -->|Yes| P[tblConfirmation\nSales confirmation record]

    K --> Q[Receipt Entry\nVType = CR or BR]
    Q --> R[tblRecVsSale\nReceipt matched to invoice]
    R --> S{Fully settled?}
    S -->|No| T[Interest Debit Note\ntblRecVsSale.IntDbNtAmt]
    S -->|Yes| U([Outstanding Cleared])

    C --> V{Return?}
    V -->|Yes| W[Sales Return\nVType = RY]
    W --> X[Credit Note\nVType = SN]
```

---

## Purchase Cycle

```mermaid
flowchart TD
    A([Supplier Invoice]) --> B{Purchase Type?}
    B -->|Trade| C[Trade Purchase\nVType = PY]
    B -->|SIT| D[SIT Purchase\nVType = PT]
    B -->|Other| E[Other Purchase\nVType = PO]
    B -->|Inward| F[Purchase Inward\nVType = PI]

    C & D & E & F --> G[tblVoucher\nHeader created]
    G --> H[tblvousub\nItem lines — bags/wt/rate]
    G --> I[tblAddLess\nFreight, charges, GST]
    G --> J[tblOutStanding\nOpen payable created]

    H --> K[tblBags\nIndividual bag tracking\nBagNo + Wt + Cone]

    J --> L{TDS applicable?}
    L -->|Yes| M[tblTDS\nTDS deduction record\nBSR + challan details]

    J --> N[Payment Entry\nVType = CP or BP]
    N --> O{Mill related?}
    O -->|Yes| P[tblMillRecPay\nMill receipt/payment\nlinked to voucher]

    J --> Q{Return?}
    Q -->|Yes| R[Purchase Return\nVType = VY]
    R --> S[Debit Note\nVType = PN]

    G --> T{Gate Pass needed?}
    T -->|Yes| U[Gate Pass\nVType = GP]
    U --> V[tblGpSub\nGP line items\nitem/wt/godown/vehicle]
    V --> W[tblBags\nBags linked to GP]
```

---

## Finance / Accounting Cycle

```mermaid
flowchart TD
    A([Transaction]) --> B{Entry Type}

    B -->|Cash Payment| C[VType = CP\ntblVoucher]
    B -->|Bank Payment| D[VType = BP\ntblVoucher]
    B -->|Cash Receipt| E[VType = CR\ntblVoucher]
    B -->|Bank Receipt| F[VType = BR\ntblVoucher]
    B -->|Journal| G[VType = JV\ntblVoucher]

    C & D & E & F & G --> H{Outstanding linked?}
    H -->|Yes| I[tblRecVsSale\nReceipt vs Invoice match\nDiscount / TDS / Adjustment]
    H -->|No| J[Direct GL posting\nVAcCode / VAcOCode\nin tblVoucher]

    I --> K[tblOutStanding\nOutRecAmt updated]
    K --> L{Fully cleared?}
    L -->|No| M[Partial outstanding\nremains open]
    L -->|Yes| N([Outstanding Closed])

    G --> O{TDS entry?}
    O -->|Yes| P[tblTDS\nTDS challan details\nBSR / challan / period]

    J --> Q[Trial Balance\ntmpGenTbl staging]
    Q --> R[Crystal Reports\nP&L / Balance Sheet\ntmptblFin staging]

    B -->|Opening Balance| S[VType = OP\nYear-start balances]
```

---

## GST / Tax Flow

```mermaid
flowchart TD
    A([Invoice Created]) --> B{Transaction Type}
    B -->|Sale| C[VAcCode = customer\nVAcOCode = sales A/C]
    B -->|Purchase| D[VAcCode = supplier\nVAcOCode = purchase A/C]

    C --> E{Interstate?}
    E -->|No - Intrastate| F[CGST + SGST\nVCGSTRt/VAmt\nVSGSTRt/VAmt in tblVoucher]
    E -->|Yes - Interstate| G[IGST\nVIGSTRt/VIGSTAmt in tblVoucher]

    F & G --> H[tblAddLess\nAdTaxCode → tblMastNarration\nGST A/C codes resolved]

    H --> I{E-Invoice threshold?}
    I -->|Yes ≥ ₹5 Cr turnover| J[IRN Generation\nAPI call to IRP\ntbleInvoiceIRN stores\nIrnNo + AckNo + QrCode]
    I -->|No| K[Regular invoice\nno IRN needed]

    J --> L[QR Code\ntmptblQrCode staging\nfor print]

    H --> M[GST Return Staging\ntmpGenTbl aggregation\nby GSTIN + month]

    D --> N{RCM applicable?}
    N -->|Yes| O[tblMastNarration.RCMPayAc\nReverse Charge posting]

    H --> P{TCS applicable?}
    P -->|Yes| Q[tblAddLess\nAdTcsRt + AdTcsAmt\ntblMastSetting.AcCodeTcsPay/Rec]
```

---

## Inventory / Stock Flow

```mermaid
flowchart TD
    A([Goods Arrival]) --> B[Purchase Inward\nVType = PI\ntblVoucher + tblvousub]
    B --> C[tblBags\nBag-level registration\nBagNo + Wt + Cone\nLinked via InwNo/InwType]

    C --> D{Godown assignment}
    D --> E[VSubGodownCode\nin tblvousub]

    E --> F[Stock in Godown]

    F --> G{Movement?}
    G -->|Sale dispatch| H[Sale Entry\nVType = SY/SO/SD\ntblvousub — bags/wt deducted]
    G -->|Gate pass| I[Gate Pass\nVType = GP\ntblGpSub — item/wt/vehicle\ntblBags linked via GPNo/GpType]

    H --> J[Stock reduced\nfrom godown]
    I --> K[tblGpSub\nDelivery date + vehicle\nSizer reference]

    K --> L[tblBags updated\nGPNo + GpType linked]

    F --> M[Stock Report\ntmpStockCal staging\nCrystal Reports]
    M --> N[Opening + Inward\n- Sales - Gate Pass\n= Closing Stock]
```
