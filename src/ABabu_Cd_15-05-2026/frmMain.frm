VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000A&
   ClientHeight    =   8805
   ClientLeft      =   165
   ClientTop       =   -6435
   ClientWidth     =   18795
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   Moveable        =   0   'False
   Picture         =   "frmMain.frx":058A
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   8505
      Width           =   18795
      _ExtentX        =   33152
      _ExtentY        =   529
      SimpleText      =   "s"
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Bevel           =   0
            Object.Width           =   5292
            MinWidth        =   5292
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Bevel           =   0
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   1
            Enabled         =   0   'False
            TextSave        =   "CAPS"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin Crystal.CrystalReport cryReport 
      Left            =   4950
      Top             =   2970
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Menu mnuAdmin 
      Caption         =   "Admin"
      Index           =   0
      Begin VB.Menu mnuAdminfrm 
         Caption         =   "&User Creation"
         Index           =   0
      End
   End
   Begin VB.Menu mnuMst 
      Caption         =   "&Master "
      Index           =   0
      Begin VB.Menu mnuMstfrm 
         Caption         =   "&Account"
         Index           =   0
      End
      Begin VB.Menu mnuMstfrm 
         Caption         =   "&Group"
         Index           =   1
      End
      Begin VB.Menu mnuMstfrm 
         Caption         =   "&Party(Weaver)Group"
         Index           =   2
      End
      Begin VB.Menu mnuMstfrm 
         Caption         =   "&Count"
         Index           =   3
      End
      Begin VB.Menu mnuMstfrm 
         Caption         =   "&Godown"
         Index           =   4
      End
      Begin VB.Menu mnuMstfrm 
         Caption         =   "&Narration"
         Index           =   5
      End
      Begin VB.Menu mnuMstfrm 
         Caption         =   "&Tax"
         Index           =   6
      End
      Begin VB.Menu mnuMstfrm 
         Caption         =   "&Sub Group Master"
         Index           =   7
      End
      Begin VB.Menu mnuMstfrm 
         Caption         =   "T&ransport"
         Index           =   8
      End
      Begin VB.Menu mnuMstfrm 
         Caption         =   "Party Bank Master"
         Index           =   9
      End
      Begin VB.Menu mnuMstfrm 
         Caption         =   "Consignee Master"
         Index           =   10
      End
      Begin VB.Menu mnuMstfrm 
         Caption         =   "Other Item"
         Index           =   11
      End
      Begin VB.Menu mnuMstfrm 
         Caption         =   "Delivery Address"
         Index           =   12
      End
      Begin VB.Menu mnuMstfrm 
         Caption         =   "Users Working Firms"
         Index           =   13
      End
      Begin VB.Menu mnuMstfrm 
         Caption         =   "Bill Serial Master"
         Index           =   14
      End
      Begin VB.Menu mnuMstfrm 
         Caption         =   "Opening Stock "
         Index           =   15
         Begin VB.Menu mnuMstfrmSubOpStk 
            Caption         =   "Trade"
            Index           =   0
         End
         Begin VB.Menu mnuMstfrmSubOpStk 
            Caption         =   "Goods Inward (Consignment / Depot)"
            Index           =   1
         End
         Begin VB.Menu mnuMstfrmSubOpStk 
            Caption         =   "Exemption"
            Index           =   2
         End
         Begin VB.Menu mnuMstfrmSubOpStk 
            Caption         =   "Sale In Transit"
            Index           =   3
         End
      End
      Begin VB.Menu mnuMstfrm 
         Caption         =   "Op. Balance Interest "
         Index           =   16
      End
      Begin VB.Menu mnuMstSetting 
         Caption         =   "&Setting"
         Index           =   0
      End
      Begin VB.Menu mnuMstHSNUpdate 
         Caption         =   "Update HSN Code"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMstFirmIntRt 
         Caption         =   "Firmwise Interest Rate"
         Index           =   0
      End
   End
   Begin VB.Menu mnutrnGst 
      Caption         =   "Transaction"
      Index           =   0
      Begin VB.Menu mnutrnGstSubGSTBK 
         Caption         =   "&Booking And Despatch Details"
         Index           =   0
         Begin VB.Menu mnutrnGstSubBKGSTfrm 
            Caption         =   "&Trade Purchase Booking"
            Index           =   0
         End
         Begin VB.Menu mnutrnGstSubBKGSTfrm 
            Caption         =   "&Mill Booking ( Direct Sale )"
            Index           =   1
         End
         Begin VB.Menu mnutrnGstSubBKGSTfrm 
            Caption         =   "Despatch Details Mill Booking"
            Index           =   2
         End
         Begin VB.Menu mnutrnGstSubBKGSTfrm 
            Caption         =   "Despatch Details Purchase Booking"
            Index           =   3
         End
         Begin VB.Menu mnutrnGstSubBKGSTfrm 
            Caption         =   "Party Booking ( Trade Sale )"
            Index           =   4
         End
         Begin VB.Menu mnutrnGstSubBKGSTfrm 
            Caption         =   "Despatch Details Depot"
            Index           =   5
         End
      End
      Begin VB.Menu mnutrnGstSub1 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnutrnGstSubGPfrm 
         Caption         =   "&Gate Pass"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnutrnGstSub2 
         Caption         =   "-"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnutrnGstSubSale 
         Caption         =   "&Sales"
         Index           =   0
         Begin VB.Menu mnutrnGstSubSalefrm 
            Caption         =   "&Trade Sales"
            Index           =   0
         End
         Begin VB.Menu mnutrnGstSubSalefrm 
            Caption         =   "&Depot Sales"
            Index           =   1
         End
         Begin VB.Menu mnutrnGstSubSalefrm 
            Caption         =   "&Other Sale"
            Index           =   2
         End
         Begin VB.Menu mnutrnGstSubSalefrm 
            Caption         =   "&Commission And Brokerage Sales"
            Index           =   3
         End
         Begin VB.Menu mnutrnGstSubSalefrm 
            Caption         =   "&Mill Billing"
            Index           =   4
         End
         Begin VB.Menu mnutrnGstSubSalefrm 
            Caption         =   "-"
            Index           =   5
         End
         Begin VB.Menu mnutrnGstSubSalefrm 
            Caption         =   "Late Payment Interest Sales"
            Index           =   6
         End
         Begin VB.Menu mnutrnGstSubSalefrm 
            Caption         =   "Late Payment Interest Sales (Without GST)"
            Index           =   7
         End
         Begin VB.Menu mnutrnGstSubSalefrm 
            Caption         =   "TCS Debit Note"
            Index           =   8
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnutrnGstSubPur 
         Caption         =   "&Purchase"
         Index           =   0
         Begin VB.Menu mnutrnGstSubPurfrm 
            Caption         =   "&Trade Purchase"
            Index           =   0
         End
         Begin VB.Menu mnutrnGstSubPurfrm 
            Caption         =   "&Other Purchase"
            Index           =   1
         End
      End
      Begin VB.Menu mnutrnGstSubInwfrm 
         Caption         =   "Goods &Inward (  Depot )"
         Index           =   0
      End
      Begin VB.Menu mnutrnGstSub3 
         Caption         =   "Goods &Outward (  Depot )"
         Index           =   0
      End
      Begin VB.Menu mnutrnGstSubPay 
         Caption         =   "Pa&yment"
         Index           =   0
         Begin VB.Menu mnutrnGstSubPayfrm 
            Caption         =   "&Bank"
            Index           =   0
         End
         Begin VB.Menu mnutrnGstSubPayfrm 
            Caption         =   "&Cash"
            Index           =   1
         End
      End
      Begin VB.Menu mnutrnGstSubRec 
         Caption         =   "&Receipt"
         Index           =   0
         Begin VB.Menu mnutrnGstSubRecfrm 
            Caption         =   "&Bank"
            Index           =   0
         End
         Begin VB.Menu mnutrnGstSubRecfrm 
            Caption         =   "&Cash"
            Index           =   1
         End
      End
      Begin VB.Menu mnutrnGstSub4 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnutrnGstSubJVfrm 
         Caption         =   "&Journal Entry"
         Index           =   0
      End
      Begin VB.Menu mnutrnGstSubCrNt 
         Caption         =   "&Credit Note "
         Index           =   0
         Begin VB.Menu mnutrnGstSubCrNtfrm 
            Caption         =   "Credit Note With GST( Sales )"
            Index           =   0
         End
         Begin VB.Menu mnutrnGstSubCrNtfrm 
            Caption         =   "Credit Note ( Sales )"
            Index           =   1
         End
         Begin VB.Menu mnutrnGstSubCrNtfrm 
            Caption         =   "Credit Note With GST( Purchase )"
            Index           =   2
         End
         Begin VB.Menu mnutrnGstSubCrNtfrm 
            Caption         =   "Credit Note ( Purchase )"
            Index           =   3
         End
      End
      Begin VB.Menu mnutrnGstSubDrNt 
         Caption         =   "&Debit Note"
         Index           =   0
         Begin VB.Menu mnutrnGstSubDrNtfrm 
            Caption         =   "Debit Note With GST( Sales )"
            Index           =   0
         End
         Begin VB.Menu mnutrnGstSubDrNtfrm 
            Caption         =   "Debit Note ( Sales )"
            Index           =   1
         End
         Begin VB.Menu mnutrnGstSubDrNtfrm 
            Caption         =   "Debit Note With GST( Purchase )"
            Index           =   2
         End
         Begin VB.Menu mnutrnGstSubDrNtfrm 
            Caption         =   "Debit Note ( Purchase )"
            Index           =   3
         End
      End
      Begin VB.Menu mnutrnGstSub5 
         Caption         =   "J.V. Against  Invoice "
         Index           =   0
      End
      Begin VB.Menu mnutrnGstSubSlRtn 
         Caption         =   "Sales &Return"
         Index           =   0
      End
      Begin VB.Menu mnutrnGstSubPurRtn 
         Caption         =   "P&urchase Return"
         Index           =   0
      End
      Begin VB.Menu mnutrnGstSub6 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnutrnGstSubLpfrm 
         Caption         =   "Late Payment Demand Note List"
         Index           =   0
      End
      Begin VB.Menu mnutrnGstSubEpfrm 
         Caption         =   "Early Payment  Credit Note List"
         Index           =   0
      End
      Begin VB.Menu mnutrnGstSubLpCashfrm 
         Caption         =   "Late Payment Demand Note List (From Cash Receipt)"
         Index           =   0
      End
      Begin VB.Menu mnutrnGstSubDailyPricefrm 
         Caption         =   "Daily Price List "
         Index           =   0
      End
      Begin VB.Menu mnutrnGstSubMillPayment 
         Caption         =   "Mill Payment"
         Index           =   0
         Begin VB.Menu mnutrnGstSubMillPaymentfrm 
            Caption         =   "Against Mill Bill"
            Index           =   0
         End
         Begin VB.Menu mnutrnGstSubMillPaymentfrm 
            Caption         =   "Against Trade Purchase"
            Index           =   1
         End
         Begin VB.Menu mnutrnGstSubMillPaymentfrm 
            Caption         =   "Against Depot Bill"
            Index           =   2
         End
         Begin VB.Menu mnutrnGstSubMillPaymentfrm 
            Caption         =   "Against Commission-Brokerage Sales"
            Index           =   3
         End
      End
      Begin VB.Menu mnutrnGstSubOsAdj 
         Caption         =   "Outstanding Adjustment"
         Index           =   0
      End
      Begin VB.Menu mnutrnGstSubJVWithTDS 
         Caption         =   "J V With TDS And Brokerage On Sales"
         Index           =   0
      End
   End
   Begin VB.Menu mnuRptR 
      Caption         =   "&Reports"
      Index           =   0
      Begin VB.Menu mnuRptRSub 
         Caption         =   "&Financial Reports"
         Index           =   0
         Begin VB.Menu mnuRptRSubFin 
            Caption         =   "&Balance Sheet"
            Index           =   0
         End
         Begin VB.Menu mnuRptRSubFin 
            Caption         =   "&Trial Balance"
            Index           =   1
            Begin VB.Menu mnuRptRSubFinTri 
               Caption         =   "&Simple"
               Index           =   0
               Visible         =   0   'False
            End
            Begin VB.Menu mnuRptRSubFinTri 
               Caption         =   "&Detail"
               Index           =   1
               Visible         =   0   'False
            End
            Begin VB.Menu mnuRptRSubFinTri 
               Caption         =   "&GroupWise"
               Index           =   2
            End
            Begin VB.Menu mnuRptRSubFinTri 
               Caption         =   "GroupWise Detail"
               Index           =   3
            End
            Begin VB.Menu mnuRptRSubFinTri 
               Caption         =   "&T Format"
               Index           =   4
               Visible         =   0   'False
            End
         End
         Begin VB.Menu mnuRptRSubFin 
            Caption         =   "&Profit & Loss A/C"
            Index           =   2
         End
         Begin VB.Menu mnuRptRSubFin 
            Caption         =   "&Shedule Printing"
            Index           =   3
         End
         Begin VB.Menu mnuRptRSubFin 
            Caption         =   "Opposite A/c wise Summery"
            Index           =   4
         End
         Begin VB.Menu mnuRptRSubFin 
            Caption         =   "Expenses Sub A/c Details"
            Index           =   5
         End
         Begin VB.Menu mnuRptRSubFin 
            Caption         =   "Loan A/c Details"
            Index           =   6
         End
         Begin VB.Menu mnuRptRSubFin 
            Caption         =   "Asset A/c Details"
            Index           =   7
         End
         Begin VB.Menu mnuRptRSubFin 
            Caption         =   "Depreciation Chart"
            Index           =   8
         End
      End
      Begin VB.Menu mnuRptRSub 
         Caption         =   "&Account Ledger"
         Index           =   1
         Begin VB.Menu mnuRptRSubAccLeg 
            Caption         =   "&Accountwise"
            Index           =   0
         End
         Begin VB.Menu mnuRptRSubAccLeg 
            Caption         =   "&Groupwise"
            Index           =   1
         End
         Begin VB.Menu mnuRptRSubAccLeg 
            Caption         =   "Account &Confirmation"
            Index           =   2
         End
         Begin VB.Menu mnuRptRSubAccLeg 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu mnuRptRSubAccLeg 
            Caption         =   "Broker Ledger"
            Index           =   4
         End
         Begin VB.Menu mnuRptRSubAccLeg 
            Caption         =   "&Bank Reconciliation"
            Index           =   5
         End
         Begin VB.Menu mnuRptRSubAccLeg 
            Caption         =   "&Paid In Next Year"
            Index           =   6
         End
      End
      Begin VB.Menu mnuRptRSub 
         Caption         =   "&Books Printing"
         Index           =   2
         Begin VB.Menu mnurptRSubBookPri 
            Caption         =   "&Cash Book"
            Index           =   0
         End
         Begin VB.Menu mnurptRSubBookPri 
            Caption         =   "&Bank Book"
            Index           =   1
         End
         Begin VB.Menu mnurptRSubBookPri 
            Caption         =   "Credit Cash Report"
            Index           =   2
         End
      End
      Begin VB.Menu mnuRptRSub 
         Caption         =   "&Register Printing"
         Index           =   3
         Begin VB.Menu mnurptRSubRegPri 
            Caption         =   "&Sales Register"
            Index           =   0
            Begin VB.Menu mnurptRSubRegPriSale 
               Caption         =   "&Trade Sales"
               Index           =   0
            End
            Begin VB.Menu mnurptRSubRegPriSale 
               Caption         =   "&Other Sales"
               Index           =   1
            End
            Begin VB.Menu mnurptRSubRegPriSale 
               Caption         =   "&Consign/Depot"
               Index           =   2
            End
            Begin VB.Menu mnurptRSubRegPriSale 
               Caption         =   "&Mill Bill"
               Index           =   3
            End
            Begin VB.Menu mnurptRSubRegPriSale 
               Caption         =   "Commission And Brokerage Sales"
               Index           =   4
            End
         End
         Begin VB.Menu mnurptRSubRegPri 
            Caption         =   "&Purchase Register"
            Index           =   1
            Begin VB.Menu mnurptRSubRegPriPur 
               Caption         =   "&Trade"
               Index           =   0
            End
            Begin VB.Menu mnurptRSubRegPriPur 
               Caption         =   "&SIT"
               Index           =   1
            End
            Begin VB.Menu mnurptRSubRegPriPur 
               Caption         =   "&Other"
               Index           =   2
            End
            Begin VB.Menu mnurptRSubRegPriPur 
               Caption         =   "&Consignment/Depot Inward"
               Index           =   3
            End
         End
         Begin VB.Menu mnurptRSubRegPri 
            Caption         =   "&Insurance Report - Purchase"
            Index           =   2
         End
         Begin VB.Menu mnurptRSubRegPri 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu mnurptRSubRegPri 
            Caption         =   "&Credit Note (Sales) Register"
            Index           =   4
         End
         Begin VB.Menu mnurptRSubRegPri 
            Caption         =   "&Debit Note (Sales) Register"
            Index           =   5
         End
         Begin VB.Menu mnurptRSubRegPri 
            Caption         =   "&Internal Voucher Register"
            Index           =   6
         End
         Begin VB.Menu mnurptRSubRegPri 
            Caption         =   "Credit Note ( Purchase ) Register"
            Index           =   7
         End
         Begin VB.Menu mnurptRSubRegPri 
            Caption         =   "Debit Note ( Purchase ) Register"
            Index           =   8
         End
         Begin VB.Menu mnurptRSubRegPri 
            Caption         =   "&Sales Return"
            Index           =   9
         End
         Begin VB.Menu mnurptRSubRegPri 
            Caption         =   "&Purchase Return"
            Index           =   10
         End
         Begin VB.Menu mnurptRSubRegPri 
            Caption         =   "-"
            Index           =   11
         End
         Begin VB.Menu mnurptRSubRegPri 
            Caption         =   "&Late Payment Debit Note Register"
            Index           =   12
         End
         Begin VB.Menu mnurptRSubRegPri 
            Caption         =   "Late Payment Debit Note Register (Mill Bill)"
            Index           =   13
         End
      End
      Begin VB.Menu mnuRptRSub 
         Caption         =   "&OutStanding Reports"
         Index           =   4
         Begin VB.Menu mnurptRSubOutSta 
            Caption         =   "&Sales"
            Index           =   0
            Begin VB.Menu mnurptRSubOutStaSale 
               Caption         =   "&Brokerwise"
               Index           =   0
            End
            Begin VB.Menu mnurptRSubOutStaSale 
               Caption         =   "Partywise"
               Index           =   1
            End
            Begin VB.Menu mnurptRSubOutStaSale 
               Caption         =   "Invoicewise"
               Index           =   2
            End
            Begin VB.Menu mnurptRSubOutStaSale 
               Caption         =   "-"
               Index           =   3
            End
            Begin VB.Menu mnurptRSubOutStaSale 
               Caption         =   "&Late Payment Interest List"
               Index           =   4
            End
            Begin VB.Menu mnurptRSubOutStaSale 
               Caption         =   "Ledger And Oustanding Sale Bal Check"
               Index           =   5
            End
            Begin VB.Menu mnurptRSubOutStaSale 
               Caption         =   "&Early Payment Interest List"
               Index           =   6
            End
         End
         Begin VB.Menu mnurptRSubOutSta 
            Caption         =   "&Purchase"
            Index           =   1
            Begin VB.Menu mnurptRSubOutStaPur 
               Caption         =   "&Partywise"
               Index           =   0
            End
            Begin VB.Menu mnurptRSubOutStaPur 
               Caption         =   "&Billwise"
               Index           =   1
            End
            Begin VB.Menu mnurptRSubOutStaPur 
               Caption         =   "&Unadjusted Payments"
               Index           =   2
            End
            Begin VB.Menu mnurptRSubOutStaPur 
               Caption         =   "&Mill Bill Mill O/s"
               Index           =   3
            End
         End
         Begin VB.Menu mnurptRSubOutSta 
            Caption         =   "TDS Receivable O/s"
            Index           =   2
         End
         Begin VB.Menu mnurptRSubOutSta 
            Caption         =   "TCS Receivable O/s"
            Index           =   3
         End
         Begin VB.Menu mnurptRSubOutSta 
            Caption         =   "TDS On Sales Receivable O/s"
            Index           =   4
         End
         Begin VB.Menu mnurptRSubOutSta 
            Caption         =   "TDS On Sales ( Mill Bill )"
            Index           =   5
         End
         Begin VB.Menu mnurptRSubOutSta 
            Caption         =   "TDS On Sales ( Depot )"
            Index           =   6
         End
      End
      Begin VB.Menu mnuRptRSub 
         Caption         =   "&Stock Reports"
         Index           =   5
         Begin VB.Menu mnurptRSubStock 
            Caption         =   "Stock Statement"
            Index           =   0
         End
         Begin VB.Menu mnurptRSubStock 
            Caption         =   "Daily Stock Statement"
            Index           =   1
         End
         Begin VB.Menu mnurptRSubStock 
            Caption         =   "Sales Summery"
            Index           =   2
         End
         Begin VB.Menu mnurptRSubStock 
            Caption         =   "Gate Pass Wise Sales"
            Index           =   3
         End
         Begin VB.Menu mnurptRSubStock 
            Caption         =   "Stock Statement Godownwsie"
            Index           =   4
         End
         Begin VB.Menu mnurptRSubStock 
            Caption         =   "-"
            Index           =   5
            Visible         =   0   'False
         End
         Begin VB.Menu mnurptRSubStock 
            Caption         =   "&Purchase/Inwardwise Stock"
            Index           =   6
            Visible         =   0   'False
         End
         Begin VB.Menu mnurptRSubStock 
            Caption         =   "Purchase/Inwardwise &Outward"
            Index           =   7
            Visible         =   0   'False
         End
         Begin VB.Menu mnurptRSubStock 
            Caption         =   "&Pending Purchase SIT"
            Index           =   8
            Visible         =   0   'False
         End
         Begin VB.Menu mnurptRSubStock 
            Caption         =   "-"
            Index           =   9
            Visible         =   0   'False
         End
         Begin VB.Menu mnurptRSubStock 
            Caption         =   "Stock &Valuation (Trade,SIT)"
            Index           =   10
            Visible         =   0   'False
            Begin VB.Menu mnurptRSubStockVal 
               Caption         =   "&Register"
               Index           =   0
            End
            Begin VB.Menu mnurptRSubStockVal 
               Caption         =   "&Summary"
               Index           =   1
            End
         End
         Begin VB.Menu mnurptRSubTax 
            Caption         =   "&TDS "
            Index           =   10
            Visible         =   0   'False
         End
         Begin VB.Menu mnurptRSubTax 
            Caption         =   "TDS &Summary"
            Index           =   11
            Visible         =   0   'False
         End
         Begin VB.Menu mnurptRSubTax 
            Caption         =   "Service Tax"
            Index           =   12
            Visible         =   0   'False
         End
         Begin VB.Menu mnurptRSubTax 
            Caption         =   "VAT Audit Reports"
            Index           =   13
            Visible         =   0   'False
            Begin VB.Menu mnurptRSubTaxAuditForm 
               Caption         =   "Annexure - G  (Received C And E1 Form)"
               Index           =   0
            End
            Begin VB.Menu mnurptRSubTaxAuditForm 
               Caption         =   "Annexure - I  (Pending  C And E1 Form)"
               Index           =   1
            End
            Begin VB.Menu mnurptRSubTaxAuditForm 
               Caption         =   "Annexure - J Section -1   (All Partywise VAT Sales)"
               Index           =   2
            End
            Begin VB.Menu mnurptRSubTaxAuditForm 
               Caption         =   "Annexure - J Section -1   (R.D. Partywise VAT Sales)"
               Index           =   3
            End
            Begin VB.Menu mnurptRSubTaxAuditForm 
               Caption         =   "Annexure - J Section -1   (U.R.D. Partywise VAT Sales)"
               Index           =   4
            End
            Begin VB.Menu mnurptRSubTaxAuditForm 
               Caption         =   "Annexure - J Section -2   (Partywise VAT Purchases)"
               Index           =   5
            End
            Begin VB.Menu mnurptRSubTaxAuditForm 
               Caption         =   "Annexure - J Section -6   (Supllierwise CST Purchases)"
               Index           =   6
            End
            Begin VB.Menu mnurptRSubTaxAuditForm 
               Caption         =   "-"
               Index           =   7
            End
            Begin VB.Menu mnurptRSubTaxAuditForm 
               Caption         =   "Partywise Sales Details"
               Index           =   8
            End
            Begin VB.Menu mnurptRSubTaxAuditForm 
               Caption         =   "Supplierwise Purchase Details"
               Index           =   9
            End
            Begin VB.Menu mnurptRSubTaxAuditForm 
               Caption         =   "Monthly Sales Details"
               Index           =   10
            End
            Begin VB.Menu mnurptRSubTaxAuditForm 
               Caption         =   "Monthly Purchase Details"
               Index           =   11
            End
            Begin VB.Menu mnurptRSubTaxAuditForm 
               Caption         =   "Pending C Form (Receivable)"
               Index           =   12
            End
            Begin VB.Menu mnurptRSubTaxAuditForm 
               Caption         =   "Pending E1 Form (Receivable)"
               Index           =   13
            End
            Begin VB.Menu mnurptRSubTaxAuditForm 
               Caption         =   "Received C Form "
               Index           =   14
            End
            Begin VB.Menu mnurptRSubTaxAuditForm 
               Caption         =   "Received E1 Form "
               Index           =   15
            End
         End
      End
      Begin VB.Menu mnuRptRSub 
         Caption         =   "Party Interest Calculation"
         Index           =   8
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRptRSub 
         Caption         =   "G S T Reports"
         Index           =   9
         Begin VB.Menu mnuRptRSubGstRp 
            Caption         =   "Trade Sale Register"
            Index           =   0
         End
         Begin VB.Menu mnuRptRSubGstRp 
            Caption         =   "Depot Sale Register"
            Index           =   1
         End
         Begin VB.Menu mnuRptRSubGstRp 
            Caption         =   "Commission And Brokerage Sales"
            Index           =   2
         End
         Begin VB.Menu mnuRptRSubGstRp 
            Caption         =   "Late Payment Interest Sales"
            Index           =   3
         End
         Begin VB.Menu mnuRptRSubGstRp 
            Caption         =   "Late Payment Interest Sales ( Without GST )"
            Index           =   4
         End
         Begin VB.Menu mnuRptRSubGstRp 
            Caption         =   "Non-GST Outward Sales"
            Index           =   5
         End
         Begin VB.Menu mnuRptRSubGstRp 
            Caption         =   "R C M Sale Register"
            Index           =   6
         End
         Begin VB.Menu mnuRptRSubGstRp 
            Caption         =   "-"
            Index           =   7
         End
         Begin VB.Menu mnuRptRSubGstRp 
            Caption         =   "Credit Note (Sales) Register"
            Index           =   8
         End
         Begin VB.Menu mnuRptRSubGstRp 
            Caption         =   "Debit  Note (Sales) Register"
            Index           =   9
         End
         Begin VB.Menu mnuRptRSubGstRp 
            Caption         =   "Credit Note (Purchase) Register"
            Index           =   10
         End
         Begin VB.Menu mnuRptRSubGstRp 
            Caption         =   "Debit  Note (Purchase) Register"
            Index           =   11
         End
         Begin VB.Menu mnuRptRSubGstRp 
            Caption         =   "-"
            Index           =   12
         End
         Begin VB.Menu mnuRptRSubGstRp 
            Caption         =   "Trade Purchase Register"
            Index           =   13
         End
         Begin VB.Menu mnuRptRSubGstRp 
            Caption         =   "Other Purchase Register"
            Index           =   14
         End
         Begin VB.Menu mnuRptRSubGstRp 
            Caption         =   "-"
            Index           =   15
         End
         Begin VB.Menu mnuRptRSubGstRp 
            Caption         =   "GSTR 3B"
            Index           =   16
         End
         Begin VB.Menu mnuRptRSubGstRp 
            Caption         =   "GSTR - 1"
            Index           =   17
         End
         Begin VB.Menu mnuRptRSubGstRp 
            Caption         =   "Monthwise GST Summery"
            Index           =   18
         End
         Begin VB.Menu mnuRptRSubGstRp 
            Caption         =   "GSTR - 2A  -  B2B"
            Index           =   19
         End
         Begin VB.Menu mnuRptRSubGstRp 
            Caption         =   "GSTR - 2A  -  B2B  (Mismatch)"
            Index           =   20
         End
         Begin VB.Menu mnuRptRSubGstRp 
            Caption         =   "-"
            Index           =   21
         End
         Begin VB.Menu mnuRptRSubGstRp 
            Caption         =   "Sales Return Register"
            Index           =   22
         End
         Begin VB.Menu mnuRptRSubGstRp 
            Caption         =   "Purchase Return Register"
            Index           =   23
         End
         Begin VB.Menu mnuRptRSubGstRp 
            Caption         =   "GSTR - 9"
            Index           =   24
         End
      End
      Begin VB.Menu mnuRptRSub 
         Caption         =   "TCS Return "
         Index           =   10
      End
      Begin VB.Menu mnuRptRSub 
         Caption         =   "TCS Bills Check"
         Index           =   11
      End
      Begin VB.Menu mnuRptRSub 
         Caption         =   "Whatsapp Message"
         Index           =   12
         Begin VB.Menu mnuRptRSubWapp 
            Caption         =   "Outstanding To Broker"
            Index           =   0
         End
         Begin VB.Menu mnuRptRSubWapp 
            Caption         =   "Outstanding To Party"
            Index           =   1
         End
      End
      Begin VB.Menu mnuRptRSub 
         Caption         =   "XXXXXXXXXXXXXX"
         Index           =   13
      End
      Begin VB.Menu mnuRptRSub 
         Caption         =   "Audit Trail Reports"
         Index           =   14
      End
   End
   Begin VB.Menu mnuRptGenr 
      Caption         =   "General Reports"
      Index           =   0
      Begin VB.Menu mnuRptGenrSub 
         Caption         =   "List"
         Index           =   0
         Begin VB.Menu mnuRptGenrSubLst 
            Caption         =   "&Accounts List"
            Index           =   0
         End
         Begin VB.Menu mnuRptGenrSubLst 
            Caption         =   "&Groupwise Accounts List"
            Index           =   1
         End
         Begin VB.Menu mnuRptGenrSubLst 
            Caption         =   "Selected Group Accounts List"
            Index           =   2
         End
         Begin VB.Menu mnuRptGenrSubLst 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu mnuRptGenrSubLst 
            Caption         =   "&Count Rate List"
            Index           =   4
         End
         Begin VB.Menu mnuRptGenrSubLst 
            Caption         =   "&Millwise Count List"
            Index           =   5
         End
         Begin VB.Menu mnuRptGenrSubLst 
            Caption         =   "Count &Brokerage List"
            Index           =   6
         End
         Begin VB.Menu mnuRptGenrSubLst 
            Caption         =   "Selected Group No Transaction A/c List"
            Index           =   7
         End
         Begin VB.Menu mnuRptGenrSubLst 
            Caption         =   "Party Groupwise List"
            Index           =   8
         End
      End
      Begin VB.Menu mnuRptGenrSub 
         Caption         =   "&Opening Balance List"
         Index           =   1
         Begin VB.Menu mnuRptGenrSubOPBal 
            Caption         =   "&Accoutwise"
            Index           =   0
         End
         Begin VB.Menu mnuRptGenrSubOPBal 
            Caption         =   "&Groupwise"
            Index           =   1
         End
      End
      Begin VB.Menu mnuRptGenrSub 
         Caption         =   "&Closing Balance List"
         Index           =   2
         Begin VB.Menu mnuRptGenrSubCLBal 
            Caption         =   "&Accoutwise"
            Index           =   0
         End
         Begin VB.Menu mnuRptGenrSubCLBal 
            Caption         =   "&Groupwise"
            Index           =   1
         End
      End
      Begin VB.Menu mnuRptGenrSub 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuRptGenrSub 
         Caption         =   "&Monthly Mill Statement"
         Index           =   4
         Begin VB.Menu mnuRptGenrSubMon 
            Caption         =   "&Consignment/Depot"
            Index           =   0
         End
         Begin VB.Menu mnuRptGenrSubMon 
            Caption         =   "&Mill Billing Statement"
            Index           =   1
         End
         Begin VB.Menu mnuRptGenrSubMon 
            Caption         =   "&Sales In Transit Statement"
            Index           =   2
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnuRptGenrSub 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuRptGenrSub 
         Caption         =   "&Sales Reports"
         Index           =   6
         Begin VB.Menu mnuRptGenrSubSale 
            Caption         =   "&Total Sales"
            Index           =   0
         End
         Begin VB.Menu mnuRptGenrSubSale 
            Caption         =   "&Count / Brokerwise Sales"
            Index           =   1
         End
         Begin VB.Menu mnuRptGenrSubSale 
            Caption         =   "&Party / Countwise Sales"
            Index           =   2
         End
         Begin VB.Menu mnuRptGenrSubSale 
            Caption         =   "&Sales Against Booking"
            Index           =   3
         End
      End
      Begin VB.Menu mnuRptGenrSub 
         Caption         =   "&Purchase Reports"
         Index           =   7
         Begin VB.Menu mnuRptGenrSubPur 
            Caption         =   "&Trade Purchase TDS Report"
            Index           =   0
         End
         Begin VB.Menu mnuRptGenrSubPur 
            Caption         =   "&Depot Sales TDS Report"
            Index           =   1
         End
         Begin VB.Menu mnuRptGenrSubPur 
            Caption         =   "&Mill Bill TDS Report"
            Index           =   2
         End
         Begin VB.Menu mnuRptGenrSubPur 
            Caption         =   "Millwise Inward"
            Index           =   3
         End
         Begin VB.Menu mnuRptGenrSubPur 
            Caption         =   "Daily Purchase Report"
            Index           =   4
         End
         Begin VB.Menu mnuRptGenrSubPur 
            Caption         =   "Daily Purchase Report ( Mill Bill )"
            Index           =   5
         End
         Begin VB.Menu mnuRptGenrSubPur 
            Caption         =   "&Other Purchase TDS Report"
            Index           =   6
         End
      End
      Begin VB.Menu mnuRptGenrSub 
         Caption         =   "&Brokerage Reports"
         Index           =   8
         Begin VB.Menu mnuRptGenrSubBro 
            Caption         =   "&Broker Ledger"
            Index           =   0
         End
         Begin VB.Menu mnuRptGenrSubBro 
            Caption         =   "Brokerage Detail "
            Index           =   1
         End
         Begin VB.Menu mnuRptGenrSubBro 
            Caption         =   "Brokerage Summery"
            Index           =   2
         End
         Begin VB.Menu mnuRptGenrSubBro 
            Caption         =   "Brokerage Voucher"
            Index           =   3
         End
         Begin VB.Menu mnuRptGenrSubBro 
            Caption         =   "Countwise Brokarage "
            Index           =   4
         End
         Begin VB.Menu mnuRptGenrSubBro 
            Caption         =   "Inward/Purchase Brokarage Detail"
            Index           =   5
         End
         Begin VB.Menu mnuRptGenrSubBro 
            Caption         =   "Inward/Purchase Brokarage Total"
            Index           =   6
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnuRptGenrSub 
         Caption         =   "&Booking Reports"
         Index           =   9
         Begin VB.Menu mnuRptGenrSubBk 
            Caption         =   "&Register"
            Index           =   0
            Begin VB.Menu mnuRptGenrSubBkReg 
               Caption         =   "&Mill Booking ( Direct Sale )"
               Index           =   0
            End
            Begin VB.Menu mnuRptGenrSubBkReg 
               Caption         =   "&Trade Purchase Booking"
               Index           =   1
            End
            Begin VB.Menu mnuRptGenrSubBkReg 
               Caption         =   "&Consignment/Depot Booking"
               Index           =   2
               Visible         =   0   'False
            End
            Begin VB.Menu mnuRptGenrSubBkReg 
               Caption         =   "&Party Booking"
               Index           =   3
            End
            Begin VB.Menu mnuRptGenrSubBkReg 
               Caption         =   "Party Booking Summery"
               Index           =   4
            End
            Begin VB.Menu mnuRptGenrSubBkReg 
               Caption         =   "Purchase Booking Summery"
               Index           =   5
            End
         End
         Begin VB.Menu mnuRptGenrSubBk 
            Caption         =   "&Pending"
            Index           =   1
            Begin VB.Menu mnuRptGenrSubBkStk 
               Caption         =   "&Mill Booking"
               Index           =   0
            End
            Begin VB.Menu mnuRptGenrSubBkStk 
               Caption         =   "&Trade Purchase Booking"
               Index           =   1
            End
            Begin VB.Menu mnuRptGenrSubBkStk 
               Caption         =   "&Consignment/Depot Booking"
               Index           =   2
               Visible         =   0   'False
            End
            Begin VB.Menu mnuRptGenrSubBkStk 
               Caption         =   "&Party Booking"
               Index           =   3
            End
            Begin VB.Menu mnuRptGenrSubBkStk 
               Caption         =   "&Despatch Goods Details"
               Index           =   4
            End
         End
         Begin VB.Menu mnuRptGenrSubBk 
            Caption         =   "&Canceled"
            Index           =   2
            Begin VB.Menu mnuRptGenrSubBkCancel 
               Caption         =   "&Mill Booking"
               Index           =   0
            End
            Begin VB.Menu mnuRptGenrSubBkCancel 
               Caption         =   "&Trade Purchase Booking"
               Index           =   1
            End
            Begin VB.Menu mnuRptGenrSubBkCancel 
               Caption         =   "&Consignment/Depot Booking"
               Index           =   2
               Visible         =   0   'False
            End
            Begin VB.Menu mnuRptGenrSubBkCancel 
               Caption         =   "&Party Booking"
               Index           =   3
            End
         End
         Begin VB.Menu mnuRptGenrSubBk 
            Caption         =   "&Vs Despatch"
            Index           =   3
            Begin VB.Menu mnuRptGenrSubBkVs 
               Caption         =   "&Mill Booking"
               Index           =   0
               Visible         =   0   'False
            End
            Begin VB.Menu mnuRptGenrSubBkVs 
               Caption         =   "&Trade Purchase Booking"
               Index           =   1
               Visible         =   0   'False
            End
            Begin VB.Menu mnuRptGenrSubBkVs 
               Caption         =   "&Consignment/Depot Booking"
               Index           =   2
               Visible         =   0   'False
            End
            Begin VB.Menu mnuRptGenrSubBkVs 
               Caption         =   "&Party Booking"
               Index           =   3
            End
         End
      End
      Begin VB.Menu mnuRptGenrSub 
         Caption         =   "&Godown Reports"
         Index           =   10
         Begin VB.Menu mnuRptGenrSubGP 
            Caption         =   "Delivery Register"
            Index           =   0
         End
         Begin VB.Menu mnuRptGenrSubGP 
            Caption         =   "Pending Despatch"
            Index           =   1
         End
      End
      Begin VB.Menu mnuRptGenrSub 
         Caption         =   "-"
         Index           =   11
      End
      Begin VB.Menu mnuRptGenrSub 
         Caption         =   "&Interest Reports"
         Index           =   12
         Begin VB.Menu mnuRptGenrSubInt 
            Caption         =   "&Sarafi Interest Report"
            Index           =   0
         End
         Begin VB.Menu mnuRptGenrSubInt 
            Caption         =   "&Ank Interest Report"
            Index           =   1
            Visible         =   0   'False
         End
         Begin VB.Menu mnuRptGenrSubInt 
            Caption         =   "&Mill Bill Payment Interest"
            Index           =   2
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnuRptGenrSub 
         Caption         =   "-"
         Index           =   13
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRptGenrSub 
         Caption         =   "&MIS Reports"
         Index           =   14
         Begin VB.Menu mnuRptGenrSubMis 
            Caption         =   "&Company Status"
            Index           =   0
            Visible         =   0   'False
         End
         Begin VB.Menu mnuRptGenrSubMis 
            Caption         =   "Daily Report"
            Index           =   1
         End
         Begin VB.Menu mnuRptGenrSubMis 
            Caption         =   "Bank Detail Mail"
            Index           =   2
         End
         Begin VB.Menu mnuRptGenrSubMis 
            Caption         =   "Daily Report Mail"
            Index           =   3
         End
         Begin VB.Menu mnuRptGenrSubMis 
            Caption         =   "Lot No wise Transit Detail"
            Index           =   4
         End
         Begin VB.Menu mnuRptGenrSubMis 
            Caption         =   "Planed Payment List"
            Index           =   5
         End
         Begin VB.Menu mnuRptGenrSubMis 
            Caption         =   "Daily Report ( Special )"
            Index           =   6
         End
      End
      Begin VB.Menu mnuRptGenrSub 
         Caption         =   "-"
         Index           =   15
      End
      Begin VB.Menu mnuRptGenrSub 
         Caption         =   "Missing &Transaction Report"
         Index           =   16
         Begin VB.Menu mnuRptGenrSubMissTrans 
            Caption         =   "&Booking"
            Index           =   0
         End
         Begin VB.Menu mnuRptGenrSubMissTrans 
            Caption         =   "&Purchase"
            Index           =   1
         End
         Begin VB.Menu mnuRptGenrSubMissTrans 
            Caption         =   "&Gate Pass"
            Index           =   2
         End
         Begin VB.Menu mnuRptGenrSubMissTrans 
            Caption         =   "&Sales"
            Index           =   3
         End
         Begin VB.Menu mnuRptGenrSubMissTrans 
            Caption         =   "&Vouchers"
            Index           =   4
         End
      End
      Begin VB.Menu mnuRptGenrSub 
         Caption         =   "-"
         Index           =   17
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRptGenrSub 
         Caption         =   "&Graph"
         Index           =   18
         Visible         =   0   'False
         Begin VB.Menu mnuRptGenrSubGrp 
            Caption         =   "&Sales"
            Index           =   0
         End
         Begin VB.Menu mnuRptGenrSubGrp 
            Caption         =   "&Countwise Sales"
            Index           =   1
         End
         Begin VB.Menu mnuRptGenrSubGrp 
            Caption         =   "&Millwise Sales"
            Index           =   2
         End
         Begin VB.Menu mnuRptGenrSubGrp 
            Caption         =   "&Brokerwise Sales"
            Index           =   3
         End
         Begin VB.Menu mnuRptGenrSubGrp 
            Caption         =   "-"
            Index           =   4
         End
         Begin VB.Menu mnuRptGenrSubGrp 
            Caption         =   "&Purchase && Vs Sales"
            Index           =   5
         End
         Begin VB.Menu mnuRptGenrSubGrp 
            Caption         =   "Countwise Purchase && Vs Sales"
            Index           =   6
         End
         Begin VB.Menu mnuRptGenrSubGrp 
            Caption         =   "Millwise Purchase && Vs Sales"
            Index           =   7
         End
      End
   End
   Begin VB.Menu mnuCust 
      Caption         =   "Customised Reports"
      Visible         =   0   'False
      Begin VB.Menu mnuCustRpt 
         Caption         =   "1"
         Index           =   0
      End
   End
   Begin VB.Menu mnuUtl 
      Caption         =   "&Utility"
      Index           =   0
      Begin VB.Menu mnuUtlSub 
         Caption         =   "&Company Selection"
         Index           =   0
      End
      Begin VB.Menu mnuUtlSub 
         Caption         =   "&Shift Working Year"
         Index           =   1
      End
      Begin VB.Menu mnuUtlSub 
         Caption         =   "C&ompany Creation"
         Index           =   2
      End
      Begin VB.Menu mnuUtlSub 
         Caption         =   "&User Security"
         Index           =   3
      End
      Begin VB.Menu mnuUtlSub 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuUtlSub 
         Caption         =   "&Last Year Outstanding"
         Index           =   5
         Begin VB.Menu mnuLastYearOs 
            Caption         =   "Trade - SIT - Exemption"
            Index           =   0
         End
         Begin VB.Menu mnuLastYearOs 
            Caption         =   "Consignment - Depot"
            Index           =   1
         End
         Begin VB.Menu mnuLastYearOs 
            Caption         =   "Mill Bill ( Direct Sale )"
            Index           =   2
         End
      End
      Begin VB.Menu mnuUtlSub 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuUtlSub 
         Caption         =   "Update All &Entries!"
         Index           =   7
      End
      Begin VB.Menu mnuUtlSub 
         Caption         =   "Carry Forward &Purchase Stock!"
         Index           =   8
         Visible         =   0   'False
      End
      Begin VB.Menu mnuUtlSub 
         Caption         =   "Account &Transfer From -> To"
         Index           =   9
      End
      Begin VB.Menu mnuUtlSub 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu mnuUtlSub 
         Caption         =   "Carry &Forword Balances"
         Index           =   11
         Begin VB.Menu mnuTransfer 
            Caption         =   "Transfer Op.Balance A/c.  to Next Year (Except Loans Liabilities and  Deposit Assets Group )"
            Index           =   0
            Visible         =   0   'False
         End
         Begin VB.Menu mnuTransfer 
            Caption         =   "Transfer Op.Balance A/c.  to Next Year"
            Index           =   1
         End
         Begin VB.Menu mnuTransfer 
            Caption         =   "Transfer Trade Stock  to Next Year"
            Index           =   2
         End
         Begin VB.Menu mnuTransfer 
            Caption         =   "Transfer Depot Stock  to Next Year"
            Index           =   3
         End
         Begin VB.Menu mnuTransfer 
            Caption         =   "TransferDebitorBroker from Old Software"
            Index           =   4
            Visible         =   0   'False
         End
         Begin VB.Menu mnuTransfer 
            Caption         =   "Transfer O/s From Old Software"
            Index           =   5
            Visible         =   0   'False
         End
         Begin VB.Menu mnuTransfer 
            Caption         =   "Transfer Op. Bal.  From Old Software"
            Index           =   6
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnuUtlSub 
         Caption         =   "&Year End"
         Index           =   12
      End
      Begin VB.Menu mnuUtlSub 
         Caption         =   "-"
         Index           =   13
      End
      Begin VB.Menu mnuUtlSub 
         Caption         =   "&Backup Data"
         Index           =   14
      End
      Begin VB.Menu mnuUtlSub 
         Caption         =   "Tables / Views &Modifications"
         Index           =   15
         Visible         =   0   'False
         Begin VB.Menu mnuUtlSubMod 
            Caption         =   "Create/Modify &Views!"
            Index           =   0
         End
         Begin VB.Menu mnuUtlSubMod 
            Caption         =   "&Create Tables!"
            Index           =   1
         End
         Begin VB.Menu mnuUtlSubMod 
            Caption         =   "&Modify Tables!"
            Index           =   2
         End
      End
   End
   Begin VB.Menu mnuExit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
Private Sub MDIForm_Load()
   Dim i As Long, rst1 As Recordset
'   i = GProcRstOpen(rst1, "select REPONAME, RDTRANGE from UDFREPORTS where repostatus = 0", "R", gSetupCn)
'   If i > 0 Then
'        With rst1
'            i = 0
'            .MoveFirst
'            Do While Not .EOF
'               If (i > 0) Then Load mnuCustRpt(i)
'               mnuCustRpt(i).Caption = .Fields(0)
'               mnuCustRpt(i).Visible = True
'               i = i + 1
'              .MoveNext
'            Loop
'        End With
'    End If
    gMyConn = "gCn"
    ProcAllowFrmMainMenu
    
'    GProcRstOpen gRstAccMast, "Select AcName,AcCity,AcCode from TblMastAccount where acname <> '-' order by AcName", "R"
    GProcRstOpen gRstAccMast, "Select AcName,AcAdd1+','+AcAdd2,AcCode,AcOurCode,AcContPer from TblMastAccount where acname <> '-' order by AcName", "R"
    

    StatusBar1.Panels(1).text = "UserName : " & gUserName
    StatusBar1.Panels(2).text = "Logon Time : " & Time()
End Sub
'----- From User Previlages Allow Main Form Menu
Public Sub ProcAllowFrmMainMenu()
On Error GoTo ErrorHandler
Dim i As Long
Dim MnuCtrl As Control
Dim X As String
Dim RsMenu As Recordset
'--- Not Admin User
If gIsAdminUser = True Then
    mnuAdmin(0).Visible = True
Else
    mnuAdmin(0).Visible = False
End If
If gIsAdminUser = False Then
   For Each MnuCtrl In frmMain
    If TypeOf MnuCtrl Is Menu Then
        Select Case LCase(Left(MnuCtrl.Name, 6))
        '============= Previlage for menu "mnutl", "mnutrn", "mnurpt", "mnumst" Only
            Case "mnuutl", "mnutrn", "mnurpt", "mnumst"
                'RsMenu.Find "menuName='" & CStr(MnuCtrl.Name) & "' and MenuIndex=" & CLng(MnuCtrl.Index)
                'RsMenu.Find "Menu='" & CStr(MnuCtrl.Name & "-" & CStr(MnuCtrl.Index)) & "'"
                X = "Select * from tblUserPrevilage where UserName='" & gUserName & "' and " _
                & " menuName='" & MnuCtrl.Name & "' and MenuIndex=" & CLng(MnuCtrl.Index)
                If GProcRstOpen(RsMenu, X, "O", gSetupCn) = 0 Then
                    On Error Resume Next
                    MnuCtrl.Visible = False
                    On Error GoTo ErrorHandler
                End If
                RsMenu.Close
        End Select
     End If
    Next
End If
Exit Sub
ErrorHandler:
If Err.Number = 387 Then
    On Error Resume Next
Else
    MsgBox Err.Number & " " & Err.Description, vbOKOnly + vbInformation, Me.Caption
End If
End Sub
Private Sub MDIForm_Unload(Cancel As Integer)
gSetupCn.Execute "update tblMastUser set IsLogOn=0 where UserName='" & gUserName & "'"
End
End Sub
'------ Admin
Private Sub mnuAdminfrm_Click(Index As Integer)
    Select Case Index
    Case 0 '--User
        frmAdmUser.Show
    End Select
End Sub

Private Sub mnuLastYearOs_Click(Index As Integer)
Set gfrmMainMenu = mnuLastYearOs(Index)
Dim Srk As Integer
Dim x1 As String
Dim tmp1 As Recordset

Srk = 2001
gCn.Execute "delete from tmpGenTbl2 where username='" & gUserName & "'"
x1 = "Select * from tmpGenTbl2 where username='" & gUserName & "'"
GProcRstOpen tmp1, x1, "O", gCn
Do While Srk < 2019
      tmp1.AddNew
      tmp1!VYear = Srk
      tmp1!UserName = gUserName
      tmp1.Update
      Srk = Srk + 1
Loop
gClsSearch.SearchMultiField "tmpGenTbl2", "Vyear", Array("Bill Year"), Array(1000), " username='" & gUserName & "'", , "Amt1"
If gClsSearch.SearchMultiRetCol(0) > 0 Then
   gOldBillYear = gClsSearch.SearchMultiRetCol(0)
Else
    Exit Sub
End If
Select Case Index
    Case 0 '-- Trade SIT
        gfrmTypeStr = "SY"
        gfrmHunkStr = ""
        frmSalesOldOs.Show
    Case 1 '--Consignment/Depot
        gfrmTypeStr = "SD"
        gfrmHunkStr = ""
        frmSalesOldOs.Show
    Case 2 '--Mill Bill
        gfrmTypeStr = "SM"
        gfrmHunkStr = ""
        frmSalesOldOs.Show
End Select
End Sub

Private Sub mnuMstFirmIntRt_Click(Index As Integer)
  frmMastFirmwiseIntRate.Show
End Sub

'-- Master
Private Sub mnuMstfrm_Click(Index As Integer)
Set gfrmMainMenu = mnuMstfrm(Index)
Select Case Index
    Case 0 '-- Account
        frmMastAccount.Show
    Case 1 '-- Group
        frmMastGroup.Show
    Case 2 '-- Company Group
        frmMastCompGroup.Show
    Case 3 '-- Item
        frmMastItem.Show
    Case 4 '-- Godown
        gfrmTypeStr = "G"
        frmMastNarration.Show
    Case 5 '-- Narration
        gfrmTypeStr = "N"
        frmMastNarration.Show
    Case 6 '-- Tax Type
        'gfrmTypeStr = "T"
        frmMastNarrationTax.Show
    Case 7 '-- Sub Group Master
        frmMastSubGroup.Show
    Case 8 '-- Transport Type
        gfrmTypeStr = "R"
        frmMastNarration.Show
    Case 9 '--Party Bank Master --Sizer
        gfrmTypeStr = "S"
        frmMastNarration.Show
    Case 10  '--Consignee Master
        gfrmTypeStr = "C"
        frmMastNarration.Show
'        frmMastTransactionCode.Show
    Case 11 '-- Other Item
        frmMastItemOther.Show
    Case 12 '-- Other Item
        frmMastDelAdd.Show
    Case 14 '-- Bill Serial Master
        frmMastBillSerial.Show
    Case 16 '-- Interest Calculation Upto
        frmIntCalDate.Show
    End Select
End Sub
Private Sub mnuMstfrmSubOpStk_Click(Index As Integer)
Set gfrmMainMenu = mnuMstfrmSubOpStk(Index)
Select Case Index
    Case 0 '-- Trade
        gfrmTypeStr = "OY"
        gfrmHunkStr = ""
        frmOpStockTrade.Show
    Case 1 '--Goods Inward
        gfrmTypeStr = "OI"
        gfrmHunkStr = ""
        frmOpStockTrade.Show
    Case 2 '--Exemption
        gfrmTypeStr = "OE"
        gfrmHunkStr = ""
        frmOpStockTrade.Show
    Case 3 '--SIT
        gfrmTypeStr = "OT"
        gfrmHunkStr = ""
        frmOpStockTrade.Show
End Select

End Sub

Private Sub mnuMstHSNUpdate_Click(Index As Integer)
frmUpdateHSN.Show
End Sub

Private Sub mnuMstSetting_Click(Index As Integer)
Select Case Index
    Case 0 '-- Setting
        frmSetting.Show
End Select
End Sub
Private Sub mnuRptGenrSubBkCancel_Click(Index As Integer)
gSelectedMenu = ""
'Select Case Index
'    Case 0: '-- Canceled Mill Booking Register
'        gSelectedMenu = "Canceled Mill Booking Register"
'    Case 1 '--  Canceled SIT Booking Register
'        gSelectedMenu = "Canceled SIT Booking Register"
'    Case 2: '-- Canceled Consignment/Depot Booking Register
'        gSelectedMenu = "Canceled Consignment/Depot Booking Register"
'    Case 3: '-- Canceled Party Booking Register
'        gSelectedMenu = "Canceled Party Booking Register"
'End Select
'If gSelectedMenu <> "" Then frmRptRegi.Show
End Sub

Private Sub mnuRptGenrSubBkReg_Click(Index As Integer)
gSelectedMenu = ""
Select Case Index
'    Case 0: '-- Mill Booking Register
'        gSelectedMenu = "Mill Booking Register"
'    Case 1 '--  SIT Booking Register
'        gSelectedMenu = "SIT Booking Register"
'    Case 2: '-- Consignment/Depot Booking Register
'        gSelectedMenu = "Consignment/Depot Booking Register"
    Case 3: '-- Party Booking Register
        gSelectedMenu = "Party Booking Register"
    Case 4: '--Party Booking Summery
        gSelectedMenu = "Party Booking Summery"
    Case 5: '--Purchase Booking Summery
        gSelectedMenu = "Purchase Booking Summery"
End Select
If gSelectedMenu <> "" Then frmRptRegi.Show
End Sub

Private Sub mnuRptGenrSubBkStk_Click(Index As Integer)
gSelectedMenu = ""
Select Case Index
'    Case 0: '-- Pending Mill Booking Register
'        gSelectedMenu = "Pending Mill Booking Register"
    Case 1 '--  Trade Purchase Booking
        gSelectedMenu = "Trade Purchase Booking"
'    Case 2: '-- Pending Consignment/Depot Booking Register
'        gSelectedMenu = "Pending Consignment/Depot Booking Register"
    Case 3: '-- Pending Party Booking Register
        gSelectedMenu = "Pending Party Booking Register"
     Case 4:  '-- Despatch Goods Details
         gSelectedMenu = "Despatch Goods Details"
End Select
If gSelectedMenu <> "" Then frmRptRegi.Show
End Sub

Private Sub mnuRptGenrSubBkVs_Click(Index As Integer)
gSelectedMenu = ""
Select Case Index
'    Case 0: '--  Mill Booking Vs Dispatch Register
'        gSelectedMenu = "Mill Booking Vs Dispatch Register"
'    Case 1 '--  SIT Booking Vs Despatch Register
'        gSelectedMenu = "SIT Booking Vs Dispatch Register"
'    Case 2: '--  Consignment/Depot Booking Vs Dispatch Register
'        gSelectedMenu = "Consignment/Depot Booking Vs Dispatch Register"
    Case 3: '-- Party Booking Vs Dispatch Register
        gSelectedMenu = "Party Booking Vs Dispatch Register"
End Select
If gSelectedMenu <> "" Then frmRptRegi.Show
End Sub

Private Sub mnuRptGenrSubBro_Click(Index As Integer)
gSelectedMenu = ""
Select Case Index
    Case 0: '-- Brokerage Detail Report
        gSelectedMenu = "Broker Ledger"
    Case 1 '--  Brokerage Detail
        gSelectedMenu = "Brokerage Detail "
    Case 2: '-- Brokerage Summery
        gSelectedMenu = "Brokerage Summery"
    Case 3: '-- Brokerage Voucher
        gSelectedMenu = "Brokerage Voucher"
    Case 4: '-- Countwise Brokarage
        gSelectedMenu = "Countwise Brokarage"
'    Case 5: '-- Inward/Purchase Brokarage Total
'        gSelectedMenu = "Inward/Purchase Brokarage Total"
End Select
If gSelectedMenu <> "" Then frmRptTaxReports.Show
End Sub

Private Sub mnuRptGenrSubCLBal_Click(Index As Integer)
gSelectedMenu = ""
'Select Case Index
'   Case 0 '-"Account Closing Balance List"
'        gSelectedMenu = "Account Closing Balance List"
'        frmRptAccount.Show
'        Exit Sub
'   Case 1 '-"Account Groupwise Closing Balance List"
'        gSelectedMenu = "Account Groupwise Closing Balance List"
'        frmRptAccount.Show
'        Exit Sub
'End Select
End Sub

Private Sub mnuRptGenrSubGP_Click(Index As Integer)
gSelectedMenu = ""
Select Case Index
     Case 0: '-- Delivery Register
        gSelectedMenu = "Delivery Register"
     Case 1 '-- Gate Pass Register
        gSelectedMenu = "Pending Despatch"
End Select
If gSelectedMenu <> "" Then frmRptRegi.Show
End Sub

Private Sub mnuRptGenrSubGPDel_Click(Index As Integer)
gSelectedMenu = ""
Select Case Index
     Case 0: '-- Countwise Goods Delivery Report
        gSelectedMenu = "Countwise Goods Delivery Report"
     Case 1 '-- Count & Party wise Goods Delivery Report
        gSelectedMenu = "Count & Partywise Goods Delivery Report"
     Case 2 '-- "Count & Sizerwise Goods Delivery Report"
        gSelectedMenu = "Count & Sizerwise Goods Delivery Report"
     Case 3 '--"Party and Bag No wise Goods Delivery Report"
        gSelectedMenu = "Party and Bag No wise Goods Delivery Report"
End Select
If gSelectedMenu <> "" Then frmRptRegi.Show
End Sub

Private Sub mnuRptGenrSubGrp_Click(Index As Integer)
gSelectedMenu = ""
'Select Case Index
'    Case 0 '- Sales Graph
'        gSelectedMenu = "Sales Graph"
'    Case 1 '- Countwise Sales Graph
'        gSelectedMenu = "Countwise Sales Graph"
'    Case 2 '- Millwise Sales Graph
'        gSelectedMenu = "Millwise Sales Graph"
'    Case 3 '- Brokerwise Sales Graph
'        gSelectedMenu = "Brokerwise Sales Graph"
'    Case 4  '-
'    Case 5 '- Purchase Graph
'        gSelectedMenu = "Purchase Graph"
'    Case 6 '- Countwise Purchase Graph
'        gSelectedMenu = "Countwise Purchase Graph"
'    Case 7 '- Millwise Purchase Graph
'        gSelectedMenu = "Millwise Purchase Graph"
'End Select
If gSelectedMenu <> "" Then frmRptGraph.Show
End Sub

Private Sub mnuRptGenrSubInt_Click(Index As Integer)
gSelectedMenu = ""
Select Case Index
 Case 0 '-- Sarafi Interest Report
    gSelectedMenu = "Sarafi Interest Report"
' Case 1 '-- Sarafi Interest Report Ank
'    gSelectedMenu = "Sarafi Interest Report ANK"
' Case 2 '---Mill Bill Payment Interest
'    gSelectedMenu = "Mill Bill Payment Interest"
End Select
If gSelectedMenu <> "" Then frmRptAccount.Show
End Sub

Public Sub mnuRptGenrSubLst_Click(Index As Integer)
gSelectedMenu = ""
Select Case Index
Case 0 '- Account List
'    gSelectedMenu = "Account List"
'    frmRptAccount.Show
    '--Tem For Dublicate Ac Name
    Dim X As String
    X = "select AcOurCode,AcName, '" & gUserName & "'" _
            & " from tblMastAccount where len(tblMastAccount.AcName) > 40"
    gCn.Execute "Delete from TmpGentbl where  UserName='" & gUserName & "'"
    gCn.Execute " insert into TmpGentbl (PartyCode,PartyName,UserName) " & X
    With frmMain.cryReport
         .SelectionFormula = "{TmpGentbl.UserName}='" & gUserName & "'"
         .ReportFileName = gReportPath & "rptAccNameDublicate.rpt"
         .Action = 1
    End With
    Exit Sub
'Case 1 '- Groupwise Account List
'    gSelectedMenu = "Groupwise Account List"
'    frmRptAccount.Show
'    Exit Sub
'Case 2 '-Selected Group Accounts List
'    gSelectedMenu = "Selected Group Accounts List"
'    frmRptAccount.Show
'    Exit Sub
'Case 3 '- -
'Case 4 '- Count Rate List
'    GProcCrystalRptPreparation cryReport, "Window"
'    cryReport.WindowTitle = "Count Rate List"
'    cryReport.ReportFileName = gReportPath & "RptLstCountRateList.rpt"
'    cryReport.Action = 1
'    Exit Sub
Case 5 '- Millwise Count List
    gSelectedMenu = "Millwise Count List"
    frmRptAccount.Show
    Exit Sub
'Case 6 '- Count Brokerage List
'    gSelectedMenu = "Count Brokerage List"
'    frmRptAccount.Show
'    Exit Sub
'Case 7 ' -- Selected Group No Transaction A/c List
'    gSelectedMenu = "Selected Group No Transaction A/c List"
'    frmRptAccount.Show
'    Exit Sub
'Case 8 '-Party Groupwise List
'    gSelectedMenu = "Party Groupwise List"
'    frmRptAccount.Show
'    Exit Sub
End Select
End Sub

Private Sub mnuRptGenrSubMis_Click(Index As Integer)
gSelectedMenu = ""
Select Case Index  '-- MIS Reports
       Case 1 '-- Purchase
            gSelectedMenu = "Daily Report"
            If gSelectedMenu <> "" Then frmRptRegi.Show
       Case 2 '--Bank Detail Mail
            gSelectedMenu = "Bank Detail Mail"
            If gSelectedMenu <> "" Then frmRpTWhatsApp.Show
       Case 3 '--Daily Report Mail
            gSelectedMenu = "Daily Report Mail"
            If gSelectedMenu <> "" Then frmRptRegi.Show
       Case 4 '-Lot No wise Transit Detail
            gSelectedMenu = "Lot No wise Transit Detail"
            If gSelectedMenu <> "" Then frmRptRegi.Show
       Case 5 '--Planed Payment List
            gSelectedMenu = "Planed Payment List"
            If gSelectedMenu <> "" Then frmRptRegi.Show
       Case 6 '--Daily Report ( Special )
            gSelectedMenu = "Daily Report ( Special )"
            If gSelectedMenu <> "" Then frmRptRegi.Show
            
End Select
'If gSelectedMenu <> "" Then frmRptRegi.Show
End Sub

Private Sub mnuRptGenrSubMissTrans_Click(Index As Integer)
gSelectedMenu = ""
'Select Case Index  '-- Missing Transaction
'    Case 0  '-- Booking
'        gSelectedMenu = "Missing Transaction Report - Booking"
'    Case 1 '-- Purchase
'        gSelectedMenu = "Missing Transaction Report - Purchase"
'    Case 2 '-- Gate Pass
'        gSelectedMenu = "Missing Transaction Report - Gate Pass"
'    Case 3 '-- Sales
'        gSelectedMenu = "Missing Transaction Report - Sales"
'    Case 4 '-- Vouchers
'        gSelectedMenu = "Missing Transaction Report - Vouchers"
'End Select
If gSelectedMenu <> "" Then frmRptRegi.Show
End Sub

Private Sub mnuRptGenrSubMon_Click(Index As Integer)
gSelectedMenu = ""
Select Case Index
     Case 0: '-- Consignment/ Depot Statement
       gSelectedMenu = "Consignment/ Depot Statement"
     Case 1 '-- Mill Billing Statement
       gSelectedMenu = "Mill Billing Statement"
'     Case 2 '-- Sales In Transit Statement
'       gSelectedMenu = "Sales In Transit Statement"
End Select
If gSelectedMenu <> "" Then frmRptStatement.Show
End Sub

Private Sub mnuRptGenrSubOPBal_Click(Index As Integer)
'gSelectedMenu = ""
Select Case Index
   Case 0 '-"Account Opening Balance List"
        gSelectedMenu = "Account Opening Balance List"
        frmRptAccount.Show
        Exit Sub
   Case 1 '-"Account Groupwise Opening Balance List"
        gSelectedMenu = "Account Groupwise Opening Balance List"
        frmRptAccount.Show
        Exit Sub
End Select
End Sub

Private Sub mnuRptGenrSubPur_Click(Index As Integer)
gSelectedMenu = ""
Select Case Index
    Case 0: '--  "Total Purchase
        gSelectedMenu = "Trade Purchase TDS Report"
    Case 1 '--"Depot Sales TDS Report"
        gSelectedMenu = "Depot Sales TDS Report"
    Case 2 '--"Mill Bill TDS Report"
        gSelectedMenu = "Mill Bill TDS Report"
    Case 3 '--Millwise Inward
        gSelectedMenu = "Millwise Inward"
    Case 4 '--Daily Purchase Report
        gSelectedMenu = "Daily Purchase Report"
    Case 5 '--Daily Purchase Report ( Mill Bill )
        gSelectedMenu = "Daily Purchase Report ( Mill Bill )"
    Case 6 '--Other Purchase TDS Report
        gSelectedMenu = "Other Purchase TDS Report"
End Select
If gSelectedMenu <> "" Then frmRptTaxReports.Show
End Sub
Private Sub mnuRptGenrSubSale_Click(Index As Integer)
gSelectedMenu = ""
Select Case Index
    Case 0: '-- "Total Sales"
        gSelectedMenu = "Total Sales Summary"
    Case 1 '-- Count / Brokerwise Sales
        gSelectedMenu = "Count / Brokerwise Sales"
    Case 2 '- Party / Countwise Sales
        gSelectedMenu = "Party / Countwise Sales"
    Case 3 '-- Sales Against Booking
        gSelectedMenu = "Sales Against Booking"
'    Case 4 '-- Countwise Sales
'        gSelectedMenu = "Countwise Sales Summary"
'    Case 5: '-- Millwise Sales
'        gSelectedMenu = "Millwise Sales Summary"
'    Case 6 '- -
'    Case 7: '-- Party and Countwise Sales
'        gSelectedMenu = "Party and Countwise Sales Summary"
'    Case 8: '-- Mill & Countwise Sales
'        gSelectedMenu = "Mill and Countwise Sales Summary"
'    Case 9: '-- Broker and Partywise Sales
'        gSelectedMenu = "Broker and Partywise Sales Summary"
'    Case 10: '-- "Broker & Countwise Sales"
'        gSelectedMenu = "Broker and Countwise Sales Summary"
End Select
If gSelectedMenu <> "" Then frmRptTaxReports.Show
End Sub

Private Sub mnuRptRSub_Click(Index As Integer)
If Index = 8 Then
   gSelectedMenu = "Party Interest Calculation"
'   frmRptPartyIntCal.Show
ElseIf Index = 10 Then
   gSelectedMenu = "TCS Return"
   frmRptGraph.Show
ElseIf Index = 11 Then
   gSelectedMenu = "TCS Bills Check"
   frmRptGraph.Show
ElseIf Index = 13 Then
   gSelectedMenu = "XXXX"
   frmRptGraph.Show
ElseIf Index = 14 Then
   gSelectedMenu = "Audit Trail Reports"
   frmRptAuditTrial.Show
End If




End Sub

'Private Sub mnuRptRSub_Click(Index As Integer)
'gSelectedMenu = "Daily Massage"
'frmRptSelection.Show
'End Sub

Private Sub mnurptRSubBookPri_Click(Index As Integer)
gSelectedMenu = ""
Select Case Index
     Case 0 '-Cash Book
       gSelectedMenu = "Cash Book"
     Case 1 '-Bank Book
       gSelectedMenu = "Bank Book"
     Case 2 '-Credit Cash Report
       gSelectedMenu = "Credit Cash Report"
End Select
'If gSelectedMenu <> "" Then frmRptAccount.Show
End Sub

Private Sub mnuRptRSubGstRp_Click(Index As Integer)
gSelectedMenu = ""
Select Case Index
    Case 0:
           gSelectedMenu = "Trade Sale Register GST" '----Trade Sale Register
    Case 1:
           gSelectedMenu = "Depot Sale Register GST" '----Depot Sale Register
    Case 2:
          gSelectedMenu = "Commission And Brokerage Sales"
    Case 3 '----"Late Pay Debit Note Register"
           gSelectedMenu = "Late Pay Debit Note Register"
    Case 4  '--"Late Pay Debit Note Register (Mill Bill)"
           gSelectedMenu = "Late Pay Debit Note Register (Mill Bill)"
    Case 5  '--"Non-GST Outward Sales"
           gSelectedMenu = "Non-GST Outward Sales"
    Case 6:
        gSelectedMenu = "R C M Sale Register GST" '----Trade Sale Register
    Case 7:  '-
    Case 8:  '-  "Credit Note (Sales) Register"
        gSelectedMenu = "Credit Note (Sales) Register"
    Case 9:  '-  "Debit Note (Sales) Register"
        gSelectedMenu = "Debit Note (Sales) Register"
    Case 10:  '- "Credit Note (Purchase) Register"
        gSelectedMenu = "Credit Note (Purchase) Register"
    Case 11:  '- "Debit Note (Purchase) Register"
        gSelectedMenu = "Debit Note (Purchase) Register"
    Case 13:
        gSelectedMenu = "Trade Purchase Register GST"
    Case 14:
        gSelectedMenu = "Other Purchase Register GST"
    Case 15:  '-
    Case 16:
        gSelectedMenu = "GSTR 3B"
        frmRptGstReports.Show
        Exit Sub
    Case 17:
        gSelectedMenu = "GSTR - 1"
        frmRptGstReports.Show
        Exit Sub
    Case 18:
        gSelectedMenu = "Monthwise GST Summery"
        frmRptGstReports.Show
        Exit Sub
    Case 19:
        gSelectedMenu = "GSTR - 2A  -  B2B"
        frmRptGstReports.Show
        Exit Sub
    Case 20:
        gSelectedMenu = "GSTR - 2A  -  B2B  (Mismatch)"
        frmRptGstReports.Show
        Exit Sub
    Case 21: '-
    Case 22:
        gSelectedMenu = "Sales Return Register"
    Case 23:
        gSelectedMenu = "Purchase Return Register"
    Case 24:
        gSelectedMenu = "GSTR - 9"
        frmRptGstReports.Show
        Exit Sub
End Select
frmRptRegi.Show
End Sub

Private Sub mnurptRSubOutSta_Click(Index As Integer)
gSelectedMenu = ""
Select Case Index
     Case 2:
         gSelectedMenu = "TDS Receivable O/s"
     Case 3:
         gSelectedMenu = "TCS Receivable O/s"
     Case 4:
         gSelectedMenu = "TDS On Sales Receivable O/s"
     Case 5:
         gSelectedMenu = "TDS On Sales ( Mill Bill )"
     Case 6:
         gSelectedMenu = "TDS On Sales ( Depot )"
End Select
If gSelectedMenu <> "" Then frmRptAccount.Show
End Sub

Private Sub mnurptRSubOutStaPur_Click(Index As Integer)
gSelectedMenu = ""
Select Case Index
     Case 0: '-"Purchase Outstanding Partywise"
        gSelectedMenu = "Purchase Outstanding Partywise"
     Case 1: '- "Purchase Outstanding Billwise"
        gSelectedMenu = "" '"Purchase Outstanding Billwise"
     Case 2: '-"Unadjusted Payment"
       gSelectedMenu = "Unadjusted Payment"
     Case 3: '"Mill Bill Mill O/s"
       gSelectedMenu = "Mill Bill Mill O/s"
End Select
'If gSelectedMenu <> "" Then frmRptRegi.Show
End Sub

Public Sub mnurptRSubOutStaSale_Click(Index As Integer)
gSelectedMenu = ""
gRptMillFlag = False
gRptBrokerFlag = False
gRptOnDateFlag = False
gRptPartyFlag = False
Select Case Index
    Case 0:
        gSelectedMenu = "Outstanding Brokerwise"
    Case 1:
        gSelectedMenu = "Outstanding Partywise"
    Case 2:
        gSelectedMenu = "Outstanding Invoicewise"
    Case 3 '-
    Case 4:
       gRptPartyFlag = True
       gSelectedMenu = "Late Payment Interest List"
       frmRptSelection.Show
       Exit Sub
    Case 5:
       gSelectedMenu = "Ledger And Oustanding Sale Bal Check"
    Case 6:
       gRptPartyFlag = True
       gSelectedMenu = "Early Payment Interest List"
       frmRptSelection.Show
       Exit Sub
End Select
If gSelectedMenu <> "" Then frmRptRegi.Show
End Sub

Private Sub mnurptRSubRegPri_Click(Index As Integer)
gSelectedMenu = ""
Select Case Index
    Case 0 '-Sale Register
    Case 1  '-Purchase Register
    Case 2 '----"Sale In Transit Register"
        gSelectedMenu = "Insurance Report - Purchase"
    Case 3 '-
    Case 4 '----"Credit Note Register"
        gSelectedMenu = "Credit Note (Sales) Register"
    Case 5  '----"Debit Note Register"
        gSelectedMenu = "Debit Note (Sales) Register"
    Case 6  '----"Internal Voucher Register"
        gSelectedMenu = "Internal Voucher Register"
    Case 7
        gSelectedMenu = "Credit Note (Purchase) Register"
    Case 8 '-
        gSelectedMenu = "Debit Note (Purchase) Register"
    
    Case 9 '-Sale Return
        gSelectedMenu = "Sales Return Register"
    Case 10  '-Purchase Return
        gSelectedMenu = "Purchase Return Register"
    Case 11 '-
    Case 12 '----"Late Pay Debit Note Register"
        gSelectedMenu = "Late Pay Debit Note Register"
    Case 13  '--"Late Pay Debit Note Register (Mill Bill)"
        gSelectedMenu = "Late Pay Debit Note Register (Mill Bill)"
End Select
If gSelectedMenu <> "" Then frmRptRegi.Show
End Sub

Private Sub mnurptRSubRegPriPur_Click(Index As Integer)
gDepotConsiMillHelp = False
gSelectedMenu = ""
Select Case Index
    Case 0:
        gSelectedMenu = "Trade Purchase Register GST" '----Trade Purchase Register
    Case 1:
        gSelectedMenu = "SIT Purchase Register" '----SIT Purchase Register
    Case 2:
        gSelectedMenu = "Other Purchase Register GST" '----Other Purchase Register
    Case 3:
        gDepotConsiMillHelp = True
        gSelectedMenu = "Consignment/Depot Purchase Register" '----Consignment/Depot Purchase Register
End Select
If gSelectedMenu <> "" Then frmRptRegi.Show
End Sub
Private Sub mnuRptRSubSMS_Click(Index As Integer)
'On Error GoTo ErrorHandler
gSelectedMenu = ""
'''Select Case Index
'''       Case 0 '---Pending C form Receivable
'''            gSelectedMenu = "Pending C form Receivable"
'''       Case 1 '---"Form 'C' Receivable SMS (Trade/SIT) Single Party"
'''            gSelectedMenu = "Form 'C' Receivable SMS (Trade/SIT) Single Party"
'''       Case 2 '---"Form 'C' Receivable SMS (Trade/SIT) Comman SMS to All Party"
'''            gSelectedMenu = "Form 'C' Receivable SMS (Trade/SIT) Comman SMS to All Party"
'''       Case 3 '-
'''
'''       Case 4 '---"Form 'E1' Receivable SMS  All Party"
'''            gSelectedMenu = "Form 'E1' Receivable SMS  All Party"
'''       Case 5 '---"Form 'E1' Receivable SMS  Single Party"
'''            gSelectedMenu = "Form 'E1' Receivable SMS  Single Party"
'''       Case 6 '-
'''       Case 7 '-Form 'E2' Receivable SMS  All Party
'''            gSelectedMenu = "Form 'E2' Receivable SMS  All Party"
'''       Case 8 '-Form 'E2' Receivable SMS  Single Party
'''            gSelectedMenu = "Form 'E2' Receivable SMS  Single Party"
'''       Case 9 '-
'''       Case 10 '---General SMS
'''            gSelectedMenu = "General SMS"
'''       Case 11 '-
'''       Case 12 '---Pending C form Receivable All Broker
'''            gSelectedMenu = "Pending C form Receivable All Broker"
'''       Case 13 '---Pending C form Receivable Single Broker
'''            gSelectedMenu = "Pending C form Receivable Single Broker"
'''       Case 14 '----Form 'E1' Receivable SMS (Trade/SIT) All Broker
'''            gSelectedMenu = "Form 'E1' Receivable SMS (Trade/SIT) All Broker"
'''       Case 15 '----Form 'E1' Receivable SMS (Trade/SIT) Single Broker
'''            gSelectedMenu = "Form 'E1' Receivable SMS (Trade/SIT) Single Broker"
'''       Case 16 '---Form 'E1' Receivable SMS Comman SMS to All Party
'''            gSelectedMenu = "Form 'E1' Receivable SMS Comman SMS to All Party"
'''       Case 17 '-
'''       Case 18 '---Outstanding To Party
'''            gSelectedMenu = "Outstanding To Party"
'''       Case 19 '---Outstanding To Broker
'''            gSelectedMenu = "Outstanding To Broker"
'''End Select
'''If gSelectedMenu <> "" Then frmRpTSMS.Show
'''Exit Sub
'''ErrorHandler:
GProcErrorHandler
End Sub

Private Sub mnurptRSubStock_Click(Index As Integer)
'''On Error GoTo ErrorHandler
gSelectedMenu = ""
Select Case Index
    Case 0 '-- Item Ledger
        gSelectedMenu = "Stock Statement"
        frmRptRegi.Show
    Case 1 '-- Daily Stock Statement
        gSelectedMenu = "Daily Stock Statement"
        frmRptRegi.Show
    Case 2 ' -- Sales Summery
        gSelectedMenu = "Sales Summery"
        frmRptRegi.Show
    Case 3 ' -- Gate Pass Wise Sales
        gSelectedMenu = "Gate Pass Wise Sales"
        frmRptRegi.Show
    Case 4 '--Stock Statement Godownwsie
        gSelectedMenu = "Stock Statement Godownwsie"
        frmRptRegi.Show

End Select
Exit Sub
'''ErrorHandler:
'''GProcErrorHandler
End Sub

Private Sub mnurptRSubStockVal_Click(Index As Integer)
'''On Error GoTo ErrorHandler
'''gSelectedMenu = ""
'''Select Case Index
'''    Case 0 '-- Stock Valuation (Trade,Sit)
'''        gSelectedMenu = "Trade & SIT Stock Valuation Register"
'''        frmRptRegi.Show
'''        Exit Sub
'''    Case 1 '-- "Trade & SIT Stock Valuation Summary"
'''        gSelectedMenu = "Trade & SIT Stock Valuation Summary"
'''        frmRptRegi.Show
'''        Exit Sub
'''End Select
'''Exit Sub
'''ErrorHandler:
'''GProcErrorHandler
End Sub
Private Sub mnurptRSubTaxAuditForm_Click(Index As Integer)
gSelectedMenu = ""
'''Select Case Index
'''    Case 0 '--"Received C And E1 Form"
'''        gSelectedMenu = "Annexure - G (Received C And E1 Form)"
'''    Case 1 '--"Annexure - I  (Pending  C And E1 Form)"
'''        gSelectedMenu = "Annexure - I  (Pending  C And E1 Form)"
'''    Case 2 '--"Annexure - J Section -1   (Partywise VAT Sales)"
'''        gSelectedMenu = "Annexure - J Section -1   (Partywise VAT Sales)"
'''    Case 3 '--"Annexure - J Section -1   (R.D. Partywise VAT Sales)"
'''        gSelectedMenu = "Annexure - J Section -1   (R.D. Partywise VAT Sales)"
'''    Case 4 '---"Annexure - J Section -1   (U.R.D. Partywise VAT Sales)"
'''        gSelectedMenu = "Annexure - J Section -1   (U.R.D. Partywise VAT Sales)"
'''    Case 5 '---"Annexure - J Section -2   (Partywise VAT Purchases)"
'''        gSelectedMenu = "Annexure - J Section -2   (Partywise VAT Purchases)"
'''    Case 6 '---"Annexure - J Section -6   (Supllierwise CST Purchases)"
'''        gSelectedMenu = "Annexure - J Section -6   (Supllierwise CST Purchases)"
'''
'''    Case 8 '--"Partywise Sales VAT Audit"
'''        gSelectedMenu = "Partywise Sales VAT Audit"
'''    Case 9 '--"Partywise Purchase VAT Audit"
'''        gSelectedMenu = "Partywise Purchase VAT Audit"
'''    Case 10 '--"Monthly Sales Detail VAT Audit"
'''        gSelectedMenu = "Monthly Sales Detail VAT Audit"
'''    Case 11 '--"Monthly Purchase Detail VAT Audit"
'''        gSelectedMenu = "Monthly Purchase Detail VAT Audit"
'''    Case 12 '--"Pending C Form (Receivable) VAT Audit"
'''        gSelectedMenu = "Pending C Form (Receivable) VAT Audit"
'''    Case 13 '--"Pending E1 Form (Receivable) VAT Audit"
'''        gSelectedMenu = "Pending E1 Form (Receivable) VAT Audit"
'''    Case 14 '--"Received C Form"
'''        gSelectedMenu = "Received C Form"
'''    Case 15 '---"Received E1 Form"
'''        gSelectedMenu = "Received E1 Form"
'''    Case 16 '---"Suplierwise CST Purchase Details"
'''        gSelectedMenu = "Suplierwise CST Purchase Details"
'''
'''End Select
'''If gSelectedMenu <> "" Then frmRptTaxReports.Show
End Sub

Private Sub mnurptRSubTaxForm_Click(Index As Integer)
gSelectedMenu = ""
'''Select Case Index
'''    Case 0 '--"C Form Receivable"
'''        gSelectedMenu = "C Form Receivable"
'''    Case 1 '-- "C Form Issuable"
'''        gSelectedMenu = "C Form Issuable"
'''    Case 2 '-- "E1 From Receivable"
'''        gSelectedMenu = "E1 Form Receivable"
'''    Case 3 '-- "E2 From Receivable"
'''        gSelectedMenu = "E2 Form Receivable"
'''    Case 4 '-- "E2 From Issuable"
'''        gSelectedMenu = "E2 From Issuable"
'''    Case 5 '-- "F Form Issuable"
'''        gSelectedMenu = "F Form Issuable"
'''    Case 6 '-
'''    Case 7 '-- "Requirement Of C From"
'''        gSelectedMenu = "Requirement Of C Form"
'''    Case 8 '-- "Requirement Of F From"
'''        gSelectedMenu = "Requirement Of F Form"
'''    Case 9 '-- F Form Statement
'''        gSelectedMenu = "F Form Statement"
'''    Case 10 '-
'''    Case 11 '-- Pending C Form Receivable Slab Wise
'''        gSelectedMenu = "Pending C Form Receivable Slab Wise"
'''    Case 12 '-- Pending E1 Form Receivable Slab Wise
'''        gSelectedMenu = "Pending E1 Form Receivable Slab Wise"
'''    Case 13 '-- C Form Issuable and E1-E2 Receivable Summery
'''        gSelectedMenu = "C Form Issuable and E1-E2 Receivable Summery"
'''    Case 14 '--"Requirment of E2 Form"
'''        gSelectedMenu = "Requirment of E2 Form"
'''    Case 15 '-Requirment of C Form Invoicewise
'''        gSelectedMenu = "Requirment of C Form Invoicewise"
'''    Case 16 '- Requirment of F Form Invoicewise
'''        gSelectedMenu = "Requirment of F Form Invoicewise"
'''    Case 17 '--"Requirment of E2 Form Invoicewise"
'''        gSelectedMenu = "Requirment of E2 Form Invoicewise"
'''    Case 18 '-
'''    Case 19 '---Notes On Forms
'''        gSelectedMenu = "Notes On Forms"
'''    Case 20 '---SIT Forms Summery
'''        gSelectedMenu = "SIT Forms Summery"
'''End Select
'''If gSelectedMenu <> "" Then frmRptTaxReports.Show
End Sub

Private Sub mnuRptRSubWapp_Click(Index As Integer)
Select Case Index '-- Bank Reconciliation Entry
    Case 0
        gSelectedMenu = "Outsading To Broker"
    Case 1
        gSelectedMenu = "Outsading To Party"
End Select
If gSelectedMenu <> "" Then frmRpTWhatsApp.Show

End Sub

Private Sub mnuTransfer_Click(Index As Integer)
Dim NextCYr As String
Dim NextDbNm As String
Dim NextYrStartDt As Date
NextCYr = Year(gCYEDate)
NextDbNm = gCCode & NextCYr
NextYrStartDt = gCYEDate + 1

Dim X As String
Dim Y As String
Dim z As String
Dim i As Integer
Dim RsP As Recordset
Dim RsP2 As Recordset
Dim RsP3 As Recordset
Dim tmpVno As Double
Dim LoopStr As String
Dim wCtr As Long
Set gfrmMainMenu = mnuTransfer(Index)
Select Case Index '-- Transfer To next Year
    Case 0, 1 '- Transfer Opening Balance to Next Year
         If MsgBox("Op. Balance Transfer To " & NextCYr & "  Are you Sure ?", vbYesNo) = vbYes Then
            X = "Select * from tblCompanyYear where CompCode = '" & gCCode & "' and CompYear =" & Right(Year(gCYEDate), 2) + Right(Year(gCYEDate), 2) + 1 & "  "
            If (GProcRstOpen(RsP3, X, "R") = 0) Then
               MsgBox ("Company Year Not Found...")
               Exit Sub
            End If
            gCn.BeginTrans
            If Index = 0 Then
               gCn.Execute "delete tblVoucher where tblVoucher.Vfirm = '" & gCCode & "' and  tblVoucher.Vtype = '0P' and tblVoucher.Vyear = " & NextCYr & " and (case when vcraccode=1 then vdraccode else vcraccode end) in (select accode from tblmastaccount where agcode not in (90006,90010,90018)) "
            Else
               gCn.Execute "delete tblVoucher where tblVoucher.Vfirm = '" & gCCode & "' and  tblVoucher.Vtype = '0P' and tblVoucher.Vyear = " & NextCYr & " and (case when vcraccode=1 then vdraccode else vcraccode end) in (select accode from tblmastaccount where agcode not in (90018)) "
            End If
            gCn.CommitTrans

            gCmd.CommandText = "PrcPrepareClBalance"
            gCmd.Parameters.Refresh
            gCmd.Parameters("@VFirm") = Trim(gCCode)
            gCmd.Parameters("@Fr_dt") = Format(gCYSDate, "yyyy/mm/dd")
            gCmd.Parameters("@To_dt") = Format(gCYEDate, "yyyy/mm/dd")
            gCmd.Parameters("@FyStart_dt") = Format(gCYSDate, "yyyy/mm/dd")
            gCmd.Parameters("@UCode") = gUserName
            gCmd.Parameters("@VFirm") = Trim(gCCode)
            gCmd.Parameters("@CGSTAc") = gCgstAcCode
            gCmd.Parameters("@SGSTAc") = gSgstAcCode
            gCmd.Parameters("@IGSTAc") = gIgstAcCode
            gCmd.Parameters("@TcsRec") = gTcsRec
            gCmd.Parameters("@TcsPay") = gTcsPay
            gCmd.Parameters("@CGSTRcmPayAc") = gCgstRCMAcCode
            gCmd.Parameters("@SGSTRcmPayAc") = gSgstRCMAcCode
            gCmd.Parameters("@IGSTRcmPayAc") = gIgstRCMAcCode
            gCmd.Parameters("@CGSTRcmRecAc") = gCgstRCMRecCode
            gCmd.Parameters("@SGSTRcmRecAc") = gSgstRCMRecCode
            gCmd.Parameters("@IGSTRcmRecAc") = gIgstRCMRecCode
            gCmd.Parameters("@RoundingAc") = gRoundOffAc
            gCmd.Execute

            DoEvents
            DoEvents
            DoEvents
            DoEvents
            z = "Select * from tblVoucher "
            GProcRstOpen RsP3, z, "O"
            
            If Index = 0 Then
               X = "Select * from TmpClosingBalance where UserName = '" & gUserName & "' and vfirm = '" & gCCode & "' and TopGrp < 90003 and (ClBal_Dr + ClBal_Cr) > 0 and AgCode not in (90018,90006,90010) and AcName <> 'Profit And Loss Of Current Year'  "
            Else
               X = "Select * from TmpClosingBalance where UserName = '" & gUserName & "' and vfirm = '" & gCCode & "' and TopGrp < 90003 and (ClBal_Dr + ClBal_Cr) > 0 and AgCode <> 90018 and AcName <> 'Profit And Loss Of Current Year'"
            End If
            If (GProcRstOpen(RsP, X, "R") > 0) Then
               tmpVno = 0
               RsP.MoveFirst
               Do While Not RsP.EOF
                   tmpVno = tmpVno + 1
                   RsP3.AddNew
                   RsP3.Fields("Vno") = RsP.Fields("AcCode")  '--   Val(tmpVno)
                   RsP3.Fields("VType") = "0P"
                   RsP3.Fields("Vdt") = NextYrStartDt
                   RsP3.Fields("VYear") = NextCYr
                   RsP3.Fields("VFirm") = RsP.Fields("Vfirm")
                   RsP3.Fields("VCtrNo") = 1
                   If RsP.Fields("ClBal_Dr") > 0 Then
                      RsP3.Fields("VCrAcCode") = 1
                      RsP3.Fields("VDrAcCode") = RsP.Fields("AcCode")
                      RsP3.Fields("vamt") = RsP.Fields("ClBal_Dr")
                   Else
                      RsP3.Fields("VCrAcCode") = RsP.Fields("AcCode")
                      RsP3.Fields("VDrAcCode") = 1
                      RsP3.Fields("vamt") = RsP.Fields("ClBal_Cr")
                   End If
                   RsP3.Fields("VNoListVYear") = NextCYr
                   RsP3.Fields("SubGroupCd") = 1
                   RsP3.Update
                   RsP.MoveNext
               Loop
            End If
             MsgBox ("Op Tfr Complited")
         End If
         Exit Sub
    Case 2, 3 '--Stock Transfer to Next Year
         If MsgBox("Op. Stock Transfer To " & NextCYr & "  Are you Sure ?", vbYesNo) = vbYes Then
            X = "Select * from tblCompanyYear where CompCode = '" & gCCode & "' and CompYear =" & Right(Year(gCYEDate), 2) + Right(Year(gCYEDate), 2) + 1 & "  "
            If (GProcRstOpen(RsP3, X, "R") = 0) Then
               MsgBox ("Company Year Not Found...")
               Exit Sub
            End If
            gCn.BeginTrans
            gCn.Execute "delete tblPurchSub where tblPurchSub.Vfirm = '" & gCCode & "' and  tblPurchSub.Vtype in ('OY','OI') and tblPurchSub.Vyear = " & NextCYr & "  "
            gCn.Execute "delete tblPurch where tblPurch.Vfirm = '" & gCCode & "' and  tblPurch.Vtype in ('OY','OI') and tblPurch.Vyear = " & NextCYr & "  "
            gCn.CommitTrans

            gCmd.CommandText = "PrcPrepareLotwiseGodownStock"
            gCmd.Parameters.Refresh
            gCmd.Parameters("@VFirm") = Trim(gCCode)
            gCmd.Parameters("@Fr_dt") = Format(gCYSDate, "yyyy/mm/dd")
            gCmd.Parameters("@To_dt") = Format(gCYEDate, "yyyy/mm/dd")
            gCmd.Parameters("@LotNo") = ""
            If Index = 2 Then
                gCmd.Parameters("@StkTP") = "T"
            Else
                gCmd.Parameters("@StkTP") = "D"
            End If
            gCmd.Parameters("@Vno") = ""
            gCmd.Parameters("@Vtype") = ""
            gCmd.Parameters("@Vyear") = ""
            gCmd.Parameters("@IsModi") = "N"
            gCmd.Parameters("@UCode") = gUserName
            gCmd.Execute

            DoEvents
            DoEvents
            DoEvents
            DoEvents
            Y = "Select * from tblPurch"
            GProcRstOpen RsP2, Y, "O"
            z = "Select * from tblPurchSub"
            GProcRstOpen RsP3, z, "O"
            
            X = "Select * from tmpLotwiseStock where UserName = '" & gUserName & "' and vfirm = '" & gCCode & "' order by VFirm,MillAcCode,GodownCd,ItemCode"
            If (GProcRstOpen(RsP, X, "R") > 0) Then
               tmpVno = 40000
               RsP.MoveFirst
               Do While Not RsP.EOF
                  If (RsP.Fields("Amt1") - RsP.Fields("Amt3")) <> 0 Or (RsP.Fields("Amt2") - RsP.Fields("Amt4") <> 0) Then
                      tmpVno = tmpVno + 1
                      RsP2.AddNew
                      RsP2.Fields("Vno") = tmpVno
                      If Index = 2 Then
                         RsP2.Fields("VType") = "OY"
                      Else
                         RsP2.Fields("VType") = "OI"
                      End If
                      RsP2.Fields("Vdt") = NextYrStartDt
                      RsP2.Fields("VYear") = NextCYr
                      RsP2.Fields("VFirm") = RsP.Fields("Vfirm")
                      RsP2.Fields("PurAcCrCode") = 1
                      RsP2.Fields("PurAcDrCode") = 1
                      RsP2.Fields("PurMillCode") = RsP.Fields("MillAcCode")
                      RsP2.Fields("PurBroker") = 1
                      RsP2.Fields("PurSubAmt") = 1
                      RsP2.Fields("PurBillAmt") = 1
                      RsP2.Fields("PurTransport") = 1
                      RsP2.Fields("PurTaxCode") = 1
                      RsP2.Fields("PurTaxPayAc") = 1
                      RsP2.Fields("SubGroupcd") = 1
                      RsP2.Fields("DueDays") = 0
    
    '
                      
                      RsP2.Update
                      
                      LoopStr = RsP.Fields("Vfirm") + Str(RsP.Fields("MillAcCode")) + Str(RsP.Fields("GodownCd"))
                      wCtr = 0
                      Do While (Not RsP.EOF) And (RsP.Fields("Vfirm") + Str(RsP.Fields("MillAcCode")) + Str(RsP.Fields("GodownCd")) = LoopStr)
                         RsP3.AddNew
                         RsP3.Fields("Vno") = tmpVno
                         If Index = 2 Then
                            RsP3.Fields("VType") = "OY"
                         Else
                            RsP3.Fields("VType") = "OI"
                         End If
                         RsP3.Fields("Vdt") = NextYrStartDt
                         RsP3.Fields("VYear") = NextCYr
                         RsP3.Fields("VFirm") = RsP.Fields("Vfirm")
                         wCtr = wCtr + 1
                         RsP3.Fields("PurSubItSrNo") = wCtr
                         RsP3.Fields("PurSubItCode") = RsP.Fields("ItemCode")
                         RsP3.Fields("PurSubBag") = RsP.Fields("Amt1") - RsP.Fields("Amt3")
                         RsP3.Fields("PurSubWt") = RsP.Fields("Amt2") - RsP.Fields("Amt4")
                         RsP3.Fields("PurSubRt") = 1
                         RsP3.Fields("PurSubRtPer") = 1
                         RsP3.Fields("PurSubAmt") = 1
                         RsP3.Fields("PurSubGodown") = RsP.Fields("GodownCd")
                         RsP3.Fields("PurSubLotNo") = RsP.Fields("LotNo")
                         RsP3.Update
                         RsP.MoveNext
                         If RsP.EOF Then GoTo aab
                      Loop
                  Else
                     RsP.MoveNext
                     If RsP.EOF Then GoTo aab
                  End If
               Loop
            End If
aab:
            MsgBox ("Op Stock Tfr Complited")
         End If
         Exit Sub
    
    End Select
End Sub

Private Sub mnutrnGstSub3_Click(Index As Integer)
        gfrmTypeStr = "PR"
        frmPurchaseInward.Show
End Sub

Private Sub mnutrnGstSub5_Click(Index As Integer)
        gfrmTypeStr = "JI"
        frmJVAgtInvSale.Show
        Exit Sub

End Sub
Private Sub mnutrnGstSubBKGSTfrm_Click(Index As Integer)
Set gfrmMainMenu = mnutrnGstSubBKGSTfrm(Index)
Select Case Index
       Case 0 '-- Trade Bill
            gfrmTypeStr = "OT"
            frmBookingTradePurch.Show
       Case 1 '-- Mill Bill
            gfrmTypeStr = "OM"
            frmBookingMillBill.Show
       Case 2
            gfrmTypeStr = "OM"
            frmBookingMillBillDespatchNew.Show
       Case 3
            gfrmTypeStr = "OT"
            frmBookingMillBillDespatchNew.Show
       Case 4
            gfrmTypeStr = "OP"
            frmBookingParty.Show
       Case 5
            gfrmTypeStr = "OD"
            frmDespatchDepot.Show
End Select
End Sub

Private Sub mnutrnGstSubCrNtfrm_Click(Index As Integer)
Select Case Index
Case 0:    '-- Credit Note Sales With GST
        gfrmTypeStr = "PX"
        frmCrnDrnGST.Show
Case 1:    '-- Credit Note Sales
        gfrmTypeStr = "PN"
        frmJV.Show
Case 2:    '--Credit Note Purchase With GST
        gfrmTypeStr = "PV"
        frmCrnDrnGST.Show

Case 3:    '--Credit Note Purchase
End Select
End Sub
Private Sub mnutrnGstSubDailyPricefrm_Click(Index As Integer)
frmPriceList.Show
End Sub
Private Sub mnutrnGstSubDrNtfrm_Click(Index As Integer)
Select Case Index
'--SN--Debit Note Sale (Voucher)
'--PN--Credit Note Sale (Voucher

'--SQ--Debit Note purchase (Voucher)
'--PQ--Credit Note purchase (Voucher

'--SV--Debit Note purchase with GST (intsale)
'--PV--Credit Note purchase with GST (intsale)

'--SI--Debit Note Sale with GST (intsale)



Case 0:    '-- Debit Note Sales With GST
        gfrmTypeStr = "SI"
        gfrmHunkStr = ""
        frmCrnDrnGST.Show
Case 1:    '-- Debit Note Sales
        gfrmTypeStr = "SN"
        frmJV.Show
Case 2:    '-- Debit Note Purchase With GST
        gfrmTypeStr = "SV"
        frmCrnDrnGST.Show
Case 3:    '--
        gfrmTypeStr = "SQ"
        frmJV.Show
End Select
End Sub

Private Sub mnutrnGstSubEpfrm_Click(Index As Integer)
    gfrmTypeStr = "BR"
'    frmLatePaymentIntCredit.Show
    frmLatePaymentIntCreditGST.Show
End Sub

Private Sub mnutrnGstSubInwfrm_Click(Index As Integer)
        gfrmTypeStr = "PI"
        frmPurchaseInward.Show
End Sub

Private Sub mnutrnGstSubJVfrm_Click(Index As Integer)
gfrmTypeStr = "JV"
frmJV.Show
End Sub

Private Sub mnutrnGstSubJVWithTDS_Click(Index As Integer)
gfrmTypeStr = "JT"
frmJVwithTDS.Show
End Sub

Private Sub mnutrnGstSubLpCashfrm_Click(Index As Integer)
    gfrmTypeStr = "CR"
    frmLatePaymentIntGST.Show
End Sub

Private Sub mnutrnGstSubLpfrm_Click(Index As Integer)
    gfrmTypeStr = "BR"
    frmLatePaymentIntGST.Show
End Sub

Private Sub mnutrnGstSubMillPaymentfrm_Click(Index As Integer)
Set gfrmMainMenu = mnutrnGstSubMillPaymentfrm(Index)
Select Case Index
    Case 0 '--Mill Bill
        gfrmTypeStr = "SM"
        gfrmHunkStr = ""
    Case 1 '--Trade Purchase
        gfrmTypeStr = "PY"
        gfrmHunkStr = ""
    Case 2 '--Depo Sales
        gfrmTypeStr = "SD"
        gfrmHunkStr = ""
    Case 3 '--Against Commission-Brokerage Sales
        gfrmTypeStr = "SB"
        gfrmHunkStr = ""
End Select
frmMillPaymentNew.Show
End Sub
Private Sub mnutrnGstSubOsAdj_Click(Index As Integer)
gfrmTypeStr = "AD"
frmOsAdjustment.Show
End Sub

Private Sub mnutrnGstSubPayfrm_Click(Index As Integer)
Set gfrmMainMenu = mnutrnGstSubPayfrm(Index)
Select Case Index
    Case 0 '--Bank
        gfrmTypeStr = "BP"
        frmPayment.Show
    Case 1 '-- Cash
        gfrmTypeStr = "CP"
        frmPayment.Show
End Select
End Sub

Private Sub mnutrnGstSubPurfrm_Click(Index As Integer)
Set gfrmMainMenu = mnutrnGstSubPurfrm(Index)
Select Case Index
    Case 0 '-- Yarn Purchase
        gfrmTypeStr = "PY"
        gfrmHunkStr = ""
        frmPurchaseTradeGST.Show
    Case 1 '-- Other Purchase
        gfrmTypeStr = "PO"
        frmPurchaseOtherGST.Show
End Select
End Sub

Private Sub mnutrnGstSubPurRtn_Click(Index As Integer)
gfrmTypeStr = "VY"
gfrmHunkStr = ""
frmPurchaseReturn.Show
End Sub

Private Sub mnutrnGstSubRecfrm_Click(Index As Integer)
Set gfrmMainMenu = mnutrnGstSubRecfrm(Index)
Select Case Index
    Case 0
        gfrmTypeStr = "BR"
        frmReceipt.Show
    Case 1
        gfrmTypeStr = "CR"
        frmReceipt.Show
    Case 2
        gfrmTypeStr = "MR"
        frmReceipt.Show
End Select
End Sub

Private Sub mnutrnGstSubSalefrm_Click(Index As Integer)
Set gfrmMainMenu = mnutrnGstSubSalefrm(Index)
Select Case Index
    Case 0 '--Sale Trade
        gfrmTypeStr = "SY"
        gfrmHunkStr = ""
        frmSalesGST.Show
    Case 1 '-- Sale Depot
        gfrmTypeStr = "SD"
        gfrmHunkStr = ""
        frmSalesGST.Show
    Case 2 '-- Sale Other
        gfrmTypeStr = "SC"
        gfrmHunkStr = ""
        frmSalesOtherGST.Show
    Case 3 '-- Sale Brokerage & Commission
        gfrmTypeStr = "SB"
        gfrmHunkStr = ""
        frmSalesOtherGST.Show
    Case 4 '-- Mill Bill
        gfrmTypeStr = "SM"
        gfrmHunkStr = ""
'        '-- Sales A/c
'        gClsSearch.SearchMultiField "tblMastaccount", "AcName,AcCity,Accode,AcOurCode", Array("Mill Name", "City", "Code", ""), Array(8000, 3000, 0, 0), "(AgCode=90029 and AcMillBillFirm = '" & gCCode & "') ", , "AcName"
'        If gClsSearch.SearchMultiRetCol(0) = "" Then Exit Sub
        frmSalesMillBill.Show
    Case 6 '-- Sale Late Payment Int
        gfrmTypeStr = "SI"
        gfrmHunkStr = ""
        frmSalesLpInt.Show
    Case 7 '-- Sale Late Payment Int
        gfrmTypeStr = "MI"
        gfrmHunkStr = ""
        frmSalesLpInt.Show
'    Case 8 '-- Old O/s Bills
 '       frmSalesOldOs.Show
    Case 8 '-- Sale Other
        gfrmTypeStr = "SC"
        gfrmHunkStr = "TCS"
        frmSalesOtherGST.Show
End Select
End Sub
Private Sub mnutrnSubGoodsInwFrm_Click(Index As Integer)
        gfrmTypeStr = "PI"
        frmPurchaseInward.Show

End Sub
Private Sub mnutrnSubLpCrNtFrm_Click(Index As Integer)
    gfrmTypeStr = "BR"
    frmLatePaymentIntCredit.Show
End Sub

Private Sub mnutrnSubLpDbNtFrm_Click(Index As Integer)
    gfrmTypeStr = "BR"
    frmLatePaymentInt.Show
End Sub
Private Sub mnuCustRpt_Click(Index As Integer)
    frmReporter.Text1 = mnuCustRpt(Index).Caption
    frmReporter.Show 1
End Sub

Public Sub mnuRptRSubAccLeg_Click(Index As Integer)
gSelectedMenu = ""
Select Case Index
    Case 0 '-"Account Ledger
        gSelectedMenu = "Account Ledger"
    Case 1 '-- Groupwise Account Ledger"
        gSelectedMenu = "Groupwise Account Ledger"
    Case 2 '-- Account Confirmation
        gSelectedMenu = "Account Confirmation"
    Case 3 '-
    Case 4 '-
        gSelectedMenu = "Broker Ledger"
        frmRptTaxReports.Show
        Exit Sub
    Case 5 '-  Bank Reconciliation
        gSelectedMenu = "Bank Reconciliation"
    Case 6 '-"Paid In Next Year"
        gSelectedMenu = "Paid In Next Year"
End Select

If gSelectedMenu <> "" Then frmRptAccount.Show
End Sub

Private Sub mnuRptRSubFin_Click(Index As Integer)
gSelectedMenu = ""
Select Case Index
    Case 0: '-"Balance Sheet"
        gSelectedMenu = "Balance Sheet"
    Case 2: '-"Profit/Loss Account"
        gSelectedMenu = "Profit/Loss Account"
    Case 3  '---"Schedule Printing"
        gSelectedMenu = "Schedule Printing"
    Case 4 '-- All Schedule Printing
        gSelectedMenu = "Opposite A/c wise Summery"
        frmRptAccount.Show
        Exit Sub
    Case 5 '---Expences Shedule
        gSelectedMenu = "Expences Shedule"
    Case 6 '--Loan A/c Details
        gSelectedMenu = "Loan A/c Details"
        frmRptAccount.Show
        Exit Sub
    Case 7 '--Asset A/c Details
        gSelectedMenu = "Asset A/c Details"
        frmRptAccount.Show
        Exit Sub
    Case 8 '--Depreciation Chart
        gSelectedMenu = "Depreciation Chart"
        frmRptAccount.Show
        Exit Sub
End Select
If gSelectedMenu <> "" Then frmRptFinal.Show
End Sub
Private Sub mnuRptRSubFinTri_Click(Index As Integer)
gSelectedMenu = ""
 Select Case Index
    Case 0: '----Simple
        gSelectedMenu = "Trial Balance Simple"
    Case 1: '----Detail
        gSelectedMenu = "Trial Balance Detail"
    Case 2: '----Groupwise
        gSelectedMenu = "Trial Balance Group Simple"
    Case 3: '----Group Detail
        gSelectedMenu = "Trial Balance Group Detail"
    Case 4: '----T-Format
        gSelectedMenu = "T-Format"
    End Select
    frmRptFinal.Show
End Sub
Private Sub mnurptRSubRegPriSale_Click(Index As Integer)
gDepotConsiMillHelp = False
gSelectedMenu = ""
Select Case Index
    Case 0:
        gSelectedMenu = "Trade Sale Register" '----Trade Sale Register
    Case 1:
        gSelectedMenu = "Other Sales" '----Other Sales Register
    Case 2:
        gSelectedMenu = "Consignment/Depot Sale Register" '----Consignment/Depot Sale Register
        gDepotConsiMillHelp = True
    Case 3:
        gSelectedMenu = "Mill Bill Register" '----Mill bill Register
    Case 4:
        gSelectedMenu = "Commission And Brokerage Sales"
End Select
frmRptRegi.Show
End Sub
Private Sub mnutrnSubTDS_Click(Index As Integer)
'Set gfrmMainMenu = mnutrnSubTDS(Index)
Select Case Index
    Case 0 '- TDS Entry
'        frmTDSEntry.Show
End Select
End Sub

Private Sub mnutrnGstSubSlRtn_Click(Index As Integer)
gfrmTypeStr = "RY"
gfrmHunkStr = ""
frmSalesReturn.Show
End Sub

Private Sub mnuUtlSub_Click(Index As Integer)
Dim NextCYr As String
Dim NextDbNm As String
Dim X As String
Dim rstbl As Recordset

Set gfrmMainMenu = mnuUtlSub(Index)
Select Case Index
    Case 0 '-- Company Select
        frmSelCompany.Show 1
        frmMain.Show
    Case 1 '-- Working Year
         X = "Select * from tblCompanyYear where CompCode='" & gCCode & "'"
         GProcRstOpen rstbl, X, "O", gSetupCn
         X = "Select Distinct CompYear  from tblCompanyYear"
         gClsSearch.SearchMultiField X, "", Array("Year"), Array(2000), "", "", "CompYear desc", 0, 0, False, gSetupCn, True
         If gClsSearch.SearchMultiRetCol(0) <> "" Then
            rstbl.MoveFirst
            rstbl.Find "CompYear='" & gClsSearch.SearchMultiRetCol(0) & "'"
            If Not rstbl.EOF Then
               GProcGetCompDetails rstbl
            End If
         End If
'        frmUtlSelectYear.Show
    Case 2 '-- Company Creation
'        frmMastCompany.Show
    Case 3 '- User Security (change Password)
        frmUtlChangePassword.Show
    Case 4 '-  -
'    Case 5 '-- Last Year Outstanding
'        frmUtlLastYrOutstanding.Show
    Case 6 '-  -
    Case 7 '-- Update Balance
        GprocUtlUpdateBalance gCYEDate
    Case 8 '-- Purchase Stock To Next Year
        ''PurchaseStockTransferToNextYear
    Case 9 '-- Transfer Account
'        frmUtlAccountTransfer.Show
    Case 10 '-- -
    Case 11 '- Carry Forword
'        NextCYr = Format(gCYEDate, "yy")
'        NextCYr = NextCYr + Format(CLng(NextCYr) + 1, "00") '- Yr Code
'        NextDbNm = gCCode & NextCYr
'        GProcUtlCarryForward NextDbNm
    Case 12 '- Year End
        GProcUtlYearEnd
    Case 13 '-
    Case 14 '-- Backup Data (For Access )
         BackUpDatabase
 '       If gBackEndDB = gBackEndAccess Then frmUtlBackUpData.Show 1
    Case 15 '-- Table/View Modification
End Select
End Sub
Private Sub mnuUtlSubMod_Click(Index As Integer)
Select Case Index
    Case 0 '-- Modify  View
        If gBackEndDB = gBackEndORA Then
            Call GProcConnectionOpenORA(gCn)
            CreateMofidyViews
        End If
    Case 1 '-- create Table
        If gBackEndDB = gBackEndORA Then
            Call GProcConnectionOpenORA(gCn)
            If gUserName = "hitrix" Then CreateTables
        End If
    Case 2 '-- Modify Table
        ModifyTables
End Select
End Sub
'-- Window
Private Sub mnuWin_Click(Index As Integer)
Select Case Index
    Case 0
        Me.Arrange vbCascade
    Case 1
        Me.Arrange vbHorizontal
    Case 2
        Me.Arrange vbVertical
End Select
End Sub
'-- Exit
Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub BackUpDatabase()
Dim BackupPath As String
Dim BkDbNM As String
BkDbNM = "Vishnu"
Dim DayOfWeek As Integer
DayOfWeek = Weekday(Now) - 1
If DayOfWeek = 0 Then
   DayOfWeek = 7
End If
    '    gCn.Execute "DBCC SHRINKfile ('ModiInv1011_log', 1)"
       ' gCn.Execute "Backup log ModiInv with truncate_only"
     '   gCn.Execute "DBCC SHRINKDATABASE (" & gDbNm & ")"
 'gCn.Execute "DBCC SHRINKDATABASE (" & gDbNm & ")"
    If Dir("E:\BackupHi-Trix", vbDirectory) = "" Then FileSystem.MkDir ("E:\BackupHi-Trix")
    BackupPath = "E:\BackupHi-Trix\Vishnu" & WeekdayName(DayOfWeek, True) & ".Bak"
    gCmd.CommandType = adCmdText
    gCmd.CommandText = "BackUp Database " & BkDbNM + " To Disk = '" + Trim(BackupPath) + "'" + " with init "
    gCmd.Execute
    DoEvents
    'MsgBox ("Backup Complited SuccessFully..............")
End Sub


