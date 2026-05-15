VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000A&
   ClientHeight    =   6900
   ClientLeft      =   165
   ClientTop       =   -6435
   ClientWidth     =   11415
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   Picture         =   "frmMain.frx":058A
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   6600
      Width           =   11415
      _ExtentX        =   20135
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
         Begin VB.Menu mnuMstfrmGSTVAT 
            Caption         =   "GST"
            Index           =   0
         End
         Begin VB.Menu mnuMstfrmGSTVAT 
            Caption         =   "VAT"
            Index           =   1
         End
      End
      Begin VB.Menu mnuMstfrm 
         Caption         =   "&Mill Expense Description"
         Index           =   7
      End
      Begin VB.Menu mnuMstfrm 
         Caption         =   "T&ransport"
         Index           =   8
      End
      Begin VB.Menu mnuMstfrm 
         Caption         =   "Si&zer / Consignee"
         Index           =   9
      End
      Begin VB.Menu mnuMstfrm 
         Caption         =   "Other Item"
         Index           =   10
      End
      Begin VB.Menu mnuMstSetting 
         Caption         =   "&Setting"
         Index           =   0
      End
   End
   Begin VB.Menu mnuGSTtrn 
      Caption         =   "Transaction ( G S T )"
      Index           =   0
      Begin VB.Menu mnuGSTtrnSubBK 
         Caption         =   "&Booking"
         Index           =   0
         Begin VB.Menu mnuGSTtrnSubBKfrm 
            Caption         =   "&Party Booking"
            Index           =   0
         End
         Begin VB.Menu mnuGSTtrnSubBKfrm 
            Caption         =   "&Trade Purchase Booking"
            Index           =   1
         End
      End
      Begin VB.Menu mnuGSTtrnSub1 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuGSTtrnSubGPFrm 
         Caption         =   "&Gate Pass"
         Index           =   0
      End
      Begin VB.Menu mnuGSTtrnSub2 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuGSTtrnSubSale 
         Caption         =   "&Sales"
         Index           =   0
         Begin VB.Menu mnuGSTtrnSubSalefrm 
            Caption         =   "&Trade Sales"
            Index           =   0
         End
         Begin VB.Menu mnuGSTtrnSubSalefrm 
            Caption         =   "&Depot Sales"
            Index           =   1
         End
         Begin VB.Menu mnuGSTtrnSubSalefrm 
            Caption         =   "&Other Sales"
            Index           =   2
         End
      End
      Begin VB.Menu mnuGSTtrnSubPur 
         Caption         =   "&Purchase"
         Index           =   0
         Begin VB.Menu mnuGSTtrnSubPurfrm 
            Caption         =   "&Trade Purchase"
            Index           =   0
         End
         Begin VB.Menu mnuGSTtrnSubPurfrm 
            Caption         =   "&Other Purchase"
            Index           =   1
         End
         Begin VB.Menu mnuGSTtrnSubPurfrm 
            Caption         =   "&Goods Inward"
            Index           =   2
         End
      End
      Begin VB.Menu mnuGSTtrnSub3 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuGSTtrnSubPay 
         Caption         =   "Pa&yment"
         Index           =   0
         Begin VB.Menu mnuGSTtrnSubPayfrm 
            Caption         =   "&Cash"
            Index           =   0
         End
         Begin VB.Menu mnuGSTtrnSubPayfrm 
            Caption         =   "&Bank"
            Index           =   1
         End
      End
      Begin VB.Menu mnuGSTtrnSubRec 
         Caption         =   "&Receipt"
         Index           =   0
         Begin VB.Menu mnuGSTtrnSubRecfrm 
            Caption         =   "&Cash"
            Index           =   0
         End
         Begin VB.Menu mnuGSTtrnSubRecfrm 
            Caption         =   "&Bank"
            Index           =   1
         End
      End
      Begin VB.Menu mnuGSTtrnSubJV 
         Caption         =   "&Journal Entry"
         Index           =   0
         Begin VB.Menu mnuGSTtrnSubJVfrm 
            Caption         =   "&J.V."
            Index           =   0
         End
         Begin VB.Menu mnuGSTtrnSubJVfrm 
            Caption         =   "Credit Note ( Sales )"
            Index           =   1
         End
         Begin VB.Menu mnuGSTtrnSubJVfrm 
            Caption         =   "Debit Note ( Sales )"
            Index           =   2
         End
         Begin VB.Menu mnuGSTtrnSubJVfrm 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu mnuGSTtrnSubJVfrm 
            Caption         =   "Credit Note ( Purchase )"
            Index           =   4
         End
         Begin VB.Menu mnuGSTtrnSubJVfrm 
            Caption         =   "Debit  Note  ( Purchase )"
            Index           =   5
         End
         Begin VB.Menu mnuGSTtrnSubJVfrm 
            Caption         =   "-"
            Index           =   6
         End
         Begin VB.Menu mnuGSTtrnSubJVfrm 
            Caption         =   "Debit  Note  ( TCS )"
            Index           =   7
         End
      End
      Begin VB.Menu mnuGSTtrnSub4 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuGSTtrnSubSlRet 
         Caption         =   "Sales &Return"
         Index           =   0
         Begin VB.Menu mnuGSTtrnSubSlRetfrm 
            Caption         =   "&Trade"
            Index           =   0
         End
         Begin VB.Menu mnuGSTtrnSubSlRetfrm 
            Caption         =   "&Depot"
            Index           =   1
         End
      End
      Begin VB.Menu mnuGSTtrnSubPuRet 
         Caption         =   "P&urchase Return"
         Index           =   0
         Begin VB.Menu mnuGSTtrnSubPuRetfrm 
            Caption         =   "&Trade"
            Index           =   0
         End
         Begin VB.Menu mnuGSTtrnSubPuRetfrm 
            Caption         =   "&Depot"
            Index           =   1
         End
      End
   End
   Begin VB.Menu mnutrn 
      Caption         =   "Transaction ( V A T )"
      Index           =   0
      Begin VB.Menu mnutrnSubBK 
         Caption         =   "&Booking"
         Index           =   0
         Begin VB.Menu mnutrnSubBKfrm 
            Caption         =   "&Mill Booking"
            Index           =   0
         End
         Begin VB.Menu mnutrnSubBKfrm 
            Caption         =   "&SIT Booking"
            Index           =   1
         End
         Begin VB.Menu mnutrnSubBKfrm 
            Caption         =   "&Consignment/Depot Booking"
            Index           =   2
         End
         Begin VB.Menu mnutrnSubBKfrm 
            Caption         =   "&Party Booking"
            Index           =   3
         End
      End
      Begin VB.Menu mnutrnSub1 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnutrnSubGPFrm 
         Caption         =   "&Gate Pass"
         Index           =   0
      End
      Begin VB.Menu mnutrnSub2 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnutrnSubSale 
         Caption         =   "&Sales"
         Index           =   0
         Begin VB.Menu mnutrnSubSalefrm 
            Caption         =   "&Trade Sales"
            Index           =   0
         End
         Begin VB.Menu mnutrnSubSalefrm 
            Caption         =   "&Consignment Sales"
            Index           =   1
         End
         Begin VB.Menu mnutrnSubSalefrm 
            Caption         =   "&Depot Sales"
            Index           =   2
         End
         Begin VB.Menu mnutrnSubSalefrm 
            Caption         =   "&Sale In Transit"
            Index           =   3
         End
         Begin VB.Menu mnutrnSubSalefrm 
            Caption         =   "&Mill Billing"
            Index           =   4
         End
         Begin VB.Menu mnutrnSubSalefrm 
            Caption         =   "Trade Hank Sales"
            Index           =   5
         End
      End
      Begin VB.Menu mnutrnSubPur 
         Caption         =   "&Purchase"
         Index           =   0
         Begin VB.Menu mnutrnSubPurfrm 
            Caption         =   "&Trade Purchase"
            Index           =   0
         End
         Begin VB.Menu mnutrnSubPurfrm 
            Caption         =   "&SIT Purchase "
            Index           =   1
         End
         Begin VB.Menu mnutrnSubPurfrm 
            Caption         =   "&Other Purchase"
            Index           =   2
         End
         Begin VB.Menu mnutrnSubPurfrm 
            Caption         =   "&Goods Inward"
            Index           =   3
         End
         Begin VB.Menu mnutrnSubPurfrm 
            Caption         =   "Trade Hunk Purchase"
            Index           =   4
         End
      End
      Begin VB.Menu mnutrnSub3 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnutrnSubPay 
         Caption         =   "Pa&yment"
         Index           =   0
         Begin VB.Menu mnutrnSubPayfrm 
            Caption         =   "&Cash"
            Index           =   0
         End
         Begin VB.Menu mnutrnSubPayfrm 
            Caption         =   "&Bank"
            Index           =   1
         End
         Begin VB.Menu mnutrnSubPayfrm 
            Caption         =   "&Mill Payment"
            Index           =   2
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnutrnSubRec 
         Caption         =   "&Receipt"
         Index           =   0
         Begin VB.Menu mnutrnSubRecfrm 
            Caption         =   "&Cash"
            Index           =   0
         End
         Begin VB.Menu mnutrnSubRecfrm 
            Caption         =   "&Bank"
            Index           =   1
         End
         Begin VB.Menu mnutrnSubRecfrm 
            Caption         =   "&Mill Receipt"
            Index           =   2
         End
      End
      Begin VB.Menu mnutrnSubJv 
         Caption         =   "&Journal Entry"
         Index           =   0
         Begin VB.Menu mnutrnSubJvfrm 
            Caption         =   "&J.V."
            Index           =   0
         End
         Begin VB.Menu mnutrnSubJvfrm 
            Caption         =   "&Credit Note"
            Index           =   1
         End
         Begin VB.Menu mnutrnSubJvfrm 
            Caption         =   "&Debit Note"
            Index           =   2
         End
      End
      Begin VB.Menu mnutrnSub4 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnutrnSubSlRet 
         Caption         =   "Sales &Return"
         Index           =   0
         Begin VB.Menu mnutrnSubSlRetfrm 
            Caption         =   "&Trade"
            Index           =   0
         End
         Begin VB.Menu mnutrnSubSlRetfrm 
            Caption         =   "&Consignment/Depot"
            Index           =   1
         End
      End
      Begin VB.Menu mnutrnSubPurRet 
         Caption         =   "P&urchase Return"
         Index           =   0
         Begin VB.Menu mnutrnSubPurRetfrm 
            Caption         =   "&Trade"
            Index           =   0
         End
         Begin VB.Menu mnutrnSubPurRetfrm 
            Caption         =   "&Inward"
            Index           =   1
         End
      End
      Begin VB.Menu mnutrnSub5 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnutrnSubform 
         Caption         =   "&Forms Entry"
         Index           =   0
         Begin VB.Menu mnutrnSubFormEtr 
            Caption         =   "Form &C Received"
            Index           =   0
         End
         Begin VB.Menu mnutrnSubFormEtr 
            Caption         =   "Form C &Issued"
            Index           =   1
         End
         Begin VB.Menu mnutrnSubFormEtr 
            Caption         =   "Form &E1 Received"
            Index           =   2
         End
         Begin VB.Menu mnutrnSubFormEtr 
            Caption         =   "Form &F Issued"
            Index           =   3
         End
         Begin VB.Menu mnutrnSubFormEtr 
            Caption         =   "Form E2 Received"
            Index           =   4
         End
         Begin VB.Menu mnutrnSubFormEtr 
            Caption         =   "Form E2 Issued"
            Index           =   5
         End
      End
      Begin VB.Menu mnutrnSub6 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnutrnSubTDS 
         Caption         =   "&TDS Entry"
         Index           =   0
      End
      Begin VB.Menu mnutrnSub7 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnutrnSubBankRecon 
         Caption         =   "&Bank Reconciliation Entry"
         Index           =   0
         Begin VB.Menu mnutrnSubBankReconEtr 
            Caption         =   "&Payment"
            Index           =   0
         End
         Begin VB.Menu mnutrnSubBankReconEtr 
            Caption         =   "&Receipt"
            Index           =   1
         End
      End
      Begin VB.Menu mnutrnSub8 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnutrnSubConfirmFrm 
         Caption         =   "Proforma Invoice"
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
            End
            Begin VB.Menu mnuRptRSubFinTri 
               Caption         =   "&Detail"
               Index           =   1
            End
            Begin VB.Menu mnuRptRSubFinTri 
               Caption         =   "&GroupWise"
               Index           =   2
            End
            Begin VB.Menu mnuRptRSubFinTri 
               Caption         =   "Grp. Detail"
               Index           =   3
            End
            Begin VB.Menu mnuRptRSubFinTri 
               Caption         =   "&T Format"
               Index           =   4
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
            Caption         =   "&All Shedule Print"
            Index           =   4
         End
         Begin VB.Menu mnuRptRSubFin 
            Caption         =   "Monthly Profit & Loss A/C"
            Index           =   5
         End
         Begin VB.Menu mnuRptRSubFin 
            Caption         =   "&Loan Reporting "
            Index           =   6
         End
         Begin VB.Menu mnuRptRSubFin 
            Caption         =   "Loan-Interest-TDS Details"
            Index           =   7
         End
         Begin VB.Menu mnuRptRSubFin 
            Caption         =   "&Countwise Profit and Loss"
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
            Caption         =   "&Mill Billing Ledger"
            Index           =   4
         End
         Begin VB.Menu mnuRptRSubAccLeg 
            Caption         =   "&Bank Reconciliation"
            Index           =   5
         End
         Begin VB.Menu mnuRptRSubAccLeg 
            Caption         =   "Multi Account Ledger"
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
               Caption         =   "&SIT Sales"
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
            Begin VB.Menu mnurptRSubRegPriPur 
               Caption         =   "Purchase Register For ITC"
               Index           =   4
            End
            Begin VB.Menu mnurptRSubRegPriPur 
               Caption         =   "&Insurance Report (SIT)"
               Index           =   5
            End
         End
         Begin VB.Menu mnurptRSubRegPri 
            Caption         =   "&SIT Purchase - Sale Register"
            Index           =   2
         End
         Begin VB.Menu mnurptRSubRegPri 
            Caption         =   "&Trade Purchase - Sale Register"
            Index           =   3
         End
         Begin VB.Menu mnurptRSubRegPri 
            Caption         =   "&Credit Note Register"
            Index           =   4
         End
         Begin VB.Menu mnurptRSubRegPri 
            Caption         =   "&Debit Note Register"
            Index           =   5
         End
         Begin VB.Menu mnurptRSubRegPri 
            Caption         =   "&Internal Voucher Register"
            Index           =   6
         End
         Begin VB.Menu mnurptRSubRegPri 
            Caption         =   "&Mill Bill Receipt Register"
            Index           =   7
         End
         Begin VB.Menu mnurptRSubRegPri 
            Caption         =   "Receipt Register"
            Index           =   8
         End
         Begin VB.Menu mnurptRSubRegPri 
            Caption         =   "&Sales Return"
            Index           =   9
            Begin VB.Menu mnurptRSubRegPriSaleRet 
               Caption         =   "&Trade"
               Index           =   0
            End
            Begin VB.Menu mnurptRSubRegPriSaleRet 
               Caption         =   "&Consignment/Depot"
               Index           =   1
            End
         End
         Begin VB.Menu mnurptRSubRegPri 
            Caption         =   "&Purchase Return"
            Index           =   10
            Begin VB.Menu mnurptRSubRegPriPurRet 
               Caption         =   "&Trade"
               Index           =   0
            End
            Begin VB.Menu mnurptRSubRegPriPurRet 
               Caption         =   "&Consignment/Depot"
               Index           =   1
            End
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
         Begin VB.Menu mnurptRSubRegPri 
            Caption         =   "Payment/RTGS Register"
            Index           =   14
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
               Caption         =   "Broker with &Partywise"
               Index           =   1
            End
            Begin VB.Menu mnurptRSubOutStaSale 
               Caption         =   "-"
               Index           =   2
            End
            Begin VB.Menu mnurptRSubOutStaSale 
               Caption         =   "&Late Payment Interest List"
               Index           =   3
            End
            Begin VB.Menu mnurptRSubOutStaSale 
               Caption         =   "Late Payment Interest &O/S"
               Index           =   4
            End
            Begin VB.Menu mnurptRSubOutStaSale 
               Caption         =   "Interest Demand &Note"
               Index           =   5
            End
            Begin VB.Menu mnurptRSubOutStaSale 
               Caption         =   "Late Payment Interest List (Mill Bill)"
               Index           =   6
            End
            Begin VB.Menu mnurptRSubOutStaSale 
               Caption         =   "Late Payment Interest O/S (Mill Bill)"
               Index           =   7
            End
            Begin VB.Menu mnurptRSubOutStaSale 
               Caption         =   "-"
               Index           =   8
            End
            Begin VB.Menu mnurptRSubOutStaSale 
               Caption         =   "&Freight O/S"
               Index           =   9
            End
            Begin VB.Menu mnurptRSubOutStaSale 
               Caption         =   "&Unadjusted Recepit"
               Index           =   10
            End
            Begin VB.Menu mnurptRSubOutStaSale 
               Caption         =   "Broker && Partywise &Ageing"
               Index           =   11
            End
            Begin VB.Menu mnurptRSubOutStaSale 
               Caption         =   "Broker, Party && Due &Dayswise Ageing"
               Index           =   12
            End
            Begin VB.Menu mnurptRSubOutStaSale 
               Caption         =   "Broker with Partywise - &Mill Bill"
               Index           =   13
            End
            Begin VB.Menu mnurptRSubOutStaSale 
               Caption         =   "Pending GST Amt In L.P. Int "
               Index           =   14
            End
            Begin VB.Menu mnurptRSubOutStaSale 
               Caption         =   "O/s In Date Range"
               Index           =   15
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
               Caption         =   "Late Payment Interest List (Purchase)"
               Index           =   3
            End
         End
      End
      Begin VB.Menu mnuRptRSub 
         Caption         =   "&Stock Reports"
         Index           =   5
         Begin VB.Menu mnurptRSubStock 
            Caption         =   "&Item Ledger"
            Index           =   0
         End
         Begin VB.Menu mnurptRSubStock 
            Caption         =   "Daily Stock To &Mill"
            Index           =   1
         End
         Begin VB.Menu mnurptRSubStock 
            Caption         =   "-"
            Index           =   2
         End
         Begin VB.Menu mnurptRSubStock 
            Caption         =   "&Stock Statement (Inward Type wise)"
            Index           =   3
         End
         Begin VB.Menu mnurptRSubStock 
            Caption         =   "Stock Statement &Datewise"
            Index           =   4
         End
         Begin VB.Menu mnurptRSubStock 
            Caption         =   "-"
            Index           =   5
         End
         Begin VB.Menu mnurptRSubStock 
            Caption         =   "&Purchase/Inwardwise Stock"
            Index           =   6
         End
         Begin VB.Menu mnurptRSubStock 
            Caption         =   "Purchase/Inwardwise &Outward"
            Index           =   7
         End
         Begin VB.Menu mnurptRSubStock 
            Caption         =   "&Pending Purchase SIT"
            Index           =   8
         End
         Begin VB.Menu mnurptRSubStock 
            Caption         =   "-"
            Index           =   9
         End
         Begin VB.Menu mnurptRSubStock 
            Caption         =   "Stock &Valuation (Trade,SIT)"
            Index           =   10
            Begin VB.Menu mnurptRSubStockVal 
               Caption         =   "&Register"
               Index           =   0
            End
            Begin VB.Menu mnurptRSubStockVal 
               Caption         =   "&Summary"
               Index           =   1
            End
         End
      End
      Begin VB.Menu mnuRptRSub 
         Caption         =   "&Tax Reports"
         Index           =   6
         Begin VB.Menu mnurptRSubTax 
            Caption         =   "&VAT Computation"
            Index           =   0
         End
         Begin VB.Menu mnurptRSubTax 
            Caption         =   "VAT Computation &Detail"
            Index           =   1
         End
         Begin VB.Menu mnurptRSubTax 
            Caption         =   "&Exemption Sales Detail"
            Index           =   2
         End
         Begin VB.Menu mnurptRSubTax 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu mnurptRSubTax 
            Caption         =   "&Sales Register"
            Index           =   4
         End
         Begin VB.Menu mnurptRSubTax 
            Caption         =   "&Purchase Register"
            Index           =   5
         End
         Begin VB.Menu mnurptRSubTax 
            Caption         =   "Sales &Return Register"
            Index           =   6
         End
         Begin VB.Menu mnurptRSubTax 
            Caption         =   "Purchase &Return Register"
            Index           =   7
         End
         Begin VB.Menu mnurptRSubTax 
            Caption         =   "-"
            Index           =   8
         End
         Begin VB.Menu mnurptRSubTax 
            Caption         =   "&Form Report"
            Index           =   9
            Begin VB.Menu mnurptRSubTaxForm 
               Caption         =   "C Form &Receivable"
               Index           =   0
            End
            Begin VB.Menu mnurptRSubTaxForm 
               Caption         =   "C Form &Issuable"
               Index           =   1
            End
            Begin VB.Menu mnurptRSubTaxForm 
               Caption         =   "&E1 Form Receivable"
               Index           =   2
            End
            Begin VB.Menu mnurptRSubTaxForm 
               Caption         =   "F Form Issuable"
               Index           =   3
            End
            Begin VB.Menu mnurptRSubTaxForm 
               Caption         =   "-"
               Index           =   4
            End
            Begin VB.Menu mnurptRSubTaxForm 
               Caption         =   "Requirment of &C Form"
               Index           =   5
            End
            Begin VB.Menu mnurptRSubTaxForm 
               Caption         =   "Requirment of &F Form"
               Index           =   6
            End
            Begin VB.Menu mnurptRSubTaxForm 
               Caption         =   "F Form &Statement"
               Index           =   7
            End
            Begin VB.Menu mnurptRSubTaxForm 
               Caption         =   "-"
               Index           =   8
            End
            Begin VB.Menu mnurptRSubTaxForm 
               Caption         =   "&Pending C Form Receivable Slab Wise"
               Index           =   9
            End
            Begin VB.Menu mnurptRSubTaxForm 
               Caption         =   "Pending E1 Form Receivable Slab Wise"
               Index           =   10
            End
            Begin VB.Menu mnurptRSubTaxForm 
               Caption         =   "E2 Form Receivable"
               Index           =   11
            End
            Begin VB.Menu mnurptRSubTaxForm 
               Caption         =   "E2 Form Issuable"
               Index           =   12
            End
            Begin VB.Menu mnurptRSubTaxForm 
               Caption         =   "Requirment of E2 Form Invoicewise"
               Index           =   13
            End
            Begin VB.Menu mnurptRSubTaxForm 
               Caption         =   "Requirment of C Form Invoicewise"
               Index           =   14
            End
            Begin VB.Menu mnurptRSubTaxForm 
               Caption         =   "Requirment of F Form Invoicewise"
               Index           =   15
            End
            Begin VB.Menu mnurptRSubTaxForm 
               Caption         =   "C Form Issued But E1Form Not Received"
               Index           =   16
            End
         End
         Begin VB.Menu mnurptRSubTax 
            Caption         =   "Interest And TDS Details"
            Index           =   10
         End
         Begin VB.Menu mnurptRSubTax 
            Caption         =   "TDS &Summary"
            Index           =   11
         End
         Begin VB.Menu mnurptRSubTax 
            Caption         =   "Service Tax"
            Index           =   12
         End
         Begin VB.Menu mnurptRSubTax 
            Caption         =   "VAT Audit Reports"
            Index           =   13
            Begin VB.Menu mnurptRSubTaxAuditForm 
               Caption         =   "Partywise Sales Details"
               Index           =   0
            End
            Begin VB.Menu mnurptRSubTaxAuditForm 
               Caption         =   "Supplierwise Purchase Details"
               Index           =   1
            End
            Begin VB.Menu mnurptRSubTaxAuditForm 
               Caption         =   "Monthly Sales Details"
               Index           =   2
            End
            Begin VB.Menu mnurptRSubTaxAuditForm 
               Caption         =   "Monthly Purchase Details"
               Index           =   3
            End
            Begin VB.Menu mnurptRSubTaxAuditForm 
               Caption         =   "Annexure - G  (Received C And E1 Form)"
               Index           =   4
            End
            Begin VB.Menu mnurptRSubTaxAuditForm 
               Caption         =   "Annexure - I  (Pending  C And E1 Form)"
               Index           =   5
            End
            Begin VB.Menu mnurptRSubTaxAuditForm 
               Caption         =   "Annexure - J Section -1   (All Partywise VAT Sales)"
               Index           =   6
            End
            Begin VB.Menu mnurptRSubTaxAuditForm 
               Caption         =   "Annexure - J Section -1   (R.D. Partywise VAT Sales)"
               Index           =   7
            End
            Begin VB.Menu mnurptRSubTaxAuditForm 
               Caption         =   "Annexure - J Section -1   (U.R.D. Partywise VAT Sales)"
               Index           =   8
            End
            Begin VB.Menu mnurptRSubTaxAuditForm 
               Caption         =   "Annexure - J Section -2   (Partywise VAT Purchases)"
               Index           =   9
            End
            Begin VB.Menu mnurptRSubTaxAuditForm 
               Caption         =   "Annexure - J Section -6   (Supllierwise CST Purchases)"
               Index           =   10
            End
            Begin VB.Menu mnurptRSubTaxAuditForm 
               Caption         =   "Annexure - J Section -6   (Supllierwise CST Sales)"
               Index           =   11
            End
         End
         Begin VB.Menu mnurptRSubTax 
            Caption         =   "VAT Computation (2 And 5 %)"
            Index           =   14
         End
         Begin VB.Menu mnurptRSubTax 
            Caption         =   "Transactionwise Sales Details"
            Index           =   15
         End
         Begin VB.Menu mnurptRSubTax 
            Caption         =   "Transactionwise Purchase  Details"
            Index           =   16
         End
      End
      Begin VB.Menu mnuRptRSub 
         Caption         =   "SMS"
         Index           =   7
         Begin VB.Menu mnuRptRSubSMS 
            Caption         =   "Form 'C' Receivable SMS (Trade/SIT) All Party"
            Index           =   0
         End
         Begin VB.Menu mnuRptRSubSMS 
            Caption         =   "Form 'C' Receivable SMS (Trade/SIT) Single Party"
            Index           =   1
         End
         Begin VB.Menu mnuRptRSubSMS 
            Caption         =   "Form 'C' Receivable SMS (Trade/SIT) Comman SMS to All Party"
            Index           =   2
         End
         Begin VB.Menu mnuRptRSubSMS 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu mnuRptRSubSMS 
            Caption         =   "Form 'E1' Receivable SMS  All Party"
            Index           =   4
         End
         Begin VB.Menu mnuRptRSubSMS 
            Caption         =   "Form 'E1' Receivable SMS  Single Party"
            Index           =   5
         End
         Begin VB.Menu mnuRptRSubSMS 
            Caption         =   "-"
            Index           =   6
         End
         Begin VB.Menu mnuRptRSubSMS 
            Caption         =   "Form 'E2' Receivable SMS  All Party"
            Index           =   7
         End
         Begin VB.Menu mnuRptRSubSMS 
            Caption         =   "Form 'E2' Receivable SMS  Single Party"
            Index           =   8
         End
         Begin VB.Menu mnuRptRSubSMS 
            Caption         =   "-"
            Index           =   9
         End
         Begin VB.Menu mnuRptRSubSMS 
            Caption         =   "General SMS"
            Index           =   10
         End
         Begin VB.Menu mnuRptRSubSMS 
            Caption         =   "-"
            Index           =   11
         End
         Begin VB.Menu mnuRptRSubSMS 
            Caption         =   "Form 'C' Receivable SMS (Trade/SIT) All Broker"
            Index           =   12
         End
         Begin VB.Menu mnuRptRSubSMS 
            Caption         =   "Form 'C' Receivable SMS (Trade/SIT) Single Broker"
            Index           =   13
         End
         Begin VB.Menu mnuRptRSubSMS 
            Caption         =   "Form 'E1' Receivable SMS (Trade/SIT) All Broker"
            Index           =   14
         End
         Begin VB.Menu mnuRptRSubSMS 
            Caption         =   "Form 'E1' Receivable SMS (Trade/SIT) Single Broker"
            Index           =   15
         End
         Begin VB.Menu mnuRptRSubSMS 
            Caption         =   "Form 'E1' Receivable SMS Comman SMS to All Party"
            Index           =   16
         End
         Begin VB.Menu mnuRptRSubSMS 
            Caption         =   "-"
            Index           =   17
         End
         Begin VB.Menu mnuRptRSubSMS 
            Caption         =   "Outstanding To Party"
            Index           =   18
         End
         Begin VB.Menu mnuRptRSubSMS 
            Caption         =   "Outstanding To Broker"
            Index           =   19
         End
      End
      Begin VB.Menu mnuRptRSub 
         Caption         =   "G S T Reports"
         Index           =   8
         Begin VB.Menu mnuRptRSubGST 
            Caption         =   "Trade Sale Register"
            Index           =   0
         End
         Begin VB.Menu mnuRptRSubGST 
            Caption         =   "R C M Sale Register"
            Index           =   1
         End
         Begin VB.Menu mnuRptRSubGST 
            Caption         =   "Debit Note Register"
            Index           =   2
         End
         Begin VB.Menu mnuRptRSubGST 
            Caption         =   "Purchase Return Register"
            Index           =   3
         End
         Begin VB.Menu mnuRptRSubGST 
            Caption         =   "-"
            Index           =   4
         End
         Begin VB.Menu mnuRptRSubGST 
            Caption         =   "Trade Purchase Register"
            Index           =   5
         End
         Begin VB.Menu mnuRptRSubGST 
            Caption         =   "Other Purchase / Expences Register"
            Index           =   6
         End
         Begin VB.Menu mnuRptRSubGST 
            Caption         =   "Credit Note Register"
            Index           =   7
         End
         Begin VB.Menu mnuRptRSubGST 
            Caption         =   "Sales Return Register"
            Index           =   8
         End
         Begin VB.Menu mnuRptRSubGST 
            Caption         =   "-"
            Index           =   9
         End
         Begin VB.Menu mnuRptRSubGST 
            Caption         =   "Form GSTR - 3B"
            Index           =   10
         End
         Begin VB.Menu mnuRptRSubGST 
            Caption         =   "Form GSTR - 1"
            Index           =   11
         End
         Begin VB.Menu mnuRptRSubGST 
            Caption         =   "Form GSTR - 2"
            Index           =   12
         End
         Begin VB.Menu mnuRptRSubGST 
            Caption         =   "Monthwise Summery"
            Index           =   13
         End
         Begin VB.Menu mnuRptRSubGST 
            Caption         =   "Insurance Report"
            Index           =   14
         End
         Begin VB.Menu mnuRptRSubGST 
            Caption         =   "GSTR - 9"
            Index           =   15
         End
         Begin VB.Menu mnuRptRSubGST 
            Caption         =   "HSN Summery"
            Index           =   16
         End
         Begin VB.Menu mnuRptRSubGST 
            Caption         =   "Partywise Sale/Purchase Summery"
            Index           =   17
         End
         Begin VB.Menu mnuRptRSubGST 
            Caption         =   "GSTR - 2A  -  B2B  (Mismatch)"
            Index           =   18
         End
         Begin VB.Menu mnuRptRSubGST 
            Caption         =   "Purchase And Expences Detail (Audit)"
            Index           =   19
         End
         Begin VB.Menu mnuRptRSubGST 
            Caption         =   "Purchase With Payment Detaill (Audit)"
            Index           =   20
         End
      End
      Begin VB.Menu mnuRptRSub 
         Caption         =   "TDS Report "
         Index           =   9
      End
      Begin VB.Menu mnuRptRSub 
         Caption         =   "Whatsapp Massage"
         Index           =   10
         Begin VB.Menu mnuRptRSubWapp 
            Caption         =   "Outsading To Broker"
            Index           =   0
         End
         Begin VB.Menu mnuRptRSubWapp 
            Caption         =   "Outsading To Party"
            Index           =   1
         End
      End
      Begin VB.Menu mnuRptRSub 
         Caption         =   "TCS  Return"
         Index           =   11
      End
      Begin VB.Menu mnuRptRSub 
         Caption         =   "TCS Report (On Receipt)"
         Index           =   12
      End
      Begin VB.Menu mnuRptRSub 
         Caption         =   "TDS (194Q) Purchase Report"
         Index           =   13
      End
      Begin VB.Menu mnuRptRSub 
         Caption         =   "TCS AND TDS (194Q) Details"
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
            Caption         =   "&Monthwise Sales"
            Index           =   1
         End
         Begin VB.Menu mnuRptGenrSubSale 
            Caption         =   "-"
            Index           =   2
         End
         Begin VB.Menu mnuRptGenrSubSale 
            Caption         =   "&Partywise Sales"
            Index           =   3
         End
         Begin VB.Menu mnuRptGenrSubSale 
            Caption         =   "&Countwise Sales"
            Index           =   4
         End
         Begin VB.Menu mnuRptGenrSubSale 
            Caption         =   "&Millwise Sales"
            Index           =   5
         End
         Begin VB.Menu mnuRptGenrSubSale 
            Caption         =   "-"
            Index           =   6
         End
         Begin VB.Menu mnuRptGenrSubSale 
            Caption         =   "Party && Countwise Sales"
            Index           =   7
         End
         Begin VB.Menu mnuRptGenrSubSale 
            Caption         =   "Mill && Countwise Sales"
            Index           =   8
         End
         Begin VB.Menu mnuRptGenrSubSale 
            Caption         =   "&Broker && Partywise Sales"
            Index           =   9
         End
         Begin VB.Menu mnuRptGenrSubSale 
            Caption         =   "Broker && Countwise Sales"
            Index           =   10
         End
         Begin VB.Menu mnuRptGenrSubSale 
            Caption         =   "-"
            Index           =   11
            Visible         =   0   'False
         End
         Begin VB.Menu mnuRptGenrSubSale 
            Caption         =   "Sales && Purchase Diffrence"
            Index           =   12
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnuRptGenrSub 
         Caption         =   "&Purchase Reports"
         Index           =   7
         Begin VB.Menu mnuRptGenrSubPur 
            Caption         =   "&Total Purchase"
            Index           =   0
         End
         Begin VB.Menu mnuRptGenrSubPur 
            Caption         =   "&Countwise Purchase"
            Index           =   1
         End
         Begin VB.Menu mnuRptGenrSubPur 
            Caption         =   "&Millwise Purchase"
            Index           =   2
         End
         Begin VB.Menu mnuRptGenrSubPur 
            Caption         =   "&Monthwise Sales/Purchase"
            Index           =   3
         End
         Begin VB.Menu mnuRptGenrSubPur 
            Caption         =   "-"
            Index           =   4
         End
         Begin VB.Menu mnuRptGenrSubPur 
            Caption         =   "Total Inward"
            Index           =   5
         End
         Begin VB.Menu mnuRptGenrSubPur 
            Caption         =   "Countwise Inward"
            Index           =   6
         End
         Begin VB.Menu mnuRptGenrSubPur 
            Caption         =   "Millwise Inward"
            Index           =   7
         End
         Begin VB.Menu mnuRptGenrSubPur 
            Caption         =   "Monthwise Purchase"
            Index           =   8
         End
         Begin VB.Menu mnuRptGenrSubPur 
            Caption         =   "Monthwise Sales/Purchase GST"
            Index           =   9
         End
      End
      Begin VB.Menu mnuRptGenrSub 
         Caption         =   "&Brokerage Reports"
         Index           =   8
         Begin VB.Menu mnuRptGenrSubBro 
            Caption         =   "&Detail"
            Index           =   0
         End
         Begin VB.Menu mnuRptGenrSubBro 
            Caption         =   "&Total"
            Index           =   1
         End
         Begin VB.Menu mnuRptGenrSubBro 
            Caption         =   "&Voucher"
            Index           =   2
         End
         Begin VB.Menu mnuRptGenrSubBro 
            Caption         =   "&Inward/Purchase Brokarage Voucher"
            Index           =   3
            Visible         =   0   'False
         End
         Begin VB.Menu mnuRptGenrSubBro 
            Caption         =   "Inward/Purchase Brokarage Detail"
            Index           =   4
         End
         Begin VB.Menu mnuRptGenrSubBro 
            Caption         =   "Inward/Purchase Brokarage Total"
            Index           =   5
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
               Caption         =   "&Mill Booking"
               Index           =   0
            End
            Begin VB.Menu mnuRptGenrSubBkReg 
               Caption         =   "&SIT Booking"
               Index           =   1
            End
            Begin VB.Menu mnuRptGenrSubBkReg 
               Caption         =   "&Consignment/Depot Booking"
               Index           =   2
            End
            Begin VB.Menu mnuRptGenrSubBkReg 
               Caption         =   "&Party Booking"
               Index           =   3
            End
            Begin VB.Menu mnuRptGenrSubBkReg 
               Caption         =   "&Trade Purchase Booking"
               Index           =   4
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
               Caption         =   "&SIT Booking"
               Index           =   1
            End
            Begin VB.Menu mnuRptGenrSubBkStk 
               Caption         =   "&Consignment/Depot Booking"
               Index           =   2
            End
            Begin VB.Menu mnuRptGenrSubBkStk 
               Caption         =   "&Party Booking"
               Index           =   3
            End
            Begin VB.Menu mnuRptGenrSubBkStk 
               Caption         =   "&Trade Purchase Booking"
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
               Caption         =   "&SIT Booking"
               Index           =   1
            End
            Begin VB.Menu mnuRptGenrSubBkCancel 
               Caption         =   "&Consignment/Depot Booking"
               Index           =   2
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
               Caption         =   "&SIT Booking"
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
            Begin VB.Menu mnuRptGenrSubBkVs 
               Caption         =   "Trade Purchase Booking"
               Index           =   4
            End
         End
      End
      Begin VB.Menu mnuRptGenrSub 
         Caption         =   "&Godown Reports"
         Index           =   10
         Begin VB.Menu mnuRptGenrSubGP 
            Caption         =   "&Pending Gate Pass "
            Index           =   0
         End
         Begin VB.Menu mnuRptGenrSubGP 
            Caption         =   "&Gate Pass Register"
            Index           =   1
         End
         Begin VB.Menu mnuRptGenrSubGP 
            Caption         =   "&Delivery Register"
            Index           =   2
         End
         Begin VB.Menu mnuRptGenrSubGP 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu mnuRptGenrSubGP 
            Caption         =   "Godownwise &Inward Register"
            Index           =   4
         End
         Begin VB.Menu mnuRptGenrSubGP 
            Caption         =   "Goods Delivery Report"
            Index           =   5
            Begin VB.Menu mnuRptGenrSubGPDel 
               Caption         =   "&CountWise"
               Index           =   0
            End
            Begin VB.Menu mnuRptGenrSubGPDel 
               Caption         =   "Count && PartyWise"
               Index           =   1
            End
            Begin VB.Menu mnuRptGenrSubGPDel 
               Caption         =   "Count && SizerWise"
               Index           =   2
            End
            Begin VB.Menu mnuRptGenrSubGPDel 
               Caption         =   "&Party && Bag No wise Register"
               Index           =   3
            End
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
         End
      End
      Begin VB.Menu mnuRptGenrSub 
         Caption         =   "-"
         Index           =   13
      End
      Begin VB.Menu mnuRptGenrSub 
         Caption         =   "&MIS Reports"
         Index           =   14
         Begin VB.Menu mnuRptGenrSubMis 
            Caption         =   "&Cash and Bank Balance"
            Index           =   0
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
      End
      Begin VB.Menu mnuRptGenrSub 
         Caption         =   "&Graph"
         Index           =   18
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
         Caption         =   "&Last Year Sales Outstanding"
         Index           =   5
      End
      Begin VB.Menu mnuUtlSub 
         Caption         =   "&Last Year Purchase Outstanding"
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
            Caption         =   "Transfer Op.Balance A/c.  to Next Year"
            Index           =   0
         End
         Begin VB.Menu mnuTransfer 
            Caption         =   "Transfer Outstanding  to Next Year"
            Index           =   1
         End
         Begin VB.Menu mnuTransfer 
            Caption         =   "Transfer Stock  to Next Year"
            Index           =   2
         End
         Begin VB.Menu mnuTransfer 
            Caption         =   "Transfer Address Deails to Next Year"
            Index           =   3
         End
         Begin VB.Menu mnuTransfer 
            Caption         =   "Transfer Trade Purchase Booking"
            Index           =   4
         End
         Begin VB.Menu mnuTransfer 
            Caption         =   "Transfer Party Booking  to Next Year"
            Index           =   5
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
         Begin VB.Menu mnuUtlSubMod 
            Caption         =   "Create/Modify &Views!"
            Index           =   0
            Visible         =   0   'False
         End
         Begin VB.Menu mnuUtlSubMod 
            Caption         =   "&Create Tables!"
            Index           =   1
            Visible         =   0   'False
         End
         Begin VB.Menu mnuUtlSubMod 
            Caption         =   "&Modify Tables!"
            Index           =   2
         End
      End
      Begin VB.Menu mnuUtlSub 
         Caption         =   "For Eof() Bof() Error"
         Index           =   16
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      Begin VB.Menu mnuWin 
         Caption         =   "&Cascade"
         Index           =   0
      End
      Begin VB.Menu mnuWin 
         Caption         =   "Tile &Horizontal"
         Index           =   1
      End
      Begin VB.Menu mnuWin 
         Caption         =   "Tile &Vertical"
         Index           =   2
      End
      Begin VB.Menu mnuWin 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuWin 
         Caption         =   "Window &List"
         Index           =   4
         WindowList      =   -1  'True
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
   i = GProcRstOpen(rst1, "select REPONAME, RDTRANGE from UDFREPORTS where repostatus = 0", "R", gSetupCn)
   If i > 0 Then
        With rst1
            i = 0
            .MoveFirst
            Do While Not .EOF
               If (i > 0) Then Load mnuCustRpt(i)
               mnuCustRpt(i).Caption = .Fields(0)
               mnuCustRpt(i).Visible = True
               i = i + 1
              .MoveNext
            Loop
        End With
    End If
    gMyConn = "gCn"
    ProcAllowFrmMainMenu
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

Private Sub mnuGSTtrnSubBKfrm_Click(Index As Integer)
Set gfrmMainMenu = mnuGSTtrnSubBKfrm(Index)
Select Case Index
    Case 0 '--Party
        gfrmTypeStr = "OP"
        frmBookingPartyNew.Show
    Case 1 '-- Trade Purchase Booking
        gfrmTypeStr = "OT"
        frmBookingTrade.Show
End Select
End Sub

Private Sub mnuGSTtrnSubGPFrm_Click(Index As Integer)
       If gCYear >= 2025 Then
          frmGatepassNew.Show
       Else
          frmGatepass.Show
       End If
End Sub

Private Sub mnuGSTtrnSubJVfrm_Click(Index As Integer)
Set gfrmMainMenu = mnuGSTtrnSubJVfrm(Index)
Select Case Index
    Case 0
        gfrmTypeStr = "JV"
        frmJV.Show
    Case 1
        gfrmHunkStr = "B"
        gfrmTypeStr = "PN"
        frmCrDrNoteGST.Show
    Case 2
        gfrmHunkStr = "B"
        gfrmTypeStr = "SN"
        frmCrDrNoteGST.Show
    Case 4
        gfrmHunkStr = "S"
        gfrmTypeStr = "PN"
        frmCrDrNoteGST.Show
    Case 5
        gfrmHunkStr = "S"
        gfrmTypeStr = "SN"
        frmCrDrNoteGST.Show
    Case 7
        gfrmHunkStr = "T"
        gfrmTypeStr = "SN"
        frmCrDrNoteGST.Show
End Select
End Sub

Private Sub mnuGSTtrnSubPayfrm_Click(Index As Integer)
Set gfrmMainMenu = mnuGSTtrnSubPayfrm(Index)
Select Case Index
    Case 0 '--Cash
        gfrmTypeStr = "CP"
        frmPaymentGST.Show
    Case 1 '-- Bank
        gfrmTypeStr = "BP"
        frmPaymentGST.Show
End Select

End Sub

Private Sub mnuGSTtrnSubPuRetfrm_Click(Index As Integer)
Set gfrmMainMenu = mnuGSTtrnSubPuRetfrm(Index)
Select Case Index
    Case 0 '-- Account
        gfrmTypeStr = "VY"
        frmPurchRtnTradeGST.Show
End Select
End Sub

Private Sub mnuGSTtrnSubPurfrm_Click(Index As Integer)
Set gfrmMainMenu = mnuGSTtrnSubPurfrm(Index)
Select Case Index
    Case 0 '-- Yarn Purchase
        gfrmTypeStr = "PY"
        frmPurchaseTradeGST.Show
    Case 1 '-- Other Purchase
        gfrmTypeStr = "PO"
        frmPurchaseOtherGST.Show
    Case 2 '-- Goods Inward
        gfrmTypeStr = "PI"
        frmPurchaseTradeGST.Show
End Select
End Sub

Private Sub mnuGSTtrnSubRecfrm_Click(Index As Integer)
Set gfrmMainMenu = mnuGSTtrnSubRecfrm(Index)
Select Case Index
    Case 0
        gfrmTypeStr = "CR"
        frmReceiptGST.Show
    Case 1
        gfrmTypeStr = "BR"
        frmReceiptGST.Show
End Select
End Sub
Private Sub mnuGSTtrnSubSalefrm_Click(Index As Integer)
Set gfrmMainMenu = mnuGSTtrnSubSalefrm(Index)
Select Case Index
    Case 0 '--Sale Trade
        gfrmTypeStr = "SY"
        frmSalesGST.Show
    Case 1 '-- Sale Depot
        gfrmTypeStr = "SD"
        '-- Sales A/c
        gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,Accode,AcCity,AcMillTypeStr", Array("Sales Account", "AcName", "acCode", "City", "Mill Type"), Array(3000, 0, 0, 2000, 2000), " ((AgCode=90029 or GpCode=90029) and AcMillType=2)", , "AcName"
        If gClsSearch.SearchMultiRetCol(0) = "" Then Exit Sub
        frmSalesGST.Show
    Case 2  '-- Other Sale
        gfrmHunkStr = "OTH"
        gfrmTypeStr = "SY"
        frmSalesOtherGST.Show
End Select
End Sub

Private Sub mnuGSTtrnSubSlRetfrm_Click(Index As Integer)
Set gfrmMainMenu = mnuGSTtrnSubSlRetfrm(Index)
Select Case Index
    Case 0 '-- Account
        gfrmTypeStr = "RY"
        frmSaleRtnTradeGST.Show
End Select
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
'    Case 6 '-- Tax Type
'        'gfrmTypeStr = "T"
'        frmMastNarrationTax.Show
    Case 7 '-- Mill Expenses
        gfrmTypeStr = "M"
        frmMastNarration.Show
    Case 8 '-- Transport Type
        gfrmTypeStr = "R"
        frmMastNarration.Show
    Case 9 '-- Sizer
        gfrmTypeStr = "S"
        frmMastNarration.Show
    Case 10 '-- Other Item
        frmMastItemOther.Show
End Select
End Sub

Private Sub mnuMstfrmGSTVAT_Click(Index As Integer)
Select Case Index
    Case 0 '-- GST
        frmMastNarrationTaxGST.Show
    Case 1 '-- VAT
        frmMastNarrationTax.Show
End Select
End Sub

Private Sub mnuMstSetting_Click(Index As Integer)
Select Case Index
    Case 0 '-- Setting
        frmSetting.Show
End Select
End Sub

Private Sub mnuRptGenrSubBkCancel_Click(Index As Integer)
gSelectedMenu = ""
Select Case Index
    Case 0: '-- Canceled Mill Booking Register
        gSelectedMenu = "Canceled Mill Booking Register"
    Case 1 '--  Canceled SIT Booking Register
        gSelectedMenu = "Canceled SIT Booking Register"
    Case 2: '-- Canceled Consignment/Depot Booking Register
        gSelectedMenu = "Canceled Consignment/Depot Booking Register"
    Case 3: '-- Canceled Party Booking Register
        gSelectedMenu = "Canceled Party Booking Register"
End Select
If gSelectedMenu <> "" Then frmRptRegi.Show
End Sub

Private Sub mnuRptGenrSubBkReg_Click(Index As Integer)
gSelectedMenu = ""
Select Case Index
    Case 0: '-- Mill Booking Register
        gSelectedMenu = "Mill Booking Register"
    Case 1 '--  SIT Booking Register
        gSelectedMenu = "SIT Booking Register"
    Case 2: '-- Consignment/Depot Booking Register
        gSelectedMenu = "Consignment/Depot Booking Register"
    Case 3: '-- Party Booking Register
        gSelectedMenu = "Party Booking Register"
    Case 4 '--  Trade Purchase Booking Register
        gSelectedMenu = "Trade Purchase Booking Register"
End Select
If gSelectedMenu <> "" Then frmRptRegi.Show
End Sub

Private Sub mnuRptGenrSubBkStk_Click(Index As Integer)
gSelectedMenu = ""
Select Case Index
    Case 0: '-- Pending Mill Booking Register
        gSelectedMenu = "Pending Mill Booking Register"
    Case 1 '--  Pending SIT Booking Register
        gSelectedMenu = "Pending SIT Booking Register"
    Case 2: '-- Pending Consignment/Depot Booking Register
        gSelectedMenu = "Pending Consignment/Depot Booking Register"
    Case 3: '-- Pending Party Booking Register
        gSelectedMenu = "Pending Party Booking Register"
    Case 4 '--  Pending Trade Purchase Booking Register
        gSelectedMenu = "Pending Trade Purchase Booking Register"
End Select
If gSelectedMenu <> "" Then frmRptRegi.Show
End Sub

Private Sub mnuRptGenrSubBkVs_Click(Index As Integer)
gSelectedMenu = ""
Select Case Index
    Case 0: '--  Mill Booking Vs Dispatch Register
        gSelectedMenu = "Mill Booking Vs Dispatch Register"
    Case 1 '--  SIT Booking Vs Despatch Register
        gSelectedMenu = "SIT Booking Vs Dispatch Register"
    Case 2: '--  Consignment/Depot Booking Vs Dispatch Register
        gSelectedMenu = "Consignment/Depot Booking Vs Dispatch Register"
    Case 3: '-- Party Booking Vs Dispatch Register
        gSelectedMenu = "Party Booking Vs Dispatch Register"
    Case 4 '--  Trade Booking Vs Despatch Register
        gSelectedMenu = "Trade Booking Vs Dispatch Register"
End Select
If gSelectedMenu <> "" Then frmRptRegi.Show
End Sub

Private Sub mnuRptGenrSubBro_Click(Index As Integer)
gSelectedMenu = ""
Select Case Index
    Case 0: '-- Brokerage Detail Report
        gSelectedMenu = "Brokerage Detail Report"
    Case 1 '--  Brokerage Total Report
        gSelectedMenu = "Brokerage Total Report"
    Case 2: '-- Brokerage Voucher Report
        gSelectedMenu = "Brokerage Voucher Report"
    Case 3: '-- Inward/Purchase Brokerage Report
        gSelectedMenu = "Inward/Purchase Brokerage Report"
    Case 4: '-- Inward/Purchase Brokerage Detail Report
        gSelectedMenu = "Inward/Purchase Brokerage Detail Report"
    Case 5: '-- Inward/Purchase Brokarage Total
        gSelectedMenu = "Inward/Purchase Brokarage Total"
End Select
If gSelectedMenu <> "" Then frmRptTaxReports.Show
End Sub

Private Sub mnuRptGenrSubCLBal_Click(Index As Integer)
gSelectedMenu = ""
Select Case Index
   Case 0 '-"Account Closing Balance List"
        gSelectedMenu = "Account Closing Balance List"
        frmRptAccount.Show
        Exit Sub
   Case 1 '-"Account Groupwise Closing Balance List"
        gSelectedMenu = "Account Groupwise Closing Balance List"
        frmRptAccount.Show
        Exit Sub
End Select
End Sub

Private Sub mnuRptGenrSubGP_Click(Index As Integer)
gSelectedMenu = ""
Select Case Index
     Case 0: '-- Pending Gate Pass Report
        gSelectedMenu = "Pending Gate Pass Report"
     Case 1 '-- Gate Pass Register
        gSelectedMenu = "Gate Pass Register"
     Case 2 '-- Delivery Register
        gSelectedMenu = "Delivery Register"
     Case 3 '-
     Case 4 '- "Godownwise Inward Register"
        gSelectedMenu = "Godownwise Inward Register"
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
Select Case Index
    Case 0 '- Sales Graph
        gSelectedMenu = "Sales Graph"
    Case 1 '- Countwise Sales Graph
        gSelectedMenu = "Countwise Sales Graph"
    Case 2 '- Millwise Sales Graph
        gSelectedMenu = "Millwise Sales Graph"
    Case 3 '- Brokerwise Sales Graph
        gSelectedMenu = "Brokerwise Sales Graph"
    Case 4  '-
    Case 5 '- Purchase Graph
        gSelectedMenu = "Purchase Graph"
    Case 6 '- Countwise Purchase Graph
        gSelectedMenu = "Countwise Purchase Graph"
    Case 7 '- Millwise Purchase Graph
        gSelectedMenu = "Millwise Purchase Graph"
End Select
If gSelectedMenu <> "" Then frmRptGraph.Show
End Sub

Private Sub mnuRptGenrSubInt_Click(Index As Integer)
gSelectedMenu = ""
Select Case Index
 Case 0 '-- Sarafi Interest Report
    gSelectedMenu = "Sarafi Interest Report"
 Case 1 '-- Sarafi Interest Report Ank
    gSelectedMenu = "Sarafi Interest Report ANK"
End Select
If gSelectedMenu <> "" Then frmRptAccount.Show
End Sub

Public Sub mnuRptGenrSubLst_Click(Index As Integer)
gSelectedMenu = ""
Select Case Index
Case 0 '- Account List
    gSelectedMenu = "Account List"
    frmRptAccount.Show
'    GProcCrystalRptPreparation cryReport, "Window"
'    cryReport.WindowTitle = "Account List"
'    cryReport.ReportFileName = gReportPath & "RptLstAccount.rpt"
'    cryReport.Action = 1
    Exit Sub
Case 1 '- Groupwise Account List
    gSelectedMenu = "Groupwise Account List"
    frmRptAccount.Show
'    GProcCrystalRptPreparation cryReport, "Window"
'    cryReport.WindowTitle = "Groupwise Account List"
'    cryReport.ReportFileName = gReportPath & "RptLstAccountGroupwise.rpt"
'    cryReport.Action = 1
    Exit Sub
Case 2 '-Selected Group Accounts List
    gSelectedMenu = "Selected Group Accounts List"
    frmRptAccount.Show
    Exit Sub
Case 3 '- -
Case 4 '- Count Rate List
    GProcCrystalRptPreparation CryReport, "Window"
    CryReport.WindowTitle = "Count Rate List"
    CryReport.ReportFileName = gReportPath & "RptLstCountRateList.rpt"
    CryReport.Action = 1
    Exit Sub
Case 5 '- Millwise Count List
    gSelectedMenu = "Millwise Count List"
    frmRptAccount.Show
    Exit Sub
Case 6 '- Count Brokerage List
    gSelectedMenu = "Count Brokerage List"
    frmRptAccount.Show
    Exit Sub
End Select
End Sub

Private Sub mnuRptGenrSubMis_Click(Index As Integer)
   gSelectedMenu = ""
   Select Case Index  '-- Missing Transaction
          Case 0 '---"Cash - Bank  Status"
               gSelectedMenu = "Cash - Bank  Status"
          End Select
   If gSelectedMenu <> "" Then frmRptTaxReports.Show
End Sub

Private Sub mnuRptGenrSubMissTrans_Click(Index As Integer)
gSelectedMenu = ""
Select Case Index  '-- Missing Transaction
    Case 0  '-- Booking
        gSelectedMenu = "Missing Transaction Report - Booking"
    Case 1 '-- Purchase
        gSelectedMenu = "Missing Transaction Report - Purchase"
    Case 2 '-- Gate Pass
        gSelectedMenu = "Missing Transaction Report - Gate Pass"
    Case 3 '-- Sales
        gSelectedMenu = "Missing Transaction Report - Sales"
    Case 4 '-- Vouchers
        gSelectedMenu = "Missing Transaction Report - Vouchers"
End Select
If gSelectedMenu <> "" Then frmRptRegi.Show
End Sub

Private Sub mnuRptGenrSubMon_Click(Index As Integer)
gSelectedMenu = ""
Select Case Index
     Case 0: '-- Consignment/ Depot Statement
       gSelectedMenu = "Consignment/ Depot Statement"
     Case 1 '-- Mill Billing Statement
       gSelectedMenu = "Mill Billing Statement"
     Case 2 '-- Sales In Transit Statement
       gSelectedMenu = "Sales In Transit Statement"
End Select
If gSelectedMenu <> "" Then frmRptStatement.Show
End Sub

Private Sub mnuRptGenrSubOPBal_Click(Index As Integer)
gSelectedMenu = ""
Select Case Index
   Case 0 '-"Account Opening Balance List"
        GProcCrystalRptPreparation CryReport, "Window"
        CryReport.WindowTitle = "Account Opening Balance List"
        CryReport.SelectionFormula = "{tblVoucher.VType}='0P'"
        CryReport.ReportFileName = gReportPath & "RptLstAccOpBal.rpt"
        CryReport.Action = 1
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
        gSelectedMenu = "Total Purchase Summary"
    Case 1 '--"Countwise Purchase"
        gSelectedMenu = "Countwise Purchase Summary"
    Case 2 '--"Millwise Purchase"
        gSelectedMenu = "Millwise Purchase Summary"
    Case 3 '--Monthwise Sales/Purchase
        gSelectedMenu = "Monthwise Sales/Purchase Summary"
    Case 4 '-- -
    Case 5 '--Total Inward
        gSelectedMenu = "Total Inward Summary"
    Case 6 '--"Countwise Inward
        gSelectedMenu = "Countwise Inward Summary"
    Case 7 '--"Millwise Inward
        gSelectedMenu = "Millwise Inward Summary"
    Case 8 '--"Monthwise Purchase Summary"
        gSelectedMenu = "Monthwise Purchase Summary"
    Case 9 '--"Monthwise Sales/Purchase GST
        gSelectedMenu = "Monthwise Sales/Purchase New"
End Select
If gSelectedMenu <> "" Then frmRptTaxReports.Show
End Sub
Private Sub mnuRptGenrSubSale_Click(Index As Integer)
gSelectedMenu = ""
Select Case Index
    Case 0: '-- "Total Sales"
        gSelectedMenu = "Total Sales Summary"
    Case 1 '-- Monthwise Sales
        gSelectedMenu = "Monthwise Sales Summary"
    Case 2 '- -
    Case 3 '-- Partywise Sales
        gSelectedMenu = "Partywise Sales Summary"
    Case 4 '-- Countwise Sales
        gSelectedMenu = "Countwise Sales Summary"
    Case 5: '-- Millwise Sales
        gSelectedMenu = "Millwise Sales Summary"
    Case 6 '- -
    Case 7: '-- Party and Countwise Sales
        gSelectedMenu = "Party and Countwise Sales Summary"
    Case 8: '-- Mill & Countwise Sales
        gSelectedMenu = "Mill and Countwise Sales Summary"
    Case 9: '-- Broker and Partywise Sales
        gSelectedMenu = "Broker and Partywise Sales Summary"
    Case 10: '-- "Broker & Countwise Sales"
        gSelectedMenu = "Broker and Countwise Sales Summary"
End Select
If gSelectedMenu <> "" Then frmRptTaxReports.Show
End Sub

Private Sub mnuRptRSub_Click(Index As Integer)
If Index = 9 Then
    gSelectedMenu = "TDS Report"
    frmRptRegi.Show
End If
If Index = 11 Then
   gSelectedMenu = "TCS  Return"
   frmRptGraph.Show
End If
If Index = 12 Then
   gSelectedMenu = "TCS Report"
   frmRptAccount.Show
End If
If Index = 13 Then
   gSelectedMenu = "TDS (194Q) Purchase Report"
   frmRptGraph.Show
End If
If Index = 14 Then
   gSelectedMenu = "TCS AND TDS (194Q) Details"
   frmRptGraph.Show
End If
End Sub

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
If gSelectedMenu <> "" Then frmRptAccount.Show
End Sub

Private Sub mnuRptRSubGST_Click(Index As Integer)
gSelectedMenu = ""
Select Case Index
    Case 0:
        gSelectedMenu = "Trade Sale Register GST"
    Case 1:
        gSelectedMenu = "R C M Sale Register"
    Case 2:
        gSelectedMenu = "Debit Note Register GST"
    Case 3:
        gSelectedMenu = "Purchase Return Register"
    Case 5:
        gSelectedMenu = "Trade Purchase Register GST"
    Case 6:
        gSelectedMenu = "Other Purchase Register GST"
    Case 7:
        gSelectedMenu = "Credit Note Register GST"
    Case 8:
        gSelectedMenu = "Sales Return Register"
    Case 10:
        gSelectedMenu = "Form GSTR - 3B"
    Case 11:
        gSelectedMenu = "Form GSTR - 1"
    Case 12:
        gSelectedMenu = "Form GSTR - 2"
    Case 13:
        gSelectedMenu = "Monthwise Summery GST"
    Case 14:
        gSelectedMenu = "Insurance Report"
    Case 15:
        gSelectedMenu = "GSTR - 9"
    Case 16:
        gSelectedMenu = "HSN Summery"
    Case 17:
        gSelectedMenu = "Partywise Sale/Purchase Summery"
    Case 18:
        gSelectedMenu = "GSTR - 2A  -  B2B  (Mismatch)"
    Case 19:
        gSelectedMenu = "Purchase And Expences Detail (Audit)"
    Case 20:
         gSelectedMenu = "Purchase With Payment Detaill (Audit)"
End Select
frmRptGST.Show
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
     Case 3: '-"Late Payment Interest Purchase List"
       gSelectedMenu = "Late Payment Interest Purchase List"
       frmRptSelection.Show
       Exit Sub
End Select
If gSelectedMenu <> "" Then frmRptRegi.Show
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
        gSelectedMenu = "Outstanding Broker with Partywise"
    Case 2 '-
    Case 3:
       gRptPartyFlag = True
       gSelectedMenu = "Late Payment Interest Register"
      ' frmRptSelection.CommonReportSelection True, True, False, True, False, False
       frmRptSelection.Show
       Exit Sub
    Case 4 '-"Outstanding Late Payment Interest Report"
        gSelectedMenu = "Outstanding Late Payment Interest Report"
    Case 5 '- Interest Demand Note
       gRptPartyFlag = True
       gSelectedMenu = "Interest Demand Note"
       frmRptSelection.Show
       Exit Sub
    Case 6 '-"Late Payment Interest Register (Mill Bill)"
       gRptPartyFlag = True
       gSelectedMenu = "Late Payment Interest Register (Mill Bill)"
       frmRptSelection.Show
       Exit Sub
    Case 7 '-"Outstanding Late Payment Interest Report (Mill Bill)"
        gSelectedMenu = "Outstanding Late Payment Interest Report (Mill Bill)"
    Case 8  '-
    Case 9:
        'gSelectedMenu = "Freight Outstanding Report"
    Case 10:
       gSelectedMenu = "Unadjusted Recepit"
    Case 11 '--"Broker  & Partywise Aging Report"
       gSelectedMenu = "Broker & Partywise Aging Report"
       gRptBrokerFlag = True
       gRptOnDateFlag = True
       frmRptSelection.Show
       Exit Sub
   Case 12 '--"Broker ,Party & Due Days wise Aging Report"
       gSelectedMenu = "Broker, Party & Due Dayswise Aging Report"
       gRptBrokerFlag = True
       gRptOnDateFlag = True
       frmRptSelection.Show
       Exit Sub
   Case 13 '--"Mill Bill Outstanding Broker with Partywise"
       gSelectedMenu = "Mill Bill Outstanding Broker with Partywise"
   Case 14:
       gSelectedMenu = "Pending GST Amt In L.P. Int "
   Case 15
       gSelectedMenu = "O/s In Date Range"
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
        gSelectedMenu = "Sale In Transit Register"
    Case 3 '----"Trade Purchase - Sale Register"
        gSelectedMenu = "Trade Purchase - Sale Register"
    Case 4 '----"Credit Note Register"
        gSelectedMenu = "Credit Note Register"
    Case 5  '----"Debit Note Register"
        gSelectedMenu = "Debit Note Register"
    Case 6  '----"Internal Voucher Register"
        gSelectedMenu = "Internal Voucher Register"
    Case 7  '----Mill Bill Receipt Register
        gSelectedMenu = "Mill Bill Receipt Register"
    Case 8 '-
         gSelectedMenu = "Receipt Register"
    Case 9 '-Sale Return
    Case 10  '-Purchase Return
    Case 11 '-
    Case 12 '----"Late Pay Debit Note Register"
        gSelectedMenu = "Late Pay Debit Note Register"
    Case 13  '--"Late Pay Debit Note Register (Mill Bill)"
        gSelectedMenu = "Late Pay Debit Note Register (Mill Bill)"
    Case 14  '----"Payment/RTGS Register"
        gSelectedMenu = "Payment/RTGS Register"
End Select
If gSelectedMenu <> "" Then frmRptRegi.Show
End Sub

Private Sub mnurptRSubRegPriPur_Click(Index As Integer)
gSelectedMenu = ""
Select Case Index
    Case 0:
        gSelectedMenu = "Trade Purchase Register" '----Trade Purchase Register
    Case 1:
        gSelectedMenu = "SIT Purchase Register" '----SIT Purchase Register
    Case 2:
        gSelectedMenu = "Other Purchase Register" '----Other Purchase Register
    Case 3:
        gSelectedMenu = "Consignment/Depot Purchase Register" '----Consignment/Depot Purchase Register
    Case 4:
        gSelectedMenu = "Purchase Register For ITC"
    Case 5:
        gSelectedMenu = "Insurance Report(SIT)"
End Select
If gSelectedMenu <> "" Then frmRptRegi.Show
End Sub

Private Sub mnurptRSubRegPriPurRet_Click(Index As Integer)
gSelectedMenu = ""
Select Case Index
    Case 0
        gSelectedMenu = "Trade Purchase Return Register" '----Trade Purchase Return Register
    Case 1
        gSelectedMenu = "Consignment/Depot Purchase Return Register" '----Consignment/Depot Purchase Return Register
End Select
'frmRptRegi.Show
End Sub

Private Sub mnurptRSubRegPriSaleRet_Click(Index As Integer)
gSelectedMenu = ""
Select Case Index
    Case 0
        gSelectedMenu = "Trade Sales Return Register" '----Trade Sales Return Register
    Case 1
        gSelectedMenu = "Consignment/Depot Salee Return Register" '----Consignment/Depot Sales Return Register
End Select
'frmRptRegi.Show
End Sub

Private Sub mnuRptRSubSMS_Click(Index As Integer)
On Error GoTo ErrorHandler
gSelectedMenu = ""
Select Case Index
       Case 0 '---Pending C form Receivable
            gSelectedMenu = "Pending C form Receivable"
       Case 1 '---"Form 'C' Receivable SMS (Trade/SIT) Single Party"
            gSelectedMenu = "Form 'C' Receivable SMS (Trade/SIT) Single Party"
       Case 2 '---"Form 'C' Receivable SMS (Trade/SIT) Comman SMS to All Party"
            gSelectedMenu = "Form 'C' Receivable SMS (Trade/SIT) Comman SMS to All Party"
       Case 3 '-

       Case 4 '---"Form 'E1' Receivable SMS  All Party"
            gSelectedMenu = "Form 'E1' Receivable SMS  All Party"
       Case 5 '---"Form 'E1' Receivable SMS  Single Party"
            gSelectedMenu = "Form 'E1' Receivable SMS  Single Party"
       Case 6 '-
       Case 7 '-Form 'E2' Receivable SMS  All Party
            gSelectedMenu = "Form 'E2' Receivable SMS  All Party"
       Case 8 '-Form 'E2' Receivable SMS  Single Party
            gSelectedMenu = "Form 'E2' Receivable SMS  Single Party"
       Case 9 '-
       Case 10 '---General SMS
            gSelectedMenu = "General SMS"
       Case 11 '-
       Case 12 '---Pending C form Receivable All Broker
            gSelectedMenu = "Pending C form Receivable All Broker"
       Case 13 '---Pending C form Receivable Single Broker
            gSelectedMenu = "Pending C form Receivable Single Broker"
       Case 14 '----Form 'E1' Receivable SMS (Trade/SIT) All Broker
            gSelectedMenu = "Form 'E1' Receivable SMS (Trade/SIT) All Broker"
       Case 15 '----Form 'E1' Receivable SMS (Trade/SIT) Single Broker
            gSelectedMenu = "Form 'E1' Receivable SMS (Trade/SIT) Single Broker"
       Case 16 '---Form 'E1' Receivable SMS Comman SMS to All Party
            gSelectedMenu = "Form 'E1' Receivable SMS Comman SMS to All Party"
       Case 17 '-
       Case 18 '---Outstanding To Party
            gSelectedMenu = "Outstanding To Party"
       Case 19 '---Outstanding To Broker
            gSelectedMenu = "Outstanding To Broker"
End Select
If gSelectedMenu <> "" Then frmRpTSMS.Show
Exit Sub
ErrorHandler:
GProcErrorHandler

End Sub

Private Sub mnurptRSubStock_Click(Index As Integer)
On Error GoTo ErrorHandler
gSelectedMenu = ""
Select Case Index
    Case 0 '-- Item Ledger
        gSelectedMenu = "Item Ledger"
        frmRptRegi.Show
    Case 1 '-- Daily Stock To Mill
        gSelectedMenu = "Daily Stock To Mill"
        frmRptRegi.Show
    Case 2 ' -
    Case 3  '--Stock Statement
        gSelectedMenu = "Stock Statement"
        frmRptSelection.Show
    Case 4  '--Stock Statement DateWise
        gSelectedMenu = "Stock Statement Datewise"
        'frmRptSelection.CommonReportSelection True, True, True, False, False, False
        frmRptSelection.Show
    Case 5 ' -
    Case 6  '-- Inward wise
        GProcCrystalRptPreparation CryReport, "Window"
        CryReport.WindowTitle = mnurptRSubStock(Index).Caption
        CryReport.ReportFileName = gReportPath & "RptStkRptInwardWise.rpt"
        CryReport.SelectionFormula = "({TBLVOUSUB.VTYPE}='PY' or {TBLVOUSUB.VTYPE}='PI' or {TBLVOUSUB.VTYPE}='PT') and ({TBLVOUSUB.VSubBag}-{TBLVOUSUB.VSubIssBag})<>0"
        CryReport.Action = 1
        Exit Sub
    Case 7 ''Purchase/Inward wise Outward
        GProcCrystalRptPreparation CryReport, "Window"
        CryReport.WindowTitle = mnurptRSubStock(Index).Caption
        CryReport.ReportFileName = gReportPath & "RptStkRptInVsOut.rpt"
        CryReport.SelectionFormula = "{TBLVOUSUB_Gp.VTYPE}='GP'"
        CryReport.Action = 1
        Exit Sub
    Case 8 '-- "Pending Purchase SIT Register"
        gSelectedMenu = "Pending Purchase SIT Register"
      '  frmRptSelection.CommonReportSelection True, True, True, False, False, False
        frmRptSelection.Show
    Case 9 '- -
    Case 10 '- Valuation
End Select
Exit Sub
ErrorHandler:
GProcErrorHandler
End Sub

Private Sub mnurptRSubStockVal_Click(Index As Integer)
On Error GoTo ErrorHandler
gSelectedMenu = ""
Select Case Index
    Case 0 '-- Stock Valuation (Trade,Sit)
        gSelectedMenu = "Trade & SIT Stock Valuation Register"
        frmRptRegi.Show
        Exit Sub
    Case 1 '-- "Trade & SIT Stock Valuation Summary"
        gSelectedMenu = "Trade & SIT Stock Valuation Summary"
        frmRptRegi.Show
        Exit Sub
End Select
Exit Sub
ErrorHandler:
GProcErrorHandler
End Sub

Private Sub mnurptRSubTax_Click(Index As Integer)
gSelectedMenu = ""
Select Case Index
    Case 0: '-- "V A T Computation"
        gSelectedMenu = "V A T Computation"
    Case 1 '--- "V A T Computation - Detail"
        gSelectedMenu = "V A T Computation - Detail"
    Case 2 '-- Exemption Sales Detail
        gSelectedMenu = "Exemption Sales Detail"
    Case 3 ' -
    Case 4 '-- "Sales Register"
        gSelectedMenu = "Sales Register"
    Case 5 '-- "Purchase Register"
        gSelectedMenu = "Purchase Register"
    Case 6 '--"Sales Tax Return Register"
       ' gSelectedMenu = "Sales Tax Return Register"
    Case 7 '-- "Purchase Tax Return Register"
       ' gSelectedMenu = "Purchase Tax Return Register"
    Case 8 '--
    Case 9 '-- Form Entry
    Case 10 '-- "TDS Report"
        gSelectedMenu = "TDS Register"
    'Case 11 '-- "TDS Summary"
    '    gSelectedMenu = "TDS Summary"
    Case 12 '-- "Service Tax Report"
        'gSelectedMenu = "Service Tax Report"
    Case 14 '-- "V A T Computation 2 And 5 %"
        gSelectedMenu = "V A T Computation 2 And 5 %"
    Case 15  '--Transactionwise Sales Details
        gSelectedMenu = "Transactionwise Sales Details"
    Case 16  '--Transactionwise Purchase Details
        gSelectedMenu = "Transactionwise Purchase Details"
End Select
If gSelectedMenu <> "" Then frmRptTaxReports.Show
End Sub

Private Sub mnurptRSubTaxAuditForm_Click(Index As Integer)
gSelectedMenu = ""
Select Case Index
    Case 0 '--"Partywise Sales VAT Audit"
        gSelectedMenu = "Partywise Sales VAT Audit"
    Case 1 '--"Partywise Purchase VAT Audit"
        gSelectedMenu = "Partywise Purchase VAT Audit"
    Case 2 '--"Monthly Sales Detail VAT Audit"
        gSelectedMenu = "Monthly Sales Detail VAT Audit"
    Case 3 '--"Monthly Purchase Detail VAT Audit"
        gSelectedMenu = "Monthly Purchase Detail VAT Audit"
    Case 4 '--"Received C And E1 Form"
        gSelectedMenu = "Annexure - G (Received C And E1 Form)"
    Case 5 '--"Annexure - I  (Pending  C And E1 Form)"
        gSelectedMenu = "Annexure - I  (Pending  C And E1 Form)"
    Case 6 '--"Annexure - J Section -1   (Partywise VAT Sales)"
        gSelectedMenu = "Annexure - J Section -1   (Partywise VAT Sales)"
    Case 7 '--"Annexure - J Section -1   (R.D. Partywise VAT Sales)"
        gSelectedMenu = "Annexure - J Section -1   (R.D. Partywise VAT Sales)"
    Case 8 '---"Annexure - J Section -1   (U.R.D. Partywise VAT Sales)"
        gSelectedMenu = "Annexure - J Section -1   (U.R.D. Partywise VAT Sales)"
    Case 9 '---"Annexure - J Section -2   (Partywise VAT Purchases)"
        gSelectedMenu = "Annexure - J Section -2   (Partywise VAT Purchases)"
    Case 10 '---"Annexure - J Section -6   (Supllierwise CST Purchases)"
        gSelectedMenu = "Annexure - J Section -6   (Supllierwise CST Purchases)"
    Case 11 '---    Annexure - J Section -6   (Supllierwise CST Sales)
        gSelectedMenu = "Annexure - J Section -6   (Supllierwise CST Sales)"
        
        
End Select
If gSelectedMenu <> "" Then frmRptTaxReports.Show
End Sub

Private Sub mnurptRSubTaxForm_Click(Index As Integer)
gSelectedMenu = ""
Select Case Index
    Case 0 '--"C Form Receivable"
        gSelectedMenu = "C Form Receivable"
    Case 1 '-- "C Form Issuable"
        gSelectedMenu = "C Form Issuable"
    Case 2 '-- "E1 From Receivable"
        gSelectedMenu = "E1 Form Receivable"
    Case 3 '-- "F Form Issuable"
        gSelectedMenu = "F Form Issuable"
    Case 4 '-
    Case 5 '-- "Requirement Of C From"
        gSelectedMenu = "Requirement Of C Form"
    Case 6 '-- "Requirement Of F From"
        gSelectedMenu = "Requirement Of F Form"
    Case 7 '-- F Form Statement
        gSelectedMenu = "F Form Statement"
    Case 8 '-
    Case 9 '-- Pending C Form Receivable Slab Wise
        gSelectedMenu = "Pending C Form Receivable Slab Wise"
    Case 10 '-- Pending E1 Form Receivable Slab Wise
        gSelectedMenu = "Pending E1 Form Receivable Slab Wise"
    Case 11 '-- "E2 From Receivable"
        gSelectedMenu = "E2 Form Receivable"
    Case 12 '-- "E2 From Issuable"
        gSelectedMenu = "E2 From Issuable"
    Case 13 '--"Requirment of E2 Form Invoicewise"
        gSelectedMenu = "Requirment of E2 Form Invoicewise"
    Case 14 '-Requirment of C Form Invoicewise
        gSelectedMenu = "Requirment of C Form Invoicewise"
    Case 15 '- Requirment of F Form Invoicewise
        gSelectedMenu = "Requirment of F Form Invoicewise"
    Case 16 '---C Form Issued But E1Form Not Received
        gSelectedMenu = "C Form Issued But E1Form Not Received"
End Select
If gSelectedMenu <> "" Then frmRptTaxReports.Show
End Sub
Private Sub mnuRptRSubWapp_Click(Index As Integer)
Set gfrmMainMenu = mnutrnSubBankReconEtr(Index)
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
NextCYr = Format(gCYEDate, "yy")
NextCYr = NextCYr + Format(CLng(NextCYr) + 1, "00") '- Yr Code
NextDbNm = gCCode & NextCYr
Set gfrmMainMenu = mnuTransfer(Index)
Select Case Index '-- Transfer To next Year
    Case 0 '- Transfer Opening Balance to Next Year
         GProcUtlCarryForwardOpening NextDbNm
    Case 1 '- Transfer Outstanding to Next Year
         GProcUtlCarryForwardOutstanding NextDbNm
    Case 2 '- Transfer Stock to Next Year
         GProcUtlCarryForwardStock NextDbNm
    Case 3 '- Transfer Late Payment Int to Next Year
         GProcUtlCarryForwardTINNo NextDbNm
    Case 4 '--Transfer Trade Purchase Booking
         GProcUtlCarryForwardTradeBooking NextDbNm
    Case 5 '- Transfer Party Booking to Next Year
         GProcUtlCarryForwardPartyBooking NextDbNm
         
End Select
End Sub

Private Sub mnutrnSubBankReconEtr_Click(Index As Integer)
Set gfrmMainMenu = mnutrnSubBankReconEtr(Index)
Select Case Index '-- Bank Reconciliation Entry
    Case 0 '- Payment
        gSelectedMenu = "PA"
    Case 1 '- Receipt
        gSelectedMenu = "RE"
End Select
frmBankReConcilEntry.Show
End Sub

Private Sub mnutrnSubBKfrm_Click(Index As Integer)
Set gfrmMainMenu = mnutrnSubBKfrm(Index)
Select Case Index
    Case 0 '-- Mill Bill
        gfrmTypeStr = "OM"
        frmBookingMillBill.Show
    Case 1 '-- Sit
        gfrmTypeStr = "OS"
        frmBookingSIT.Show
    Case 2  '-- Consignment /Depot Booking
        gfrmTypeStr = "OI"
        frmBookingInward.Show
    Case 3 '--Party
        gfrmTypeStr = "OP"
        frmBookingParty.Show
End Select
End Sub
Private Sub mnutrnSubFormfrm_Click(Index As Integer)
End Sub

'==== Confirmation
Private Sub mnutrnSubConfirmfrm_Click(Index As Integer)
Set gfrmMainMenu = mnutrnSubConfirmFrm(Index)
Select Case Index
    Case 0 '-- Sales Confirmation
        frmConfirmationSales.Show
End Select
End Sub

Private Sub mnutrnSubFormEtr_Click(Index As Integer)
Set gfrmMainMenu = mnutrnSubFormEtr(Index)
gSelectedMenu = ""
Select Case Index
    Case 0 '-- Form c Received
        gSelectedMenu = "CR"
    Case 1 '-- Form c Issue
        gSelectedMenu = "CI"
    Case 2 '-- Form E1 Received
        gSelectedMenu = "E1R"
    Case 3 '-- Form F Issue
        gSelectedMenu = "FI"
    Case 4 '-- Form E2 Received
        gSelectedMenu = "E2R"
    Case 5 '-- Form E2 Issue
        gSelectedMenu = "E2I"
        
End Select
If gSelectedMenu <> "" Then frmFormEntry.Show
End Sub

Private Sub mnutrnSubGPFrm_Click(Index As Integer)
Set gfrmMainMenu = mnutrnSubGPFrm(Index)
Select Case Index
    Case 0 '- Gate Pass
       frmGatepass.Show
End Select
End Sub

'---- Purchase
Private Sub mnuTrnSubPurfrm_Click(Index As Integer)
Set gfrmMainMenu = mnutrnSubPurfrm(Index)
Select Case Index
    Case 0 '-- Yarn Purchase
        gfrmTypeStr = "PY"
        gfrmHunkStr = ""
        frmPurchaseTrade.Show
    Case 1 '--SIT Purchase
        gfrmTypeStr = "PT"
        frmPurchaseSIT.Show
    Case 2 '-- Other Purchase
        gfrmTypeStr = "PO"
        frmPurchaseOther.Show
    Case 3 '-- Goods Inward
        gfrmTypeStr = "PI"
        frmPurchaseInward.Show
    Case 4 '-- Yarn Purchase
        gfrmTypeStr = "PY"
        gfrmHunkStr = "HNK"
        frmPurchaseTrade.Show
End Select
End Sub
'---- Purchase Return
Private Sub mnuTrnSubPurRetfrm_Click(Index As Integer)
Set gfrmMainMenu = mnutrnSubPurRetfrm(Index)
Select Case Index
    Case 0 '-- Trade
        gfrmTypeStr = "VY"
        frmPurchaseReturn.Show
    Case 1 '--Inward
        gfrmTypeStr = "VI"
        frmPurchaseReturn.Show
End Select
End Sub
'--- Sales
Private Sub mnuTrnSubSalefrm_Click(Index As Integer)
Set gfrmMainMenu = mnutrnSubSalefrm(Index)
Select Case Index
    Case 0 '--Sale Trade
        gfrmTypeStr = "SY"
        gfrmHunkStr = "TRD"
        frmSales.Show
    Case 1 '-- Sale Consignment
        gfrmTypeStr = "SO"
        gfrmHunkStr = ""
'        '-- Sales A/c
'        gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,AcCode,AcCity,AcMillTypeStr", Array("Sales Account", "AcName", "acCode", "City", "Mill Type"), Array(3000, 0, 0, 2000, 2000), " ((AgCode=90029 or GpCode=90029) and (AcMillType=1))", , "AcName"
'        If gClsSearch.SearchMultiRetCol(0) = "" Then Exit Sub
        frmSales.Show
    Case 2 '-- Sale Depot
        gfrmTypeStr = "SD"
        gfrmHunkStr = ""
        '-- Sales A/c
        gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,Accode,AcCity,AcMillTypeStr", Array("Sales Account", "AcName", "acCode", "City", "Mill Type"), Array(3000, 0, 0, 2000, 2000), " ((AgCode=90029 or GpCode=90029) and AcMillType=2)", , "AcName"
        If gClsSearch.SearchMultiRetCol(0) = "" Then Exit Sub
        frmSales.Show
    Case 3 '-- Sale SIT
        gfrmTypeStr = "ST"
        gfrmHunkStr = ""
        frmSales.Show
    Case 4 '-- Sale Mill
        gfrmTypeStr = "SM"
        gfrmHunkStr = ""
        '-- Sales A/c
        gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,Accode,AcCity,AcMillTypeStr", Array("Sales Account", "AcName", "acCode", "City", "Mill Type"), Array(3000, 0, 0, 2000, 2000), " ((AgCode=90029 or GpCode=90029) and AcMillType=3)", , "AcName"
        If gClsSearch.SearchMultiRetCol(0) = "" Then Exit Sub
        frmSalesMillBill.Show
    Case 5 '--Sale Trade Hank
        gfrmTypeStr = "SY"
        gfrmHunkStr = "HNK"
        frmSales.Show
End Select
End Sub
'-- Sale Return
Private Sub mnuTrnSubSlRetfrm_Click(Index As Integer)
Set gfrmMainMenu = mnutrnSubSlRetfrm(Index)
Select Case Index
    Case 0 '-- Trade
        gfrmTypeStr = "RY"
        frmSalesReturn.Show
    Case 1 '--Cosign/Depot
        gfrmTypeStr = "RO"
        frmSalesReturn.Show
End Select
End Sub
'---- Payment
Private Sub mnuTrnSubPayfrm_Click(Index As Integer)
Set gfrmMainMenu = mnutrnSubPayfrm(Index)
Select Case Index
    Case 0 '--Cash
        gfrmTypeStr = "CP"
        frmPayment.Show
    Case 1 '-- Bank
        gfrmTypeStr = "BP"
        frmPayment.Show
End Select
End Sub
'----- JV
Private Sub mnutrnSubJvfrm_Click(Index As Integer)
Set gfrmMainMenu = mnutrnSubJvfrm(Index)
Select Case Index
    Case 0
        gfrmTypeStr = "JV"
    Case 1
        gfrmTypeStr = "PN"
    Case 2
        gfrmTypeStr = "SN"
End Select
frmJV.Show
End Sub
'----- Receipt
Private Sub mnuTrnSubRecfrm_Click(Index As Integer)
Set gfrmMainMenu = mnutrnSubRecfrm(Index)
Select Case Index
    Case 0
        gfrmTypeStr = "CR"
        frmReceipt.Show
    Case 1
        gfrmTypeStr = "BR"
        frmReceipt.Show
    Case 2 '--Mill Receipt
        gfrmTypeStr = "MR"
        frmReceiptMill.Show
End Select
End Sub
Private Sub mnuCMV_Click(Index As Integer)
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
    Case 4 '- "Mill Bill Account Ledger"
        gSelectedMenu = "Mill Bill Account Ledger"
    Case 5 '-  Bank Reconciliation
        gSelectedMenu = "Bank Reconciliation"
    Case 6 '-  Multi Account Ledger
         frmMultiAccount.Show
         Exit Sub
End Select
If gSelectedMenu <> "" Then frmRptAccount.Show
End Sub

Private Sub mnuRptRSubFin_Click(Index As Integer)
gSelectedMenu = ""
Select Case Index
    Case 0: '-"Balance Sheet"
        gSelectedMenu = "Balance Sheet"
        frmRptFinal.Show
    Case 2: '-"Profit/Loss Account"
        gSelectedMenu = "Profit/Loss Account"
        frmRptFinal.Show
    Case 3  '---"Schedule Printing"
        gSelectedMenu = "Schedule Printing"
        frmRptFinal.Show
    Case 4 '-- All Schedule Printing
        gSelectedMenu = "All Schedule Printing"
        frmRptFinal.Show
    Case 5
        gSelectedMenu = "Monthly Profit & Loss A/C"
        frmRptFinal.Show
    Case 6 '---Loan Reporting
        gSelectedMenu = "Loan Reporting"
        frmRptFinal.Show
    Case 7 '-- Loan-Interest-TDS Details
        gSelectedMenu = "Loan-Interest-TDS Details"
        frmRptFinal.Show
    Case 8 '---Countwise Profit and Loss
        gSelectedMenu = "Countwise Profit and Loss"
        frmRptSelection.Show
        
End Select
'If gSelectedMenu <> "" Then frmRptFinal.Show
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
gSelectedMenu = ""
Select Case Index
    Case 0:
        gSelectedMenu = "Trade Sale Register" '----Trade Sale Register
    Case 1:
        gSelectedMenu = "Sale In Transit Sale Register" '----SIT Sale Register
    Case 2:
        gSelectedMenu = "Consignment/Depot Sale Register" '----Consignment/Depot Sale Register
    Case 3:
        gSelectedMenu = "Mill Bill Register" '----Mill bill Register
End Select
frmRptRegi.Show
End Sub
Private Sub mnutrnSubTDS_Click(Index As Integer)
Set gfrmMainMenu = mnutrnSubTDS(Index)
Select Case Index
    Case 0 '- TDS Entry
        frmTDSEntry.Show
End Select
End Sub
Private Sub mnuUtlSub_Click(Index As Integer)
Dim NextCYr As String
Dim NextDbNm As String

Set gfrmMainMenu = mnuUtlSub(Index)
Select Case Index
    Case 0 '-- Company Select
        frmSelCompany.Show 1
        frmMain.Show
    Case 1 '-- Working Year
        frmUtlSelectYear.Show
    Case 2 '-- Company Creation
        frmMastCompany.Show
    Case 3 '- User Security (change Password)
        frmUtlChangePassword.Show
    Case 4 '-  -
    Case 5 '-- Last Year Outstanding
        frmUtlLastYrOutstanding.Show
    Case 6 '-  -
        frmUtlLastYrOutstandingPurch.Show
    Case 7 '-- Update Balance
        GprocUtlUpdateBalance gCYEDate
    Case 8 '-- Purchase Stock To Next Year
        ''PurchaseStockTransferToNextYear
    Case 9 '-- Transfer Account
        frmUtlAccountTransfer.Show
    Case 10 '-- -
    Case 11 '- Carry Forword
'        NextCYr = Format(gCYEDate, "yy")
'        NextCYr = NextCYr + Format(CLng(NextCYr) + 1, "00") '- Yr Code
'        NextDbNm = gCCode & NextCYr
'        GProcUtlCarryForward NextDbNm
    Case 12 '- Year End
       If MsgBox("Are You Sure For Create New Year ?", vbQuestion + vbYesNo + vbDefaultButton2, "Year End") = vbYes Then
            GProcUtlYearEnd
        End If
    Case 13 '-
    Case 14 '-- Backup Data (For Access )
        If gBackEndDB = gBackEndAccess Then frmUtlBackUpData.Show 1
    Case 15 '-- Table/View Modification
    Case 16
        gSelectedMenu = "For Eof() Bof() Error"
        frmRptTaxReports.Show
        Exit Sub
    
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



