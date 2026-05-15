VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmRptAccount 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Report"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13320
   FillColor       =   &H00C0C0C0&
   Icon            =   "frmRptAccount.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   13320
   Begin VB.TextBox txtCode 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   2
      Left            =   1035
      TabIndex        =   7
      Text            =   "txtCode(2)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1695
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.TextBox txtCode 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   1
      Left            =   1035
      TabIndex        =   5
      Text            =   "txtCode(1)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1215
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.TextBox txtCode 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   0
      Left            =   1035
      TabIndex        =   3
      Text            =   "txtCode(0)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   735
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.CheckBox ChkConsolidate 
      BackColor       =   &H00EDFDFE&
      Caption         =   "Is All Firm Report ?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   1620
      TabIndex        =   47
      Top             =   2610
      Visible         =   0   'False
      Width           =   2430
   End
   Begin VB.TextBox txttmp 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   11745
      TabIndex        =   45
      Text            =   "0"
      Top             =   1260
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.TextBox txttmp 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   10170
      TabIndex        =   22
      Text            =   "4"
      Top             =   3285
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.TextBox txttmp 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   12510
      TabIndex        =   27
      Text            =   "3"
      Top             =   4170
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.TextBox txttmp 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   12510
      TabIndex        =   26
      Text            =   "2"
      Top             =   3720
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.CheckBox ChkSummary 
      BackColor       =   &H00EDFDFE&
      Caption         =   "Is Summary ?"
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   8910
      TabIndex        =   23
      Top             =   3720
      Visible         =   0   'False
      Width           =   2190
   End
   Begin VB.TextBox txttmp 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   12510
      TabIndex        =   25
      Text            =   "1"
      Top             =   3285
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.TextBox txttmp 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   10170
      TabIndex        =   21
      Text            =   "0"
      Top             =   2850
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.CheckBox ChkIntRateMonthly 
      BackColor       =   &H00EDFDFE&
      Caption         =   "Is Interest Rate Monthly ?"
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   8910
      TabIndex        =   24
      Top             =   4125
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   2
      Left            =   2505
      TabIndex        =   8
      Top             =   1695
      Visible         =   0   'False
      Width           =   5865
   End
   Begin VB.CheckBox ChkIsCashBankGroup 
      BackColor       =   &H00EDFDFE&
      Caption         =   "With Cash And  Bank"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   10875
      TabIndex        =   20
      Top             =   2070
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CheckBox ChkIsRefNo 
      BackColor       =   &H00EDFDFE&
      Caption         =   "&Ref. No Required"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   8910
      TabIndex        =   19
      Top             =   2055
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CommandButton cmdSPE 
      Height          =   495
      Index           =   0
      Left            =   1065
      Picture         =   "frmRptAccount.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4920
      Width           =   1335
   End
   Begin VB.CommandButton cmdSPE 
      Height          =   495
      Index           =   1
      Left            =   2640
      Picture         =   "frmRptAccount.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Print"
      Top             =   4920
      Width           =   1335
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   0
      Left            =   2520
      TabIndex        =   4
      Top             =   735
      Visible         =   0   'False
      Width           =   5865
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000018&
      Caption         =   "Report Format"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   855
      Left            =   285
      TabIndex        =   42
      Top             =   3135
      Visible         =   0   'False
      Width           =   5580
      Begin VB.OptionButton optReportOn 
         BackColor       =   &H80000018&
         Caption         =   "&Grid"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   315
         Index           =   2
         Left            =   3885
         TabIndex        =   10
         Top             =   360
         Width           =   1425
      End
      Begin VB.OptionButton optReportOn 
         BackColor       =   &H80000018&
         Caption         =   "&Windows"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   315
         Index           =   1
         Left            =   2085
         TabIndex        =   9
         Top             =   360
         Value           =   -1  'True
         Width           =   1425
      End
      Begin VB.OptionButton optReportOn 
         BackColor       =   &H80000018&
         Caption         =   "MS &DOS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   315
         Index           =   0
         Left            =   210
         TabIndex        =   32
         Top             =   360
         Width           =   1425
      End
   End
   Begin VB.OptionButton optSelType 
      BackColor       =   &H00EDFDFE&
      Caption         =   "&Status"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   315
      Index           =   3
      Left            =   10320
      TabIndex        =   17
      Top             =   915
      Width           =   2685
   End
   Begin VB.OptionButton optSelType 
      BackColor       =   &H00EDFDFE&
      Caption         =   "&Monthly"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   315
      Index           =   2
      Left            =   8880
      TabIndex        =   16
      Top             =   900
      Width           =   1425
   End
   Begin VB.OptionButton optSelType 
      BackColor       =   &H00EDFDFE&
      Caption         =   "&Partywise"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   315
      Index           =   1
      Left            =   10320
      TabIndex        =   15
      Top             =   540
      Width           =   2685
   End
   Begin VB.OptionButton optSelType 
      BackColor       =   &H00EDFDFE&
      Caption         =   "&Simple"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   315
      Index           =   0
      Left            =   8880
      TabIndex        =   14
      Top             =   525
      Value           =   -1  'True
      Width           =   1425
   End
   Begin VB.CheckBox chkDaySum 
      BackColor       =   &H00EDFDFE&
      Caption         =   "Daywise &Summary"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   8880
      TabIndex        =   18
      Top             =   1305
      Visible         =   0   'False
      Width           =   4215
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   5520
      Top             =   4050
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdSPE 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   4215
      TabIndex        =   13
      Top             =   4920
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   2550
      TabIndex        =   1
      Top             =   255
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   122552323
      CurrentDate     =   38050
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   6795
      TabIndex        =   2
      Top             =   255
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   122552323
      CurrentDate     =   38050
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   1
      Left            =   2505
      TabIndex        =   6
      Top             =   1200
      Visible         =   0   'False
      Width           =   5865
   End
   Begin VB.ListBox lstParty 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2505
      Sorted          =   -1  'True
      TabIndex        =   41
      Top             =   1200
      Width           =   4335
   End
   Begin MSComctlLib.ProgressBar prgBar 
      Height          =   255
      Left            =   180
      TabIndex        =   43
      Top             =   4140
      Visible         =   0   'False
      Width           =   5595
      _ExtentX        =   9869
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&TDS % :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Index           =   7
      Left            =   10965
      TabIndex        =   46
      Top             =   1305
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Side && Less Days :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Index           =   4
      Left            =   11115
      TabIndex        =   44
      Top             =   2940
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Less Payment"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Index           =   6
      Left            =   11115
      TabIndex        =   39
      Top             =   3720
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Days of Year       :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   450
      Index           =   5
      Left            =   8910
      TabIndex        =   37
      Top             =   3240
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Less Receipt"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Index           =   3
      Left            =   11115
      TabIndex        =   40
      Top             =   4170
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Side (Sales)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Index           =   1
      Left            =   11115
      TabIndex        =   38
      Top             =   3285
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Interest %   :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Index           =   2
      Left            =   8910
      TabIndex        =   36
      Top             =   2850
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "&Sarafi Interest :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   8835
      TabIndex        =   35
      Top             =   2535
      Visible         =   0   'False
      Width           =   3540
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000040C0&
      X1              =   13155
      X2              =   8715
      Y1              =   2430
      Y2              =   2430
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000040C0&
      X1              =   13140
      X2              =   8670
      Y1              =   1650
      Y2              =   1650
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000040C0&
      X1              =   8550
      X2              =   75
      Y1              =   2295
      Y2              =   2295
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Mill       :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Left            =   210
      TabIndex        =   31
      Top             =   1695
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "A/c G&roupwise Ledger   :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   8835
      TabIndex        =   34
      Top             =   1755
      Visible         =   0   'False
      Width           =   2640
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A/C &Gp :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Left            =   210
      TabIndex        =   29
      Top             =   735
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Accoutwise &Ledger   :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   8835
      TabIndex        =   33
      Top             =   165
      Width           =   2115
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To Date       :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Left            =   5400
      TabIndex        =   28
      Top             =   300
      Width           =   1215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Account:  "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Left            =   210
      TabIndex        =   30
      Top             =   1200
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "From    :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Index           =   0
      Left            =   210
      TabIndex        =   0
      Top             =   255
      Width           =   795
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   4530
      Left            =   75
      Top             =   90
      Width           =   8475
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00BDDFD1&
      FillStyle       =   0  'Solid
      Height          =   765
      Left            =   615
      Shape           =   4  'Rounded Rectangle
      Top             =   4800
      Width           =   5220
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   4530
      Left            =   8685
      Top             =   90
      Width           =   4470
   End
End
Attribute VB_Name = "frmRptAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer, LnCnt As Integer, PgCnt As Integer, wDotLine As String
Dim wAcCode As Long, wAgCode As Long, mScreen As Boolean, mP12 As String, mPc As String
Dim wPartyCd(20) As String
'==== Format Text Boxes Value
Private Sub CheckForNumber()
Dim i As Integer
'-- Format txttmp
For i = 0 To 4
    If IsNumeric(txttmp(i)) = False Then
        txttmp(i) = 0
    End If
    '--format
    If i = 0 Then '-Rate Double
        txttmp(i) = Format(txttmp(i), FStr)
    Else
        txttmp(i) = Format(txttmp(i), PStr)
    End If
Next i
End Sub
'-- Validate Data
Private Function ValidateData() As Boolean
ValidateData = True
Select Case gSelectedMenu
    Case "Sarafi Interest Report", "Sarafi Interest Report ANK", "Mill Bill Payment Interest" '-"Sarafi Interest Report"
        '--Interest %
        If CDbl(txttmp(0)) = 0 Or CDbl(txttmp(0)) > 100 Then
            ValidateData = False
            MsgBox "Check Interest %", vbOKOnly + vbInformation
            txttmp(0).SetFocus
            Exit Function
        End If
        '-- Total Days for int Rate Yearly
        '-Days of year
        If ChkIntRateMonthly = 0 And CLng(txttmp(4)) = 0 Then txttmp(4) = DateDiff("d", gCYSDate, gCYEDate) + 1
            
'            ValidateData = False
'            MsgBox "Check Total Days of Year", vbOKOnly + vbInformation
'            txttmp(4).SetFocus
'            Exit Function
'        End If
End Select
End Function
Private Sub cmdSPE_Click(Index As Integer)
Dim sTp As String
Dim FromDt As Date
Dim ToDt As Date
Dim X As String
Dim Y As String
Dim x1 As String
Dim Cdt As String
Dim ShowRepo As Boolean
Dim tmp1 As Recordset
Dim tmp2 As Recordset
Dim Opbal1 As Double
Dim Opbal2 As Double
ShowRepo = False
Dim AcCd As Long
Dim ShowInRpt As Integer
Select Case Index
Case 0
    mScreen = True
Case 1
    mScreen = False
Case 2
   Unload Me
   Exit Sub
End Select
'If ValidateData = False Then Exit Sub
CheckForNumber
FromDt = Format(DTPicker1.Value, "dd/MM/yyyy") '-From Dt
ToDt = Format(DTPicker2.Value, "dd/MM/yyyy") '-To Dt
'--- Ledger Type
If optSelType(0) = True Then '-Simple
     sTp = "D"
ElseIf optSelType(1) = True Then '-Partywise
     sTp = "P"
ElseIf optSelType(2) = True Then '-Monthly
     sTp = "M"
ElseIf optSelType(3) = True Then '-T Format
      sTp = "T"
End If
'== Dos
If optReportOn(0) = True Then
'   Pty_Ldr1 (sTp)
'== Window
ElseIf optReportOn(1) = True Or optReportOn(2) = True Then
  If optReportOn(1) = True Then GProcCrystalRptPreparation CrystalReport1, IIf(mScreen, "Window", "Printer")
    With CrystalReport1
        Select Case gSelectedMenu
'            '============ List End
            Case "Account Ledger", "Sarafi Interest Report", "Account Confirmation", "Groupwise Account Ledger", "Opposite A/c wise Summery", "Loan A/c Details", "Asset A/c Details", "Depreciation Chart", "TDS Receivable O/s", "TCS Receivable O/s", "TDS On Sales Receivable O/s" '====="Account Ledger"
                 ShowRepo = True
                 gCmd.CommandText = "PrcPrepareOpBalance"
                 gCmd.Parameters.Refresh
                 If gSelectedMenu = "Account Ledger" Then
                    gCmd.Parameters("@Ldr") = "L"
                 Else
                    gCmd.Parameters("@Ldr") = ""
                 End If
                 If ChkConsolidate.Value = 1 Then
                    gCmd.Parameters("@VFirm") = ""
                 Else
                    gCmd.Parameters("@VFirm") = Trim(gCCode)
                 End If
                 gCmd.Parameters("@To_dt") = Format(FromDt, "yyyy/mm/dd")
                 gCmd.Parameters("@FyStart_dt") = Format(gCYSDate, "yyyy/mm/dd")
                 gCmd.Parameters("@UCode") = gUserName
'                 gCmd.Parameters("@VFirm") = Trim(gCCode)
                 gCmd.Parameters("@CGSTAc") = gCgstAcCode
                 gCmd.Parameters("@SGSTAc") = gSgstAcCode
                 gCmd.Parameters("@IGSTAc") = gIgstAcCode
                 gCmd.Parameters("@CGSTRcmPayAc") = gCgstRCMAcCode
                 gCmd.Parameters("@SGSTRcmPayAc") = gSgstRCMAcCode
                 gCmd.Parameters("@IGSTRcmPayAc") = gIgstRCMAcCode
                 gCmd.Parameters("@TcsRec") = gTcsRec
                 gCmd.Parameters("@TcsPay") = gTcsPay
                 gCmd.Parameters("@CGSTRcmRecAc") = gCgstRCMRecCode
                 gCmd.Parameters("@SGSTRcmRecAc") = gSgstRCMRecCode
                 gCmd.Parameters("@IGSTRcmRecAc") = gIgstRCMRecCode
                 gCmd.Parameters("@RoundingAc") = gRoundOffAc
                 
                 If txt(1) = "" Then
                    wAcCode = 0
                 End If
                 If txt(0) = "" Then
                    wAgCode = 0
                 End If
                 
                 If wAcCode <> 0 Then gCmd.Parameters("@Accode") = wAcCode
                 If wAgCode <> 0 Then gCmd.Parameters("@Agcode") = wAgCode
                 If gSelectedMenu = "Depreciation Chart" Then
                    gCmd.Parameters("@Agcode") = 90008
                 End If
                 gCmd.Execute
                gCmd.CommandText = "PrcPrepareAccLedger"
                gCmd.Parameters.Refresh
                gCmd.Parameters("@From_dt") = Format(FromDt, "yyyy/mm/dd")
                gCmd.Parameters("@To_dt") = Format(ToDt, "yyyy/mm/dd")
                gCmd.Parameters("@UCode") = gUserName
                gCmd.Parameters("@IsDepotFirm") = gCIsDepotFirm
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
                If ChkConsolidate.Value = 1 Then
                   gCmd.Parameters("@VFirm") = ""
                Else
                   gCmd.Parameters("@VFirm") = Trim(gCCode)
                End If
                If wAcCode <> 0 Then gCmd.Parameters("@Accode") = wAcCode
                If wAgCode <> 0 Then gCmd.Parameters("@Agcode") = wAgCode
                If gSelectedMenu = "Depreciation Chart" Then
                   gCmd.Parameters("@Agcode") = 90008
                End If
                gCmd.Execute
                If gSelectedMenu = "Groupwise Account Ledger" And (wAgCode = 90018 Or wAgCode = 0) Then
                    '--- For Closing Stock
                    X = "select '" & Format(ToDt, "yyyy/mm/dd") & "','" & gCCode & "','" & gCName & "','Yarn Stock ( Closing )', 'Closing Balance',StAmount,0,'Stock In Hand',2,'" & gUserName & "',tblMastAccount.AcOurCode,tblMastAccount.AgCode,90018    from tblMastAccount,tblClosingStock  where tblMastAccount.AcCode = tblClosingStock.StAcCode " _
                         & " and tblMastAccount.AcName = 'Yarn Stock' and tblClosingStock.stDate = ('" & Format(ToDt, "yyyy/mm/dd") & "')  and tblClosingStock.Vfirm = '" & gCCode & "' and tblClosingStock.StAmount > 0    "

                    If GProcRstOpen(tmp1, X, "O", gCn) > 0 Then
                        
                 gCn.Execute " Insert Into TmpAccLedger (vdt,vfirm,VFirmName,AcName,AcOname,VDrAmt,VCrAmt,AgName,VSrNo,UserName,AcOurCode,AgCode,TopGroup)" _
                             & " select '" & Format(ToDt, "yyyy/mm/dd") & "','" & gCCode & "','" & gCName & "','Yarn Stock ( Closing )', 'Closing Balance',StAmount,0,'Stock In Hand',2,'" & gUserName & "',tblMastAccount.AcOurCode,tblMastAccount.AgCode,90001    from tblMastAccount,tblClosingStock  where tblMastAccount.AcCode = tblClosingStock.StAcCode " _
                             & " and tblMastAccount.AcName = 'Yarn Stock' and tblClosingStock.stDate = ('" & Format(ToDt, "yyyy/mm/dd") & "')  and tblClosingStock.Vfirm = '" & gCCode & "'"
                    End If
                End If
                
                If gSelectedMenu = "Sarafi Interest Report" Then
                   wAgCode = GProcGetColumnValue("tblMastGroup", "AgName", txt(0), "S", "AgCode", "N")
                   gCn.BeginTrans
                   gCn.Execute "delete from TmpAccLedger where username = '" & gUserName & "' and  Vtype ='0P'"
                   X = "delete from TmpAccLedger where vfirm+vtype+convert(char(10),isnull(vno,0))+AcName in " _
                       & " (select t.vfirm+vtype+convert(char(10),isnull(vno,0))+AcName from TmpAccLedger T,(select vfirm,partycd,max(vdt) as dt from tblInterestCalDate where vyear= " & gCYear & " group by vfirm,partycd) as A  where username = '" & gUserName & "' and vdt < dt  and T.vfirm=a.VFirm and a.PartyCd =T.AcCode)"
                   gCn.Execute X
                   gCn.CommitTrans
                   gCn.BeginTrans
                   x1 = "select Vfirm,PartyCd,max(vdt) as dt from tblInterestCalDate where vyear= " & gCYear & " and vfirm = '" & gCCode & "'   group by vfirm,partycd "
                   If GProcRstOpen(tmp1, x1, "R", gCn) > 0 Then
                      tmp1.MoveFirst
                      Do While (Not tmp1.EOF)
                         If wAgCode <> 0 Then
                            Y = "Select tblInterestCalDate.VNo,tblInterestCalDate.Vdt,tblInterestCalDate.VYear,tblInterestCalDate.VFirm,tblInterestCalDate.CtrNo, tblInterestCalDate.PartyCd,tblInterestCalDate.OpAmt,tblInterestCalDate.CrDr," _
                                & " tblMastAccount.AcName,tblMastAccount.AgCode,tblMastGroup.AgName,tblMastAccount.AcPan,tblMastAccount.AcOurCode,tblMastGroup.TopGroup,tblMastAccount.AcIntPer" _
                                & " from tblInterestCalDate,tblMastAccount,tblMastGroup" _
                                & " where vfirm = '" & tmp1.Fields("Vfirm") & "' and PartyCd = " & tmp1.Fields("PartyCd") & " and vdt = '" & Format(tmp1.Fields("dt"), "yyyy/mm/dd") & "' " _
                                & " and tblInterestCalDate.PartyCd = tblMastAccount.AcCode and tblMastAccount.AgCode = tblMastGroup.AgCode"
                            If GProcRstOpen(tmp2, Y, "R", gCn) > 0 Then
                               If tmp2.Fields("AgCode") = wAgCode Then
                                  If tmp2.Fields("OpAmt") > 0 Then
                                    X = "'0P'" & ",'" & gCCode & "','" & gCName & "'," & tmp2.Fields("PartyCd") & ",'" & tmp2.Fields("AcName") & "','Opening Balance'," & IIf(tmp2.Fields("CrDr") = 1, 0, tmp2.Fields("OpAmt")) & "," & IIf(tmp2.Fields("CrDr") = 1, tmp2.Fields("OpAmt"), 0) & ",'" & tmp2.Fields("AgName") & "','" & tmp2.Fields("AcPan") & "'," _
                                        & "'" & tmp2.Fields("AcOurCode") & "'," & tmp2.Fields("TopGroup") & "," & tmp2.Fields("AgCode") & ",'" & Format(tmp2.Fields("Vdt"), "yyyy/mm/dd") & "'," & tmp2.Fields("AcIntPer") & ",'" & gUserName & "'"
                                    gCn.Execute "Insert into TmpAccLedger(vtype,vfirm,VFirmName,AcCode,AcName,AcOname,VCrAmt,VDrAmt,AgName,VAcPan,AcOurCode,TopGroup,AgCode,Vdt,AcIntPer,UserName)" _
                                                 & " Values(" & X & ")"
                                  End If
                               End If
                            End If
                         End If
                         If wAcCode <> 0 Then
                            If wAcCode = tmp1.Fields("PartyCd") Then
                               Y = "Select tblInterestCalDate.VNo,tblInterestCalDate.Vdt,tblInterestCalDate.VYear,tblInterestCalDate.VFirm,tblInterestCalDate.CtrNo, tblInterestCalDate.PartyCd,tblInterestCalDate.OpAmt,tblInterestCalDate.CrDr," _
                                   & " tblMastAccount.AcName,tblMastAccount.AgCode,tblMastGroup.AgName,tblMastAccount.AcPan,tblMastAccount.AcOurCode,tblMastGroup.TopGroup,tblMastAccount.AcIntPer" _
                                   & " from tblInterestCalDate,tblMastAccount,tblMastGroup" _
                                   & " where vfirm = '" & tmp1.Fields("Vfirm") & "' and PartyCd = " & tmp1.Fields("PartyCd") & " and vdt = '" & Format(tmp1.Fields("dt"), "yyyy/mm/dd") & "' " _
                                   & " and tblInterestCalDate.PartyCd = tblMastAccount.AcCode and tblMastAccount.AgCode = tblMastGroup.AgCode"
                               If GProcRstOpen(tmp2, Y, "R", gCn) > 0 Then
                                  If tmp2.Fields("OpAmt") > 0 Then
                                     X = "'0P'" & ",'" & gCCode & "','" & gCName & "'," & tmp2.Fields("PartyCd") & ",'" & tmp2.Fields("AcName") & "','Opening Balance'," & IIf(tmp2.Fields("CrDr") = 1, 0, tmp2.Fields("OpAmt")) & "," & IIf(tmp2.Fields("CrDr") = 1, tmp2.Fields("OpAmt"), 0) & ",'" & tmp2.Fields("AgName") & "','" & tmp2.Fields("AcPan") & "'," _
                                         & "'" & tmp2.Fields("AcOurCode") & "'," & tmp2.Fields("TopGroup") & "," & tmp2.Fields("AgCode") & ",'" & Format(tmp2.Fields("Vdt"), "yyyy/mm/dd") & "'," & tmp2.Fields("AcIntPer") & ",'" & gUserName & "'"
                                     gCn.Execute "Insert into TmpAccLedger(vtype,vfirm,VFirmName,AcCode,AcName,AcOname,VCrAmt,VDrAmt,AgName,VAcPan,AcOurCode,TopGroup,AgCode,Vdt,AcIntPer,UserName)" _
                                                  & " Values(" & X & ")"
                                  End If
                               End If
                            
                            
                            End If
                         End If
                         tmp1.MoveNext
                      Loop
                   End If
                   gCn.CommitTrans
                     X = "select * from tmpaccledger where tmpaccledger.UserName = '" & gUserName & "' "
                     If ChkConsolidate.Value <> 1 Then
                        X = X + " and tmpaccledger.Vfirm = '" & Trim(gCCode) & "'"
                     End If
                     If GProcRstOpen(tmp1, X, "R", gCn) > 0 Then
                         tmp1.MoveFirst
                         Do While Not tmp1.EOF
                            ShowInRpt = GProcGetColumnValue("tblMastAccount", "AcName", "" & tmp1.Fields("AcName") & "", "S", "AcIsShowOuts", "N")
                            AcCd = GProcGetColumnValue("tblMastAccount", "AcName", "" & tmp1.Fields("AcName") & "", "S", "AcCode", "N")
                            gCn.Execute "update tmpaccledger Set UserCode = " & ShowInRpt & " where tmpaccledger.AcName = '" & tmp1.Fields("AcName") & "' and tmpaccledger.VFirm = '" & tmp1!VFirm & "' and tmpaccledger.username='" & gUserName & "'"
                            tmp1.MoveNext
                         Loop
                     End If
                    gCn.Execute "Delete from tmpaccledger where charindex('Int',Vnar1) > 0 and  UserName='" & gUserName & "'"
                    gCn.Execute "Delete from tmpaccledger where charindex('Int',Vnar1) > 0 and  UserName='" & gUserName & "'"
                    If chkDaySum.Value = 0 Then
                       gCn.Execute "Delete from tmpaccledger where UserCode = 1 and  UserName='" & gUserName & "'"
                    End If
                    '--For Firmwise Int. Rate
                    gCn.BeginTrans
                    X = "select PartyCd,IntRate from tblMastFirmwiseIntRate where tblMastFirmwiseIntRate.VFirm = '" & gCCode & "' "
                    If GProcRstOpen(tmp1, X, "R", gCn) > 0 Then
                       tmp1.MoveFirst
                       Do While Not tmp1.EOF
                          gCn.Execute "update tmpaccledger Set AcIntPer = " & tmp1.Fields(1) & " where tmpaccledger.AcCode = " & tmp1.Fields(0) & " and tmpaccledger.VFirm = '" & gCCode & "' and tmpaccledger.username='" & gUserName & "'"
                          tmp1.MoveNext
                       Loop
                    End If
                    gCn.CommitTrans
                    tmp1.Close
                
                End If
                '--For Withwout Cash & Bank Bal
                If gSelectedMenu = "Groupwise Account Ledger" And ChkIsCashBankGroup = 0 Then
                    gCn.Execute "Delete from tmpaccledger where agcode =  90019 and  UserName='" & gUserName & "'"
                End If
                '-- end
                If gSelectedMenu = "Opposite A/c wise Summery" Then       '-- Partywise
                    If optSelType(0).Value = True Then
                       gCn.Execute "Delete from tmpaccledger where OppPartywiseSumm =  0 and  UserName='" & gUserName & "'"
                    Else
                       gCn.Execute "Delete from tmpaccledger where OppPartywiseSummAudit =  0 and  UserName='" & gUserName & "'"
                    End If
                End If
                
                If gSelectedMenu = "Loan A/c Details" Then
                    gCn.BeginTrans
                    gCn.Execute "update tmpaccledger set UserCode = 0 where  UserName='" & gUserName & "'"
                    gCn.CommitTrans
                    
                    gCn.BeginTrans
                    gCn.Execute "Delete from tmpaccledger where LoanDet =  0 and  UserName='" & gUserName & "'"
                    gCn.CommitTrans
                    
                    gCn.BeginTrans
                    gCn.Execute "update tmpaccledger set AcOname = 'To Interest', UserCode = 1 where  UserName='" & gUserName & "' and vdramt > 0 and vtype <> '0P' and left(AcOname,8) = 'Interest' "
                    gCn.CommitTrans
                    
                    gCn.BeginTrans
                    gCn.Execute "update tmpaccledger set AcOname = 'To T. D. S.',UserCode = 1 where  UserName='" & gUserName & "' and vdramt > 0 and vtype <> '0P' and (left(AcOname,8) = 'T. D. S.' or left(AcOname,5) = 'T D S' or left(AcOname,3) = 'Tds' or left(AcOname,5) = 'T.D.S') "
                    gCn.CommitTrans
                    
                    gCn.BeginTrans
                    gCn.Execute "update tmpaccledger set AcOname = 'To Debited During The Year' where  UserName='" & gUserName & "' and vdramt > 0 and vtype <> '0P' and UserCode = 0"
                    gCn.CommitTrans
                
                    gCn.BeginTrans
                    gCn.Execute "update tmpaccledger set AcOname = 'By Interest', UserCode = 1 where  UserName='" & gUserName & "' and vcramt > 0 and vtype <> '0P' and left(AcOname,8) = 'Interest' "
                    gCn.CommitTrans
                    
                    gCn.BeginTrans
                    gCn.Execute "update tmpaccledger set AcOname = 'By T. D. S.',UserCode = 1 where  UserName='" & gUserName & "' and vcramt > 0 and vtype <> '0P' and (left(AcOname,8) = 'T. D. S.' or left(AcOname,5) = 'T D S' or left(AcOname,3) = 'Tds' or left(AcOname,5) = 'T.D.S') "
                    gCn.CommitTrans
                    
                    gCn.BeginTrans
                    gCn.Execute "update tmpaccledger set AcOname = 'By Credited During The Year' where  UserName='" & gUserName & "' and vcramt > 0 and vtype <> '0P' and UserCode = 0"
                    gCn.CommitTrans
                
                End If
                gCn.BeginTrans
                gCn.Execute "delete tmpaccledger  where  UserName='" & gUserName & "' and AcName = '-'"
                gCn.CommitTrans
                If gSelectedMenu = "Asset A/c Details" Then
                    gCn.Execute "Delete from tmpaccledger where AssetDet =  0 and  UserName='" & gUserName & "'"
                End If
                
                If ChkConsolidate.Value = 1 Then
                   X = "{tmpaccledger.UserName}= '" & gUserName & "' "
                Else
                   X = "{tmpaccledger.UserName}= '" & gUserName & "' and {tmpaccledger.Vfirm}='" & gCCode & "'   "
                End If
                .SelectionFormula = X
                .Formulas(4) = "wFrToDt = 'From : ' & '" & FromDt & "' & '  To ' & '" & ToDt & "'"
                .Formulas(5) = "wRptName = 'Account Ledger'"
                If optReportOn(2).Value = True Then
                    procAccounts Str(wAcCode), FromDt, ToDt
                   Exit Sub
                Else
                    If gSelectedMenu = "Sarafi Interest Report" Then
                      .Formulas(6) = "wToDt = '" & ToDt & "'"
                      .Formulas(7) = "wIntRt = " & Val(txttmp(0)) & ""
                      If optSelType(0).Value = True Then
                         .ReportFileName = gReportPath & "rptInterestSummery.rpt"
                      Else
                         .ReportFileName = gReportPath & "rptInterest.rpt"
                      End If
                    ElseIf gSelectedMenu = "Account Ledger" Then
                        If ChkConsolidate.Value = 1 Then
                           .ReportFileName = gReportPath & "rptAccountLedgerAllFirm.rpt"
                        Else
                           If optSelType(0).Value = True Then '-- Simple
                             .ReportFileName = gReportPath & "rptAccountLedger.rpt"
'                             .ReportFileName = gReportPath & "rptAccountLedger-Mum.rpt"
                           End If
                           If optSelType(1).Value = True Then       '-- Partywise
                             .Formulas(5) = "wRptName = 'Account Ledger Partywise'"
                             .ReportFileName = gReportPath & "rptAccountLedgerPartywise.rpt"
                           End If
                           If optSelType(2).Value = True Then '-- Monthwise
                             .Formulas(5) = "wRptName = 'Account Ledger Monthwise'"
                             .ReportFileName = gReportPath & "rptAccountLedgerMonthwise.rpt"
                           End If
                           If optSelType(3).Value = True Then '-- Status
                             .Formulas(5) = "wRptName = 'Account Status'"
                             .ReportFileName = gReportPath & "rptAccountLedgerStatus.rpt"
                           End If
                        End If
                    ElseIf gSelectedMenu = "Opposite A/c wise Summery" Then
                          .Formulas(5) = "wRptName = 'Account Ledger Summery'"
                          .Formulas(6) = "wCapitalNm = '" & gCCapitalName & "'"
                          .Formulas(7) = "wDesi = '(" & gCDesignation & ")'"
                          If ChkIsCashBankGroup.Value = 0 Then
                             .Formulas(8) = "wPrintCapital = 'No'"
                          Else
                             .Formulas(8) = "wPrintCapital = 'Yes'"
                          End If
                          '--- For Delete 0 bal Ac
                          gCn.Execute "Delete TmpAccLedger where vfirm+AcOname  in (select vfirm+AcOname " _
                                      & " from TmpAccLedger where Vfirm = '" & gCCode & "' and UserName = '" & gUserName & "' " _
                                      & " group by vfirm,AcOname having sum(vdramt) = sum(vcramt)) " _
                                      & " select *   from TmpAccLedger where vfirm+AcOname  in (select vfirm+AcOname   from TmpAccLedger where Vfirm = '" & gCCode & "' and UserName = '" & gUserName & "' " _
                                      & " group by vfirm,AcOname " _
                                      & " having sum(vdramt) = sum(vcramt))"
                          .ReportFileName = gReportPath & "rptAccountLedgerOppPtySummery.rpt"
                    ElseIf gSelectedMenu = "Loan A/c Details" Then
                          .Formulas(5) = "wRptName = 'Loan A/c Details'"
                          .Formulas(6) = "wCapitalNm = '" & gCCapitalName & "'"
                          .Formulas(7) = "wDesi = '(" & gCDesignation & ")'"
                          If optSelType(1).Value = True Then
                             .Formulas(8) = "wPrintCapital = 'No'"
                          Else
                             .Formulas(8) = "wPrintCapital = 'Yes'"
                          End If
                          .ReportFileName = gReportPath & "rptAccountLedgerLoanSummery.rpt"
                    ElseIf gSelectedMenu = "Asset A/c Details" Then
                          .Formulas(5) = "wRptName = 'Asset A/c Details'"
                          .Formulas(6) = "wCapitalNm = '" & gCCapitalName & "'"
                          .Formulas(7) = "wDesi = '(" & gCDesignation & ")'"
                          If optSelType(1).Value = True Then
                             .Formulas(8) = "wPrintCapital = 'No'"
                          Else
                             .Formulas(8) = "wPrintCapital = 'Yes'"
                          End If
                          .ReportFileName = gReportPath & "rptAccountLedgerAssetSummery.rpt"
                    ElseIf gSelectedMenu = "Groupwise Account Ledger" Then
                         If ChkIsRefNo.Value = 1 Then
                           .ReportFileName = gReportPath & "rptAccountLedgerGroupwiseIndex.rpt"
                         Else
                           .ReportFileName = gReportPath & "rptAccountLedgerGroupwise.rpt"
                         End If
'                        .ReportFileName = gReportPath & "rptAccountLedgerGroupwise-SRK.rpt"
                    ElseIf gSelectedMenu = "Account Confirmation" Then
                           .Formulas(5) = "wRptName = 'Account Confirmation'"
                           .ReportFileName = gReportPath & "rptAccountConfirmation.rpt"
                    ElseIf gSelectedMenu = "Depreciation Chart" Then
                          .Formulas(4) = "wFrToDt = 'From : ' & '" & FromDt & "' & '  To ' & '" & ToDt & "'"
                          .Formulas(6) = "PropNm = ' Proprietor Name : ' & '" & gCCapitalName & "'"
                          If Month(ToDt) >= 4 Then
                             Opbal1 = Month(ToDt) - 3
                          Else
                             Opbal1 = Month(ToDt) + 9
                          End If
                          .Formulas(5) = "wCalMon=" & Opbal1
                          .Formulas(7) = "wCapitalNm = '" & gCCapitalName & "'"
                          .Formulas(8) = "wDesi = '(" & gCDesignation & ")'"
                          If optSelType(1).Value = True Then
                             .Formulas(9) = "wPrintCapital = 'No'"
                          Else
                             .Formulas(9) = "wPrintCapital = 'Yes'"
                          End If
                          gCn.BeginTrans
                          gCn.Execute "delete tmpaccledger where left(AcOname,12) = 'Depreciation' and  UserName='" & gUserName & "' and  vcramt > 0 and vfirm = '" & gCCode & "'"
                          gCn.CommitTrans
                          .ReportFileName = gReportPath & "rptDeapeciationChart.rpt"
                    ElseIf gSelectedMenu = "TDS Receivable O/s" Or gSelectedMenu = "TCS Receivable O/s" Or gSelectedMenu = "TDS On Sales Receivable O/s" Then
                          If gSelectedMenu = "TCS Receivable O/s" Then
                             gCn.BeginTrans
                             If ChkConsolidate.Value = 1 Then
                                gCn.Execute "delete tmpaccledger where VType in ('PY','PT','PO','PE') and  UserName='" & gUserName & "' "
                             Else
                                gCn.Execute "delete tmpaccledger where VType in ('PY','PT','PO','PE') and  UserName='" & gUserName & "' and vfirm = '" & gCCode & "'"
                             End If
                             gCn.CommitTrans
                             
                             gCn.BeginTrans
                             
                             gCmd.CommandText = "PrcPrepareTcsReceivable"
                             gCmd.Parameters.Refresh
                             gCmd.Parameters("@From_dt") = Format(FromDt, "yyyy/mm/dd")
                             gCmd.Parameters("@To_dt") = Format(ToDt, "yyyy/mm/dd")
                             gCmd.Parameters("@UCode") = gUserName
                             gCmd.Parameters("@IsDepotFirm") = gCIsDepotFirm
                             gCmd.Parameters("@TcsRec") = gTcsRec
                             If ChkConsolidate.Value = 1 Then
                                gCmd.Parameters("@VFirm") = ""
                             Else
                                gCmd.Parameters("@VFirm") = Trim(gCCode)
                             End If
                             If wAcCode <> 0 Then gCmd.Parameters("@Accode") = wAcCode
                             If wAgCode <> 0 Then gCmd.Parameters("@Agcode") = wAgCode
                             gCmd.Execute
                             gCn.CommitTrans
                          
                          End If
                          gCn.Execute "update tmpaccledger Set Vnar4 = '-' where isnull(Vnar4,'') = ''  and tmpaccledger.UserName= '" & gUserName & "'"
                          gCn.Execute "update tmpaccledger Set Vno = 0 where isnull(Vno,'') = ''  and tmpaccledger.UserName= '" & gUserName & "'"
                          
                          If ChkConsolidate.Value = 1 Then
                             x1 = "Select * From tmpaccledger where tmpaccledger.UserName= '" & gUserName & "' "
                          Else
                             x1 = "Select * From tmpaccledger where tmpaccledger.UserName= '" & gUserName & "' and tmpaccledger.Vfirm='" & gCCode & "'   "
                          End If
                          gCn.BeginTrans
                          If GProcRstOpen(tmp1, x1, "O", gCn) > 0 Then
                             tmp1.MoveFirst
                              Do While Not tmp1.EOF
                                 If Len(tmp1.Fields("Vnar4")) <= 1 Then
                                    sTp = GProcGetColumnValue("tblMastAccount", "AcName", "" & tmp1.Fields("AcOName") & "", "S", "AcOurCode", "S")
                                    Y = GProcGetColumnValue("tblMastAccount", "AcName", "" & tmp1.Fields("AcOName") & "", "S", "AcPropNm", "S")
                                    X = GProcGetColumnValue("tblMastAccount", "AcName", "" & tmp1.Fields("AcOName") & "", "S", "AcTDSNature", "S")
                                  
                                 Else
                                    sTp = GProcGetColumnValue("tblMastAccount", "AcName", "" & tmp1.Fields("Vnar4") & "", "S", "AcOurCode", "S")
                                    Y = GProcGetColumnValue("tblMastAccount", "AcName", "" & tmp1.Fields("Vnar4") & "", "S", "AcPropNm", "S")
                                    X = GProcGetColumnValue("tblMastAccount", "AcName", "" & tmp1.Fields("Vnar4") & "", "S", "AcTDSNature", "S")
                                 End If
                                 gCn.Execute "update tmpaccledger set Vnar3 = '" & sTp & "'  where vtype =  '" & tmp1.Fields("Vtype") & "' and Vno = " & tmp1.Fields("Vno") & " and vfirm = '" & tmp1.Fields("VFirm") & "' and vdt = '" & Format(tmp1.Fields("Vdt"), "yyyy/mm/dd") & "'  and  UserName='" & gUserName & "' and AcOname = '" & tmp1.Fields("AcOname") & "' and Vnar4 = '" & tmp1.Fields("VNar4") & "' "
                                 gCn.Execute "update tmpaccledger set SubGrpAcname = '" & Y & "'  where vtype =  '" & tmp1.Fields("Vtype") & "' and Vno = " & tmp1.Fields("Vno") & " and vfirm = '" & tmp1.Fields("VFirm") & "' and vdt = '" & Format(tmp1.Fields("Vdt"), "yyyy/mm/dd") & "'  and  UserName='" & gUserName & "' and AcOname = '" & tmp1.Fields("AcOname") & "' and Vnar4 = '" & tmp1.Fields("VNar4") & "' "
                                 gCn.Execute "update tmpaccledger set VBillNo = '" & X & "'  where vtype =  '" & tmp1.Fields("Vtype") & "' and Vno = " & tmp1.Fields("Vno") & " and vfirm = '" & tmp1.Fields("VFirm") & "' and vdt = '" & Format(tmp1.Fields("Vdt"), "yyyy/mm/dd") & "'  and  UserName='" & gUserName & "' and AcOname = '" & tmp1.Fields("AcOname") & "' and Vnar4 = '" & tmp1.Fields("VNar4") & "' "
                                 tmp1.MoveNext
                              Loop
                         End If
                         tmp1.Close
                         gCn.CommitTrans
                          If gSelectedMenu = "TDS Receivable O/s" Then
                             .Formulas(5) = "wRptName = 'TDS Receivable O/s'"
                          ElseIf gSelectedMenu = "TDS On Sales Receivable O/s" Then
                             .Formulas(5) = "wRptName = 'TDS On Sales Receivable O/s'"
                          Else
                             .Formulas(5) = "wRptName = 'TCS Receivable O/s'"
                          End If
                         .Formulas(6) = "wCPAN = 'PAN No  :  ' & '" & gCPAN & "'"
                         .Formulas(7) = "wCapitalNm = '" & gCCapitalName & "'"
                          If chkDaySum.Value = 1 Then
                            gCn.BeginTrans
                            gCn.Execute "delete TmpAccLedger where VNar3 in (select VNar3  from TmpAccLedger where UserName = '" & gUserName & "' and Vfirm = '" & gCCode & "'" _
                                        & " group by VNar3 having sum(VDrAmt)-sum(VCrAmt) = 0) "
                            gCn.CommitTrans
                          End If
                          .ReportFileName = gReportPath & "rptTdsTcsOs.rpt"
                    End If
                End If
        Case "Account Groupwise Opening Balance List", "Account Opening Balance List"
             ShowRepo = True
             X = "Select iif(c.VDrAcCode = 1,crpt.AcOurCode,drpt.AcOurCode),iif(c.VDrAcCode = 1,crpt.AcName,drpt.AcName),c.VAmt,iif(c.VDrAcCode = 1,'Cr','Dr'),iif(c.VDrAcCode = 1,GpCrpt.AgName,GpDrpt.AgName),'" & gUserName & "' " _
                 & " from tblMastAccount Crpt,tblMastAccount Drpt,tblVoucher C,tblMastGroup GpCrpt,tblMastGroup GpDrpt where c.VCrAcCode = crpt.AcCode and c.VDrAcCode  = drpt.AcCode and " _
                 & " Crpt.AgCode = GpCrpt.AgCode and Drpt.AgCode = GpDrpt.AgCode and " _
                 & "c.Vfirm = '" & gCCode & "' and c.Vtype = '0P' And c.VYear = " & gCYear
             gCn.Execute "Delete from TmpGentbl where  UserName='" & gUserName & "'"
     
             gCn.Execute " insert into TmpGentbl (PartyCode,PartyName,Amt1,VBillType,Nar1,UserName) " & X
             
             X = "{TmpGentbl.UserName}= '" & gUserName & "' "
             If txt(0) <> "" Then
                X = X + " and  {TmpGentbl.Nar1} = '" & txt(0) & "' "
             End If
             .SelectionFormula = X
             .Formulas(4) = "wFrToDt = '" & gCYSDate & "'   "
             .Formulas(5) = "wRptName = 'Account Op. Balance List'"
             If gSelectedMenu = "Account Opening Balance List" Then
               .ReportFileName = gReportPath & "rptOpBalList.rpt"
             Else
               .ReportFileName = gReportPath & "rptOpBalListGroupwise.rpt"
             End If
        Case "Millwise Count List"
             ShowRepo = True
'             X = "isnull({tblMastItem.ItHsn})"
             X = ""
'             If txt(2) <> "" Then X = X + " and {tblMastAccount.ACNAME}='" & txt(2) & "'" '-Mill
             If txt(2) <> "" Then X = X + "{tblMastAccount.ACNAME}='" & txt(2) & "'" '-Mill
            
             .SelectionFormula = X
             .ReportFileName = gReportPath & "RptLstCountMillwise.rpt"
        Case "TDS On Sales ( Mill Bill )", "TDS On Sales ( Depot )"
             ShowRepo = True
             If gSelectedMenu = "TDS On Sales ( Depot )" Then
                X = "select tblMastAccount_Mill.AcOurCode, tblMastAccount_Mill.AcName,tblMastAccount.AcOurCode,tblMastAccount.AcName,tblMastAccount.AcPropNm,tblMastAccount.AcTDSNature,sum(SlTdsAmt),'" & gUserName & "' " _
                    & " from tblSale,tblMastAccount,tblMastAccount tblMastAccount_Mill " _
                    & " where tblSale.SlAcDrCode = tblMastAccount.AcCode and tblSale.SlAcCrCode = tblMastAccount_Mill.AcCode and " _
                    & " VType = 'SD' and VFirm = '" & gCCode & "' and vdt >= '" & Format(DTPicker1.Value, "YYYY/MM/dd") & "' and vdt <= '" & Format(DTPicker2.Value, "YYYY/MM/dd") & "' " _
                    & " group by tblMastAccount_Mill.AcOurCode,tblMastAccount_Mill.AcName,tblMastAccount.AcOurCode,tblMastAccount.AcName,tblMastAccount.AcPropNm,tblMastAccount.AcTDSNature " _
                    & " having sum(SlTdsAmt) > 0"
             Else
                X = "select tblMastAccount_Mill.AcOurCode, tblMastAccount_Mill.AcName,tblMastAccount.AcOurCode,tblMastAccount.AcName,tblMastAccount.AcPropNm,tblMastAccount.AcTDSNature, sum(SlTdsAmt),'" & gUserName & "' " _
                    & " from tblSale,tblMastAccount,tblMastAccount tblMastAccount_Mill " _
                    & " where tblSale.SlAcDrCode = tblMastAccount.AcCode and tblSale.SlAcCrCode = tblMastAccount_Mill.AcCode and " _
                    & " VType = 'SM' and VFirm = '" & gCCode & "' and vdt >= '" & Format(DTPicker1.Value, "YYYY/MM/dd") & "' and vdt <= '" & Format(DTPicker2.Value, "YYYY/MM/dd") & "' " _
                    & " group by tblMastAccount_Mill.AcOurCode,tblMastAccount_Mill.AcName,tblMastAccount.AcOurCode,tblMastAccount.AcName,tblMastAccount.AcPropNm,tblMastAccount.AcTDSNature " _
                    & " having sum(SlTdsAmt) > 0"
             End If
             gCn.Execute "Delete from TmpGentbl2 where  UserName='" & gUserName & "'"
             gCn.Execute " insert into TmpGentbl2 (MillCode,MillName,PartyCode,PartyName,Nar1,Nar3,Amt1,UserName) " & X
             '---For TDS JV Against InVoice
             If gSelectedMenu = "TDS On Sales ( Depot )" Then
                X = "select tblMastAccount_Mill.AcOurCode, tblMastAccount_Mill.AcName,tblMastAccount.AcOurCode,tblMastAccount.AcName,tblMastAccount.AcPropNm,tblMastAccount.AcTDSNature ,sum(tblVoucher.Vamt),'JI','" & gUserName & "' " _
                    & " from tblVoucher,tblSale,tblMastAccount,tblMastAccount tblMastAccount_Mill " _
                    & " where tblSale.SlAcDrCode = tblMastAccount.AcCode and tblSale.SlAcCrCode = tblMastAccount_Mill.AcCode and " _
                    & " tblVoucher.VBillType = 'SD' and tblVoucher.VFirm = '" & gCCode & "' and tblVoucher.vdt >= '" & Format(DTPicker1.Value, "YYYY/MM/dd") & "' and tblVoucher.vdt <= '" & Format(DTPicker2.Value, "YYYY/MM/dd") & "' " _
                    & " and tblVoucher.vtype = 'JI' and tblVoucher.VRefTp = '1' and tblVoucher.VBillVno = tblSale.vno and tblVoucher.VBillType = tblSale.Vtype and tblVoucher.VBillVYear = tblSale.Vyear and tblVoucher.VBillFirm = tblSale.Vfirm " _
                    & " group by tblMastAccount_Mill.AcOurCode,tblMastAccount_Mill.AcName,tblMastAccount.AcOurCode,tblMastAccount.AcName,tblMastAccount.AcPropNm,tblMastAccount.AcTDSNature "
             Else
                X = "select tblMastAccount_Mill.AcOurCode, tblMastAccount_Mill.AcName,tblMastAccount.AcOurCode,tblMastAccount.AcName,tblMastAccount.AcPropNm,tblMastAccount.AcTDSNature ,sum(tblVoucher.Vamt),'JI','" & gUserName & "' " _
                    & " from tblVoucher,tblSale,tblMastAccount,tblMastAccount tblMastAccount_Mill " _
                    & " where tblSale.SlAcDrCode = tblMastAccount.AcCode and tblSale.SlAcCrCode = tblMastAccount_Mill.AcCode and " _
                    & " tblVoucher.VBillType = 'SM' and tblVoucher.VFirm = '" & gCCode & "' and tblVoucher.vdt >= '" & Format(DTPicker1.Value, "YYYY/MM/dd") & "' and tblVoucher.vdt <= '" & Format(DTPicker2.Value, "YYYY/MM/dd") & "' " _
                    & " and tblVoucher.vtype = 'JI' and tblVoucher.VRefTp = '1' and tblVoucher.VBillVno = tblSale.vno and tblVoucher.VBillType = tblSale.Vtype and tblVoucher.VBillVYear = tblSale.Vyear and tblVoucher.VBillFirm = tblSale.Vfirm " _
                    & " group by tblMastAccount_Mill.AcOurCode,tblMastAccount_Mill.AcName,tblMastAccount.AcOurCode,tblMastAccount.AcName,tblMastAccount.AcPropNm,tblMastAccount.AcTDSNature "
             End If
             gCn.Execute " insert into TmpGentbl2 (MillCode,MillName,PartyCode,PartyName,Nar1,Nar3,Amt1,Nar2,UserName) " & X
             X = "{TmpGentbl2.UserName}= '" & gUserName & "' "
             If txt(2).text <> "" Then X = X + " and {TmpGentbl2.MillName}= '" & txt(2) & "'"
             .SelectionFormula = X
             .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(DTPicker1.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(DTPicker2.Value, "dd/MM/yyyy") & "'"
             If gSelectedMenu = "TDS On Sales ( Depot )" Then
                .Formulas(5) = "wRptName = 'TDS On Depot Sales'"
             Else
                .Formulas(5) = "wRptName = 'TDS On Mill Bill'"
             End If
             .ReportFileName = gReportPath & "rptTDSonMillBill.rpt"
        End Select
        If ShowRepo = True Then
           .Action = 1
           .PageZoom (120)
        End If
    End With
    Screen.MousePointer = vbDefault
ElseIf optReportOn(2) = True Then '-Grid
'    Pty_Ldr1 ("F")
End If
End Sub
Private Sub DTPicker2_LostFocus()
If (gSelectedMenu = "R2M") Then
    lstParty_GotFocus
    Exit Sub
End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case 13
    Sendkeys "{TAB}"
Case 27
    Unload Me
End Select
End Sub
Private Sub Form_Load()
Dim i As Long
Dim Cdt As String
Me.Top = 700
Me.Left = Screen.Width / 2 - Me.Width / 2
 '-- Activate & Set  Controls
GProcActivateControls Me, False
DTPicker1.Enabled = True
DTPicker2.Enabled = True
lstParty.Visible = False
For i = 0 To 2 '-- Report On
    optReportOn(i).Enabled = True
Next i
'-- Set Values
DTPicker1.Value = Format(gCYSDate, "dd/MM/yyyy")

Cdt = Format(Now, "dd/MM/yyyy")

If CDate(Cdt) > gCYEDate Then
   DTPicker2.Value = Format(gCYEDate, "dd/MM/yyyy")
Else
   DTPicker2.Value = Format(Now, "dd/MM/yyyy")
End If

'DTPicker2.Value = Format(gCYEDate, "dd/MM/yyyy")
For i = 0 To 4
    txttmp(i) = 0
Next i
For i = 0 To 2
    txtCode(i) = ""
Next i

'-Days of year
txttmp(4) = DateDiff("d", gCYSDate, gCYEDate) + 1
CheckForNumber
'------------------------------------
mScreen = True
Me.Caption = gSelectedMenu
Select Case gSelectedMenu
    Case "Depreciation Chart", "Loan A/c Details", "Asset A/c Details"
         Label4.Visible = False
         optSelType(0).Visible = False
         optSelType(2).Visible = False
         optSelType(1).Enabled = True
         optSelType(3).Enabled = True
         optSelType(1).Value = True
         optSelType(1).Caption = "Without Capital Name"
         optSelType(3).Caption = "With Capital Name"
    Case "Opposite A/c wise Summery"
         optSelType(0).Enabled = True
         optSelType(0).Caption = "Total"
         optSelType(0).Value = True
         optSelType(2).Enabled = True
         optSelType(2).Caption = "For Audit"
         Label4.Visible = False
         optSelType(1).Visible = False
         optSelType(3).Visible = False
         ChkIsCashBankGroup.Visible = True
         ChkIsCashBankGroup.Enabled = True
         ChkIsCashBankGroup.Value = 0
         ChkIsCashBankGroup.Caption = "With Capital Name"
    Case "Account Ledger", "Sarafi Interest Report", "Account Confirmation", "Groupwise Account Ledger" '-"Account Ledger"
        txt(1).Visible = True  '-A/c
        txt(1).Enabled = True '-A/c
        txtCode(1).Visible = True
        txtCode(1).Enabled = True
        Label3.Visible = True
        ChkConsolidate.Visible = True
        ChkConsolidate.Enabled = True
        Label4.Visible = False
        If gSelectedMenu = "Account Ledger" Then
            For i = 0 To 3 '-- Report format
                optSelType(i).Enabled = True
            Next i
        End If
        Frame1.Visible = True
        If gSelectedMenu = "Account Confirmation" Then
           txtCode(0).Visible = True
           txtCode(0).Enabled = True
           txt(0).Visible = True  '-Grp A/c
           txt(0).Enabled = True '-grp A/c
           Label5.Visible = True
        End If
        If gSelectedMenu = "Groupwise Account Ledger" Then
           txtCode(0).Visible = True
           txtCode(0).Enabled = True
           txt(0).Visible = True  '-Grp A/c
           txt(0).Enabled = True '-grp A/c
           Label5.Visible = True
           ChkConsolidate.Visible = False
           ChkConsolidate.Enabled = False
           ChkIsCashBankGroup.Visible = True
           ChkIsCashBankGroup.Enabled = True
           ChkIsCashBankGroup.Value = 0
           ChkIsRefNo.Value = 0
           ChkIsRefNo.Caption = "Index Page"
           ChkIsRefNo.Visible = True
           ChkIsRefNo.Enabled = True
        End If
        If gSelectedMenu = "Sarafi Interest Report" Then
           ChkConsolidate.Visible = False
           ChkConsolidate.Enabled = False
'           txttmp(0).Visible = True
'           txttmp(0).Enabled = True
'           Label1(2).Visible = True
           txtCode(0).Visible = True
           txtCode(0).Enabled = True
           txt(0).Visible = True  '-Grp A/c
           txt(0).Enabled = True '-grp A/c
           txt(0) = "Loans Liabilities"
           Label5.Visible = True
           optSelType(0).Value = True
           optSelType(0).Enabled = True
           optSelType(0).Caption = "Summery"
           optSelType(1).Enabled = True
           optSelType(1).Caption = "Detail"
           optSelType(2).Visible = False
           optSelType(3).Visible = False
           chkDaySum.Visible = True
           chkDaySum.Enabled = True
           chkDaySum.Caption = "All Accounts ?"
        End If
    Case "Account Groupwise Opening Balance List", "Account Opening Balance List"
       If gSelectedMenu = "Account Opening Balance List" Then
         Label5.Visible = False
         txt(0).Visible = False
         txt(0).Enabled = False
         txtCode(0).Visible = False
         txtCode(0).Enabled = False
       Else
         Label5.Visible = True
         txt(0).Visible = True
         txt(0).Enabled = True
         txtCode(0).Visible = True
         txtCode(0).Enabled = True
       End If
       Label1(0).Visible = False
       Label2.Visible = False
       DTPicker1.Visible = False
       DTPicker2.Visible = False
       Label4.Visible = False
       For i = 0 To 3 '-- Report format
           optSelType(i).Visible = False
       Next i
       Shape1.Visible = False
       Line1.Visible = False
       Line2.Visible = False
   Case "Millwise Count List"
      DTPicker1.Visible = False
      DTPicker2.Visible = False
      txtCode(2).Visible = True
      txt(2).Visible = True
      DTPicker1.Enabled = True
      DTPicker2.Enabled = True
      txtCode(2).Enabled = True
      txt(2).Enabled = True
      Label7.Visible = True
   Case "TDS Receivable O/s", "TCS Receivable O/s", "TDS On Sales Receivable O/s"
      Label4.Visible = False
      optSelType(0).Visible = False
      optSelType(1).Visible = False
      optSelType(2).Visible = False
      optSelType(3).Visible = False
      Label3.Visible = True
      txtCode(1).Visible = True
      txt(1).Visible = True
      chkDaySum.Visible = True
      chkDaySum.Enabled = True
      chkDaySum.Caption = "Without 0 Balance"
      If gSelectedMenu = "TDS On Sales ( Mill Bill ) Receivable O/s" Then
         txtCode(1).Visible = False
         txt(1).Visible = False
      Else
        If gSelectedMenu = "TDS Receivable O/s" Then
           wAcCode = gTDSAcCodeRec
        ElseIf gSelectedMenu = "TDS On Sales Receivable O/s" Then
           wAcCode = gTdsOnSalesCode
        Else
           wAcCode = gTcsRec
        End If
      End If
      txtCode(1) = GProcGetColumnValue("tblMastAccount", "AcCode", "" & wAcCode & "", "N", "AcOurCode", "S")
      txt(1) = GProcGetColumnValue("tblMastAccount", "AcCode", "" & wAcCode & "", "N", "AcName", "S")
Case "TDS On Sales ( Mill Bill )", "TDS On Sales ( Depot )"
      txtCode(2).Visible = True
      txt(2).Visible = True
      txtCode(2).Enabled = True
      txt(2).Enabled = True
      Label7.Visible = True
End Select
 '-- Set InActivate Ctrl Back Color
GProcSetInActivateCtrlBackClr Me
End Sub

Private Sub lstParty_GotFocus()
'lstParty.Clear
'Dim i As Integer
'Call iisGPrActivateControls(Me, False)
'TopPos = Me.Top
''+ lstParty.Top
'LeftPos = Me.Left
''+ lstParty.Left
'    '----Account Name
'    mSelHead = "Select Multiple Account Name "
'    mySele = lstParty.Text
'    mySql = "select ANAME from ACCT where ACTYPE='A'"
'    myField = "ANAME"
'    myCnd = " and "
'    DoEvents
'    frmSearchMultiSel.Show 1
'    Do While mRet(i) <> ""
'        lstParty.AddItem (mRet(i))
'        DoEvents
'        wPartyCd(i) = iisGFnGetColumnValue(gCn, "select ACCODE from ACCT where ANAME = '" & mRet(i) & "'")
'        i = i + 1
'    Loop
'Call iisGPrActivateControls(Me, True)
'Command1.SetFocus
End Sub
Private Sub txt_GotFocus(Index As Integer)
'If index = 1 And lstParty.Visible = True Then
 '   lstParty.SetFocus
'Else
    GProcSelectBox Me.ActiveControl
'End If
End Sub
Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
Dim LeftPos As Long
Dim TopPos As Long
TopPos = txt(Index).Top + Me.Top + 650
LeftPos = txt(Index).Left + Me.Left
If KeyAscii = 13 Then Exit Sub
Select Case Index
    Case 0 '-Group
        If gSelectedMenu = "Party Groupwise List" Then
           gClsSearch.SearchMultiField "tblMastCompGroup", "CGName", Array("Company Group"), Array(2000), " ", Chr(KeyAscii), "CGName", LeftPos, TopPos
           KeyAscii = 0
           If gClsSearch.SearchMultiRetCol(0) <> "" Then
              txt(Index) = gClsSearch.SearchMultiRetCol(0)
              txtCode(Index) = GProcGetColumnValue("tblMastAccount", "AcName", txt(Index), "S", "AcOurCode", "S")
           End If
        Else
           gClsSearch.SearchMultiField "tblMastGroup", "AGName,Agcode", Array("Account Group", "Agcode"), Array(txt(Index).Width, 0), " Not(Agcode = TopGroup)  ", Trim(Chr(KeyAscii)), "", LeftPos, TopPos, False, , False
           KeyAscii = 0
           If gClsSearch.SearchMultiRetCol(0) <> "" Then
              txt(Index) = gClsSearch.SearchMultiRetCol(0)
              wAgCode = gClsSearch.SearchMultiRetCol(1)
           End If
        End If
    Case 1 '-Account
        '-Selected Group
'        If txt(0) = "" And gSelectedMenu <> "Account List" Then
'            gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,Accode,AcCity,AgName", Array("Account", "AcCode", "AcName", "City", "Group"), Array(4000, 0, 0, 2000, 3000), "", Chr(KeyAscii), "AcName", LeftPos, TopPos
'        ElseIf gSelectedMenu <> "Account List" And txt(0) <> "" Then
'            gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,Accode,AcCity,AgName", Array("Account", "AcCode", "AcName", "City", "Group"), Array(3000, 0, 0, 2000, 3000), "AgName='" & txt(0) & "'", Chr(KeyAscii), "AcName", LeftPos, TopPos
'        ElseIf gSelectedMenu = "Account List" Then
'               gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,Accode,AcCity,AgName", Array("Account", "AcCode", "AcName", "City", "Group"), Array(3000, 0, 0, 2000, 3000), "Agcode=90016", Chr(KeyAscii), "AcName", LeftPos, TopPos
'        End If
        gClsSearch.SearchMultiFieldAccMaster Chr(KeyAscii), LeftPos, TopPos
        KeyAscii = 0
        If gClsSearch.SearchMultiRetColAccMaster(0) <> "" Then
            txt(Index) = gClsSearch.SearchMultiRetColAccMaster(0)
            wAcCode = gClsSearch.SearchMultiRetColAccMaster(3)
            txtCode(Index) = GProcGetColumnValue("tblMastAccount", "AcName", txt(Index), "S", "AcOurCode", "S")
        End If
'        If gClsSearch.SearchMultiRetColAccMaster(0) <> "" And gClsSearch.SearchMultiRetColAccMaster(1) = "" Then
'            txt(Index) = gClsSearch.SearchMultiRetColAccMaster(0)
'            wAcCode = gClsSearch.SearchMultiRetColAccMaster(2)
'        ElseIf gClsSearch.SearchMultiRetColAccMaster(0) <> "" And gClsSearch.SearchMultiRetColAccMaster(1) <> "" Then
'            txt(Index) = gClsSearch.SearchMultiRetColAccMaster(1)
'            wAcCode = gClsSearch.SearchMultiRetColAccMaster(2)
'        End If
    Case 2 '-Mill
        If gSelectedMenu = "Account Confirmation" Or gSelectedMenu = "Sarafi Interest Report" Or gSelectedMenu = "Sarafi Interest Report ANK" Or gSelectedMenu = "Account Closing Balance List" Then
           gClsSearch.SearchMultiField "tblMastCompGroup", "CGName", Array("Company Group"), Array(2000), " ", Chr(KeyAscii), "CGName", LeftPos, TopPos
           KeyAscii = 0
           If gClsSearch.SearchMultiRetCol(0) <> "" Then
              txt(Index) = gClsSearch.SearchMultiRetCol(0)
           End If
        Else
            gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,Accode,AcCity,AgName", Array("Account", "AcCode", "AcName", "City", "Group"), Array(3000, 0, 0, 2000, 3000), " AgCode=90029", Chr(KeyAscii), "AcName", LeftPos, TopPos
            KeyAscii = 0
            If gClsSearch.SearchMultiRetCol(0) <> "" And gClsSearch.SearchMultiRetCol(1) = "" Then
                txt(Index) = gClsSearch.SearchMultiRetCol(0)
                wAcCode = gClsSearch.SearchMultiRetCol(2)
                txtCode(Index) = GProcGetColumnValue("tblMastAccount", "AcName", txt(Index), "S", "AcOurCode", "S")
            ElseIf gClsSearch.SearchMultiRetCol(0) <> "" And gClsSearch.SearchMultiRetCol(1) <> "" Then
                txt(Index) = gClsSearch.SearchMultiRetCol(1)
                wAcCode = gClsSearch.SearchMultiRetCol(2)
                txtCode(Index) = GProcGetColumnValue("tblMastAccount", "AcName", txt(Index), "S", "AcOurCode", "S")
            End If
        End If
End Select
End Sub
Private Sub txtCode_LostFocus(Index As Integer)
If Trim(txtCode(Index)) = "" Then Exit Sub
txt(Index) = GProcGetColumnValue("tblMastAccount", "AcOurCode", UCase(txtCode(Index)), "S", "AcName", "S")
Select Case Index
Case 0
    wAgCode = GProcGetColumnValue("tblMastAccount", "AcName", txt(0), "S", "AcCode", "N")
Case 1
    wAcCode = GProcGetColumnValue("tblMastAccount", "AcName", txt(1), "S", "AcCode", "N")
Case 2
    wAcCode = GProcGetColumnValue("tblMastAccount", "AcName", txt(2), "S", "AcCode", "N")
End Select
End Sub
Private Sub txttmp_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub

Private Sub txttmp_LostFocus(Index As Integer)
CheckForNumber
End Sub
