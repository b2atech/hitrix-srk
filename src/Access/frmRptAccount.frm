VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmRptAccount 
   BackColor       =   &H80000013&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Report"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10770
   Icon            =   "frmRptAccount.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   10770
   Begin VB.CommandButton cmdSPE 
      Caption         =   "W.App"
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
      Index           =   3
      Left            =   4410
      TabIndex        =   46
      Top             =   4950
      Visible         =   0   'False
      Width           =   1290
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
      Index           =   6
      Left            =   7680
      TabIndex        =   31
      Text            =   "0"
      Top             =   4560
      Width           =   1275
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
      Left            =   7680
      TabIndex        =   32
      Text            =   "0"
      Top             =   5040
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
      Index           =   4
      Left            =   7590
      TabIndex        =   25
      Text            =   "4"
      Top             =   3285
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
      Left            =   8595
      TabIndex        =   28
      Text            =   "3"
      Top             =   2925
      Width           =   1455
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
      Left            =   9945
      TabIndex        =   27
      Text            =   "2"
      Top             =   4170
      Width           =   555
   End
   Begin VB.CheckBox ChkSummary 
      BackColor       =   &H80000018&
      Caption         =   "Is Summary ?"
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   6345
      TabIndex        =   29
      Top             =   3720
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
      Left            =   9945
      TabIndex        =   26
      Text            =   "1"
      Top             =   3735
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
      Left            =   7605
      TabIndex        =   24
      Text            =   "0"
      Top             =   2850
      Width           =   780
   End
   Begin VB.CheckBox ChkIntRateMonthly 
      BackColor       =   &H80000018&
      Caption         =   "Is Interest Rate Monthly ?"
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   6345
      TabIndex        =   30
      Top             =   4125
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.TextBox txt 
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
      Height          =   375
      Index           =   2
      Left            =   1605
      TabIndex        =   9
      Top             =   1695
      Width           =   4335
   End
   Begin VB.CheckBox ChkIsCashBankGroup 
      BackColor       =   &H80000018&
      Caption         =   "&Cash and Bank Book"
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
      Left            =   8040
      TabIndex        =   21
      Top             =   2070
      Value           =   1  'Checked
      Width           =   2490
   End
   Begin VB.CheckBox ChkIsRefNo 
      BackColor       =   &H80000018&
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
      Left            =   6210
      TabIndex        =   20
      Top             =   2055
      Value           =   1  'Checked
      Width           =   1830
   End
   Begin VB.CommandButton cmdSPE 
      Height          =   495
      Index           =   0
      Left            =   270
      Picture         =   "frmRptAccount.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   4950
      Width           =   1245
   End
   Begin VB.CommandButton cmdSPE 
      Height          =   495
      Index           =   1
      Left            =   1650
      Picture         =   "frmRptAccount.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   "Print"
      Top             =   4950
      Width           =   1245
   End
   Begin VB.TextBox txt 
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
      Height          =   375
      Index           =   0
      Left            =   1605
      TabIndex        =   5
      Top             =   735
      Width           =   4335
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
      TabIndex        =   41
      Top             =   2595
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   10
         Top             =   360
         Width           =   1425
      End
   End
   Begin VB.OptionButton optSelType 
      BackColor       =   &H80000018&
      Caption         =   "&T Format"
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
      Left            =   7755
      TabIndex        =   17
      Top             =   915
      Width           =   1425
   End
   Begin VB.OptionButton optSelType 
      BackColor       =   &H80000018&
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
      Left            =   6315
      TabIndex        =   16
      Top             =   900
      Width           =   1425
   End
   Begin VB.OptionButton optSelType 
      BackColor       =   &H80000018&
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
      Left            =   7755
      TabIndex        =   15
      Top             =   540
      Width           =   1425
   End
   Begin VB.OptionButton optSelType 
      BackColor       =   &H80000018&
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
      Left            =   6315
      TabIndex        =   14
      Top             =   525
      Value           =   -1  'True
      Width           =   1425
   End
   Begin VB.CheckBox chkDaySum 
      BackColor       =   &H80000018&
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
      Left            =   6315
      TabIndex        =   18
      Top             =   1305
      Width           =   3135
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
      Left            =   3045
      TabIndex        =   39
      Top             =   4950
      Width           =   1245
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1605
      TabIndex        =   1
      Top             =   255
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   96993283
      CurrentDate     =   38050
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   4380
      TabIndex        =   3
      Top             =   255
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   96993283
      CurrentDate     =   38050
   End
   Begin VB.TextBox txt 
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
      Height          =   375
      Index           =   1
      Left            =   1605
      TabIndex        =   7
      Top             =   1200
      Width           =   4335
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
      Left            =   1605
      Sorted          =   -1  'True
      TabIndex        =   40
      Top             =   1200
      Width           =   4335
   End
   Begin MSComctlLib.ProgressBar prgBar 
      Height          =   255
      Left            =   195
      TabIndex        =   42
      Top             =   3765
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
      Caption         =   "T.D.S. Limit  :"
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
      Index           =   8
      Left            =   6345
      TabIndex        =   45
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "T.D.S. Rate   :"
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
      Left            =   6345
      TabIndex        =   44
      Top             =   5070
      Width           =   1245
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
      Left            =   8550
      TabIndex        =   43
      Top             =   3390
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
      Left            =   8550
      TabIndex        =   36
      Top             =   4170
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
      Left            =   6345
      TabIndex        =   34
      Top             =   3240
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
      Left            =   8550
      TabIndex        =   37
      Top             =   2610
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
      Left            =   8550
      TabIndex        =   35
      Top             =   3735
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
      Left            =   6345
      TabIndex        =   23
      Top             =   2850
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
      Left            =   6270
      TabIndex        =   22
      Top             =   2535
      Width           =   3540
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000040C0&
      X1              =   10590
      X2              =   6150
      Y1              =   2430
      Y2              =   2430
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000040C0&
      X1              =   10575
      X2              =   6105
      Y1              =   1650
      Y2              =   1650
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000040C0&
      X1              =   6120
      X2              =   75
      Y1              =   2295
      Y2              =   2295
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Mill             :"
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
      TabIndex        =   8
      Top             =   1695
      Width           =   1140
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
      Left            =   6270
      TabIndex        =   19
      Top             =   1755
      Width           =   2640
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A/C &Group   :"
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
      TabIndex        =   4
      Top             =   735
      Width           =   1155
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
      Left            =   6270
      TabIndex        =   13
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
      Left            =   3285
      TabIndex        =   2
      Top             =   255
      Width           =   1215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Account       :"
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
      TabIndex        =   6
      Top             =   1200
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "From &Date   :"
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
      Width           =   1230
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   4530
      Left            =   75
      Top             =   90
      Width           =   6000
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00BDDFD1&
      FillStyle       =   0  'Solid
      Height          =   765
      Left            =   75
      Shape           =   4  'Rounded Rectangle
      Top             =   4800
      Width           =   5760
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   5370
      Left            =   6120
      Top             =   120
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
Dim wAcCode As Long, mScreen As Boolean, mP12 As String, mPc As String
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
'    If i = 0 Then '-Rate Double
'        txttmp(i) = Format(txttmp(i), FStr)
'    Else
'        txttmp(i) = Format(txttmp(i), PStr)
'    End If
    If i = 0 Then '-Rate Double
        If gSelectedMenu = "Sarafi Interest Report ANK" Then
           txttmp(i) = Format(txttmp(i), "###0.00000")
        ElseIf gSelectedMenu = "TCS Report" Then
           txttmp(i) = Format(txttmp(i), WtStr)
        Else
           txttmp(i) = Format(txttmp(i), FStr)
        End If
    Else
        txttmp(i) = Format(txttmp(i), PStr)
    End If

Next i
End Sub
'-- Validate Data
Private Function ValidateData() As Boolean
ValidateData = True
Select Case gSelectedMenu
    Case "Sarafi Interest Report", "Sarafi Interest Report ANK" '-"Sarafi Interest Report"
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
            
     Case "Groupwise Account List"
           If txt(0).text = "" And ChkIsCashBankGroup.Value = 1 Then
               ValidateData = False
               MsgBox "select Group Name ", vbOKOnly + vbInformation
               txt(0).SetFocus
               Exit Function
           End If

'            ValidateData = False
'            MsgBox "Check Total Days of Year", vbOKOnly + vbInformation
'            txttmp(4).SetFocus
'            Exit Function
'        End If
End Select
End Function

Private Sub chkDaySum_Click()
If gSelectedMenu = "Sarafi Interest Report" Then
   If chkDaySum.Value = 1 Then
      Label1(8).Caption = "CC Limit :"
      Label1(7).Visible = False
      txttmp(5).Visible = False
   Else
      Label1(8).Caption = "T.D.S. Limit  :"
      Label1(7).Visible = True
      txttmp(5).Visible = True
   End If
End If

End Sub

Private Sub cmdSPE_Click(Index As Integer)
Dim sTp As String
Dim FromDt As Date
Dim ToDt As Date
Dim X As String
Dim X2 As String
Dim X3 As String
Dim X4 As String
Select Case Index
Case 0
    mScreen = True
Case 1
    mScreen = False
Case 2
   Unload Me
   Exit Sub
End Select
If ValidateData = False Then Exit Sub
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
   Pty_Ldr1 (sTp)
'== Window
ElseIf optReportOn(1) = True Then
    If mScreen Then
        GProcCrystalRptPreparation CrystalReport1, "Window"
    Else
        GProcCrystalRptPreparation CrystalReport1, "Printer"
    End If
    With CrystalReport1
        Select Case gSelectedMenu
           Case "TCS Report"
                X2 = " qrytblVoucherSr1.vdt>=cdate('" & Format(gCYSDate, "yyyy,mm,dd") & "')  and qrytblVoucherSr1.vdt < Cdate('" & Format(FromDt, "yyyy,mm,dd") & "')"
                X3 = " qrytblVoucherSr1.vdt>=cdate('" & Format(FromDt, "yyyy,mm,dd") & "')  and qrytblVoucherSr1.vdt<=Cdate('" & Format(ToDt, "yyyy,mm,dd") & "')"
                If ChkIsRefNo.Value = 1 Then
                    X = "SELECT qrytblVoucherSr1.VacCode, tblMastAccount.AcName, tblMastAccount.Agcode,tblMastAccount.AcPAN,qrytblVoucherSr1.Vno,qrytblVoucherSr1.Vtype, qrytblVoucherSr1.Vdt,iif(isnull(qrytblVoucherSr1.VNar1),'',left(qrytblVoucherSr1.VNar1,40)), " _
                        & " iif( " & X2 & ", qrytblVoucherSr1.Vamt,0),iif( " & X3 & ", qrytblVoucherSr1.Vamt,0) ," _
                        & " '" & gUserName & "' FROM qrytblVoucherSr1,tblMastAccount WHERE qrytblVoucherSr1.MILLEXPNARRCODE <> 1  and  qrytblVoucherSr1.MILLEXPNARRCODE <> 1  and qrytblVoucherSr1.VIsAudited <> 1 and qrytblVoucherSr1.VacCode=tblMastAccount.AcCode AND qrytblVoucherSr1.vtype in ('BR','CR') and tblMastAccount.Agcode = 90017 and tblMastAccount.AcIsTDSPay <> 1 and qrytblVoucherSr1.vdt<=Cdate('" & Format(ToDt, "yyyy,mm,dd") & "')"
                
                Else
                    X = "SELECT qrytblVoucherSr1.VacCode, tblMastAccount.AcName, tblMastAccount.Agcode,tblMastAccount.AcPAN,qrytblVoucherSr1.Vno,qrytblVoucherSr1.Vtype, qrytblVoucherSr1.Vdt,iif( " & X2 & ", qrytblVoucherSr1.Vamt,0),iif( " & X3 & ", qrytblVoucherSr1.Vamt,0) ," _
                        & " iif(isnull(tblRecVsSale.InvBillNo),'',tblRecVsSale.InvBillNo),iif(isnull(tblRecVsSale.Invdt),qrytblVoucherSr1.Vdt,tblRecVsSale.Invdt)," _
                        & " iif(isnull(tblRecVsSale.Invtp),'',tblRecVsSale.Invtp), iif( " & X2 & ", iif(isnull(tblRecVsSale.AdjAmt),0, tblRecVsSale.AdjAmt),0),iif( " & X3 & ", iif(isnull(tblRecVsSale.AdjAmt),0, tblRecVsSale.AdjAmt),0) ," _
                        & " '" & gUserName & "' FROM (qrytblVoucherSr1 LEFT JOIN tblRecVsSale ON (qrytblVoucherSr1.VYear = tblRecVsSale.RecVYear) AND (qrytblVoucherSr1.VType = tblRecVsSale.RecVType) AND (qrytblVoucherSr1.VNo = tblRecVsSale.RecVNo)) INNER JOIN tblMastAccount ON qrytblVoucherSr1.VAcCode = tblMastAccount.AcCode " _
                        & " WHERE qrytblVoucherSr1.MILLEXPNARRCODE <> 1  and qrytblVoucherSr1.VIsAudited <> 1 and (((qrytblVoucherSr1.VacCode)=[tblMastAccount].[AcCode]) AND (qrytblVoucherSr1.vtype in ('BR','CR'))) and tblMastAccount.AcIsTDSPay <> 1"
                End If
                If txt(1) <> "" Then
                   X = X + " and qrytblVoucherSr1.VacCode = " & wAcCode & ""
                End If
                gCn.BeginTrans
                gCn.Execute "delete from tmpRptTbl Where Username='" & gUserName & "'"
                If ChkIsRefNo.Value = 1 Then
                   gCn.Execute "insert into tmpRptTbl (VACCODE,Nar1,VACOCODE,Nar2,Vno,VTYPE,Vdt,Nar3,Amt1,Amt2,USERNAME) " & X
                Else
                   gCn.Execute "insert into tmpRptTbl (VACCODE,Nar1,VACOCODE,Nar2,Vno,VTYPE,Vdt,Amt1,Amt2,VBillNo,VBillDt,Nar3,Amt3,Amt4,USERNAME) " & X
                End If
                gCn.CommitTrans
                 .DataFiles(0) = gDbLocation
                 gReportPath = App.Path & "\HiReports_Access\"
                 .Formulas(4) = "wFrToDt = 'From : ' & '" & FromDt & "' & '  To ' & '" & ToDt & "'"
                 .Formulas(5) = "wFromDt=   '" & FromDt & "'"
                 .Formulas(6) = "wcPan=   '" & gCPAN & "'"
                 .Formulas(10) = "wcTan=   '" & gCTAN & "'"
                 .Formulas(7) = "wTcsRt=" & Format(Val(txttmp(0)), "###0.000")
                 .Formulas(8) = "wTcsLimit=" & Format(Val(txttmp(3)), "######0.00")
                 
                 X = "{tmpRptTbl.UserName}='" & gUserName & "' "
                 If txt(1) <> "" Then X = X + " and {tmpRptTbl.nar1}='" & txt(1) & "'  "
                 
                 .SelectionFormula = X
                 If optSelType(1).Value = True Then
                    If ChkIsRefNo.Value = 1 Then
                       .ReportFileName = gReportPath & "rptTCSOnReceiptSummeryWithInt.rpt"
                    Else
                       .ReportFileName = gReportPath & "rptTCSOnReceiptSummery.rpt"
                    End If
                 Else
                    If ChkIsRefNo.Value = 1 Then
                      .ReportFileName = gReportPath & "rptTCSOnReceiptWithInt.rpt"
                    Else
                      .ReportFileName = gReportPath & "rptTCSOnReceipt.rpt"
                    End If
                 End If
           '========== List Start
           Case "Account List" '== "Accounts List"
                   'GProcCrystalRptPreparation cryReport, "Window"
                   '.WindowTitle = "Account List"
                    If txt(1) <> "" Then X = "{tblMastaccount.acbrkcode}=" & GProcGetColumnValue("tblMastAccount", "AcName", txt(1), "S", "AcCode", "N") & "" '-boker
                    .SelectionFormula = X
                   .ReportFileName = gReportPath & "RptLstAccount.rpt"
            
    
            Case "Selected Group Accounts List" '== "Selected Group Accounts List"
                 If txt(0) <> "" Then X = "{tblMastGroup.AGNAME}='" & txt(0) & "'" '-Group
                 .SelectionFormula = X
                 If ChkIsRefNo.Value = 1 Then
                     .ReportFileName = gReportPath & "RptLstAccountGrpTrn.rpt"
                 ElseIf ChkIsCashBankGroup.Value = 1 Then
                    .ReportFileName = gReportPath & "RptLstAccEml.rpt"
                 Else
                    .ReportFileName = gReportPath & "RptLstSelGroupAccount.rpt"
                 End If
            Case "Mill wise Count List" '=== "Millwise Count List"
                If txt(2) <> "" Then X = "{tblMastAccount.ACNAME}='" & txt(2) & "'" '-Mill
                .SelectionFormula = X
                .ReportFileName = gReportPath & "RptLstCountMillwise.rpt"
            Case "Count Brokerage List" '== "Count Brokerage List"
                X = "({tblMastItem.ITBROKRT}>0 or {tblMastItem.ITOUTBRKRT}>0)"
                If txt(2) <> "" Then X = X + " and {tblMastAccount.ACNAME}='" & txt(2) & "'" '-Mill
                .SelectionFormula = X
                .ReportFileName = gReportPath & "RptLstCountBrokerage.rpt"
            Case "Account Groupwise Opening Balance List"  '== "Account Groupwise Opening Balance List"
                X = "{tblVoucher.VType}='0P'"
                If txt(0) <> "" Then X = X + " and {tblMastGroup.AGNAME}='" & txt(0) & "'"    '-Group
                .SelectionFormula = X
                .ReportFileName = gReportPath & "RptLstAccOpBalGroupwise.rpt"
            Case "Account Closing Balance List", "Account Groupwise Closing Balance List" '=== "Account Closing Balance List"
                    gCn.Execute "delete from tmpselection where UserName='" & gUserName & "'"
                    If gBackEndDB = gBackEndAccess Then
                        gCn.Execute "insert into tmpselection (dt1,Dt2,UserName) values ('" & gCYSDate & "','" & ToDt & "','" & gUserName & "')"
                    Else
                        gCn.Execute "insert into tmpselection (dt1,dt2,UserName) values (to_date('" & gCYSDate & "','dd/MM/yyyy'),to_Date('" & ToDt & "','dd/MM/yyyy'),'" & gUserName & "') "
                    End If
                    '--Set Loacation
                    If gBackEndDB = gBackEndAccess Then
                        .DataFiles(0) = gDbLocation
                        gReportPath = App.Path & "\HiReports_Access\"
                    Else '-- Oracle
                        .Connect = "DSN=" + gODBCDSN + " ;UID=Scott;PWD=" + gODBCPwd + ";DBQ=" + gODBCHostStr
                        gReportPath = App.Path & "\HiReports_ORA\"
                    End If
                    '.Formulas(4) = "wFrToDt = 'From : ' & '" & FromDt & "' & '  To ' & '" & ToDt & "'"
                    .Formulas(4) = "wFrToDt = 'On Date : ' & '" & ToDt & " '"
                    '-- Selection Formula
                    X = "{qryAccountBalance.UserName}='" & gUserName & "'"
                    If txt(0) <> "" Then X = X + " and {qryAccountBalance.AGNAME}='" & txt(0) & "'"
                    .SelectionFormula = X
                    If gSelectedMenu = "Account Closing Balance List" Then
                        .ReportFileName = gReportPath & "RptLstAccCLBal.rpt"
                    Else
                        .ReportFileName = gReportPath & "RptLstAccCLBalGroupwise.rpt"
                    End If
            '============ List End
            Case "Account Ledger"  '====="Account Ledger"
            
               If txt(1) <> "" And cmdSPE(3).Value = True Then
                        Dim objCrystal As CRAXDRT.Application
                        Dim objReport As CRAXDRT.Report
                        Dim Tattach As String
                        Dim Mfile As String
                        Dim Mcp As String
                        Dim PtyMailId As String
                        Dim PtyMailIdCc As String
                        Dim PtyMailIdBcc As String
                        Dim MailSub As String
                        Dim MailStr As String
                        Dim wMillMailId As String
                        Dim j As Integer
                        Set objCrystal = New CRAXDRT.Application
                        Tattach = gReportPath & "rptAccountLedgerWapp.rpt"
                        Set objReport = objCrystal.OpenReport(Tattach, 1)
                        objReport.RecordSelectionFormula = "{QryAccountLedger.UserName}='" & gUserName & "' and {QryAccountLedger.ACNAME}='" & txt(1) & "'"

                        Dim crTab As CRAXDRT.DatabaseTable
                        For Each crTab In objReport.Database.Tables
                            crTab.Location = gDbLocation
                        Next
                        '---   X"
                        For j = 1 To objReport.FormulaFields.Count
                        Select Case objReport.FormulaFields(j).Name
                        Case "{@FirmNm}"
                                  objReport.FormulaFields(j).text = "Trim(" & Chr(39) & UCase(gCName) & Chr(39) & ")"
                        Case "{@wAdd1}"
                                   objReport.FormulaFields(j).text = "Trim(" & Chr(39) & gCAdd1 & Chr(39) & ")"
                        Case "{@wAdd2}"
                                   objReport.FormulaFields(j).text = "Trim(" & Chr(39) & gCAdd2 & Chr(39) & ")"
                        Case "{@wPhNo}"
                            objReport.FormulaFields(j).text = "Trim(" & Chr(39) & gCPhNo & Chr(39) & ")"
                        Case "{@wFrToDt}"
                           objReport.FormulaFields(j).text = "" & Chr(39) & " From : " & DTPicker1 & "   To " & DTPicker2 & "" & Chr(39) & ""
                        Case "{@WRptName}"
                           objReport.FormulaFields(j).text = "" & Chr(39) & " ACCOUNT LEDGER " & Chr(39) & ""
                       End Select
                       Next j
                       
                      Dim PtyMbl As String
                      Dim PtyName  As String
                      Dim FileNMWithPath As String
                      Dim FileNM As String
                      Dim aFile As String
                      FileNM = "AcLedger" & Format(Now, "ddMMyyyyhhmmss") & "" & ".pdf"
   '                   FileNM = "" & Format(Now, "ddMMyyyyhhmmss") & "" & ".pdf"
                      FileNMWithPath = gReportPath & "" & FileNM & ""
                      PtyMbl = GProcGetColumnValue("tblMastAccount", "AcName", txt(1), "S", "AcFaxNo", "S")
                      PtyName = GProcGetColumnValue("tblMastAccount", "AcName", txt(1), "S", "AcName", "S")
                      
                      ExportReportToPDF objReport, FileNMWithPath, "foo"
                       
                      'ExportReportToPDF objReport, Mid(Tattach, 1, Len(Tattach) - 4) & ".pdf", "foo"
                      Gdelay (10)
                      DoEvents
                      SentWhatsApp "PDF", FileNM, "", PtyMbl, PtyName
                      Kill Mid(Tattach, 1, Len(Tattach) - 4) & ".pdf"
                   Exit Sub
                Else
                    ''---- Prepare TmpRptTbl For Receipt Vs Invoice Narration
                     gCn.Execute "delete from tmpselection where UserName='" & gUserName & "'"
                     If gBackEndDB = gBackEndAccess Then
                         gCn.Execute "insert into tmpselection (dt1,Dt2,UserName) values ('" & FromDt & "','" & ToDt & "','" & gUserName & "')"
                     Else
                         gCn.Execute "insert into tmpselection (dt1,dt2,UserName) values (to_date('" & FromDt & "','dd/MM/yyyy'),to_Date('" & ToDt & "','dd/MM/yyyy'),'" & gUserName & "') "
                     End If
                     '--Set Loacation
                     If gBackEndDB = gBackEndAccess Then
                         .DataFiles(0) = gDbLocation
                         gReportPath = App.Path & "\HiReports_Access\"
                     Else '-- Oracle
                         .Connect = "DSN=" + gODBCDSN + " ;UID=Scott;PWD=" + gODBCPwd + ";DBQ=" + gODBCHostStr
                         gReportPath = App.Path & "\HiReports_ORA\"
                     End If
                     .Formulas(4) = "wFrToDt = 'From : ' & '" & FromDt & "' & '  To ' & '" & ToDt & "'"
                     '-- Selection Formula
                     X = "{QryAccountLedger.UserName}='" & gUserName & "'"
                     If txt(0) <> "" Then X = X + " and {QryAccountLedger.AGNAME}='" & txt(0) & "'"
                     If txt(1) <> "" Then X = X + " and {QryAccountLedger.ACNAME}='" & txt(1) & "'"
                     .SelectionFormula = X
                    '-- Report Name
                    Select Case sTp
                      Case "D"  '-Detail
                        If chkDaySum = 1 Then
                            .ReportFileName = gReportPath & "rptAccLedgerVdtwise.rpt" '-Dt wise
                        Else
''                            .Formulas(5) = "wRptName = 'ACCOUNT LEDGER'"
''                            .ReportFileName = gReportPath & "rptAccountLedger.rpt"
                                .Formulas(5) = "wRptName = 'ACCOUNT LEDGER'"
                                 'If ChkIntRateMonthly.Value = 1 Then
                                         X = "Select tblmastaccount.AcName,  tblvoucher.vdt as vdt,tblvoucher.vtype as Vtype,tblvoucher.vno as vno,Vctrno, tblAccO.acname as  AcOName,iif(vamt<0,abs(vamt),0) AS Debit,  iif(vamt>0,vamt,0)  AS Credit, tblvoucher.vbillno, tblvoucher.vbilldt, vnar1, vnar2, vnar3, vnar4,tblvoucher.vaccode," _
                                             & " AgName,2 as orderno,'" & gUserName & "',tblMastAccount.ACCITY,tblMastAccount.ACPan" _
                                             & " From tblvoucher, tblMastAccount , TblMastGroup, tblMastAccount as  tblAccO Where tblMastAccount.AGCODE=  TblMastGroup.AGCODE and tblvoucher.vacCode = tblMastAccount.AcCode  and tblvoucher.VAcOCode = tblAccO.AcCode  and tblvoucher.Vdt>= Cdate('" & FromDt & "') and tblvoucher.Vdt<=  Cdate('" & ToDt & "')"
                                        
                                         If txt(0) <> "" Then X = X + " and TblMastGroup.AgName ='" & txt(0) & "'"
                                         If txt(1) <> "" Then X = X + " and tblmastaccount.AcName='" & txt(1) & "'"
                                         X = X + " ORDER BY tblmastaccount.AcName, vdt, vtype, vno"
                                         
                                         gCn.BeginTrans
                                         gCn.Execute "Delete from tmpGenTbl where USERNAME='" & gUserName & "'"
                                         
                                           gCn.Execute "insert into tmpGenTbl (Nar1,Vdt,Vtype,Vno,VsrNo,Nar2,Amt1,Amt2,VBillNo,VBillDt,Nar3,Nar4,Nar5,Nar6,VACCODE,Nar7,Amt3,UserName,Nar8,Nar9) " & X & ""
                                                     
                                         X = "Select AcName,Cdate('" & FromDt & "'),'0P',0,0,'Opening Balance', iif(sum(vamt)<0,abs(sum(vamt)),0),  iif(sum(vamt)>0,sum(vamt),0),  '',null,'','','','',0,AgName,1,'" & gUserName & "',tblMastAccount.ACCITY,tblMastAccount.ACPan  From tblvoucher, tblmastaccount,TblMastGroup Where tblMastAccount.AGCODE= TblMastGroup.AGCODE and  tblvoucher.vacCode = tblMastAccount.AcCode And tblvoucher.Vdt < Cdate('" & FromDt & "')  "
                                         If txt(0) <> "" Then X = X + " and TblMastGroup.AgName ='" & txt(0) & "'"
                                         If txt(1) <> "" Then X = X + " and tblmastaccount.AcName='" & txt(1) & "'"
                                         X = X + " group by AcName,AgName,'" & gUserName & "',tblMastAccount.ACCITY,tblMastAccount.ACPan"
                                         gCn.Execute "insert into tmpGenTbl (Nar1,Vdt,Vtype,Vno,VsrNo,Nar2,Amt1,Amt2,VBillNo,VBillDt,Nar3,Nar4,Nar5,Nar6,VACCODE,Nar7,Amt3,UserName,Nar8,Nar9) " & X & ""
                                         
                                         gCn.CommitTrans
                                         
                                         X = "{tmpGenTbl.UserName}='" & gUserName & "'"
                                         .SelectionFormula = X
                                         .Formulas(5) = "wRptName = 'ACCOUNT LEDGER'"
                                         .ReportFileName = gReportPath & "rptAccountLedgerNew.rpt"


                        End If
                      Case "P"  '-Party
                         .ReportFileName = gReportPath & "rptAccLedgerpartywise.rpt"
                      Case "M"  '-Monthly
                        .ReportFileName = gReportPath & "rptAccLedgermonthly.rpt"
                    End Select
                    If ChkIsRefNo.Value = 1 Then
                        .ReportFileName = gReportPath & "rptAccLedgerdwiseBalance.rpt"
                    End If
                    
                    
                    
                    
                    
              End If
             Case "Mill Bill Account Ledger" '-"Mill Bill Account Ledger"
                 ''---- Prepare TmpRptTbl For Receipt Vs Invoice Narration
                 If txt(1) <> "" Then GProcPrepareTmpRptTblForRecInvNarration FromDt, ToDt, wAcCode, " RECVTYPE='MR'"
                '-- Prepare Mill Bill Ledger
                 If txt(1) = "" Then
                    GProcPrepareMillBillLedger FromDt, ToDt
                 Else
                    GProcPrepareMillBillLedger FromDt, ToDt, wAcCode
                 End If
                '--Set Loacation
                If gBackEndDB = gBackEndAccess Then
                    .DataFiles(0) = gDbLocation
                    gReportPath = App.Path & "\HiReports_Access\"
                Else '-- Oracle
                    .Connect = "DSN=" + gODBCDSN + " ;UID=Scott;PWD=" + gODBCPwd + ";DBQ=" + gODBCHostStr
                    gReportPath = App.Path & "\HiReports_ORA\"
                End If
                '-- Selection Formula
                .SelectionFormula = "{tmpGenTbl.UserName}='" & gUserName & "'"
                .Formulas(4) = "wFrToDt = 'From : ' & '" & FromDt & "' & '  To ' & '" & ToDt & "'"
                .ReportFileName = gReportPath & "rptAccLedgerMillBill.rpt"
             Case "Bank Reconciliation" '-- "Bank Reconciliation"
                '-insert in tmpGenTbl  Date Range,UserName
                 gCn.Execute "delete from tmpGenTbl where UserName='" & gUserName & "'"
                 If gBackEndDB = gBackEndAccess Then
                     gCn.Execute "insert into tmpGenTbl (dt1,Dt2,UserName) values ('" & FromDt & "','" & ToDt & "','" & gUserName & "')"
                 Else
                     gCn.Execute "insert into tmpGenTbl (dt1,dt2,UserName) values (to_date('" & FromDt & "','dd/MM/yyyy'),to_Date('" & ToDt & "','dd/MM/yyyy'),'" & gUserName & "') "
                 End If
                 '--Set Loacation
                 If gBackEndDB = gBackEndAccess Then
                     .DataFiles(0) = gDbLocation
                     gReportPath = App.Path & "\HiReports_Access\"
                 Else '-- Oracle
                     .Connect = "DSN=" + gODBCDSN + " ;UID=Scott;PWD=" + gODBCPwd + ";DBQ=" + gODBCHostStr
                     gReportPath = App.Path & "\HiReports_ORA\"
                 End If
                .Formulas(4) = "wFrToDt = 'From : ' & '" & FromDt & "' & '  To ' & '" & ToDt & "'"
                '-- Selection Formula
                X = "{qryBankReconcilLedger.UserName}='" & gUserName & "'"
                If txt(0) <> "" Then X = X + " and {qryBankReconcilLedger.AGNAME}='" & txt(0) & "'"
                If txt(1) <> "" Then X = X + " and {qryBankReconcilLedger.ACNAME}='" & txt(1) & "'"
                .SelectionFormula = X
                .ReportFileName = gReportPath & "rptAccBankReconcil.rpt"
             Case "Cash Book", "Bank Book", "Credit Cash Report"  '-===Cash Book", "Bank Book","Credit Cash Report"
                 '-insert Date Range,UserName
                       gCn.Execute "delete from tmpselection where UserName='" & gUserName & "'"
                        If gBackEndDB = gBackEndAccess Then
                            gCn.Execute "insert into tmpselection (dt1,Dt2,UserName) values ('" & FromDt & "','" & ToDt & "','" & gUserName & "')"
                        Else
                            gCn.Execute "insert into tmpselection (dt1,dt2,UserName) values (to_date('" & FromDt & "','dd/MM/yyyy'),to_Date('" & ToDt & "','dd/MM/yyyy'),'" & gUserName & "') "
                        End If
                        DoEvents

''                 gCn.Execute "delete from tmpGenTbl where UserName='" & gUserName & "'"
''                 If gBackEndDB = gBackEndAccess Then
''                     gCn.Execute "insert into tmpGenTbl (dt1,Dt2,UserName) values ('" & FromDt & "','" & ToDt & "','" & gUserName & "')"
''                 Else
''                     gCn.Execute "insert into tmpGenTbl (dt1,dt2,UserName) values (to_date('" & FromDt & "','dd/MM/yyyy'),to_Date('" & ToDt & "','dd/MM/yyyy'),'" & gUserName & "') "
''                 End If
                 '--Set Loacation
                 If gBackEndDB = gBackEndAccess Then
                     .DataFiles(0) = gDbLocation
                     gReportPath = App.Path & "\HiReports_Access\"
                 Else '-- Oracle
                     .Connect = "DSN=" + gODBCDSN + " ;UID=Scott;PWD=" + gODBCPwd + ";DBQ=" + gODBCHostStr
                     gReportPath = App.Path & "\HiReports_ORA\"
                 End If
                 .Formulas(4) = "wFrToDt = 'From : ' & '" & FromDt & "' & '  To ' & '" & ToDt & "'"
                 '-- Selection Formula
                 X = "{QryAccountLedger.UserName}='" & gUserName & "'"
                 If txt(0) <> "" Then X = X + " and {QryAccountLedger.AGNAME}='" & txt(0) & "'"
                 If txt(1) <> "" Then X = X + " and {QryAccountLedger.ACNAME}='" & txt(1) & "'"
                 If gSelectedMenu = "Bank Book" Then X = X + " and {QryAccountLedger.ACNAME}<>'Cash In Hand'"
                .SelectionFormula = X
                '-- Report Name
                If gSelectedMenu = "Credit Cash Report" Then '-- Credit Cash Report
                    .Formulas(5) = "wRptName = 'Credit Cash Report'"
                    .Formulas(6) = "wIsCrCashRpt = 'Y'"
                    .ReportFileName = gReportPath & "rptAccountLedger.rpt"
                Else
                    .ReportFileName = gReportPath & "rptAccountLedgerDateBal.rpt"
                    If gSelectedMenu = "Cash Book" Then '-- Cash Book
                        .Formulas(5) = "wRptName = 'Cash Book'"
                    Else '-- Bank Book'--- Date wise balance Ledger Report
                        .Formulas(5) = "wRptName = 'Bank Book'"
                    End If
                End If
                
            Case "Groupwise Account List"
                   If txt(0) <> "" Then
                      X = " {tblMastGroup.AGNAME}='" & txt(0) & "'" 'Group
                     .SelectionFormula = X
                   End If
                   If ChkIsCashBankGroup.Value = 1 Then
                        X = X + " and (isnull({tblMastAccount.AcCST}) or isnull({tblMastAccount.AcBST})) "
                   End If
                  .SelectionFormula = X
                   If ChkIsCashBankGroup.Value = 1 Then
                       .ReportFileName = gReportPath & "RptLstAccountGrpTin.rpt"
                   Else
                      .ReportFileName = gReportPath & "RptLstAccountGroupwise.rpt"
                   End If
            Case "Groupwise Account Ledger" '===="Groupwise Account Ledger"
                '-insert Date Range,UserName
                     gCn.Execute "delete from tmpGenTbl where UserName='" & gUserName & "'"
                     If gBackEndDB = gBackEndAccess Then
                         gCn.Execute "insert into tmpGenTbl (dt1,Dt2,UserName) values ('" & FromDt & "','" & ToDt & "','" & gUserName & "')"
                     Else
                         gCn.Execute "insert into tmpGenTbl (dt1,dt2,UserName) values (to_date('" & FromDt & "','dd/MM/yyyy'),to_Date('" & ToDt & "','dd/MM/yyyy'),'" & gUserName & "') "
                     End If
                     '--Set Loacation
                     If gBackEndDB = gBackEndAccess Then
                         .DataFiles(0) = gDbLocation
                         gReportPath = App.Path & "\HiReports_Access\"
                     Else '-- Oracle
                         .Connect = "DSN=" + gODBCDSN + " ;UID=Scott;PWD=" + gODBCPwd + ";DBQ=" + gODBCHostStr
                         gReportPath = App.Path & "\HiReports_ORA\"
                     End If
                     .Formulas(4) = "wFrToDt = 'From : ' & '" & FromDt & "' & '  To ' & '" & ToDt & "'"
                     '-- Selection Formula AcName ,AgName,Bank Book ,cash Book
                     X = "{qryAccGroupWiseLedger.UserName}='" & gUserName & "'"
                     If txt(1) <> "" Then X = X + " and {qryAccGroupWiseLedger.ACNAME}='" & txt(1) & "'" 'A/c
                     If txt(0) <> "" Then X = X + " and {qryAccGroupWiseLedger.AGNAME}='" & txt(0) & "'" 'Group
                     If ChkIsCashBankGroup = 0 Then X = X + " and {qryAccGroupWiseLedger.AGCode}<>90019"  'No Cash Bank Group
                     .SelectionFormula = X
                     If chkDaySum = 1 Then
                        .ReportFileName = gReportPath & "rptAccGroupLedgerNaNp.rpt"
                     ElseIf ChkIsRefNo = 1 Then '-Ref No in Ledger
                        .ReportFileName = gReportPath & "rptAccGroupLedger.rpt"
                     ElseIf ChkIntRateMonthly.Value = 1 Then
                        .ReportFileName = gReportPath & "rptAccGroupBrkLst.rpt"
                     Else
                        .ReportFileName = gReportPath & "rptAccGroupLedgerNoRefNo.rpt"
                     End If
            Case "Account Confirmation" '-"Account Confirmation"
                  '-insert Date Range,UserName
                     gCn.Execute "delete from tmpGenTbl where UserName='" & gUserName & "'"
                     If gBackEndDB = gBackEndAccess Then
                         gCn.Execute "insert into tmpGenTbl (dt1,Dt2,UserName) values ('" & FromDt & "','" & ToDt & "','" & gUserName & "')"
                     Else
                         gCn.Execute "insert into tmpGenTbl (dt1,dt2,UserName) values (to_date('" & FromDt & "','dd/MM/yyyy'),to_Date('" & ToDt & "','dd/MM/yyyy'),'" & gUserName & "') "
                     End If
                     '--Set Loacation
                     If gBackEndDB = gBackEndAccess Then
                         .DataFiles(0) = gDbLocation
                         gReportPath = App.Path & "\HiReports_Access\"
                     Else '-- Oracle
                         .Connect = "DSN=" + gODBCDSN + " ;UID=Scott;PWD=" + gODBCPwd + ";DBQ=" + gODBCHostStr
                         gReportPath = App.Path & "\HiReports_ORA\"
                     End If
                     .Formulas(4) = "wFrToDt = 'From ' & '" & FromDt & "' & '  To ' & '" & ToDt & "'"
                     .Formulas(5) = "PANNo= '" & gCPAN & "'"
                     '-- Selection Formula
                     X = "{QryAccountLedger.UserName}='" & gUserName & "'"
                     If txt(1) <> "" Then X = X + " and {QryAccountLedger.ACNAME}='" & txt(1) & "'" 'A/c
                     If txt(0) <> "" Then X = X + " and {QryAccountLedger.AGNAME}='" & txt(0) & "'" 'Group
                     .SelectionFormula = X
                     '-- Report Name
                    .ReportFileName = gReportPath & "rptAccConfirmation.rpt"
             Case "Sarafi Interest Report", "Sarafi Interest Report ANK"  '=== "Sarafi Interest Report"
                    '-- Credit/Debit Balance on To Date
                    gCn.Execute "delete from tmpselection"
                    'where UserName='" & gUserName & "'"
                    If gBackEndDB = gBackEndAccess Then
                        gCn.Execute "insert into tmpselection (dt1,Dt2,UserName) values ('" & gCYSDate & "','" & ToDt & "','" & gUserName & "')"
                    Else
                        gCn.Execute "insert into tmpselection (dt1,dt2,UserName) values (to_date('" & gCYSDate & "','dd/MM/yyyy'),to_Date('" & ToDt & "','dd/MM/yyyy'),'" & gUserName & "') "
                    End If
                 '-Prepare For Ledger (insert in tmpGenTbl  Date Range,Sale Side Days,UserName)
                 gCn.Execute "delete from tmpGenTbl " 'where UserName='" & gUserName & "'"
                 If gBackEndDB = gBackEndAccess Then
                     gCn.Execute "insert into tmpGenTbl (dt1,Dt2,Amt1,UserName) values ('" & FromDt & "','" & ToDt & "'," & txttmp(1) & ",'" & gUserName & "')"
                 Else
                     gCn.Execute "insert into tmpGenTbl (dt1,dt2,Amt1,UserName) values (to_date('" & FromDt & "','dd/MM/yyyy'),to_Date('" & ToDt & "','dd/MM/yyyy')," & txttmp(1) & ",'" & gUserName & "') "
                 End If
                 .Formulas(4) = "wFrToDt = 'From : ' & '" & FromDt & "' & '  To ' & '" & ToDt & "'"
                 '-- Is Interest Rate Monthly
                 If ChkIntRateMonthly = 1 Then
                   .Formulas(6) = "IsIntRateMonthly ='1'"
                 Else
                   .Formulas(6) = "IsIntRateMonthly ='0'"
                 End If
                 .Formulas(7) = "IntRate ='" & txttmp(0) & "'" '- Int Rate
                 .Formulas(8) = "TotDays ='" & txttmp(4) & "'" '- Total Days
                 .Formulas(9) = "SideSADay ='" & txttmp(1) & "'" '- Side SA Days
                 .Formulas(10) = "LessPayDay ='" & txttmp(2) & "'" '- Less Pay
                 .Formulas(11) = "LessRecDay ='" & txttmp(3) & "'" '- Less Rec
                 .Formulas(12) = "ToDate ='" & Format(ToDt, "yyyy,mm,dd") & "'"  '- To Date
                 .Formulas(14) = "TdsLimit =" & Val(txttmp(6)) & ""
                 .Formulas(15) = "TdsRate =" & Val(txttmp(5)) & ""
                 '--Set Loacation
                 If gBackEndDB = gBackEndAccess Then
                     .DataFiles(0) = gDbLocation
                     gReportPath = App.Path & "\HiReports_Access\"
                 Else '-- Oracle
                     .Connect = "DSN=" + gODBCDSN + " ;UID=Scott;PWD=" + gODBCPwd + ";DBQ=" + gODBCHostStr
                     gReportPath = App.Path & "\HiReports_ORA\"
                 End If
                 '-- Selection Formula
                 gCn.BeginTrans
                 X = "{qryInterestLedger.UserName}='" & gUserName & "'"
                 gCn.CommitTrans
                 If txt(0) <> "" Then X = X + " and {qryInterestLedger.AGNAME}='" & txt(0) & "'"
                 If txt(1) <> "" Then X = X + " and {qryInterestLedger.ACNAME}='" & txt(1) & "'"
                  '- Report Name
                                  
                 If chkDaySum.Value = 1 Then
                    .SelectionFormula = X
                    .ReportFileName = gReportPath & "rptInterestRpt.rpt"
                 Else
                    If ChkSummary = 0 Then
                       X = X + " and ({@IntPayAmt}>0 or {@IntRecAmt}>0)"
                       .SelectionFormula = X
                       If gSelectedMenu = "Sarafi Interest Report ANK" Then
                          .ReportFileName = gReportPath & "rptInterestRptANK.rpt"
                       Else
                          .ReportFileName = gReportPath & "rptInterestRptSarafi.rpt"
                       End If
                    Else '-  Summary
                       X = X + " and (({qryInterestLedger.Vtype}<>'CL'  and ({@IntPayAmt}>0 or {@IntRecAmt}>0)) or {qryInterestLedger.Vtype}='CL')"
                      .SelectionFormula = X
                       If gSelectedMenu = "Sarafi Interest Report ANK" Then
                          .ReportFileName = gReportPath & "rptInterestRptANKSumm.rpt"
                       Else
                         .ReportFileName = gReportPath & "rptInterestRptSarafiSumm.rpt"
                       End If
                       
                       
                    End If
                 End If
            Case "R25":
                .ReportFileName = App.Path & "\Reports\rptAcCon.rpt"
            Case "R2M":
                .Formulas(7) = "wRepoHead='MULTIPLE ACCOUNT LEDGER'"
                .ReportFileName = App.Path & "\Reports\rptSingleAcLed.rpt"
        End Select
        .Action = 1
    End With
    Screen.MousePointer = vbDefault
ElseIf optReportOn(2) = True Then '-Grid
    Pty_Ldr1 ("F")
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
Me.Top = 700
Me.Left = Screen.Width / 2 - Me.Width / 2
 '-- Activate & Set  Controls
GProcActivateControls Me, False
DTPicker1.Enabled = True
DTPicker2.Enabled = True
lstParty.Visible = False
For i = 0 To 1 '-- Report On
    optReportOn(i).Enabled = True
Next i
'-- Set Values
DTPicker1.Value = Format(gCYSDate, "dd/MM/yyyy")
DTPicker2.Value = Format(gCYEDate, "dd/MM/yyyy")
For i = 0 To 4
    txttmp(i) = 0
Next i
'-Days of year
txttmp(4) = DateDiff("d", gCYSDate, gCYEDate) + 1
CheckForNumber
'------------------------------------
mScreen = True
Me.Caption = gSelectedMenu
Select Case gSelectedMenu
    Case "TCS Report"
         txt(1).Enabled = True
         txt(0).Visible = False
         txt(2).Visible = False
         Label5.Visible = False
         Label7.Visible = False
         Label1(7).Visible = False
         Label1(8).Visible = False
         optSelType(0).Enabled = True
         optSelType(1).Enabled = True
         optSelType(0).Caption = "Detail"
         optSelType(1).Caption = "Summery"
         optSelType(2).Visible = False
         optSelType(3).Visible = False
         Label4.Visible = False
         Label6.Visible = False
         chkDaySum.Visible = False
         ChkIsRefNo.Visible = True
         ChkIsRefNo.Enabled = True
         ChkIsRefNo.Caption = "On Total Receipt"
         ChkIsCashBankGroup.Visible = True
         ChkIsCashBankGroup.Enabled = True
         ChkIsCashBankGroup.Caption = "On Rec. Without Interest"
         ChkIsCashBankGroup.Value = 0
         Label8.Visible = False
         Label1(2).Caption = "TCS % :"
         txttmp(3).Enabled = True
         txttmp(0).Enabled = True
         txttmp(4).Visible = False
         txttmp(5).Visible = False
         txttmp(6).Visible = False
         Label1(5).Visible = False
         Label1(4).Visible = False
         Label1(1).Visible = False
         Label1(6).Visible = False
         txttmp(1).Visible = False
         txttmp(2).Visible = False
         ChkIntRateMonthly.Visible = False
         ChkSummary.Visible = False
         Frame1.Visible = False
         Label1(3).Caption = "TCS Limit :"
         txttmp(3).text = 5000000
         txttmp(0).text = 0.1
    Case "Account List"
         DTPicker1.Enabled = False
         DTPicker2.Enabled = False
        txt(0).Enabled = False
         txt(1).Enabled = True
        txt(2).Enabled = False
        Label3.Caption = "Broker"
    Case "Selected Group Accounts List", "Account Groupwise Opening Balance List" '== "Selected Group Accounts List","Account Groupwise Opening Balance List"
        DTPicker1.Enabled = False
        DTPicker2.Enabled = False
        txt(0).Enabled = True '- A/c Group
        optReportOn(0).Enabled = False '-Dos
        ChkIsCashBankGroup.Enabled = True
        ChkIsCashBankGroup.Caption = "Email And Mobile"
        ChkIsCashBankGroup.Value = 0
        ChkIsRefNo.Enabled = True
        ChkIsRefNo.Caption = "Pan/Tin List"
        ChkIsRefNo.Value = 0

    Case "Millwise Count List", "Count Brokerage List" '== "Millwise Count List","Count Brokerage List"
        DTPicker1.Enabled = False
        DTPicker2.Enabled = False
        txt(2).Enabled = True '- Mill
        optReportOn(0).Enabled = False '-Dos
    Case "Account Closing Balance List" '== "Account Closing Balance List"
        DTPicker1.Enabled = False
        optReportOn(0).Enabled = False '-Dos
        optReportOn(2).Enabled = False '-Grid
    Case "Account Groupwise Closing Balance List"  '=== "Account Closing Balance List"
        DTPicker1.Enabled = False
        txt(0).Enabled = True '- A/c Group
        optReportOn(0).Enabled = False '-Dos
    Case "Account Ledger" '-"Account Ledger"
        txt(1).Enabled = True '-A/c
       ' lstParty.Enabled = True
        chkDaySum.Enabled = True '-Summ
        ChkIsRefNo.Value = 0
        ChkIsRefNo.Enabled = True
        ChkIsRefNo.Caption = "Daywise Balance"
        cmdSPE(3).Enabled = True
        cmdSPE(3).Visible = True
        For i = 0 To 3 '-- Report format
            optSelType(i).Enabled = True
        Next i
        
    Case "Groupwise Account Ledger" '== "Groupwise Account Ledger"
        txt(0).Enabled = True '-A/c Group
        txt(1).Enabled = True '-A/c
        ChkIsCashBankGroup.Enabled = True '-CAsh ,Bank Group
        ChkIsRefNo.Enabled = True '-Ref No
        ChkIntRateMonthly.Enabled = True
        ChkIntRateMonthly.Caption = "Pan List"
        ChkIntRateMonthly.Value = 0
        chkDaySum.Enabled = True
        chkDaySum.Caption = "Accountwise Print ?"
    Case "Groupwise Account List"
        txt(0).Enabled = True '-A/c Group
        ChkIsCashBankGroup.Caption = "Without Tin List"
        ChkIsCashBankGroup.Enabled = True
        ChkIsCashBankGroup.Value = 0

    Case "Account Confirmation" '== "Account Confirmation"
        txt(1).Enabled = True '-A/c
        txt(0).Enabled = True '-A/c Group
    Case "Mill Bill Account Ledger" '- "Mill Bill Account Ledger"
        txt(1).Enabled = True '-A/c
    Case "Bank Reconciliation" '-- "Bank Reconciliation"
        txt(0) = GProcGetColumnValue("tblMastGroup", "AgCode", 90019, "N", "AgName", "S") '-Bank Group
        txt(1).Enabled = True '-A/c
    Case "Cash Book", "Credit Cash Report" '== "Cash Book", '== "Credit Cash Report"
        txt(1) = "Cash In Hand" '-Cash A/c
    Case "Bank Book" '== Bank Book"
        txt(1).Enabled = True '-A/c
        txt(0) = GProcGetColumnValue("tblMastGroup", "AgCode", 90019, "N", "AgName", "S") '-Bank Group
    Case "Sarafi Interest Report", "Sarafi Interest Report ANK"
 '-"Sarafi Interest Report"
        optReportOn(0).Enabled = False '-Dos
        txt(0).Enabled = True '-- A/c Group
        txt(1).Enabled = True '--A/c
        ChkSummary.Enabled = True '-- Interest  Setting Ctrls
        ChkIntRateMonthly.Enabled = True
        For i = 0 To 6
            txttmp(i).Enabled = True
        Next i
        If gSelectedMenu = "Sarafi Interest Report" Then
           chkDaySum.Enabled = True
           chkDaySum.Caption = "CC A/c OD Int. Report"
        End If
    'Case "R22"
    '    Label3.Caption = "Group Name :"
    '    Me.Caption = "Single Group Ledger"
    '    txt(0).Enabled = True '-Group
    'Case "R23"
    '    Me.Caption = "All Account Ledger"
    '    txt(1).Visible = False
    '    Label3.Visible = False
'    Case "R25"
'        Me.Caption = "Account Confirmation"
    'Case "R26"
    '    Me.Caption = "Account Ledger (Monthly)"
    'Case "R27"
    '    Me.Caption = "Account Ledger (Partywise)"
'    Case "R2M"
'        Me.Caption = "Multi Account Ledger"
'        lstParty.Enabled = True
'        txt(1).Visible = False
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
Private Sub Pty_Ldr1(wRpTp As String)
Dim tdt As Date, currdt As Date, CrSide(2000) As String, DrSide(2000) As String, CrCnt As Integer, DrCnt As Integer
Dim t1 As Double, t2 As Double, t3 As Double, wnp As Double, T11 As Double
Dim wTp As String, X As String, X1 As String, ttp As String, wAcCode1 As Long, wwaname As String, wwagcode As Long
Dim rst1 As Recordset, r_s As String, wTTrno As String, wTContP As Boolean
Dim rst2 As Recordset, rstTran33 As Recordset, j As Integer
Dim wHdPrn As Boolean, wNamePrn As Boolean
Dim wRecLen As Long, z1 As String, rstIt As Recordset, rstItP As Recordset
Dim wDPay As Double, wDRec As Double, mbkMk As Variant
wDotLine = String(96, "-")
PgCnt = 1
CrCnt = 0
DrCnt = 0
If (gSelectedMenu <> "R23") And (gSelectedMenu <> "R2M") Then
    If (Len(wAcCode) = 0) Then
        MsgBox "Select Party Name", vbCritical + vbOKOnly, "Select Party Name"
        txt_GotFocus (1)
        Exit Sub
    End If
End If

If mScreen = True Then
    mP12 = ""
    mPc = ""
Else
    mP12 = Chr(27) + Chr(33) + Chr(1)
    mPc = Chr(27) + Chr(15)
End If

Select Case gSelectedMenu
Case "Account Ledger", "R25", "R26", "R27"
    r_s = "and vACCODE = " & wAcCode
Case "R22"
    r_s = "and AGCODE = " & wAcCode
Case "R23"
    r_s = ""
   ' gCn.Execute "Update ACCT set PGPR =0 where ACCODE ='" & wAcCode & "' "
Case "R2M":
    z1 = ""
    For i = 0 To (lstParty.ListCount - 1)
        If i < (lstParty.ListCount - 1) Then
            z1 = z1 + " (vACCODE = " & wPartyCd(i) & ") or "
        ElseIf i = (lstParty.ListCount - 1) Then
            z1 = z1 + " (vACCODE = " & wPartyCd(i) & ") "
        End If
    Next i
    r_s = "and (" & z1 & ")"
End Select
DoEvents


'======================================= PRINTING START ========================
wHdPrn = True
If gBackEndDB = gBackEndORA Then
T11 = GProcRstOpen(rstTran33, "Select * from qryPartyLedgerSingle where vdt <= to_date('" & DTPicker2 & "','dd/MM/yyyy') " & r_s & " order by vaccode,vdt,vtype,vno,Vctrno", "R", gCn)
Else
T11 = GProcRstOpen(rstTran33, "Select * from qryPartyLedgerSingle where format(vdt,'yyyy/MM/dd') <= '" & Format(DTPicker2, "yyyy/MM/dd") & "' " & r_s & " order by vaccode,vdt,vtype,vno,Vctrno", "R", gCn)
End If
If (T11 = 0) Then
   rstTran33.Close
   MsgBox "No Report available", vbCritical + vbOKOnly, "No Report available"
   Exit Sub
End If
Dim mOpBal As Double, TmpVno As Long, TmpVType As String, TmpVamt As Double
mOpBal = 0
prgBar.Value = 0
prgBar.Max = rstTran33.RecordCount
prgBar.Visible = True
If (wRpTp = "D") Or (wRpTp = "T") Or (wRpTp = "M") Or (wRpTp = "P") Then
    Dim wChgDate As Date
    wDPay = 0
    wDRec = 0
    wChgDate = DateAdd("d", -1, gCYSDate)
    Open "c:\rpt1.txt" For Output As #1
       With rstTran33
            .MoveFirst
            Do While Not .EOF
               wAcCode = !vacCode
                If (wRpTp = "T") Or (wRpTp = "M") Or (wRpTp = "P") Then
                    For i = 1 To 2000
                        CrSide(i) = ""
                        DrSide(i) = ""
                    Next i
                End If
               t1 = 0
               t2 = 0
               wNamePrn = True
Loop2:
               Do While Not .EOF
                  If (wAcCode <> !vacCode) Then Exit Do
                  If (!Vdt < DTPicker1) Then
                     mOpBal = mOpBal + !Credit - !Debit
                     .MoveNext
                     GoTo Loop2
                  End If
                  Select Case wRpTp
                    Case "M":
                        If (!Debit > 0) Then
                           DrSide(Month(!Vdt)) = GProcNoToStr(!Debit + Val(DrSide(Month(!Vdt))), 12, 2)
                        Else
                           t2 = t2 + T11
                           CrSide(Month(!Vdt)) = GProcNoToStr(!Credit + Val(CrSide(Month(!Vdt))), 12, 2)
                        End If
                    Case "P":
                        For j = 1 To DrCnt
                            If Mid(DrSide(j), 25) = !AcName Then Exit For
                        Next j
                        If j > DrCnt Then
                            DrCnt = DrCnt + 1
                            If (!Debit > 0) Then
                               DrSide(j) = GProcNoToStr(!Debit, 12, 2) + Space(12) + !AcName
                            Else
                               DrSide(j) = Space(12) + GProcNoToStr(!Credit, 12, 2) + !AcName
                            End If
                        Else
                            If (!Debit > 0) Then
                               DrSide(j) = GProcNoToStr(!Debit + Val(Left(DrSide(j), 12)), 12, 2) + Mid(DrSide(j), 13)
                            Else
                               DrSide(j) = Left(DrSide(j), 12) + GProcNoToStr(!Credit + Val(Mid(DrSide(j), 13, 12)), 12, 2) + Mid(DrSide(j), 25)
                            End If
                        End If
                    Case "T":
                        Select Case i
                        Case (i < 3):
                           X1 = Left(!AcName + Space(19), 19) + " No." + Left(!VbillNo, 7)
                        Case (i < 5):
                           X1 = Left(!AcName + Space(19), 19) + " No." + Left(!VbillNo, 7)
                        Case Else
                           X1 = Left(!AcName + Space(30), 30)
                        End Select
                        If (!Debit > 0) Then
                           DrCnt = DrCnt + 1
                           DrSide(DrCnt) = GProcNoToStr(!Debit, 12, 2) + "  " + Format(!Vdt, "dd/MM/yyyy") + "  " + Left(!Vtype, 2) + "  " + X1
                        Else
                           CrCnt = CrCnt + 1
                           CrSide(CrCnt) = GProcNoToStr(!Credit, 12, 2) + "  " + Format(!Vdt, "dd/MM/yyyy") + "  " + Left(!Vtype, 2) + "  " + X1
                        End If
                           If (Len(Trim(!VNar1)) > 0) Then
                              If (!Debit > 0) Then
                                 DrCnt = DrCnt + 1
                                 DrSide(DrCnt) = Space(30) + Left(!VNar1, 30)
                              Else
                                 CrCnt = CrCnt + 1
                                 CrSide(CrCnt) = Space(30) + Left(!VNar1, 30)
                              End If
                           End If
                           If (Len(Trim(!VNar2)) > 0) Then
                              If (!Debit > 0) Then
                                 DrCnt = DrCnt + 1
                                 DrSide(DrCnt) = Space(30) + Left(!VNar2, 30)
                              Else
                                 CrCnt = CrCnt + 1
                                 CrSide(CrCnt) = Space(30) + Left(!VNar2, 30)
                              End If
                           End If
                           If (Len(Trim(!VNar3)) > 0) Then
                              If (!Debit > 0) Then
                                 DrCnt = DrCnt + 1
                                 DrSide(DrCnt) = Space(30) + Left(!VNar3, 30)
                              Else
                                 CrCnt = CrCnt + 1
                                 CrSide(CrCnt) = Space(30) + Left(!VNar3, 30)
                              End If
                           End If
                           If (Len(Trim(!VNAR4)) > 0) Then
                              If (!Debit > 0) Then
                                 DrCnt = DrCnt + 1
                                 DrSide(DrCnt) = Space(30) + Left(!VNAR4, 30)
                              Else
                                 CrCnt = CrCnt + 1
                                 CrSide(CrCnt) = Space(30) + Left(!VNAR4, 30)
                              End If
                           End If
                  Case "D":
                        If wHdPrn Then
                           wHdPrn = False
                           DosHdg
                        End If
                        DoEvents
                        If wNamePrn Then
                           wNamePrn = False
                           wwaname = GProcGetColumnValue("tblmastaccount", "accode", CStr(wAcCode), "N", "acname", "S")
                           Print #1, Space(40 - (Len(wwaname) / 2)) + wwaname
                           Print #1, Space(40 - (Len(wwaname) / 2)) + String(Len(wwaname), "-") + mP12
                          ' If gSelectedmenu = "R23" Then Cn.Execute "Update ACCT set PGPR =" & (PgCnt - 1) & " where ACCODE ='" & wAcCode & "' and PGPR = 0 "
                           LnCnt = LnCnt + 2
                        End If
                        If (!Vdt <> wChgDate) Then
                           If ((wDPay > 0) Or (wDRec > 0)) And (chkDaySum.Value = 1) Then
                                If (LnCnt > 65) Then
                                    If mScreen = True Then
                                        Print #1, wDotLine
                                    Else
                                        Print #1, wDotLine + Chr(27) + Chr(33) + Chr(0)
                                    End If
                                    wHdPrn = True
                                    LnCnt = 0
                                End If
                                Print #1, Space(58) + String(21, "-")
                                If mScreen = True Then
                                    Print #1, Space(10) + "Day Totals : " + Space(35) + GProcNoToStr(wDRec, 11, 2) + GProcNoToStr(wDPay, 10, 2)
                                Else
                                    Print #1, Chr(27) + Chr(69) & Space(10) + "Day Totals : " + Space(35) + GProcNoToStr(wDRec, 11, 2) + GProcNoToStr(wDPay, 10, 2) + Chr(27) + Chr(70)
                                End If
                                Print #1, ""
                                LnCnt = LnCnt + 3
                                wDPay = 0
                                wDRec = 0
                            End If
                           X = Format(!Vdt, "dd/MM/yy") + "  "
                        Else
                           X = Space(10)
                        End If
                        wChgDate = !Vdt
                        If (mOpBal <> 0) Then
                            X1 = Space(10) + "             " + Left("Opening Balance" + Space(32), 32) + "  "
                            If (mOpBal < 0) Then
                               t1 = t1 + Abs(mOpBal)
                               wDRec = wDRec + Abs(mOpBal)
                               X1 = X1 + GProcNoToStr(Abs(mOpBal), 11, 2) + Space(10)
                            Else
                               t2 = t2 + Abs(mOpBal)
                               wDPay = wDPay + Abs(mOpBal)
                               X1 = X1 + Space(10) + GProcNoToStr(Abs(mOpBal), 11, 2)
                            End If
                            X1 = X1 + "  " + GProcNoToStr(Abs(Round((t1 - t2), 2)), 12, 2)
                            If (t1 > t2) Then
                               X1 = X1 + " Dr"
                            Else
                               X1 = X1 + " Cr"
                            End If
                            Print #1, mP12 + X1
                            LnCnt = LnCnt + 1
                            mOpBal = 0
                        End If
                        If mScreen = True Then
                           X = X + Left(Trim(!Vtype) + Space(2), 2) + "  " + Left(Trim(!Vno & vbNullString) + Space(7), 7) + "  " + Left(!AcName + Space(32), 32) + "  "
                        Else
                           X = X + Left(Trim(!Vtype) + Space(2), 2) + "  " + Left(Trim(!VbillNo & vbNullString) + Space(7), 7) + "  " + Left(!AcName + Space(32), 32) + "  "
                        End If
                        TmpVType = !Vtype
                        TmpVno = !Vno
                        TmpVamt = 0
                        Do While Not .EOF
                            If (TmpVType <> !Vtype) Or (TmpVno <> !Vno) Then Exit Do
                            If (!Debit > 0) Then
                              TmpVamt = TmpVamt - !Debit
                            Else
                              TmpVamt = TmpVamt + !Credit
                            End If
                            .MoveNext
                        Loop
                        .MovePrevious
                            If (TmpVamt < 0) Then
                               t1 = t1 + Abs(TmpVamt)
                               wDRec = wDRec + Abs(TmpVamt)
                               X = X + GProcNoToStr(Abs(TmpVamt), 11, 2) + Space(10)
                            Else
                               t2 = t2 + TmpVamt
                               wDPay = wDPay + TmpVamt
                               X = X + Space(10) + GProcNoToStr(TmpVamt, 11, 2)
                            End If
                        X1 = GProcNoToStr(Abs(Round((t1 - t2), 2)), 12, 2)
                        If (t1 > t2) Then
                           X1 = X + "  " + X1 + " Dr"
                        Else
                           X1 = X + "  " + X1 + " Cr"
                        End If
                        Print #1, mP12 + X1
                        LnCnt = LnCnt + 1
                        If (Len(Trim(!VNar1)) > 0) Then
                           Print #1, Space(25) & mPc & Trim(!VNar1) & mP12
                           LnCnt = LnCnt + 1
                        End If
                        If (Len(Trim(!VNar2)) > 0) Then
                           Print #1, Space(25) & mPc & Trim(!VNar2) & mP12
                           LnCnt = LnCnt + 1
                        End If
                        If (Len(Trim(!VNar3)) > 0) Then
                           Print #1, Space(25) & mPc & Trim(!VNar3) & mP12
                           LnCnt = LnCnt + 1
                        End If
                        If (Len(Trim(!VNAR4)) > 0) Then
                           Print #1, Space(25) & mPc & Trim(!VNAR4) & mP12
                           LnCnt = LnCnt + 1
                        End If
                        If (LnCnt > 60) Then
                            If mScreen = True Then
                                Print #1, wDotLine
                            Else
                                Print #1, wDotLine + Chr(27) + Chr(33) + Chr(0) + Chr(12)
                            End If
                            wHdPrn = True
                            LnCnt = 0
                        End If
                   End Select
                  .MoveNext
               Loop
                If (wRpTp = "D") And ((wDPay > 0) Or (wDRec > 0)) And (chkDaySum.Value = 1) Then
                     Print #1, Space(58) + String(21, "-")
                     If mScreen = True Then
                         Print #1, Space(10) + "Day Totals : " + Space(35) + GProcNoToStr(wDRec, 11, 2) + GProcNoToStr(wDPay, 10, 2)
                     Else
                         Print #1, Chr(27) + Chr(69) & Space(10) + "Day Totals : " + Space(35) + GProcNoToStr(wDRec, 11, 2) + GProcNoToStr(wDPay, 10, 2) + Chr(27) + Chr(70)
                     End If
                     Print #1, ""
                     LnCnt = LnCnt + 3
                        If (LnCnt > 60) Then
                            If mScreen = True Then
                                Print #1, wDotLine
                            Else
                                Print #1, wDotLine + Chr(27) + Chr(33) + Chr(0) + Chr(12)
                            End If
                            wHdPrn = True
                            LnCnt = 0
                        End If
                 End If
                Select Case wRpTp
                Case "D":
                    If ((t1 > 0) Or (t2 > 0)) Then
                        If wHdPrn Then
                            wHdPrn = False
                            DosHdg
                        End If
                        If wNamePrn Then
                            wNamePrn = False
                           wwaname = GProcGetColumnValue("tblmastaccount", "accode", CStr(wAcCode), "N", "acname", "S", , gCn)
                            Print #1, Space(40 - (Len(wwaname) / 2)) + wwaname
                            Print #1, Space(40 - (Len(wwaname) / 2)) + String(Len(wwaname), "-") + mP12
                            LnCnt = LnCnt + 2
                        End If
                        X1 = GProcNoToStr(Abs(Round((t1 - t2), 2)), 12, 2)
                        If (t1 > t2) Then
                           X1 = X1 + " Dr"
                        Else
                           X1 = X1 + " Cr"
                        End If
                        Print #1, Space(54) + String(42, "-")
                        Print #1, Space(54) + GProcNoToStr(t1, 12, 2) + " " + GProcNoToStr(t2, 12, 2) + " " + X1
                        Print #1, Space(54) + String(42, "-") + Chr(27) + Chr(33) + Chr(0)
                        LnCnt = LnCnt + 4
                        If (LnCnt > 60) Then
                            If mScreen = True Then
                                Print #1, wDotLine
                            Else
                                Print #1, wDotLine + Chr(27) + Chr(33) + Chr(0) + Chr(12)
                            End If
                            wHdPrn = True
                            LnCnt = 0
                        End If
                    End If
                Case "T":      'T-Format Party Ledger
                    wDotLine = String(126, "-")
                    If (CrCnt > DrCnt) Then
                       wnp = CrCnt
                    Else
                       wnp = DrCnt
                    End If
                    t1 = 0
                    t2 = 0
                    For t3 = 1 To wnp
                        If wHdPrn Then
                           wHdPrn = False
                           If mScreen = True Then
                               Print #1, gCName
                               Print #1, "Party Ledger For Period From " & DTPicker1.Value & "  To  " & DTPicker2.Value & "           Page No. " + Str(PgCnt)
                               Print #1, wDotLine
                           Else
                               Print #1, Chr(18) + Chr(27) + Chr(14) & gCName & Chr(27) + Chr(18)
                               Print #1, Chr(27) + Chr(69) & "Party Ledger " & Chr(27) + Chr(70); " From " & DTPicker1.Value & "  To  " & DTPicker2.Value & "           Page No. " + Str(PgCnt)
                               Print #1, mP12 & wDotLine
                           End If
                           Print #1, "Credit Amt.   Date     Tp   Description                           Debit Amt.       Date     Tp   Description"
                           Print #1, wDotLine
                           LnCnt = 5
                           PgCnt = PgCnt + 1
                        End If
                        If wNamePrn Then
                           wNamePrn = False
                           wwaname = GProcGetColumnValue("tblmastaccount", "accode", CStr(wAcCode), "N", "acname", "S", , gCn)
                           Print #1, Space(60 - (Len(wwaname) / 2)) + wwaname
                           Print #1, Space(60 - (Len(wwaname) / 2)) + String(Len(wwaname), "-")
                           LnCnt = LnCnt + 2
                        End If
                        If (t3 <= CrCnt) Then
                           X = Left(CrSide(t3) + Space(63), 63)
                           t1 = t1 + Val(Left(CrSide(t3), 12))
                        Else
                           X = Space(63)
                        End If
                        If (t3 <= DrCnt) Then
                           X = X + DrSide(t3)
                           t2 = t2 + Val(Left(DrSide(t3), 12))
                        End If
                        Print #1, X
                        LnCnt = LnCnt + 1
                        If (LnCnt > 65) And (wRpTp = "D") Then
                            If mScreen = True Then
                                Print #1, wDotLine
                            Else
                                Print #1, wDotLine & Chr(12)
                            End If
                            wHdPrn = True
                            LnCnt = 0
                        End If
                    Next t3
                    If (t1 <> t2) Then
                       X = GProcNoToStr(Abs(t2 - t1), 12, 2)
                       If (t2 > t1) Then
                          t1 = t2
                          X = X + "  " + Format(DTPicker2.Value, "dd/MM/yyyy") + Space(6) + "Closing Balance"
                       Else
                          t2 = t1
                          X = Space(63) + X + "  " + Format(DTPicker2.Value, "dd/MM/yyyy") + Space(6) + "Closing Balance"
                       End If
                       Print #1, X
                    End If
                    Print #1, wDotLine
                    Print #1, GProcNoToStr(t1, 12, 2) + Space(51) + GProcNoToStr(t2, 12, 2)
                    Print #1, wDotLine + Chr(27) + Chr(18) + Chr(12)
                Case "M":      'Monthly Party Ledger
                    wDotLine = String(66, "-")
                    T11 = 0
                    t1 = 0
                    t2 = 0
                    For wnp = 0 To 12
                        If wHdPrn Then
                           wHdPrn = False
                           If mScreen = True Then
                               Print #1, gCName
                               Print #1, "Monthly Party Ledger For Period From " & DTPicker1.Value & "  To  " & DTPicker2.Value & "           Page No. " + Str(PgCnt)
                           Else
                               Print #1, Chr(18) + Chr(27) + Chr(14) & gCName & Chr(27) + Chr(18)
                               Print #1, Chr(27) + Chr(69) & "Monthly Party Ledger " & Chr(27) + Chr(70); " From " & DTPicker1.Value & "  To  " & DTPicker2.Value & "           Page No. " + Str(PgCnt)
                           End If
                           wwaname = GProcGetColumnValue("tblmastaccount", "accode", CStr(wAcCode), "N", "acname", "S", , gCn)
                           Print #1, Space(25 - (Len(wwaname) / 2)) + wwaname
                           Print #1, Space(25 - (Len(wwaname) / 2)) + String(Len(wwaname), "-")
                           Print #1, mP12 & wDotLine
                           Print #1, "Month                     Debit        Credit       Balance"
                           Print #1, wDotLine
                           LnCnt = 7
                           PgCnt = PgCnt + 1
                        End If
                        t3 = (wnp - Int(wnp / 10) * 10) + 3 - (Int(wnp / 10) * 2)
                        If (Val(CrSide(t3)) > 0) Or (Val(DrSide(t3)) > 0) Then
                        If (Val(DrSide(t3)) > 0) Then
                           X = DrSide(t3)
                           t1 = t1 + Val(DrSide(t3))
                           T11 = T11 - Val(DrSide(t3))
                        Else
                           X = Space(12)
                        End If
                        If (Val(CrSide(t3)) > 0) Then
                           X = X + "  " + CrSide(t3)
                           T11 = T11 + Val(CrSide(t3))
                           t2 = t2 + Val(CrSide(t3))
                        Else
                           X = X + "  " + Space(12)
                        End If
                        X = X + "  " + GProcNoToStr(Abs(T11), 12, 2)
                        If (T11 < 0) Then
                           X = X + "  Dr"
                        Else
                           X = X + "  Cr"
                        End If
                        Print #1, MonthName(Int(t3)) + Space(11) + X
                        LnCnt = LnCnt + 1
                        End If
                    Next wnp
                    Print #1, wDotLine
                    Print #1, Space(20) + GProcNoToStr(t1, 12, 2) + "  " + GProcNoToStr(t2, 12, 2)
                    Print #1, wDotLine + Chr(12)
                Case "P":      'Partywise Party Ledger
                    wDotLine = String(70, "-")
                    T11 = 0
                    t1 = 0
                    t2 = 0
                    For t3 = 0 To DrCnt
                        If wHdPrn Then
                           wHdPrn = False
                           If mScreen = True Then
                               Print #1, gCName
                               Print #1, "Accountwise Party Ledger For Period From " & DTPicker1.Value & "  To  " & DTPicker2.Value & "           Page No. " + Str(PgCnt)
                           Else
                               Print #1, Chr(18) + Chr(27) + Chr(14) & gCName & Chr(27) + Chr(18)
                               Print #1, Chr(27) + Chr(69) & "Accountwise Party Ledger " & Chr(27) + Chr(70); " From " & DTPicker1.Value & "  To  " & DTPicker2.Value & "           Page No. " + Str(PgCnt)
                           End If
                           wwaname = GProcGetColumnValue("tblmastaccount", "accode", CStr(wAcCode), "N", "acname", "S", , gCn)
                           Print #1, Tab(25 - (Len(wwaname) / 2)), wwaname
                           Print #1, Tab(25 - (Len(wwaname) / 2)), String(Len(wwaname), "-")
                           Print #1, mP12 & wDotLine
                           Print #1, "Account Name                                   Debit        Credit "
                           Print #1, wDotLine
                           LnCnt = 7
                           PgCnt = PgCnt + 1
                        End If
                        If (Val(DrSide(t3)) > 0) Then
                            X = Left(Mid(DrSide(t3), 25) + Space(40), 40)
                            X = X + "  " + Left(DrSide(t3), 12)
                            X = X + "  " + Mid(DrSide(t3), 13, 12)
                            t1 = t1 + Val(Left(DrSide(t3), 12))
                            t2 = t2 + Val(Mid(DrSide(t3), 13, 12))
                            Print #1, X
                            LnCnt = LnCnt + 1
                        End If
                    Next t3
                    Print #1, wDotLine
                    Print #1, Space(42) + GProcNoToStr(t1, 12, 2) + "  " + GProcNoToStr(t2, 12, 2)
                    Print #1, wDotLine + Chr(12)
                End Select
            Loop
        End With
        If wRpTp = "D" Then
            If mScreen = True Then
                Print #1, wDotLine + Chr(27) + Chr(33) + Chr(0) + Chr(27) + Chr(12)
            Else
                Print #1, wDotLine
            End If
        End If
'        If gSelectedmenu = "R23" Then
'            Print #1, wDotLine + Chr(12)
'            '----Display INDEX
'            Dim rstAc As Recordset
'            If GRstOpen(Cn, rstAc, "select ANAME, PGPR from ACCT where PGPR > 0 order by ANAME", "R") > 0 Then
'                PgCnt = 1
'                wHdPrn = True
'                LnCnt = 0
'                With rstAc
'                    .MoveFirst
'                    Do While Not .EOF
'                       If wHdPrn Then
'                            wHdPrn = False
'                            Print #1, "INDEX REPORT                               Page No. " + Str(PgCnt)
'                            Print #1, String(55, "-")
'                            Print #1, "ACCOUNT NAME                                    PAGE NO."
'                            Print #1, String(55, "-")
'                            LnCnt = 4
'                            PgCnt = PgCnt + 1
'                       End If
'                       Print #1, Left(!ANAME + Space(40), 40) + "  " + gprocnotostr(!PGPR, 13, 0)
'                       LnCnt = LnCnt + 1
'                       If (LnCnt > 65) Then
'                          If mScreen = True Then
'                            Print #1, String(55, "-")
'                          Else
'                            Print #1, String(55, "-") & Chr(12)
'                          End If
'                          wHdPrn = True
'                          LnCnt = 0
'                       End If
'                       .MoveNext
'                    Loop
'                    .Close
'                End With
'                Print #1, String(55, "-")
'            End If
'        End If
        
    Close #1
End If

Select Case wRpTp
Case "D", "T", "P", "M":
    Dim RetVal
    If mScreen = False Then
       RetVal = Shell(App.Path & "\txtprint.exe", vbMaximizedFocus)
    ElseIf mScreen = True Then
       RetVal = Shell(App.Path & "\WordPad c:\rpt1.txt", vbMaximizedFocus)
    End If
Case "F":
'    Dim wTotDr As Double, wTotCr As Double, q As Integer, wBal As Double, wCrDr As String
'    Dim wX As String
'    wTotCr = 0
'    wTotDr = 0
'    wBal = 0
'    frmRptFlexLed.msfLed.Clear
'    DoEvents
'    frmRptFlexLed.msfLed.FormatString = "<Date             |<Type  |>No          |<Account Name                                      |>Debit                        |>Credit                           |>Amount                  |       "
'    frmRptFlexLed.lblHead.Caption = (iisGFnGetColumnValue(Cn, "Select ANAME from acct where accode='" & wAcCode1 & "'")) & "  From  " & DTPicker1.Value & "  To  " & DTPicker2.Value
'    With frmRptFlexLed.msfLed
'        rstTran33.MoveFirst
'        Do While Not rstTran33.EOF
'            .TextMatrix(.Row, 0) = rstTran33!Vdt
'            .TextMatrix(.Row, 1) = Left(rstTran33!Nar1, 2)
'            .TextMatrix(.Row, 2) = rstTran33!DocNo & vbNullString
'            .TextMatrix(.Row, 3) = rstTran33!AcName & vbNullString
'            .TextMatrix(.Row, 4) = Format(rstTran33!Debit, "#######0.00")
'            .TextMatrix(.Row, 5) = Format(rstTran33!Credit, "#######0.00")
'            wTotDr = wTotDr + rstTran33!Debit
'            wTotCr = wTotCr + rstTran33!Credit
'            .TextMatrix(.Row, 6) = Format(Abs(wTotDr - wTotCr), "#######0.00")
'            If wTotDr > wTotCr Then
'                wCrDr = "Dr"
'            Else
'                wCrDr = "Cr"
'            End If
'            .TextMatrix(.Row, 7) = wCrDr
'            .Rows = .Rows + 1
'            .Row = .Row + 1
'            For j = 3 To 6
'                Select Case j
'                Case 3:
'                    wX = rstTran33!Nar3 & vbNullString
'                Case 4:
'                    wX = rstTran33!Nar4 & vbNullString
'                Case 5:
'                    wX = rstTran33!NAR5 & vbNullString
'                Case 6:
'                    wX = rstTran33!NAR6 & vbNullString
'                End Select
'                If Len(Trim(wX)) > 0 Then
'                    .TextMatrix(.Row, 3) = Space(5) + wX & vbNullString
'                    .Rows = .Rows + 1
'                    .Row = .Row + 1
'                End If
'           Next j
'           rstTran33.MoveNext
'        Loop
'        .Rows = .Rows + 1
'        .Row = .Row + 1
'        For q = 3 To 6
'            .Col = q
'            .CellFontBold = True
'            .CellForeColor = vbHighlight
'        Next q
'        .TextMatrix(.Row, 3) = "Total"
'        .TextMatrix(.Row, 4) = Format(Round(wTotDr, 2), "#############0.00")
'        .TextMatrix(.Row, 5) = Format(Round(wTotCr, 2), "#############0.00")
'    End With
'    frmRptFlexLed.Show 1
End Select
rstTran33.Close
End Sub
Private Sub DosHdg()
    Dim wAgCode As String, wReHead As String
    wAgCode = GProcGetColumnValue("tblmastaccount", "accode", CStr(wAcCode), "N", "agcode", "N", , gCn)
    wReHead = "Party Ledger"
    If wAgCode = 90019 Then
        wReHead = "Bank Book"
    End If
    
    If mScreen = True Then
       Print #1, gCName
       Print #1, wReHead & " For Period From " & DTPicker1.Value & "  To  " & DTPicker2.Value & "        Page No. " + Str(PgCnt)
       Print #1, wDotLine
    Else
       Print #1, Chr(18) + Chr(27) + Chr(14) & gCName & Chr(27) + Chr(18)
       Print #1, Chr(27) + Chr(69) & wReHead & " For Period " & Chr(27) + Chr(70) & " From " & DTPicker1.Value & "  To  " & DTPicker2.Value & "        Page No. " + Str(PgCnt)
       Print #1, mP12 & wDotLine
    End If
    If (wAgCode = 90019) Then
        Print #1, "   Date   Tp   No.     Account Name                         Receipt    Payment      Balance"
    Else
        Print #1, "   Date   Tp   No.     Account Name                           Debit     Credit      Balance"
    End If
    Print #1, wDotLine
    LnCnt = 5
    PgCnt = PgCnt + 1
End Sub
Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
Dim LeftPos As Long
Dim TopPos As Long
TopPos = txt(Index).Top + Me.Top + 650
LeftPos = txt(Index).Left + Me.Left
If KeyAscii = 13 Then Exit Sub
Select Case Index
    Case 0 '-Group
        gClsSearch.SearchMultiField "tblMastGroup", "AGName", Array("Account Group"), Array(txt(Index).Width), "", Chr(KeyAscii), "AgName", LeftPos, TopPos
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" Then txt(Index) = gClsSearch.SearchMultiRetCol(0)
    Case 1 '-Account
        '-Selected Group
        If txt(0) = "" And gSelectedMenu <> "Account List" Then
            gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,Accode,AcCity,AgName", Array("Account", "AcCode", "AcName", "City", "Group"), Array(4000, 0, 0, 2000, 3000), "", Chr(KeyAscii), "AcName", LeftPos, TopPos
        ElseIf gSelectedMenu <> "Account List" And txt(0) <> "" Then
            gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,Accode,AcCity,AgName", Array("Account", "AcCode", "AcName", "City", "Group"), Array(3000, 0, 0, 2000, 3000), "AgName='" & txt(0) & "'", Chr(KeyAscii), "AcName", LeftPos, TopPos
        ElseIf gSelectedMenu = "Account List" Then
               gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,Accode,AcCity,AgName", Array("Account", "AcCode", "AcName", "City", "Group"), Array(3000, 0, 0, 2000, 3000), "Agcode=90016", Chr(KeyAscii), "AcName", LeftPos, TopPos
        End If
        
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" And gClsSearch.SearchMultiRetCol(1) = "" Then
            txt(Index) = gClsSearch.SearchMultiRetCol(0)
            wAcCode = gClsSearch.SearchMultiRetCol(2)
        ElseIf gClsSearch.SearchMultiRetCol(0) <> "" And gClsSearch.SearchMultiRetCol(1) <> "" Then
            txt(Index) = gClsSearch.SearchMultiRetCol(1)
            wAcCode = gClsSearch.SearchMultiRetCol(2)
        End If
    Case 2 '-Mill
        gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,Accode,AcCity,AgName", Array("Account", "AcCode", "AcName", "City", "Group"), Array(3000, 0, 0, 2000, 3000), " AgCode=90029", Chr(KeyAscii), "AcName", LeftPos, TopPos
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" And gClsSearch.SearchMultiRetCol(1) = "" Then
            txt(Index) = gClsSearch.SearchMultiRetCol(0)
            wAcCode = gClsSearch.SearchMultiRetCol(2)
        ElseIf gClsSearch.SearchMultiRetCol(0) <> "" And gClsSearch.SearchMultiRetCol(1) <> "" Then
            txt(Index) = gClsSearch.SearchMultiRetCol(1)
            wAcCode = gClsSearch.SearchMultiRetCol(2)
        End If
End Select
End Sub

Private Sub txttmp_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub

Private Sub txttmp_LostFocus(Index As Integer)
CheckForNumber
End Sub
