VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmRptGstReports 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Register"
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13245
   Icon            =   "frmRptGstReports.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   13245
   Begin VB.CheckBox chkBox6 
      BackColor       =   &H00EDFDFE&
      Caption         =   "chk Box 6 ?"
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
      Left            =   10215
      TabIndex        =   47
      Top             =   2160
      Visible         =   0   'False
      Width           =   2790
   End
   Begin VB.CheckBox chkBox5 
      BackColor       =   &H00EDFDFE&
      Caption         =   "chk Box 5 ?"
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
      Left            =   10215
      TabIndex        =   46
      Top             =   1740
      Visible         =   0   'False
      Width           =   2790
   End
   Begin VB.CommandButton cmdCrystal 
      Caption         =   "Mail"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   2
      Left            =   6060
      Picture         =   "frmRptGstReports.frx":058A
      TabIndex        =   41
      Top             =   6885
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   2
      Left            =   5175
      TabIndex        =   15
      ToolTipText     =   "Press F3 for Creation"
      Top             =   5535
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   1
      Left            =   5175
      TabIndex        =   14
      ToolTipText     =   "Press F3 for Creation"
      Top             =   5040
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   0
      Left            =   5175
      TabIndex        =   13
      ToolTipText     =   "Press F3 for Creation"
      Top             =   4545
      Visible         =   0   'False
      Width           =   2175
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
      Left            =   1470
      TabIndex        =   7
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1807
      Visible         =   0   'False
      Width           =   1095
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
      Index           =   3
      Left            =   1470
      TabIndex        =   5
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1312
      Visible         =   0   'False
      Width           =   1095
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
      Index           =   2
      Left            =   1470
      TabIndex        =   3
      ToolTipText     =   "Press F3 for Creation"
      Top             =   832
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CheckBox chkBox4 
      BackColor       =   &H00EDFDFE&
      Caption         =   "chk Box 4 ?"
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
      Left            =   10215
      TabIndex        =   36
      Top             =   1365
      Visible         =   0   'False
      Width           =   2790
   End
   Begin VB.CheckBox chkBox3 
      BackColor       =   &H00EDFDFE&
      Caption         =   "chk Box 3 ?"
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
      Left            =   10215
      TabIndex        =   35
      Top             =   630
      Visible         =   0   'False
      Width           =   2835
   End
   Begin VB.CheckBox chkNewPage 
      BackColor       =   &H00EDFDFE&
      Caption         =   "Is New Page After Broker ?"
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
      Left            =   10215
      TabIndex        =   34
      Top             =   1020
      Visible         =   0   'False
      Width           =   2790
   End
   Begin VB.TextBox txtTemp 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   1425
      TabIndex        =   11
      Top             =   3270
      Visible         =   0   'False
      Width           =   6705
   End
   Begin VB.TextBox txtTemp 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   1425
      TabIndex        =   9
      Top             =   2295
      Visible         =   0   'False
      Width           =   6705
   End
   Begin VB.TextBox txtTemp 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   1425
      TabIndex        =   12
      Top             =   3780
      Visible         =   0   'False
      Width           =   6705
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
      Left            =   10215
      TabIndex        =   22
      Top             =   240
      Visible         =   0   'False
      Width           =   2430
   End
   Begin VB.TextBox txtTemp 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   2640
      TabIndex        =   6
      Top             =   1320
      Visible         =   0   'False
      Width           =   6705
   End
   Begin VB.CommandButton cmdCrystal 
      Height          =   480
      Index           =   0
      Left            =   3420
      Picture         =   "frmRptGstReports.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   6885
      Width           =   1185
   End
   Begin VB.CommandButton cmdCrystal 
      Height          =   480
      Index           =   1
      Left            =   4725
      Picture         =   "frmRptGstReports.frx":0E0E
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Print"
      Top             =   6885
      Width           =   1185
   End
   Begin VB.TextBox txtTemp 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   2610
      TabIndex        =   4
      Top             =   840
      Visible         =   0   'False
      Width           =   6705
   End
   Begin VB.TextBox txtTemp 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1425
      TabIndex        =   10
      Top             =   2790
      Visible         =   0   'False
      Width           =   6705
   End
   Begin VB.TextBox txtTemp 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   2640
      TabIndex        =   8
      Top             =   1815
      Visible         =   0   'False
      Width           =   6705
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EDFDFE&
      Height          =   1200
      Left            =   10260
      TabIndex        =   23
      Top             =   2880
      Width           =   2715
      Begin VB.OptionButton opsummary 
         BackColor       =   &H00EDFDFE&
         Caption         =   "Summary"
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
         Left            =   120
         TabIndex        =   33
         Top             =   840
         Visible         =   0   'False
         Width           =   2505
      End
      Begin VB.OptionButton OptShort 
         BackColor       =   &H00EDFDFE&
         Caption         =   "Short"
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
         Left            =   120
         TabIndex        =   25
         Top             =   495
         Visible         =   0   'False
         Width           =   2490
      End
      Begin VB.OptionButton OptDetail 
         BackColor       =   &H00EDFDFE&
         Caption         =   "Detail"
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
         Left            =   120
         TabIndex        =   24
         Top             =   120
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   2550
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7425
      TabIndex        =   28
      Top             =   6878
      Width           =   1335
   End
   Begin Crystal.CrystalReport CrystalReport2 
      Left            =   225
      Top             =   7290
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComctlLib.ProgressBar prgBar 
      Height          =   255
      Left            =   105
      TabIndex        =   29
      Top             =   6240
      Visible         =   0   'False
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComCtl2.DTPicker dtpToDt 
      Height          =   360
      Left            =   6300
      TabIndex        =   2
      Top             =   255
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   635
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
      Format          =   123928579
      CurrentDate     =   38050
   End
   Begin MSComCtl2.DTPicker dtpFrDt 
      Height          =   360
      Left            =   1470
      TabIndex        =   0
      Top             =   255
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   635
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
      Format          =   123928579
      CurrentDate     =   38050
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Bindings        =   "frmRptGstReports.frx":1478
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileType   =   15
      PrintFileLinesPerPage=   60
   End
   Begin MSComCtl2.DTPicker dtpBillsFrDt 
      Height          =   360
      Left            =   11580
      TabIndex        =   42
      Top             =   4680
      Visible         =   0   'False
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   635
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
      Format          =   123928579
      CurrentDate     =   38050
   End
   Begin MSComCtl2.DTPicker dtpBillsToDt 
      Height          =   360
      Left            =   11580
      TabIndex        =   43
      Top             =   5400
      Visible         =   0   'False
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   635
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
      Format          =   123928579
      CurrentDate     =   38050
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bills To Date     :"
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
      Left            =   9900
      TabIndex        =   45
      Top             =   5460
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bills From Date  :"
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
      Left            =   9840
      TabIndex        =   44
      Top             =   4740
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "3)   RCM  I  Gst   : "
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
      Left            =   3150
      TabIndex        =   40
      Top             =   5625
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2)   RCM S Gst    : "
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
      Left            =   3150
      TabIndex        =   39
      Top             =   5130
      Visible         =   0   'False
      Width           =   1710
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1)   RCM C Gst    : "
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
      Left            =   3150
      TabIndex        =   38
      Top             =   4635
      Visible         =   0   'False
      Width           =   1710
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RCM Last Month Credit  :"
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
      Left            =   315
      TabIndex        =   37
      Top             =   4590
      Visible         =   0   'False
      Width           =   2340
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   1710
      Left            =   45
      Top             =   4410
      Width           =   9735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000080&
      X1              =   9990
      X2              =   13095
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "From &Date  :"
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
      Left            =   120
      TabIndex        =   32
      Top             =   300
      Width           =   1170
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00BDDFD1&
      FillStyle       =   0  'Solid
      Height          =   765
      Left            =   3180
      Shape           =   4  'Rounded Rectangle
      Top             =   6780
      Width           =   5895
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000040C0&
      X1              =   13095
      X2              =   13095
      Y1              =   2670
      Y2              =   2655
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Sizer        :"
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
      Left            =   165
      TabIndex        =   31
      Top             =   3270
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Label Label7 
      ForeColor       =   &H000040C0&
      Height          =   330
      Left            =   10485
      TabIndex        =   30
      Top             =   6210
      Width           =   2460
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Item          :"
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
      Left            =   165
      TabIndex        =   17
      Top             =   2295
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Godown    :"
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
      Left            =   165
      TabIndex        =   21
      Top             =   3780
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Broker       :"
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
      Left            =   165
      TabIndex        =   20
      Top             =   1320
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label lblParty 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Party         : "
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
      Left            =   165
      TabIndex        =   19
      Top             =   840
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Label lblTaxType 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Tax Type   :"
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
      Left            =   165
      TabIndex        =   18
      Top             =   2790
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label lblMillName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Mill            :"
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
      Left            =   165
      TabIndex        =   16
      Top             =   1800
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To :"
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
      Left            =   5745
      TabIndex        =   1
      Top             =   255
      Width           =   360
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   4140
      Left            =   45
      Top             =   90
      Width           =   9750
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   4140
      Left            =   9990
      Top             =   135
      Width           =   3120
   End
End
Attribute VB_Name = "frmRptGstReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim wNarrCode As Long, wMillCode As Long, wPartyCode As Long, wBrkCode As Long, wItemCode As Long
Private Sub dtpToDt_LostFocus()
If gSelectedMenu = "GSTR 3B" Then
    Dim LMonStDt As Date
    Dim Lmon As Integer
    Dim LYear As Integer
    Dim x1 As String
    Dim Y1 As String
    Dim tmp1 As Recordset
    x1 = "select sum(tblPurch.PurTaxableAmt),sum(tblpurch.PurRCMCgstAmt) ,sum(tblpurch.PurRCMSgstAmt),sum(tblpurch.PurRCMIgstAmt) " _
         & " From tblPurch where tblPurch.PurBillDt >= ('" & Format(dtpFrDt.Value, "yyyy/mm/dd") & "')  and tblPurch.PurBillDt <= ('" & Format(dtpToDt.Value, "yyyy/mm/dd") & "')  and tblPurch.vfirm =  '" & gCCode & "' and " _
         & " tblPurch.VType = 'PO' and (tblpurch.PurRCMCgstAmt+tblpurch.PurRCMSgstAmt+tblpurch.PurRCMigstAmt) > 0"

    If (GProcRstOpen(tmp1, x1, "R") > 0) Then
       txt(0) = IIf(IsNull(tmp1.Fields(1)), Format(0, FStr), Format(tmp1.Fields(1), FStr))
       txt(1) = IIf(IsNull(tmp1.Fields(2)), Format(0, FStr), Format(tmp1.Fields(2), FStr))
       txt(2) = IIf(IsNull(tmp1.Fields(3)), Format(0, FStr), Format(tmp1.Fields(3), FStr))
    Else
       txt(0) = Format(0, FStr)
       txt(1) = Format(0, FStr)
       txt(2) = Format(0, FStr)
    End If
End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case 13:
       Sendkeys "{TAB}"
    Case 27:
        Unload Me
    End Select
End Sub
Private Sub Form_Load()
wBrkCode = 0
wMillCode = 0
wPartyCode = 0
wNarrCode = 0
wItemCode = 0
Me.Top = 700
Me.Left = Screen.Width / 2 - Me.Width / 2
 '-- Activate Controls to False
GProcActivateControls Me, False
dtpFrDt.Enabled = True
dtpToDt.Enabled = True

dtpFrDt.Value = Format(gCYSDate, "dd/MM/yyyy")
If Format(Now, "dd/MM/yyyy") > Format(gCYEDate, "dd/MM/yyyy") Then
   dtpToDt.Value = Format(gCYEDate, "dd/MM/yyyy")
Else
   dtpToDt.Value = Format(Now, "dd/MM/yyyy")
End If
dtpBillsFrDt = dtpToDt
dtpBillsToDt = dtpToDt
Me.Caption = gSelectedMenu
Select Case gSelectedMenu
''''==== Outstanding Report
Case "GSTR - 1", "Monthwise GST Summery"
     If gSelectedMenu = "Monthwise GST Summery" Then
        OptDetail.Visible = True
        OptDetail.Enabled = True
        OptDetail.Caption = "Monthwise"
        OptShort.Enabled = True
        OptShort.Visible = True
        OptShort.Caption = "Yearly"
      End If
Case "GSTR 3B"
     txt(0).Visible = True
     txt(1).Visible = True
     txt(2).Visible = True
     txt(0).Enabled = True
     txt(1).Enabled = True
     txt(2).Enabled = True
     Label6.Visible = True
     Label8.Visible = True
     Label10.Visible = True
     Label11.Visible = True
Case "GSTR - 2A  -  B2B"
     lblParty.Visible = True
     txtCode(2).Visible = True
     txtTemp(2).Visible = True
     txtCode(2).Enabled = True
     txtTemp(2).Enabled = True
End Select
 '-- Set InActivate Ctrl Back Color
GProcSetInActivateCtrlBackClr Me
End Sub
Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub txt_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub

Private Sub txtCode_LostFocus(Index As Integer)
If Trim(txtCode(Index)) = "" Then Exit Sub
Select Case Index
    Case 2  '-- Party  Code
         txtTemp(2) = GProcGetColumnValue("tblMastAccount", "AcOurCode", UCase(txtCode(2)), "S", "AcName", "S")
         wPartyCode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(Index), "S", "AcCode", "N")
    Case 0  '-- Mill  Code
         txtTemp(0) = GProcGetColumnValue("tblMastAccount", "AcOurCode", UCase(txtCode(0)), "S", "AcName", "S")
    Case 3  '-- Broker  Code
         txtTemp(3) = GProcGetColumnValue("tblMastAccount", "AcOurCode", UCase(txtCode(3)), "S", "AcName", "S")
End Select
End Sub
Private Sub txttemp_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub txtTemp_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then Exit Sub '-- Enter
    Select Case Index
    Case 0 '----Mill Name
        If gSelectedMenu = "Trade Purchase Register GST" Or gSelectedMenu = "Other Purchase Register GST" Or gSelectedMenu = "Purchase Return Register" Then
           gClsSearch.SearchMultiFieldAccMaster Chr(KeyAscii), txtTemp(Index).Left + Me.Left, txtTemp(Index).Top + Me.Top + 1000
            KeyAscii = 0
            If gClsSearch.SearchMultiRetColAccMaster(0) <> "" Then
                txtTemp(Index) = gClsSearch.SearchMultiRetColAccMaster(0)
                wMillCode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(Index), "S", "AcCode", "N")
            End If
        ElseIf gDepotConsiMillHelp = True Then
           gClsSearch.SearchMultiField "qryMillHelpWitIni", "BillSr,MillName,AcOurCode,AcBST,AcMsOmsParty", Array("Mill Ini.", "Mill Name", "Code", "", ""), Array(1000, 8000, 1000, 0, 0), "SaleType in ('27','28') and vfirm = '" & gCCode & "' ", Trim(Chr(KeyAscii)), "MillName", txtTemp(Index).Left + Me.Left, txtTemp(Index).Top + Me.Top + 1000, True
            KeyAscii = 0
            If gClsSearch.SearchMultiRetCol(1) <> "" Then
                txtTemp(Index) = gClsSearch.SearchMultiRetCol(1)
                wMillCode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(Index), "S", "AcCode", "N")
            End If
        Else
            gClsSearch.SearchMultiField "tblMastaccount", "AcName,AcOurCode,AcCity,ACCode", Array("Mill Name", "Code", "City", "Master Code"), Array(8000, 1500, 2000, 0), "(AgCode=90029 ) ", Chr(KeyAscii), "AcName", txtTemp(Index).Left + Me.Left, txtTemp(Index).Top + Me.Top + 1000
            KeyAscii = 0
            If gClsSearch.SearchMultiRetCol(0) <> "" Then
                txtTemp(Index) = gClsSearch.SearchMultiRetCol(0)
                wMillCode = gClsSearch.SearchMultiRetCol(3)
            End If
        End If
    Case 2 '----Party Name
'        gClsSearch.SearchMultiField "QryMastAccount", "AcOurCode,Account,CGName,AcContPer,AcCity,AgName,AcBST,ACCode", Array("Code", "Account Name", "Party Group", "Contact Person", "City", "Group", "TIN", "Code"), Array(1000, 4000, 2000, 2000, 2000, 2500, 0, 0), "(Agcode in (90017) )", Chr(KeyAscii), "", txtTemp(Index).Left + Me.Left, txtTemp(Index).Top + Me.Top + 1000
        gClsSearch.SearchMultiFieldAccMaster Chr(KeyAscii), txtTemp(Index).Left + Me.Left, txtTemp(Index).Top + Me.Top + 1000
        KeyAscii = 0
        If gClsSearch.SearchMultiRetColAccMaster(0) <> "" Then
            txtTemp(Index) = gClsSearch.SearchMultiRetColAccMaster(0)
            wPartyCode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(Index), "S", "AcCode", "N")
        End If
     Case 3 '--Broker
        gClsSearch.SearchMultiField "tblMastaccount", "AcName,AcOurCode,AcCity,ACCode", Array("Broker Name", "Code", "City", "Master Code"), Array(8000, 1500, 2000, 0), "(AgCode=90016 ) ", Chr(KeyAscii), "AcName", txtTemp(Index).Left + Me.Left, txtTemp(Index).Top + Me.Top + 1000
        
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" Then
            txtTemp(Index) = gClsSearch.SearchMultiRetCol(0)
            wBrkCode = gClsSearch.SearchMultiRetCol(3)
        End If
     Case 5 '-- Item
        If txtTemp(0) <> "" Then
           wMillCode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(0), "S", "AcCode", "N")
           gClsSearch.SearchMultiField "tblMastItem,tblMastAccount", "ItName,ItCode,Itticket,Acname", Array("Item", "", "Ticket", "Mill"), Array(3000, 0, 0, 3000), "  tblMastItem.ITMillCode=tblMastaccount.acCode and ITMillCode=" & wMillCode, Chr(KeyAscii), "ItName", txtTemp(Index).Left + Me.Left, txtTemp(Index).Top + Me.Top + 1000
        Else
           gClsSearch.SearchMultiField "tblMastItem,tblMastAccount", "ItName,ItCode,Itticket,Acname", Array("Item", "", "Ticket", "Mill"), Array(3000, 0, 0, 3000), "  tblMastItem.ItMillcode=tblMastAccount.AcCode", Chr(KeyAscii), "ItName", txtTemp(Index).Left + Me.Left, txtTemp(Index).Top + Me.Top + 1000
        End If
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" Then
           wItemCode = gClsSearch.SearchMultiRetCol(1) '-Itcode
           txtTemp(Index) = gClsSearch.SearchMultiRetCol(0) '-ItName
           txtTemp(0) = gClsSearch.SearchMultiRetCol(3) '-Mill
           wMillCode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(0), "S", "AcCode", "N")
        End If
    End Select
End Sub
Private Sub cmdCrystal_Click(Index As Integer)
Dim x1 As String
Dim X As String
Dim DtCondStr As String
Dim ABC As String
Dim ShowRepo As Boolean
ShowRepo = False
Dim MillDepoCd As Long
Dim IsMillDepo As Long
Dim flnm As String
Dim BarCnt As Double
Dim RecYN As Long
Dim tmp1 As Recordset
Dim tmp2 As Recordset
Dim excel_app As Excel.Application
Dim workbook As Excel.workbook
Dim sheet_name As String
Dim sheet As Excel.Worksheet
Dim header_range As Excel.Range
Dim a As Long
Dim xlApp As Object
Dim xlWB As Object
                                
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
x1 = ""
If txtTemp(0) = "" Then wMillCode = 0 '- Mill Code
With CrystalReport1
    GProcCrystalRptPreparation CrystalReport1, IIf(Index = 0, "Window", "Printer")
    Select Case gSelectedMenu
        Case "GSTR 3B"
                ShowRepo = True
                gCmd.CommandText = "PrcPrepareGSTR3B"
                gCmd.Parameters.Refresh
                gCmd.Parameters("@Fr_dt") = Format(dtpFrDt, "yyyy/mm/dd")
                gCmd.Parameters("@To_dt") = Format(dtpToDt, "yyyy/mm/dd")
                gCmd.Parameters("@UCode") = gUserName
                gCmd.Parameters("@VFirm") = Trim(gCCode)
                gCmd.Execute
                Dim ig As Integer
                Dim rstTmp As Recordset
                Dim SaleTaxable As Double, SaleIgst As Double, SaleCgst As Double, SaleSgst As Double
                Dim RcmTaxable As Double, RcmIgst As Double, RcmCgst As Double, RcmSgst As Double
                Dim PurIgst As Double, PurCgst As Double, PurSgst As Double, PurIgstInw As Double
                Dim RTxableSale As Double, RIgstSale As Double, RCgstSale As Double, RSgstSale As Double
                Dim NonGstSale As Double, IntSaleWithOutGst As Double
                
                '--Sale
                x1 = "Select amt1,Amt2,Amt3,Amt4 from TmpGentbl where UserName = '" & gUserName & "' and Vfirm = '" & gCCode & "' and Nar6 = 'C'  "
                ig = GProcRstOpen(rstTmp, x1, "R")
                If ig > 0 Then
                   SaleTaxable = Val(rstTmp.Fields(0) & vbNullString)
                   SaleIgst = Val(rstTmp.Fields(3) & vbNullString)
                   SaleCgst = Val(rstTmp.Fields(1) & vbNullString)
                   SaleSgst = Val(rstTmp.Fields(2) & vbNullString)
                End If
                '--Sale Return
                x1 = "Select amt1,Amt2,Amt3,Amt4 from TmpGentbl where UserName = '" & gUserName & "' and Vfirm = '" & gCCode & "' and Nar6 = 'I'  "
                ig = GProcRstOpen(rstTmp, x1, "R")
                If ig > 0 Then
                   SaleTaxable = SaleTaxable - Val(rstTmp.Fields(0) & vbNullString)
                   SaleIgst = SaleIgst - Val(rstTmp.Fields(3) & vbNullString)
                   SaleCgst = SaleCgst - Val(rstTmp.Fields(1) & vbNullString)
                   SaleSgst = SaleSgst - Val(rstTmp.Fields(2) & vbNullString)
                End If
                '--Non Gst Sale Power Sale)
                x1 = "Select amt1 from TmpGentbl where UserName = '" & gUserName & "' and Vfirm = '" & gCCode & "' and Nar6 = 'F'  "
                ig = GProcRstOpen(rstTmp, x1, "R")
                If ig > 0 Then
                   NonGstSale = Val(rstTmp.Fields(0) & vbNullString)
                End If
               '--Int Sales Without GST
                x1 = "Select amt1 from TmpGentbl where UserName = '" & gUserName & "' and Vfirm = '" & gCCode & "' and Nar6 = 'G'  "
                ig = GProcRstOpen(rstTmp, x1, "R")
                If ig > 0 Then
                   IntSaleWithOutGst = Val(rstTmp.Fields(0) & vbNullString)
                End If
                NonGstSale = NonGstSale + IntSaleWithOutGst
                
                '--RCM Payable
                x1 = "Select amt1,Amt2,Amt3,Amt4 from TmpGentbl where UserName = '" & gUserName & "' and Vfirm = '" & gCCode & "' and Nar6 = 'D'  "
                ig = GProcRstOpen(rstTmp, x1, "R")
                If ig > 0 Then
                   RcmTaxable = Val(rstTmp.Fields(0) & vbNullString)
                   RcmIgst = Val(rstTmp.Fields(3) & vbNullString)
                   RcmCgst = Val(rstTmp.Fields(1) & vbNullString)
                   RcmSgst = Val(rstTmp.Fields(2) & vbNullString)
                End If
                '--Purch
                x1 = "Select sum(amt1),sum(Amt2),sum(Amt3),sum(Amt4) from TmpGentbl where UserName = '" & gUserName & "' and Vfirm = '" & gCCode & "' and Nar6 = 'E'  "
                ig = GProcRstOpen(rstTmp, x1, "R")
                If ig > 0 Then
                   PurIgst = Val(rstTmp.Fields(3) & vbNullString)
                   PurCgst = Val(rstTmp.Fields(1) & vbNullString)
                   PurSgst = Val(rstTmp.Fields(2) & vbNullString)
                End If
                '--Purch Return
                x1 = "Select sum(amt1),sum(Amt2),sum(Amt3),sum(Amt4) from TmpGentbl where UserName = '" & gUserName & "' and Vfirm = '" & gCCode & "' and Nar6 = 'H'  "
                ig = GProcRstOpen(rstTmp, x1, "R")
                If ig > 0 Then
                   PurIgst = PurIgst - Val(rstTmp.Fields(3) & vbNullString)
                   PurCgst = PurCgst - Val(rstTmp.Fields(1) & vbNullString)
                   PurSgst = PurSgst - Val(rstTmp.Fields(2) & vbNullString)
                End If
                
                '--Credit Note Purch gst
                x1 = "Select sum(amt1),sum(Amt2),sum(Amt3),sum(Amt4) from TmpGentbl where UserName = '" & gUserName & "' and Vfirm = '" & gCCode & "' and Nar6 = 'K'  "
                ig = GProcRstOpen(rstTmp, x1, "R")
                If ig > 0 Then
                   PurIgst = PurIgst + Val(rstTmp.Fields(3) & vbNullString)
                   PurCgst = PurCgst + Val(rstTmp.Fields(1) & vbNullString)
                   PurSgst = PurSgst + Val(rstTmp.Fields(2) & vbNullString)
                End If
                '--Debit Note Purch gst
                x1 = "Select sum(amt1),sum(Amt2),sum(Amt3),sum(Amt4) from TmpGentbl where UserName = '" & gUserName & "' and Vfirm = '" & gCCode & "' and Nar6 = 'L'  "
                ig = GProcRstOpen(rstTmp, x1, "R")
                If ig > 0 Then
                   PurIgst = PurIgst - Val(rstTmp.Fields(3) & vbNullString)
                   PurCgst = PurCgst - Val(rstTmp.Fields(1) & vbNullString)
                   PurSgst = PurSgst - Val(rstTmp.Fields(2) & vbNullString)
                End If
                '--Goods Iward Depot
                If gCIsDepotFirm = 1 Then
                    x1 = "Select sum(Amt4) from TmpGentbl where UserName = '" & gUserName & "' and Vfirm = '" & gCCode & "' and Nar6 = 'J' "
                    gCn.BeginTrans
                    ig = GProcRstOpen(rstTmp, x1, "R")
                    If ig > 0 Then
                       PurIgstInw = Val(rstTmp.Fields(0) & vbNullString)
                    End If
                    gCn.CommitTrans
                End If
                RTxableSale = SaleTaxable '+ DbNtTaxable - CrNtTaxable
                RIgstSale = SaleIgst '+ DbNtIgst - CrNtIgst
                RCgstSale = SaleCgst '+ DbNtCgst - CrNtCgst
                RSgstSale = SaleSgst '+ DbNtSgst - CrNtSgst
                PurIgst = PurIgst + PurIgstInw
                .Formulas(4) = "wFrToDt='From : '& '" & dtpFrDt.Value & "' & ' To ' & '" & dtpToDt.Value & "'"
                .Formulas(5) = "wRTxableSale= " & RTxableSale & ""
                .Formulas(6) = "wRIgstSale= " & RIgstSale & ""
                .Formulas(7) = "wRCgstSale= " & RCgstSale & ""
                .Formulas(8) = "wRSgstSale= " & RSgstSale & ""
                .Formulas(9) = "wRcmTaxable= " & RcmTaxable & ""
                .Formulas(10) = "wRcmIgst= " & RcmIgst & ""
                .Formulas(11) = "wRcmCgst= " & RcmCgst & ""
                .Formulas(12) = "wRcmSgst= " & RcmSgst & ""
                .Formulas(13) = "wPurIgst= " & PurIgst & ""
                .Formulas(14) = "wPurCgst= " & PurCgst & ""
                .Formulas(15) = "wPurSgst= " & PurSgst & ""
                .Formulas(16) = "wCompGstin = '" & gCGStin & "'"
                .Formulas(17) = "wLMRcmCgst = " & Val(txt(0)) & ""
                .Formulas(18) = "wLMRcmSgst = " & Val(txt(1)) & ""
                .Formulas(19) = "wLMRcmIgst = " & Val(txt(2)) & ""
                .Formulas(20) = "wNonGstSale = " & NonGstSale & "     "
                x1 = "{TmpGentbl.UserName}='" & gUserName & "'"
               .SelectionFormula = x1
               .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt, "dd/MM/yyyy") & "'"
               .ReportFileName = gReportPath & "rptFormGSTR-3B.rpt"
         Case "GSTR - 2A  -  B2B"
                ShowRepo = True
                gCmd.CommandText = "PrcPrepareGSTR2A"
                gCmd.Parameters.Refresh
                gCmd.Parameters("@Fr_dt") = Format(dtpFrDt, "yyyy/mm/dd")
                gCmd.Parameters("@To_dt") = Format(dtpToDt, "yyyy/mm/dd")
                gCmd.Parameters("@UCode") = gUserName
                gCmd.Parameters("@VFirm") = Trim(gCCode)
                If txtTemp(2) = "" Then
                   wPartyCode = 0
                End If
                If wPartyCode <> 0 Then gCmd.Parameters("@Accode") = wPartyCode
                gCmd.Execute
                x1 = "{TmpGentbl.UserName}='" & gUserName & "'"
               .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt, "dd/MM/yyyy") & "'"
               .Formulas(5) = "wCompGstin = '" & gCGStin & "'"
               .SelectionFormula = x1
               .ReportFileName = gReportPath & "rptFormGSTR-2A.rpt"
         Case "GSTR - 2A  -  B2B  (Mismatch)"
                ShowRepo = True
                gCmd.CommandText = "PrcPrepareGSTR2A"
                gCmd.Parameters.Refresh
                gCmd.Parameters("@Fr_dt") = Format(dtpFrDt, "yyyy/mm/dd")
                gCmd.Parameters("@To_dt") = Format(dtpToDt, "yyyy/mm/dd")
                gCmd.Parameters("@UCode") = gUserName
                gCmd.Parameters("@VFirm") = Trim(gCCode)
                gCmd.Execute
'                flnm = "D:\GSTR1-" + gCCode + "-" + Format(dtpFrDt, "mmm") + ".xlsx"
                flnm = "D:\GSTR2A" + gCCode + ".xlsx"
                Set excel_app = New Excel.Application
                excel_app.Visible = False
                Set workbook = excel_app.Workbooks.Open(filename:=flnm)
                Set sheet = workbook.Sheets.Item("b2b")
                gCn.BeginTrans
                gCn.Execute "delete from tmpGenTbl2 where username='" & gUserName & "'"
                gCn.CommitTrans
                x1 = "Select * from tmpGenTbl2 where username='" & gUserName & "'"
                GProcRstOpen tmp1, x1, "O", gCn
                gCn.BeginTrans
                a = 4
                Do While Len(sheet.cells(a, 1)) > 0
                   tmp1.AddNew
                   tmp1!VFirmName = gCName
                   tmp1!VFirm = gCCode
                   tmp1!PartyName = Left(sheet.cells(a, 1), 50)
                   tmp1!Nar3 = sheet.cells(a, 2)
                   tmp1!Vtype = Left(sheet.cells(a, 2), 2)
                   tmp1!VBillNo = sheet.cells(a, 4)
                   tmp1!VBILLDT = Format(CDate(sheet.cells(a, 5)), "yyyy/MM/DD")
                   tmp1!Amt1 = Val(sheet.cells(a, 6))
                   tmp1!Amt2 = Val(sheet.cells(a, 7))
                   tmp1!Amt3 = Val(sheet.cells(a, 8))
                   tmp1!Amt4 = Val(sheet.cells(a, 9))
                   tmp1!Amt5 = Val(sheet.cells(a, 10))
                   tmp1!UserName = gUserName
                   tmp1.Update
                   a = a + 1
                Loop
                gCn.CommitTrans
                workbook.Close SaveChanges:=True
                excel_app.Quit
                '-- For Matched Date
                gCn.Execute "Update tmpGenTbl set Nar1 = 'Match' from tmpGenTbl,tmpGenTbl2 where tmpGenTbl.Nar3 = tmpGenTbl2.Nar3 and tmpGenTbl.VBillNo = tmpGenTbl2.VBillNo and tmpGenTbl.VBillDt = tmpGenTbl2.VBillDt  and  " _
                            & " tmpGenTbl.amt1 = tmpGenTbl2.amt1 and tmpGenTbl.amt2 = tmpGenTbl2.amt2 and tmpGenTbl.amt3 = tmpGenTbl2.amt3 and tmpGenTbl.amt4 = tmpGenTbl2.amt4 and tmpGenTbl.amt5 = tmpGenTbl2.amt5 and tmpGenTbl.UserName = tmpGenTbl2.UserName and " _
                            & " tmpGenTbl.UserName = '" & gUserName & "' "
                
                gCn.Execute "Update tmpGenTbl2 set Nar1 = 'Match' from tmpGenTbl,tmpGenTbl2 where tmpGenTbl.Nar3 = tmpGenTbl2.Nar3 and tmpGenTbl.VBillNo = tmpGenTbl2.VBillNo and tmpGenTbl.VBillDt = tmpGenTbl2.VBillDt  and  " _
                            & " tmpGenTbl.amt1 = tmpGenTbl2.amt1 and tmpGenTbl.amt2 = tmpGenTbl2.amt2 and tmpGenTbl.amt3 = tmpGenTbl2.amt3 and tmpGenTbl.amt4 = tmpGenTbl2.amt4 and tmpGenTbl.amt5 = tmpGenTbl2.amt5 and tmpGenTbl.UserName = tmpGenTbl2.UserName and " _
                            & " tmpGenTbl2.UserName = '" & gUserName & "' "
                
                gCn.Execute "delete from tmpGenTbl where tmpGenTbl.Nar1 = 'Match' and tmpGenTbl.UserName = '" & gUserName & "' "
                gCn.Execute "delete from tmpGenTbl2 where tmpGenTbl2.Nar1 = 'Match' and tmpGenTbl2.UserName = '" & gUserName & "' "
                '-- For Not Matched Data
                gCn.Execute "Update tmpGenTbl set tmpGenTbl.Amt6 = tmpGenTbl2.Amt3,tmpGenTbl.Amt7 = tmpGenTbl2.Amt4,tmpGenTbl.Amt8 = tmpGenTbl2.Amt5, Nar1 = 'A - Mismatch' from tmpGenTbl,tmpGenTbl2 where tmpGenTbl.Nar3 = tmpGenTbl2.Nar3 and tmpGenTbl.VBillNo = tmpGenTbl2.VBillNo and tmpGenTbl.VBillDt = tmpGenTbl2.VBillDt  and  " _
                            & " (tmpGenTbl.Amt3 <> tmpGenTbl2.Amt3 or tmpGenTbl.Amt4 <> tmpGenTbl2.Amt4 or tmpGenTbl.Amt5 <> tmpGenTbl2.Amt5)  and tmpGenTbl.UserName = tmpGenTbl2.UserName and  tmpGenTbl.UserName = '" & gUserName & "' "
                                
                gCn.Execute "Update tmpGenTbl2 set Nar1 = 'A - Mismatch' from tmpGenTbl,tmpGenTbl2 where tmpGenTbl.Nar3 = tmpGenTbl2.Nar3 and tmpGenTbl.VBillNo = tmpGenTbl2.VBillNo and tmpGenTbl.VBillDt = tmpGenTbl2.VBillDt  and  " _
                            & " (tmpGenTbl.Amt3 <> tmpGenTbl2.Amt3 or tmpGenTbl.Amt4 <> tmpGenTbl2.Amt4 or tmpGenTbl.Amt5 <> tmpGenTbl2.Amt5)  and tmpGenTbl.UserName = tmpGenTbl2.UserName and  tmpGenTbl.UserName = '" & gUserName & "' "
                
                gCn.Execute "delete from tmpGenTbl2 where tmpGenTbl2.Nar1 = 'A - Mismatch' and tmpGenTbl2.UserName = '" & gUserName & "' "
                '-- For Credit Not Received
                gCn.Execute "Update tmpGenTbl set Nar1 = 'B - GST Credit Not Received From Supplier' from tmpGenTbl where tmpGenTbl.UserName = '" & gUserName & "' and tmpGenTbl.Nar1 IS NULL"
                gCn.Execute "Update tmpGenTbl2 set Nar1 = 'C - GST Credit Received From Supplier But Not In Our Books' from tmpGenTbl2 where tmpGenTbl2.UserName = '" & gUserName & "' and tmpGenTbl2.Nar1 IS NULL"
                gCn.BeginTrans
                x1 = "Select VFirmName,Vfirm,PartyName,Nar3,Vtype,VBillNo,VBillDt,Amt1,Amt2,Amt3,Amt4,Amt5,Nar1,UserName from tmpGenTbl2 where username='" & gUserName & "'"
                GProcRstOpen tmp1, x1, "O", gCn
                gCn.Execute " Insert Into TmpGentbl (VFirmName,Vfirm,PartyName,Nar3,Vtype,VBillNo,VBillDt,Amt1,Amt2,Amt3,Amt4,Amt5,Nar1,UserName) " & x1
                gCn.CommitTrans
                '---Bill date check is pending
                x1 = "{TmpGentbl.UserName}='" & gUserName & "'"
               .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt, "dd/MM/yyyy") & "'"
               .Formulas(5) = "wCompGstin = '" & gCGStin & "'"
               .SelectionFormula = x1
               .ReportFileName = gReportPath & "rptFormGSTR-2AMisMatch.rpt"
         Case "GSTR - 1"
               ShowRepo = True
               If gCIsDepotFirm = 1 Then
                  x1 = " tblSale.VType in ('SD')  and (tblSale.VDt)>= ('" & Format(dtpFrDt.Value, "yyyy/mm/dd") & "') and (tblSale.VDt) <= ('" & Format(dtpToDt.Value, "yyyy/mm/dd") & "') and isnull(tblmastaccount.AcGSTIN,'')='' and tblSale.Vfirm = '" & gCCode & "' "  '
               Else
                  x1 = " tblSale.VType in ('SY','SC','SB')  and (tblSale.VDt)>= ('" & Format(dtpFrDt.Value, "yyyy/mm/dd") & "') and (tblSale.VDt) <= ('" & Format(dtpToDt.Value, "yyyy/mm/dd") & "') and isnull(tblmastaccount.AcGSTIN,'')='' and tblSale.Vfirm = '" & gCCode & "' "  '
               End If
               x1 = "Select Distinct AcName,'" & gUserName & "' from tblSale,tblMastAccount where tblSale.SlAcDrCode = tblMastAccount.Accode  and " & x1
                 If GProcRstOpen(tmp1, x1, "R", gCn) > 0 Then
                    gCn.BeginTrans
                    gCn.Execute "delete from tmpGenTbl where username='" & gUserName & "'"
                    gCn.Execute " Insert Into TmpGentbl (Nar1,UserName) " & x1
                    gCn.CommitTrans
                    x1 = "{TmpGentbl.UserName}='" & gUserName & "'"
                   .SelectionFormula = x1
                   .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt, "dd/MM/yyyy") & "'"
                   .ReportFileName = gReportPath & "rptGstinCheck.rpt"
                    .Action = 1
                    Exit Sub
                 End If
                 '---FOr BarCount
                 gCn.BeginTrans
                 gCn.Execute "delete from tmpGenTbl where username='" & gUserName & "'"
                 '-B2B
                 If gCIsDepotFirm = 1 Then
                    x1 = " tblSale.VType in ('SD')  and (tblSale.VDt)>= ('" & Format(dtpFrDt.Value, "yyyy/mm/dd") & "') and (tblSale.VDt) <= ('" & Format(dtpToDt.Value, "yyyy/mm/dd") & "') and len(tblmastaccount.AcGSTIN) > 3 and tblSale.Vfirm = '" & gCCode & "'"  '
                 Else
                    x1 = " tblSale.VType in ('SY','SC','SB')  and (tblSale.VDt)>= ('" & Format(dtpFrDt.Value, "yyyy/mm/dd") & "') and (tblSale.VDt) <= ('" & Format(dtpToDt.Value, "yyyy/mm/dd") & "') and len(tblmastaccount.AcGSTIN) > 3 and tblSale.Vfirm = '" & gCCode & "'"  '
                 End If
                 gCn.Execute " Insert Into TmpGentbl (vno,Nar6,vyear,Vfirm,Nar1,vtype,VbillNo,Vdt,Amt1,Nar4,Nar5,Nar2,Nar3,Amt2,Amt3,Amt4,UserName)" _
                             & " select tblSale.vno,tblSale.vtype,tblSale.vyear,tblSale.Vfirm,upper(tblmastaccount.AcGSTIN),left(tblmastaccount.AcGSTIN,2),tblSale.SlBillNo,tblSale.Vdt,tblSale.SlBillAmt,tblMastState.StateName, 'N','Regular','', iif(tblSale.SlTaxRate3>0,tblSale.SlTaxRate3,tblSale.SlTaxRate+tblSale.SlTaxRate2),tblSale.SlTaxableAmt,0,'" & gUserName & "'" _
                             & " From tblmastaccount,tblSale,tblMastState where tblSale.SlAcDrCode = tblMastAccount.Accode  and left(tblMastAccount.AcGSTIN,2) = tblMastState.StateCd  and " & x1 & ""
                 gCn.CommitTrans
                 '--LP Int. Sales
                 gCn.BeginTrans
                 x1 = " tblIntSale.VType = 'SI'  and (tblIntSale.VDt)>= ('" & Format(dtpFrDt.Value, "yyyy/mm/dd") & "') and (tblIntSale.VDt) <= ('" & Format(dtpToDt.Value, "yyyy/mm/dd") & "') and len(tblmastaccount.AcGSTIN) > 3 and tblIntSale.Vfirm = '" & gCCode & "'"  '
                 
                 gCn.Execute " Insert Into TmpGentbl (vno,Nar6,vyear,Vfirm,Nar1,vtype,PartyName,VbillNo,Vdt,Amt1,Nar4,Nar5,Nar2,Nar3,Amt2,Amt3,Amt4,UserName)" _
                             & " select tblIntSale.vno,tblIntSale.vtype,tblIntSale.vyear,tblIntSale.Vfirm,upper(tblmastaccount.AcGSTIN),left(tblmastaccount.AcGSTIN,2),tblMastAccount.Acname,tblIntSale.CrDrNoteNo,tblIntSale.Vdt,tblIntSale.VAmt,tblMastState.StateName, 'N','Regular','', iif(tblIntSale.IgstRt>0,tblIntSale.IgstRt,tblIntSale.CgstRt+tblIntSale.SgstRt),tblIntSale.Interest,0,'" & gUserName & "'" _
                             & " From tblmastaccount,tblIntSale,tblMastState where tblIntSale.VDrAcCode = tblMastAccount.Accode  and left(tblMastAccount.AcGSTIN,2) = tblMastState.StateCd  and " & x1 & ""
                 gCn.CommitTrans
                 DoEvents
                 DoEvents
                 DoEvents
                 BarCnt = GProcRstOpen(tmp1, "select * from TmpGentbl where TmpGentbl.UserName ='" & gUserName & "' order by vdt,vno  ", "R", gCn)
                 BarCnt = BarCnt * 3
                 If BarCnt = 0 Then BarCnt = 50000
                 '---End Bar Count
                 
                 
                 flnm = "D:\GSTR1-" + gCCode + "-" + Format(dtpFrDt, "mmm") + ".xlsx"
'                 FileCopy gReportPath & "\GSTR1-Blank-1.7.xlsx", flnm
'                 FileCopy gReportPath & "\GSTR1-Blank.xlsx", flnm
                 FileCopy gReportPath & "\GSTR1_V21_Blank.xlsx", flnm
                 gCn.BeginTrans
                 gCn.Execute "delete from tmpGenTbl where username='" & gUserName & "'"
                 '---Sales Detail
                 '--B2B
                 If gCIsDepotFirm = 1 Then
                    x1 = " tblSale.VType in ('SD')  and (tblSale.VDt)>= ('" & Format(dtpFrDt.Value, "yyyy/mm/dd") & "') and (tblSale.VDt) <= ('" & Format(dtpToDt.Value, "yyyy/mm/dd") & "')  and   (tblSale.SlTaxAmt+tblSale.SlTaxAmt2+tblSale.SlTaxAmt3)  > 0    and len(tblmastaccount.AcGSTIN) > 3 and tblSale.Vfirm = '" & gCCode & "'"  '
                 Else
                    x1 = " tblSale.VType in ('SY','SC','SB')  and (tblSale.VDt)>= ('" & Format(dtpFrDt.Value, "yyyy/mm/dd") & "') and (tblSale.VDt) <= ('" & Format(dtpToDt.Value, "yyyy/mm/dd") & "')  and   (tblSale.SlTaxAmt+tblSale.SlTaxAmt2+tblSale.SlTaxAmt3)  > 0    and len(tblmastaccount.AcGSTIN) > 3 and tblSale.Vfirm = '" & gCCode & "'"  '
                 End If
                 gCn.Execute " Insert Into TmpGentbl (vno,Nar6,vyear,Vfirm,Nar1,vtype,PartyName,VbillNo,Vdt,Amt1,Nar4,Nar5,Nar2,Nar3,Amt2,Amt3,Amt4,Amt5,Amt6,Amt7,UserName)" _
                             & " select tblSale.vno,tblSale.vtype,tblSale.vyear,tblSale.Vfirm,upper(tblmastaccount.AcGSTIN),left(tblmastaccount.AcGSTIN,2),tblMastAccount.Acname,tblSale.SlBillNo,tblSale.Vdt,tblSale.SlBillAmt,tblMastState.StateName, 'N','Regular','', iif(tblSale.SlTaxRate3>0,tblSale.SlTaxRate3,tblSale.SlTaxRate+tblSale.SlTaxRate2),tblSale.SlTaxableAmt,0, " _
                             & " tblSale.SlTaxAmt3,tblSale.SlTaxAmt,tblSale.SlTaxAmt2,'" & gUserName & "'" _
                             & " From tblmastaccount,tblSale,tblMastState where tblSale.SlAcDrCode = tblMastAccount.Accode  and left(tblMastAccount.AcGSTIN,2) = tblMastState.StateCd  and " & x1 & ""
                 gCn.CommitTrans
                 Set excel_app = New Excel.Application
                 prgBar.Max = BarCnt
                 prgBar.Value = 0
                 prgBar.Visible = True
    '             excel_app.Visible = True
                 excel_app.Visible = False
                 Set workbook = excel_app.Workbooks.Open(filename:=flnm)
                 RecYN = GProcRstOpen(tmp1, "select * from TmpGentbl where TmpGentbl.UserName ='" & gUserName & "' order by vdt,Nar6,vno   ", "R", gCn)
                 
'                 Set sheet = workbook.Sheets.Item("Home")
'                 If (sheet Is Nothing) Then
'                     MsgBox "Error Sales Sheet"
'                 End If
'                 sheet.cells(3, 2) = UCase(gCGStin)
'                 sheet.cells(4, 2) = gCName
'                 sheet.cells(5, 2) = CStr(Year(gCYSDate)) + "-" + CStr(Year(gCYEDate))
'                 sheet.cells(5, 4) = Format(dtpFrDt.Value, "mmmm")
                 
                 If RecYN > 0 Then
                    Set sheet = workbook.Sheets.Item("b2b,sez,de")
                    If (sheet Is Nothing) Then
                        MsgBox "Error Sales Sheet"
                    End If
                    a = 5
                    With tmp1
                         .MoveFirst
                         Do While Not .EOF
                            sheet.cells(a, 1) = UCase(!Nar1)
                            If Val(Right(!PartyName, 2)) > 0 Then
                               sheet.cells(a, 2) = Left(!PartyName, Len(!PartyName) - 2)
                            Else
                               sheet.cells(a, 2) = !PartyName
                            End If
                            sheet.cells(a, 3) = !VBillNo
                            sheet.cells(a, 4) = Format(!Vdt, "dd-MMM-YYYY")
                            sheet.cells(a, 5) = !Amt1
                            sheet.cells(a, 6) = !Nar4
                            sheet.cells(a, 7) = "N"
                            sheet.cells(a, 9) = "Regular B2B"
                            sheet.cells(a, 11) = !Amt2
                            sheet.cells(a, 12) = !Amt3
                            sheet.cells(a, 13) = 0
                            a = a + 1
'                            prgBar.Value = prgBar.Value + 1
                           .MoveNext
                         Loop
                    End With
                 End If

'                --B2CL
                 gCn.BeginTrans
                 gCn.Execute "delete from tmpGenTbl where username='" & gUserName & "'"
                 If gCIsDepotFirm = 1 Then
                    x1 = " tblSale.VType in ('SD')  and (tblSale.VDt)>= ('" & Format(dtpFrDt.Value, "yyyy/mm/dd") & "') and (tblSale.VDt) <= ('" & Format(dtpToDt.Value, "yyyy/mm/dd") & "') and (len(tblmastaccount.AcGSTIN) <= 3 or isnull(tblmastaccount.AcGSTIN,'')='')    and   (tblSale.SlTaxAmt+tblSale.SlTaxAmt2+tblSale.SlTaxAmt3)  > 0    and tblSale.Vfirm = '" & gCCode & "' and tblSale.SlBillAmt >= 250000"  '
                 Else
                    x1 = " tblSale.VType in ('SY','SC','SB')  and (tblSale.VDt)>= ('" & Format(dtpFrDt.Value, "yyyy/mm/dd") & "') and (tblSale.VDt) <= ('" & Format(dtpToDt.Value, "yyyy/mm/dd") & "') and (len(tblmastaccount.AcGSTIN) <= 3 or isnull(tblmastaccount.AcGSTIN,'')='')    and   (tblSale.SlTaxAmt+tblSale.SlTaxAmt2+tblSale.SlTaxAmt3)  > 0    and tblSale.Vfirm = '" & gCCode & "' and tblSale.SlBillAmt >= 250000"  '
                 End If
                 gCn.Execute " Insert Into TmpGentbl (vno,Nar6,vyear,Vfirm,Nar1,vtype,PartyName,VbillNo,Vdt,Amt1,Nar4,Nar5,Nar2,Nar3,Amt2,Amt3,Amt4,Amt5,UserName)" _
                             & " select tblSale.vno,tblSale.vtype,tblSale.vyear,tblSale.Vfirm,upper(tblmastaccount.AcGSTIN),left(tblmastaccount.AcGSTIN,2),tblMastAccount.Acname,tblSale.SlBillNo,tblSale.Vdt,tblSale.SlBillAmt,tblMastState.StateName, 'N','Regular','', iif(tblSale.SlTaxRate3>0,tblSale.SlTaxRate3,tblSale.SlTaxRate+tblSale.SlTaxRate2),tblSale.SlTaxableAmt,0,tblSale.SlTaxAmt3,'" & gUserName & "'" _
                             & " From tblmastaccount,tblSale,tblMastState where tblSale.SlAcDrCode = tblMastAccount.Accode and left(tblMastAccount.AcGSTIN,2) = tblMastState.StateCd  and " & x1 & ""
                 gCn.CommitTrans
                 RecYN = GProcRstOpen(tmp1, "select * from TmpGentbl where TmpGentbl.UserName ='" & gUserName & "'", "R", gCn)
                 If RecYN > 0 Then
                    Set sheet = workbook.Sheets.Item("b2cl")
                    If (sheet Is Nothing) Then
                        MsgBox "Error Sales Sheet"
                    End If
                    a = 5
                    With tmp1
                         .MoveFirst
                         Do While Not .EOF
                            sheet.cells(a, 1) = !VBillNo
                            sheet.cells(a, 2) = Format(!Vdt, "dd-MMM-YYYY")
                            sheet.cells(a, 3) = !Amt1
                            sheet.cells(a, 4) = !Nar4
                            sheet.cells(a, 6) = !Amt2
                            sheet.cells(a, 7) = !Amt3
                            sheet.cells(a, 8) = 0
                            a = a + 1
                            prgBar.Value = prgBar.Value + 1
                           .MoveNext
                         Loop
                    End With
                 End If
'                 '--B2CS
                 gCn.BeginTrans
                 gCn.Execute "delete from tmpGenTbl where username='" & gUserName & "'"
                 If gCIsDepotFirm = 1 Then
                    x1 = " tblSale.VType in ('SD')  and (tblSale.VDt)>= ('" & Format(dtpFrDt.Value, "yyyy/mm/dd") & "') and (tblSale.VDt) <= ('" & Format(dtpToDt.Value, "yyyy/mm/dd") & "') and (len(tblmastaccount.AcGSTIN) <= 3 or isnull(tblmastaccount.AcGSTIN,'')='')  and   (tblSale.SlTaxAmt+tblSale.SlTaxAmt2+tblSale.SlTaxAmt3)  > 0    and tblSale.Vfirm = '" & gCCode & "' and tblSale.SlBillAmt < 250000"  '
                 Else
                    x1 = " tblSale.VType in ('SY','SC','SB')  and (tblSale.VDt)>= ('" & Format(dtpFrDt.Value, "yyyy/mm/dd") & "') and (tblSale.VDt) <= ('" & Format(dtpToDt.Value, "yyyy/mm/dd") & "') and (len(tblmastaccount.AcGSTIN) <= 3 or isnull(tblmastaccount.AcGSTIN,'')='')  and   (tblSale.SlTaxAmt+tblSale.SlTaxAmt2+tblSale.SlTaxAmt3)  > 0    and tblSale.Vfirm = '" & gCCode & "' and tblSale.SlBillAmt < 250000"  '
                 End If
                 gCn.Execute " Insert Into TmpGentbl (Nar4,Nar5,Nar2,Amt2,Amt3,Amt4,UserName)" _
                             & " select upper(tblmastaccount.AcGSTIN),left(tblmastaccount.AcGSTIN,2),tblMastState.StateName,iif(tblSale.SlTaxRate3>0,tblSale.SlTaxRate3,tblSale.SlTaxRate+tblSale.SlTaxRate2),tblSale.SlTaxableAmt,0,'" & gUserName & "'" _
                             & " From tblmastaccount,tblSale,tblMastState where tblSale.SlAcDrCode = tblMastAccount.Accode and left(tblMastAccount.AcGSTIN,2) = tblMastState.StateCd  and " & x1 & ""

                 gCn.CommitTrans
                 If GProcRstOpen(tmp1, "select * from TmpGentbl where TmpGentbl.UserName ='" & gUserName & "'", "R", gCn) > 0 Then
                    Set sheet = workbook.Sheets.Item("b2cs")
                    If (sheet Is Nothing) Then
                        MsgBox "Error Sales Sheet"
                    End If
                    a = 5
                    With tmp1
                         .MoveFirst
                         Do While Not .EOF
                            sheet.cells(a, 1) = "OE"
                            DoEvents
                            sheet.cells(a, 2) = !Nar2
                            sheet.cells(a, 4) = !Amt2
                            sheet.cells(a, 5) = !Amt3
                            sheet.cells(a, 6) = 0
                            a = a + 1
  '                          prgBar.Value = prgBar.Value + 1
                           .MoveNext
                         Loop
                    End With
                 End If
                 
                 
                 '---Sales Return              '--CDNR
                 x1 = " tblSale.VType = 'RY'  and (tblSale.VDt)>= ('" & Format(dtpFrDt.Value, "yyyy/mm/dd") & "') and (tblSale.VDt) <= ('" & Format(dtpToDt.Value, "yyyy/mm/dd") & "')  and   (tblSale.SlTaxAmt+tblSale.SlTaxAmt2+tblSale.SlTaxAmt3)  > 0    and len(tblmastaccount.AcGSTIN) > 3 and tblSale.Vfirm = '" & gCCode & "'"  '
                 gCn.BeginTrans
                 gCn.Execute "delete from tmpGenTbl where username='" & gUserName & "'"
                 gCn.Execute " Insert Into TmpGentbl (Nar7,Nar6,vyear,Vfirm,Nar1,vtype,PartyName,VbillNo,Vdt,Amt1,Nar4,Nar5,Nar2,Nar3,Amt2,Amt3,Amt4,Amt5,Amt6,Amt7,UserName)" _
                             & " select 'SRTN/'+ltrim(str(tblSale.vno)),tblSale.vtype,tblSale.vyear,tblSale.Vfirm,upper(tblmastaccount.AcGSTIN),left(tblmastaccount.AcGSTIN,2),tblMastAccount.Acname,tblSale.SlBillNo,tblSale.Vdt,tblSale.SlBillAmt,tblMastState.StateName, 'N','Regular','', iif(tblSale.SlTaxRate3>0,tblSale.SlTaxRate3,tblSale.SlTaxRate+tblSale.SlTaxRate2),tblSale.SlTaxableAmt,0,tblSale.SlTaxAmt3,tblSale.SlTaxAmt,tblSale.SlTaxAmt2,'" & gUserName & "'" _
                             & " From tblmastaccount,tblSale,tblMastState where tblSale.SlAcDrCode = tblMastAccount.Accode  and left(tblMastAccount.AcGSTIN,2) = tblMastState.StateCd  and " & x1 & ""


                 gCn.CommitTrans
                 '--LP Int. Sales
                 gCn.BeginTrans
                 x1 = " tblIntSale.VType = 'SI'  and (tblIntSale.VDt)>= ('" & Format(dtpFrDt.Value, "yyyy/mm/dd") & "') and (tblIntSale.VDt) <= ('" & Format(dtpToDt.Value, "yyyy/mm/dd") & "') and len(tblmastaccount.AcGSTIN) > 3 and tblIntSale.Vfirm = '" & gCCode & "'"  '

                 gCn.Execute " Insert Into TmpGentbl (vno,Nar6,vyear,Vfirm,Nar1,vtype,PartyName,VbillNo,Vdt,Amt1,Nar4,Nar5,Nar2,Nar3,Amt2,Amt3,Amt4,Amt5,Amt6,Amt7,UserName)" _
                             & " select tblIntSale.vno,tblIntSale.vtype,tblIntSale.vyear,tblIntSale.Vfirm,upper(tblmastaccount.AcGSTIN),left(tblmastaccount.AcGSTIN,2),tblMastAccount.Acname,tblIntSale.CrDrNoteNo,tblIntSale.Vdt,tblIntSale.VAmt,tblMastState.StateName, 'N','Regular','', iif(tblIntSale.IgstRt>0,tblIntSale.IgstRt,tblIntSale.CgstRt+tblIntSale.SgstRt),tblIntSale.Interest,0,tblIntSale.IgstAmt,tblIntSale.CgstAmt,tblIntSale.SgstAmt,'" & gUserName & "'" _
                             & " From tblmastaccount,tblIntSale,tblMastState where tblIntSale.VDrAcCode = tblMastAccount.Accode  and left(tblMastAccount.AcGSTIN,2) = tblMastState.StateCd  and " & x1 & ""
                 gCn.CommitTrans
                 '--Credit Note Sales
                 gCn.BeginTrans
                 x1 = " tblIntSale.VType = 'PX'  and (tblIntSale.VDt)>= ('" & Format(dtpFrDt.Value, "yyyy/mm/dd") & "') and (tblIntSale.VDt) <= ('" & Format(dtpToDt.Value, "yyyy/mm/dd") & "') and len(tblmastaccount.AcGSTIN) > 3 and tblIntSale.Vfirm = '" & gCCode & "'"  '

                 gCn.Execute " Insert Into TmpGentbl (vno,Nar6,vyear,Vfirm,Nar1,vtype,PartyName,VbillNo,Vdt,Amt1,Nar4,Nar5,Nar2,Nar3,Amt2,Amt3,Amt4,Amt5,Amt6,Amt7,UserName)" _
                             & " select tblIntSale.vno,tblIntSale.vtype,tblIntSale.vyear,tblIntSale.Vfirm,upper(tblmastaccount.AcGSTIN),left(tblmastaccount.AcGSTIN,2),tblMastAccount.Acname,tblIntSale.Vno,tblIntSale.Vdt,tblIntSale.VAmt,tblMastState.StateName, 'N','Regular','', iif(tblIntSale.IgstRt>0,tblIntSale.IgstRt,tblIntSale.CgstRt+tblIntSale.SgstRt),tblIntSale.Interest,0,tblIntSale.IgstAmt,tblIntSale.CgstAmt,tblIntSale.SgstAmt,'" & gUserName & "'" _
                             & " From tblmastaccount,tblIntSale,tblMastState where tblIntSale.VCrAcCode = tblMastAccount.Accode  and left(tblMastAccount.AcGSTIN,2) = tblMastState.StateCd  and " & x1 & ""
                 gCn.CommitTrans
                 DoEvents
                 RecYN = GProcRstOpen(tmp1, "select * from TmpGentbl where TmpGentbl.UserName ='" & gUserName & "' order by vdt,Nar6,vno   ", "R", gCn)
                 If RecYN > 0 Then
                    Set sheet = workbook.Sheets.Item("cdnr")
                    If (sheet Is Nothing) Then
                        MsgBox "Error Sales Sheet"
                    End If
                    a = 5
                    With tmp1
                         .MoveFirst
                         Do While Not .EOF
                            sheet.cells(a, 1) = !Nar1
                            If Val(Right(!PartyName, 2)) > 0 Then
                               sheet.cells(a, 2) = Left(!PartyName, Len(!PartyName) - 2)
                            Else
                               sheet.cells(a, 2) = !PartyName
                            End If
                            If !Nar6 = "SI" Then
                               sheet.cells(a, 3) = !VBillNo
                            ElseIf !Nar6 = "RY" Then
                               sheet.cells(a, 3) = !Nar7
                            Else
                               sheet.cells(a, 3) = !Vno
                            End If
                            sheet.cells(a, 4) = Format(!Vdt, "dd-MMM-YYYY")
                            If !Nar6 = "SI" Then
                               sheet.cells(a, 5) = "D"
                            Else
                               sheet.cells(a, 5) = "C"
                            End If
                            sheet.cells(a, 6) = !Nar4
                            sheet.cells(a, 7) = "N"
                            sheet.cells(a, 8) = "Regular B2B"
                            sheet.cells(a, 9) = !Amt1
                            sheet.cells(a, 11) = !Amt2
                            sheet.cells(a, 12) = !Amt3
                            sheet.cells(a, 13) = 0
                            a = a + 1
                            prgBar.Value = prgBar.Value + 1
                           .MoveNext
                         Loop
                    End With
                 End If
'                 '--Exempt
                 gCn.BeginTrans
                 gCn.Execute "delete from tmpGenTbl where username='" & gUserName & "'"
                 If gCIsDepotFirm = 1 Then
                    x1 = " tblSale.VType in ('SD')  and (tblSale.VDt)>= ('" & Format(dtpFrDt.Value, "yyyy/mm/dd") & "') and (tblSale.VDt) <= ('" & Format(dtpToDt.Value, "yyyy/mm/dd") & "')   and   (tblSale.SlTaxAmt+tblSale.SlTaxAmt2+tblSale.SlTaxAmt3)  = 0    and tblSale.Vfirm = '" & gCCode & "' "  '
                 Else
                    x1 = " tblSale.VType in ('SY','SC','SB')  and (tblSale.VDt)>= ('" & Format(dtpFrDt.Value, "yyyy/mm/dd") & "') and (tblSale.VDt) <= ('" & Format(dtpToDt.Value, "yyyy/mm/dd") & "')   and   (tblSale.SlTaxAmt+tblSale.SlTaxAmt2+tblSale.SlTaxAmt3)  = 0    and tblSale.Vfirm = '" & gCCode & "' "  '
                 End If
                 gCn.Execute " Insert Into TmpGentbl (Amt1,UserName)" _
                             & " select sum(tblSale.SlBillAmt-tblSale.SlTcsAmt),'" & gUserName & "'" _
                             & " From tblSale where " & x1 & ""
                 gCn.CommitTrans
                 RecYN = GProcRstOpen(tmp1, "select * from TmpGentbl where TmpGentbl.UserName ='" & gUserName & "'", "R", gCn)
                 a = 5
                 If RecYN > 0 Then
                    Set sheet = workbook.Sheets.Item("exemp")
                    If (sheet Is Nothing) Then
                        MsgBox "Error Sales Sheet"
                    End If
                    a = 5
                    With tmp1
                         .MoveFirst
                         Do While Not .EOF
                            sheet.cells(a, 1) = "Intra-state supplies to registered person"
                            DoEvents
                            sheet.cells(a, 2) = 0
                            sheet.cells(a, 3) = 0
                            sheet.cells(a, 4) = !Amt1
                            a = a + 1
'                            prgBar.Value = prgBar.Value + 1
                           .MoveNext
                         Loop
                    End With
                 End If
'                 '--Exempt Non GST Int Sale
                 gCn.BeginTrans
                 gCn.Execute "delete from tmpGenTbl where username='" & gUserName & "'"
                 x1 = " tblIntSale.VType = 'MI'  and (tblIntSale.VDt)>= ('" & Format(dtpFrDt.Value, "yyyy/mm/dd") & "') and (tblIntSale.VDt) <= ('" & Format(dtpToDt.Value, "yyyy/mm/dd") & "')   and   (tblIntSale.CgstAmt+tblIntSale.SgstAmt+tblIntSale.IgstAmt)  = 0    and tblIntSale.Vfirm = '" & gCCode & "' "  '
                 gCn.Execute " Insert Into TmpGentbl (Amt1,UserName)" _
                             & " select sum(tblIntSale.Interest),'" & gUserName & "'" _
                             & " From tblIntSale where " & x1 & ""
                 gCn.CommitTrans
                 RecYN = GProcRstOpen(tmp1, "select * from TmpGentbl where TmpGentbl.UserName ='" & gUserName & "'", "R", gCn)
                 If RecYN > 0 Then
                    Set sheet = workbook.Sheets.Item("exemp")
                    If (sheet Is Nothing) Then
                        MsgBox "Error Sales Sheet"
                    End If
                    With tmp1
                         .MoveFirst
                         Do While Not .EOF
                            sheet.cells(a, 1) = "Non GST Interest"
                            DoEvents
                            sheet.cells(a, 2) = 0
                            sheet.cells(a, 3) = 0
                            sheet.cells(a, 4) = !Amt1
                            a = a + 1
    '                        prgBar.Value = prgBar.Value + 1
                           .MoveNext
                         Loop
                    End With
                 End If

'                 '--HSN
                 gCn.BeginTrans
                 gCn.Execute "delete from tmpGenTbl where username='" & gUserName & "'"
                 If gCIsDepotFirm = 1 Then
                    x1 = " tblSale.VType in ('SD')  and (tblSale.VDt)>= ('" & Format(dtpFrDt.Value, "yyyy/mm/dd") & "') and (tblSale.VDt) <= ('" & Format(dtpToDt.Value, "yyyy/mm/dd") & "') and tblSale.Vfirm = '" & gCCode & "' and (tblSale.SlTaxAmt+tblSale.SlTaxAmt2+tblSale.SlTaxAmt3) > 0   "  '
                 Else
                    x1 = " tblSale.VType in ('SY','SC','SB')  and (tblSale.VDt)>= ('" & Format(dtpFrDt.Value, "yyyy/mm/dd") & "') and (tblSale.VDt) <= ('" & Format(dtpToDt.Value, "yyyy/mm/dd") & "') and tblSale.Vfirm = '" & gCCode & "' and (tblSale.SlTaxAmt+tblSale.SlTaxAmt2+tblSale.SlTaxAmt3) > 0   "  '
                 End If
                 If gCYear >= 2021 Then
                    gCn.Execute " Insert Into TmpGentbl (Nar1,Nar2,Nar3,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,Amt7,UserName)" _
                                & " select tblMastItem.ItHsn2,iif(tblMastItem.IsOtherItem=0,'Yarn',''),iif(tblMastItem.ItUnit= 'Bag','BAG-BAGS',iif(tblMastItem.ItUnit = 'Cartoon','CTN-CARTONS',iif(Left(tblMastItem.ItUnit,2) = 'No','NOS-NUMBERS',tblMastItem.ItUnit))),sum(tblSaleSub.SlSubBag),sum(tblSale.SlBillAmt),sum(tblSale.SlTaxableAmt),sum(tblSale.SlTaxAmt3),sum(tblSale.SlTaxAmt),sum(tblSale.SlTaxAmt2),(tblSale.SlTaxRate+tblSale.SlTaxRate2+tblSale.SlTaxRate3),'" & gUserName & "'" _
                                & " From tblSale,tblSaleSub,tblMastItem where tblSaleSub.SlSubItCode = tblMastItem.ItCode and tblSale.Vno = tblSaleSub.Vno and tblSale.Vtype = tblSaleSub.Vtype and tblSale.Vyear = tblSaleSub.Vyear and tblSale.Vfirm = tblSaleSub.Vfirm  and " & x1 & "" _
                                & " group by tblMastItem.ItHsn2,iif(tblMastItem.IsOtherItem=0,'Yarn',''),tblMastItem.ItUnit,(tblSale.SlTaxRate+tblSale.SlTaxRate2+tblSale.SlTaxRate3)"
                 Else
                    gCn.Execute " Insert Into TmpGentbl (Nar1,Nar2,Nar3,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,Amt7,UserName)" _
                                & " select tblMastItem.ItHsn,iif(tblMastItem.IsOtherItem=0,'Yarn',''),iif(tblMastItem.ItUnit= 'Bag','BAG-BAGS',iif(tblMastItem.ItUnit = 'Cartoon','CTN-CARTONS',iif(Left(tblMastItem.ItUnit,2) = 'No','NOS-NUMBERS',tblMastItem.ItUnit))),sum(tblSaleSub.SlSubBag),sum(tblSale.SlBillAmt),sum(tblSale.SlTaxableAmt),sum(tblSale.SlTaxAmt3),sum(tblSale.SlTaxAmt),sum(tblSale.SlTaxAmt2),(tblSale.SlTaxRate+tblSale.SlTaxRate2+tblSale.SlTaxRate3),'" & gUserName & "'" _
                                & " From tblSale,tblSaleSub,tblMastItem where tblSaleSub.SlSubItCode = tblMastItem.ItCode and tblSale.Vno = tblSaleSub.Vno and tblSale.Vtype = tblSaleSub.Vtype and tblSale.Vyear = tblSaleSub.Vyear and tblSale.Vfirm = tblSaleSub.Vfirm  and " & x1 & "" _
                                & " group by tblMastItem.ItHsn,iif(tblMastItem.IsOtherItem=0,'Yarn',''),tblMastItem.ItUnit,(tblSale.SlTaxRate+tblSale.SlTaxRate2+tblSale.SlTaxRate3)"
                 End If
                 gCn.CommitTrans

                 '--LP Int. Sales
                 gCn.BeginTrans
                 x1 = " tblIntSale.VType = 'SI'  and (tblIntSale.VDt)>= ('" & Format(dtpFrDt.Value, "yyyy/mm/dd") & "') and (tblIntSale.VDt) <= ('" & Format(dtpToDt.Value, "yyyy/mm/dd") & "') and len(tblmastaccount.AcGSTIN) > 3 and tblIntSale.Vfirm = '" & gCCode & "'"  '
                 If gCYear >= 2021 Then
                    gCn.Execute " Insert Into TmpGentbl (Nar1,Nar2,Nar3,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,Amt7,UserName)" _
                                & " select tblMastItem.ItHsn2,'Interest','NOS-NUMBERS',0,sum(tblIntSale.VAmt),sum(tblIntSale.Interest),sum(tblIntSale.IgstAmt),sum(tblIntSale.CgstAmt),sum(tblIntSale.SgstAmt),tblIntSale.CgstRt+tblIntSale.SgstRt+tblIntSale.IgstRt,'" & gUserName & "'" _
                                & " From tblIntSale,tblmastaccount,tblMastState,tblMastItem where tblIntSale.VDrAcCode = tblMastAccount.Accode  and left(tblMastAccount.AcGSTIN,2) = tblMastState.StateCd and tblIntSale.ItCode = tblMastItem.ItCode and " & x1 & " " _
                                & " group by tblMastItem.ItHsn2,tblIntSale.VCtrNo,tblIntSale.CgstRt+tblIntSale.SgstRt+tblIntSale.IgstRt"
                 Else
                    gCn.Execute " Insert Into TmpGentbl (Nar1,Nar2,Nar3,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,Amt7,UserName)" _
                                & " select tblMastItem.ItHsn,'Interest','NOS-NUMBERS',0,sum(tblIntSale.VAmt),sum(tblIntSale.Interest),sum(tblIntSale.IgstAmt),sum(tblIntSale.CgstAmt),sum(tblIntSale.SgstAmt),tblIntSale.CgstRt+tblIntSale.SgstRt+tblIntSale.IgstRt,'" & gUserName & "'" _
                                & " From tblIntSale,tblmastaccount,tblMastState,tblMastItem where tblIntSale.VDrAcCode = tblMastAccount.Accode  and left(tblMastAccount.AcGSTIN,2) = tblMastState.StateCd and tblIntSale.ItCode = tblMastItem.ItCode and " & x1 & " " _
                                & " group by tblMastItem.ItHsn,tblIntSale.VCtrNo,tblIntSale.CgstRt+tblIntSale.SgstRt+tblIntSale.IgstRt"
                 End If
                 gCn.CommitTrans
                 
                 '--HSN Sales Rtn
                 
                 x1 = " tblSale.VType = 'RY'  and (tblSale.VDt)>= ('" & Format(dtpFrDt.Value, "yyyy/mm/dd") & "') and (tblSale.VDt) <= ('" & Format(dtpToDt.Value, "yyyy/mm/dd") & "')  and   (tblSale.SlTaxAmt+tblSale.SlTaxAmt2+tblSale.SlTaxAmt3)  > 0   and tblSale.Vfirm = '" & gCCode & "'"  '
                 gCn.BeginTrans
                 gCn.Execute " Insert Into TmpGentbl (Nar1,Nar2,Nar3,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,Amt7,UserName)" _
                                & " select tblMastItem.ItHsn2,iif(tblMastItem.IsOtherItem=0,'Credit Not Sl Rtn.',''),iif(tblMastItem.ItUnit= 'Bag','BAG-BAGS',iif(tblMastItem.ItUnit = 'Cartoon','CTN-CARTONS',iif(Left(tblMastItem.ItUnit,2) = 'No','NOS-NUMBERS',tblMastItem.ItUnit))),sum(tblSaleSub.SlSubBag)*-1,sum(tblSale.SlBillAmt)*-1,sum(tblSale.SlTaxableAmt)*-1,sum(tblSale.SlTaxAmt3)*-1,sum(tblSale.SlTaxAmt)*-1,sum(tblSale.SlTaxAmt2)*-1,(tblSale.SlTaxRate+tblSale.SlTaxRate2+tblSale.SlTaxRate3),'" & gUserName & "'" _
                                & " From tblSale,tblSaleSub,tblMastItem where tblSaleSub.SlSubItCode = tblMastItem.ItCode and tblSale.Vno = tblSaleSub.Vno and tblSale.Vtype = tblSaleSub.Vtype and tblSale.Vyear = tblSaleSub.Vyear and tblSale.Vfirm = tblSaleSub.Vfirm  and " & x1 & "" _
                                & " group by tblMastItem.ItHsn2,iif(tblMastItem.IsOtherItem=0,'Credit Not Sl Rtn.',''),tblMastItem.ItUnit,(tblSale.SlTaxRate+tblSale.SlTaxRate2+tblSale.SlTaxRate3)"
                 gCn.CommitTrans
                 
                 
                 '--Credit Note Sales
                 gCn.BeginTrans
                 x1 = " tblIntSale.VType = 'PX'  and (tblIntSale.VDt)>= ('" & Format(dtpFrDt.Value, "yyyy/mm/dd") & "') and (tblIntSale.VDt) <= ('" & Format(dtpToDt.Value, "yyyy/mm/dd") & "') and len(tblmastaccount.AcGSTIN) > 3 and tblIntSale.Vfirm = '" & gCCode & "'"  '
                 gCn.Execute " Insert Into TmpGentbl (Nar1,Nar2,Nar3,Amt1,Amt2,Amt3,Amt4,Amt5,Amt6,Amt7,UserName)" _
                                & " select tblMastItem.ItHsn2,'Credit Note','NOS-NUMBERS',0,sum(tblIntSale.VAmt),sum(tblIntSale.Interest),sum(tblIntSale.IgstAmt),sum(tblIntSale.CgstAmt),sum(tblIntSale.SgstAmt),tblIntSale.CgstRt+tblIntSale.SgstRt+tblIntSale.IgstRt,'" & gUserName & "'" _
                                & " From tblIntSale,tblmastaccount,tblMastState,tblMastItem where tblIntSale.VCrAcCode = tblMastAccount.Accode  and left(tblMastAccount.AcGSTIN,2) = tblMastState.StateCd and tblIntSale.ItCode = tblMastItem.ItCode and " & x1 & " " _
                                & " group by tblMastItem.ItHsn2,tblIntSale.VCtrNo,tblIntSale.CgstRt+tblIntSale.SgstRt+tblIntSale.IgstRt"
                 gCn.CommitTrans
                 
                 
                 
                 If GProcRstOpen(tmp1, "select * from TmpGentbl where TmpGentbl.UserName ='" & gUserName & "'", "R", gCn) > 0 Then

                    Set sheet = workbook.Sheets.Item("HSN(b2b)")
                    If (sheet Is Nothing) Then
                        MsgBox "Error Sales Sheet"
                    End If
                    a = 5
                    With tmp1
                         .MoveFirst
                         Do While Not .EOF
                            sheet.cells(a, 1) = !Nar1   '-- HSN Code
                            sheet.cells(a, 2) = !Nar2   '--Description
                            sheet.cells(a, 3) = !Nar3   '-- Unit
                            sheet.cells(a, 4) = !Amt1   '-- Qty
                            sheet.cells(a, 5) = !Amt2   '-- Total Value
                            sheet.cells(a, 6) = !Amt7   '-- Tax Rate
                            sheet.cells(a, 7) = !Amt3   '-- Taxable
                            sheet.cells(a, 8) = !Amt4   '-- I amt
                            sheet.cells(a, 9) = !Amt5   '-- C Amt
                            sheet.cells(a, 10) = !Amt6  '-- S Amt
                            sheet.cells(a, 11) = 0      '-- Cess
                            a = a + 1
     '                       prgBar.Value = prgBar.Value + 1
                           .MoveNext
                         Loop
                    End With
                 End If
                 workbook.Close SaveChanges:=True
                 Set xlApp = CreateObject("Excel.Application")
                 xlApp.Visible = True
                 Set xlWB = xlApp.Workbooks.Open("" & flnm & "")
                 excel_app.Quit
                 MsgBox "Compileted"
                 Exit Sub
        Case "GSTR - 9"
             GProcPurpareGSTR9 dtpFrDt.Value, dtpToDt.Value, OptDetail.Value, OptShort.Value
             Exit Sub
        Case "Monthwise GST Summery"
            ShowRepo = True
             GProcPurpareGSTSummery dtpFrDt.Value, dtpToDt.Value, OptShort.Value
             x1 = "{TmpGentbl.UserName}='" & gUserName & "' "
             .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt, "dd/MM/yyyy") & "'"
             .Formulas(7) = "wCompGstin = '" & gCGStin & "'"
             If OptDetail.Value = True Then
                .Formulas(5) = "wRptHd = 'Report On : Bill Date'"
             ElseIf OptShort.Value = True Then
                .Formulas(5) = "wRptHd = 'Report On : Received Date'"
             End If
             If OptShort.Value = True Then
                .Formulas(6) = "wRptNm = 'GST Details'"
             Else
                .Formulas(6) = "wRptNm = 'Monthwise GST Details'"
             End If
             .SelectionFormula = x1
             .ReportFileName = gReportPath & "rptMonthwiseGSTSumm.rpt"
        End Select
     If ShowRepo = True Then
       .Action = 1
       .PageZoom (120)
     End If
End With
End Sub
