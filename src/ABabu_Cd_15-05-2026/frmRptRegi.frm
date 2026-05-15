VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmRptRegi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Register"
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13245
   Icon            =   "frmRptRegi.frx":0000
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
      Picture         =   "frmRptRegi.frx":058A
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
      Top             =   4560
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
      Picture         =   "frmRptRegi.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   6885
      Width           =   1185
   End
   Begin VB.CommandButton cmdCrystal 
      Height          =   480
      Index           =   1
      Left            =   4725
      Picture         =   "frmRptRegi.frx":0E0E
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
      Format          =   116916227
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
      Format          =   116916227
      CurrentDate     =   38050
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Bindings        =   "frmRptRegi.frx":1478
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
      Left            =   10800
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
      Format          =   116916227
      CurrentDate     =   38050
   End
   Begin MSComCtl2.DTPicker dtpBillsToDt 
      Height          =   360
      Left            =   10800
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
      Format          =   116916227
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
      Left            =   8340
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
      Left            =   8280
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
      Width           =   8115
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
Attribute VB_Name = "frmRptRegi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim wNarrCode As Long, wMillCode As Long, wPartyCode As Long, wBrkCode As Long, wItemCode As Long, wPtyGrpCode As Long, wGodCd As Long
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
Dim Y As String
Dim G_Rs As Recordset
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



If gSelectedMenu = "Despatch Goods Details" Then
   dtpFrDt.Value = dtpFrDt.Value
   cmdCrystal(2).Visible = True
End If
If gSelectedMenu = "Daily Stock Statement" Then
   gDepotConsiMillHelp = True
End If
Me.Caption = gSelectedMenu
Select Case gSelectedMenu
''''==== Outstanding Report
Case "Outstanding Partywise", "Outstanding Brokerwise", "Outstanding Invoicewise"
     txtTemp(2).Visible = True 'party
     txtTemp(3).Visible = True 'broker
     txtTemp(0).Visible = True 'mill
     txtTemp(2).Enabled = True 'party
     txtTemp(3).Enabled = True 'broker
     txtTemp(0).Enabled = True 'mill
     
     txtCode(0).Visible = True
     txtCode(0).Enabled = True
     txtCode(2).Visible = True
     txtCode(2).Enabled = True
     txtCode(3).Visible = True
     txtCode(3).Enabled = True
     
     OptDetail.Visible = True
     OptDetail.Enabled = True
     OptDetail.Caption = "All O/s Bills "
     OptShort.Visible = True
     OptShort.Enabled = True
     OptShort.Caption = "Only Due Bill"
     opsummary.Visible = True
     opsummary.Enabled = True
     opsummary.Caption = "O/s Bills In Date Range"
     dtpFrDt.Visible = False
     Label1.Visible = False
     lblParty.Visible = True
     Label3.Visible = True
     lblMillName.Visible = True
     chkBox4.Visible = True
     chkBox4.Enabled = True
     chkBox4.Caption = "Only 2nd Day O/s"
     chkBox3.Visible = True
     chkBox3.Enabled = True
     chkBox3.Caption = "Payment To Mill O/S"
     chkBox6.Visible = True
     chkBox6.Enabled = True
     chkBox6.Caption = "TDS O/s"
     If gSelectedMenu = "Outstanding Invoicewise" Then
        ChkConsolidate.Visible = False
        ChkConsolidate.Enabled = False
        ChkConsolidate.Value = 0
        chkNewPage.Visible = False
        chkNewPage.Enabled = False
     Else
        ChkConsolidate.Visible = True
        ChkConsolidate.Enabled = True
        ChkConsolidate.Value = 0
        chkNewPage.Visible = True
        chkNewPage.Enabled = True
     End If
     If gSelectedMenu = "Outstanding Brokerwise" Then
        chkBox5.Visible = True
        chkBox5.Enabled = True
        chkBox5.Caption = "TCS O/s For Db.Nt."
     End If
Case "Ledger And Oustanding Sale Bal Check"
     dtpFrDt.Visible = False
     Label1.Visible = False
    ChkConsolidate.Visible = True
    ChkConsolidate.Enabled = True
Case "Trade Purchase Register GST", "Other Purchase Register GST"
     txtTemp(2).Visible = True 'party
     txtTemp(3).Visible = True 'broker
     txtTemp(0).Visible = True 'mill
     txtTemp(2).Enabled = True 'party
     txtTemp(3).Enabled = True 'broker
     txtTemp(0).Enabled = True 'mill
     lblMillName.Caption = "Pur. A/c     :"
     txtCode(0).Visible = True
     txtCode(0).Enabled = True
     txtCode(2).Visible = True
     txtCode(2).Enabled = True
     txtCode(3).Visible = True
     txtCode(3).Enabled = True
     
     lblParty.Visible = True
     Label3.Visible = True
     lblMillName.Visible = True
     ChkConsolidate.Visible = True
     ChkConsolidate.Value = 0
     ChkConsolidate.Enabled = True
     ChkConsolidate.Caption = "Register VAT"
     chkNewPage.Visible = True
     chkNewPage.Enabled = True
     chkNewPage.Caption = "Product Summery ?"
     chkNewPage.Value = 1
     chkBox3.Enabled = True
     chkBox3.Visible = True
     chkBox3.Caption = "Register Exempt."
     chkBox4.Visible = True
     chkBox4.Enabled = True
     chkBox4.Caption = "Partywise Summery"
     chkBox4.Value = 1
Case "Consignment/Depot Purchase Register", "SIT Purchase Register", "Other Sales", "Commission And Brokerage Sales"
     txtTemp(2).Visible = True 'party
     txtTemp(3).Visible = True 'broker
     txtTemp(0).Visible = True 'mill
     txtTemp(2).Enabled = True 'party
     txtTemp(3).Enabled = True 'broker
     txtTemp(0).Enabled = True 'mill
     txtCode(0).Visible = True
     txtCode(0).Enabled = True
     txtCode(2).Visible = True
     txtCode(2).Enabled = True
     txtCode(3).Visible = True
     txtCode(3).Enabled = True
     lblParty.Visible = True
     Label3.Visible = True
     lblMillName.Visible = True
     ChkConsolidate.Visible = True
     ChkConsolidate.Value = 0
     ChkConsolidate.Enabled = True
     ChkConsolidate.Caption = "Register"
     chkNewPage.Visible = True
     chkNewPage.Enabled = True
     chkNewPage.Caption = "Product Summery ?"
     If gSelectedMenu = "Consignment/Depot Purchase Register" Then
        chkBox3.Visible = True
        chkBox3.Enabled = True
        chkBox3.Caption = "Mill Summery"
        OptDetail.Visible = True
        OptDetail.Enabled = True
        OptDetail.Caption = "On Recd. Date "
        OptShort.Visible = True
        OptShort.Enabled = True
        OptShort.Caption = "On Desp. Date"
        OptShort.Value = True
     End If
     chkNewPage.Value = 1
Case "Trade Sale Register", "Consignment/Depot Sale Register", "Mill Bill Register"
     txtTemp(2).Visible = True 'party
     txtTemp(3).Visible = True 'broker
     txtTemp(0).Visible = True 'mill
     txtTemp(2).Enabled = True 'party
     txtTemp(3).Enabled = True 'broker
     txtTemp(0).Enabled = True 'mill
     txtCode(0).Visible = True
     txtCode(0).Enabled = True
     txtCode(2).Visible = True
     txtCode(2).Enabled = True
     txtCode(3).Visible = True
     txtCode(3).Enabled = True
     lblParty.Visible = True
     Label3.Visible = True
     lblMillName.Visible = True
     chkNewPage.Visible = True
     chkNewPage.Enabled = True
     chkNewPage.Caption = "Product Summery ?"
     chkNewPage.Value = 1
     If gSelectedMenu = "Consignment/Depot Sale Register" Then
        OptDetail.Visible = True
        OptDetail.Enabled = True
        OptDetail.Value = True
        OptShort.Visible = True
        OptShort.Enabled = True
        OptShort.Caption = "Datewise"
        opsummary.Visible = True
        opsummary.Enabled = True
        chkBox4.Visible = True
        chkBox4.Enabled = True
        chkBox4.Caption = "Only 2nd Day Bills"
        chkBox3.Visible = True
        chkBox3.Enabled = True
        chkBox3.Caption = "Only Regular Bills"
        txtTemp(1).Visible = True
        txtTemp(1).Enabled = True
        lblTaxType.Visible = True
        lblTaxType.Caption = "MailId :"
        cmdCrystal(2).Visible = True
        cmdCrystal(2).Enabled = True
     End If
     If gSelectedMenu = "Mill Bill Register" Then
        OptDetail.Visible = True
        OptDetail.Enabled = True
        OptDetail.Caption = "All Bills"
        OptDetail.Value = True
        OptShort.Visible = True
        OptShort.Enabled = True
        OptShort.Caption = "Without Direct Payment"
        opsummary.Visible = True
        opsummary.Enabled = True
        opsummary.Caption = "Olny Direct Payment"
     End If
     
Case "Despatch Goods Details", "Party Booking Register", "Pending Party Booking Register", "Party Booking Summery", "Trade Purchase Booking", "Purchase Booking Summery", "Party Booking Vs Dispatch Register"
     txtTemp(2).Visible = True 'party
     txtTemp(3).Visible = True 'broker
     txtTemp(0).Visible = True 'mill
     txtTemp(2).Enabled = True 'party
     txtTemp(3).Enabled = True 'broker
     txtTemp(0).Enabled = True 'mill
     txtCode(0).Visible = True
     txtCode(0).Enabled = True
     txtCode(2).Visible = True
     txtCode(2).Enabled = True
     txtCode(3).Visible = True
     txtCode(3).Enabled = True
     lblParty.Visible = True
     Label3.Visible = True
     lblMillName.Visible = True
     If gSelectedMenu = "Despatch Goods Details" Then
        OptDetail.Visible = True
        OptDetail.Enabled = True
        opsummary.Visible = True
        opsummary.Enabled = True
        ChkConsolidate.Visible = True
        ChkConsolidate.Enabled = True
     End If
     If gSelectedMenu = "Party Booking Register" Or gSelectedMenu = "Pending Party Booking Register" Then cmdCrystal(2).Visible = True
     If gSelectedMenu = "Trade Purchase Booking" Then cmdCrystal(2).Visible = True
     If gSelectedMenu = "Party Booking Register" Or gSelectedMenu = "Pending Party Booking Register" Or gSelectedMenu = "Trade Purchase Booking" Then
        txtTemp(5).Visible = True
        txtTemp(5).Enabled = True
     End If
     If gSelectedMenu = "Party Booking Summery" Or gSelectedMenu = "Party Booking Register" Or gSelectedMenu = "Pending Party Booking Register" Or gSelectedMenu = "Trade Purchase Booking" Or gSelectedMenu = "Purchase Booking Summery" Then
        ChkConsolidate.Value = 1
        ChkConsolidate.Visible = True
        ChkConsolidate.Enabled = True
      End If
      If gSelectedMenu = "Party Booking Vs Dispatch Register" Then
         Label8.Visible = True
         Label8.Caption = "Booking No : "
         txt(0).Visible = True
         txt(0).Enabled = True
         ChkConsolidate.Visible = True
         ChkConsolidate.Enabled = True
      End If
Case "Stock Statement", "Daily Stock Statement", "Sales Summery", "Gate Pass Wise Sales", "Stock Statement Godownwsie"
     txtTemp(0).Visible = True 'mill
     txtTemp(0).Enabled = True 'mill
     txtCode(0).Visible = True
     txtCode(0).Enabled = True
     txtTemp(5).Visible = True
     txtTemp(5).Enabled = True
     Label5.Visible = True
     lblMillName.Visible = True
     ChkConsolidate.Visible = True
     ChkConsolidate.Enabled = True
     dtpFrDt.Value = Format(Now, "dd/MM/yyyy")
     dtpToDt.Value = Format(Now, "dd/MM/yyyy")
     If gSelectedMenu = "Stock Statement" Then
        chkBox3.Visible = True
        chkBox3.Enabled = True
        chkBox3.Caption = "Lotwise Stock"
        chkBox3.Value = 1
     End If
     If gSelectedMenu = "Stock Statement Godownwsie" Then
        txtTemp(4).Visible = True
        txtTemp(4).Enabled = True
        Label4.Visible = True
     End If
Case "Depot Sale Register GST"
     txtCode(0).Visible = True
     txtCode(0).Enabled = True
     txtTemp(0).Visible = True 'mill
     txtTemp(0).Enabled = True
     lblMillName.Visible = True
Case "Trade Purchase Register GST", "Purchase Return Register", "Insurance Report - Purchase"
     txtTemp(2).Visible = True 'party
     txtTemp(2).Enabled = True 'party
     txtCode(2).Visible = True
     txtCode(2).Enabled = True
     lblParty.Visible = True
     If gSelectedMenu = "Insurance Report - Purchase" Then
        txtTemp(3).Visible = False
        txtTemp(3).Enabled = False
        txtCode(3).Visible = False
        txtCode(3).Enabled = False
        txtTemp(0).Visible = False
        txtTemp(0).Enabled = False
        txtCode(0).Visible = False
        txtCode(0).Enabled = False
        ChkConsolidate.Visible = False
        chkBox3.Visible = False
        chkNewPage.Visible = False
     End If
Case "Daily Report", "Daily Report Mail", "Lot No wise Transit Detail", "Daily Report ( Special )"
      dtpBillsFrDt.Visible = True
      dtpBillsFrDt.Enabled = True
      dtpBillsFrDt = Format(DateAdd("YYYY", -1, gCYSDate), "dd/MM/yyyy")
      Label12.Visible = True
      Label12.Caption = "Pur.Book.Pend.From  :"
      dtpFrDt.Visible = False
      Label1.Visible = False
      txt(0).Visible = True
      txt(0).Enabled = True
      txt(1).Visible = True
      txt(1).Enabled = True
      Label8.Visible = True
      Label10.Visible = True
      Y = "Select Min(Acsrno) from tblMastAccount where acsrno > 0"
      GProcRstOpen G_Rs, Y, "R", gCn
      If G_Rs.Fields(0) = Null Or IsNumeric(G_Rs.Fields(0)) = False Then
         txt(0) = 0
      Else
         txt(0) = CLng(G_Rs.Fields(0))
      End If
      Y = "Select Max(Acsrno) from tblMastAccount where acsrno > 0"
      GProcRstOpen G_Rs, Y, "R", gCn
      If G_Rs.Fields(0) = Null Or IsNumeric(G_Rs.Fields(0)) = False Then
         txt(1) = 0
      Else
         txt(1) = CLng(G_Rs.Fields(0))
      End If
      Label8.Caption = "From Sr No :"
      Label10.Caption = "To   Sr No :"
      G_Rs.Close
      
      If gSelectedMenu = "Daily Report Mail" Then
         txtCode(0).Visible = True
         txtCode(0).Enabled = True
         txtTemp(0).Visible = True
         txtTemp(0).Enabled = True
         lblMillName.Visible = True
         txtTemp(1).Visible = True
         txtTemp(1).Enabled = True
         lblTaxType.Visible = True
         lblTaxType.Caption = "Mill Group :"
         cmdCrystal(2).Visible = True
         cmdCrystal(2).Enabled = True
      End If
      
Case "Trade Sale Register GST", "Sales Return Register"
     txtTemp(2).Visible = True 'party
     txtTemp(0).Visible = True 'mill
     txtTemp(2).Enabled = True  'party
     txtTemp(0).Enabled = True  'mill
     txtCode(0).Visible = True
     txtCode(0).Enabled = True
     txtCode(2).Visible = True
     txtCode(2).Enabled = True
     lblParty.Visible = True
     lblMillName.Visible = True
Case "Internal Voucher Register", "Credit Note (Sales) Register", "Debit Note (Sales) Register", "Credit Note (Purchase) Register", "Debit Note (Purchase) Register", "Late Pay Debit Note Register", "Late Pay Debit Note Register (Mill Bill)", "Non-GST Outward Sales"
     txtTemp(2).Visible = True 'party
     txtTemp(2).Enabled = True  'party
     lblParty.Visible = True
     txtCode(2).Visible = True
     txtCode(2).Enabled = True
     If gSelectedMenu = "Internal Voucher Register" Then
         OptDetail.Visible = True
         OptDetail.Enabled = True
         OptDetail.Caption = "Broker JV Reg."
         OptShort.Visible = True
         OptShort.Enabled = True
         OptShort.Caption = "Other JV Reg."
         opsummary.Visible = True
         opsummary.Enabled = True
         opsummary.Caption = "All JV Reg."
         opsummary.Value = True
     End If
     If gSelectedMenu = "Credit Note (Purchase) Register" Or gSelectedMenu = "Debit Note (Purchase) Register" Or gSelectedMenu = "Credit Note (Sales) Register" Then
        OptDetail.Value = True
        OptDetail.Visible = True
        OptDetail.Enabled = True
        OptDetail.Caption = "With GST"
        OptShort.Visible = True
        OptShort.Enabled = True
        OptShort.Caption = "Without GST"
     End If
Case "Delivery Register", "Pending Despatch"
     ChkConsolidate.Visible = True
     ChkConsolidate.Enabled = True
     lblParty.Visible = True
     txtCode(2).Visible = True
     txtCode(2).Enabled = True
     txtTemp(2).Visible = True  '--Party
     txtTemp(2).Enabled = True
     
     lblMillName.Visible = True
     txtCode(0).Visible = True
     txtCode(0).Enabled = True
     txtTemp(0).Visible = True  '-- Mill
     txtTemp(0).Enabled = True
     
     Label5.Visible = True
     txtTemp(5).Visible = True  '--Item
     txtTemp(5).Enabled = True

     Label9.Visible = True
     txtTemp(6).Visible = True  '--Sizer
     txtTemp(6).Enabled = True
Case "Planed Payment List"
     Label1.Visible = False
     dtpFrDt.Visible = False
     lblParty.Visible = False
     txtCode(2).Visible = False
     txtTemp(2).Visible = False
     Label3.Visible = False
     txtCode(3).Visible = False
     txtTemp(3).Visible = False
     lblMillName.Visible = False
     txtCode(0).Visible = False
     txtTemp(0).Visible = False
     Label5.Visible = False
     txtTemp(5).Visible = False
     lblTaxType.Visible = False
     txtTemp(1).Visible = False
     Label9.Visible = False
     txtTemp(6).Visible = False
     Label4.Visible = False
     txtTemp(4).Visible = False
     ChkConsolidate.Visible = False
     chkBox3.Visible = False
     chkNewPage.Visible = False
     chkBox4.Visible = False
     chkBox5.Visible = False
     chkBox6.Visible = False
     OptDetail.Visible = False
     OptShort.Visible = False
     opsummary.Visible = False
     Label8.Visible = True
     Label8.Caption = "Day Of Week :"
     txt(0).Visible = True
     txt(0).Enabled = True
     Label2.Visible = False
     dtpToDt.Visible = False
End Select
 '-- Set InActivate Ctrl Back Color
GProcSetInActivateCtrlBackClr Me
End Sub
Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub opsummary_Click()
If gSelectedMenu = "Outstanding Brokerwise" Or gSelectedMenu = "Outstanding Invoicewise" Then
    If opsummary.Value = True Then
        Label12.Visible = True
        Label13.Visible = True
        dtpBillsFrDt.Visible = True
        dtpBillsToDt.Visible = True
        dtpBillsFrDt.Enabled = True
        dtpBillsToDt.Enabled = True
    Else
        Label12.Visible = False
        Label13.Visible = False
        dtpBillsFrDt.Visible = False
        dtpBillsToDt.Visible = False
    End If
End If
End Sub

Private Sub OptDetail_Click()
If gSelectedMenu = "Outstanding Brokerwise" Or gSelectedMenu = "Outstanding Invoicewise" Then
    Label12.Visible = False
    Label13.Visible = False
    dtpBillsFrDt.Visible = False
    dtpBillsToDt.Visible = False
End If
End Sub

Private Sub OptShort_Click()
If gSelectedMenu = "Outstanding Brokerwise" Or gSelectedMenu = "Outstanding Invoicewise" Then
    Label12.Visible = False
    Label13.Visible = False
    dtpBillsFrDt.Visible = False
    dtpBillsToDt.Visible = False
End If
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
     Case 4 '--Godown
        gClsSearch.SearchMultiField "tblMastNarration", "Narration,NarrCode", Array("Godown", ""), Array(6000, 0), " NarrType='G' ", Chr(KeyAscii), "Narration", txtTemp(Index).Left + Me.Left, txtTemp(Index).Top + Me.Top + 1000
        KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" Then
           wGodCd = gClsSearch.SearchMultiRetCol(1) '-Godown Cd
           txtTemp(Index) = gClsSearch.SearchMultiRetCol(0) '-Godown Nm
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
    Case 1 '----Party Group
         gClsSearch.SearchMultiField "tblMastCompGroup", "CGName", Array("Company Group"), Array(txtTemp(Index).Width), " ", Chr(KeyAscii), "CGName", txtTemp(Index).Left + Me.Left, txtTemp(Index).Top + Me.Top + 1000, True
         KeyAscii = 0
        If gClsSearch.SearchMultiRetCol(0) <> "" Then
            txtTemp(Index) = gClsSearch.SearchMultiRetCol(0)
            wPtyGrpCode = GProcGetColumnValue("tblMastCompGroup", "CgName", txtTemp(1), "S", "CgCode", "N")
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
'''        '===== Outstanding Report
        Case "Outstanding Partywise", "Outstanding Brokerwise", "Outstanding Invoicewise"
              ShowRepo = True
                If txtTemp(2) <> "" Then
                   wPartyCode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(2), "S", "AcCode", "N")
                Else
                   wPartyCode = 0
                End If
                If txtTemp(3) <> "" Then
                   wBrkCode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(3), "S", "AcCode", "N")
                Else
                   wBrkCode = 0
                End If
                If txtTemp(0) <> "" Then
                   wMillCode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(0), "S", "AcCode", "N")
                Else
                   wMillCode = 0
                End If
                If chkBox3.Value = 1 Then
                   gCmd.CommandText = "PrcPrepareOutStangingSaleDirectPayment"
                Else
                   gCmd.CommandText = "PrcPrepareOutStangingSale"
                End If
                gCmd.Parameters.Refresh
'                gCmd.Parameters("@From_dt") = Format(FromDt, "yyyy/mm/dd")
                gCmd.Parameters("@To_dt") = Format(dtpToDt, "yyyy/mm/dd")
                gCmd.Parameters("@UCode") = gUserName
                gCmd.Parameters("@IntDbAc") = gLatePayIntAcCodeRecGST
                gCmd.Parameters("@IntDbAcWoGST") = gLPIntAcCodeWithoutGST
                gCmd.Parameters("@IntTDSAc") = gTDSAcCodeRec
                gCmd.Parameters("@IntCrAc") = gLatePayIntAcCodePay
                If chkBox4.Value = 1 Then gCmd.Parameters("@Is2ndDay") = 1
                If ChkConsolidate.Value = 1 Then
                   gCmd.Parameters("@VFirm") = ""
                Else
                   gCmd.Parameters("@VFirm") = Trim(gCCode)
                End If
                If OptShort.Value = True Then
                   gCmd.Parameters("@DueOn") = "D"
                ElseIf OptDetail.Value = True Then
                   gCmd.Parameters("@DueOn") = "B"
                ElseIf opsummary.Value = True Then
                   gCmd.Parameters("@DueOn") = "X"
                End If
'                If chkBox3.Value = 1 Then gCmd.Parameters("@IsDirectPay") = 1
                gCmd.Parameters("@X_dt1") = Format(dtpBillsFrDt, "yyyy/mm/dd")
                gCmd.Parameters("@X_dt2") = Format(dtpBillsToDt, "yyyy/mm/dd")
                If wPartyCode <> 0 Then gCmd.Parameters("@Accode") = wPartyCode
                If wBrkCode <> 0 Then gCmd.Parameters("@Accode_Br") = wBrkCode
                If wMillCode <> 0 Then gCmd.Parameters("@MillCode") = wMillCode
                gCmd.Execute
                If chkBox6.Value = 1 Then
                    gCn.BeginTrans
                    gCn.Execute "delete from tmpGenTbl where username='" & gUserName & "'  and Amt6 =0 "
                    gCn.CommitTrans
                End If
                x1 = "{TmpGentbl.UserName}='" & gUserName & "' "
               .SelectionFormula = x1
               .Formulas(4) = "wFrToDt = 'Up To ' & '" & Format(dtpToDt.Value, "dd/MM/yyyy") & "'"
               If ChkConsolidate.Value = 1 Then
                  .Formulas(5) = "wAllFirm = 'All Firm '"
               Else
                  .Formulas(5) = "wAllFirm = 'Single Firm '"
               End If
               .Formulas(6) = "wNewPage = " & chkNewPage.Value & " "
               If chkBox4.Value = 1 Then .Formulas(7) = "w2ndDayOs = '2nd Day Bills O/s'"
               If chkBox3.Value = 1 Then .Formulas(8) = "wDirectPaymentOs = 'Direct Payment To Mill O/s'"
               
               If gSelectedMenu = "Outstanding Brokerwise" Then
                 If chkBox5.Value = 1 Then
                    gCn.BeginTrans
                    gCn.Execute "delete from tmpGenTbl where username='" & gUserName & "'  and Amt8 = 0 "
                    gCn.CommitTrans
                    gCn.BeginTrans
                    gCn.Execute "delete from tmpGenTbl where username='" & gUserName & "'  and Amt1+Amt3-Amt2 <=0 "
                    gCn.CommitTrans
                    .ReportFileName = gReportPath & "rptTcsOsDbNt.rpt"
                 Else
                    .ReportFileName = gReportPath & "rptOutStandingSaleBrokerwise.rpt"
                 End If
               ElseIf gSelectedMenu = "Outstanding Partywise" Then
                  If ChkConsolidate.Value = 1 Then
                     .ReportFileName = gReportPath & "rptOutStandingSalePartywise.rpt"
                  Else
                     .ReportFileName = gReportPath & "rptOutStandingSalePartyOneFirm.rpt"
                  End If
               ElseIf gSelectedMenu = "Outstanding Invoicewise" Then
                     .ReportFileName = gReportPath & "rptOutStandingSaleInvoicewise.rpt"
               End If
        Case "Ledger And Oustanding Sale Bal Check"
              ShowRepo = True
                gCmd.CommandText = "PrcPrepareOutStangingSale"
                gCmd.Parameters.Refresh
'                gCmd.Parameters("@From_dt") = Format(FromDt, "yyyy/mm/dd")
                gCmd.Parameters("@To_dt") = Format(dtpToDt, "yyyy/mm/dd")
                gCmd.Parameters("@UCode") = gUserName
                gCmd.Parameters("@IntDbAc") = gLatePayIntAcCodeRecGST
                gCmd.Parameters("@IntDbAcWoGST") = gLPIntAcCodeWithoutGST
                gCmd.Parameters("@IntTDSAc") = gTDSAcCodeRec
                gCmd.Parameters("@IntCrAc") = gLatePayIntAcCodePay
                If ChkConsolidate.Value = 1 Then
                   gCmd.Parameters("@VFirm") = ""
                Else
                   gCmd.Parameters("@VFirm") = Trim(gCCode)
                End If
                gCmd.Parameters("@X_dt1") = Format(dtpToDt, "yyyy/mm/dd")
                gCmd.Parameters("@X_dt2") = Format(dtpToDt, "yyyy/mm/dd")
                gCmd.Execute
                 gCmd.CommandText = "PrcPrepareClBalance"
                 gCmd.Parameters.Refresh
                 If ChkConsolidate.Value = 1 Then
                    gCmd.Parameters("@VFirm") = ""
                 Else
                    gCmd.Parameters("@VFirm") = Trim(gCCode)
                 End If
                 gCmd.Parameters("@Fr_dt") = Format(dtpFrDt.Value, "yyyy/mm/dd")
                 gCmd.Parameters("@To_dt") = Format(dtpToDt.Value, "yyyy/mm/dd")
                 gCmd.Parameters("@FyStart_dt") = Format(gCYSDate, "yyyy/mm/dd")
                 gCmd.Parameters("@UCode") = gUserName
'                 gCmd.Parameters("@VFirm") = Trim(gCCode)
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
                 X = "Select  'A',TmpGentbl.Vfirm,TmpGentbl.VFirmName,TmpGentbl.PartyCode,TmpGentbl.PartyName,sum(TmpGentbl.Amt1+TmpGentbl.Amt3-TmpGentbl.Amt2), '" & gUserName & "' from TmpGentbl  where  TmpGentbl.UserName = '" & gUserName & "' " _
                        & " Group By TmpGentbl.Vfirm,TmpGentbl.VFirmName,TmpGentbl.PartyCode,TmpGentbl.PartyName"
                 x1 = "Select  'B',TmpClosingBalance.Vfirm,TmpClosingBalance.VFirmName,TmpClosingBalance.AcOurCode  ,TmpClosingBalance.AcName,TmpClosingBalance.ClBal_Dr,TmpClosingBalance.ClBal_Cr, '" & gUserName & "' from TmpClosingBalance  where  TmpClosingBalance.UserName = '" & gUserName & "'   and  TmpClosingBalance.agcode = 90017 "
             
                gCn.BeginTrans
                gCn.Execute "delete from tmpGenTbl2 where username='" & gUserName & "'"
                gCn.Execute " Insert Into TmpGentbl2 (Vtype,Vfirm,VFirmName,PartyCode,PartyName,Amt3,UserName) " & X
                gCn.Execute " Insert Into TmpGentbl2 (Vtype,Vfirm,VFirmName,PartyCode,PartyName,Amt1,Amt2,UserName) " & x1
                gCn.CommitTrans
             
                X = "Select 'C',Vfirm,VFirmName,PartyCode,PartyName,sum(amt1),sum(amt2),sum(amt3),'" & gUserName & "'  from TmpGentbl2 where TmpGentbl2.UserName = '" & gUserName & "' group by Vfirm,VFirmName,PartyCode,PartyName"
                
                gCn.BeginTrans
                gCn.Execute " Insert Into TmpGentbl2 (Vtype,Vfirm,VFirmName,PartyCode,PartyName,Amt1,Amt2,Amt3,UserName) " & X
                gCn.Execute "delete from tmpGenTbl2 where username='" & gUserName & "'  and vtype in ('A','B')  "
                gCn.Execute "delete from tmpGenTbl2 where username='" & gUserName & "'  and vtype = 'C' and Amt1+Amt2-abs(amt3) = 0 "
                gCn.CommitTrans
                x1 = "{TmpGentbl2.UserName}='" & gUserName & "' "
                If ChkConsolidate.Value = 0 Then
                   x1 = x1 + " and {TmpGentbl2.Vfirm} = '" & gCCode & "' "
                End If
               .SelectionFormula = x1
               .Formulas(4) = "wFrToDt = 'Up To ' & '" & Format(dtpToDt.Value, "dd/MM/yyyy") & "'"
               If ChkConsolidate.Value = 1 Then
                  .Formulas(5) = "wAllFirm = 'All Firm '"
               Else
                  .Formulas(5) = "wAllFirm = 'Single Firm '"
               End If
               .ReportFileName = gReportPath & "rptOutStandingSaleLegerChk.rpt"
        Case "Trade Purchase Register"
              ShowRepo = True
                If txtTemp(2) <> "" Then
                   wPartyCode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(2), "S", "AcCode", "N")
                Else
                   wPartyCode = 0
                End If
                If txtTemp(3) <> "" Then
                   wBrkCode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(3), "S", "AcCode", "N")
                Else
                   wBrkCode = 0
                End If
                If txtTemp(0) <> "" Then
                   wMillCode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(0), "S", "AcCode", "N")
                Else
                   wMillCode = 0
                End If
                gCmd.CommandText = "PrcPreparePurchaseDetail"
                gCmd.Parameters.Refresh
                gCmd.Parameters("@Fr_dt") = Format(dtpFrDt, "yyyy/mm/dd")
                gCmd.Parameters("@To_dt") = Format(dtpToDt, "yyyy/mm/dd")
                gCmd.Parameters("@UCode") = gUserName
                gCmd.Parameters("@VFirm") = Trim(gCCode)
                gCmd.Parameters("@Vtype") = "PY"
                
                If wPartyCode <> 0 Then gCmd.Parameters("@Accode") = wPartyCode
                If wBrkCode <> 0 Then gCmd.Parameters("@Accode_Br") = wBrkCode
                If wMillCode <> 0 Then gCmd.Parameters("@MillCode") = wMillCode
                gCmd.Execute
                x1 = "{TmpGentbl.UserName}='" & gUserName & "'"
                ABC = ""
                If ChkConsolidate.Value = 1 And chkBox3.Value = 0 Then 'register Vat
                   ABC = ABC + "A"
                   x1 = x1 + "and {TmpGentbl.Amt10} = 0 "
                   .Formulas(6) = "wVAT/Exempt = 'VAT Purchase' "
                End If
                If chkBox3.Value = 1 And ChkConsolidate.Value = 0 Then 'register Exempt
                   ABC = ABC + "A"
                   x1 = x1 + "and {TmpGentbl.Amt10} > 0 "
                   .Formulas(6) = "wVAT/Exempt = 'Exempt Purchase' "
                End If
                If ChkConsolidate.Value = 1 And chkBox3.Value = 1 Then 'register Vat+ Exempt
                   ABC = ABC + "A"
                   .Formulas(6) = "wVAT/Exempt = 'VAT and Exempt Purchase' "
                End If
                If chkNewPage.Value = 1 Then
                   ABC = ABC + "B"
                End If
                x1 = x1 + "and {TmpGentbl.Nar6} in ( '" & ABC & "') "
               .SelectionFormula = x1
               .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt, "dd/MM/yyyy") & "'"
               .ReportFileName = gReportPath & "rptRegisterPurchTradeDetail.rpt"
        Case "SIT Purchase Register"
              ShowRepo = True
                If txtTemp(2) <> "" Then
                   wPartyCode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(2), "S", "AcCode", "N")
                Else
                   wPartyCode = 0
                End If
                If txtTemp(3) <> "" Then
                   wBrkCode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(3), "S", "AcCode", "N")
                Else
                   wBrkCode = 0
                End If
                If txtTemp(0) <> "" Then
                   wMillCode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(0), "S", "AcCode", "N")
                Else
                   wMillCode = 0
                End If
                gCmd.CommandText = "PrcPreparePurchaseDetailSIT"
                gCmd.Parameters.Refresh
                gCmd.Parameters("@Fr_dt") = Format(dtpFrDt, "yyyy/mm/dd")
                gCmd.Parameters("@To_dt") = Format(dtpToDt, "yyyy/mm/dd")
                gCmd.Parameters("@UCode") = gUserName
                gCmd.Parameters("@VFirm") = Trim(gCCode)
                gCmd.Parameters("@Vtype") = "PT"
                
                If wPartyCode <> 0 Then gCmd.Parameters("@Accode") = wPartyCode
                If wBrkCode <> 0 Then gCmd.Parameters("@Accode_Br") = wBrkCode
                If wMillCode <> 0 Then gCmd.Parameters("@MillCode") = wMillCode
                gCmd.Execute
                x1 = "{TmpGentbl.UserName}='" & gUserName & "'"
                ABC = ""
                If ChkConsolidate.Value = 1 Then 'Inward register
                   ABC = ABC + "A"
                End If
                If chkNewPage.Value = 1 Then
                   ABC = ABC + "B"
                End If
                If chkBox3.Value = 1 Then
                   ABC = ABC + "C"
                End If
                x1 = x1 + "and {TmpGentbl.Nar6} in ( '" & ABC & "') "
               .SelectionFormula = x1
               .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt, "dd/MM/yyyy") & "'"
               .ReportFileName = gReportPath & "rptRegisterPurchSITDetail.rpt"
        Case "Consignment/Depot Purchase Register"
              ShowRepo = True
                If txtTemp(2) <> "" Then
                   wPartyCode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(2), "S", "AcCode", "N")
                Else
                   wPartyCode = 0
                End If
                If txtTemp(3) <> "" Then
                   wBrkCode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(3), "S", "AcCode", "N")
                Else
                   wBrkCode = 0
                End If
                If txtTemp(0) <> "" Then
                   wMillCode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(0), "S", "AcCode", "N")
                Else
                   wMillCode = 0
                End If
                gCmd.CommandText = "PrcPreparePurchaseDetail"
                gCmd.Parameters.Refresh
                gCmd.Parameters("@Fr_dt") = Format(dtpFrDt, "yyyy/mm/dd")
                gCmd.Parameters("@To_dt") = Format(dtpToDt, "yyyy/mm/dd")
                gCmd.Parameters("@UCode") = gUserName
                gCmd.Parameters("@VFirm") = Trim(gCCode)
                gCmd.Parameters("@Vtype") = "PI"
                If OptShort.Value = True Then
                   gCmd.Parameters("@OnDate") = "D"
                Else
                   gCmd.Parameters("@OnDate") = "R"
                End If
                If wPartyCode <> 0 Then gCmd.Parameters("@Accode") = wPartyCode
                If wBrkCode <> 0 Then gCmd.Parameters("@Accode_Br") = wBrkCode
                If wMillCode <> 0 Then gCmd.Parameters("@MillCode") = wMillCode
                gCmd.Execute
                x1 = "{TmpGentbl.UserName}='" & gUserName & "'"
                ABC = ""
                If ChkConsolidate.Value = 1 Then 'Inward register
                   ABC = ABC + "A"
                End If
                If chkNewPage.Value = 1 Then
                   ABC = ABC + "B"
                End If
                If chkBox3.Value = 1 Then
                   ABC = ABC + "C"
                End If
                x1 = x1 + " and {TmpGentbl.Nar6} in ( '" & ABC & "') "
               .SelectionFormula = x1
               .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt, "dd/MM/yyyy") & "'"
               If OptShort.Value = True Then
                  .Formulas(5) = "wOnDate = 'D'"
               Else
                  .Formulas(5) = "wOnDate = 'R'"
               End If
               .ReportFileName = gReportPath & "rptRegisterPurchInwardDetail.rpt"
        Case "Lot No wise Transit Detail"
             x1 = "select tblMastAccount.AcOurCode,tblMastAccount.AcName,tblMastItem.ItName,tblBooKingDesp.BkDespDate,BkDespTmpChr2,BkDespBag,'" & gUserName & "' " _
                  & " From tblBooKing, tblBooKingDesp, tblMastItem, tblMastAccount " _
                  & " where tblBooKing.vtype in ('OT','OD') and BkDespTmpChr2 <> '' and  BkDespTmpChr1 <> 'Y' and BkDespBag > 0 and " _
                  & " tblBooking.VNo = tblBooKingDesp.VNo and  tblBooking.VType = tblBooKingDesp.VType and tblBooking.VYear  = tblBooKingDesp.VYear and " _
                  & " tblBooKing.VFirm = tblBooKingDesp.VFirm And tblBooKing.BkSrNo = 1 " _
                  & " and tblBooKing.BkItCode = tblMastItem.ItCode and tblMastItem.ItMillCode = tblMastAccount.AcCode  "
             
             gCn.BeginTrans
             gCn.Execute "delete from tmpGenTbl where username='" & gUserName & "' "
             gCn.CommitTrans
             
             gCn.BeginTrans
             gCn.Execute "insert into tmpGentbl (MillCode,MillName,Nar1,Vdt,Nar2,Amt1,UserName) " & x1
             gCn.CommitTrans
             ShowRepo = True
             x1 = "{TmpGentbl.UserName}='" & gUserName & "' "
             .SelectionFormula = x1
             .Formulas(4) = "wFrToDt = 'Date : ' & '" & Format(dtpToDt.Value, "dd/MM/yyyy") & "' "
             .ReportFileName = gReportPath & "rptLotwiseTransit.rpt"
        Case "Daily Report", "Daily Report ( Special )"
              GProcDailyReport dtpFrDt.Value, dtpToDt.Value, Val(txt(0)), Val(txt(1)), dtpBillsFrDt.Value
              ShowRepo = True
                x1 = "{TmpGentbl2.UserName}='" & gUserName & "' "
               .SelectionFormula = x1
               .Formulas(4) = "wFrToDt = 'Date : ' & '" & Format(dtpToDt.Value, "dd/MM/yyyy") & "' "
                If gSelectedMenu = "Daily Report ( Special )" Then
                   .Formulas(5) = "wMillRt = 'Y'"
                Else
                   .Formulas(5) = "wMillRt = 'N'"
                End If
               .ReportFileName = gReportPath & "rptDailyReport.rpt"
        Case "Daily Report Mail"
              .Formulas(0) = ""
              .Formulas(1) = ""
              .Formulas(2) = ""
              .Formulas(3) = ""
              If txtTemp(0) <> "" Then
                 wMillCode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(0), "S", "AcCode", "N")
              Else
                 wMillCode = 0
              End If
              If txtTemp(1) <> "" Then
                 wPtyGrpCode = GProcGetColumnValue("tblMastCompGroup", "CgName", txtTemp(1), "S", "CgCode", "N")
              Else
                 wPtyGrpCode = 0
              End If
              If wMillCode = 0 And wPtyGrpCode = 0 Then
                MsgBox ("Please Select Mill Name")
                Exit Sub
              End If
              GProcDailyReportMail dtpFrDt.Value, dtpToDt.Value, wMillCode, wPtyGrpCode, dtpBillsFrDt.Value
              ShowRepo = True
              If cmdCrystal(2).Value = True Then
                 Set objCrystal = New CRAXDRT.Application
                 Tattach = gReportPath & "rptDailyReportMail.rpt"
                 Set objReport = objCrystal.OpenReport(Tattach, 1)
                 objReport.RecordSelectionFormula = "{TmpGentbl.UserName}='" & gUserName & "'"
                 If gCIBOffice = "I" Then
                    objReport.Database.LogOnServer "pdsodbc.dll", "Hitrix", "Vishnu", "sa", "123456"
                 Else
                    objReport.Database.LogOnServer "pdsodbc.dll", "Hitrix", "VishnuMum", "sa", "123456"
                 End If
                 For j = 1 To objReport.FormulaFields.Count
                     Select Case objReport.FormulaFields(j).Name
                     Case "{@wFrToDt}"
                          objReport.FormulaFields(j).text = "" & Chr(39) & "Date : " & dtpToDt & "" & Chr(39) & ""
                     End Select
                 Next j
                 If wMillCode > 0 Then
                    wMillMailId = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(0), "S", "AcEmail", "S")
                 End If
                 If txtTemp(1) <> "" Then
                    wMillMailId = GProcGetColumnValue("tblMastAccount", "AcCmpCode", "" & wPtyGrpCode & "", "N", "AcEmail", "S")
                 End If
                 
                 x1 = "Select vno from TmpGentbl where nar10 = 'A'"
                 RecYN = GProcRstOpen(tmp1, x1, "R")
                 If RecYN > 0 Then
                    If tmp1.Fields(0) > 0 Then
                       a = 1
                    Else
                       a = 0
                    End If
                 End If
'                 wMillMailId = "sanju.khamkar@gmail.com"
                 If wMillMailId = "" Or IsNull(wMillMailId) Then
                     MsgBox "Mail ID Not Found"
                     Exit Sub
                 End If
                 PtyMailId = wMillMailId
                 Mfile = CStr(App.Path) + "\MailPDF\" + gCCode + "-" + gUserName + "-" + "Daily_Report"
                 ExportReportToPDF objReport, Mfile & ".pdf", "foo"
                 Mfile = Mfile + ".pdf"
                 MailSub = "Daily Report"
                 MailStr = "Dear Sir, " + Chr(13) + Chr(10)
                 MailStr = MailStr + "Please find attached daily report." + Chr(13) + Chr(10) + Chr(13) + Chr(10)
'                 If a = 1 Then
'                    MailStr = MailStr + "required for Depot Work" & vbLf & vbLf
'                 Else
'                    MailStr = MailStr + "required for Direct Billing Work" & vbLf & vbLf
'                 End If
'                 MailStr = MailStr + "Thanks and Regards" & vbLf & vbLf
'                 MailStr = MailStr + "Vishnu Textile Corporation" & vbLf
                 gMailsendToMill = False
                 Call SendEmail1("" & Trim(gCMailId) & "", "" & PtyMailId & "", MailSub, MailStr, "" & PtyMailIdCc & "", "" & PtyMailIdBcc & "", Mfile)
'                 Kill Mid(Tattach, 1, Len(Tattach) - 4) & ".pdf"
                 Exit Sub
              Else
                x1 = "{TmpGentbl.UserName}='" & gUserName & "' "
               .SelectionFormula = x1
               .Formulas(4) = "wFrToDt = 'Date : ' & '" & Format(dtpToDt.Value, "dd/MM/yyyy") & "' "
               .ReportFileName = gReportPath & "rptDailyReportMail.rpt"
              End If
        Case "Despatch Goods Details"
              ShowRepo = True
                If txtTemp(2) <> "" Then
                   wPartyCode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(2), "S", "AcCode", "N")
                Else
                   wPartyCode = 0
                End If
                If txtTemp(3) <> "" Then
                   wBrkCode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(3), "S", "AcCode", "N")
                Else
                   wBrkCode = 0
                End If
                If txtTemp(0) <> "" Then
                   wMillCode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(0), "S", "AcCode", "N")
                Else
                   wMillCode = 0
                End If
                gCmd.CommandText = "PrcPrepareDespatchDetail"
                gCmd.Parameters.Refresh
                gCmd.Parameters("@Fr_dt") = Format(dtpFrDt, "yyyy/mm/dd")
                gCmd.Parameters("@To_dt") = Format(dtpToDt, "yyyy/mm/dd")
                gCmd.Parameters("@UCode") = gUserName
                If ChkConsolidate.Value = 1 Then
                   gCmd.Parameters("@VFirm") = ""
                Else
                   gCmd.Parameters("@VFirm") = Trim(gCCode)
                End If
                gCmd.Parameters("@Vtype") = "OM"
                If wPartyCode <> 0 Then gCmd.Parameters("@Accode") = wPartyCode
                If wBrkCode <> 0 Then gCmd.Parameters("@Accode_Br") = wBrkCode
                If wMillCode <> 0 Then gCmd.Parameters("@MillCode") = wMillCode
                gCmd.Execute
                x1 = "{TmpGentbl.UserName}='" & gUserName & "'"
               .SelectionFormula = x1
               .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt, "dd/MM/yyyy") & "'"
               If cmdCrystal(2).Value = True Then
                  RecYN = GProcRstOpen(tmp1, "select Distinct millname,vfirmname from TmpGentbl where TmpGentbl.UserName ='" & gUserName & "' order by millname", "R", gCn)
                  If RecYN > 0 Then
                     With tmp1
                          .MoveFirst
                          Do While Not .EOF
                             Set objCrystal = New CRAXDRT.Application
                             Tattach = gReportPath & "rptBookingDespDetail.rpt"
                             Set objReport = objCrystal.OpenReport(Tattach, 1)
                             x1 = "{TmpGentbl.UserName}='" & gUserName & "' and {TmpGentbl.MillName} = '" & tmp1.Fields(0) & "' and {TmpGentbl.vfirmname} = '" & tmp1.Fields(1) & "' "
                             objReport.RecordSelectionFormula = x1
                             If gCIBOffice = "I" Then
                                objReport.Database.LogOnServer "pdsodbc.dll", "Hitrix", "Vishnu", "sa", "123456"
                             Else
                                objReport.Database.LogOnServer "pdsodbc.dll", "Hitrix", "VishnuMum", "sa", "123456"
                             End If
                             '---   X"
    '                        Dim j As Integer
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
                                             objReport.FormulaFields(j).text = "" & Chr(39) & " From : " & dtpFrDt & "   To " & dtpToDt & "" & Chr(39) & ""
                                        Case "{@wNewPage}"
                                             objReport.FormulaFields(j).text = "0"
                                       End Select
                             Next j
                             wMillMailId = GProcGetColumnValue("tblMastAccount", "AcName", tmp1.Fields(0), "S", "AcEmail", "S")
                             If wMillMailId = "" Or IsNull(wMillMailId) Then
                                 MsgBox "Mail ID Not Found" + tmp1.Fields(0)
                                 Exit Sub
                             End If
                             PtyMailId = wMillMailId
'                             Mfile = CStr(App.Path) + "\MailPDF\" + gCCode + "-" + gUserName + "-" + "DesPacth_Detail"
                             Mfile = CStr(App.Path) + "\MailPDF\" + .Fields(0) + "-" + .Fields(1) + "-" + gUserName + "-" + "DesPacth_Detail"
                             ExportReportToPDF objReport, Mfile & ".pdf", "foo"
                             Mfile = Mfile + ".pdf"
                             MailSub = "Despatch Goods Details "
                             MailStr = "Dear Sir, " + Chr(13) + Chr(10)
                             MailStr = MailStr + "Despatch Goods Details"
                             gMailsendToMill = False
                             Call SendEmail1("" & Trim(gCMailId) & "", "" & PtyMailId & "", MailSub, MailStr, "" & PtyMailIdCc & "", "" & PtyMailIdBcc & "", Mfile)
                             MsgBox ("Mail Send....." + tmp1.Fields(0))
                             .MoveNext
                          Loop
                     End With
                  Else
                     MsgBox "Mail ID Not Found"
                  End If
                  Exit Sub
               Else
                   If opsummary.Value = True Then
                      .ReportFileName = gReportPath & "rptBookingDespDetailSummery.rpt"
                   Else
                      .ReportFileName = gReportPath & "rptBookingDespDetail.rpt"
                   End If
               End If
        Case "Trade Purchase Booking"
              ShowRepo = True
                If txtTemp(2) <> "" Then
                   wPartyCode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(2), "S", "AcCode", "N")
                Else
                   wPartyCode = 0
                End If
                If txtTemp(3) <> "" Then
                   wBrkCode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(3), "S", "AcCode", "N")
                Else
                   wBrkCode = 0
                End If
                If txtTemp(0) <> "" Then
                   wMillCode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(0), "S", "AcCode", "N")
                Else
                   wMillCode = 0
                End If
                gCmd.CommandText = "PrcPreparePendingBookingPurch"
                gCmd.Parameters.Refresh
                gCmd.Parameters("@Fr_dt") = Format(dtpFrDt, "yyyy/mm/dd")
                gCmd.Parameters("@To_dt") = Format(dtpToDt, "yyyy/mm/dd")
                gCmd.Parameters("@UCode") = gUserName
                gCmd.Parameters("@VFirm") = Trim(gCCode)
                gCmd.Parameters("@Vtype") = "OT"
                gCmd.Parameters("@RptTp") = "P"
                If ChkConsolidate.Value = 1 Then
                   gCmd.Parameters("@VFirm") = ""
                Else
                   gCmd.Parameters("@VFirm") = Trim(gCCode)
                End If
                If wPartyCode <> 0 Then gCmd.Parameters("@Accode") = wPartyCode
                If wBrkCode <> 0 Then gCmd.Parameters("@Accode_Br") = wBrkCode
                If wMillCode <> 0 Then gCmd.Parameters("@MillCode") = wMillCode
                If wItemCode <> 0 Then gCmd.Parameters("@ItemCode") = wItemCode
                gCmd.Execute
                x1 = "{TmpGentbl.UserName}='" & gUserName & "'"
               .SelectionFormula = x1
               .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt, "dd/MM/yyyy") & "'"
               .ReportFileName = gReportPath & "rptBookingPurchPending.rpt"
        
               If cmdCrystal(2).Value = True Then
'                  X = CStr(App.Path) + "\MailPDF\*" & ".pdf"
'                  If Len(Dir$(X)) > 0 Then
'                     Kill CStr(App.Path) + "\MailPDF\*" & ".pdf"
'                  End If
'                  Set objCrystal = New CRAXDRT.Application
'                  Tattach = gReportPath & "rptBookingPurchPending.rpt"
'                  Set objReport = objCrystal.OpenReport(Tattach, 1)
'                  objReport.RecordSelectionFormula = x1
'                  If gCIBOffice = "I" Then
'                     objReport.Database.LogOnServer "pdsodbc.dll", "Hitrix", "Vishnu", "sa", "123456"
'                  Else
'                     objReport.Database.LogOnServer "pdsodbc.dll", "Hitrix", "VishnuMum", "sa", "123456"
'                  End If
'                  For j = 1 To objReport.FormulaFields.Count
'                      Select Case objReport.FormulaFields(j).Name
'                             Case "{@FirmNm}"
'                                   objReport.FormulaFields(j).text = "Trim(" & Chr(39) & UCase(gCName) & Chr(39) & ")"
'                             Case "{@wAdd1}"
'                                   objReport.FormulaFields(j).text = "Trim(" & Chr(39) & gCAdd1 & Chr(39) & ")"
'                             Case "{@wAdd2}"
'                                   objReport.FormulaFields(j).text = "Trim(" & Chr(39) & gCAdd2 & Chr(39) & ")"
'                             Case "{@wPhNo}"
'                                   objReport.FormulaFields(j).text = "Trim(" & Chr(39) & gCPhNo & Chr(39) & ")"
'                             Case "{@wFrToDt}"
'                                   objReport.FormulaFields(j).text = "" & Chr(39) & " From : " & dtpFrDt & "   To " & dtpToDt & "" & Chr(39) & ""
'                       End Select
'                  Next j
                  
'--FOr Loop
                  RecYN = GProcRstOpen(tmp1, "select Distinct millname,vfirmname from TmpGentbl where TmpGentbl.UserName ='" & gUserName & "' order by millname", "R", gCn)
                  If RecYN > 0 Then
                     With tmp1
                          .MoveFirst
                          Do While Not .EOF
                             Set objCrystal = New CRAXDRT.Application
                             Tattach = gReportPath & "rptBookingPurchPending.rpt"
                             Set objReport = objCrystal.OpenReport(Tattach, 1)
                             x1 = "{TmpGentbl.UserName}='" & gUserName & "' and {TmpGentbl.MillName} = '" & tmp1.Fields(0) & "' and {TmpGentbl.vfirmname} = '" & tmp1.Fields(1) & "' "
                             objReport.RecordSelectionFormula = x1
                             If gCIBOffice = "I" Then
                                objReport.Database.LogOnServer "pdsodbc.dll", "Hitrix", "Vishnu", "sa", "123456"
                             Else
                                objReport.Database.LogOnServer "pdsodbc.dll", "Hitrix", "VishnuMum", "sa", "123456"
                             End If
                             For j = 1 To objReport.FormulaFields.Count
                                 Select Case objReport.FormulaFields(j).Name
                                        Case "{@FirmNm}"
                                              objReport.FormulaFields(j).text = "Trim(" & Chr(39) & UCase(tmp1.Fields(1)) & Chr(39) & ")"
                                        Case "{@wAdd1}"
                                              objReport.FormulaFields(j).text = "Trim(" & Chr(39) & gCAdd1 & Chr(39) & ")"
                                        Case "{@wAdd2}"
                                              objReport.FormulaFields(j).text = "Trim(" & Chr(39) & gCAdd2 & Chr(39) & ")"
                                        Case "{@wPhNo}"
                                              objReport.FormulaFields(j).text = "Trim(" & Chr(39) & gCPhNo & Chr(39) & ")"
                                        Case "{@wFrToDt}"
                                              objReport.FormulaFields(j).text = "" & Chr(39) & " From : " & dtpFrDt & "   To " & dtpToDt & "" & Chr(39) & ""
                                  End Select
                             Next j
                             wMillMailId = GProcGetColumnValue("tblMastAccount", "AcName", tmp1.Fields(0), "S", "AcEmail", "S")
                             If wMillMailId = "" Or IsNull(wMillMailId) Then
                                MsgBox "Mail ID Not Found " + tmp1.Fields(0)
                             Else
                                PtyMailId = wMillMailId
'                                Mfile = CStr(App.Path) + "\MailPDF\" + gCCode + "-" + gUserName + "-" + "Purch_Bk_Detail"
                                Mfile = CStr(App.Path) + "\MailPDF\" + .Fields(0) + "-" + .Fields(1) + "-" + gUserName + "-" + "Purch_Bk_Detail"
                                ExportReportToPDF objReport, Mfile & ".pdf", "foo"
                                Mfile = Mfile + ".pdf"
                                MailSub = "Pending Purchase Booking Details "
                                MailStr = "Dear Sir, " + Chr(13) + Chr(10)
                                MailStr = MailStr + "Attaching herewith the Pending Purchase Booking Details"
                                gMailsendToMill = False
                                Call SendEmail1("" & Trim(gCMailId) & "", "" & PtyMailId & "", MailSub, MailStr, "" & PtyMailIdCc & "", "" & PtyMailIdBcc & "", Mfile)
'                                MsgBox ("Mail Send....." + tmp1.Fields(0))
                             End If
                             .MoveNext
                          Loop
                     End With
                  Else
                     MsgBox "Mail ID Not Found"
                     Exit Sub
                  End If
                  MsgBox "All Mail Sent"


'                  X = CStr(App.Path) + "\MailPDF\*" & ".pdf"
'                  If Len(Dir$(X)) > 0 Then
'                     Kill CStr(App.Path) + "\MailPDF\*" & ".pdf"
'                  End If
                  
                  
                  Exit Sub
                                   
                  
                  
                  
                  
'                  wMillMailId = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(0), "S", "AcEmail", "S")
'                  If wMillMailId = "" Or IsNull(wMillMailId) Then
'                     MsgBox "Mail ID Not Found"
'                     Exit Sub
'                  End If
'                  PtyMailId = wMillMailId
'                  Mfile = CStr(App.Path) + "\MailPDF\" + gCCode + "-" + gUserName + "-" + "Purch_Bk_Detail"
'                  ExportReportToPDF objReport, Mfile & ".pdf", "foo"
'                  Mfile = Mfile + ".pdf"
'                  MailSub = "Pending Purchase Booking Details "
'                  MailStr = "Dear Sir, " + Chr(13) + Chr(10)
'                  MailStr = MailStr + "Attaching herewith the Pending Purchase Booking Details"
'                  gMailsendToMill = False
'                  Call SendEmail1("" & Trim(gCMailId) & "", "" & PtyMailId & "", MailSub, MailStr, "" & PtyMailIdCc & "", "" & PtyMailIdBcc & "", Mfile)
'                  Exit Sub
              End If
        Case "Party Booking Vs Dispatch Register"
              ShowRepo = True
                If txtTemp(2) <> "" Then
                   wPartyCode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(2), "S", "AcCode", "N")
                Else
                   wPartyCode = 0
                End If
                If txtTemp(3) <> "" Then
                   wBrkCode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(3), "S", "AcCode", "N")
                Else
                   wBrkCode = 0
                End If
                If txtTemp(0) <> "" Then
                   wMillCode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(0), "S", "AcCode", "N")
                Else
                   wMillCode = 0
                End If
                gCmd.CommandText = "PrcPrepareBookingPartyVsDesp"
                gCmd.Parameters.Refresh
                gCmd.Parameters("@Fr_dt") = Format(dtpFrDt, "yyyy/mm/dd")
                gCmd.Parameters("@To_dt") = Format(dtpToDt, "yyyy/mm/dd")
                gCmd.Parameters("@UCode") = gUserName
                gCmd.Parameters("@VFirm") = Trim(gCCode)
                gCmd.Parameters("@Vtype") = "OP"
                If ChkConsolidate.Value = 1 Then
                   gCmd.Parameters("@VFirm") = ""
                Else
                   gCmd.Parameters("@VFirm") = Trim(gCCode)
                End If
                If wPartyCode <> 0 Then gCmd.Parameters("@Accode") = wPartyCode
                If wBrkCode <> 0 Then gCmd.Parameters("@Accode_Br") = wBrkCode
                If wMillCode <> 0 Then gCmd.Parameters("@MillCode") = wMillCode
                If wItemCode <> 0 Then gCmd.Parameters("@ItemCode") = wItemCode
                If Val(txt(0)) <> 0 Then gCmd.Parameters("@BkNo") = txt(0)
                gCmd.Execute
                x1 = "{TmpGentbl.UserName}='" & gUserName & "'"
               .SelectionFormula = x1
               .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt, "dd/MM/yyyy") & "'"
               .ReportFileName = gReportPath & "rptBookingPartyVsDesp.rpt"
        Case "Party Booking Register", "Pending Party Booking Register", "Party Booking Summery", "Purchase Booking Summery"
              ShowRepo = True
                If txtTemp(2) <> "" Then
                   wPartyCode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(2), "S", "AcCode", "N")
                Else
                   wPartyCode = 0
                End If
                If txtTemp(3) <> "" Then
                   wBrkCode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(3), "S", "AcCode", "N")
                Else
                   wBrkCode = 0
                End If
                If txtTemp(0) <> "" Then
                   wMillCode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(0), "S", "AcCode", "N")
                Else
                   wMillCode = 0
                End If
                gCmd.CommandText = "PrcPreparePendingBookingParty"
                gCmd.Parameters.Refresh
                gCmd.Parameters("@Fr_dt") = Format(dtpFrDt, "yyyy/mm/dd")
                gCmd.Parameters("@To_dt") = Format(dtpToDt, "yyyy/mm/dd")
                gCmd.Parameters("@UCode") = gUserName
                gCmd.Parameters("@VFirm") = Trim(gCCode)
                gCmd.Parameters("@Vtype") = "OP"
                If gSelectedMenu = "Pending Party Booking Register" Then
                   gCmd.Parameters("@RptTp") = "P"
                    If ChkConsolidate.Value = 1 Then
                       gCmd.Parameters("@VFirm") = ""
                    Else
                       gCmd.Parameters("@VFirm") = Trim(gCCode)
                    End If
                ElseIf gSelectedMenu = "Party Booking Summery" Then
                   gCmd.Parameters("@RptTp") = "S"
                    If ChkConsolidate.Value = 1 Then
                       gCmd.Parameters("@VFirm") = ""
                    Else
                       gCmd.Parameters("@VFirm") = Trim(gCCode)
                    End If
                ElseIf gSelectedMenu = "Purchase Booking Summery" Then
                   gCmd.Parameters("@Vtype") = "OT"
                   gCmd.Parameters("@RptTp") = "S"
                    If ChkConsolidate.Value = 1 Then
                       gCmd.Parameters("@VFirm") = ""
                    Else
                       gCmd.Parameters("@VFirm") = Trim(gCCode)
                    End If
                Else
                   gCmd.Parameters("@RptTp") = "R"
                    If ChkConsolidate.Value = 1 Then
                       gCmd.Parameters("@VFirm") = ""
                    Else
                       gCmd.Parameters("@VFirm") = Trim(gCCode)
                    End If
                End If
                If wPartyCode <> 0 Then gCmd.Parameters("@Accode") = wPartyCode
                If wBrkCode <> 0 Then gCmd.Parameters("@Accode_Br") = wBrkCode
                If wMillCode <> 0 Then gCmd.Parameters("@MillCode") = wMillCode
                If wItemCode <> 0 Then gCmd.Parameters("@ItemCode") = wItemCode
                gCmd.Execute
                x1 = "{TmpGentbl.UserName}='" & gUserName & "'"
               .SelectionFormula = x1
               .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt, "dd/MM/yyyy") & "'"
                If gSelectedMenu = "Pending Party Booking Register" Then
                   If cmdCrystal(2).Value = True Then
                      Set objCrystal = New CRAXDRT.Application
                      Tattach = gReportPath & "rptBookingPartyPending.rpt"
                      Set objReport = objCrystal.OpenReport(Tattach, 1)
                      objReport.RecordSelectionFormula = "{TmpGentbl.UserName}='" & gUserName & "'"
                      If gCIBOffice = "I" Then
                         objReport.Database.LogOnServer "pdsodbc.dll", "Hitrix", "Vishnu", "sa", "123456"
                      Else
                         objReport.Database.LogOnServer "pdsodbc.dll", "Hitrix", "VishnuMum", "sa", "123456"
                      End If
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
                               objReport.FormulaFields(j).text = "" & Chr(39) & " From : " & dtpFrDt & "   To " & dtpToDt & "" & Chr(39) & ""
                          End Select
                      Next j
                      wMillMailId = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(0), "S", "AcEmail", "S")
                      If wMillMailId = "" Or IsNull(wMillMailId) Then
                         MsgBox "Mail ID Not Found"
                         Exit Sub
                      End If
                      PtyMailId = wMillMailId
                      Mfile = CStr(App.Path) + "\MailPDF\" + gCCode + "-" + gUserName + "-" + "Pnd_Pty_Bk"
                      ExportReportToPDF objReport, Mfile & ".pdf", "foo"
                      Mfile = Mfile + ".pdf"
                      MailSub = "Pending Booking Details "
                      MailStr = "Dear Sir, " + Chr(13) + Chr(10)
                      MailStr = MailStr + "Attaching herewith the Pending Booking Details"
                                
                      gMailsendToMill = False
                      Call SendEmail1("" & Trim(gCMailId) & "", "" & PtyMailId & "", MailSub, MailStr, "" & PtyMailIdCc & "", "" & PtyMailIdBcc & "", Mfile)
'                                 Kill Mid(Tattach, 1, Len(Tattach) - 4) & ".pdf"
                      Exit Sub
                   Else
                     .ReportFileName = gReportPath & "rptBookingPartyPending.rpt"
                   End If
                ElseIf gSelectedMenu = "Purchase Booking Summery" Then
                  .Formulas(5) = "wRptName = 'Purchase Booking Summery'"
                  .ReportFileName = gReportPath & "rptBookingPartySummery.rpt"
                ElseIf gSelectedMenu = "Party Booking Summery" Then
                  .Formulas(5) = "wRptName = 'Party Booking Summery'"
                  .ReportFileName = gReportPath & "rptBookingPartySummery.rpt"
                Else
                       If cmdCrystal(2).Value = True Then
                                Set objCrystal = New CRAXDRT.Application
                                 Tattach = gReportPath & "rptBookingPartyReg.rpt"
                                Set objReport = objCrystal.OpenReport(Tattach, 1)
                                objReport.RecordSelectionFormula = "{TmpGentbl.UserName}='" & gUserName & "'"
                                If gCIBOffice = "I" Then
                                    objReport.Database.LogOnServer "pdsodbc.dll", "Hitrix", "Vishnu", "sa", "123456"
                                Else
                                    objReport.Database.LogOnServer "pdsodbc.dll", "Hitrix", "VishnuMum", "sa", "123456"
                                 End If
                                '---   X"
        '                        Dim j As Integer
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
                                   objReport.FormulaFields(j).text = "" & Chr(39) & " From : " & dtpFrDt & "   To " & dtpToDt & "" & Chr(39) & ""
                               End Select
                               Next j
                               wMillMailId = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(0), "S", "AcEmail", "S")
                               If wMillMailId = "" Or IsNull(wMillMailId) Then
                                    MsgBox "Mail ID Not Found"
                                     Exit Sub
                                End If
                                PtyMailId = wMillMailId
                                Mfile = CStr(App.Path) + "\MailPDF\" + gCCode + "-" + gUserName + "-" + "Booking_Detail"
                                ExportReportToPDF objReport, Mfile & ".pdf", "foo"
                                Mfile = Mfile + ".pdf"
                                
'                                ExportReportToPDF objReport, Mid(Tattach, 1, Len(Tattach) - 4) & ".pdf", "foo"
'                                Mfile = gReportPath & "rptBookingPartyReg.pdf"
                                MailSub = "Booking Details "
                                MailStr = "Dear Sir, " + Chr(13) + Chr(10)
                                MailStr = MailStr + "Attaching herewith the Booking Details"
                                
                    
                    '                   MailStr = MailStr + "Please find attached herewith GST Booking Order" & vbLf
                    '                   MailStr = MailStr + "Note : This is system generated email, Please do not reply."
                                 gMailsendToMill = False
                                 Call SendEmail1("" & Trim(gCMailId) & "", "" & PtyMailId & "", MailSub, MailStr, "" & PtyMailIdCc & "", "" & PtyMailIdBcc & "", Mfile)
'                                 Kill Mid(Tattach, 1, Len(Tattach) - 4) & ".pdf"
                           Exit Sub
                      Else
                           .ReportFileName = gReportPath & "rptBookingPartyReg.rpt"
                      End If
                End If
        Case "Stock Statement Godownwsie"
              ShowRepo = True
              gCn.BeginTrans
              gCmd.CommandText = "PrcPrepareLotwiseGodownStock"
              gCmd.Parameters.Refresh
              If ChkConsolidate.Value = 1 Then
                 gCmd.Parameters("@VFirm") = ""
              Else
                 gCmd.Parameters("@VFirm") = Trim(gCCode)
              End If
              gCmd.Parameters("@Fr_dt") = Format(gCYSDate, "yyyy/mm/dd")
              gCmd.Parameters("@To_dt") = Format(dtpToDt, "yyyy/mm/dd")
              If wMillCode <> 0 Then gCmd.Parameters("@MillCode") = wMillCode
              If wItemCode <> 0 Then gCmd.Parameters("@ItemCode") = wItemCode
              If wGodCd <> 0 Then gCmd.Parameters("@GodownCd") = wGodCd
              gCmd.Parameters("@LotNo") = ""
              gCmd.Parameters("@StkTP") = ""
              gCmd.Parameters("@Vno") = ""
              gCmd.Parameters("@Vtype") = ""
              gCmd.Parameters("@Vyear") = ""
              gCmd.Parameters("@IsModi") = "N"
              gCmd.Parameters("@UCode") = gUserName
              gCmd.Execute
              gCn.CommitTrans
              x1 = "{tmpLotwiseStock.UserName}='" & gUserName & "'"
              .SelectionFormula = x1
              .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(gCYSDate, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt, "dd/MM/yyyy") & "'"
              .ReportFileName = gReportPath & "rptStockStLotGodown.rpt"
        Case "Stock Statement"
              ShowRepo = True
                If txtTemp(0) <> "" Then
                   wMillCode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(0), "S", "AcCode", "N")
                Else
                   wMillCode = 0
                End If
                If chkBox3.Value = 1 Then
                   gCn.BeginTrans
                   gCmd.CommandText = "PrcPrepareLotwiseStock"
                   gCmd.Parameters.Refresh
                   If ChkConsolidate.Value = 1 Then
                      gCmd.Parameters("@VFirm") = ""
                   Else
                      gCmd.Parameters("@VFirm") = Trim(gCCode)
                   End If
                   gCmd.Parameters("@Fr_dt") = Format(gCYSDate, "yyyy/mm/dd")
                   gCmd.Parameters("@To_dt") = Format(dtpToDt, "yyyy/mm/dd")
                   If wMillCode <> 0 Then gCmd.Parameters("@MillCode") = wMillCode
                   If wItemCode <> 0 Then gCmd.Parameters("@ItemCode") = wItemCode
                   gCmd.Parameters("@LotNo") = ""
                   gCmd.Parameters("@StkTP") = ""
                   gCmd.Parameters("@Vno") = ""
                   gCmd.Parameters("@Vtype") = ""
                   gCmd.Parameters("@Vyear") = ""
                   gCmd.Parameters("@IsModi") = "N"
                   gCmd.Parameters("@UCode") = gUserName
                   gCmd.Execute
                   gCn.CommitTrans
                   x1 = "{tmpLotwiseStock.UserName}='" & gUserName & "'"
                  .SelectionFormula = x1
                  .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(gCYSDate, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt, "dd/MM/yyyy") & "'"
                  .ReportFileName = gReportPath & "rptStockStatementLot.rpt"
                Else
                   gCmd.CommandText = "PrcPrepareStock"
                   gCmd.Parameters.Refresh
                   gCmd.Parameters("@YearStart_Dt") = Format(gCYSDate, "yyyy/mm/dd")
                   gCmd.Parameters("@Fr_dt") = Format(dtpFrDt, "yyyy/mm/dd")
                   gCmd.Parameters("@To_dt") = Format(dtpToDt, "yyyy/mm/dd")
                   gCmd.Parameters("@UCode") = gUserName
                   gCmd.Parameters("@VFirm") = Trim(gCCode)
                   If wItemCode <> 0 Then gCmd.Parameters("@ItemCode") = wItemCode
                   If wMillCode <> 0 Then gCmd.Parameters("@MillCode") = wMillCode
                   gCmd.Execute
                   x1 = "{TmpGentbl.UserName}='" & gUserName & "'"
                  .SelectionFormula = x1
                  .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt, "dd/MM/yyyy") & "'"
                  .ReportFileName = gReportPath & "rptStockStatement.rpt"
                End If
        Case "Daily Stock Statement"
              ShowRepo = True
                If txtTemp(0) <> "" Then
                   wMillCode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(0), "S", "AcCode", "N")
                Else
                   wMillCode = 0
                End If
                gCmd.CommandText = "PrcPrepareStockDaily"
                gCmd.Parameters.Refresh
                gCmd.Parameters("@YearStart_Dt") = Format(gCYSDate, "yyyy/mm/dd")
                gCmd.Parameters("@Fr_dt") = Format(dtpFrDt, "yyyy/mm/dd")
                gCmd.Parameters("@To_dt") = Format(dtpToDt, "yyyy/mm/dd")
                gCmd.Parameters("@UCode") = gUserName
                gCmd.Parameters("@VFirm") = Trim(gCCode)
                If wItemCode <> 0 Then gCmd.Parameters("@ItemCode") = wItemCode
                If wMillCode <> 0 Then gCmd.Parameters("@MillCode") = wMillCode
                gCmd.Execute
                x1 = "{TmpGentbl.UserName}='" & gUserName & "' and {TmpGentbl.Amt5} > 1"
               .SelectionFormula = x1
               .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt, "dd/MM/yyyy") & "'"
               .ReportFileName = gReportPath & "rptStockStatementDaily.rpt"
        Case "Sales Summery"
              ShowRepo = True
                If txtTemp(0) <> "" Then
                   wMillCode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(0), "S", "AcCode", "N")
                Else
                   wMillCode = 0
                End If
                gCmd.CommandText = "PrcPrepareSaleStatusRatewise"
                gCmd.Parameters.Refresh
                gCmd.Parameters("@Fr_dt") = Format(dtpFrDt, "yyyy/mm/dd")
                gCmd.Parameters("@To_dt") = Format(dtpToDt, "yyyy/mm/dd")
                gCmd.Parameters("@UCode") = gUserName
                If ChkConsolidate.Value = 1 Then
                   gCmd.Parameters("@VFirm") = ""
                Else
                   gCmd.Parameters("@VFirm") = Trim(gCCode)
                End If
                If wPartyCode <> 0 Then gCmd.Parameters("@Accode") = wPartyCode
                If wBrkCode <> 0 Then gCmd.Parameters("@Accode_Br") = wBrkCode
                If wMillCode <> 0 Then gCmd.Parameters("@MillCode") = wMillCode
                If wItemCode <> 0 Then gCmd.Parameters("@ItemCode") = wItemCode


                gCmd.Execute
                x1 = "{TmpGentbl.UserName}='" & gUserName & "' "
               .SelectionFormula = x1
               .Formulas(4) = "wFrToDt = 'From : ' & '" & dtpFrDt & "' & '  To ' & '" & dtpToDt & "'"
               If ChkConsolidate.Value = 1 Then
                  .Formulas(5) = "wAllFirm = 'All Firm '"
               Else
                  .Formulas(5) = "wAllFirm = 'Single Firm '"
               End If
               .ReportFileName = gReportPath & "rptSaleSummaryStock.rpt"
        Case "Gate Pass Wise Sales"
              ShowRepo = True
                If txtTemp(0) <> "" Then
                   wMillCode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(0), "S", "AcCode", "N")
                Else
                   wMillCode = 0
                End If
                gCmd.CommandText = "PrcPrepareGetPassWiseSale"
                gCmd.Parameters.Refresh
                gCmd.Parameters("@Fr_dt") = Format(dtpFrDt, "yyyy/mm/dd")
                gCmd.Parameters("@To_dt") = Format(dtpToDt, "yyyy/mm/dd")
                gCmd.Parameters("@UCode") = gUserName
                If ChkConsolidate.Value = 1 Then
                   gCmd.Parameters("@VFirm") = ""
                Else
                   gCmd.Parameters("@VFirm") = Trim(gCCode)
                End If
                If wPartyCode <> 0 Then gCmd.Parameters("@Accode") = wPartyCode
                If wBrkCode <> 0 Then gCmd.Parameters("@Accode_Br") = wBrkCode
                If wMillCode <> 0 Then gCmd.Parameters("@MillCode") = wMillCode
                If wItemCode <> 0 Then gCmd.Parameters("@ItemCode") = wItemCode
                gCmd.Execute
                x1 = "{TmpGentbl.UserName}='" & gUserName & "'"
               .SelectionFormula = x1
               .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt, "dd/MM/yyyy") & "'"
               .ReportFileName = gReportPath & "rptGatePasswiseSales.rpt"
        Case "Trade Sale Register GST", "Depot Sale Register GST", "Trade Sale Register", "Consignment/Depot Sale Register", "Mill Bill Register", "Other Sales", "Other Sales", "Commission And Brokerage Sales", "Non-GST Outward Sales", "Sales Return Register"
                If txtTemp(2) <> "" Then
                   wPartyCode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(2), "S", "AcCode", "N")
                Else
                   wPartyCode = 0
                End If
                If txtTemp(3) <> "" Then
                   wBrkCode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(3), "S", "AcCode", "N")
                Else
                   wBrkCode = 0
                End If
                If txtTemp(0) <> "" Then
                   wMillCode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(0), "S", "AcCode", "N")
                Else
                   wMillCode = 0
                End If
              ShowRepo = True
              gCmd.CommandText = "PrcPrepareSaleRegisterGST"
                gCmd.Parameters.Refresh
                gCmd.Parameters("@Fr_dt") = Format(dtpFrDt, "yyyy/mm/dd")
                gCmd.Parameters("@To_dt") = Format(dtpToDt, "yyyy/mm/dd")
                gCmd.Parameters("@UCode") = gUserName
                gCmd.Parameters("@VFirm") = Trim(gCCode)
                If wPartyCode <> 0 Then gCmd.Parameters("@Accode") = wPartyCode
                If wMillCode <> 0 Then gCmd.Parameters("@MillCode") = wMillCode
                If gSelectedMenu = "Trade Sale Register GST" Or gSelectedMenu = "Trade Sale Register" Then
                   gCmd.Parameters("@Vtype") = "SY"
                   .Formulas(5) = "RepoName_A = 'Trade Sales  Register' "
                   .Formulas(6) = "RepoName_B = 'Product Summery Trade Sales' "
                   .Formulas(7) = "RepoName_C = '' "
                End If
                If gSelectedMenu = "Sales Return Register" Then
                   gCmd.Parameters("@Vtype") = "RY"
                   .Formulas(5) = "RepoName_A = 'Sales Return Register' "
                   .Formulas(6) = "RepoName_B = 'Product Summery Sales Return' "
                   .Formulas(7) = "RepoName_C = '' "
                End If
                
                If gSelectedMenu = "Other Sales" Or gSelectedMenu = "Non-GST Outward Sales" Then
                   gCmd.Parameters("@Vtype") = "SC"
                   If gSelectedMenu = "Other Sales" Then
                       .Formulas(5) = "RepoName_A = 'Other Sales Register' "
                   Else
                       .Formulas(5) = "RepoName_A = 'Non-GST Outward Sales' "
                    End If
                   .Formulas(6) = "RepoName_B = 'Product Summery Trade Sales' "
                   .Formulas(7) = "RepoName_C = '' "
                End If
                If gSelectedMenu = "Commission And Brokerage Sales" Then
                   gCmd.Parameters("@Vtype") = "SB"
                   .Formulas(5) = "RepoName_A = 'Commission And Brokerage Sales Register' "
                   .Formulas(6) = "RepoName_B = 'Product Summery Commission And Brokerage Sales' "
                   .Formulas(7) = "RepoName_C = '' "
                End If
                If gSelectedMenu = "Depot Sale Register GST" Or gSelectedMenu = "Consignment/Depot Sale Register" Then
                    MillDepoCd = GProcGetColumnValue("tblMastNarration", "Narration", "Depot", "S", "NarrCode", "N")
                    IsMillDepo = GProcGetColumnValue("tblMastBillSerial", "MillCode", CStr(wMillCode), "N", "SaleType", "N", "Vfirm = '" & gCCode & "'")
                    If chkBox4.Value = 1 Then gCmd.Parameters("@Is2ndDay") = 1
                    If chkBox3.Value = 1 Then gCmd.Parameters("@IsRegiDay") = 1
                    
                    If gCIsDepotFirm = 1 Then
                       gCmd.Parameters("@Vtype") = "SD"
                    Else
                        If MillDepoCd = IsMillDepo Then
                           gCmd.Parameters("@Vtype") = "SD"
                        Else
                           gCmd.Parameters("@Vtype") = "SO"
                        End If
                    End If
                   .Formulas(5) = "RepoName_A = 'Consignment/Depot Sales  Register' "
                   .Formulas(6) = "RepoName_B = 'Product Summery Consignment/Depot Sales' "
                   .Formulas(7) = "RepoName_C = 'Mill Summery Consignment/Depot Sales' "
                   If txtTemp(0) = "" Then
                      .Formulas(8) = "MillName_A = '' "
                   Else
                      .Formulas(8) = "MillName_A = 'Mill Name : '+'" & txtTemp(0) & "' "
                   End If
                End If
                If gSelectedMenu = "Mill Bill Register" Then
                   gCmd.Parameters("@Vtype") = "SM"
                   .Formulas(5) = "RepoName_A = 'Mill Bill Sales  Register' "
                   .Formulas(6) = "RepoName_B = 'Product Summery Mill Bill' "
                   .Formulas(7) = "RepoName_C = 'Mill Summery Mill Bill' "
                   .Formulas(8) = "MillName_A = 'Mill Name : '+'" & txtTemp(0) & "' "
                End If
                gCmd.Execute
                x1 = "{TmpGentbl.UserName}='" & gUserName & "'"
               .SelectionFormula = x1
               .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt, "dd/MM/yyyy") & "'"
               If gSelectedMenu = "Mill Bill Register" Then
                  If OptShort.Value = True Then
                     gCn.BeginTrans
                     gCn.Execute "delete from tmpGenTbl where username='" & gUserName & "' and Amt15 = 1 "
                     gCn.CommitTrans
                  End If
                  If opsummary.Value = True Then
                     gCn.BeginTrans
                     gCn.Execute "delete from tmpGenTbl where username='" & gUserName & "' and Amt15 = 0 "
                     gCn.CommitTrans
                  End If
                  .ReportFileName = gReportPath & "rptRegisterSalesDetailMillBill.rpt"
               ElseIf gSelectedMenu = "Non-GST Outward Sales" Then
                   x1 = "{TmpGentbl.Amt2} + {TmpGentbl.Amt4} + {TmpGentbl.Amt6} = 0 and  {TmpGentbl.UserName}=  '" & gUserName & "'"
                  .ReportFileName = gReportPath & "rptRegisterSalesOtherNonGst.rpt"
               ElseIf gSelectedMenu = "Consignment/Depot Sale Register" And gCIsDepotFirm = 1 Then
                    Dim RptNM As String
                    
                    If chkBox4.Value = 1 Then
                      .Formulas(9) = "Is2ndDay = 'Only 2nd Day Bills' "
                       RptNM = "Only 2nd Day Bills"
                    ElseIf chkBox3.Value = 1 Then
                      .Formulas(9) = "Is2ndDay = 'Only Regular Bills' "
                       RptNM = "Only Regular Bills"
                    Else
                      .Formulas(9) = "Is2ndDay = 'All Bills' "
                       RptNM = "All Bills"
                    End If
                    If OptDetail.Value = True Then
                       If cmdCrystal(2).Value = True Then
                          Set objCrystal = New CRAXDRT.Application
                          Tattach = gReportPath & "rptRegisterSalesDetailGSTDepotFirm2.rpt"
                          Set objReport = objCrystal.OpenReport(Tattach, 1)
                          objReport.RecordSelectionFormula = "{TmpGentbl.UserName}='" & gUserName & "'"
                          If gCIBOffice = "I" Then
                             objReport.Database.LogOnServer "pdsodbc.dll", "Hitrix", "Vishnu", "sa", "123456"
                          Else
                              objReport.Database.LogOnServer "pdsodbc.dll", "Hitrix", "VishnuMum", "sa", "123456"
                          End If
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
                                   objReport.FormulaFields(j).text = "" & Chr(39) & " From : " & dtpFrDt & "   To " & dtpToDt & "" & Chr(39) & ""
                              Case "{@wNewPage}"
                                   objReport.FormulaFields(j).text = "0"
                              Case "{@Is2ndDay}"
                                   objReport.FormulaFields(j).text = "Trim(" & Chr(39) & RptNM & Chr(39) & ")"
                              End Select
                          Next j
                          wMillMailId = "" & txtTemp(1) & ""
                          If wMillMailId = "" Or IsNull(wMillMailId) Then
                              MsgBox "Mail ID Not Found"
                              Exit Sub
                          End If
                          PtyMailId = wMillMailId
                          Mfile = CStr(App.Path) + "\MailPDF\" + gCCode + "-" + gUserName + "-" + "Depot_Sale_Regi"
                          ExportReportToPDF objReport, Mfile & ".pdf", "foo"
                          Mfile = Mfile + ".pdf"
                          MailSub = "Sale Register"
                          MailStr = "Dear Sir, " + Chr(13) + Chr(10)
                          MailStr = MailStr + "Sale Register"
                          gMailsendToMill = False
                          Call SendEmail1("" & Trim(gCMailId) & "", "" & PtyMailId & "", MailSub, MailStr, "" & PtyMailIdCc & "", "" & PtyMailIdBcc & "", Mfile)
'                         Kill Mid(Tattach, 1, Len(Tattach) - 4) & ".pdf"
                      Exit Sub
                    Else
                       .ReportFileName = gReportPath & "rptRegisterSalesDetailGSTDepotFirm2.rpt"
                    End If
                  ElseIf OptShort.Value = True Then
                     x1 = "{TmpGentbl.Nar6} = 'A' and {TmpGentbl.UserName}='" & gUserName & "'"
                    .SelectionFormula = x1
                     .ReportFileName = gReportPath & "rptRegisterSalesDepotDatewise.rpt"
                  Else
                     x1 = "{TmpGentbl.Nar6} = 'A' and {TmpGentbl.UserName}='" & gUserName & "'"
                    .SelectionFormula = x1
                     .ReportFileName = gReportPath & "rptRegisterSalesDetailGSTDepotFirmSum.rpt"
                  End If
               Else
                  .ReportFileName = gReportPath & "rptRegisterSalesDetailGST.rpt"
               End If
        Case "Insurance Report - Purchase"
                gCn.BeginTrans
                gCn.Execute "delete from tmpGenTbl where username='" & gUserName & "' "
                gCn.CommitTrans
                ShowRepo = True
                x1 = "select tblMastNarration.Narration,PurLrDate,PurLrNo,PurBillNo,PurBillDt,tblMastAccount.AcName,PurBillAmt,'" & gUserName & "'  " _
                     & " From tblPurch, tblMastNarration, tblMastAccount " _
                     & " where tblPurch.PurTransport = tblMastNarration.NarrCode and tblPurch.PurAcCrCode  = tblMastAccount.AcCode and " _
                     & " tblPurch.Vfirm = '" & gCCode & "' and tblPurch.Vtype = 'PY' and tblPurch.PurBillDt >= ('" & Format(dtpFrDt.Value, "yyyy/mm/dd") & "') and tblPurch.PurBillDt <= ('" & Format(dtpToDt.Value, "yyyy/mm/dd") & "') "
                If wPartyCode <> 0 Then
                   x1 = x1 + " and tblPurch.PurAcCrCode = " & wPartyCode & ""
                End If
                
                x1 = x1 + " and left(tblMastAccount.AcRemark,1) = 'Y' "
                
                gCn.BeginTrans
                gCn.Execute "insert into tmpGentbl (Nar1,Vdt,Nar2,VBillNo,VBillDt,MillName,Amt1,UserName) " & x1
                gCn.CommitTrans
               .SelectionFormula = "{TmpGentbl.UserName}='" & gUserName & "'"
               .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt, "dd/MM/yyyy") & "'"
               .ReportFileName = gReportPath & "rptInsurancePurch.rpt"
        Case "Trade Purchase Register GST", "Other Purchase Register GST", "Purchase Return Register"
                ShowRepo = True
                gCmd.CommandText = "PrcPreparePurchaseGST"
                gCmd.Parameters.Refresh
                gCmd.Parameters("@Fr_dt") = Format(dtpFrDt, "yyyy/mm/dd")
                gCmd.Parameters("@To_dt") = Format(dtpToDt, "yyyy/mm/dd")
                gCmd.Parameters("@UCode") = gUserName
                gCmd.Parameters("@VFirm") = Trim(gCCode)
                If gSelectedMenu = "Other Purchase Register GST" Then
                   gCmd.Parameters("@Vtype") = "PO"
                ElseIf gSelectedMenu = "Purchase Return Register" Then
                   gCmd.Parameters("@Vtype") = "VY"
                Else
                   gCmd.Parameters("@Vtype") = "PY"
                End If
                If wPartyCode <> 0 Then gCmd.Parameters("@Accode") = wPartyCode
                If wBrkCode <> 0 Then gCmd.Parameters("@Accode_Br") = wBrkCode
                If wMillCode <> 0 Then gCmd.Parameters("@MillCode") = wMillCode
                gCmd.Execute
                x1 = "{TmpGentbl.UserName}='" & gUserName & "'"
                ABC = "A"
                If chkNewPage.Value = 1 Then ABC = ABC + "B"
                If chkBox4.Value = 1 Then ABC = ABC + "C"
                x1 = x1 + " and {TmpGentbl.Nar6} in ( '" & ABC & "') "
               
               
               .SelectionFormula = x1
               .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt, "dd/MM/yyyy") & "'"
               If gSelectedMenu = "Other Purchase Register GST" Then
                  .ReportFileName = gReportPath & "rptRegisterPurchOtherGST.rpt"
               Else
                  If gSelectedMenu = "Purchase Return Register" Then
                    .Formulas(5) = "RepoName = 'Purchase Return / Shortage Register' "
                  Else
                    .Formulas(5) = "RepoName = 'Trade Purchase Register' "
                  End If
                  .ReportFileName = gReportPath & "rptRegisterPurchTradeGST.rpt"
               End If
        Case "R C M Sale Register GST"
                ShowRepo = True
                gCmd.CommandText = "PrcPrepareSalesRegisterRCM"
                gCmd.Parameters.Refresh
                gCmd.Parameters("@Fr_dt") = Format(dtpFrDt, "yyyy/mm/dd")
                gCmd.Parameters("@To_dt") = Format(dtpToDt, "yyyy/mm/dd")
                gCmd.Parameters("@UCode") = gUserName
                gCmd.Parameters("@VFirm") = Trim(gCCode)
                gCmd.Parameters("@Vtype") = "PO"
                If wPartyCode <> 0 Then gCmd.Parameters("@Accode") = wPartyCode
                If wBrkCode <> 0 Then gCmd.Parameters("@Accode_Br") = wBrkCode
                If wMillCode <> 0 Then gCmd.Parameters("@MillCode") = wMillCode
                gCmd.Execute
                x1 = "{TmpGentbl.UserName}='" & gUserName & "'"
               .SelectionFormula = x1
               .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt, "dd/MM/yyyy") & "'"
               .ReportFileName = gReportPath & "rptRegisterSalesRCM.rpt"
        Case "Internal Voucher Register", "Credit Note (Sales) Register", "Debit Note (Sales) Register", "Credit Note (Purchase) Register", "Debit Note (Purchase) Register"
             ShowRepo = True
             If OptDetail.Value = True And (gSelectedMenu = "Credit Note (Purchase) Register" Or gSelectedMenu = "Debit Note (Purchase) Register") Or (gSelectedMenu = "Credit Note (Sales) Register" And OptDetail.Value = True) Then
                gCmd.CommandText = "PrcPrepareRegisterCrDrNoteGST"
                gCmd.Parameters.Refresh
                gCmd.Parameters("@From_dt") = Format(dtpFrDt.Value, "yyyy/mm/dd")
                gCmd.Parameters("@To_dt") = Format(dtpToDt.Value, "yyyy/mm/dd")
                gCmd.Parameters("@UCode") = gUserName
                If ChkConsolidate.Value = 1 Then
                   gCmd.Parameters("@VFirm") = ""
                Else
                   gCmd.Parameters("@VFirm") = Trim(gCCode)
                End If
                If gSelectedMenu = "Credit Note (Purchase) Register" Then gCmd.Parameters("@RegTp") = "PV"
                If gSelectedMenu = "Credit Note (Sales) Register" Then gCmd.Parameters("@RegTp") = "PX"
                If gSelectedMenu = "Debit Note (Purchase) Register" Then gCmd.Parameters("@RegTp") = "SV"
                gCmd.Execute
                If ChkConsolidate.Value = 1 Then
                   X = "{TmpGentbl.UserName}= '" & gUserName & "' "
                Else
                   X = "{TmpGentbl.UserName}= '" & gUserName & "' and {TmpGentbl.Vfirm}='" & gCCode & "'   "
                End If
                .SelectionFormula = X
                .Formulas(4) = "wFrToDt = 'From : ' & '" & dtpFrDt.Value & "' & '  To ' & '" & dtpToDt.Value & "'"
                If gSelectedMenu = "Credit Note (Purchase) Register" Then .Formulas(5) = "wRptName = 'Credit Note (Purchase) GST Register '"
                If gSelectedMenu = "Debit Note (Purchase) Register" Then .Formulas(5) = "wRptName = 'Debit Note (Purchase) GST Register'"
                If gSelectedMenu = "Credit Note (Sales) Register" Then .Formulas(5) = "wRptName = 'Credit Note (Sales GST) Register'"
                .ReportFileName = gReportPath & "rptRegisterCrDrNtGST.rpt"
             Else
                gCmd.CommandText = "PrcPrepareRegisterJV"
                gCmd.Parameters.Refresh
                gCmd.Parameters("@From_dt") = Format(dtpFrDt.Value, "yyyy/mm/dd")
                gCmd.Parameters("@To_dt") = Format(dtpToDt.Value, "yyyy/mm/dd")
                gCmd.Parameters("@UCode") = gUserName
                If ChkConsolidate.Value = 1 Then
                   gCmd.Parameters("@VFirm") = ""
                Else
                   gCmd.Parameters("@VFirm") = Trim(gCCode)
                End If
                If gSelectedMenu = "Internal Voucher Register" Then gCmd.Parameters("@RegTp") = "JV"
                If gSelectedMenu = "Credit Note (Sales) Register" Then gCmd.Parameters("@RegTp") = "PN"
                If gSelectedMenu = "Debit Note (Sales) Register" Then gCmd.Parameters("@RegTp") = "SN"
                If gSelectedMenu = "Credit Note (Purchase) Register" Then gCmd.Parameters("@RegTp") = "PQ"
                If gSelectedMenu = "Debit Note (Purchase) Register" Then gCmd.Parameters("@RegTp") = "SQ"
                gCmd.Execute
                If ChkConsolidate.Value = 1 Then
                   X = "{tmpaccledger.UserName}= '" & gUserName & "' "
                Else
                   X = "{tmpaccledger.UserName}= '" & gUserName & "' and {tmpaccledger.Vfirm}='" & gCCode & "'   "
                End If
                 If gSelectedMenu = "Internal Voucher Register" And OptDetail.Value = True Then X = X + " and  left({tmpaccledger.Vnar2},1) = '9' and  left({tmpaccledger.Vnar3},1) = '9'  "
                 If gSelectedMenu = "Internal Voucher Register" And OptShort.Value = True Then X = X + " and   left({tmpaccledger.Vnar2},1) <> '9' and  left({tmpaccledger.Vnar3},1) <>  '9'  "
                .SelectionFormula = X
                .Formulas(4) = "wFrToDt = 'From : ' & '" & dtpFrDt.Value & "' & '  To ' & '" & dtpToDt.Value & "'"
                 If gSelectedMenu = "Internal Voucher Register" Then
                    If OptDetail.Value = True Then
                       .Formulas(5) = "wRptName = 'Broker JV Register'"
                    ElseIf OptShort.Value = True Then
                      .Formulas(5) = "wRptName = 'Other JV Register'"
                    ElseIf opsummary.Value = True Then
                      .Formulas(5) = "wRptName =  'All JV Register'"
                    End If
                 End If
                 If gSelectedMenu = "Credit Note (Sales) Register" Then .Formulas(5) = "wRptName = 'Credit Note (Sales) Register'"
                 If gSelectedMenu = "Debit Note (Sales) Register" Then .Formulas(5) = "wRptName = 'Debit Note (Sales) Register'"
                 If gSelectedMenu = "Credit Note (Purchase) Register" Then .Formulas(5) = "wRptName = 'Credit Note (Purchase) Register'"
                 If gSelectedMenu = "Debit Note (Purchase) Register" Then .Formulas(5) = "wRptName = 'Debit Note (Purchase) Register'"
                 .ReportFileName = gReportPath & "rptRegisterJV.rpt"
             End If
        Case "Late Pay Debit Note Register", "Late Pay Debit Note Register (Mill Bill)"
             ShowRepo = True
             gCmd.CommandText = "PrcPrepareRegisterLpInt"
             gCmd.Parameters.Refresh
             gCmd.Parameters("@From_dt") = Format(dtpFrDt.Value, "yyyy/mm/dd")
             gCmd.Parameters("@To_dt") = Format(dtpToDt.Value, "yyyy/mm/dd")
             gCmd.Parameters("@UCode") = gUserName
             If ChkConsolidate.Value = 1 Then
                gCmd.Parameters("@VFirm") = ""
             Else
                gCmd.Parameters("@VFirm") = Trim(gCCode)
             End If
             If gSelectedMenu = "Late Pay Debit Note Register" Then gCmd.Parameters("@RegTp") = "SI"
             If gSelectedMenu = "Late Pay Debit Note Register (Mill Bill)" Then gCmd.Parameters("@RegTp") = "MI"
             gCmd.Execute
             If ChkConsolidate.Value = 1 Then
                If txtTemp(2).text <> "" Then
                   X = "{TmpGentbl.PartyName}= '" & txtTemp(2) & "' and  {TmpGentbl.UserName}= '" & gUserName & "' "
                Else
                   X = "{TmpGentbl.UserName}= '" & gUserName & "' "
                End If
             Else
                If txtTemp(2).text <> "" Then
                   X = "{TmpGentbl.PartyName}= '" & txtTemp(2) & "' and {TmpGentbl.UserName}= '" & gUserName & "' and {TmpGentbl.Vfirm}='" & gCCode & "'   "
                Else
                   X = "{TmpGentbl.UserName}= '" & gUserName & "' and {TmpGentbl.Vfirm}='" & gCCode & "'   "
                End If
             End If
             .SelectionFormula = X
             .Formulas(4) = "wFrToDt = 'From : ' & '" & dtpFrDt.Value & "' & '  To ' & '" & dtpToDt.Value & "'"
             If gSelectedMenu = "Late Pay Debit Note Register" Then .Formulas(5) = "wRptName = 'Late Pay Debit Note Register'"
             If gSelectedMenu = "Late Pay Debit Note Register (Mill Bill)" Then .Formulas(5) = "wRptName = 'Late Payment Debit Note Register (Mill Bill)'"
             If gSelectedMenu = "Late Pay Debit Note Register" Then
                .ReportFileName = gReportPath & "rptRegisterLPInt.rpt"
             Else
                .ReportFileName = gReportPath & "rptRegisterLPIntMillBill.rpt"
             End If
        Case "Delivery Register"
             ShowRepo = True
             If txtTemp(2) <> "" Then
                wPartyCode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(2), "S", "AcCode", "N")
             Else
                wPartyCode = 0
             End If
             If txtTemp(3) <> "" Then
                wBrkCode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(3), "S", "AcCode", "N")
             Else
                wBrkCode = 0
             End If
             If txtTemp(0) <> "" Then
                wMillCode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(0), "S", "AcCode", "N")
             Else
                wMillCode = 0
             End If
             gCmd.CommandText = "PrcPrepareGodownDelivery"
             gCmd.Parameters.Refresh
             gCmd.Parameters("@Fr_dt") = Format(dtpFrDt, "yyyy/mm/dd")
             gCmd.Parameters("@To_dt") = Format(dtpToDt, "yyyy/mm/dd")
             gCmd.Parameters("@UCode") = gUserName
             gCmd.Parameters("@GodCd") = "" '--Trim(gCGodCd)
             If ChkConsolidate.Value = 1 Then
                gCmd.Parameters("@VFirm") = ""
             Else
                gCmd.Parameters("@VFirm") = Trim(gCCode)
             End If
             If wPartyCode <> 0 Then gCmd.Parameters("@Accode") = wPartyCode
             If wMillCode <> 0 Then gCmd.Parameters("@MillCode") = wMillCode
             If wItemCode <> 0 Then gCmd.Parameters("@ItemCode") = wItemCode
             gCmd.Execute
             x1 = "{TmpGentbl.UserName}='" & gUserName & "'"
             .SelectionFormula = x1
             .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt, "dd/MM/yyyy") & "'"
             
             .ReportFileName = gReportPath & "rptGodownDelivery.rpt"
        Case "Pending Despatch"
             ShowRepo = True
             If txtTemp(2) <> "" Then
                wPartyCode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(2), "S", "AcCode", "N")
             Else
                wPartyCode = 0
             End If
             If txtTemp(3) <> "" Then
                wBrkCode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(3), "S", "AcCode", "N")
             Else
                wBrkCode = 0
             End If
             If txtTemp(0) <> "" Then
                wMillCode = GProcGetColumnValue("tblMastAccount", "AcName", txtTemp(0), "S", "AcCode", "N")
             Else
                wMillCode = 0
             End If
             gCmd.CommandText = "PrcPrepareGodownDeliveryPending"
             gCmd.Parameters.Refresh
             gCmd.Parameters("@Fr_dt") = Format(dtpFrDt, "yyyy/mm/dd")
             gCmd.Parameters("@To_dt") = Format(dtpToDt, "yyyy/mm/dd")
             gCmd.Parameters("@UCode") = gUserName
             gCmd.Parameters("@GodCd") = "" '--Trim(gCGodCd)
             If ChkConsolidate.Value = 1 Then
                gCmd.Parameters("@VFirm") = ""
             Else
                gCmd.Parameters("@VFirm") = Trim(gCCode)
             End If
             gCmd.Parameters("@Rptp") = "P"
             If wPartyCode <> 0 Then gCmd.Parameters("@Accode") = wPartyCode
             If wMillCode <> 0 Then gCmd.Parameters("@MillCode") = wMillCode
             If wItemCode <> 0 Then gCmd.Parameters("@ItemCode") = wItemCode
             gCmd.Execute
             x1 = "{TmpGentbl.UserName}='" & gUserName & "'"
             .SelectionFormula = x1
             .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(dtpFrDt.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(dtpToDt, "dd/MM/yyyy") & "'"
             .ReportFileName = gReportPath & "rptGodownDeliveryPend.rpt"
        Case "Planed Payment List"
             ShowRepo = True
             gCn.BeginTrans
             gCn.Execute "delete from tmpGenTbl where username='" & gUserName & "' "
             gCn.CommitTrans
             '--For Depo Sale
             X = "Select dbo.tblSale.Vfirm,dbo.tblMastAccount.AcOurCode,dbo.tblMastAccount.AcName,dbo.tblMastAccount.AcRemark ,dbo.tblSale.Vdt,iif(dbo.tblSale.DueDays+dbo.tblSale.SlCrDays = 0 and dbo.tblSale.VType = 'SD','*',''),dbo.tblSale.SlTmpChr2,sum(dbo.tblSale.SlBillAmt-dbo.tblSale.SlTdsAmt)-SlBillRecdAmt,sum(dbo.tblSale.SlTdsOnAmt),sum(dbo.tblSale.SlTdsAmt),sum(dbo.tblSale.SlBillAmt),dbo.tblMastAccount.Acsrno,'" & gUserName & "' " _
                 & " From dbo.tblSale inner join dbo.tblMastAccount  on dbo.tblSale.SlAcCrCode = dbo.tblMastAccount.AcCode where  dbo.tblSale.VType in ('SM','SD') and (dbo.tblSale.SlTmpChr1 is null or dbo.tblSale.SlTmpChr1 in ('P','C') or dbo.tblSale.SlTmpChr1 = '') and dbo.tblSale.SlIsHank = 0 And tblMastAccount.Acsrno > 0 " _
                 & " group by dbo.tblSale.Vfirm,dbo.tblMastAccount.AcOurCode,dbo.tblMastAccount.AcName,dbo.tblMastAccount.AcRemark,dbo.tblSale.Vdt,iif(dbo.tblSale.DueDays+dbo.tblSale.SlCrDays = 0 and dbo.tblSale.VType = 'SD','*',''),dbo.tblSale.SlTmpChr2,SlBillRecdAmt,dbo.tblMastAccount.Acsrno "
             
             gCn.BeginTrans
             gCn.Execute "insert into tmpGentbl (Vfirm,MillCode,MillName,Nar1,Vdt,Nar2,Nar3,Amt1,Amt2,Amt3,Amt4,VSrNo,UserName) " & X
             gCn.CommitTrans
             '-- For Commission & brokerage Bill
             X = "Select dbo.tblSale.Vfirm,dbo.tblMastAccount.AcOurCode,dbo.tblMastAccount.AcName,dbo.tblMastAccount.AcRemark ,dbo.tblSale.Vdt,'*',dbo.tblSale.SlCFromNote,(sum(dbo.tblSale.SlBillAmt-dbo.tblSale.SlTdsAmt)-SlBillRecdAmt)*-1,dbo.tblMastAccount.Acsrno,'" & gUserName & "' " _
                 & " From dbo.tblSale inner join dbo.tblMastAccount  on dbo.tblSale.SlAcDrCode = dbo.tblMastAccount.AcCode where  dbo.tblSale.VType = 'SB' and (isnull(dbo.tblSale.SlCFormNo,'') = '' or dbo.tblSale.SlCFormNo in ('P','C') ) " _
                 & " and dbo.tblSale.SlIsHank = 0  and tblMastAccount.Acsrno > 0 and tblMastAccount.AcIsVATDbnt = 1 " _
                 & " group by dbo.tblSale.Vfirm,dbo.tblMastAccount.AcOurCode,dbo.tblMastAccount.AcName,dbo.tblMastAccount.AcRemark,dbo.tblSale.Vdt,dbo.tblSale.SlCFromNote,SlBillRecdAmt,dbo.tblMastAccount.Acsrno "
             gCn.BeginTrans
             gCn.Execute "insert into tmpGentbl (Vfirm,MillCode,MillName,Nar1,Vdt,Nar2,Nar3,Amt1,VSrNo,UserName) " & X
             gCn.CommitTrans
             '-- For Trade Purchase
             X = "Select dbo.tblPurch.Vfirm,dbo.tblMastAccount.AcOurCode,dbo.tblMastAccount.AcName,dbo.tblMastAccount.AcRemark,dbo.tblPurch.PurBillDt,dbo.tblPurch.PurTmpChr2,sum(dbo.tblPurch.PurBillAmt-dbo.tblPurch.PurTdsAmt)-PurBillPaidAmt,sum(dbo.tblPurch.PurTdsOnAmt),sum(dbo.tblPurch.PurTdsAmt),sum(dbo.tblPurch.PurBillAmt),dbo.tblMastAccount.Acsrno,'" & gUserName & "' " _
                 & " From dbo.tblPurch inner join dbo.tblMastAccount  on dbo.tblPurch.PurAcCrCode = dbo.tblMastAccount.AcCode where  dbo.tblPurch.VType = 'PY' and (dbo.tblPurch.PurTmpChr1 is null or dbo.tblPurch.PurTmpChr1 in ('P','C') or dbo.tblPurch.PurTmpChr1 = '') and " _
                 & " tblMastAccount.Acsrno >= 0 group by dbo.tblPurch.Vfirm,dbo.tblMastAccount.AcOurCode,dbo.tblMastAccount.AcName,dbo.tblMastAccount.AcRemark,dbo.tblPurch.PurBillDt,dbo.tblPurch.PurTmpChr2,PurBillPaidAmt,dbo.tblMastAccount.Acsrno "
             gCn.BeginTrans
             gCn.Execute "insert into tmpGentbl (Vfirm,MillCode,MillName,Nar1,Vdt,Nar3,Amt1,Amt2,Amt3,Amt4,VSrNo,UserName) " & X
             gCn.CommitTrans
             X = "{TmpGentbl.UserName}='" & gUserName & "' "
             If txt(0).text <> "" Then X = X + " and left({TmpGentbl.Nar3},3) = '" & txt(0) & "'"
             .SelectionFormula = X
             .ReportFileName = gReportPath & "rptPlanedPaymentList.rpt"
        End Select
     If ShowRepo = True Then
       .Action = 1
       .PageZoom (120)
     End If
End With
End Sub
