VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmRpTSMS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Report"
   ClientHeight    =   8925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9885
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRptSMS.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8925
   ScaleWidth      =   9885
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtLate 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   4440
      TabIndex        =   52
      Top             =   4320
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame frameSMSdetail 
      Height          =   3795
      Left            =   300
      TabIndex        =   33
      Top             =   4800
      Visible         =   0   'False
      Width           =   7155
      Begin VB.Line Line4 
         X1              =   60
         X2              =   7140
         Y1              =   780
         Y2              =   780
      End
      Begin VB.Label smsLineHeading 
         Caption         =   "SMS To All Parties whose 'C' Forms Are Pending"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   360
         Width           =   6435
      End
      Begin VB.Label smsLine1 
         Caption         =   "Pending C Form Details"
         Height          =   255
         Left            =   1980
         TabIndex        =   40
         Top             =   1140
         Width           =   4335
      End
      Begin VB.Label smsLine2 
         Caption         =   "Period :"
         Height          =   255
         Left            =   1980
         TabIndex        =   39
         Top             =   1440
         Width           =   4335
      End
      Begin VB.Label smsLine3 
         Caption         =   "Amt -"
         Height          =   255
         Left            =   1980
         TabIndex        =   38
         Top             =   1740
         Width           =   4335
      End
      Begin VB.Label smsLine4 
         Caption         =   "No Of Bills"
         Height          =   255
         Left            =   1980
         TabIndex        =   37
         Top             =   2040
         Width           =   4335
      End
      Begin VB.Label smsLine5 
         Caption         =   "Firm -"
         Height          =   255
         Left            =   1980
         TabIndex        =   36
         Top             =   2340
         Width           =   4335
      End
      Begin VB.Label smsLine6 
         Caption         =   "MBL -"
         Height          =   255
         Left            =   1980
         TabIndex        =   35
         Top             =   2640
         Width           =   4335
      End
      Begin VB.Label smsLine7 
         Caption         =   "Pl Send C Form"
         Height          =   255
         Left            =   1980
         TabIndex        =   34
         Top             =   2940
         Width           =   4335
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Standard SMS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2715
      Left            =   240
      TabIndex        =   22
      Top             =   2100
      Width           =   9315
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Bank Details Universal Synthetic"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   5040
         TabIndex        =   50
         Top             =   1920
         Width           =   4155
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "SMS FOR PAN NO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   5040
         TabIndex        =   49
         Top             =   2280
         Width           =   4155
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Pending E-1 Forms Reminder"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   5040
         TabIndex        =   48
         Top             =   1560
         Width           =   4155
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "City Union UTC"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   5040
         TabIndex        =   47
         Top             =   1140
         Width           =   4155
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "E-Mail Id Requiest"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   5040
         TabIndex        =   32
         Top             =   720
         Width           =   4155
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Bank Balance"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   5040
         TabIndex        =   31
         Top             =   360
         Width           =   4155
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "C Form Form Reminder"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   180
         TabIndex        =   30
         Top             =   1920
         Width           =   4155
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "TIN Details - Universal  Spintex"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   180
         TabIndex        =   29
         Top             =   2280
         Visible         =   0   'False
         Width           =   4155
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "TIN Details - Universal Textile Corporation"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   180
         TabIndex        =   26
         Top             =   1560
         Width           =   4155
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Firm Address,Phone No And Email Id."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   180
         TabIndex        =   25
         Top             =   1140
         Width           =   4155
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "HDFC Details For Universal Textile Corp."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   24
         Top             =   720
         Width           =   4155
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Blank"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   23
         Top             =   360
         Width           =   4155
      End
   End
   Begin VB.CheckBox chkQtr4 
      BackColor       =   &H00EFFCFB&
      Caption         =   "4 th Quarter"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3300
      TabIndex        =   46
      Top             =   3900
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.CheckBox chkQtr3 
      BackColor       =   &H00EFFCFB&
      Caption         =   "3 rd Quarter"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3300
      TabIndex        =   45
      Top             =   3360
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.CheckBox chkQtr2 
      BackColor       =   &H00EFFCFB&
      Caption         =   "2 nd Quarter"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3300
      TabIndex        =   44
      Top             =   2880
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.CheckBox chkQtr1 
      BackColor       =   &H00EFFCFB&
      Caption         =   "1 st Quarter"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3300
      TabIndex        =   43
      Top             =   2400
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "View"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   7500
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   6360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtTemp 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3195
      Index           =   3
      Left            =   720
      MaxLength       =   320
      MultiLine       =   -1  'True
      OLEDropMode     =   1  'Manual
      TabIndex        =   9
      Top             =   5100
      Width           =   5655
   End
   Begin VB.TextBox txtTemp 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   8
      Left            =   6660
      MaxLength       =   15
      TabIndex        =   8
      Top             =   1575
      Width           =   2655
   End
   Begin VB.TextBox txtTemp 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   7
      Left            =   6660
      MaxLength       =   21
      TabIndex        =   6
      Top             =   1125
      Width           =   2655
   End
   Begin VB.TextBox txtTemp 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   6
      Left            =   6660
      MaxLength       =   15
      TabIndex        =   4
      Top             =   690
      Width           =   2655
   End
   Begin VB.TextBox txtTemp 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   5
      Left            =   6720
      MaxLength       =   70
      TabIndex        =   11
      Top             =   7980
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtTemp 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   4
      Left            =   9540
      MaxLength       =   70
      TabIndex        =   10
      Top             =   8340
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.TextBox txtTemp 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1620
      TabIndex        =   3
      Top             =   720
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.TextBox txtTemp 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   1620
      TabIndex        =   7
      Top             =   1605
      Width           =   4335
   End
   Begin VB.TextBox txtTemp 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1620
      TabIndex        =   5
      Top             =   1155
      Visible         =   0   'False
      Width           =   4335
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
      Left            =   7500
      TabIndex        =   13
      Top             =   7620
      Width           =   915
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Send"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   7500
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6960
      Width           =   975
   End
   Begin VB.CommandButton cmdPrint 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   9420
      Picture         =   "frmRptSMS.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Print"
      Top             =   6480
      Visible         =   0   'False
      Width           =   195
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   8880
      Top             =   7200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1635
      TabIndex        =   1
      Top             =   225
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
      Format          =   112328707
      CurrentDate     =   38050
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   4410
      TabIndex        =   2
      Top             =   225
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
      Format          =   112328707
      CurrentDate     =   38050
   End
   Begin MSComctlLib.ProgressBar prgBar 
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   8580
      Visible         =   0   'False
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Outstanding Morer Than late Days    :"
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
      Left            =   660
      TabIndex        =   51
      Top             =   4380
      Visible         =   0   'False
      Width           =   3525
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SMS Length"
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
      Left            =   7200
      TabIndex        =   28
      Top             =   5160
      Width           =   1155
   End
   Begin VB.Label lblSmsLength 
      Alignment       =   2  'Center
      BackColor       =   &H00EFFCFB&
      Caption         =   "SMS Length"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   7200
      TabIndex        =   27
      Top             =   5580
      Width           =   1395
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000040C0&
      X1              =   9630
      X2              =   120
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile No "
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
      Left            =   7380
      TabIndex        =   21
      Top             =   240
      Width           =   1020
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SMS"
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
      Left            =   240
      TabIndex        =   20
      Top             =   4800
      Width           =   435
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000040C0&
      X1              =   9600
      X2              =   90
      Y1              =   2115
      Y2              =   2115
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00BDDFD1&
      FillStyle       =   0  'Solid
      Height          =   2145
      Left            =   6840
      Shape           =   4  'Rounded Rectangle
      Top             =   6090
      Width           =   2160
   End
   Begin VB.Label lblMillName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Mill              :"
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
      Left            =   240
      TabIndex        =   15
      Top             =   720
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Broker         :"
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
      Left            =   240
      TabIndex        =   17
      Top             =   1605
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Item            :"
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
      Left            =   240
      TabIndex        =   16
      Top             =   1155
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To Date   :"
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
      Left            =   3315
      TabIndex        =   14
      Top             =   225
      Width           =   975
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
      Left            =   240
      TabIndex        =   0
      Top             =   225
      Width           =   1230
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   8340
      Left            =   120
      Top             =   60
      Width           =   9540
   End
End
Attribute VB_Name = "frmRpTSMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer
Dim wMillCode As Long, wBrkCode As Long, wItemCode As Long

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdView_Click(Index As Integer)

Dim X As String
If txtTemp(0) = "" Then wMillCode = 0
If txtTemp(1) = "" Then wItemCode = 0
If txtTemp(2) = "" Then wBrkCode = 0
Dim FrDt As Date
Dim ToDt As Date
FrDt = Format(DTPicker1.Value, "dd/MM/yyyy")
ToDt = Format(DTPicker2.Value, "dd/MM/yyyy")


Select Case gSelectedMenu
       Case "Outstanding To Party"
            GProcCollectCommonOutstanding FrDt, ToDt, txtTemp(0), txtTemp(2), txtTemp(1), True, False
            With CrystalReport1
                GProcCrystalRptPreparation CrystalReport1, IIf(Index = 0, "Window", "Printer")
                CrystalReport1.DataFiles(0) = App.Path & "\HIData\setup.mdb"
                X = "{tmptbloutstanding.Username}='" & gUserName & "'"
                If txtLate.text <> "" Then
                    If CDbl(txtLate.text) > 0 Then
                        X = X + " and {@late} >= " & CDbl(txtLate.text)
                    End If
                End If
                X = X + " and {tmptbloutstanding.OutPartyPhoneSMS} <> ''"
                .SelectionFormula = X
                gReportPath = App.Path & "\HiReports\"
                .ReportFileName = gReportPath & "rptSMSOsParty.rpt"
                CrystalReport1.Formulas(4) = "wFrToDt = 'From : ' & '" & Format(DTPicker1.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(DTPicker2.Value, "dd/MM/yyyy") & "'"
                .Action = 1
             End With
       Case "Outstanding To Broker"
            GProcCollectCommonOutstanding FrDt, ToDt, txtTemp(0), txtTemp(2), txtTemp(1), True, False
            With CrystalReport1
                GProcCrystalRptPreparation CrystalReport1, IIf(Index = 0, "Window", "Printer")
                CrystalReport1.DataFiles(0) = App.Path & "\HIData\setup.mdb"
                X = "{tmptbloutstanding.OutBrokerPhoneSMS} <> '' and  {tmptbloutstanding.Username}='" & gUserName & "'"
                If txtLate.text <> "" Then
                    If CDbl(txtLate.text) > 0 Then
                        X = X + " and {@late} >= " & CDbl(txtLate.text)
                    End If
                End If
                .SelectionFormula = X
                gReportPath = App.Path & "\HiReports\"
                .ReportFileName = gReportPath & "rptSMSOsBroker.rpt"
                CrystalReport1.Formulas(4) = "wFrToDt = 'From : ' & '" & Format(DTPicker1.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(DTPicker2.Value, "dd/MM/yyyy") & "'"
                .Action = 1
             End With
       
       Case "Pending C form Receivable", "Form 'C' Receivable SMS (Trade/SIT) Single Party", "Form 'C' Receivable SMS (Trade/SIT) Comman SMS to All Party", "Form 'E1' Receivable SMS  All Party", "Form 'E1' Receivable SMS  Single Party", "Form 'E2' Receivable SMS  All Party", "Form 'E2' Receivable SMS  Single Party", "Pending C form Receivable All Broker", "Form 'E1' Receivable SMS (Trade/SIT) All Broker", "Form 'E1' Receivable SMS (Trade/SIT) Single Broker", "Form 'E1' Receivable SMS Comman SMS to All Party"
            Dim CondStr As String
            Dim AllFirmFlag As Boolean
            AllFirmFlag = False
            
            If gSelectedMenu = "Form 'E1' Receivable SMS (Trade/SIT) All Broker" Or gSelectedMenu = "Form 'E1' Receivable SMS (Trade/SIT) Single Broker" And gSelectedMenu = "Form 'E2' Receivable SMS  Single Party" Or gSelectedMenu = "Form 'E2' Receivable SMS  All Party" Or gSelectedMenu = "Form 'E1' Receivable SMS  Single Party" Or gSelectedMenu = "Form 'E1' Receivable SMS  All Party" Or gSelectedMenu = "Form 'E1' Receivable SMS Comman SMS to All Party" Then
               CondStr = " tbladdLess.AdGpDt>= cdate('" & Format(DTPicker1.Value, "dd/MM/yyyy") & "')  and tbladdLess.AdGpDt<= cdate('" & Format(DTPicker2.Value, "dd/MM/yyyy") & "')"
               CondStr = CondStr + " and (ADE1FORMNO ='' or ADE1FORMNO is null)" ' (ADCFORMNO is null or ADCFORMNO='') " '-
            Else
               CondStr = " tbladdLess.vdt>= cdate('" & Format(DTPicker1.Value, "dd/MM/yyyy") & "')  and tbladdLess.vdt<= cdate('" & Format(DTPicker2.Value, "dd/MM/yyyy") & "')"
               CondStr = CondStr + " and (ADCFORMNO='' or ADCFORMNO is null)" ' (ADCFORMNO is null or ADCFORMNO='') " '-
            End If
            If gSelectedMenu = "Form 'C' Receivable SMS (Trade/SIT) Single Party" Then
               CondStr = CondStr + " and TblMastAccount.AcName='" & txtTemp(1).text & "'"  '--Party
            End If
            If gSelectedMenu = "Form 'E1' Receivable SMS  Single Party" Then
               CondStr = CondStr + " and TblMastAccount.AcName='" & txtTemp(1).text & "'"  '--Party
            End If
            If gSelectedMenu = "Form 'E2' Receivable SMS  Single Party" Then
               CondStr = CondStr + " and TblMastAccount.AcName='" & txtTemp(1).text & "'"  '--Party
            End If
            If gSelectedMenu = "Pending C form Receivable Single Broker" Or gSelectedMenu = "Form 'E1' Receivable SMS (Trade/SIT) Single Broker" Then
               CondStr = CondStr + " and TblBrk.AcName='" & txtTemp(2).text & "'"  '--Broker
            End If
            
            If gSelectedMenu = "Form 'E1' Receivable SMS  All Party" Or gSelectedMenu = "Form 'E1' Receivable SMS  Single Party" Or gSelectedMenu = "Form 'E1' Receivable SMS Comman SMS to All Party" Then
                 CondStr = CondStr + " and not (ADCFORMNO ='' or ADCFORMNO is null)" ' (ADCFORMNO is null or ADCFORMNO='') " '-

                X = " select tbladdLess.VType,tblAddless.AdFormE1RecdDate,ADINVNO,tbladdLess.AdGpdt,ADBILLAMT,TblMastAccount.AcName,tblBrk.AcName,TblMastAccount.AcFaxNo,tblMill.AcName,'" & gUserName & "'" _
                & " ,ADE1FORMNO,TblMastAccount.AcCst,TblMastAccount.AcContno,tblmastaccount.accity,tblmastaccount.acADD1 from tbladdLess,tblmastaccount,tblmastaccount tblBrk,tblmastaccount  tblMill,tblOutstanding,tblMastNarration where" _
                & " tblOutstanding.vNo=tbladdLess.vno and tblOutstanding.vType=tbladdLess.vType and tblOutstanding.vdt=tbladdLess.vdt" _
                & " and ADACCRCODE=tblmastaccount.Accode and ADMILLCODE=tblMill.AcCode and " _
                & " OUTBROKERCODE=tblBrk.AcCode and ADTAXCODE=NARRCODE and (instr(narration,'C Form')>0 ) and tbladdLess.Vtype in ('PT') and tbladdless.AdGpNo = 0 and " & CondStr
            
            ElseIf gSelectedMenu = "Form 'E2' Receivable SMS  All Party" Or gSelectedMenu = "Form 'E2' Receivable SMS  Single Party" Then
        
                X = " select tbladdLess.VType,tblAddless.AdFormE1RecdDate,ADINVNO,tbladdLess.AdGpdt,ADBILLAMT,TblMastAccount.AcName,tblBrk.AcName,TblMastAccount.AcFaxNo,tblMill.AcName,'" & gUserName & "'" _
                & " ,ADE1FORMNO,TblMastAccount.AcCst,TblMastAccount.AcContno,tblmastaccount.accity,tblmastaccount.acADD1 from tbladdLess,tblmastaccount,tblmastaccount tblBrk,tblmastaccount  tblMill,tblOutstanding,tblMastNarration where" _
                & " tblOutstanding.vNo=tbladdLess.vno and tblOutstanding.vType=tbladdLess.vType and tblOutstanding.vdt=tbladdLess.vdt" _
                & " and ADACCRCODE=tblmastaccount.Accode and ADMILLCODE=tblMill.AcCode and " _
                & " OUTBROKERCODE=tblBrk.AcCode and ADTAXCODE=NARRCODE and (instr(narration,'C Form')>0 ) and tbladdLess.Vtype in ('PT') and tbladdless.AdGpNo = 1 and " & CondStr
        
            ElseIf gSelectedMenu = "Pending C form Receivable All Broker" Then
                       X = "select  '',cdate('" & Format(DTPicker1.Value, "dd/MM/yyyy") & "'),'',cdate('" & Format(DTPicker2.Value, "dd/MM/yyyy") & "'),sum(ADBILLAMT),tblBrk.AcName,TblMastAccount.AcName,tblBrk.AcFaxNo,'','" & gUserName & "' from  tbladdLess,tblmastaccount,tblmastaccount tblBrk,tblOutstanding,tblMastNarration" _
                       & " where tblOutstanding.vNo=tbladdLess.vno and tblOutstanding.vType=tbladdLess.vType and tblOutstanding.vdt=tbladdLess.vdt and ADACDRCODE=tblmastaccount.Accode  and  OUTBROKERCODE=tblBrk.AcCode " _
                       & " and ADTAXCODE=NARRCODE and (instr(narration,'C Form')>0 or instr(narration,'C S T')>0)  and tbladdLess.Vtype in ('SY','ST','SD','SO') and  " & CondStr & "" _
                       & " GROUP BY tblBrk.AcName, TblMastAccount.AcName, tblBrk.AcFaxNo"
                         
            ElseIf gSelectedMenu = "Form 'E1' Receivable SMS (Trade/SIT) All Broker" Or gSelectedMenu = "Form 'E1' Receivable SMS (Trade/SIT) Single Broker" Then
                      CondStr = CondStr + " and not (ADCFORMNO ='' or ADCFORMNO is null)" ' (ADCFORMNO is null or ADCFORMNO='') " '-
                         
                       X = "select  '',cdate('" & Format(DTPicker1.Value, "dd/MM/yyyy") & "'),'',cdate('" & Format(DTPicker2.Value, "dd/MM/yyyy") & "'),sum(ADBILLAMT),tblBrk.AcName,TblMastAccount.AcName,tblBrk.AcFaxNo,'','" & gUserName & "' from  tbladdLess,tblmastaccount,tblmastaccount tblBrk,tblOutstanding,tblMastNarration" _
                       & " where tblOutstanding.vNo=tbladdLess.vno and tblOutstanding.vType=tbladdLess.vType and tblOutstanding.vdt=tbladdLess.vdt and ADACcRCODE=tblmastaccount.Accode  and  OUTBROKERCODE=tblBrk.AcCode " _
                       & " and ADTAXCODE=NARRCODE and (instr(narration,'C Form')>0 ) and tbladdLess.Vtype in ('PT') and tbladdless.AdGpNo = 0 and  " & CondStr & "" _
                       & " GROUP BY tblBrk.AcName, TblMastAccount.AcName, tblBrk.AcFaxNo"
            Else
                X = " select tbladdLess.VType,tblAddless.AdFormRecdDate,ADINVNO,tbladdLess.Vdt,ADBILLAMT,TblMastAccount.AcName,tblBrk.AcName,TblMastAccount.AcFaxNo,tblMill.AcName,'" & gUserName & "'" _
                & " ,ADCFORMNO,TblMastAccount.AcCst,TblMastAccount.AcContno,tblmastaccount.accity,tblmastaccount.acADD1 from tbladdLess,tblmastaccount,tblmastaccount tblBrk,tblmastaccount  tblMill,tblOutstanding,tblMastNarration where" _
                & " tblOutstanding.vNo=tbladdLess.vno and tblOutstanding.vType=tbladdLess.vType and tblOutstanding.vdt=tbladdLess.vdt" _
                & " and ADACDRCODE=tblmastaccount.Accode and ADMILLCODE=tblMill.AcCode and " _
                & " OUTBROKERCODE=tblBrk.AcCode and ADTAXCODE=NARRCODE and (instr(narration,'C Form')>0 or instr(narration,'C S T')>0)     and tbladdLess.Vtype in ('SY','ST','SD','SO')  and " & CondStr
            End If
            '--- insert to setup db Table For All Firm
            GProcInsertToSetupDbTbl AllFirmFlag, X, "tmpGenTbl", "Vtype,dt1,VbillNo,VBillDt,Amt1,Nar2,Nar1,Nar11,Nar4,username,Nar6,nar7,nar8,nar9,nar10 ", " username='" & gUserName & "'", "Nar3", "Nar5"
            
            With CrystalReport1
                GProcCrystalRptPreparation CrystalReport1, IIf(Index = 0, "Window", "Printer")
                .Formulas(4) = "wFrToDt = 'From : ' & '" & Format(DTPicker1.Value, "dd/MM/yyyy") & "' & '  To ' & '" & Format(DTPicker2.Value, "dd/MM/yyyy") & "'"
                .Formulas(15) = "wCSTNo = '" & gCCST & "'"
                .Formulas(16) = "wBSTNo = '" & gCBST & "'"
                .Formulas(17) = "wPanNo = '" & gCPAN & "'"
                Select Case gSelectedMenu
                Case "Pending C form Receivable"
                     .Formulas(18) = "wRrportHead = 'SMS List For Pending C Form To All Party'"
                Case "Form 'C' Receivable SMS (Trade/SIT) Single Party"
                     .Formulas(18) = "wRrportHead = 'SMS List For Pending C Form To Single Party'"
                Case "Form 'C' Receivable SMS (Trade/SIT) Comman SMS to All Party"
                     .Formulas(18) = "wRrportHead = 'Comman SMS List For Pending C Form To All Party'"
                Case "Form 'E1' Receivable SMS  All Party"
                     .Formulas(18) = "wRrportHead = 'SMS List For Pending E1 Form To All Party'"
                Case "Form 'E1' Receivable SMS  Single Party"
                     .Formulas(18) = "wRrportHead = 'SMS List For Pending E1 Form To Single Party'"
                Case "Form 'E2' Receivable SMS  All Party"
                     .Formulas(18) = "wRrportHead = 'SMS List For Pending E2 Form To All Party'"
                Case "Form 'E2' Receivable SMS  Single Party"
                     .Formulas(18) = "wRrportHead = 'SMS List For Pending E2 Form To Single Party'"
                Case "Pending C form Receivable All Broker"
                     .Formulas(18) = "wRrportHead = 'SMS List For Pending C form Receivable All Broker'"
                Case "Form 'E1' Receivable SMS (Trade/SIT) All Broker"
                     .Formulas(18) = "wRrportHead = 'Form E1 Receivable SMS (Trade/SIT) All Broker'"
                Case "Form 'E1' Receivable SMS (Trade/SIT) Single Broker"
                     .Formulas(18) = "wRrportHead = 'Form E1 Receivable SMS (Trade/SIT) Single Broker'"
                Case "Form 'E1' Receivable SMS Comman SMS to All Party"
                     .Formulas(18) = "wRrportHead = 'Form E1 Receivable SMS Comman SMS to All Party'"
                End Select
                .DataFiles(0) = App.Path & "\HIData\setup.mdb"
                .SelectionFormula = "{tmpgenTbl.Username}='" & gUserName & "'"
                gReportPath = App.Path & "\HiReports\"
                .ReportFileName = gReportPath & "rptSMSDetails.rpt"
                .Action = 1
            End With
End Select
Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case 13
'     SendKeys "{TAB}"
Case 27
    Unload Me
End Select
End Sub
Private Sub Form_Load()
Dim i As Long
Me.Top = 700
Me.Left = Screen.Width / 2 - Me.Width / 2
 '-- Activate Controls
GProcActivateControls Me, False
DTPicker1.Enabled = True
DTPicker2.Enabled = True
'-- Set Values
DTPicker1.Value = Format(gCYSDate, "dd/MM/yyyy")
DTPicker2.Value = Format(gCYEDate, "dd/MM/yyyy")
'For i = 0 To 2 '-- Graph Type
'    OptGraphType(i).Enabled = True
'Next i
'For i = 0 To 4 '-- Number Format
'    OptNumFormat(i).Enabled = True
'Next i
'ChkConsolidate.Enabled = True
'OptAmt.Enabled = True
'OptBag.Enabled = True
Me.Caption = gSelectedMenu
Select Case gSelectedMenu
       Case "Pending C form Receivable", "Form 'C' Receivable SMS (Trade/SIT) Single Party", "Form 'C' Receivable SMS (Trade/SIT) Comman SMS to All Party", "Form 'E1' Receivable SMS  All Party", "Form 'E1' Receivable SMS  Single Party", "Form 'E2' Receivable SMS  All Party", "Form 'E2' Receivable SMS  Single Party", "Pending C form Receivable All Broker", "Pending C form Receivable Single Broker", "Form 'E1' Receivable SMS (Trade/SIT) All Broker", "Form 'E1' Receivable SMS (Trade/SIT) Single Broker", "Form 'E1' Receivable SMS Comman SMS to All Party"
            For i = 0 To 8
                txtTemp(i).Visible = False
            Next
             
            chkQtr1.Visible = True
            chkQtr1.Enabled = True
            
            chkQtr2.Visible = True
            chkQtr2.Enabled = True
            
            chkQtr3.Visible = True
            chkQtr3.Enabled = True
            
            chkQtr4.Visible = True
            chkQtr4.Enabled = True
            
            DTPicker1.Visible = False
            DTPicker2.Visible = False
            Label1.Visible = False
            Label2.Visible = False
            
            Frame1.Visible = False
            Label6.Visible = False
            Label4.Visible = False
            Label7.Visible = False
            
            Line3.Visible = False
            Line1.Visible = False
            lblSmsLength.Visible = False
            cmdView(0).Visible = True
            cmdView(0).Enabled = True
            Option1(11).Visible = False
'            frameSMSdetail.Visible = True
            
            If gSelectedMenu = "Form 'C' Receivable SMS (Trade/SIT) Single Party" Then
               Label6.Visible = True
               txtTemp(1).Visible = True
               txtTemp(1).Enabled = True
               txtTemp(7).Visible = True
               txtTemp(7).Enabled = True
               Label5.Visible = True
               Label5.Caption = "Party  : "
            End If
            If gSelectedMenu = "Form 'E1' Receivable SMS  Single Party" Then
               Label6.Visible = True
               txtTemp(1).Visible = True
               txtTemp(1).Enabled = True
               txtTemp(7).Visible = True
               txtTemp(7).Enabled = True
               Label5.Visible = True
               Label5.Caption = "Party  : "
            End If
            If gSelectedMenu = "Form 'E2' Receivable SMS  Single Party" Then
               Label6.Visible = True
               txtTemp(1).Visible = True
               txtTemp(1).Enabled = True
               txtTemp(7).Visible = True
               txtTemp(7).Enabled = True
               Label5.Visible = True
               Label5.Caption = "Party  : "
            End If
            If gSelectedMenu = "Pending C form Receivable All Broker" Then
                chkQtr1.Visible = False
                chkQtr1.Enabled = False
                
                chkQtr2.Visible = False
                chkQtr2.Enabled = False
                
                chkQtr3.Visible = False
                chkQtr3.Enabled = False
                
                chkQtr4.Visible = False
                chkQtr4.Enabled = False
                DTPicker1.Visible = True
                DTPicker2.Visible = True
            End If
            If gSelectedMenu = "Form 'E1' Receivable SMS (Trade/SIT) All Broker" Then
                chkQtr1.Visible = False
                chkQtr1.Enabled = False
                
                chkQtr2.Visible = False
                chkQtr2.Enabled = False
                
                chkQtr3.Visible = False
                chkQtr3.Enabled = False
                
                chkQtr4.Visible = False
                chkQtr4.Enabled = False
                DTPicker1.Visible = True
                DTPicker2.Visible = True
            End If
            If gSelectedMenu = "Pending C form Receivable Single Broker" Then
               Label6.Visible = True
               Label3.Visible = True
               txtTemp(2).Visible = True
               txtTemp(2).Enabled = True
               txtTemp(8).Visible = True
               txtTemp(8).Enabled = True
               chkQtr1.Visible = False
               chkQtr1.Enabled = False
                
               chkQtr2.Visible = False
               chkQtr2.Enabled = False
                
               chkQtr3.Visible = False
               chkQtr3.Enabled = False
                
               chkQtr4.Visible = False
               chkQtr4.Enabled = False
               DTPicker1.Visible = True
               DTPicker2.Visible = True
            End If
            If gSelectedMenu = "Form 'E1' Receivable SMS (Trade/SIT) Single Broker" Then
               Label6.Visible = True
               Label3.Visible = True
               txtTemp(2).Visible = True
               txtTemp(2).Enabled = True
               txtTemp(8).Visible = True
               txtTemp(8).Enabled = True
               chkQtr1.Visible = False
               chkQtr1.Enabled = False
                
               chkQtr2.Visible = False
               chkQtr2.Enabled = False
                
               chkQtr3.Visible = False
               chkQtr3.Enabled = False
                
               chkQtr4.Visible = False
               chkQtr4.Enabled = False
               DTPicker1.Visible = True
               DTPicker2.Visible = True
            End If
            If gSelectedMenu = "Form 'C' Receivable SMS (Trade/SIT) Comman SMS to All Party" Or gSelectedMenu = "Form 'E1' Receivable SMS Comman SMS to All Party" Then
                chkQtr1.Visible = False
                chkQtr1.Enabled = False
                
                chkQtr2.Visible = False
                chkQtr2.Enabled = False
                
                chkQtr3.Visible = False
                chkQtr3.Enabled = False
                
                chkQtr4.Visible = False
                chkQtr4.Enabled = False
                
                DTPicker1.Visible = True
                DTPicker2.Visible = True
                Label1.Visible = True
                Label2.Visible = True
               lblSmsLength.Visible = True
               Label7.Visible = True
               txtTemp(3).Visible = True
               txtTemp(3).Enabled = True
               If gSelectedMenu = "Form 'E1' Receivable SMS Comman SMS to All Party" Then
                   txtTemp(3).text = "Clear Pending E1 Forms " & vbCrLf
               Else
                   txtTemp(3).text = "Clear Pending C Forms " & vbCrLf
               End If
               txtTemp(3).text = txtTemp(3).text + "For the Period : " + Trim(CStr(Format(DTPicker1.Value, "dd/mm/yyyy"))) + " -" + Trim(CStr(Format(DTPicker2.Value, "dd/mm/yyyy"))) & vbCrLf
               'txtTemp(3).Text = txtTemp(3).Text + "As early as possible" & vbCrLf
               txtTemp(3).text = txtTemp(3).text + vbCrLf
               txtTemp(3).text = txtTemp(3).text + "Firm-" & vbCrLf
               txtTemp(3).text = txtTemp(3).text + gCName & vbCrLf
               txtTemp(3).text = txtTemp(3).text + "Ph-(0230) 2433670" & vbCrLf
               txtTemp(3).text = Left(txtTemp(3).text, 160)
               frameSMSdetail.Visible = False
            
            End If
            '---For SMS Detail
            Select Case gSelectedMenu
            Case "Form 'C' Receivable SMS (Trade/SIT) Single Party"
                  smsLineHeading.Caption = "SMS To Single Party whose 'C' Forms Are Pending"
            Case "Form 'C' Receivable SMS (Trade/SIT) Comman SMS to All Party"
                  smsLineHeading.Caption = "Comman SMS To All Parties whose 'C' Forms Are Pending"
                  smsLine1.Caption = "Pl Clear Pending C Forms"
                  smsLine2.Caption = "For the Period :"
                  smsLine3.Caption = "As early as possible"
                  smsLine4.Caption = ""
                  smsLine5.Caption = "Firm"
                  smsLine6.Caption = "MBL"
                  smsLine7.Caption = ""
            Case "Form 'E1' Receivable SMS  All Party"
                  smsLineHeading.Caption = "SMS To All Parties whose 'E1' Forms Are Pending"
                  smsLine1.Caption = "Pending E1 Form Details"
                  smsLine7.Caption = "Pl Send E1 Form"
            Case "Form 'E1' Receivable SMS  Single Party"
                  smsLineHeading.Caption = "SMS To Single Party whose 'E1' Forms Are Pending"
                  smsLine1.Caption = "Pending E1 Form Details"
                  smsLine7.Caption = "Pl Send E1 Form"
            Case "Form 'E2' Receivable SMS  All Party"
                  smsLineHeading.Caption = "SMS To All Parties whose 'E2' Forms Are Pending"
                  smsLine1.Caption = "Pending E2 Form Details"
                  smsLine7.Caption = "Pl Send E2 Form"
            Case "Form 'E2' Receivable SMS  Single Party"
                  smsLineHeading.Caption = "SMS To Single Party whose 'E2' Forms Are Pending"
                  smsLine1.Caption = "Pending E2 Form Details"
            Case "Form 'E1' Receivable SMS Comman SMS to All Party"
                  smsLineHeading.Caption = "Comman SMS To All Parties whose 'E1' Forms Are Pending"
                  smsLine1.Caption = "Pl Clear Pending E1 Forms"
                  smsLine2.Caption = "For the Period :"
                  smsLine3.Caption = "As early as possible"
                  smsLine4.Caption = ""
                  smsLine5.Caption = "Firm"
                  smsLine6.Caption = "MBL"
                  smsLine7.Caption = ""
            End Select
            '---
            
       Case "General SMS" '== Sales Graph
             lblMillName.Caption = "Name 1  :"
             Label5.Caption = "Name 2  :"
             Label3.Caption = "Name 3  :"
             
             lblMillName.Visible = True
             txtTemp(0).Visible = True
             txtTemp(0).Enabled = True
             Label5.Visible = True
             txtTemp(1).Visible = True
             txtTemp(1).Enabled = True
             
             Label5.Visible = True
             txtTemp(2).Visible = True
             txtTemp(2).Enabled = True
             
             Label1.Visible = False
             Label2.Visible = False
             Label3.Visible = True
             DTPicker1.Visible = False
             DTPicker2.Visible = False
             txtTemp(3).Visible = True
             txtTemp(3).Enabled = True
'             txtTemp(4).Visible = True
'             txtTemp(4).Enabled = True
'             txtTemp(5).Visible = True
'             txtTemp(5).Enabled = True
             txtTemp(6).Visible = True
             txtTemp(6).Enabled = True
             txtTemp(7).Visible = True
             txtTemp(7).Enabled = True
             txtTemp(8).Visible = True
             txtTemp(8).Enabled = True
             Option1(0).Value = True
             Option1(0).Enabled = True
             Option1(1).Enabled = True
             Option1(2).Enabled = True
             Option1(3).Enabled = True
             Option1(4).Enabled = True
             Option1(5).Enabled = True
             Option1(6).Enabled = True
             Option1(7).Enabled = True
             Option1(8).Enabled = True
             Option1(9).Enabled = True
             Option1(10).Enabled = True
             Option1(11).Enabled = True
        Case "Outstanding To Party"
             Frame1.Visible = False
             frameSMSdetail.Visible = False
             txtTemp(3).Visible = False
             txtTemp(0).Visible = False
             txtTemp(0).Enabled = False
             txtTemp(1).Visible = True
             txtTemp(1).Enabled = True
             txtTemp(2).Visible = True
             txtTemp(2).Enabled = True
             
             txtTemp(6).Visible = False
             txtTemp(6).Enabled = False
             txtTemp(7).Visible = True
             txtTemp(7).Enabled = True
             txtTemp(8).Visible = False
             txtTemp(8).Enabled = False
             cmdView(0).Visible = True
             cmdView(0).Enabled = True
             txtLate.Visible = True
             Label8.Visible = True
             txtLate.Enabled = True
             Label4.Visible = False
             Line1.Visible = False
             
             lblMillName.Visible = False
             Label5.Visible = True
             Label3.Visible = True
             Label5.Caption = "Party : "
       Case "Outstanding To Broker"
             Frame1.Visible = False
             frameSMSdetail.Visible = False
             txtTemp(3).Visible = False
             txtTemp(0).Visible = False
             txtTemp(0).Enabled = False
             txtTemp(1).Visible = True
             txtTemp(1).Enabled = True
             txtTemp(2).Visible = True
             txtTemp(2).Enabled = True
             
             txtTemp(6).Visible = False
             txtTemp(6).Enabled = False
             txtTemp(7).Visible = False
             txtTemp(7).Enabled = False
             txtTemp(8).Visible = True
             txtTemp(8).Enabled = True
             cmdView(0).Visible = True
             cmdView(0).Enabled = True
             txtLate.Visible = True
             Label8.Visible = True
             txtLate.Enabled = True
             Label4.Visible = False
             Line1.Visible = False
             
             lblMillName.Visible = False
             Label5.Visible = True
             Label3.Visible = True
             Label5.Caption = "Party : "
End Select
 '-- Set InActivate Ctrl Back Color
GProcSetInActivateCtrlBackClr Me
End Sub
Private Sub cmdPrint_Click(Index As Integer)
'On Error GoTo ErrorRoutine
Dim X As String
If txtTemp(0) = "" Then wMillCode = 0
If txtTemp(1) = "" Then wItemCode = 0
If txtTemp(2) = "" Then wBrkCode = 0




Dim HttpReq As New WinHttpRequest
Dim url As String
Dim params As String
Dim params2 As String
Dim SmsMblNo As String
Dim SmsString As String
Dim SmsSent As Boolean
SmsSent = False

Dim mTmprst1 As Recordset
Dim mTmprst2 As Recordset
Dim PartyName As String
Dim BrokerName As String

Dim CFromAmount As Double
Dim NoOfBills As Integer
Dim NoOfSms As Integer
Dim SrNo  As Integer
NoOfSms = 0

Dim CondStr As String
Dim AllFirmFlag As Boolean
AllFirmFlag = False
Dim QtrArray As Integer
QtrArray = 1
Dim FrDt As Date
Dim ToDt As Date
Dim PartyNm As String
Dim Ldays As Integer
Dim currdt As Date
FrDt = Format(DTPicker1.Value, "dd/MM/yyyy")
ToDt = Format(DTPicker2.Value, "dd/MM/yyyy")
currdt = Format(Now, "dd/MM/yyyy")


'--- For Party O/s SMS

If gSelectedMenu = "Outstanding To Party" Then
   GProcCollectCommonOutstanding FrDt, ToDt, txtTemp(0), txtTemp(2), txtTemp(1), True, False
   If txtLate.text = "" Then
      txtLate.text = "0"
   End If
   
   If GProcRstOpen(mTmprst1, "select * from tmptblOutStanding where  OutPartyPhoneSMS <> '' and  UserName='" & gUserName & "'  order by OutAcName,OutBillNo", "R", gSetupCn) > 0 Then
      
      MsgBox "Start For SMS Send Process..."
   
      With mTmprst1
           mTmprst1.MoveFirst
Lp1:
           PartyNm = mTmprst1.Fields("OutAcName")
           If txtTemp(7).text = "" Then
              SmsMblNo = mTmprst1.Fields("OutPartyPhoneSMS") & vbNullString
           Else
              SmsMblNo = txtTemp(7)
           End If
           SmsString = ""
           Do While Not .EOF
              If !OutAcName <> "" & PartyNm & "" Then
                  GoTo Lp1
              End If
              Ldays = DateDiff("d", !OutDueDate, currdt) - 10
              If Ldays >= CDbl(txtLate.text) Then
                 If (!Vtype = "ST") Then
                     SmsString = SmsString + !OutBillNo + " Rs." + CStr(!OutAmount - !OutRecAmt) + " Late(" + CStr(Ldays) + ")" & vbLf
                 Else
                     SmsString = SmsString + "B.No:" + !OutBillNo + " Rs." + CStr(!OutAmount - !OutRecAmt) + " Late(" + CStr(Ldays) + ")" & vbLf
                 End If
              End If
              mTmprst1.MoveNext
              If mTmprst1.EOF = True Then
                 If SmsString <> "" Then
                    SmsString = "Pymt. Due " & vbLf + "Firm-" + PartyNm & vbLf + SmsString & vbLf
                    SmsString = SmsString + "Firm-" & vbLf
                    SmsString = SmsString + Left(gCName, 22) & vbLf
                    SmsString = SmsString + "Ph- 2433648"
                    If SmsMblNo = "" Then
                    Else
                        url = "http://alerts.sinfini.com/api/web2sms.php"
                        params = "workingkey= 98702uv6wtl7bqnh450&sender=UTCYRN&to=" & SmsMblNo & " &message=" & SmsString & ""
                        
                        HttpReq.Open "POST", url, False
                        HttpReq.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
                        HttpReq.send params
                        NoOfSms = NoOfSms + 1
                    
                    
''                       url = "http://bhashsms.com/api/sendmsg.php"
''                       params = "user=universaltex&pass=123456&sender=UTCYRN&phone=" & SmsMblNo & "&text=" & SmsString & "&priority=ndnd&stype=normal"
''                       HttpReq.Open "POST", url, False
''                       HttpReq.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
''                       HttpReq.Send params
''                       NoOfSms = NoOfSms + 1
'                       delay (3)
                     End If
                 End If
              ElseIf !OutAcName <> "" & PartyNm & "" Then
                 If SmsString <> "" Then
                    SmsString = "Pymt. Due " & vbLf + "Firm-" + PartyNm & vbLf + SmsString & vbLf
                    SmsString = SmsString + "Firm-" & vbLf
                    SmsString = SmsString + Left(gCName, 22) & vbLf
                    SmsString = SmsString + "Ph- 2433648"
                    If SmsMblNo = "" Then
                    Else
                        url = "http://alerts.sinfini.com/api/web2sms.php"
                        params = "workingkey= 98702uv6wtl7bqnh450&sender=UTCYRN&to=" & SmsMblNo & " &message=" & SmsString & ""
                        HttpReq.Open "POST", url, False
                        HttpReq.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
                        HttpReq.send params
                        NoOfSms = NoOfSms + 1
                    
''                       url = "http://bhashsms.com/api/sendmsg.php"
''                       params = "user=universaltex&pass=123456&sender=UTCYRN&phone=" & SmsMblNo & "&text=" & SmsString & "&priority=ndnd&stype=normal"
''                       HttpReq.Open "POST", url, False
''                       HttpReq.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
''                       HttpReq.Send params
''                       NoOfSms = NoOfSms + 1
'                       delay (3)
                    End If
                 End If
              End If
           Loop
      End With
      MsgBox CStr(NoOfSms) + " SMS Send..."
      MsgBox "All SMS Send..."
   End If
   Exit Sub
End If
'--- End For Party O/s SMS


'--- For Broker O/s SMS
If gSelectedMenu = "Outstanding To Broker" Then
   GProcCollectCommonOutstanding FrDt, ToDt, txtTemp(0), txtTemp(2), txtTemp(1), True, False
   If txtLate.text = "" Then
      txtLate.text = "0"
   End If
   
   If GProcRstOpen(mTmprst1, "select * from tmptblOutStanding where  OutBrokerPhoneSMS <> '' and UserName='" & gUserName & "'  order by OutBrokerName,OutAcName,OutBillNo", "R", gSetupCn) > 0 Then
      
      MsgBox "Start For SMS Send Process..."
      
      With mTmprst1
           mTmprst1.MoveFirst
BLp1:
           BrokerName = mTmprst1.Fields("OutBrokerName")
           If txtTemp(8).text = "" Then
              SmsMblNo = mTmprst1.Fields("OutBrokerPhoneSMS") & vbNullString
           Else
              SmsMblNo = txtTemp(8)
           End If
           Do While Not .EOF
              If !OutBrokerName <> "" & BrokerName & "" Then
             
                  GoTo BLp1
              End If
BLp2:
              PartyNm = mTmprst1.Fields("OutAcName")
              SmsString = ""
              Do While Not .EOF
                 If !OutBrokerName <> "" & BrokerName & "" Then
                     GoTo BLp1
                 End If
                 If !OutAcName <> "" & PartyNm & "" Then
                     GoTo BLp2
                 End If
                 Ldays = DateDiff("d", !OutDueDate, currdt) - 10
                 If Ldays >= CDbl(txtLate.text) Then
                     If (!Vtype = "ST") Then
                         SmsString = SmsString + !OutBillNo + " Rs." + CStr(!OutAmount - !OutRecAmt) + " Late(" + CStr(Ldays) + ")" & vbLf
                     Else
                         SmsString = SmsString + "B.No:" + !OutBillNo + " Rs." + CStr(!OutAmount - !OutRecAmt) + " Late(" + CStr(Ldays) + ")" & vbLf
                     End If
                 End If
                 mTmprst1.MoveNext
                 If mTmprst1.EOF = True Then
                     If SmsString <> "" Then
                        SmsString = "Pymt. Due " & vbLf + "Firm-" + PartyNm & vbLf + SmsString & vbLf
                        SmsString = SmsString + "Firm-" & vbLf
                        SmsString = SmsString + gCName & vbLf
                        SmsString = SmsString + "Ph- 2433648"
                        If SmsMblNo = "" Then
                        Else
                        
                        url = "http://alerts.sinfini.com/api/web2sms.php"
                        params = "workingkey= 98702uv6wtl7bqnh450&sender=UTCYRN&to=" & SmsMblNo & " &message=" & SmsString & ""
                        HttpReq.Open "POST", url, False
                        HttpReq.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
                        HttpReq.send params
                        NoOfSms = NoOfSms + 1
                        
''                           url = "http://bhashsms.com/api/sendmsg.php"
''                           params = "user=universaltex&pass=123456&sender=UTCYRN&phone=" & SmsMblNo & "&text=" & SmsString & "&priority=ndnd&stype=normal"
''                           HttpReq.Open "POST", url, False
''                           HttpReq.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
''                           HttpReq.Send params
''                           NoOfSms = NoOfSms + 1
'                           delay (3)
                        End If
                     End If
                  ElseIf !OutAcName <> "" & PartyNm & "" Then
                     If SmsString <> "" Then
                        SmsString = "Pymt. Due " & vbLf + "Firm-" + PartyNm & vbLf + SmsString & vbLf
                        SmsString = SmsString + "Firm-" & vbLf
                        SmsString = SmsString + gCName & vbLf
                        SmsString = SmsString + "Ph- 2433648"
                        If SmsMblNo = "" Then
                        Else
                        
                        url = "http://alerts.sinfini.com/api/web2sms.php"
                        params = "workingkey= 98702uv6wtl7bqnh450&sender=UTCYRN&to=" & SmsMblNo & " &message=" & SmsString & ""
                        HttpReq.Open "POST", url, False
                        HttpReq.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
                        HttpReq.send params
                        NoOfSms = NoOfSms + 1
                        
''                           url = "http://bhashsms.com/api/sendmsg.php"
''                           params = "user=universaltex&pass=123456&sender=UTCYRN&phone=" & SmsMblNo & "&text=" & SmsString & "&priority=ndnd&stype=normal"
''                           HttpReq.Open "POST", url, False
''                           HttpReq.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
''                           HttpReq.Send params
''                           NoOfSms = NoOfSms + 1
'                           delay (3)
                        End If
                     End If
                  End If
              Loop
           Loop
      End With
      MsgBox CStr(NoOfSms) + " SMS Send..."
      MsgBox "All SMS Send..."
   End If
   Exit Sub
End If
'--- End For Broker O/s SMS

If gSelectedMenu = "Form 'C' Receivable SMS (Trade/SIT) Comman SMS to All Party" Then
   SmsString = txtTemp(3).text
   SmsString = Left(SmsString, 360)
   
   CondStr = " tbladdLess.vdt>= cdate('" & Format(DTPicker1.Value, "dd/MM/yyyy") & "')  and tbladdLess.vdt<= cdate('" & Format(DTPicker2.Value, "dd/MM/yyyy") & "')"
   CondStr = CondStr + " and (ADCFORMNO='' or ADCFORMNO is null)" ' (ADCFORMNO is null or ADCFORMNO='') " '-
   
   X = " select tbladdLess.VType,tblAddless.AdFormRecdDate,ADINVNO,tbladdLess.Vdt,ADBILLAMT,TblMastAccount.AcName,tblBrk.AcName,TblMastAccount.AcFaxNo,tblMill.AcName,'" & gUserName & "'" _
       & " ,ADCFORMNO,TblMastAccount.AcCst,TblMastAccount.AcContno,tblmastaccount.accity,tblmastaccount.acADD1 from tbladdLess,tblmastaccount,tblmastaccount tblBrk,tblmastaccount  tblMill,tblOutstanding,tblMastNarration where" _
       & " tblOutstanding.vNo=tbladdLess.vno and tblOutstanding.vType=tbladdLess.vType and tblOutstanding.vdt=tbladdLess.vdt" _
       & " and ADACDRCODE=tblmastaccount.Accode and ADMILLCODE=tblMill.AcCode and " _
       & " OUTBROKERCODE=tblBrk.AcCode and ADTAXCODE=NARRCODE and (instr(narration,'C Form')>0 or instr(narration,'C S T')>0)     and tbladdLess.Vtype in ('SY','ST','SD','SO')  and " & CondStr

   GProcInsertToSetupDbTbl AllFirmFlag, X, "tmpGenTbl", "Vtype,dt1,VbillNo,VBillDt,Amt1,Nar2,Nar1,Nar11,Nar4,username,Nar6,nar7,nar8,nar9,nar10 ", " username='" & gUserName & "'", "Nar3", "Nar5"

   If GProcRstOpen(mTmprst1, "select DISTINCT nar11 from tmpGenTbl where UserName='" & gUserName & "'  order by Nar11", "R", gSetupCn) > 0 Then
      With mTmprst1
           mTmprst1.MoveFirst
           Do While Not .EOF
'SmsString = "Trial 2 SMS Testing " + CStr(NoOfSms)
              
              SmsMblNo = !nar11 & vbNullString
'              SmsMblNo = "9422049083"
              If SmsMblNo = "" Then
'                  MsgBox ("Party Mobile No Not Found....")
              Else

'--- For smscountry
'                    url = "http://www.smscountry.com/sms-api-documentation.aspx"
''                    params = "user=universaltex&apikey=YRPUla1IJxTUU5iKRTQQ&number='" & SmsMblNo & "'&message='" & SmsString & "'&senderid=SMSTXT&type=txt"
'
'                    params = "User=universaltex&passwd=99712453&mobilenumber='" & SmsMblNo & "'&message='" & SmsString & "'&sid=xxxxxxxx&mtype=N&DR=Y"
'                    HttpReq.Open "POST", url, False
'                    HttpReq.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
'                    HttpReq.Send params
'                    NoOfSms = NoOfSms + 1


'--- end

'''--- For FoftSMS
''                    url = "http://smshorizon.co.in/api/sendsms.php?"
''                    params = "user=universaltex&apikey=YRPUla1IJxTUU5iKRTQQ&number=" & SmsMblNo & "&message=" & SmsString & "&senderid=UTCYRN&type=text"
'''                    params = "user=universaltex&apikey=123456&number=" & SmsMblNo & "&message=" & SmsString & "&senderid=SMSTXT&type=txt"
''                    HttpReq.Open "POST", url, False
''                    HttpReq.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
''                    HttpReq.Send params
''                    NoOfSms = NoOfSms + 1
'--- end harizon
'
'
'
'
'
''--- For harizon
'                    url = "http://smshorizon.co.in/api/sendsms.php?"
'                    params = "user=universaltex&apikey=YRPUla1IJxTUU5iKRTQQ&number=" & SmsMblNo & "&message=" & SmsString & "&senderid=SMSTXT&type=txt"
'                    HttpReq.Open "POST", url, False
'                    HttpReq.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
'                    HttpReq.Send params
'                    NoOfSms = NoOfSms + 1
''--- end harizon



'--- For Bhashsms
''                    url = "http://bhashsms.com/api/sendmsg.php"
''                    params = "user=universaltex&pass=123456&sender=UTCYRN&phone=" & SmsMblNo & "&text=" & SmsString & "&priority=ndnd&stype=normal"
''                    HttpReq.Open "POST", url, False
''                    HttpReq.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
''                    HttpReq.Send params
''                    NoOfSms = NoOfSms + 1
''delay (3)
'---end bhash
'


'---For sinfini
                  url = "http://alerts.sinfini.com/api/web2sms.php"
                  params = "workingkey= 98702uv6wtl7bqnh450&sender=UTCYRN&to=" & SmsMblNo & " &message=" & SmsString & ""
                  HttpReq.Open "POST", url, False
                  HttpReq.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
                  HttpReq.send params
                  NoOfSms = NoOfSms + 1

'---
               End If
               mTmprst1.MoveNext
               
'               If NoOfSms >= 50 Then
'                  NoOfSms = NoOfSms + 1
'                   DoEvents
'                   MsgBox CStr(NoOfSms) + " SMS Send..."
'               End If
           Loop
           MsgBox CStr(NoOfSms) + " SMS Send..."
      End With
   End If

ElseIf gSelectedMenu = "Form 'E1' Receivable SMS Comman SMS to All Party" Then
   SmsString = txtTemp(3).text
   SmsString = Left(SmsString, 360)
                                
   CondStr = " tbladdLess.AdGpDt>= cdate('" & Format(DTPicker1.Value, "dd/MM/yyyy") & "')  and tbladdLess.AdGpDt<= cdate('" & Format(DTPicker2.Value, "dd/MM/yyyy") & "')"
   CondStr = CondStr + " and (ADE1FORMNO ='' or ADE1FORMNO is null) and  not ((ADCFORMNO ='' or ADCFORMNO is null))        " ' (ADCFORMNO is null or ADCFORMNO='') " '-
   
   
   X = " select tbladdLess.VType,tblAddless.AdFormE1RecdDate,ADINVNO,tbladdLess.AdGpdt,ADBILLAMT,TblMastAccount.AcName,tblBrk.AcName,TblMastAccount.AcFaxNo,tblMill.AcName,'" & gUserName & "'" _
       & " ,ADE1FORMNO,TblMastAccount.AcCst,TblMastAccount.AcContno,tblmastaccount.accity,tblmastaccount.acADD1 from tbladdLess,tblmastaccount,tblmastaccount tblBrk,tblmastaccount  tblMill,tblOutstanding,tblMastNarration where" _
       & " tblOutstanding.vNo=tbladdLess.vno and tblOutstanding.vType=tbladdLess.vType and tblOutstanding.vdt=tbladdLess.vdt" _
       & " and ADACCRCODE=tblmastaccount.Accode and ADMILLCODE=tblMill.AcCode and " _
       & " OUTBROKERCODE=tblBrk.AcCode and ADTAXCODE=NARRCODE and (instr(narration,'C Form')>0 ) and tbladdLess.Vtype in ('PT') and tbladdless.AdGpNo = 0 and " & CondStr

   GProcInsertToSetupDbTbl AllFirmFlag, X, "tmpGenTbl", "Vtype,dt1,VbillNo,VBillDt,Amt1,Nar2,Nar1,Nar11,Nar4,username,Nar6,nar7,nar8,nar9,nar10 ", " username='" & gUserName & "'", "Nar3", "Nar5"
   If GProcRstOpen(mTmprst1, "select DISTINCT nar11 from tmpGenTbl where UserName='" & gUserName & "'  order by Nar11", "R", gSetupCn) > 0 Then
      With mTmprst1
           mTmprst1.MoveFirst
           Do While Not .EOF
              SmsMblNo = !nar11 & vbNullString
              If SmsMblNo = "" Then
'                  MsgBox ("Party Mobile No Not Found....")
              Else
                  
''                    url = "http://bhashsms.com/api/sendmsg.php"
''                    params = "user=universaltex&pass=123456&sender=UTCYRN&phone=" & SmsMblNo & "&text=" & SmsString & "&priority=ndnd&stype=normal"
''                    HttpReq.Open "POST", url, False
''                    HttpReq.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
''                    HttpReq.Send params
delay (3)
                  
                  
                  url = "http://alerts.sinfini.com/api/web2sms.php"
                  params = "workingkey= 98702uv6wtl7bqnh450&sender=UTCYRN&to=" & SmsMblNo & " &message=" & SmsString & ""
                  HttpReq.Open "POST", url, False
                  HttpReq.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
                  HttpReq.send params
                  NoOfSms = NoOfSms + 1
               End If
               mTmprst1.MoveNext
           Loop
           MsgBox CStr(NoOfSms) + " SMS Send..."
      End With
   End If
Else
        Select Case gSelectedMenu
               Case "Pending C form Receivable", "Form 'C' Receivable SMS (Trade/SIT) Single Party", "Form 'C' Receivable SMS (Trade/SIT) Comman SMS to All Party", "Form 'E1' Receivable SMS  All Party", "Form 'E1' Receivable SMS  Single Party", "Form 'E2' Receivable SMS  All Party", "Form 'E2' Receivable SMS  Single Party", "Pending C form Receivable All Broker", "Pending C form Receivable Single Broker", "Form 'E1' Receivable SMS (Trade/SIT) All Broker", "Form 'E1' Receivable SMS (Trade/SIT) Single Broker", "Form 'E1' Receivable SMS Comman SMS to All Party"
'                    Dim CondStr As String
'                    Dim AllFirmFlag As Boolean
'                    AllFirmFlag = False
'                    Dim QtrArray As Integer
                    QtrArray = 1
Loop1:
                    Do While QtrArray <= 4
                        If gSelectedMenu = "Form 'C' Receivable SMS (Trade/SIT) Comman SMS to All Party" Or gSelectedMenu = "Pending C form Receivable All Broker" Or gSelectedMenu = "Pending C form Receivable Single Broker" Or gSelectedMenu = "Form 'E1' Receivable SMS (Trade/SIT) All Broker" Or gSelectedMenu = "Form 'E1' Receivable SMS (Trade/SIT) Single Broker" Or gSelectedMenu = "Form 'E1' Receivable SMS Comman SMS to All Party" Then
                            QtrArray = 4
                        Else
                            If QtrArray = 1 And chkQtr1.Value = 0 Then
                               QtrArray = QtrArray + 1
                               GoTo Loop1
                            End If
                            If QtrArray = 2 And chkQtr2.Value = 0 Then
                               QtrArray = QtrArray + 1
                               GoTo Loop1
                            End If
                            If QtrArray = 3 And chkQtr3.Value = 0 Then
                               QtrArray = QtrArray + 1
                               GoTo Loop1
                            End If
                            If QtrArray = 4 And chkQtr4.Value = 0 Then
                               QtrArray = QtrArray + 1
                               GoTo Loop1
                            End If
                            Select Case QtrArray
                            Case 1
                                 DTPicker1.Value = CDate(Format("01/04/" & Year(gCYSDate), "dd/mm/yyyy"))
                                 DTPicker2.Value = CDate(Format("30/06/" & Year(gCYSDate), "dd/mm/yyyy"))
                            Case 2
                                 DTPicker1.Value = CDate(Format("01/07/" & Year(gCYSDate), "dd/mm/yyyy"))
                                 DTPicker2.Value = CDate(Format("30/09/" & Year(gCYSDate), "dd/mm/yyyy"))
                            Case 3
                                 DTPicker1.Value = CDate(Format("01/10/" & Year(gCYSDate), "dd/mm/yyyy"))
                                 DTPicker2.Value = CDate(Format("31/12/" & Year(gCYSDate), "dd/mm/yyyy"))
                            Case 4
                                 DTPicker1.Value = CDate(Format("01/01/" & Year(gCYEDate), "dd/mm/yyyy"))
                                 DTPicker2.Value = CDate(Format("31/03/" & Year(gCYEDate), "dd/mm/yyyy"))
                            End Select
                        End If
                    
                            If gSelectedMenu = "Form 'E1' Receivable SMS (Trade/SIT) All Broker" Or gSelectedMenu = "Form 'E1' Receivable SMS (Trade/SIT) Single Broker" And gSelectedMenu = "Form 'E2' Receivable SMS  Single Party" Or gSelectedMenu = "Form 'E2' Receivable SMS  All Party" Or gSelectedMenu = "Form 'E1' Receivable SMS  Single Party" Or gSelectedMenu = "Form 'E1' Receivable SMS  All Party" Or gSelectedMenu = "Form 'E1' Receivable SMS Comman SMS to All Party" Then
                               CondStr = " tbladdLess.AdGpDt>= cdate('" & Format(DTPicker1.Value, "dd/MM/yyyy") & "')  and tbladdLess.AdGpDt<= cdate('" & Format(DTPicker2.Value, "dd/MM/yyyy") & "')"
                               CondStr = CondStr + " and (ADE1FORMNO ='' or ADE1FORMNO is null)   " ' (ADCFORMNO is null or ADCFORMNO='') " '-
                            Else
                               CondStr = " tbladdLess.vdt>= cdate('" & Format(DTPicker1.Value, "dd/MM/yyyy") & "')  and tbladdLess.vdt<= cdate('" & Format(DTPicker2.Value, "dd/MM/yyyy") & "')"
                               CondStr = CondStr + " and (ADCFORMNO='' or ADCFORMNO is null)" ' (ADCFORMNO is null or ADCFORMNO='') " '-
                            End If
                            If gSelectedMenu = "Form 'C' Receivable SMS (Trade/SIT) Single Party" Then
                               CondStr = CondStr + " and TblMastAccount.AcName='" & txtTemp(1).text & "'"  '--Party
                            End If
                    
                            If gSelectedMenu = "Form 'E1' Receivable SMS  Single Party" Then
                               CondStr = CondStr + " and TblMastAccount.AcName='" & txtTemp(1).text & "'"  '--Party
                            End If
                            If gSelectedMenu = "Form 'E2' Receivable SMS  Single Party" Then
                               CondStr = CondStr + " and TblMastAccount.AcName='" & txtTemp(1).text & "'"  '--Party
                            End If
                            If gSelectedMenu = "Pending C form Receivable Single Broker" Or gSelectedMenu = "Form 'E1' Receivable SMS (Trade/SIT) Single Broker" Then
                               CondStr = CondStr + " and TblBrk.AcName='" & txtTemp(2).text & "'"  '--Broker
                            End If
                    
                            If gSelectedMenu = "Form 'E1' Receivable SMS  All Party" Or gSelectedMenu = "Form 'E1' Receivable SMS  Single Party" Or gSelectedMenu = "Form 'E1' Receivable SMS Comman SMS to All Party" Then
                
                                X = " select tbladdLess.VType,tblAddless.AdFormE1RecdDate,ADINVNO,tbladdLess.AdGpdt,ADBILLAMT,TblMastAccount.AcName,tblBrk.AcName,TblMastAccount.AcFaxNo,tblMill.AcName,'" & gUserName & "'" _
                                & " ,ADE1FORMNO,TblMastAccount.AcCst,TblMastAccount.AcContno,tblmastaccount.accity,tblmastaccount.acADD1 from tbladdLess,tblmastaccount,tblmastaccount tblBrk,tblmastaccount  tblMill,tblOutstanding,tblMastNarration where" _
                                & " tblOutstanding.vNo=tbladdLess.vno and tblOutstanding.vType=tbladdLess.vType and tblOutstanding.vdt=tbladdLess.vdt" _
                                & " and ADACCRCODE=tblmastaccount.Accode and ADMILLCODE=tblMill.AcCode and " _
                                & " OUTBROKERCODE=tblBrk.AcCode and ADTAXCODE=NARRCODE and (instr(narration,'C Form')>0 ) and tbladdLess.Vtype in ('PT') and tbladdless.AdGpNo = 0 and " & CondStr
                            
                            ElseIf gSelectedMenu = "Form 'E2' Receivable SMS  All Party" Or gSelectedMenu = "Form 'E2' Receivable SMS  Single Party" Then
                        
                                X = " select tbladdLess.VType,tblAddless.AdFormE1RecdDate,ADINVNO,tbladdLess.AdGpdt,ADBILLAMT,TblMastAccount.AcName,tblBrk.AcName,TblMastAccount.AcFaxNo,tblMill.AcName,'" & gUserName & "'" _
                                & " ,ADE1FORMNO,TblMastAccount.AcCst,TblMastAccount.AcContno,tblmastaccount.accity,tblmastaccount.acADD1 from tbladdLess,tblmastaccount,tblmastaccount tblBrk,tblmastaccount  tblMill,tblOutstanding,tblMastNarration where" _
                                & " tblOutstanding.vNo=tbladdLess.vno and tblOutstanding.vType=tbladdLess.vType and tblOutstanding.vdt=tbladdLess.vdt" _
                                & " and ADACCRCODE=tblmastaccount.Accode and ADMILLCODE=tblMill.AcCode and " _
                                & " OUTBROKERCODE=tblBrk.AcCode and ADTAXCODE=NARRCODE and (instr(narration,'C Form')>0 ) and tbladdLess.Vtype in ('PT') and tbladdless.AdGpNo = 1 and " & CondStr
                            ElseIf gSelectedMenu = "Pending C form Receivable All Broker" Or gSelectedMenu = "Pending C form Receivable Single Broker" Then
                                X = "select  '',cdate('" & Format(DTPicker1.Value, "dd/MM/yyyy") & "'),'',cdate('" & Format(DTPicker1.Value, "dd/MM/yyyy") & "'),sum(ADBILLAMT),tblBrk.AcName,TblMastAccount.AcName,tblBrk.AcFaxNo,'','" & gUserName & "' from  tbladdLess,tblmastaccount,tblmastaccount tblBrk,tblOutstanding,tblMastNarration" _
                                & " where tblOutstanding.vNo=tbladdLess.vno and tblOutstanding.vType=tbladdLess.vType and tblOutstanding.vdt=tbladdLess.vdt and ADACDRCODE=tblmastaccount.Accode  and  OUTBROKERCODE=tblBrk.AcCode " _
                                & " and ADTAXCODE=NARRCODE and (instr(narration,'C Form')>0 or instr(narration,'C S T')>0)  and tbladdLess.Vtype in ('SY','ST','SD','SO') and  " & CondStr & "" _
                                & " GROUP BY tblBrk.AcName, TblMastAccount.AcName, tblBrk.AcFaxNo"
                            ElseIf gSelectedMenu = "Form 'E1' Receivable SMS (Trade/SIT) All Broker" Or gSelectedMenu = "Form 'E1' Receivable SMS (Trade/SIT) Single Broker" Then
                                X = "select  '',cdate('" & Format(DTPicker1.Value, "dd/MM/yyyy") & "'),'',cdate('" & Format(DTPicker1.Value, "dd/MM/yyyy") & "'),sum(ADBILLAMT),tblBrk.AcName,TblMastAccount.AcName,tblBrk.AcFaxNo,'','" & gUserName & "' from  tbladdLess,tblmastaccount,tblmastaccount tblBrk,tblOutstanding,tblMastNarration" _
                                & " where tblOutstanding.vNo=tbladdLess.vno and tblOutstanding.vType=tbladdLess.vType and tblOutstanding.vdt=tbladdLess.vdt and ADACcRCODE=tblmastaccount.Accode  and  OUTBROKERCODE=tblBrk.AcCode " _
                                & " and ADTAXCODE=NARRCODE and (instr(narration,'C Form')>0 ) and tbladdLess.Vtype in ('PT') and tbladdless.AdGpNo = 0 and  " & CondStr & "" _
                                & " GROUP BY tblBrk.AcName, TblMastAccount.AcName, tblBrk.AcFaxNo"
                            Else
                                X = " select tbladdLess.VType,tblAddless.AdFormRecdDate,ADINVNO,tbladdLess.Vdt,ADBILLAMT,TblMastAccount.AcName,tblBrk.AcName,TblMastAccount.AcFaxNo,tblMill.AcName,'" & gUserName & "'" _
                                & " ,ADCFORMNO,TblMastAccount.AcCst,TblMastAccount.AcContno,tblmastaccount.accity,tblmastaccount.acADD1 from tbladdLess,tblmastaccount,tblmastaccount tblBrk,tblmastaccount  tblMill,tblOutstanding,tblMastNarration where" _
                                & " tblOutstanding.vNo=tbladdLess.vno and tblOutstanding.vType=tbladdLess.vType and tblOutstanding.vdt=tbladdLess.vdt" _
                                & " and ADACDRCODE=tblmastaccount.Accode and ADMILLCODE=tblMill.AcCode and " _
                                & " OUTBROKERCODE=tblBrk.AcCode and ADTAXCODE=NARRCODE and (instr(narration,'C Form')>0 or instr(narration,'C S T')>0)     and tbladdLess.Vtype in ('SY','ST','SD','SO')  and " & CondStr
                            End If
                            '--- insert to setup db Table For All Firm
                            
                            GProcInsertToSetupDbTbl AllFirmFlag, X, "tmpGenTbl", "Vtype,dt1,VbillNo,VBillDt,Amt1,Nar2,Nar1,Nar11,Nar4,username,Nar6,nar7,nar8,nar9,nar10 ", " username='" & gUserName & "'", "Nar3", "Nar5"
                
                            
                            
                            
                            
                            
'                            Dim mTmprst1 As Recordset
'                            Dim mTmprst2 As Recordset
'                            Dim PartyName As String
'                            Dim CFromAmount As Double
'                            Dim NoOfBills As Integer
'                            Dim NoOfSms As Integer
'                            Dim SrNo  As Integer
                            
                            
                            
                            
                            NoOfSms = 0
                         If GProcRstOpen(mTmprst1, "select * from tmpGenTbl where UserName='" & gUserName & "'  order by Nar2", "R", gSetupCn) > 0 Then
                            With mTmprst1
                                 mTmprst1.MoveFirst
                                 Do While Not .EOF
                                    PartyName = !Nar2
                                    If gSelectedMenu = "Form 'C' Receivable SMS (Trade/SIT) Single Party" Then
                                        SmsMblNo = txtTemp(7)
                                    ElseIf gSelectedMenu = "Pending C form Receivable Single Broker" Then
                                        SmsMblNo = txtTemp(8)
                                    ElseIf gSelectedMenu = "Form 'E1' Receivable SMS (Trade/SIT) Single Broker" Then
                                        SmsMblNo = txtTemp(8)
                                    ElseIf gSelectedMenu = "Form 'E1' Receivable SMS  Single Party" Then
                                        SmsMblNo = txtTemp(7)
                                    Else
                                        SmsMblNo = !nar11 & vbNullString
                                        'txttemp (7)
                                        
                                    End If
                                    CFromAmount = 0
                                    NoOfBills = 0
                                    GProcRstOpen mTmprst2, "select * from tmpGenTbl where UserName='" & gUserName & "' and Nar2 = '" & PartyName & "' ", "R", gSetupCn
                                    With mTmprst2
                                         mTmprst2.MoveFirst
                                         SrNo = 1
                                         Do While Not .EOF
                                            If gSelectedMenu = "Pending C form Receivable All Broker" Or gSelectedMenu = "Pending C form Receivable Single Broker" Then
                                               If SrNo = 1 Then
                                                  SmsString = "Pending C Form FY " + CStr(Right(Year(Format(DTPicker1.Value, "dd/MM/yyyy")), 2)) + "-" + CStr(Right(Year(Format(DTPicker2.Value, "dd/MM/yyyy")), 2)) & vbLf
                                                  SmsString = SmsString & vbLf
                                                End If
                                                SmsString = SmsString + CStr(SrNo) + "." + !Nar1 & vbLf
                                                SrNo = SrNo + 1
                                            ElseIf gSelectedMenu = "Form 'E1' Receivable SMS (Trade/SIT) All Broker" Then
                                               If SrNo = 1 Then
                                                  SmsString = "Pending E1 Form FY " + CStr(Right(Year(Format(DTPicker1.Value, "dd/MM/yyyy")), 2)) + "-" + CStr(Right(Year(Format(DTPicker2.Value, "dd/MM/yyyy")), 2)) & vbLf
                                                  SmsString = SmsString & vbLf
                                                End If
                                                SmsString = SmsString + CStr(SrNo) + "." + !Nar1 & vbLf
                                                SrNo = SrNo + 1
                                            ElseIf gSelectedMenu = "Form 'E1' Receivable SMS (Trade/SIT) Single Broker" Then
                                               If SrNo = 1 Then
                                                  SmsString = "Pending E1 Form FY " + CStr(Right(Year(Format(DTPicker1.Value, "dd/MM/yyyy")), 2)) + "-" + CStr(Right(Year(Format(DTPicker2.Value, "dd/MM/yyyy")), 2)) & vbLf
                                                  SmsString = SmsString & vbLf
                                                End If
                                                SmsString = SmsString + CStr(SrNo) + "." + !Nar1 & vbLf
                                                SrNo = SrNo + 1
                                            Else
                                               CFromAmount = CFromAmount + !Amt1
                                               NoOfBills = NoOfBills + 1
                                            End If
                                            mTmprst2.MoveNext
                                         Loop
                                    End With
                                    '--- For Sending SMS
                                    
'                                    If gSelectedMenu = "Form 'C' Receivable SMS (Trade/SIT) Comman SMS to All Party" Or gSelectedMenu = "Form 'E1' Receivable SMS Comman SMS to All Party" Then
'                                         SmsString = txtTemp(3).Text
 '                                        SmsString = Left(SmsString, 360)
                
                
                
'                                    ElseIf gSelectedMenu = "Form 'E1' Receivable SMS  All Party" Then
                                    If gSelectedMenu = "Form 'E1' Receivable SMS  All Party" Then
                                        SmsString = "Pending E-1 Form Detail" & vbLf
                                        SmsString = SmsString & vbLf
                                        SmsString = SmsString + "Pty-" + Left(PartyName, 30) & vbLf
                                        SmsString = SmsString + "Period-" + Trim(CStr(Format(DTPicker1.Value, "dd/mm/yy"))) + " -" + Trim(CStr(Format(DTPicker2.Value, "dd/mm/yy"))) & vbLf
                                        SmsString = SmsString + "Amt-" + CStr(CFromAmount) + "/-" & vbLf
                                        SmsString = SmsString + "Bills-" + Trim(CStr(NoOfBills)) & vbLf
                                        SmsString = SmsString & vbLf
                                        SmsString = SmsString + "Firm-" & vbLf
                                        SmsString = SmsString + Left(gCName, 22) & vbLf
                                        SmsString = SmsString + "Ph: 0230-2433670"
                                    ElseIf gSelectedMenu = "Form 'E1' Receivable SMS  Single Party" Then
                                        SmsString = "Pending E-1 Form Detail" & vbLf
                                        SmsString = SmsString & vbLf
                                        SmsString = SmsString + "Pty-" + Left(PartyName, 30) & vbLf
                                        SmsString = SmsString + "Period-" + Trim(CStr(Format(DTPicker1.Value, "dd/mm/yy"))) + " -" + Trim(CStr(Format(DTPicker2.Value, "dd/mm/yy"))) & vbLf
                                        SmsString = SmsString + "Amt-" + CStr(CFromAmount) + "/-" & vbLf
                                        SmsString = SmsString + "Bills-" + Trim(CStr(NoOfBills)) & vbLf
                                        SmsString = SmsString & vbLf
                                        SmsString = SmsString + "Firm-" & vbLf
                                        SmsString = SmsString + gCName & vbLf
                                        SmsString = SmsString + "Ph:0230-2433670"
                                    ElseIf gSelectedMenu = "Form 'E2' Receivable SMS  All Party" Then
                                        SmsString = "Pending E-2 Form Detail" & vbLf
                                        SmsString = SmsString & vbLf
                                        SmsString = SmsString + "Pty-" + Left(PartyName, 30) & vbLf
                                        SmsString = SmsString + "Period-" + Trim(CStr(Format(DTPicker1.Value, "dd/mm/yy"))) + " -" + Trim(CStr(Format(DTPicker2.Value, "dd/mm/yy"))) & vbLf
                                        SmsString = SmsString + "Amt-" + CStr(CFromAmount) + "/-" & vbLf
                                        SmsString = SmsString + "Bills-" + Trim(CStr(NoOfBills)) & vbLf
                                        SmsString = SmsString & vbLf
                                        SmsString = SmsString + "Firm-" & vbLf
                                        SmsString = SmsString + gCName & vbLf
                                        SmsString = SmsString + "Ph- 2433670"
                                    ElseIf gSelectedMenu = "Form 'E2' Receivable SMS  Single Party" Then
                                        SmsString = "Pending E-2 Form Detail" & vbLf
                                        SmsString = SmsString & vbLf
                                        SmsString = SmsString + "Pty-" + Left(PartyName, 30) & vbLf
                                        SmsString = SmsString + "Period-" + Trim(CStr(Format(DTPicker1.Value, "dd/mm/yy"))) + " -" + Trim(CStr(Format(DTPicker2.Value, "dd/mm/yy"))) & vbLf
                                        SmsString = SmsString + "Amt-" + CStr(CFromAmount) + "/-" & vbLf
                                        SmsString = SmsString + "Bills-" + Trim(CStr(NoOfBills)) & vbLf
                                        SmsString = SmsString & vbLf
                                        SmsString = SmsString + "Firm-" & vbLf
                                        SmsString = SmsString + gCName & vbLf
                                        SmsString = SmsString + "Ph- 2433670"
                                        
                                    ElseIf gSelectedMenu = "Pending C form Receivable All Broker" Then
                                        SmsString = SmsString & vbLf
                                        SmsString = SmsString + "Firm-" & vbLf
                                        SmsString = SmsString + gCName & vbLf
                                    ElseIf gSelectedMenu = "Pending C form Receivable Single Broker" Then
                                        SmsString = SmsString & vbLf
                                        SmsString = SmsString + "Firm-" & vbLf
                                        SmsString = SmsString + gCName & vbLf
                                    ElseIf gSelectedMenu = "Form 'E1' Receivable SMS (Trade/SIT) All Broker" Then
                                        SmsString = SmsString & vbLf
                                        SmsString = SmsString + "Firm-" & vbLf
                                        SmsString = SmsString + gCName & vbLf
                                    ElseIf gSelectedMenu = "Form 'E1' Receivable SMS (Trade/SIT) Single Broker" Then
                                        SmsString = SmsString & vbLf
                                        SmsString = SmsString + "Firm-" & vbLf
                                        SmsString = SmsString + gCName & vbLf
        
                                    
                                    Else
        '                                SmsString = Left(PartyName, 20) & vbLf
                                        SmsString = "Pending C Form Detail" & vbLf
                                        SmsString = SmsString & vbLf
                                        SmsString = SmsString + "Pty-" + Left(PartyName, 30) & vbLf
                                        SmsString = SmsString + "Period-" + Trim(CStr(Format(DTPicker1.Value, "dd/mm/yy"))) + " -" + Trim(CStr(Format(DTPicker2.Value, "dd/mm/yy"))) & vbLf
                                        SmsString = SmsString + "Amt-" + CStr(CFromAmount) + "/-" & vbLf
                                        SmsString = SmsString + "Bills-" + Trim(CStr(NoOfBills)) & vbLf
                                        SmsString = SmsString & vbLf
                                        SmsString = SmsString + "Firm-" & vbLf
                                        SmsString = SmsString + gCName & vbLf
                                        SmsString = SmsString + "Ph- 2433670"
                                    End If
                                    
                                    Select Case gSelectedMenu
                                    Case "Pending C form Receivable All Broker"
                                         SmsString = SmsString
                                    Case "Pending C form Receivable Single Broker"
                                         SmsString = SmsString
                                    Case "Form 'E1' Receivable SMS (Trade/SIT) All Broker"
                                         SmsString = SmsString
                                    Case "Form 'E1' Receivable SMS (Trade/SIT) Single Broker"
                                         SmsString = SmsString
                                    Case Else
                                        SmsString = Left(SmsString, 160)
                                    End Select
                                    If SmsMblNo = "" Then
                '                       MsgBox ("Party Mobile No Not Found....")
                                    Else
                                        
''                    url = "http://smshorizon.co.in/api/sendsms.php?"
''                    params = "user=universaltex&apikey=YRPUla1IJxTUU5iKRTQQ&number=" & SmsMblNo & "&message=" & SmsString & "&senderid=UTCYRN&type=txt"
'''                    params = "user=universaltex&apikey=123456&number=" & SmsMblNo & "&message=" & SmsString & "&senderid=SMSTXT&type=txt"
''                    HttpReq.Open "POST", url, False
''                    HttpReq.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
''                    HttpReq.Send params
''                    NoOfSms = NoOfSms + 1
''
''                                        url = "http://bhashsms.com/api/sendmsg.php"
''                                        params = "user=universaltex&pass=123456&sender=UTCYRN&phone=" & SmsMblNo & "&text=" & SmsString & "&priority=ndnd&stype=normal"
''                                        HttpReq.Open "POST", url, False
''                                        HttpReq.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
''                                        HttpReq.Send params
''delay (3)
''

                                        url = "http://alerts.sinfini.com/api/web2sms.php"
                                        params = "workingkey= 98702uv6wtl7bqnh450&sender=UTCYRN&to=" & SmsMblNo & " &message=" & SmsString & ""
                                        HttpReq.Open "POST", url, False
                                        HttpReq.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
                                        HttpReq.send params
                                        NoOfSms = NoOfSms + 1
                                    End If
                                    Do While Not mTmprst1.EOF
                                       If PartyName = !Nar2 Then
                                          mTmprst1.MoveNext
                                       Else
                                          Exit Do
                                       End If
                                    Loop
                                 Loop
                            End With
                            If gSelectedMenu = "Form 'C' Receivable SMS (Trade/SIT) Comman SMS to All Party" Or gSelectedMenu = "Pending C form Receivable All Broker" Or gSelectedMenu = "Pending C form Receivable Single Broker" Or gSelectedMenu = "Form 'E1' Receivable SMS (Trade/SIT) Single Broker" Or gSelectedMenu = "Form 'E1' Receivable SMS (Trade/SIT) All Broker" Or gSelectedMenu = "Form 'E1' Receivable SMS Comman SMS to All Party" Then
                                MsgBox CStr(NoOfSms) + " SMS Send..."
                            Else
                                Select Case QtrArray
                                Case 1
                                    MsgBox "1 st Quarter " + (CStr(NoOfSms) + " SMS Send...")
                                Case 2
                                    MsgBox "2 nd Quarter " + (CStr(NoOfSms) + " SMS Send...")
                                Case 3
                                    MsgBox "3 rd Quarter " + (CStr(NoOfSms) + " SMS Send...")
                                Case 4
                                    MsgBox "4 th Quarter " + (CStr(NoOfSms) + " SMS Send...")
                                End Select
                            End If
                          End If
                          QtrArray = QtrArray + 1
                    Loop
                    MsgBox "All SMS Send..."
               Case "General SMS" '== Sales Graph
                    SmsMblNo = txtTemp(6).text
                    SmsString = txtTemp(3).text + " " + txtTemp(4).text + " " + txtTemp(5).text
                    SmsString = Left(SmsString, 360)
                    If SmsMblNo = "" Then
                    '   MsgBox ("Party Mobile No Not Found....")
                    Else
                        
                        
''                    url = "http://smshorizon.co.in/api/sendsms.php?"
''                    params = "user=universaltex&apikey=YRPUla1IJxTUU5iKRTQQ&mobile=" & SmsMblNo & "&message=" & SmsString & "&senderid=UTCYRN&type=txt"
''                    HttpReq.Open "POST", url, False
''                    HttpReq.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
''                    HttpReq.Send params
''                    NoOfSms = NoOfSms + 1
                       
                        
''                        url = "http://bhashsms.com/api/sendmsg.php"
''                        params = "user=universaltex&pass=123456&sender=UTCYRN&phone=" & SmsMblNo & "&text=" & SmsString & "&priority=ndnd&stype=normal"
''                        HttpReq.Open "POST", url, False
''                        HttpReq.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
''                        HttpReq.Send params
''delay (3)
                        
                        url = "http://alerts.sinfini.com/api/web2sms.php"
                        params = "workingkey= 98702uv6wtl7bqnh450&sender=UTCYRN&to=" & SmsMblNo & " &message=" & SmsString & ""

                       HttpReq.Open "POST", url, False
                       HttpReq.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
                       HttpReq.send params
                       SmsSent = True
                       'NoOfSms = NoOfSms + 1
                    End If
                    SmsMblNo = txtTemp(7).text
                    If SmsMblNo = "" Then
                    '   MsgBox ("Party Mobile No Not Found....")
                    Else
                        
                    url = "http://smshorizon.co.in/api/sendsms.php?"
'                    params = "user=universaltex&apikey=YRPUla1IJxTUU5iKRTQQ&number=" & SmsMblNo & "&message=" & SmsString & "&senderid=UTCYRN&type=text"
'                    params = "user=universaltex&apikey=123456&number=" & SmsMblNo & "&message=" & SmsString & "&senderid=SMSTXT&type=txt"
                    params = "user=universaltex&apikey=YRPUla1IJxTUU5iKRTQQ&mobile=" & SmsMblNo & "&message=" & SmsString & "&senderid=UTCYRN&type=txt"
                    
                    HttpReq.Open "POST", url, False
                    HttpReq.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
                    HttpReq.send params
                    NoOfSms = NoOfSms + 1
                        
                        
''                        url = "http://bhashsms.com/api/sendmsg.php"
''                        params = "user=universaltex&pass=123456&sender=UTCYRN&phone=" & SmsMblNo & "&text=" & SmsString & "&priority=ndnd&stype=normal"
''                        HttpReq.Open "POST", url, False
''                        HttpReq.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
''                        HttpReq.Send params
''delay (3)
                        
                        
'                        url = "http://alerts.sinfini.com/api/web2sms.php"
'                        params = "workingkey= 98702uv6wtl7bqnh450&sender=UTCYRN&to=" & SmsMblNo & " &message=" & SmsString & ""
'
'                       HttpReq.Open "POST", url, False
'                       HttpReq.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
'                       HttpReq.Send params
                       SmsSent = True
                       'NoOfSms = NoOfSms + 1
                    End If
                    SmsMblNo = txtTemp(8).text
                    If SmsMblNo = "" Then
                    '   MsgBox ("Party Mobile No Not Found....")
                    Else
                        
''                    url = "http://smshorizon.co.in/api/sendsms.php?"
'''                    params = "user=universaltex&apikey=YRPUla1IJxTUU5iKRTQQ&number=" & SmsMblNo & "&message=" & SmsString & "&senderid=UTCYRN&type=text"
'''                    params = "user=universaltex&apikey=123456&number=" & SmsMblNo & "&message=" & SmsString & "&senderid=SMSTXT&type=txt"
''                    params = "user=universaltex&apikey=YRPUla1IJxTUU5iKRTQQ&mobile=" & SmsMblNo & "&message=" & SmsString & "&senderid=UTCYRN&type=txt"
''
''                    HttpReq.Open "POST", url, False
''                    HttpReq.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
''                    HttpReq.Send params
''                    NoOfSms = NoOfSms + 1
                        
                        
''                        url = "http://bhashsms.com/api/sendmsg.php"
''                        params = "user=universaltex&pass=123456&sender=UTCYRN&phone=" & SmsMblNo & "&text=" & SmsString & "&priority=ndnd&stype=normal"
''                        HttpReq.Open "POST", url, False
''                        HttpReq.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
''                        HttpReq.Send params
''delay (3)
                        
                        url = "http://alerts.sinfini.com/api/web2sms.php"
                        params = "workingkey= 98702uv6wtl7bqnh450&sender=UTCYRN&to=" & SmsMblNo & " &message=" & SmsString & ""

                       HttpReq.Open "POST", url, False
                       HttpReq.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
                       HttpReq.send params
                       SmsSent = True
                       NoOfSms = NoOfSms + 1
                    End If
                    If SmsSent = True Then
                        MsgBox (CStr(NoOfSms) + " SMS Send...")
                    End If
        End Select
End If
'------End SMS
Exit Sub
ErrorRoutine:
    GProcErrorHandler
End Sub

Private Sub Option1_Click(Index As Integer)
txtTemp(3).text = ""
txtTemp(4).text = ""
txtTemp(5).text = ""

If Option1(0).Value = True Then '-----Blank
   txtTemp(3).text = ""
   txtTemp(4).text = ""
   txtTemp(5).text = ""
End If

If Option1(1).Value = True Then '---Bank Details
   txtTemp(3).text = "UNIVERSAL TEXTILE CORPORATION" & vbCrLf
   txtTemp(3).text = txtTemp(3).text + "HDFC BANK LTD." & vbCrLf
   txtTemp(3).text = txtTemp(3).text + "A/C NO - 50200006229274" & vbCrLf
   txtTemp(3).text = txtTemp(3).text + "IFSC- HDFC 0000736" & vbCrLf
   txtTemp(3).text = txtTemp(3).text + "BRANCH- ICHALKARANJI" & vbCrLf
End If

If Option1(2).Value = True Then '---Firm Add
   txtTemp(3).text = "UNIVERSAL TEXTILE CORPORATION" & vbCrLf
   txtTemp(3).text = txtTemp(3).text + "12/487,UNIVERSAL HOUSE" & vbCrLf
   txtTemp(3).text = txtTemp(3).text + "SHAHU CORNER,ICHALKARANJI-416115 (M.H)" & vbCrLf
   txtTemp(3).text = txtTemp(3).text + "MOB- 9422025001" & vbCrLf
   txtTemp(3).text = txtTemp(3).text + "EMAIL - universaltex@yahoo.co.in"
End If

If Option1(3).Value = True Then '---TIN Universal Text Corpo
   txtTemp(3).text = "FIRM- UNIVERSAL TEXTILE CORPORATION" & vbCrLf
   txtTemp(3).text = txtTemp(3).text + "TIN NO.- 27540710952  V-25/05/09" & vbCrLf
   txtTemp(3).text = txtTemp(3).text + "                27540710952  C-25/05/09" & vbCrLf
   txtTemp(3).text = txtTemp(3).text + "PAN NO.- AAEHJ9137K" & vbCrLf
   txtTemp(3).text = txtTemp(3).text + "MOB NO.- 9422025001" & vbCrLf
End If
If Option1(4).Value = True Then '---TIN Universal Spintex
   txtTemp(3).text = "FIRM- UNIVERSAL SPINTEX" & vbCrLf
   txtTemp(3).text = txtTemp(3).text + "TIN NO.- 27570895090 V" & vbCrLf
   txtTemp(3).text = txtTemp(3).text + "                27570895090 C" & vbCrLf
   txtTemp(3).text = txtTemp(3).text + "PAN NO.- AFCPC9174L" & vbCrLf
   txtTemp(3).text = txtTemp(3).text + "MOB NO.- 9422025001" & vbCrLf
End If
If Option1(5).Value = True Then '---C / E1 Form Reminder
   txtTemp(3).text = "CLEAR PENDING" & vbCrLf
   txtTemp(3).text = txtTemp(3).text + "C-FORMS FOR FY " + Right(CStr(Year(gCYSDate)), 4) + "-" + Right(CStr(Year(gCYEDate)), 2) & vbCrLf
   txtTemp(3).text = txtTemp(3).text + "AS EARLY AS POSSIBLE" & vbCrLf
   txtTemp(3).text = txtTemp(3).text + "" & vbCrLf
   txtTemp(3).text = txtTemp(3).text + "FROM" & vbCrLf
   txtTemp(3).text = txtTemp(3).text + gCName & vbCrLf
   
'   txtTemp(3).Text = txtTemp(3).Text + "FY " + Right(CStr(Year(gCYSDate) - 2), 4) + "-" + Right(CStr(Year(gCYEDate) - 2), 2) + "/" + Right(CStr(Year(gCYSDate) - 1), 4) + "-" + Right(CStr(Year(gCYEDate) - 1), 2) + "/" + Right(CStr(Year(gCYSDate)), 4) + "-" + Right(CStr(Year(gCYEDate)), 2) & vbCrLf
'   txtTemp(3).Text = txtTemp(3).Text + "FY " + Right(CStr(Year(gCYSDate)), 4) + "-" + Right(CStr(Year(gCYEDate)), 2) & vbCrLf
   
   
   '"UNIVERSAL TEXTILE CORPORATION" & vbCrLf
   'txtTemp(3).Text = txtTemp(3).Text + "UNIVERSAL SPINTEX" & vbCrLf
   txtTemp(3).text = txtTemp(3).text + "PH-0230-2433670 " & vbCrLf
End If
If Option1(6).Value = True Then '---Bank Balance
   txtTemp(3).text = "Bank Balnce Dt." + Left(CStr(CDate(Now)), 10) & vbCrLf
   'txtTemp(3).Text = txtTemp(3).Text + "FIRM- Universal Spintex/" & vbCrLf
   txtTemp(3).text = txtTemp(3).text + "Universal textile Corporation " & vbCrLf
   txtTemp(3).text = txtTemp(3).text + "HDFC    -" & vbCrLf
   'txtTemp(3).Text = txtTemp(3).Text + "TMB C/A -" & vbCrLf
   'txtTemp(3).Text = txtTemp(3).Text + "TMB C/C -" & vbCrLf
   txtTemp(3).text = txtTemp(3).text + "Clearing-"
End If
If Option1(7).Value = True Then '---Bank Balance
   txtTemp(3).text = "SMS Your Email Id On" & vbCrLf
   txtTemp(3).text = txtTemp(3).text + "Mobile No - 9673235146" & vbCrLf
   txtTemp(3).text = txtTemp(3).text + "" & vbCrLf
   txtTemp(3).text = txtTemp(3).text + "From" & vbCrLf
   txtTemp(3).text = txtTemp(3).text + "Universal textile Corporation" & vbCrLf
   'txtTemp(3).Text = txtTemp(3).Text + "Universal Spintex" & vbCrLf
End If
If Option1(8).Value = True Then '---Bank Balance
   txtTemp(3).text = "UNIVERSAL TEXTILE CORPORATION" & vbCrLf
   txtTemp(3).text = txtTemp(3).text + "CITY UNION BANK LTD." & vbCrLf
   txtTemp(3).text = txtTemp(3).text + "A/C NO - 512020010013684" & vbCrLf
   txtTemp(3).text = txtTemp(3).text + "IFSC- CIUB 0000388" & vbCrLf
   txtTemp(3).text = txtTemp(3).text + "BRANCH- BALOTARA" & vbCrLf

'   txtTemp(3).Text = "Bank Details " & vbCrLf
'   txtTemp(3).Text = txtTemp(3).Text + "FIRM- Universal Spintex" & vbCrLf
'   txtTemp(3).Text = txtTemp(3).Text + "Bank - HDFC BANK " & vbCrLf
'   txtTemp(3).Text = txtTemp(3).Text + "A/C No. 07362560009397" & vbCrLf
'   txtTemp(3).Text = txtTemp(3).Text + "IFSC - HDFC 0000736" & vbCrLf
'   txtTemp(3).Text = txtTemp(3).Text + "BRANCH - ICHALKARANJI" & vbCrLf
End If
If Option1(9).Value = True Then '---Bank Balance

   txtTemp(3).text = "CLEAR PENDING" & vbCrLf
   txtTemp(3).text = txtTemp(3).text + "E1-FORMS FOR FY " + Right(CStr(Year(gCYSDate)), 4) + "-" + Right(CStr(Year(gCYEDate)), 2) & vbCrLf
   txtTemp(3).text = txtTemp(3).text + "AS EARLY AS POSSIBLE." & vbCrLf
   txtTemp(3).text = txtTemp(3).text + "" & vbCrLf
   txtTemp(3).text = txtTemp(3).text + "FROM" & vbCrLf
   txtTemp(3).text = txtTemp(3).text + gCName & vbCrLf
   txtTemp(3).text = txtTemp(3).text + "PH-0230-2433670 " & vbCrLf

''   txtTemp(3).Text = "Bank Details " & vbCrLf
''   txtTemp(3).Text = txtTemp(3).Text + "FIRM- Universal Textile Corporation" & vbCrLf
''   txtTemp(3).Text = txtTemp(3).Text + "Bank - HDFC BANK " & vbCrLf
''   txtTemp(3).Text = txtTemp(3).Text + "A/C No. 06432560005752" & vbCrLf
''   txtTemp(3).Text = txtTemp(3).Text + "IFSC - HDFC 0000643" & vbCrLf
''   txtTemp(3).Text = txtTemp(3).Text + "BRANCH - BALOTRA" & vbCrLf
End If
If Option1(10).Value = True Then '---Bank Balance
   txtTemp(3).text = txtTemp(3).text + "SMS PAN NO. OF YOUR FIRM ON" & vbCrLf
   txtTemp(3).text = txtTemp(3).text + "MOBILE NO 9673235146 OR EMAIL US ON" & vbCrLf
   txtTemp(3).text = txtTemp(3).text + "universaltex@yahoo.co.in" & vbCrLf
   txtTemp(3).text = txtTemp(3).text + "" & vbCrLf
   txtTemp(3).text = txtTemp(3).text + "FROM" & vbCrLf
   txtTemp(3).text = txtTemp(3).text + gCName & vbCrLf
   txtTemp(3).text = txtTemp(3).text + "PH-0230-2433670 " & vbCrLf

'   txtTemp(3).Text = "Bank Details " & vbCrLf
'   txtTemp(3).Text = txtTemp(3).Text + "FIRM- Universal Spintex" & vbCrLf
'   txtTemp(3).Text = txtTemp(3).Text + "Bank - TAMILNAD MERCANTILE BANK LTD" & vbCrLf
'   txtTemp(3).Text = txtTemp(3).Text + "A/C No. 143150310876750" & vbCrLf
'   txtTemp(3).Text = txtTemp(3).Text + "IFSC - TMBL 0000143" & vbCrLf
'   txtTemp(3).Text = txtTemp(3).Text + "BRANCH - ICHALKARANJI" & vbCrLf
End If
If Option1(11).Value = True Then '---Bank Balance
   txtTemp(3).text = "Bank Details " & vbCrLf
   txtTemp(3).text = txtTemp(3).text + "FIRM- Universal Synthetics" & vbCrLf
   txtTemp(3).text = txtTemp(3).text + "Bank - HDFC BANK " & vbCrLf
   txtTemp(3).text = txtTemp(3).text + "A/C No. 06438730000675" & vbCrLf
   txtTemp(3).text = txtTemp(3).text + "IFSC - HDFC 0000643" & vbCrLf
   txtTemp(3).text = txtTemp(3).text + "BRANCH - BALOTRA" & vbCrLf
End If



lblSmsLength.Caption = Len(txtTemp(3))
End Sub
Private Sub txtTemp_Click(Index As Integer)
    If Index = 3 Then '-- SMS Length
       lblSmsLength.Caption = Len(txtTemp(3))
    End If
End Sub

Private Sub txttemp_GotFocus(Index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub txtTemp_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then Exit Sub '-- Enter
    If Index = 3 Then '-- SMS Length
       lblSmsLength.Caption = Len(txtTemp(3))
    End If
    Select Case Index
     Case 0, 1, 2 '----Mill Name
'        gClsSearch.SearchMultiField "QryMastaccount", "Account,ACCode", Array("Name", ""), Array(txtTemp(Index).Width, 0, 0), "", Chr(KeyAscii), "AcName", txtTemp(Index).Left + Me.Left, txtTemp(Index).Top + Me.Top + 1000
        gClsSearch.SearchMultiField "QryMastaccount", "Account,AcCODE,AgName", Array("Party", "AcName", "Group"), Array(txtTemp(Index).Width, 0, 3000), " ", Chr(KeyAscii), "acName", , txtTemp(Index).Left + Me.Left, txtTemp(Index).Top + Me.Top + 1000

        KeyAscii = 0
        
        If gClsSearch.SearchMultiRetCol(1) <> "" Then
            txtTemp(Index) = gClsSearch.SearchMultiRetCol(0)
            wMillCode = gClsSearch.SearchMultiRetCol(1)
            If Index = 0 Then
               txtTemp(6).text = GProcGetColumnValue("TblMastAccount", "AcName", txtTemp(0), "S", "AcFaxNo", "S")
            End If
            If Index = 1 Then
               txtTemp(7).text = GProcGetColumnValue("TblMastAccount", "AcName", txtTemp(1), "S", "AcFaxNo", "S")
            End If
            If Index = 2 Then
               txtTemp(8).text = GProcGetColumnValue("TblMastAccount", "AcName", txtTemp(2), "S", "AcFaxNo", "S")
            End If
        End If
'     Case 1 '-- Item
'        gClsSearch.SearchMultiField "tblMastItem,tblMastAccount", "ItName,ItCode,Itticket,Acname", Array("Item", "", "Ticket", "Mill"), Array(2000, 0, 1000, 2000), " tblMastItem.ItMillcode=tblMastAccount.AcCode", Chr(KeyAscii), "ItName", txtTemp(Index).Left + Me.Left, txtTemp(Index).Top + Me.Top + 1000
'        KeyAscii = 0
'        If gClsSearch.SearchMultiRetCol(0) <> "" Then
'            txtTemp(Index) = gClsSearch.SearchMultiRetCol(0) '-ItName
'            txtTemp(0) = gClsSearch.SearchMultiRetCol(3) '-Mill
'            wItemCode = gClsSearch.SearchMultiRetCol(1)
'        End If
'     Case 2 '--Broker
'        gClsSearch.SearchMultiField "QryMastaccount", "Account,AcName,ACCode", Array("Broker", "AcName", ""), Array(txtTemp(Index).Width, 0, 0), " (AgCode=90016 or GpCode=90016)", Chr(KeyAscii), "AcName", txtTemp(Index).Left + Me.Left, txtTemp(Index).Top + Me.Top + 1000
'        KeyAscii = 0
'        If gClsSearch.SearchMultiRetCol(1) <> "" Then
'            txtTemp(Index) = gClsSearch.SearchMultiRetCol(1)
'            wBrkCode = gClsSearch.SearchMultiRetCol(2)
'        End If
    End Select
End Sub

Private Sub txttemp_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 3 Then '-- SMS Length
       lblSmsLength.Caption = Len(txtTemp(3))
    End If
End Sub


Public Sub delay(PauseTime As Integer)
    Dim start As Single
    start = Timer
   Do While Timer < start + PauseTime
    If (Timer < start) Then 'midnight crossover
        start = start - (86400 + 1)
    End If
    DoEvents   ' Yield to other processes.
    Loop
End Sub


