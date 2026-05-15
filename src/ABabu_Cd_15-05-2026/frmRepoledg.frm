VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmReport 
   BackColor       =   &H00C2E3FC&
   BorderStyle     =   0  'None
   Caption         =   "."
   ClientHeight    =   8955
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14235
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8955
   ScaleWidth      =   14235
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Account"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Index           =   0
      Left            =   5040
      TabIndex        =   33
      Top             =   3240
      Visible         =   0   'False
      Width           =   5055
      Begin VB.TextBox txtAccode 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   4200
         MaxLength       =   7
         TabIndex        =   40
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox txtclno 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1950
         MaxLength       =   7
         TabIndex        =   37
         Top             =   600
         Width           =   1575
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Ok"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   0
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox txtparty 
         Appearance      =   0  'Flat
         Height          =   360
         Index           =   1
         Left            =   1950
         TabIndex        =   35
         Top             =   1080
         Width           =   2895
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   2
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Column No. :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   480
         TabIndex        =   39
         Top             =   600
         Width           =   1230
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filter         :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   615
         TabIndex        =   38
         Top             =   1080
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdFilter 
      Caption         =   "Filter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   0
      Left            =   6240
      TabIndex        =   32
      Top             =   600
      Width           =   1185
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Outstanding"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2955
      Index           =   1
      Left            =   5280
      TabIndex        =   18
      Top             =   3120
      Visible         =   0   'False
      Width           =   3990
      Begin VB.CommandButton cmdOut 
         Caption         =   "&Ok"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   1
         Left            =   705
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   2220
         Width           =   1215
      End
      Begin VB.CommandButton cmdOut 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   2
         Left            =   2220
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   2190
         Width           =   1215
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   645
         Index           =   2
         Left            =   420
         TabIndex        =   19
         Top             =   1455
         Visible         =   0   'False
         Width           =   3420
         Begin VB.Frame Frame4 
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            Height          =   450
            Index           =   0
            Left            =   285
            TabIndex        =   23
            Top             =   585
            Visible         =   0   'False
            Width           =   420
            Begin VB.OptionButton optDS 
               BackColor       =   &H00FFC0C0&
               Caption         =   "Pending"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   3
               Left            =   1470
               TabIndex        =   24
               Top             =   75
               Width           =   1260
            End
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            Height          =   465
            Index           =   1
            Left            =   45
            TabIndex        =   20
            Top             =   150
            Width           =   3180
            Begin VB.OptionButton optDS 
               BackColor       =   &H00FFC0C0&
               Caption         =   "Simple"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   0
               Left            =   270
               TabIndex        =   22
               Top             =   45
               Width           =   1365
            End
            Begin VB.OptionButton optDS 
               BackColor       =   &H00FFC0C0&
               Caption         =   "Detail"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   1
               Left            =   1695
               TabIndex        =   21
               Top             =   45
               Width           =   1365
            End
         End
      End
      Begin MSComCtl2.DTPicker dtposfrom 
         Height          =   375
         Left            =   2040
         TabIndex        =   27
         Top             =   480
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   661
         _Version        =   393216
         Format          =   94371841
         CurrentDate     =   39505
      End
      Begin MSComCtl2.DTPicker dtposto 
         Height          =   375
         Left            =   2040
         TabIndex        =   28
         Top             =   960
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   661
         _Version        =   393216
         Format          =   94371841
         CurrentDate     =   39505
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From Date:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   600
         TabIndex        =   30
         Top             =   480
         Width           =   1050
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To Date    :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   600
         TabIndex        =   29
         Top             =   960
         Width           =   1035
      End
   End
   Begin VB.TextBox TxtPrefix 
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
      Left            =   12000
      TabIndex        =   16
      Top             =   1283
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.OptionButton OpCrDr 
      BackColor       =   &H00E4F8FC&
      Caption         =   "Both"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00645285&
      Height          =   255
      Index           =   2
      Left            =   10290
      TabIndex        =   14
      Top             =   1343
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.OptionButton OpCrDr 
      BackColor       =   &H00E4F8FC&
      Caption         =   "Debit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00645285&
      Height          =   255
      Index           =   1
      Left            =   9360
      TabIndex        =   13
      Top             =   1343
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.OptionButton OpCrDr 
      BackColor       =   &H00E4F8FC&
      Caption         =   "Credit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00645285&
      Height          =   255
      Index           =   0
      Left            =   8520
      TabIndex        =   12
      Top             =   1343
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.CommandButton cmdOut 
      Caption         =   "Outstanding"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   0
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "To Cancle Entered Information"
      Top             =   600
      Visible         =   0   'False
      Width           =   1395
   End
   Begin Crystal.CrystalReport cryReport 
      Left            =   120
      Top             =   840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileType   =   15
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   3240
      TabIndex        =   6
      Top             =   1270
      Width           =   1065
   End
   Begin VB.TextBox Txtche 
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
      Left            =   1320
      TabIndex        =   5
      Top             =   1283
      Width           =   1740
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   7800
   End
   Begin VB.CommandButton cmdBtnRetrieve 
      Caption         =   "&Retrieve Data"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Refresh Information"
      Top             =   600
      Width           =   1635
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "&P"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   2.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   10920
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Show Information"
      Top             =   600
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "E&xit"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "To Come Out From Software"
      Top             =   600
      Width           =   1035
   End
   Begin MSFlexGridLib.MSFlexGrid msgReport 
      Height          =   6390
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   13920
      _ExtentX        =   24553
      _ExtentY        =   11271
      _Version        =   393216
      Cols            =   15
      FixedCols       =   0
      RowHeightMin    =   275
      BackColor       =   16777215
      BackColorFixed  =   16512
      ForeColorFixed  =   -2147483639
      BackColorSel    =   16512
      ForeColorSel    =   -2147483640
      BackColorBkg    =   12632256
      GridColorFixed  =   8421504
      ScrollTrack     =   -1  'True
      FillStyle       =   1
      GridLinesFixed  =   1
      AllowUserResizing=   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComCtl2.DTPicker gfromdt 
      Height          =   390
      Left            =   4800
      TabIndex        =   9
      Top             =   1275
      Visible         =   0   'False
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   688
      _Version        =   393216
      Format          =   94371841
      CurrentDate     =   39505
   End
   Begin MSComCtl2.DTPicker gtodt 
      Height          =   390
      Left            =   6720
      TabIndex        =   10
      Top             =   1275
      Visible         =   0   'False
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   688
      _Version        =   393216
      Format          =   94371841
      CurrentDate     =   39505
   End
   Begin MSComctlLib.ProgressBar prgBar 
      Height          =   255
      Left            =   7800
      TabIndex        =   15
      Top             =   8520
      Visible         =   0   'False
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Please Wait............."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   315
      Left            =   600
      TabIndex        =   31
      Top             =   8520
      Visible         =   0   'False
      Width           =   3240
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00C1E3FD&
      BackStyle       =   0  'Transparent
      Caption         =   "Prefix :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00645285&
      Height          =   240
      Left            =   11280
      TabIndex        =   17
      Top             =   1320
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   315
      Left            =   240
      TabIndex        =   8
      Top             =   600
      Width           =   7080
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C1E3FD&
      BackStyle       =   0  'Transparent
      Caption         =   "Search  :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00645285&
      Height          =   240
      Left            =   240
      TabIndex        =   7
      Top             =   1350
      Width           =   795
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H000000C0&
      BorderWidth     =   5
      FillColor       =   &H00C00000&
      Height          =   8940
      Left            =   0
      Top             =   0
      Width           =   14205
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00808080&
      BorderColor     =   &H00000080&
      FillColor       =   &H00ECFFFF&
      FillStyle       =   0  'Solid
      Height          =   1335
      Left            =   120
      Top             =   480
      Width           =   13905
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000C0&
      Caption         =   "   Report"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   14175
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mGenOk As Boolean
Dim mLedgCtr As Integer
Dim mPresData As Boolean

Private Sub cmdMultiRef_Click()
'      mCallFirstTime = True
'      ReDim mAcArray(1)
'      mAcArray = MultipleAc
'      mCallForm = 2
'
End Sub

Private Sub cmd_Click(Index As Integer)
Select Case Index
Case 0
    cmdBtnRetrieve_Click
Case 2
   Frame1(0).Visible = False
End Select
   Frame1(0).Visible = False
   txtAccode = ""
   txtparty(1) = ""
End Sub

Private Sub cmdBtnRetrieve_Click()
If Label2 = "Daywise Account Ledger" Then
    If OpCrDr(0).Value = True Then
       procAccounts , gfromdt, gtodt, "Credit", TxtPrefix
    ElseIf OpCrDr(1).Value = True Then
       procAccounts , gfromdt, gtodt, "Debit", TxtPrefix
    Else
        procAccounts , gfromdt, gtodt, , txtAccode
    End If
End If

End Sub

Private Sub cmdFilter_Click(Index As Integer)
Frame1(0).Visible = True
txtclno.Visible = True
txtclno.Enabled = True
txtparty(1).Visible = True
txtparty(1).Enabled = True
txtAccode.Visible = False
End Sub

Private Sub cmdFind_Click()
Dim j As Integer, k As Integer, l As Integer
Dim FldName As String
Dim mCurCol As Integer
Dim mCurRow As Integer

mCurCol = msgReport.Col + 1
mCurRow = msgReport.Row


With msgReport

  For i = mCurRow To .rows - 1
    For j = mCurCol To .cols - 1
     If Txtche <> "" Then
        
        If UCase(Left(msgReport.TextMatrix(i, j), Len(Txtche))) = UCase(Txtche) Then
           .Row = i
           .Col = j
           .TopRow = .Row
           msgReport.SetFocus
           msgReport.CellBackColor = vbBlue
           Exit Sub
        End If
        
      Else
        MsgBox " First Select Record ", vbOKOnly
        Txtche.SetFocus
        Exit Sub
    End If
    
  Next j
  mCurCol = 0
  mCurRow = 0
  Next i
  
  If i <> msgReport.rows Or j <> msgReport.cols Then
        MsgBox "Record is Not Present", vbOKOnly
  End If

  End With
  
End Sub

Private Sub cmdOut_Click(Index As Integer)
Dim x1 As String
Dim AcCode As String
With msgReport
Select Case Index

Case 0
        DoEvents
        Frame1(1).Visible = True
        Frame3(2).Visible = True
        Frame4(1).Visible = True
        dtposfrom.Enabled = True
        dtposto.Enabled = True
        dtposto.Value = CDate(gCYEDate)
        dtposfrom.Value = CDate(gCYSDate)
        optDS(0).Enabled = True
        optDS(1).Enabled = True
         
Case 1

If msgReport.TextMatrix(4, 6) <> "" Then
                
                x1 = ""
        If msgReport.TextMatrix(4, 9) = 90017 Then
                AcCode = msgReport.TextMatrix(4, 6)
'''                   GProcCrystalRptPreparation cryReport, "Window", "Outstanding Sale Patywise"
                   If Trim(AcCode) <> "" Then x1 = Trim(AcCode)
                    If optDS(0).Value = True Then
                  '     CreateOutstandingSaleSimple "B", dtposto.Value, x1, , , dtposfrom.Value '-----------Simple
                    Else
                  '     CreateOutstandingSaleDetial "B", dtposto.Value, x1, , , dtposfrom.Value    '-----------Detial
                    End If
                  With cryReport
                        .Formulas(4) = "wFrToDt='From : '& '" & dtposfrom & "' & ' To ' & '" & dtposto & "'"
                        .Formulas(5) = "wRptName='Outstanding Sale Patywise '"
                        .SelectionFormula = "{tmpAccLedger.username} = '" & gUserName & "'"
                        .ReportTitle = "Outstanding Sale Patywise"
                     If optDS(0).Value = True Then
                        .ReportFileName = gReportPath & "rptOutstPartywise.rpt"  '--------------Simple
                     Else
                        .ReportFileName = gReportPath & "rptOutstanding.rpt"     '---------------Detial
                     End If
                        .Action = 1
                  End With
        ElseIf msgReport.TextMatrix(4, 9) = 90015 Or msgReport.TextMatrix(4, 9) = 90016 Or msgReport.TextMatrix(4, 9) = 90039 Then
                   AcCode = msgReport.TextMatrix(4, 6)
'''                   GProcCrystalRptPreparation cryReport, "Window", "Outstanding Purchase Patywise"
                   If Trim(AcCode) <> "" Then x1 = Trim(AcCode)
                    '    CreateOutstandingPurchDetial "B", dtposto.Value, x1, , dtposfrom.Value, optDS(0).Value  '--------------Detial
                      With cryReport
                            .Formulas(4) = "wFrToDt='From : '& '" & dtposfrom & "' & ' To ' & '" & dtposto & "'"
                            .Formulas(5) = "wRptName='" & frmTypeStr & "'"
                            .SelectionFormula = "{tmpAccLedger.username} = '" & gUserName & "'"
                            .ReportFileName = gReportPath & "rptOutstandingPurch.rpt"     '------------------Detial
                            .Action = 1
                      End With
        End If
End If

Case 2
        Frame1(1).Visible = False
        Frame3(2).Visible = False
        Frame4(0).Visible = False
        Frame4(1).Visible = False
End Select
End With
End Sub

Private Sub cmdPreview_Click()
cryReport.Destination = crptToWindow

'''GProcCrystalRptPreparation cryReport, "Window", ""

'frmRptSelection.cmdCrystal_Click (0)

End Sub



Private Sub CMDSTOP_Click()
    DoEvents
    Unload Me
   ' frmRptSelection.Show
End Sub

Private Sub Form_Activate()
'     Me.Width = 15675
'     Me.Height = 8490
'     Me.Top = 50
'     Me.Left = Screen.Width / 2 - Me.Width / 2
'
'     If mCallRow <> 0 And mCallFirstTime = False Then
'       If msfDataGrid.TopRow > mCallRow Then
'          msfDataGrid.TopRow = mCallRow
'       End If
'     End If
'     mCallFirstTime = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyEscape Then
         DoEvents
         Unload Me
     End If
End Sub

Private Sub Form_Load()
    Me.Top = 50
    Me.Left = Screen.Width / 2 - Me.Width / 2
     gfromdt = gCYSDate
     gtodt = gCYEDate
     If mCallRow <> 0 And mCallFirstTime = False Then
       If msfDataGrid.TopRow > mCallRow Then
          msfDataGrid.TopRow = mCallRow
       End If
     End If
     mCallFirstTime = False
End Sub
Private Sub cmdCancle_Click()
     Me.Hide
End Sub

Private Sub msgReport_Click()
With msgReport
  Select Case .Col
        Case 0, 1, 2, 3, 4, 5, 6, 7
            .CellForeColor = vbWhite
            .CellBackColor = vbBlue
  End Select
End With

End Sub

Private Sub msgReport_DblClick()
Label7.Visible = True
If Label2 = "Simple Sale Register" Then
    gfrmTypeStr = msgReport.TextMatrix(msgReport.Row, 1)
    mShowRec = Trim(msgReport.TextMatrix(msgReport.Row, 0))
    frmSale.Show
End If
If Label2 = "Trial Balance Simple" Then
'    gProcPreparePartyLedger gCYSDate, gCYEDate, Trim(msgReport.TextMatrix(msgReport.Row, 0))
    MenuRecStr = "Account"
    ReportRecStr = "Account Ledger"
    procAccounts Trim(msgReport.TextMatrix(msgReport.Row, 0)), gCYSDate, gCYEDate
End If
 If msgReport.Col > 5 Then
 
        If Val(msgReport.TextMatrix(msgReport.Row, 5)) <> 0 Then
         If Trim(msgReport.TextMatrix(msgReport.Row, 11)) <> "" Then
            Select Case msgReport.TextMatrix(msgReport.Row, 11)
                      Case "BR"
                               gfrmTypeStr = msgReport.TextMatrix(msgReport.Row, 11)
                               mShowRec = Trim(msgReport.TextMatrix(msgReport.Row, 10))
                               frmReceipt.Show
                       Case "CR"
                               gfrmTypeStr = msgReport.TextMatrix(msgReport.Row, 11)
                               mShowRec = Trim(msgReport.TextMatrix(msgReport.Row, 10))
                               frmCashReceipt.Show

                       Case "CN"
                               gfrmTypeStr = msgReport.TextMatrix(msgReport.Row, 11)
                               mShowRec = Trim(msgReport.TextMatrix(msgReport.Row, 10))
                               frmCrNote.Show
                       Case "CB"
                               gfrmTypeStr = msgReport.TextMatrix(msgReport.Row, 11)
                               mShowRec = Trim(msgReport.TextMatrix(msgReport.Row, 10))
                               frmCashBill.Show
                       Case "BP"
                               gfrmTypeStr = msgReport.TextMatrix(msgReport.Row, 11)
                               mShowRec = Trim(msgReport.TextMatrix(msgReport.Row, 10))
                               frmPayment.Show

                       Case "CP"
                               gfrmTypeStr = msgReport.TextMatrix(msgReport.Row, 11)
                               mShowRec = Trim(msgReport.TextMatrix(msgReport.Row, 10))
                               frmPayment.Show
                       Case "JV"
                               gfrmTypeStr = msgReport.TextMatrix(msgReport.Row, 11)
                               mShowRec = Trim(msgReport.TextMatrix(msgReport.Row, 10))
                               frmJV.Show
                       Case "SC"
                               gfrmTypeStr = msgReport.TextMatrix(msgReport.Row, 11)
                               mShowRec = Trim(msgReport.TextMatrix(msgReport.Row, 10))
                               frmSaleBill.Show
                       Case "SD"
                               gfrmTypeStr = msgReport.TextMatrix(msgReport.Row, 11)
                               mShowRec = Trim(msgReport.TextMatrix(msgReport.Row, 10))
                               frmCreditSaleBill.Show
                       Case "SB"
                               gfrmTypeStr = msgReport.TextMatrix(msgReport.Row, 11)
                               mShowRec = Trim(msgReport.TextMatrix(msgReport.Row, 10))
                               frmShortSaleBill.Show
                       Case "PR"
                               gfrmTypeStr = msgReport.TextMatrix(msgReport.Row, 11)
                               mShowRec = Trim(msgReport.TextMatrix(msgReport.Row, 10))
                               frmPurchase.Show
                       Case "PP"
                               gfrmTypeStr = msgReport.TextMatrix(msgReport.Row, 11)
                               mShowRec = Trim(msgReport.TextMatrix(msgReport.Row, 10))
                               frmPurchaseProcess.Show
                       Case "PC"
                               gfrmTypeStr = msgReport.TextMatrix(msgReport.Row, 11)
                               mShowRec = Trim(msgReport.TextMatrix(msgReport.Row, 10))
                               frmPurchaseCloth.Show
                               
                       Case "NP"
                               gfrmTypeStr = msgReport.TextMatrix(msgReport.Row, 11)
                               If gfrmTypeStr = "NP" Then gfrmTypeStr = "PR"
                               mShowRec = Trim(msgReport.TextMatrix(msgReport.Row, 10))
                               frmPurchase.Show
                               
                               
                       Case "PO"
                               gfrmTypeStr = msgReport.TextMatrix(msgReport.Row, 11)
                               mShowRec = Trim(msgReport.TextMatrix(msgReport.Row, 10))
                               frmPurchaseOther.Show
                       Case "DN"
                                gfrmTypeStr = msgReport.TextMatrix(msgReport.Row, 11)
                                mShowRec = Val(msgReport.TextMatrix(msgReport.Row, 10))
                                frmDrNote.Show
                       Case "IT"
                                gfrmTypeStr = msgReport.TextMatrix(msgReport.Row, 11)
                                mShowRec = Val(msgReport.TextMatrix(msgReport.Row, 10))
                                frmItNote.Show
                       Case "CV"
                                gfrmTypeStr = msgReport.TextMatrix(msgReport.Row, 11)
                                mShowRec = Val(msgReport.TextMatrix(msgReport.Row, 10))
                                frmContra.Show
            End Select
            End If
            End If
     Else
     
        If Val(msgReport.TextMatrix(msgReport.Row, 0)) <> 0 Then
        If Trim(msgReport.TextMatrix(msgReport.Row, 2)) <> "" Then      '-- And Trim(msgReport.TextMatrix(msgReport.Row, 2)) <> mFromEntry Then
            Select Case msgReport.TextMatrix(msgReport.Row, 2)
               Case "BR", "CR"
                       gfrmTypeStr = msgReport.TextMatrix(msgReport.Row, 2)
                       mShowRec = Val(msgReport.TextMatrix(msgReport.Row, 0))
                       frmReceipt.Show
               Case "BP", "CP"
                       gfrmTypeStr = msgReport.TextMatrix(msgReport.Row, 2)
                       mShowRec = Val(msgReport.TextMatrix(msgReport.Row, 0))
                       frmPayment.Show
               Case "JV", "PN", "SN"
                       gfrmTypeStr = msgReport.TextMatrix(msgReport.Row, 2)
                       mShowRec = Val(msgReport.TextMatrix(msgReport.Row, 0))
                       frmJV.Show
               Case "SY", "SD", "SO"
                       gfrmTypeStr = msgReport.TextMatrix(msgReport.Row, 2)
                       mShowRec = msgReport.TextMatrix(msgReport.Row, 0)
                       mShowRecSaleAc = GProcGetColumnValue("tblMastAccount", "AcName", msgReport.TextMatrix(msgReport.Row, 5), "S", "AcCode", "N")
                       mShowRecDt = (msgReport.TextMatrix(msgReport.Row, 1))
                       frmSales.Show
               Case "ST"
                       gfrmTypeStr = msgReport.TextMatrix(msgReport.Row, 2)
                       mShowRec = msgReport.TextMatrix(msgReport.Row, 0)
                       mShowRecDt = (msgReport.TextMatrix(msgReport.Row, 1))
                       frmSalesSIT.Show
               Case "SM"
                       gfrmTypeStr = msgReport.TextMatrix(msgReport.Row, 2)
                       mShowRec = msgReport.TextMatrix(msgReport.Row, 0)
                       mShowRecSaleAc = GProcGetColumnValue("tblMastAccount", "AcName", msgReport.TextMatrix(msgReport.Row, 5), "S", "AcCode", "N")
                       mShowRecDt = (msgReport.TextMatrix(msgReport.Row, 1))
                       frmSalesMillBill.Show
               Case "RY"
                       gfrmTypeStr = msgReport.TextMatrix(msgReport.Row, 2)
                       mShowRec = Val(msgReport.TextMatrix(msgReport.Row, 0))
                       frmSalesReturn.Show
               Case "PY"
                       gfrmTypeStr = msgReport.TextMatrix(msgReport.Row, 2)
                       mShowRec = Val(msgReport.TextMatrix(msgReport.Row, 0))
                       frmPurchaseTrade.Show
               Case "PT"
                       gfrmTypeStr = msgReport.TextMatrix(msgReport.Row, 2)
                       mShowRec = Val(msgReport.TextMatrix(msgReport.Row, 0))
                       frmPurchaseSIT.Show
               Case "PO"
                       gfrmTypeStr = msgReport.TextMatrix(msgReport.Row, 2)
                       mShowRec = Val(msgReport.TextMatrix(msgReport.Row, 0))
                       frmPurchaseOther.Show
               Case "PI"
                       gfrmTypeStr = msgReport.TextMatrix(msgReport.Row, 2)
                       mShowRec = Val(msgReport.TextMatrix(msgReport.Row, 0))
                       frmPurchaseInward.Show
               Case "VY"
                       gfrmTypeStr = msgReport.TextMatrix(msgReport.Row, 2)
                       mShowRec = Val(msgReport.TextMatrix(msgReport.Row, 0))
                       frmPurchaseReturn.Show
              End Select
          End If
          End If
End If
Label7.Visible = False
End Sub


Private Sub msgReport_EnterCell()
    With msgReport
      If .Row > 1 Then
        .CellForeColor = vbWhite
        .CellBackColor = vbBlue
      End If
    End With
End Sub

Private Sub msgReport_KeyPress(KeyAscii As Integer)
      If KeyAscii = vbKeyReturn Then
           Label7.Visible = True
           msgReport_DblClick
      End If

End Sub


Private Sub msgReport_LeaveCell()
  With msgReport
    If .Row > 1 Then
        .CellForeColor = vbBalck
        .CellBackColor = vbWhite
    End If
  End With
End Sub

Private Sub Timer1_Timer()
    DoEvents
End Sub

Private Sub txtparty_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 Then '-- Select
    gClsSearch.SearchMultiField "TblMastAccount,TblMastGroup", "AcName,AcCity,AcCode,AgName", Array("Party Name", "City Name", "", "Group Name"), Array(4200, 2200, 0, 2000), "TblMastAccount.AgCode=TblMastGroup.AgCode", Trim(Chr(KeyAscii)), "AcName", 0, 0, False
    KeyAscii = 0
    If gClsSearch.SearchMultiRetCol(0) <> "" Then
        txtparty(1) = gClsSearch.SearchMultiRetCol(0)
        txtcity = gClsSearch.SearchMultiRetCol(1)
        txtAccode = gClsSearch.SearchMultiRetCol(2)
    End If
End If

End Sub
