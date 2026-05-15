VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmUtlSelectYear 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Year To Work With"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6030
   Icon            =   "frmUtlSelectYear.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   6030
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   400
      Left            =   4395
      TabIndex        =   1
      Top             =   3870
      Width           =   1200
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Select"
      Default         =   -1  'True
      Height          =   400
      Left            =   3075
      TabIndex        =   0
      Top             =   3870
      Width           =   1200
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   3060
      Left            =   315
      TabIndex        =   2
      Top             =   345
      Width           =   5250
      _ExtentX        =   9260
      _ExtentY        =   5398
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      Icons           =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00000080&
      FillColor       =   &H00EDFDFE&
      FillStyle       =   0  'Solid
      Height          =   3495
      Left            =   120
      Top             =   150
      Width           =   5655
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00BDDFD1&
      FillStyle       =   0  'Solid
      Height          =   705
      Left            =   2880
      Shape           =   4  'Rounded Rectangle
      Top             =   3735
      Width           =   2880
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   2520
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmUtlSelectYear.frx":058A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmUtlSelectYear.frx":08A4
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmUtlSelectYear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Unloadflag As Boolean
Dim rstbl As Recordset
Private Sub CmdClose_Click()
Unload Me
End Sub
Private Sub cmdSelect_Click()
Dim FromYr As String
With ListView1.SelectedItem
    FromYr = Mid(.Text, 3, 2) + Mid(.Text, 8, 2)
End With
rstbl.MoveFirst
rstbl.Find "CompYear='" & FromYr & "'"
If Not rstbl.EOF Then
    GProcGetCompDetails rstbl
End If
Unload Me
End Sub
Private Sub Form_Activate()
If Unloadflag = True Then
    MsgBox "No year exist to Select data"
    Unload Me
End If
End Sub
Private Sub Form_Load()
    Dim X As String
    Dim Str As String
    Me.Top = 300
    Me.Left = Screen.Width / 2 - Me.Width / 2
    X = "Select * from tblCompanyYear where CompCode='" & gCCode & "'"
    GProcRstOpen rstbl, X, "O", gSetupCn
    Do While Not rstbl.EOF
        Str = Year(rstbl.Fields("CompFdt")) & "-" & Year(rstbl.Fields("CompTdt"))
        ListView1.ListItems.Add , , Str, 1
        rstbl.MoveNext
    Loop
    Unloadflag = False
    If ListView1.ListItems.Count = 0 Then
        Unloadflag = True
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
rstbl.Close
End Sub
Private Sub ListView1_BeforeLabelEdit(Cancel As Integer)
Cancel = True
End Sub
Private Sub ListView1_Click()
Dim i As Long
    For i = 1 To ListView1.ListItems.Count
        ListView1.ListItems.Item(i).Icon = 1
    Next i
    ListView1.SelectedItem.Icon = 2
End Sub


