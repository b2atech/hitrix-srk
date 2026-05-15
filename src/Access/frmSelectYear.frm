VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmSelectYear 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Year To Work With"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6030
   Icon            =   "frmSelectYear.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   6030
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   400
      Left            =   4500
      TabIndex        =   1
      Top             =   3585
      Width           =   1200
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Select"
      Default         =   -1  'True
      Height          =   400
      Left            =   3300
      TabIndex        =   0
      Top             =   3585
      Width           =   1200
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   3285
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   5794
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
            Picture         =   "frmSelectYear.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSelectYear.frx":075C
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmSelectYear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Unloadflag As Boolean
Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdSelect_Click()
Dim FromDt As String
Dim ToDt As String

With ListView1.SelectedItem
    FromDt = Left(.Text, InStr(.Text, "-") - 1)
    ToDt = Right(.Text, Len(.Text) - InStr(.Text, "-"))
End With
HiMain.Caption = gCompName & " " & FromDt & "-" & ToDt
gDbYr = Right(Trim(FromDt), 2) & Right(Trim(ToDt), 2)
gDbNm = gCompCode + gDbYr
Unload Me
Call GProcConnectionOpen(gCn)
Unload Me
End Sub

Private Sub Form_Activate()
If Unloadflag = True Then
    MsgBox "No year exist to Select data"
    Unload Me
End If
End Sub

Private Sub Form_Load()
    Dim Rstbl As Recordset
    Dim x As String
    Dim Str As String
    
    x = "Select * from tblCompanyYear where CompCode='" & gCompCode & "'"
    GProcRstOpen Rstbl, x, "O", gSetupCn
    Do While Not Rstbl.EOF
        Str = Year(Rstbl.Fields("CompFdt")) & "-" & Year(Rstbl.Fields("CompTdt"))
        ListView1.ListItems.Add , , Str, 1
        Rstbl.MoveNext
    Loop
    Unloadflag = False
    If ListView1.ListItems.Count = 0 Then
        Unloadflag = True
    End If
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


