VERSION 5.00
Begin VB.Form FrmQueryChecking 
   Caption         =   "Form1"
   ClientHeight    =   3795
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6405
   LinkTopic       =   "Form1"
   ScaleHeight     =   3795
   ScaleWidth      =   6405
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Show"
      Height          =   375
      Index           =   1
      Left            =   2880
      TabIndex        =   5
      Top             =   3240
      Width           =   1335
   End
   Begin VB.TextBox txt 
      Height          =   1095
      Index           =   1
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1800
      Width           =   5415
   End
   Begin VB.TextBox txt 
      Height          =   1095
      Index           =   0
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   5415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Update"
      Height          =   375
      Index           =   0
      Left            =   720
      TabIndex        =   0
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Result"
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   4
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Query"
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "FrmQueryChecking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
On Error Resume Next
If Index = 0 Then
   gCn.Execute txt(0)
Else
   Dim tmp1 As New ADODB.Recordset, I As Integer
   If GProcRstOpen(tmp1, txt(0)) > 0 Then
   txt(1) = ""
   tmp1.MoveFirst
   Do While Not tmp1.EOF
   For I = 0 To tmp1.Fields.Count - 1
   txt(1) = txt(1) & tmp1.Fields(I) & "  "
   Next I
   txt(1) = txt(1) & Chr(13)
   tmp1.MoveNext
   Loop
   End If
End If
End Sub
