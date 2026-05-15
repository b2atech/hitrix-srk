VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmPayRec 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C2E3FC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gate Pass Entry"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8715
   Icon            =   "frmPayRec.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   8715
   Begin VB.ListBox List1 
      Height          =   1035
      Left            =   4560
      TabIndex        =   52
      Top             =   2400
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txttemp1 
      Height          =   325
      Index           =   7
      Left            =   1560
      MaxLength       =   20
      TabIndex        =   10
      Text            =   "txttemp1(7)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   3360
      Width           =   4800
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   17
      Left            =   9120
      TabIndex        =   51
      Text            =   "17"
      Top             =   6000
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   16
      Left            =   9120
      TabIndex        =   50
      Text            =   "16"
      Top             =   5520
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txttemp1 
      Height          =   345
      Index           =   8
      Left            =   4560
      TabIndex        =   6
      Text            =   "8"
      Top             =   2400
      Width           =   375
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   1935
      Left            =   120
      TabIndex        =   13
      Top             =   4440
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   3413
      _Version        =   393216
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   8
      Left            =   1560
      MaxLength       =   40
      TabIndex        =   11
      Text            =   "8"
      Top             =   3690
      Width           =   4800
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   15
      Left            =   9120
      TabIndex        =   49
      Text            =   "15"
      Top             =   5040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
      Height          =   375
      Index           =   12
      Left            =   9030
      TabIndex        =   47
      Text            =   "12"
      Top             =   2925
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   10
      Left            =   9000
      TabIndex        =   46
      Text            =   "10"
      Top             =   2640
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   1560
      TabIndex        =   1
      Text            =   "0"
      Top             =   840
      Width           =   1320
   End
   Begin VB.TextBox txttemp 
      Height          =   375
      Index           =   5
      Left            =   1560
      MaxLength       =   40
      TabIndex        =   4
      Text            =   "txttemp(5)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1800
      Width           =   5040
   End
   Begin VB.TextBox txttemp 
      Height          =   375
      Index           =   4
      Left            =   1560
      MaxLength       =   40
      TabIndex        =   3
      Text            =   "txttemp(4)"
      ToolTipText     =   "Press F3 for Creation"
      Top             =   1320
      Width           =   5040
   End
   Begin VB.TextBox txt 
      Height          =   375
      Index           =   14
      Left            =   9120
      TabIndex        =   21
      Text            =   "14"
      Top             =   3720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   360
      Index           =   13
      Left            =   9000
      TabIndex        =   20
      Text            =   "13"
      Top             =   3360
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   0
      Left            =   7230
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "To Add New Record"
      Top             =   650
      Width           =   495
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   1
      Left            =   7830
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "To Add New Record"
      Top             =   650
      Width           =   495
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   2
      Left            =   7230
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "To Add New Record"
      Top             =   1100
      Width           =   495
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   3
      Left            =   7830
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "To Add New Record"
      Top             =   1100
      Width           =   495
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "&Find"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   4
      Left            =   7230
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "To Add New Record"
      Top             =   1550
      Width           =   1095
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "&List"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   5
      Left            =   7230
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "To Add New Record"
      Top             =   2000
      Width           =   1095
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "&Add"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   6
      Left            =   7230
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "To Add New Record"
      Top             =   2700
      Width           =   1095
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "&Modify"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   7
      Left            =   7230
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "To Add New Record"
      Top             =   3150
      Width           =   1095
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   8
      Left            =   7230
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "To Add New Record"
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   9
      Left            =   7230
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "To Add New Record"
      Top             =   4050
      Width           =   1095
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   13
      Left            =   7230
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "To Add New Record"
      Top             =   4750
      Width           =   1095
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   14
      Left            =   7230
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   "To Add New Record"
      Top             =   5200
      Width           =   1095
   End
   Begin VB.CommandButton cmdBtn 
      Cancel          =   -1  'True
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   10
      Left            =   7230
      Style           =   1  'Graphical
      TabIndex        =   39
      ToolTipText     =   "To Add New Record"
      Top             =   5650
      Width           =   1095
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "&Label"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   11
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   40
      ToolTipText     =   "To Add New Record"
      Top             =   4170
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "Optional"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   12
      Left            =   9060
      Style           =   1  'Graphical
      TabIndex        =   41
      ToolTipText     =   "To Add New Record"
      Top             =   4605
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   3
      Left            =   9120
      TabIndex        =   44
      Text            =   "3"
      Top             =   1320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   360
      Index           =   5
      Left            =   9120
      TabIndex        =   19
      Text            =   "5"
      Top             =   2040
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   360
      Index           =   6
      Left            =   1560
      TabIndex        =   5
      Text            =   "6"
      Top             =   2400
      Width           =   1560
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   9
      Left            =   1560
      MaxLength       =   40
      TabIndex        =   12
      Text            =   "9"
      Top             =   4020
      Width           =   4800
   End
   Begin VB.TextBox txt 
      Height          =   325
      Index           =   7
      Left            =   9000
      TabIndex        =   17
      Text            =   "7"
      Top             =   2400
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   11
      Left            =   5040
      TabIndex        =   7
      Text            =   "11"
      Top             =   2400
      Width           =   1575
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   360
      Index           =   4
      Left            =   9120
      TabIndex        =   18
      Text            =   "4"
      Top             =   1680
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.TextBox txt 
      Height          =   345
      Index           =   2
      Left            =   9120
      TabIndex        =   43
      Text            =   "2"
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   1
      Left            =   9000
      TabIndex        =   42
      Text            =   "1"
      Top             =   600
      Visible         =   0   'False
      Width           =   555
   End
   Begin MSMask.MaskEdBox mskDt 
      Height          =   360
      Index           =   2
      Left            =   5040
      TabIndex        =   2
      Top             =   840
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   635
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskDt 
      Height          =   360
      Index           =   12
      Left            =   5040
      TabIndex        =   8
      Top             =   2880
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   635
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskDt 
      Height          =   360
      Index           =   13
      Left            =   1560
      TabIndex        =   9
      Top             =   2880
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   635
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label lblActionStatus 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "View Record.."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   6075
      TabIndex        =   48
      Top             =   45
      Width           =   2475
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Entry No.   :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   240
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Rcon.Date :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Index           =   10
      Left            =   360
      TabIndex        =   26
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount    :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Index           =   9
      Left            =   360
      TabIndex        =   23
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date        :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Index           =   8
      Left            =   3960
      TabIndex        =   22
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "To            :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Index           =   4
      Left            =   360
      TabIndex        =   16
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "From        :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Index           =   3
      Left            =   360
      TabIndex        =   15
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date   :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   240
      Index           =   1
      Left            =   4080
      TabIndex        =   14
      Top             =   840
      Width           =   675
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00BDDFD1&
      FillStyle       =   0  'Solid
      Height          =   5610
      Left            =   6990
      Shape           =   4  'Rounded Rectangle
      Top             =   570
      Width           =   1575
   End
   Begin VB.Label Label6 
      BackColor       =   &H004D177D&
      Caption         =   "   Bank Payments"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Left            =   0
      TabIndex        =   45
      Top             =   0
      Width           =   9015
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Narration  :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Index           =   11
      Left            =   360
      TabIndex        =   25
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ref No. :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   240
      Index           =   13
      Left            =   3600
      TabIndex        =   24
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00808080&
      BorderColor     =   &H004D177D&
      FillColor       =   &H00ECFFFF&
      FillStyle       =   0  'Solid
      Height          =   3945
      Left            =   105
      Top             =   540
      Width           =   6690
   End
End
Attribute VB_Name = "frmPayRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const tblName As String = "tblVoucher"
Const IdField As String = "vno"
Const MaxNo As Long = 18
Dim CriteriaStr As String
Dim rstbl As Recordset, ShowFrmFlag As Boolean
Dim varBookMark As Variant
Dim FirstTimeFlag As Boolean
Dim FormAction As Integer
Private Sub cmdBtn_Click(index As Integer)
'On Error GoTo ErrorRoutine
Select Case index
Case 0 'Previous
    rstbl.MovePrevious
    If rstbl.BOF = True Then
        rstbl.MoveFirst
        MsgBox ("This is First Record"), vbInformation + vbOKOnly, "First Record"
    End If
    ReadFields
Case 1 'Next
    rstbl.MoveNext
    If rstbl.EOF = True Then
        rstbl.MoveLast
        MsgBox ("This is Last Record"), vbInformation + vbOKOnly, "Last Record"
    End If
    ReadFields
Case 2 'First
    rstbl.MoveFirst
    ReadFields
Case 3 'Last
    rstbl.MoveLast
    ReadFields
Case 4 'Find
'    FindFlag = True
'    gSearchMultiCaption(0) = "Gate Pass No"
'    gSearchMultiCaption(1) = "Date"
'    gSearchMultiCaption(2) = "Party"
'    gSearchMultiCaption(3) = "Mill"
'    GProcSearchMultiField Me, tblName, "vno,Vdt,tblMastAccount.AcName,M.", "", True, ""
Case 5 'List
    'GenMastList ("LAC")
Case 6 'Add
      FormAction = vbDataActionAddNew
     Call GProcActivateControls(Me, True)
     Call GProcSetButtons(Me, vbDataActionAddNew)
     Call GProcClearForm(Me, rstbl, MaxNo, True)
     txt(0).Enabled = False
     DoEvents
     txt(0).Text = GProcGenerateId(tblName, "VNO", CriteriaStr, gCn)
    If FirstTimeFlag = True Then
        FirstTimeFlag = False
'    Else
'        txt(1).SetFocus
    End If
    FillCombo
       
Case 7 'Modify
    'If mURecModi = True Then
    '    SearchHelp
    '    DoEvents
        FormAction = vbDataActionUpdate
        Call GProcActivateControls(Me, True)
        Call GProcSetButtons(Me, vbDataActionUpdate)
        txt(0).Enabled = False
'        txt(1).SetFocus
    'End If
Case 8 'Delete
   ' If mURecDel = True Then
     If IsValidForDelete = True Then
        If MsgBox("Do You Really Delete This Record?", vbQuestion + vbYesNo + vbDefaultButton1, "Confirm Delete") = vbYes Then
           gCn.Execute "Delete from tblvoucher where " & CriteriaStr & " and vno = " & txt(0)
            If rstbl.RecordCount > 1 Then
                cmdBtn_Click (1)
            ElseIf rstbl.RecordCount = 1 Then
                Call GProcActivateControls(Me, False)
                Call GProcSetButtons(Me, vbDataActionCancel)
                Form_Load
            End If
        End If
        DoEvents
     Else
        MsgBox "You can't Delete This Record. Detail Adjustments are Done."
     End If
   ' End If
Case 9 'Print
Case 10 'Exit
    DoEvents
    Unload Me
Case 13 'Save
    If ValidateData = True Then
        FillTxtFromTemp
   '     GProcSaveRecord Me, rstbl, gFormAction
        If gfrmTypeStr = "BP" Or gfrmTypeStr = "CP" Then
            GProcCreateVoucher FormAction, CLng(txt(0)), 1, CDate(txt(2)), gfrmTypeStr, CLng(txt(4)), CLng(txt(5)), Val(txt(6)), txt(11), txt(12), txt(13), txt(7), txt(8), txt(9), txt(10)
            GProcCreateVoucher FormAction, txt(0), 2, txt(2), gfrmTypeStr, txt(5), txt(4), Val(txt(6)) * (-1), txt(11), txt(12), txt(13), txt(7), txt(8), txt(9), txt(10)
        Else
            GProcCreateVoucher FormAction, txt(0), 1, txt(2), gfrmTypeStr, txt(4), txt(5), Val(txt(6)) * (-1), txt(11), txt(12), txt(13), txt(7), txt(8), txt(9), txt(10)
            GProcCreateVoucher FormAction, txt(0), 2, txt(2), gfrmTypeStr, txt(5), txt(4), Val(txt(6)), txt(11), txt(12), txt(13), txt(7), txt(8), txt(9), txt(10)
        End If
        rstbl.Requery
        Call GProcActivateControls(Me, False)
        Call GProcSetButtons(Me, True)
        FormAction = vbDataActionClose
    End If
Case 14 'Cancel
    Call GProcClearForm(Me, rstbl, MaxNo, True)
    If rstbl.RecordCount <= 0 Then
        Call GProcActivateControls(Me, False)
        Call GProcSetButtons(Me, vbDataActionCancel)
    Else
        If FormAction = vbDataActionAddNew Then
            rstbl.MoveLast
        Else
            rstbl.CancelUpdate
        End If
        ReadFields
        Call GProcActivateControls(Me, False)
        Call GProcSetButtons(Me, True)
    End If
    FormAction = vbDataActionClose
End Select
Exit Sub
ErrorRoutine:
    GProcErrorHandler
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 13:
            SendKeys "{TAB}"
        Case 27:
            If (cmdBtn(13).Enabled = False) Then
                Unload Me
            End If
    End Select
End Sub
Private Sub Form_Load()
    Dim x As String
    Dim Y As String
    Me.Top = 300
    List1.Clear
    List1.AddItem "Chq. "
    List1.AddItem "T.T. "
    List1.AddItem "D.D. "
    List1.AddItem "Hundi"
    List1.AddItem "Other"
    Me.Left = Screen.Width / 2 - Me.Width / 2
    Me.Caption = "Bank Payments for " & gCName & " " & gFinYr
    CriteriaStr = " VType = '" & gfrmTypeStr & "'"
    Y = "Select * from " & tblName & " where " & CriteriaStr & " and vctrno = 1"
    GProcRstOpen rstbl, Y, "O"
    Select Case gfrmTypeStr
    Case "BP"
    Case "BR"
        Label6.Caption = "Bank Receipts"
        lbl1(3).Top = 1800
        lbl1(4).Top = 1320
    Case "CR"
        Label6.Caption = "Cash Receipts"
        lbl1(3).Top = 1800
        lbl1(4).Top = 1320
        lbl1(10).Visible = False
        mskDt(13).Visible = False
        mskDt(12).Visible = False
        lbl1(8).Visible = False
        lbl1(13).Visible = False
        txttemp1(8).Visible = False
        txt(11).Visible = False
        txttemp1(7).MaxLength = 40
    Case "CP"
        Label6.Caption = "Cash Payments"
        lbl1(10).Visible = False
        mskDt(13).Visible = False
        mskDt(12).Visible = False
        lbl1(8).Visible = False
        lbl1(13).Visible = False
        txttemp1(8).Visible = False
        txt(11).Visible = False
        txttemp1(7).MaxLength = 40
    End Select
    FirstTimeFlag = True
    cmdBtn_Click (6)
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'If gFormAction = vbDataActionUpdate Then
'    MsgBox "You are about to modify the record. Save the record and then exit.", vbInformation + vbOKOnly, Me.Caption
'    Cancel = True
'    Exit Sub
'End If
rstbl.Requery
rstbl.Close
End Sub
Private Sub txt_GotFocus(index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub txt_LostFocus(index As Integer)
GProcCheckForNumber Me, rstbl, MaxNo
End Sub
Private Sub txt_KeyPress(index As Integer, KeyAscii As Integer)
KeyAscii = GProcValidateKey(rstbl, index, KeyAscii, txt(index).Text)
End Sub
Private Sub FillCombo()
If ((gfrmTypeStr = "CP") Or (gfrmTypeStr = "CR")) Then
   txt(4) = 2
   txttemp(4) = GProcGetColumnValue("tblmastaccount", "Accode", 2, "N", "Acname", "S")
End If
'    cbo(1).Clear
'    cbo(1).AddItem "Cheque"
'    cbo(1).AddItem "D.D."
'    cbo(1).AddItem "T.T."
'    cbo(1).AddItem "Hundi"
'    cbo(1).AddItem "Other"
End Sub
Private Sub SaveBtnEd()
If txt(0) = "" Or txt(4) = "" Then
    cmdBtn(13).Enabled = False
    cmdBtn(13).BackColor = &H80000004
Else
    cmdBtn(13).Enabled = True
    cmdBtn(13).BackColor = &HFFFF80
End If
End Sub
Private Sub FillTxtFromTemp()
txt(1) = 1
txt(2) = IIf(IsDate(mskDt(2)), mskDt(2), "") '--Vdt
txt(3) = gfrmTypeStr
txt(12) = IIf(IsDate(mskDt(12)), mskDt(12), "") '--Chqdt
txt(13) = IIf(IsDate(mskDt(13)), mskDt(13), "") '--Rcondt
If Left(gfrmTypeStr, 1) = "B" Then
   txt(7) = Left(txttemp1(8) + Space(5), 5) + " No. " + txttemp1(7)
Else
   txt(7) = txttemp1(7)
End If
txt(4) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(4), "S", "Accode", "N") ', gCn)
txt(5) = GProcGetColumnValue("tblMastAccount", "AcName", txttemp(5), "S", "AcCode", "N") ', gCn)
End Sub
Private Sub FillTempFromTxt()
mskDt(2) = IIf(IsDate(txt(2)), CDate(txt(2)), "__/__/____") '--Vdt
If IsDate(txt(12)) Then
   mskDt(12) = CDate(txt(12))
Else
   mskDt(12) = "__/__/____"
End If
If IsDate(txt(13)) Then
   mskDt(13) = CDate(txt(13))
Else
   mskDt(13) = "__/__/____"
End If
txt(6) = Format(Abs(Val(txt(6))), "##0.00")
If Left(gfrmTypeStr, 1) = "B" Then
    txttemp1(7) = Mid(txt(7), 21)
    txttemp1(8) = Left(txt(7), 5)
Else
    txttemp1(7) = txt(7)
End If
txttemp(4) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(4), "N", "AcName", "S") ', gCn)
txttemp(5) = GProcGetColumnValue("tblMastAccount", "AcCode", txt(5), "N", "AcName", "S") ', gCn)
End Sub
Private Function ValidateData() As Boolean
If txttemp(4) = "" Or txttemp(5) = "" Or Val(txt(6)) = 0 Then
   MsgBox "Incomplete Entry.."
   ValidateData = False
   Exit Function
End If
ValidateData = True
End Function
Private Function IsValidForDelete() As Boolean
'If Then
'   MsgBox "Incomplete Entry.."
'   IsValidForDelete = False
'   Exit Function
'End If
IsValidForDelete = True
End Function
Private Sub ReadFields()
GProcShowRecord Me, rstbl, MaxNo
FillTempFromTxt
End Sub
Private Sub txttemp_GotFocus(index As Integer)
GProcSelectBox Me.ActiveControl
End Sub
Private Sub txttemp_KeyPress(index As Integer, KeyAscii As Integer)
Dim LeftPos As Long
Dim TopPos As Long
TopPos = txttemp(index).Top + Me.Top + 650
LeftPos = txttemp(index).Left + Me.Left
If KeyAscii = 13 Then Exit Sub
If ((gfrmTypeStr = "CP") Or (gfrmTypeStr = "CR")) And (index = 4) Then
   KeyAscii = 0
   txt(index) = 2
   txttemp(index) = GProcGetColumnValue("tblmastaccount", "Accode", 2, "N", "Acname", "S")
   Exit Sub
End If
Select Case index
Case 4: '--Bank Cash
    gClsSearch.SearchMultiField "tblMastAccount", "AcName,Accode", Array("Account Name", "Aid"), Array(txttemp(index).Width, 0), "AgCode=90019 ", Chr(KeyAscii), "AcName", LeftPos, TopPos, True
Case 5: '--Account
    gClsSearch.SearchMultiField "tblMastAccount", "AcName,Accode", Array("Account Name"), Array(txttemp(index).Width, 0), "", Chr(KeyAscii), "acName", LeftPos, TopPos, True
End Select
DoEvents
KeyAscii = 0
If gClsSearch.SearchMultiRetCol(0) <> "" Then
    txttemp(index) = gClsSearch.SearchMultiRetCol(0)
    txt(index) = gClsSearch.SearchMultiRetCol(1)
End If
End Sub
Private Sub txttemp1_GotFocus(index As Integer)
If index = 8 Then
   List1.Visible = True
   List1.SetFocus
End If
End Sub
Private Sub list1_DblClick()
    Call list1_KeyPress(13)
    DoEvents
End Sub
Private Sub list1_GotFocus()
If Trim(txttemp1(8).Text) <> "" Then
    List1.Text = txttemp1(8).Text
Else
    List1.ListIndex = 0
End If
End Sub
Private Sub list1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txttemp1(8).Text = Trim(List1.Text)
    List1.Visible = False
    DoEvents
    txt(11).SetFocus
End If
End Sub
Private Sub list1_LostFocus()
    Call list1_KeyPress(13)
    DoEvents
End Sub

