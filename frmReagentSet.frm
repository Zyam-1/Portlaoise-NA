VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frmReagentSet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Set Reagent Levels"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6570
   Icon            =   "frmReagentSet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   6570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   690
      Left            =   5040
      Picture         =   "frmReagentSet.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4140
      Width           =   1245
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   690
      Left            =   5040
      Picture         =   "frmReagentSet.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5040
      Width           =   1245
   End
   Begin VB.Frame Frame1 
      Caption         =   "Reagent INformation"
      Height          =   1755
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   4875
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Default         =   -1  'True
         Enabled         =   0   'False
         Height          =   750
         Left            =   3600
         Picture         =   "frmReagentSet.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   900
         Width           =   1155
      End
      Begin VB.TextBox txtAmount 
         Height          =   285
         Left            =   2640
         TabIndex        =   13
         Top             =   540
         Width           =   675
      End
      Begin VB.ComboBox cmbTest 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1200
         Width           =   2415
      End
      Begin VB.ComboBox cmbReagents 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   540
         Width           =   2415
      End
      Begin VB.Label Label3 
         Caption         =   "Amount"
         Height          =   255
         Left            =   2640
         TabIndex        =   12
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Test Name"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   1515
      End
      Begin VB.Label Label1 
         Caption         =   "Reagent Name"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   300
         Width           =   1515
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   1680
      Left            =   4920
      TabIndex        =   0
      Top             =   0
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   2963
      _StockProps     =   15
      Caption         =   "Discipline"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   0
      Begin VB.OptionButton optDisp 
         Caption         =   "Biochemistry"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   6
         Tag             =   "Bio"
         Top             =   270
         Value           =   -1  'True
         Width           =   1275
      End
      Begin VB.OptionButton optDisp 
         Caption         =   "Coagulation"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   5
         Tag             =   "Coag"
         Top             =   495
         Width           =   1185
      End
      Begin VB.OptionButton optDisp 
         Caption         =   "Haematology"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   4
         Tag             =   "Haem"
         Top             =   720
         Width           =   1275
      End
      Begin VB.OptionButton optDisp 
         Caption         =   "Endocrinology"
         Height          =   195
         Index           =   3
         Left            =   105
         TabIndex        =   3
         Tag             =   "End"
         Top             =   930
         Width           =   1365
      End
      Begin VB.OptionButton optDisp 
         Caption         =   "Immunology"
         Height          =   195
         Index           =   4
         Left            =   105
         TabIndex        =   2
         Tag             =   "Imm"
         Top             =   1155
         Width           =   1320
      End
      Begin VB.OptionButton optDisp 
         Caption         =   "Blood Gas"
         Height          =   195
         Index           =   5
         Left            =   105
         TabIndex        =   1
         Tag             =   "BGA"
         Top             =   1380
         Width           =   1275
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdReg 
      Height          =   3945
      Left            =   60
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1800
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   6959
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      FormatString    =   "^Reagent Name         |^Test Name            |^Amount   "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmReagentSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbReagents_Change()

10        On Error GoTo cmbReagents_Change_Error

20        cmdAdd.Enabled = True

30        Exit Sub

cmbReagents_Change_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmReagentSet", "cmbReagents_Change", intEL, strES


End Sub

Private Sub cmbReagents_Click()


10        On Error GoTo cmbReagents_Click_Error

20        cmdAdd.Enabled = True

30        Exit Sub

cmbReagents_Click_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmReagentSet", "cmbReagents_Click", intEL, strES

End Sub

Private Sub cmdadd_Click()

10        On Error GoTo cmdadd_Click_Error

20        If cmbReagents <> "" And txtAmount <> "" And cmbTest <> "" Then

30            grdReg.AddItem cmbReagents & vbTab & cmbTest & vbTab & txtAmount
40            If grdReg.Rows > 2 And grdReg.TextMatrix(1, 0) = "" Then
50                grdReg.RemoveItem 1
60            End If
70            cmdSave.Enabled = True

80        End If



90        Exit Sub

cmdadd_Click_Error:

          Dim strES As String
          Dim intEL As Integer



100       intEL = Erl
110       strES = Err.Description
120       LogError "frmReagentSet", "cmdAdd_Click", intEL, strES


End Sub

Private Sub cmdCancel_Click()

10        Unload Me

End Sub

Private Sub Form_Load()

10        On Error GoTo Form_Load_Error

20        FillTests ("bio")
30        FillReagents

40        Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer



50        intEL = Erl
60        strES = Err.Description
70        LogError "frmReagentSet", "Form_Load", intEL, strES


End Sub

Private Sub optDisp_Click(Index As Integer)

10        On Error GoTo optDisp_Click_Error

20        FillTests (optDisp(Index).Tag)

30        Exit Sub

optDisp_Click_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmReagentSet", "optDisp_Click", intEL, strES


End Sub

Private Sub FillTests(ByVal Disp As String)
          Dim sql As String
          Dim tb As Recordset

10        On Error GoTo FillTests_Error

20        sql = "Select distinct(shortname) from " & Disp & "testdefinitions"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        Do While Not tb.EOF
60            cmbTest.AddItem tb!ShortName
70            tb.MoveNext
80        Loop

90        Exit Sub

FillTests_Error:

          Dim strES As String
          Dim intEL As Integer



100       intEL = Erl
110       strES = Err.Description
120       LogError "frmReagentSet", "FillTests", intEL, strES, sql

End Sub

Private Sub FillReagents()
          Dim sql As String
          Dim tb As Recordset

10        On Error GoTo FillReagents_Error

20        sql = "Select * from lists where listtype = 'RG'"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        Do While Not tb.EOF
60            cmbReagents.AddItem tb!Text
70            tb.MoveNext
80        Loop

90        Exit Sub

FillReagents_Error:

          Dim strES As String
          Dim intEL As Integer



100       intEL = Erl
110       strES = Err.Description
120       LogError "frmReagentSet", "FillReagents", intEL, strES, sql

End Sub

