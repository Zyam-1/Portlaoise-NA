VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmAddCat 
   Caption         =   "NetAcquire - Categories "
   ClientHeight    =   4860
   ClientLeft      =   2730
   ClientTop       =   2070
   ClientWidth     =   4425
   Icon            =   "frmAddCat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   4425
   Begin VB.CommandButton cmdMoveUp 
      Caption         =   "Move Up"
      Enabled         =   0   'False
      Height          =   825
      Left            =   2880
      Picture         =   "frmAddCat.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2205
      Width           =   1245
   End
   Begin VB.CommandButton cmdMoveDown 
      Caption         =   "Move Down"
      Enabled         =   0   'False
      Height          =   780
      Left            =   2880
      Picture         =   "frmAddCat.frx":1C8C
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3060
      Width           =   1245
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   780
      Left            =   2880
      Picture         =   "frmAddCat.frx":360E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1260
      Width           =   1245
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   735
      Left            =   2880
      Picture         =   "frmAddCat.frx":4F90
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4050
      Width           =   1245
   End
   Begin VB.Frame Frame1 
      Caption         =   "Add new Category"
      Height          =   900
      Left            =   180
      TabIndex        =   7
      Top             =   180
      Width           =   3945
      Begin VB.TextBox txtCat 
         Height          =   285
         Left            =   120
         MaxLength       =   25
         TabIndex        =   0
         Top             =   300
         Width           =   1965
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   690
         Left            =   2835
         Picture         =   "frmAddCat.frx":6912
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   135
         Width           =   1050
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdCat 
      Height          =   3375
      Left            =   270
      TabIndex        =   6
      Top             =   1380
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   5953
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      FormatString    =   "<Category                 "
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
Attribute VB_Name = "frmAddCat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdadd_Click()

          Dim tb As New Recordset
          Dim sql As String

10        On Error GoTo cmdadd_Click_Error

20        txtCat = UCase$(Trim$(txtCat))

30        If txtCat = "" Then
40            txtCat.SetFocus
50            Exit Sub
60        End If

70        sql = "SELECT * from Categorys WHERE " & _
                "Cat = '" & txtCat & "'"
80        Set tb = New Recordset
90        RecOpenServer 0, tb, sql
100       If tb.EOF Then tb.AddNew
110       tb!Cat = txtCat
120       tb!ListOrder = grdCat.Rows
130       tb.Update

140       FillG

150       txtCat = ""
160       txtCat.SetFocus

170       Exit Sub

cmdadd_Click_Error:

          Dim strES As String
          Dim intEL As Integer

180       intEL = Erl
190       strES = Err.Description
200       LogError "frmAddCat", "cmdAdd_Click", intEL, strES

End Sub

Private Sub cmdCancel_Click()

10        Unload Me

End Sub

Private Sub cmdMoveDown_Click()

          Dim Num As Long
          Dim Str As String

10        On Error GoTo cmdMoveDown_Click_Error

20        If grdCat.Row = grdCat.Rows - 1 Then Exit Sub
30        Num = grdCat.Row

40        Str = grdCat

50        grdCat.RemoveItem Num
60        If Num < grdCat.Rows Then
70            grdCat.AddItem Str, Num + 1
80            grdCat.Row = Num + 1
90        Else
100           grdCat.AddItem Str
110           grdCat.Row = grdCat.Rows - 1
120       End If

130       grdCat.CellBackColor = vbYellow

140       cmdSave.Enabled = True

150       Exit Sub

cmdMoveDown_Click_Error:

          Dim strES As String
          Dim intEL As Integer

160       intEL = Erl
170       strES = Err.Description
180       LogError "frmAddCat", "cmdMoveDown_Click", intEL, strES

End Sub

Private Sub cmdMoveUp_Click()

          Dim Num As Long
          Dim Str As String

10        On Error GoTo cmdMoveUp_Click_Error

20        If grdCat.Row = 1 Then Exit Sub

30        Num = grdCat.Row

40        Str = grdCat

50        grdCat.RemoveItem Num
60        grdCat.AddItem Str, Num - 1

70        grdCat.Row = Num - 1
80        grdCat.CellBackColor = vbYellow

90        cmdSave.Enabled = True

100       Exit Sub

cmdMoveUp_Click_Error:

          Dim strES As String
          Dim intEL As Integer

110       intEL = Erl
120       strES = Err.Description
130       LogError "frmAddCat", "cmdMoveUp_Click", intEL, strES

End Sub

Private Sub cmdSave_Click()

          Dim tb As New Recordset
          Dim sql As String
          Dim Num As Long

10        On Error GoTo cmdSave_Click_Error

20        For Num = 1 To grdCat.Rows - 1
30            sql = "SELECT * from Categorys WHERE " & _
                    "Cat = '" & grdCat.TextMatrix(Num, 0) & "'"
40            Set tb = New Recordset
50            RecOpenServer 0, tb, sql
60            If tb.EOF Then tb.AddNew
70            tb!Cat = grdCat.TextMatrix(Num, 0)
80            tb!ListOrder = Num
90            tb.Update
100       Next

110       cmdSave.Enabled = False
120       cmdMoveUp.Enabled = False
130       cmdMoveDown.Enabled = False

140       Exit Sub

cmdSave_Click_Error:

          Dim strES As String
          Dim intEL As Integer

150       intEL = Erl
160       strES = Err.Description
170       LogError "frmAddCat", "cmdsave_Click", intEL, strES, sql

End Sub

Private Sub FillG()

          Dim tb As New Recordset
          Dim sql As String

10        On Error GoTo FillG_Error

20        ClearFGrid grdCat

30        sql = "SELECT * FROM Categorys"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql

60        Do While Not tb.EOF
70            grdCat.AddItem Trim(tb!Cat & "")
80            tb.MoveNext
90        Loop

100       FixG grdCat

110       Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer

120       intEL = Erl
130       strES = Err.Description
140       LogError "frmAddCat", "FillG", intEL, strES, sql

End Sub

Private Sub Form_Load()

10        FillG

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

10        On Error GoTo Form_QueryUnload_Error

20        If cmdSave.Enabled Then
30            If iMsg("Cancel without saving?", vbQuestion + vbYesNo) = vbNo Then
40                Cancel = True
50            End If
60        End If

70        Exit Sub

Form_QueryUnload_Error:

          Dim strES As String
          Dim intEL As Integer

80        intEL = Erl
90        strES = Err.Description
100       LogError "frmAddCat", "Form_QueryUnload", intEL, strES


End Sub

Private Sub grdCat_Click()

          Static SortOrder As Boolean
          Dim Num As Long
          Dim NumSave As Long

10        On Error GoTo grdCat_Click_Error

20        If grdCat.MouseRow = 0 Then
30            If SortOrder Then
40                grdCat.Sort = flexSortGenericAscending
50            Else
60                grdCat.Sort = flexSortGenericDescending
70            End If
80            SortOrder = Not SortOrder
90            cmdMoveUp.Enabled = False
100           cmdMoveDown.Enabled = False
110           Exit Sub
120       End If

130       NumSave = grdCat.Row

140       grdCat.Visible = False
150       grdCat.Col = 0
160       For Num = 1 To grdCat.Rows - 1
170           grdCat.Row = Num
180           If grdCat.CellBackColor = vbYellow Then
190               grdCat.CellBackColor = 0
200           End If
210       Next
220       grdCat.Row = NumSave
230       grdCat.Visible = True

240       grdCat.CellBackColor = vbYellow

250       cmdMoveUp.Enabled = True
260       cmdMoveDown.Enabled = True

270       Exit Sub

grdCat_Click_Error:

          Dim strES As String
          Dim intEL As Integer

280       intEL = Erl
290       strES = Err.Description
300       LogError "frmAddCat", "grdCat_Click", intEL, strES

End Sub

Private Sub txtCat_GotFocus()

10        On Error GoTo txtCat_GotFocus_Error

20        cmdMoveUp.Enabled = False
30        cmdMoveDown.Enabled = False

40        Exit Sub

txtCat_GotFocus_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmAddCat", "txtCat_GotFocus", intEL, strES

End Sub
