VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmSystemErrorLog 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "NetAcquire - Error Log"
   ClientHeight    =   7020
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   15270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   15270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExport 
      Caption         =   "Export to Excel"
      Height          =   915
      Left            =   13995
      Picture         =   "frmSystemErrorLog.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4080
      Width           =   1230
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   915
      Left            =   13995
      Picture         =   "frmSystemErrorLog.frx":5C1E
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3090
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.ComboBox cmbMachine 
      Height          =   315
      Left            =   7410
      TabIndex        =   7
      Text            =   "cmbMachine"
      Top             =   390
      Width           =   2445
   End
   Begin VB.ComboBox cmbProcedure 
      Height          =   315
      Left            =   4860
      TabIndex        =   6
      Text            =   "cmbProcedure"
      Top             =   390
      Width           =   2445
   End
   Begin VB.ComboBox cmbModule 
      Height          =   315
      Left            =   2310
      TabIndex        =   5
      Text            =   "cmbModule"
      Top             =   390
      Width           =   2445
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Back"
      Height          =   915
      Left            =   13995
      Picture         =   "frmSystemErrorLog.frx":75A0
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5940
      Width           =   1230
   End
   Begin VB.Frame Frame1 
      Caption         =   "Max Records"
      Height          =   675
      Left            =   90
      TabIndex        =   1
      Top             =   180
      Width           =   1845
      Begin MSComCtl2.UpDown udRecords 
         Height          =   255
         Left            =   1080
         TabIndex        =   2
         Top             =   270
         Width           =   510
         _ExtentX        =   900
         _ExtentY        =   450
         _Version        =   393216
         Value           =   50
         BuddyControl    =   "lblMaxRecords"
         BuddyDispid     =   196617
         OrigLeft        =   1170
         OrigTop         =   270
         OrigRight       =   1680
         OrigBottom      =   525
         Increment       =   50
         Max             =   10000
         Min             =   50
         Orientation     =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label lblMaxRecords 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "50"
         Height          =   255
         Left            =   150
         TabIndex        =   3
         Top             =   270
         Width           =   915
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdError 
      Height          =   5955
      Left            =   90
      TabIndex        =   0
      Top             =   930
      Width           =   13845
      _ExtentX        =   24421
      _ExtentY        =   10504
      _Version        =   393216
      Cols            =   12
      RowHeightMin    =   400
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      WordWrap        =   -1  'True
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      HighLight       =   2
      GridLines       =   3
      GridLinesFixed  =   3
      AllowUserResizing=   3
      FormatString    =   $"frmSystemErrorLog.frx":7A3C
   End
   Begin VB.Label lblExcelInfo 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exporting..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   14010
      TabIndex        =   13
      Top             =   5070
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Module"
      Height          =   195
      Left            =   2340
      TabIndex        =   10
      Top             =   210
      Width           =   525
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Procedure"
      Height          =   195
      Left            =   4860
      TabIndex        =   9
      Top             =   210
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Machine"
      Height          =   195
      Left            =   7440
      TabIndex        =   8
      Top             =   210
      Width           =   615
   End
End
Attribute VB_Name = "frmSystemErrorLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim SortOrder As Boolean

Private Sub FillG()

          Dim sql As String
          Dim tb As Recordset
          Dim s As String

10        On Error GoTo FillG_Error

20        grdError.Rows = 2
30        grdError.AddItem ""
40        grdError.RemoveItem 1
50        grdError.Visible = False

60        sql = "SELECT TOP " & Val(lblMaxRecords) & " * FROM ErrorLog "
70        If cmbModule <> "" Then
80            sql = sql & "WHERE ModuleName = '" & cmbModule & "'"
90        ElseIf cmbProcedure <> "" Then
100           sql = sql & "WHERE ProcedureName = '" & cmbProcedure & "'"
110       ElseIf cmbMachine <> "" Then
120           sql = sql & "WHERE MachineName = '" & cmbMachine & "'"
130       End If
140       sql = sql & " ORDER BY DateTime DESC"
150       Set tb = New Recordset
160       RecOpenServer 0, tb, sql
170       Do While Not tb.EOF
180           s = vbTab & _
                  Format$(tb!Datetime, "General Date") & vbTab & _
                  tb!ModuleName & vbTab & _
                  tb!ProcedureName & vbTab & _
                  tb!ErrorLineNumber & vbTab & _
                  tb!SQLStatement & vbTab & _
                  tb!ErrorDescription & vbTab & _
                  tb!Username & vbTab & _
                  tb!MachineName & vbTab & _
                  tb!EventDesc & vbTab & _
                  tb!AppName & vbTab & _
                  tb!Guid & ""
190           grdError.AddItem s
200           tb.MoveNext
210       Loop

220       If grdError.Rows > 2 Then
230           grdError.RemoveItem 1
240       Else
250           s = vbTab & Format$(Now, "General Date") & vbTab & vbTab & "No Entries"
260           grdError.AddItem s
270           grdError.RemoveItem 1
280       End If
290       grdError.Visible = True

300       cmdDelete.Visible = False

310       Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer

320       intEL = Erl
330       strES = Err.Description
340       LogError "frmSystemErrorLog", "FillG", intEL, strES


End Sub

Private Sub FillCombos()

          Dim sql As String
          Dim tb As Recordset

10        On Error GoTo FillCombos_Error


20        cmbModule.Clear
30        cmbProcedure.Clear
40        cmbMachine.Clear

50        sql = "SELECT DISTINCT ModuleName FROM ErrorLog " & _
                "ORDER BY ModuleName"
60        Set tb = New Recordset
70        RecOpenServer 0, tb, sql
80        Do While Not tb.EOF
90            cmbModule.AddItem tb!ModuleName & ""
100           tb.MoveNext
110       Loop
120       cmbModule.AddItem "", 0

130       sql = "SELECT DISTINCT ProcedureName FROM ErrorLog " & _
                "ORDER BY ProcedureName"
140       Set tb = New Recordset
150       RecOpenServer 0, tb, sql
160       Do While Not tb.EOF
170           cmbProcedure.AddItem tb!ProcedureName & ""
180           tb.MoveNext
190       Loop
200       cmbProcedure.AddItem "", 0

210       sql = "SELECT DISTINCT MachineName FROM ErrorLog " & _
                "ORDER BY MachineName"
220       Set tb = New Recordset
230       RecOpenServer 0, tb, sql
240       Do While Not tb.EOF
250           cmbMachine.AddItem tb!MachineName & ""
260           tb.MoveNext
270       Loop
280       cmbMachine.AddItem "", 0

290       Exit Sub

FillCombos_Error:

          Dim strES As String
          Dim intEL As Integer

300       intEL = Erl
310       strES = Err.Description
320       LogError "frmSystemErrorLog", "FillCombos", intEL, strES


End Sub

Private Sub cmbMachine_Click()

10        On Error GoTo cmbMachine_Click_Error

20        cmbModule = ""
30        cmbProcedure = ""

40        FillG

50        Exit Sub

cmbMachine_Click_Error:

          Dim strES As String
          Dim intEL As Integer

60        intEL = Erl
70        strES = Err.Description
80        LogError "frmSystemErrorLog", "cmbMachine_Click", intEL, strES


End Sub


Private Sub cmbModule_Click()

10        On Error GoTo cmbModule_Click_Error

20        cmbProcedure = ""
30        cmbMachine = ""

40        FillG

50        Exit Sub

cmbModule_Click_Error:

          Dim strES As String
          Dim intEL As Integer

60        intEL = Erl
70        strES = Err.Description
80        LogError "frmSystemErrorLog", "cmbModule_Click", intEL, strES


End Sub


Private Sub cmbProcedure_Click()

10        On Error GoTo cmbProcedure_Click_Error

20        cmbModule = ""
30        cmbMachine = ""

40        FillG

50        Exit Sub

cmbProcedure_Click_Error:

          Dim strES As String
          Dim intEL As Integer

60        intEL = Erl
70        strES = Err.Description
80        LogError "frmSystemErrorLog", "cmbProcedure_Click", intEL, strES


End Sub


Private Sub cmdDelete_Click()

          Dim sql As String
          Dim StartRow As Integer
          Dim StopRow As Integer
          Dim n As Integer

10        On Error GoTo cmdDelete_Click_Error


20        If grdError.RowSel > grdError.Row Then
30            StartRow = grdError.Row
40            StopRow = grdError.RowSel
50        Else
60            StartRow = grdError.RowSel
70            StopRow = grdError.Row
80        End If

90        For n = StartRow To StopRow
100           sql = "DELETE FROM ErrorLog WHERE " & _
                    "GUID = '" & grdError.TextMatrix(n, 11) & "'"
110           Cnxn(0).Execute sql
120       Next

130       FillG

140       Exit Sub

cmdDelete_Click_Error:

          Dim strES As String
          Dim intEL As Integer

150       intEL = Erl
160       strES = Err.Description
170       LogError "frmSystemErrorLog", "cmdDelete_Click", intEL, strES


End Sub

Private Sub cmdExit_Click()

10        Unload Me

End Sub

Private Sub cmdExport_Click()

10        On Error GoTo cmdExport_Click_Error

20        ExportFlexGrid grdError, Me

30        Exit Sub

cmdExport_Click_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmSystemErrorLog", "cmdExport_Click", intEL, strES


End Sub

Private Sub Form_Load()

10        On Error GoTo Form_Load_Error

20        grdError.ColWidth(11) = 0

30        FillCombos
40        FillG

50        Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

60        intEL = Erl
70        strES = Err.Description
80        LogError "frmSystemErrorLog", "Form_Load", intEL, strES


End Sub




Private Sub grdError_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)

          Dim d1 As String
          Dim d2 As String

10        On Error GoTo grdError_Compare_Error

20        With grdError
30            If Not IsDate(.TextMatrix(Row1, .Col)) Then
40                Cmp = 0
50                Exit Sub
60            End If

70            If Not IsDate(.TextMatrix(Row2, .Col)) Then
80                Cmp = 0
90                Exit Sub
100           End If

110           d1 = Format(.TextMatrix(Row1, .Col), "General Date")
120           d2 = Format(.TextMatrix(Row2, .Col), "General Date")
130       End With

140       If SortOrder Then
150           Cmp = Sgn(DateDiff("s", d1, d2))
160       Else
170           Cmp = Sgn(DateDiff("s", d2, d1))
180       End If

190       Exit Sub

grdError_Compare_Error:

          Dim strES As String
          Dim intEL As Integer

200       intEL = Erl
210       strES = Err.Description
220       LogError "frmSystemErrorLog", "grdError_Compare", intEL, strES


End Sub




Private Sub grdError_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        On Error GoTo grdError_MouseUp_Error

20        With grdError

30            If .MouseRow = 0 And .MouseCol > 0 Then

40                If .MouseCol = 1 Then
50                    .Sort = 9
60                Else
70                    If SortOrder Then
80                        .Sort = flexSortGenericAscending
90                    Else
100                       .Sort = flexSortGenericDescending
110                   End If
120               End If
130               SortOrder = Not SortOrder

140               .ColSel = .Col
150               .RowSel = .Row
160               cmdDelete.Visible = False

170           Else
180               If (.ColSel <> .Col) Or (.RowSel <> .Row) Then
190                   cmdDelete.Visible = True
200               Else
210                   cmdDelete.Visible = False
220               End If
230           End If

240       End With

250       Exit Sub

grdError_MouseUp_Error:

          Dim strES As String
          Dim intEL As Integer

260       intEL = Erl
270       strES = Err.Description
280       LogError "frmSystemErrorLog", "grdError_MouseUp", intEL, strES


End Sub

Private Sub udRecords_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        On Error GoTo udRecords_MouseUp_Error

20        FillG

30        Exit Sub

udRecords_MouseUp_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmSystemErrorLog", "udRecords_MouseUp", intEL, strES


End Sub


