VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmExternalBatch 
   Caption         =   "NetAcquire - External Batches"
   ClientHeight    =   6690
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9885
   ClipControls    =   0   'False
   Icon            =   "frmExternalBatch.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6690
   ScaleWidth      =   9885
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save Changes"
      Height          =   795
      Left            =   8220
      Picture         =   "frmExternalBatch.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3780
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   795
      Left            =   8220
      Picture         =   "frmExternalBatch.frx":0974
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4770
      Width           =   1575
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "Insert Result"
      Height          =   795
      Left            =   8220
      Picture         =   "frmExternalBatch.frx":0DB6
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2550
      Width           =   1575
   End
   Begin VB.TextBox txtResult 
      Height          =   285
      Left            =   8190
      TabIndex        =   4
      Text            =   "Received"
      Top             =   2220
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   795
      Left            =   8220
      Picture         =   "frmExternalBatch.frx":11F8
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5730
      Width           =   1575
   End
   Begin VB.CommandButton cmdXL 
      Caption         =   "Export to Excel"
      Height          =   795
      Left            =   8220
      Picture         =   "frmExternalBatch.frx":1862
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   330
      Width           =   1575
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   6165
      Left            =   120
      TabIndex        =   0
      Top             =   330
      Width           =   8025
      _ExtentX        =   14155
      _ExtentY        =   10874
      _Version        =   393216
      Cols            =   6
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      BackColorSel    =   -2147483647
      ForeColorSel    =   -2147483624
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   2
      GridLines       =   3
      GridLinesFixed  =   3
      SelectionMode   =   1
      AllowUserResizing=   1
      FormatString    =   $"frmExternalBatch.frx":1B6C
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Result"
      Height          =   195
      Left            =   8220
      TabIndex        =   6
      Top             =   2010
      Width           =   450
   End
   Begin VB.Label lblExcelInfo 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exporting..."
      Height          =   345
      Left            =   8220
      TabIndex        =   3
      Top             =   1140
      Visible         =   0   'False
      Width           =   1545
   End
End
Attribute VB_Name = "frmExternalBatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub FillG()

          Dim tb As Recordset
          Dim sql As String
          Dim s As String

10        On Error GoTo FillG_Error

20        g.Rows = 2
30        g.AddItem ""
40        g.RemoveItem 1

50        sql = "SELECT TOP 50 D.SampleID, D.PatName, D.Chart, E.Analyte, E.Result, E.SendTo FROM " & _
                "Demographics AS D, ExtResults AS E WHERE " & _
                "COALESCE(E.Result, '') = '' " & _
                "AND D.SampleID = E.SampleID " & _
                "ORDER BY D.SampleID"
60        Set tb = New Recordset
70        RecOpenClient 0, tb, sql
80        Do While Not tb.EOF
90            s = CStr(tb!SampleID) & vbTab & _
                  tb!PatName & vbTab & _
                  tb!Chart & vbTab & _
                  tb!Analyte & vbTab & _
                  vbTab & _
                  tb!SendTo & ""
100           g.AddItem s
110           tb.MoveNext
120       Loop

130       If g.Rows > 2 Then
140           g.RemoveItem 1
150       End If
160       g.Row = 0
170       g.RowSel = 0

180       Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer

190       intEL = Erl
200       strES = Err.Description
210       LogError "frmExternalBatch", "FillG", intEL, strES, sql

End Sub

Private Sub cmdCancel_Click()

10        On Error GoTo cmdCancel_Click_Error

20        If cmdSave.Visible Then
30            If iMsg("Cancel without Saving?", vbYesNo + vbQuestion) = vbNo Then
40                Exit Sub
50            End If
60        End If

70        Unload Me

80        Exit Sub

cmdCancel_Click_Error:

          Dim strES As String
          Dim intEL As Integer

90        intEL = Erl
100       strES = Err.Description
110       LogError "frmExternalBatch", "cmdCancel_Click", intEL, strES

End Sub

Private Sub cmdInsert_Click()

          Dim intStart As Integer
          Dim intEnd As Integer
          Dim n As Integer

10        On Error GoTo cmdInsert_Click_Error

20        intStart = g.Row
30        intEnd = g.RowSel

40        If intStart = 0 Or intEnd = 0 Then
50            Exit Sub
60        End If

70        If intStart > intEnd Then
80            n = intStart
90            intStart = intEnd
100           intEnd = n
110       End If

120       For n = intStart To intEnd
130           g.TextMatrix(n, 4) = txtResult
140       Next

150       cmdSave.Visible = True

160       Exit Sub

cmdInsert_Click_Error:

          Dim strES As String
          Dim intEL As Integer

170       intEL = Erl
180       strES = Err.Description
190       LogError "frmExternalBatch", "cmdInsert_Click", intEL, strES

End Sub

Private Sub cmdRefresh_Click()

10        On Error GoTo cmdRefresh_Click_Error

20        FillG

30        Exit Sub

cmdRefresh_Click_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmExternalBatch", "cmdRefresh_Click", intEL, strES

End Sub

Private Sub cmdSave_Click()

          Dim tb As Recordset
          Dim n As Integer
          Dim sql As String

10        On Error GoTo cmdSave_Click_Error

20        For n = 1 To g.Rows - 1
30            If Trim(g.TextMatrix(n, 4)) <> "" Then    'result
40                sql = "Select * from ExtResults where " & _
                        "sampleid = '" & g.TextMatrix(n, 0) & "' " & _
                        "and Analyte = '" & g.TextMatrix(n, 3) & "'"
50                Set tb = New Recordset
60                RecOpenServer 0, tb, sql
70                If tb.EOF Then
80                    tb.AddNew
90                End If
100               tb!SampleID = g.TextMatrix(n, 0)
110               tb!Analyte = g.TextMatrix(n, 3)
120               tb!Result = g.TextMatrix(n, 4)
130               tb!Date = Format(Now, "dd/mmm/yyyy")
140               tb!Username = Username
150               tb!savetime = Format(Now, "dd/MMM/yyyy hh:mm:ss")
160               tb.Update
170           End If
180       Next

190       cmdSave.Visible = False

200       Exit Sub

cmdSave_Click_Error:

          Dim strES As String
          Dim intEL As Integer

210       intEL = Erl
220       strES = Err.Description
230       LogError "frmExternalBatch", "cmdsave_Click", intEL, strES, sql

End Sub

Private Sub cmdXL_Click()

10        On Error GoTo cmdXL_Click_Error

20        ExportFlexGrid g, Me

30        Exit Sub

cmdXL_Click_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmExternalBatch", "cmdXL_Click", intEL, strES

End Sub

Private Sub Form_Load()

10        On Error GoTo Form_Load_Error

20        FillG

30        Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmExternalBatch", "Form_Load", intEL, strES

End Sub


Private Sub g_Click()

          Static SortOrder As Boolean

10        On Error GoTo g_Click_Error

20        If g.MouseRow = 0 Then
30            If SortOrder Then
40                g.Sort = flexSortGenericAscending
50            Else
60                g.Sort = flexSortGenericDescending
70            End If
80            SortOrder = Not SortOrder
90        End If

100       Exit Sub

g_Click_Error:

          Dim strES As String
          Dim intEL As Integer
110       intEL = Erl
120       strES = Err.Description
130       LogError "frmExternalBatch", "g_Click", intEL, strES

End Sub


