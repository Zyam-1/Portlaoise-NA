VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frmImmBatch 
   Caption         =   "NetAcquirre - Batch AutoImmune Entry"
   ClientHeight    =   7470
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12090
   Icon            =   "frmImmBatch.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7470
   ScaleWidth      =   12090
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1275
      Left            =   150
      TabIndex        =   8
      Top             =   60
      Width           =   2475
      Begin VB.TextBox txtFrom 
         Height          =   285
         Left            =   90
         TabIndex        =   10
         Top             =   390
         Width           =   1110
      End
      Begin VB.TextBox txtTo 
         Height          =   285
         Left            =   90
         TabIndex        =   9
         Top             =   900
         Width           =   1110
      End
      Begin Threed.SSCommand cmdCo 
         Height          =   810
         Left            =   1320
         TabIndex        =   13
         Top             =   390
         Width           =   1035
         _Version        =   65536
         _ExtentX        =   1826
         _ExtentY        =   1429
         _StockProps     =   78
         Caption         =   "Get Names"
         Picture         =   "frmImmBatch.frx":030A
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   " Sample ID's "
         Height          =   195
         Left            =   750
         TabIndex        =   14
         Top             =   0
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "To"
         Height          =   195
         Left            =   90
         TabIndex        =   12
         Top             =   720
         Width           =   195
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "From"
         Height          =   195
         Left            =   90
         TabIndex        =   11
         Top             =   210
         Width           =   345
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   780
      Left            =   2655
      Picture         =   "frmImmBatch.frx":0624
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6615
      Width           =   1290
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   780
      Left            =   8010
      Picture         =   "frmImmBatch.frx":092E
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6615
      Width           =   1110
   End
   Begin VB.CommandButton cmdNeg 
      Caption         =   "Negative"
      Height          =   600
      Index           =   3
      Left            =   10170
      Picture         =   "frmImmBatch.frx":0C38
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   810
      Width           =   1005
   End
   Begin VB.CommandButton cmdNeg 
      Caption         =   "Negative"
      Height          =   600
      Index           =   2
      Left            =   8775
      Picture         =   "frmImmBatch.frx":0F42
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   810
      Width           =   960
   End
   Begin VB.CommandButton cmdNeg 
      Caption         =   "Negative"
      Height          =   600
      Index           =   1
      Left            =   7290
      Picture         =   "frmImmBatch.frx":124C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   810
      Width           =   825
   End
   Begin VB.CommandButton cmdNeg 
      Caption         =   "Negative"
      Height          =   600
      Index           =   0
      Left            =   4275
      Picture         =   "frmImmBatch.frx":1556
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   810
      Width           =   915
   End
   Begin VB.TextBox tInput 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4260
      TabIndex        =   1
      Top             =   2220
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid gBio 
      Height          =   5130
      Left            =   135
      TabIndex        =   0
      ToolTipText     =   "AutoImmune Panel Results"
      Top             =   1395
      Width           =   11685
      _ExtentX        =   20611
      _ExtentY        =   9049
      _Version        =   393216
      Cols            =   6
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      Enabled         =   0   'False
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      FormatString    =   $"frmImmBatch.frx":1860
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
Attribute VB_Name = "frmImmBatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ResetPosition()

10        On Error GoTo ResetPosition_Error

20        tinput.Top = gBio.CellTop + gBio.Top
30        tinput.Left = gBio.CellLeft + gBio.Left
40        tinput.Width = gBio.ColWidth(gBio.Col)
50        tinput.Visible = True
60        tinput = gBio
70        tinput.SelStart = 0
80        tinput.SelLength = Len(tinput)
90        tinput.SetFocus

100       Exit Sub

ResetPosition_Error:

          Dim strES As String
          Dim intEL As Integer

110       intEL = Erl
120       strES = Err.Description
130       LogError "frmImmBatch", "ResetPosition", intEL, strES

End Sub

Private Sub SaveResult(ByVal SampleID As Long, _
                       ByVal Code As String, _
                       ByVal Result As String)

          Dim sql As String

10        On Error GoTo SaveResult_Error

20        sql = "IF EXISTS (SELECT * FROM ImmResults WHERE " & _
                "           SampleID = '" & SampleID & "' " & _
                "           AND Code = '" & Code & "') " & _
                "  UPDATE ImmResults " & _
                "  SET Result = '" & Result & "', " & _
                "  Valid = 0, " & _
                "  Printed = 0, " & _
                "  Operator = '" & UserCode & "', " & _
                "  SampleType = 'S', " & _
                "  RunDate = CONVERT(nvarchar, getdate(), 112), " & _
                "  RunTime = getdate() " & _
                "  WHERE SampleID = '" & SampleID & "' " & _
                "  AND Code = '" & Code & "' "
30        sql = sql & "ELSE " & _
                "  INSERT INTO ImmResults " & _
                "  (SampleID, Code, Result, Valid, Printed, Operator, " & _
                "   SampleType, RunDate, RunTime) " & _
                "   VALUES " & _
                "  ('" & SampleID & "', " & _
                "   '" & Code & "', " & _
                "   '" & Result & "', " & _
                "   0, 0, " & _
                "   '" & UserCode & "', " & _
                "   'S', " & _
                "   CONVERT(nvarchar, getdate(), 112), " & _
                "   getdate())"
40        Cnxn(0).Execute sql

50        sql = "DELETE FROM ImmRequests " & _
                "WHERE code = '" & Code & "' " & _
                "AND SampleID = '" & SampleID & "'"
60        Cnxn(0).Execute sql

70        Exit Sub

SaveResult_Error:

          Dim strES As String
          Dim intEL As Integer

80        intEL = Erl
90        strES = Err.Description
100       LogError "frmImmBatch", "SaveResult", intEL, strES, sql

End Sub

Private Sub cmdCancel_Click()

10        Unload Me

End Sub



Private Sub cmdCo_Click()

          Dim tb As Recordset
          Dim sn As Recordset
          Dim sql As String
          Dim ANA As String
          Dim ASMA As String
          Dim GPC As String
          Dim AMA As String
          Dim s As String

10        On Error GoTo cmdCo_Click_Error

20        With gBio
30            .Rows = 2
40            .AddItem ""
50            .RemoveItem 1
60        End With

70        If Trim(txtFrom) = "" Then Exit Sub
80        If Trim(txtTo) = "" Then Exit Sub

90        sql = "SELECT distinct(sampleid) from immrequests WHERE sampleid between " & _
                " " & txtFrom & " and " & txtTo & " and " & _
                "(code = 'ANA' or code = 'ASMA' or code = 'AMA' or code = 'GPC')"
100       Set tb = New Recordset
110       RecOpenServer 0, tb, sql
120       Do While Not tb.EOF
130           sql = "SELECT * FROM Demographics WHERE " & _
                    "SampleID = '" & tb!SampleID & "'"

140           Set sn = New Recordset
150           RecOpenServer 0, sn, sql
160           If Not sn.EOF Then
170               s = Trim(sn!PatName) & vbTab
180           Else
190               s = vbTab
200           End If
210           AMA = ""
220           ASMA = ""
230           GPC = ""
240           ANA = ""
250           sql = "SELECT * from immresults WHERE sampleid = " & tb!SampleID & ""
260           Set sn = New Recordset
270           RecOpenServer 0, sn, sql
280           Do While Not sn.EOF
290               If sn!Code = "AMA" Then AMA = sn!Result
300               If sn!Code = "GPC" Then GPC = sn!Result
310               If sn!Code = "ASMA" Then ASMA = sn!Result
320               If sn!Code = "ANA" Then ANA = sn!Result
330               sn.MoveNext
340           Loop
350           s = s & tb!SampleID & vbTab & ANA & vbTab & ASMA & vbTab & AMA & vbTab & GPC
360           gBio.AddItem s
370           s = ""
380           tb.MoveNext
390       Loop

400       If gBio.TextMatrix(1, 0) = "" And gBio.Rows > 2 Then
410           gBio.RemoveItem 1
420           gBio.Enabled = True
430       End If

440       Exit Sub

cmdCo_Click_Error:

          Dim strES As String
          Dim intEL As Integer



450       intEL = Erl
460       strES = Err.Description
470       LogError "frmImmBatch", "cmdCo_Click", intEL, strES, sql

End Sub

Private Sub cmdNeg_Click(Index As Integer)
          Dim n As Long

10        On Error GoTo cmdNeg_Click_Error

20        If gBio.Rows = 2 And gBio.TextMatrix(1, 0) = "" Then Exit Sub

30        For n = 1 To gBio.Rows - 1
40            gBio.TextMatrix(n, Index + 2) = "Negative"
50        Next

60        cmdSave.Enabled = True

70        Exit Sub

cmdNeg_Click_Error:

          Dim strES As String
          Dim intEL As Integer



80        intEL = Erl
90        strES = Err.Description
100       LogError "frmImmBatch", "cmdNeg_Click", intEL, strES


End Sub

Private Sub cmdSave_Click()

          Dim n As Long

10        On Error GoTo cmdSave_Click_Error

20        If gBio.Rows = 2 And gBio.TextMatrix(1, 1) = "" Then Exit Sub

30        For n = 1 To gBio.Rows - 1

40            If Trim(gBio.TextMatrix(n, 2)) <> "" Then
50                SaveResult gBio.TextMatrix(n, 1), "ANA", gBio.TextMatrix(n, 2)
60            End If

70            If Trim(gBio.TextMatrix(n, 3)) <> "" Then
80                SaveResult gBio.TextMatrix(n, 1), "ASMA", gBio.TextMatrix(n, 3)
90            End If

100           If Trim(gBio.TextMatrix(n, 4)) <> "" Then
110               SaveResult gBio.TextMatrix(n, 1), "AMA", gBio.TextMatrix(n, 4)
120           End If

130           If Trim(gBio.TextMatrix(n, 5)) <> "" Then
140               SaveResult gBio.TextMatrix(n, 1), "GPC", gBio.TextMatrix(n, 5)
150           End If
160       Next

170       Exit Sub

cmdSave_Click_Error:

          Dim strES As String
          Dim intEL As Integer

180       intEL = Erl
190       strES = Err.Description
200       LogError "frmImmBatch", "cmdsave_Click", intEL, strES

End Sub
Private Sub gBio_Click()

10        On Error GoTo gBio_Click_Error

20        If gBio.Col > 1 And gBio.MouseRow > 0 Then
30            tinput.Top = gBio.CellTop + gBio.Top
40            tinput.Left = gBio.CellLeft + gBio.Left
50            tinput.Width = gBio.ColWidth(gBio.Col)
60            tinput.Visible = True
70            tinput = gBio
80            tinput.SelStart = 0
90            tinput.SelLength = Len(tinput)
100           tinput.SetFocus
110           cmdSave.Enabled = True
120       End If

130       Exit Sub

gBio_Click_Error:

          Dim strES As String
          Dim intEL As Integer

140       intEL = Erl
150       strES = Err.Description
160       LogError "frmImmBatch", "gBio_Click", intEL, strES

End Sub

Private Sub tInput_KeyUp(KeyCode As Integer, Shift As Integer)

          Dim sql As String
          Dim tb As Recordset

10        On Error GoTo tInput_KeyUp_Error

20        Select Case KeyCode

          Case vbKeyF2:
30            sql = "SELECT * FROM Lists WHERE " & _
                    "ListType = 'IR' " & _
                    "AND Code = '" & tinput & "'"
40            Set tb = New Recordset
50            RecOpenServer 0, tb, sql
60            If Not tb.EOF Then
70                gBio = Trim(tb!Text)
80                tinput.Visible = False
90            End If
100           gBio.SetFocus

110       Case vbKeyDown:
120           gBio = tinput
130           If gBio.Row < gBio.Rows - 1 Then
140               gBio.Row = gBio.Row + 1
150               ResetPosition
160           End If

170       Case vbKeyUp:
180           gBio = tinput
190           If gBio.Row > 1 Then
200               gBio.Row = gBio.Row - 1
210               ResetPosition
220           End If

230       Case vbKeyLeft:
240           gBio = tinput
250           If gBio.Col > 2 Then
260               gBio.Col = gBio.Col - 1
270               ResetPosition
280           End If

290       Case vbKeyRight:
300           gBio = tinput
310           If gBio.Col < gBio.Cols - 1 Then
320               gBio.Col = gBio.Col + 1
330               ResetPosition
340           End If

350       Case 65 To 90, 48 To 57
360           gBio = tinput

370       End Select

380       Exit Sub

tInput_KeyUp_Error:

          Dim strES As String
          Dim intEL As Integer

390       intEL = Erl
400       strES = Err.Description
410       LogError "frmImmBatch", "tInput_KeyUp", intEL, strES, sql

End Sub

Private Sub txtFrom_LostFocus()

10        On Error GoTo txtFrom_LostFocus_Error

20        If Trim(txtFrom) = "" Then Exit Sub

30        If Not IsNumeric(txtFrom) Then
40            iMsg "Number not numeric!"
50            txtFrom = ""
60        End If

70        Exit Sub

txtFrom_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer



80        intEL = Erl
90        strES = Err.Description
100       LogError "frmImmBatch", "txtFrom_LostFocus", intEL, strES


End Sub


Private Sub txtTo_LostFocus()

10        On Error GoTo txtTo_LostFocus_Error

20        If Trim(txtTo) = "" Then Exit Sub

30        If Not IsNumeric(txtTo) Then
40            iMsg "Number not numeric!"
50            txtTo = ""
60        End If

70        Exit Sub

txtTo_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

80        intEL = Erl
90        strES = Err.Description
100       LogError "frmImmBatch", "txtTo_LostFocus", intEL, strES

End Sub


