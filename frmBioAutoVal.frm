VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmBioAutoVal 
   Caption         =   "NetAcquire - Biochemistry AutoValidate Ranges"
   ClientHeight    =   7305
   ClientLeft      =   2805
   ClientTop       =   765
   ClientWidth     =   5745
   Icon            =   "frmBioAutoVal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   5745
   Begin VB.ComboBox cmbSampleType 
      Height          =   315
      Left            =   4320
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   780
      Width           =   1305
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   720
      Left            =   4320
      Picture         =   "frmBioAutoVal.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5625
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   675
      Left            =   4320
      Picture         =   "frmBioAutoVal.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6435
      Width           =   1035
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   225
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSFlexGridLib.MSFlexGrid grdVal 
      Height          =   6855
      Left            =   120
      TabIndex        =   3
      Top             =   300
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   12091
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   2
      ScrollBars      =   2
      FormatString    =   "<Parameter               |^Low       |<High      "
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
Attribute VB_Name = "frmBioAutoVal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Disp As String

Private Sub cmdCancel_Click()

10        Unload Me

End Sub

Private Sub cmdPrint_Click()

          Dim Y As Long

10        On Error GoTo cmdPrint_Click_Error

20        Printer.Print
30        Printer.Font.Name = "Courier New"
40        Printer.Font.Size = 12

50        If Disp = "End" Then
60            Printer.Print "List of Endocrinology AutoValidate Ranges."
70        ElseIf Disp = "Imm" Then
80            Printer.Print "List of Immunology AutoValidate Ranges."
90        Else
100           Printer.Print "List of Biochemistry AutoValidate Ranges."
110       End If
120       Printer.Print

130       For Y = 0 To grdVal.Rows - 1
140           grdVal.Row = Y
150           grdVal.Col = 0
160           Printer.Print grdVal; Tab(20);
170           grdVal.Col = 1
180           Printer.Print grdVal; Tab(30);
190           grdVal.Col = 2
200           Printer.Print grdVal
210       Next

220       Printer.EndDoc

230       Exit Sub

cmdPrint_Click_Error:

          Dim strES As String
          Dim intEL As Integer

240       intEL = Erl
250       strES = Err.Description
260       LogError "frmBioAutoVal", "cmdPrint_Click", intEL, strES

End Sub

Private Sub cmbSampleType_Click()

10        On Error GoTo cmbSampleType_Click_Error

20        FillG

30        Exit Sub

cmbSampleType_Click_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmBioAutoVal", "cmbSampleType_Click", intEL, strES


End Sub

Private Sub FillG()

          Dim tb As New Recordset
          Dim sql As String



10        On Error GoTo FillG_Error

20        ClearFGrid grdVal

30        sql = "SELECT distinct ShortName, PrintPriority, " & _
                "AutoValLow, AutoValHigh " & _
                "from " & Disp & "TestDefinitions " & _
                "Order by PrintPriority"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        Do While Not tb.EOF
70            grdVal.AddItem tb!ShortName & vbTab & _
                             tb!AutoValLow & vbTab & _
                             tb!AutoValHigh & ""
80            tb.MoveNext
90        Loop

100       FixG grdVal




110       Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer

120       intEL = Erl
130       strES = Err.Description
140       LogError "frmBioAutoVal", "FillG", intEL, strES, sql


End Sub

Private Sub FillSampleType()

          Dim sql As String
          Dim tb As New Recordset



10        On Error GoTo FillSampleType_Error

20        cmbSampleType.Clear

30        sql = "SELECT * from lists WHERE listtype = 'ST'"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        Do While Not tb.EOF
70            cmbSampleType.AddItem Trim(tb!Text)
80            tb.MoveNext
90        Loop

100       If cmbSampleType.ListCount > 0 Then
110           cmbSampleType.ListIndex = 0
120       End If




130       Exit Sub

FillSampleType_Error:

          Dim strES As String
          Dim intEL As Integer

140       intEL = Erl
150       strES = Err.Description
160       LogError "frmBioAutoVal", "FillSampleType", intEL, strES, sql


End Sub

Private Sub Form_Load()

10        On Error GoTo Form_Load_Error

20        grdVal.Font.Bold = True

30        FillSampleType
40        FillG

50        Set_Font Me


60        Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

70        intEL = Erl
80        strES = Err.Description
90        LogError "frmBioAutoVal", "Form_Load", intEL, strES


End Sub

Private Sub grdVal_Click()

          Static SortOrder As Boolean



10        On Error GoTo grdVal_Click_Error

20        If grdVal.MouseRow = 0 Then
30            If SortOrder Then
40                grdVal.Sort = flexSortGenericAscending
50            Else
60                grdVal.Sort = flexSortGenericDescending
70            End If
80            SortOrder = Not SortOrder
90            Exit Sub
100       End If

110       If grdVal.Col = 0 Then
120           Exit Sub
130       ElseIf grdVal.Col = 1 Then
140           grdVal.Enabled = False
150           grdVal = iBOX("AutoValidate Low?", , grdVal)
160           SaveG
170           grdVal.Enabled = True
180       Else
190           grdVal.Enabled = False
200           grdVal = iBOX("AutoValidate High?", , grdVal)
210           SaveG
220           grdVal.Enabled = True
230       End If




240       Exit Sub

grdVal_Click_Error:

          Dim strES As String
          Dim intEL As Integer

250       intEL = Erl
260       strES = Err.Description
270       LogError "frmBioAutoVal", "grdVal_Click", intEL, strES


End Sub

Private Sub SaveG()

          Dim sql As String
          Dim DoDelta As Boolean



10        On Error GoTo SaveG_Error

20        DoDelta = grdVal.TextMatrix(grdVal.Row, 1) = "Yes"

30        sql = "UPDATE " & Disp & "TestDefinitions " & _
                "Set AutoValLow = '" & Val(grdVal.TextMatrix(grdVal.Row, 1)) & "', " & _
                "AutoValHigh = '" & Val(grdVal.TextMatrix(grdVal.Row, 2)) & "' " & _
                "WHERE ShortName = '" & grdVal.TextMatrix(grdVal.Row, 0) & "'"
40        Cnxn(0).Execute sql



50        Exit Sub

SaveG_Error:

          Dim strES As String
          Dim intEL As Integer

60        intEL = Erl
70        strES = Err.Description
80        LogError "frmBioAutoVal", "SaveG", intEL, strES, sql


End Sub

