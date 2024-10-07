VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmTestCodeMapping 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NetAcquire --- Biochemistry Test Code Mapping"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9645
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   9645
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtText 
      Height          =   315
      Left            =   8640
      TabIndex        =   8
      Top             =   1680
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   1100
      Left            =   8220
      Picture         =   "frmTestCodeMapping.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4440
      Width           =   1200
   End
   Begin VB.CommandButton cmdSave 
      Appearance      =   0  'Flat
      Caption         =   "&Save"
      Height          =   1100
      Left            =   8220
      Picture         =   "frmTestCodeMapping.frx":049C
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3240
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.ComboBox cmbAnalyser 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   480
      Width           =   2835
   End
   Begin VB.ComboBox cmbTestName 
      Height          =   315
      Left            =   2400
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   6000
      Width           =   1215
   End
   Begin VB.ComboBox cmbTestCode 
      Height          =   315
      Left            =   1260
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   6060
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   4575
      Left            =   180
      TabIndex        =   0
      Top             =   960
      Width           =   7875
      _ExtentX        =   13891
      _ExtentY        =   8070
      _Version        =   393216
      RowHeightMin    =   315
      ScrollTrack     =   -1  'True
   End
   Begin MSComctlLib.ProgressBar PBar 
      Height          =   195
      Left            =   270
      TabIndex        =   1
      Top             =   120
      Width           =   7755
      _ExtentX        =   13679
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label lblSelectAnalyser 
      AutoSize        =   -1  'True
      Caption         =   "Select Analyser"
      Height          =   195
      Left            =   180
      TabIndex        =   5
      Top             =   540
      Width           =   1095
   End
End
Attribute VB_Name = "frmTestCodeMapping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private PrevAnalyser As String
Private grd As MSFlexGrid
Private m_sDiscipline As String
Private Department As String

Public Property Get Discipline() As String

10        Discipline = m_sDiscipline

End Property

Public Property Let Discipline(ByVal sDiscipline As String)

10        m_sDiscipline = sDiscipline

End Property


Private Sub InitializeGrid()
          Dim i As Integer
10        With g
20            .Rows = 2: .FixedRows = 1
30            .Cols = 4: .FixedCols = 0
40            .Rows = 1
              '.Font.Size = 10         'fgcFontSize
              '.Font.Name = fgcFontName
              '.ForeColor = fgcForeColor
              '.BackColor = fgcBackColor
              '.ForeColorFixed = fgcForeColorFixed
              '.BackColorFixed = fgcBackColorFixed
50            .ScrollBars = flexScrollBarBoth
              'Name                                                                      |Code
60            .TextMatrix(0, 0) = "Code": .ColWidth(0) = 1500: .ColAlignment(0) = flexAlignLeftCenter
70            .TextMatrix(0, 1) = "Test Name": .ColWidth(1) = 3000: .ColAlignment(1) = flexAlignLeftCenter
80            .TextMatrix(0, 2) = "Short Name": .ColWidth(2) = 1500: .ColAlignment(2) = flexAlignLeftCenter
90            .TextMatrix(0, 3) = "Analyser Code": .ColWidth(3) = 1500: .ColAlignment(3) = flexAlignLeftCenter
100           For i = 0 To .Cols - 1
110               If .ColWidth(i) < Me.TextWidth(.TextMatrix(0, i)) Then
120                   .ColWidth(i) = Me.TextWidth(.TextMatrix(0, i)) + 150    'fgcExtraSpace
130               End If
140           Next i
150       End With
End Sub

Private Sub FillGrid()

          Dim tb As Recordset
          Dim sql As String
          Dim s As String

10        On Error GoTo FillGrid_Error

20        sql = "Select Distinct D.Code, D.LongName, D.ShortName, M.EquipmentAnalyserCode " & _
                "From " & Discipline & "TestDefinitions D Left Outer Join " & _
                "(Select * From AnalyserTestCodeMapping Where AnalyserName = '" & cmbAnalyser & "' AND Department = '" & Discipline & "') M " & _
                "On D.Code = M.NetAcquireTestCode " & _
                "Order By M.EquipmentAnalyserCode"
30        Set tb = New Recordset
40        RecOpenClient 0, tb, sql
50        cmdSave.Visible = False
60        InitializeGrid
70        If Not tb.EOF Then
80            While Not tb.EOF
90                s = tb!Code & "" & vbTab & _
                      tb!LongName & "" & vbTab & _
                      tb!ShortName & "" & vbTab & _
                      tb!EquipmentAnalyserCode
100               g.AddItem s
110               tb.MoveNext
120           Wend

130       End If
140       g.TextMatrix(0, 3) = cmbAnalyser & " Code"

150       Exit Sub

FillGrid_Error:

          Dim strES As String
          Dim intEL As Integer

160       intEL = Erl
170       strES = Err.Description
180       LogError "frmTestCodeMapping", "FillGrid", intEL, strES

End Sub

Private Sub cmbAnalyser_Click()

10        If cmdSave.Visible = True Then
20            If iMsg("Do you want to save changes?", vbQuestion + vbYesNo) = vbYes Then
30                Save PrevAnalyser
40            End If
50        End If

60        If cmbAnalyser <> "" Then
70            FillGrid
80        End If
End Sub

Private Sub cmbAnalyser_DropDown()
10        PrevAnalyser = cmbAnalyser
End Sub

Private Sub cmbTestCode_Click()
          Dim Code As String
10        Code = cmbTestCode.Text
20        g.TextMatrix(g.row, g.Col) = Code
30        cmbTestCode.Visible = False
40        g.TextMatrix(g.row, 1) = LongNamebyCode(Code, "Bio")
End Sub

Private Sub cmbTestName_Click()
          Dim LongName As String
10        LongName = cmbTestName.Text
20        g.TextMatrix(g.row, g.Col) = LongName
30        cmbTestName.Visible = False
40        g.TextMatrix(g.row, 0) = CodebyLongName(LongName, "Bio")
End Sub

Private Sub cmdCancel_Click()
10        Unload Me
End Sub

Private Sub Save(AnalyserName As String)

          Dim tb As Recordset
          Dim sql As String
          Dim i As Integer

10        On Error GoTo Save_Error

20        For i = 1 To g.Rows - 1
30            If g.TextMatrix(i, 3) = "" Then
40                sql = "Delete From AnalyserTestCodeMapping " & _
                        "Where AnalyserName = '" & AnalyserName & "' " & _
                        "And NetAcquireTestCode = '" & g.TextMatrix(i, 0) & "' " & _
                        "AND Department = '" & Discipline & "'"
50                Cnxn(0).Execute sql
60            Else
70                sql = "Select * From AnalyserTestCodeMapping " & _
                        "Where AnalyserName = '" & AnalyserName & "' " & _
                        "And NetAcquireTestCode = '" & g.TextMatrix(i, 0) & "' " & _
                        "AND Department = '" & Discipline & "'"

80                Set tb = New Recordset
90                RecOpenClient 0, tb, sql
100               If tb.EOF Then
110                   tb.AddNew
120               End If

130               tb!NetAcquireTestCode = g.TextMatrix(i, 0)
140               tb!EquipmentAnalyserCode = g.TextMatrix(i, 3)
150               tb!TestName = g.TextMatrix(i, 1)
160               tb!AnalyserName = AnalyserName
170               tb!Department = Discipline
180               tb!DateTimeOfRecord = Format(Now, "YYYY-MM-DD hh:mm:ss")

190               tb.Update
200           End If

210       Next i

220       Exit Sub

Save_Error:

          Dim strES As String
          Dim intEL As Integer

230       intEL = Erl
240       strES = Err.Description
250       LogError "frmTestCodeMapping", "Save", intEL, strES, sql

End Sub

Private Sub cmdSave_Click()



10        On Error GoTo cmdSave_Click_Error

20        Save cmbAnalyser
30        FillGrid
40        cmdSave.Visible = False
50        Exit Sub

cmdSave_Click_Error:

          Dim strES As String
          Dim intEL As Integer

60        intEL = Erl
70        strES = Err.Description
80        LogError "frmTestCodeMapping", "cmdSave_Click", intEL, strES

End Sub

Private Sub Form_Load()

10        CheckAnalyserTestCodeMappingInDb

20        InitializeGrid
30        Me.Caption = "NetAcquire --- " & Discipline & " Test Code Mapping"
          'Fill all combos

'40        Select Case Discipline
'          Case "Biochemistry":
'50            Department = "Bio"
'60        End Select

70        FillGenericList cmbAnalyser, Discipline & "Analysers"

End Sub

'Private Sub FillTestCombos()
'
'          Dim tb As Recordset
'          Dim sql As String
'
'10        On Error GoTo FillTestCombos_Error
'
'20        sql = "Select Distinct Code, LongName From " & Discipline & "TestDefinitions"
'30        Set tb = New Recordset
'40        RecOpenClient 0, tb, sql
'
'50        cmbTestCode.Clear
'
'60        cmbTestName.Clear
'70        If Not tb.EOF Then
'80            While Not tb.EOF
'90                cmbTestCode.AddItem tb!Code & ""
'100               cmbTestName.AddItem tb!LongName & ""
'110               tb.MoveNext
'
'120           Wend
'130       End If
'
'140       Exit Sub
'
'FillTestCombos_Error:
'
'          Dim strES As String
'          Dim intEL As Integer
'
'150       intEL = Erl
'160       strES = Err.Description
'170       LogError "frmTestCodeMapping", "FillTestCombos", intEL, strES, sql
'
'End Sub





Private Function LongNamebyCode(ByVal Code As String, _
                                ByVal Department As String) As String

          Dim tb As New Recordset
          Dim sql As String

10        On Error GoTo LongNamebyCode_Error

20        LongNamebyCode = "???"

30        sql = "SELECT LongName FROM " & Department & "TestDefinitions WHERE " & _
                "Code = '" & Code & "'"

40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        If Not tb.EOF Then
70            LongNamebyCode = Trim(tb!LongName & "")
80        End If

90        Exit Function

LongNamebyCode_Error:

          Dim strES As String
          Dim intEL As Integer

100       intEL = Erl
110       strES = Err.Description
120       LogError "basDisciplineFunctions", "LongNamebyCode", intEL, strES, sql

End Function

Private Function CodebyLongName(ByVal LongName As String, _
                                ByVal Department As String) As String

          Dim sql As String
          Dim tb As New Recordset

10        On Error GoTo CodebyLongName_Error

20        CodebyLongName = "???"

30        sql = "SELECT Code FROM " & Department & "TestDefinitions WHERE " & _
                "LongName = '" & LongName & "'"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        If Not tb.EOF Then
70            CodebyLongName = Trim(tb!Code & "")
80        End If

90        Exit Function

CodebyLongName_Error:

          Dim strES As String
          Dim intEL As Integer

100       intEL = Erl
110       strES = Err.Description
120       LogError "basDisciplineFunctions", "CodebyLongName", intEL, strES, sql

End Function




Private Sub g_Click()

          Static SortOrder As Boolean


10        If g.MouseRow = 0 Then
20            If SortOrder Then
30                g.Sort = flexSortGenericAscending
40            Else
50                g.Sort = flexSortGenericDescending
60            End If
70            SortOrder = Not SortOrder
80            Exit Sub
90        End If

100       If g.ColSel = 3 Then
110           If g.MouseRow > 0 Then
120               Set grd = g
130               grd.row = grd.MouseRow
140               grd.Col = grd.MouseCol
150               LoadControls
160           End If
170           Exit Sub
180       End If

End Sub

Private Sub g_KeyUp(KeyCode As Integer, Shift As Integer)
10        If g.Col = 3 Then
20            If EditGrid(g, KeyCode, Shift) Then
30                cmdSave.Visible = True
40            End If
50        End If
End Sub

Private Sub g_LeaveCell()

10        On Error GoTo g_LeaveCell_Error

20        txtText.Visible = False

30        Exit Sub

g_LeaveCell_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmTestCodeMapping", "g_LeaveCell", intEL, strES

End Sub

Private Sub g_Scroll()

10        On Error GoTo g_Scroll_Error

20        txtText.Visible = False

30        Exit Sub

g_Scroll_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmTestCodeMapping", "g_Scroll", intEL, strES

End Sub

Private Sub txtText_LostFocus()

10        On Error GoTo txtText_LostFocus_Error



20        txtText.Visible = False

30        Exit Sub

txtText_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmTestCodeMapping", "txtText_LostFocus", intEL, strES

End Sub

Private Sub LoadControls()
10        On Error GoTo LoadControls_Error

20        txtText.Visible = False
30        txtText = ""
          'gRD.SetFocus

40        Select Case grd.Col
          Case 3:
50            txtText.Move grd.Left + grd.CellLeft + 5, _
                           grd.Top + grd.CellTop + 5, _
                           grd.CellWidth - 20, grd.CellHeight - 20
60            txtText.Text = grd.TextMatrix(grd.row, grd.Col)
70            txtText.Visible = True
80            txtText.SelStart = 0
90            txtText.SelLength = Len(txtText)
100           txtText.SetFocus

110       End Select

120       Exit Sub

LoadControls_Error:

          Dim strES As String
          Dim intEL As Integer

130       intEL = Erl
140       strES = Err.Description
150       LogError "frmOptions", "LoadControls", intEL, strES

End Sub

Private Sub txtText_KeyUp(KeyCode As Integer, Shift As Integer)
10        If KeyCode = vbKeyUp Then
              'GoOneRowUp
20        ElseIf KeyCode = vbKeyDown Then
              'GoOneRowDown
30        ElseIf KeyCode = 13 Then
40            txtText.Visible = False
50        Else
60            grd.TextMatrix(grd.row, grd.Col) = txtText
70            cmdSave.Visible = True
80        End If
End Sub

Private Sub GoOneRowUp()
10        If grd.row > 1 Then
20            grd.row = grd.row - 1
30            LoadControls
40        End If
End Sub
Private Sub GoOneRowDown()
10        If grd.row < grd.Rows - 1 Then
20            grd.row = grd.row + 1
30            LoadControls
40        End If
End Sub



