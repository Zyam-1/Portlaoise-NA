VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmPrintPriorities 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NetAcquire --- Print Priorities"
   ClientHeight    =   7950
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8385
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   8385
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Height          =   315
      Left            =   6180
      TabIndex        =   12
      Top             =   1080
      Width           =   795
   End
   Begin VB.TextBox txtPrintPriority 
      Enabled         =   0   'False
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
      Left            =   4920
      TabIndex        =   11
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton bcancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   825
      Left            =   7380
      Picture         =   "frmPrintPriorities.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6945
      Width           =   795
   End
   Begin VB.CommandButton bMoveUp 
      Caption         =   "Move Up"
      Enabled         =   0   'False
      Height          =   825
      Left            =   7035
      Picture         =   "frmPrintPriorities.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3105
      Width           =   795
   End
   Begin VB.CommandButton bMoveDown 
      Caption         =   "Move Down"
      Enabled         =   0   'False
      Height          =   825
      Left            =   7035
      Picture         =   "frmPrintPriorities.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3960
      Width           =   795
   End
   Begin VB.CommandButton bSave 
      BackColor       =   &H0000FFFF&
      Caption         =   "Save"
      Height          =   825
      Left            =   7380
      Picture         =   "frmPrintPriorities.frx":0B8E
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6000
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Frame Frame1 
      Caption         =   "Department"
      Height          =   1200
      Left            =   240
      TabIndex        =   0
      Top             =   180
      Width           =   4245
      Begin VB.OptionButton o 
         Caption         =   "Biochemistry"
         Height          =   225
         Index           =   0
         Left            =   360
         TabIndex        =   4
         Top             =   390
         Value           =   -1  'True
         Width           =   1785
      End
      Begin VB.OptionButton o 
         Caption         =   "Coagulation"
         Enabled         =   0   'False
         Height          =   225
         Index           =   1
         Left            =   1260
         TabIndex        =   3
         Top             =   960
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.OptionButton o 
         Caption         =   "Endrocrinology"
         Height          =   225
         Index           =   2
         Left            =   2400
         TabIndex        =   2
         Top             =   390
         Width           =   1785
      End
      Begin VB.OptionButton o 
         Caption         =   "Immunology"
         Height          =   225
         Index           =   3
         Left            =   360
         TabIndex        =   1
         Top             =   720
         Width           =   1785
      End
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   6285
      Left            =   240
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1500
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   11086
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483630
      BackColorFixed  =   -2147483638
      ForeColorFixed  =   -2147483635
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      FormatString    =   "<Code   |Text                                                                                      "
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Print Priority"
      Height          =   195
      Left            =   4920
      TabIndex        =   10
      Top             =   840
      Width           =   825
   End
End
Attribute VB_Name = "frmPrintPriorities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private OldRowIndex As Integer
Private NewRowIndex As Integer
Private bDontEnterCell As Boolean

Private Enum Departments
    Biochemistry = 0
    Coagulation = 1
    Endrocrinology = 2
    Immunology = 3
End Enum

Private Sub LoadTestDefinitions(Department As Departments)

          Dim tb As Recordset
          Dim sql As String
          Dim s As String
          Dim Dept As String

10        On Error GoTo LoadTestDefinitions_Error

20        Select Case Department
          Case 0: Dept = "Bio"
30        Case 1: Dept = "Coag"
40        Case 2: Dept = "End"
50        Case 3: Dept = "Imm"
60        End Select


70        sql = "select distinct Code, longname, shortname,Analyser, PrintPriority " & _
                "from " & Dept & "testdefinitions " & _
                "order by PrintPriority"
80        Set tb = New Recordset
90        RecOpenClient 0, tb, sql
100       InitGrid
110       If Not tb.EOF Then
120           Screen.MousePointer = vbHourglass
130           With g
140               .Visible = False
150               While Not tb.EOF
160                   s = g.Rows & vbTab & _
                          tb!Code & "" & vbTab & _
                          tb!LongName & "" & vbTab & _
                          tb!ShortName & "" & vbTab & _
                          tb!PrintPriority & "" & vbTab & _
                          tb!Analyser & ""
170                   .AddItem s, .Rows

180                   tb.MoveNext
190               Wend

200               .Visible = True
210           End With
220           Screen.MousePointer = vbNormal
230       End If
240       Exit Sub

LoadTestDefinitions_Error:

          Dim strES As String
          Dim intEL As Integer

250       g.Visible = True
260       Screen.MousePointer = vbNormal
270       intEL = Erl
280       strES = Err.Description
290       LogError "frmPrintPriorities", "LoadTestDefinitions", intEL, strES, sql

End Sub

Private Sub InitGrid()

10        On Error GoTo InitGrid_Error

20        With g
30            .Rows = 2: .Cols = 6
40            .FixedRows = 1: .FixedCols = 1

50            .Rows = 1
60            .SelectionMode = flexSelectionByRow

70            .TextMatrix(0, 0) = "PP"
80            .TextMatrix(0, 1) = "Code"
90            .TextMatrix(0, 2) = "Long Name"
100           .TextMatrix(0, 3) = "Short Name"
110           .TextMatrix(0, 4) = "Print Priority"
120           .TextMatrix(0, 5) = "Analyser"

130           .ColWidth(0) = 400
140           .ColWidth(1) = 1000
150           .ColWidth(2) = 2300
160           .ColWidth(3) = 1300
170           .ColWidth(4) = 0
180           .ColWidth(5) = 1300

190           .ColAlignment(0) = flexAlignLeftCenter
200           .ColAlignment(1) = flexAlignLeftCenter
210           .ColAlignment(2) = flexAlignLeftCenter
220           .ColAlignment(3) = flexAlignLeftCenter
230           .ColAlignment(4) = flexAlignLeftCenter
240           .ColAlignment(5) = flexAlignLeftCenter

250       End With

260       Exit Sub

InitGrid_Error:

          Dim strES As String
          Dim intEL As Integer

270       intEL = Erl
280       strES = Err.Description
290       LogError "frmPrintPriorities", "InitGrid", intEL, strES

End Sub

Private Sub HighlightGridRow(RowIndex As Integer)
          Dim i As Integer
10        bDontEnterCell = True
20        g.Row = RowIndex
30        For i = 1 To g.Cols - 1
40            g.Col = i
50            g.CellBackColor = &H8000000D
60            g.CellForeColor = &H8000000E
70        Next i
80        bDontEnterCell = False
End Sub
Private Sub UnHighlightGridRow(RowIndex As Integer)
          Dim i As Integer
10        bDontEnterCell = True
20        g.Row = RowIndex
30        For i = 1 To g.Cols - 1
40            g.Col = i
50            g.CellBackColor = &H80000018
60            g.CellForeColor = &H80000012
70        Next i
80        bDontEnterCell = False
End Sub

Private Sub ReArrangeSerial()
          Dim i As Integer
          Dim o As Integer

10        o = g.Row
20        For i = 1 To g.Rows - 1
30            g.TextMatrix(i, 0) = i
40        Next i
50        g.Row = o
End Sub

Private Sub SavePriorities(Department As String)

          Dim sql As String
          Dim i As Integer

10        On Error GoTo SavePriorities_Error

20        Screen.MousePointer = vbHourglass
30        For i = 1 To g.Rows - 1
40            sql = "Update " & Department & "TestDefinitions " & _
                    "Set PrintPriority = " & g.TextMatrix(i, 0) & " " & _
                    "Where Code = '" & g.TextMatrix(i, 1) & "' "
50            sql = sql & "And IsNull(Analyser,'') = '" & Trim(g.TextMatrix(i, 5)) & "'"
60            Cnxn(0).Execute sql
70        Next i
80        Screen.MousePointer = vbNormal
90        bsave.Visible = False
100       iMsg "Print priorities saved sucessfully"

110       Exit Sub

SavePriorities_Error:

          Dim strES As String
          Dim intEL As Integer
120       Screen.MousePointer = vbNormal
130       intEL = Erl
140       strES = Err.Description
150       LogError "frmPrintPriorities", "SavePriorities", intEL, strES, sql

End Sub

Private Sub bcancel_Click()
10        Unload Me
End Sub

Private Sub bMoveDown_Click()
10        g.RowPosition(g.Row) = g.Row + 1
20        g.Row = g.Row + 1
30        ReArrangeSerial
40        bsave.Visible = True
End Sub

Private Sub bMoveUp_Click()
10        g.RowPosition(g.Row) = g.Row - 1
20        g.Row = g.Row - 1
30        ReArrangeSerial
40        bsave.Visible = True
End Sub

Private Sub bSave_Click()


10        On Error GoTo bSave_Click_Error

20        If o(0).Value = True Then
30            SavePriorities "Bio"
40        ElseIf o(2).Value = True Then
50            SavePriorities "End"
60        ElseIf o(3).Value = True Then
70            SavePriorities "Imm"
80        End If
90        Exit Sub

bSave_Click_Error:

          Dim strES As String
          Dim intEL As Integer
100       Screen.MousePointer = vbNormal
110       intEL = Erl
120       strES = Err.Description
130       LogError "frmPrintPriorities", "bSave_Click", intEL, strES

End Sub

Private Sub cmdUpdate_Click()
          Dim NewIndex As Integer
10        If g.Row = txtPrintPriority Then Exit Sub
20        If txtPrintPriority > g.Rows Then
30            iMsg "Print priority should be in between 1 and " & g.Rows
40            txtPrintPriority.SetFocus
50            Exit Sub
60        End If
70        NewIndex = txtPrintPriority
80        g.RowPosition(g.Row) = NewIndex
90        g.Row = NewIndex
100       ReArrangeSerial
110       bsave.Visible = True
End Sub

Private Sub Form_Load()
10        bDontEnterCell = False
20        LoadTestDefinitions Biochemistry
End Sub




Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
10        If bsave.Visible = True Then
20            If iMsg("Changes are not saved. Are you sure you wish to exit without saving", vbYesNo) = vbNo Then
30                Cancel = True
40            End If
50        End If
End Sub

Private Sub g_EnterCell()

10        If bDontEnterCell = False Then

20            NewRowIndex = g.Row
30            txtPrintPriority = g.Row
40            bMoveUp.Enabled = (g.Row > 1)
50            bMoveDown.Enabled = (g.Row < g.Rows - 1)
60            txtPrintPriority.Enabled = g.Row > 0
70            UnHighlightGridRow OldRowIndex
80            HighlightGridRow NewRowIndex

90        End If

End Sub

Private Sub g_KeyDown(KeyCode As Integer, Shift As Integer)
10        Debug.Print KeyCode
End Sub

Private Sub g_LeaveCell()
10        OldRowIndex = g.Row
End Sub



Private Sub o_Click(Index As Integer)
10        Select Case Index
          Case 0: LoadTestDefinitions Biochemistry
20        Case 1: LoadTestDefinitions Coagulation
30        Case 2: LoadTestDefinitions Endrocrinology
40        Case 3: LoadTestDefinitions Immunology
50        End Select
End Sub

Private Sub txtPrintPriority_KeyPress(KeyAscii As Integer)
10        KeyAscii = VI(KeyAscii, Numeric_Only)
End Sub
