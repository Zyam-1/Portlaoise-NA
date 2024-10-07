VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmIsolateReport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Isolate Report"
   ClientHeight    =   9195
   ClientLeft      =   105
   ClientTop       =   495
   ClientWidth     =   15375
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9195
   ScaleWidth      =   15375
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton bcancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   735
      Left            =   5670
      Picture         =   "frmIsolateReport.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   240
      Width           =   1155
   End
   Begin VB.CommandButton bGen 
      Caption         =   "&Start"
      Height          =   735
      Left            =   1950
      MaskColor       =   &H8000000F&
      Picture         =   "frmIsolateReport.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   240
      Width           =   1155
   End
   Begin VB.CommandButton cmdPrint 
      Appearance      =   0  'Flat
      Caption         =   "&Print to Local Printer"
      Height          =   735
      Left            =   3660
      Picture         =   "frmIsolateReport.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   240
      Width           =   1635
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   8010
      Left            =   45
      TabIndex        =   0
      Top             =   1170
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   14129
      _Version        =   393216
      Cols            =   18
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      AllowUserResizing=   1
      FormatString    =   $"frmIsolateReport.frx":091E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComCtl2.DTPicker calToDate 
      Height          =   315
      Left            =   510
      TabIndex        =   6
      Top             =   630
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   556
      _Version        =   393216
      Format          =   59310081
      CurrentDate     =   38503
   End
   Begin MSComCtl2.DTPicker calFromDate 
      Height          =   315
      Left            =   510
      TabIndex        =   7
      Top             =   240
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   556
      _Version        =   393216
      Format          =   59310081
      CurrentDate     =   38503
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Between Dates"
      Height          =   195
      Left            =   540
      TabIndex        =   8
      Top             =   30
      Width           =   1095
   End
   Begin VB.Label lblTotal 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   12210
      TabIndex        =   2
      Top             =   420
      Width           =   1140
   End
   Begin VB.Label Label1 
      Caption         =   "Total"
      Height          =   240
      Left            =   11715
      TabIndex        =   1
      Top             =   465
      Width           =   690
   End
End
Attribute VB_Name = "frmIsolateReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit





Function HowManyWeeks() As Integer


10        On Error GoTo HowManyWeeks_Error

20        If g.Rows < 2 Then
30            HowManyWeeks = 0
40            Exit Function
50        End If

60        HowManyWeeks = Abs(DateDiff("w", Format(calFromDate, "dd/mmm/yyyy"), Format(calToDate, "dd/mmm/yyyy"))) + 1

70        Exit Function

HowManyWeeks_Error:

          Dim strES As String
          Dim intEL As Integer



80        intEL = Erl
90        strES = Err.Description
100       LogError "frmIsolateReport", "HowManyWeeks", intEL, strES


End Function

Private Sub bcancel_Click()

10        Unload Me

End Sub


Private Sub bGen_Click()

10        FillG

End Sub

Private Sub cmdPrint_Click()

          Dim n As Integer

10        On Error GoTo cmdPrint_Click_Error

20        Printer.Orientation = vbPRORLandscape
30        Printer.Font.Name = "Courier New"
40        Printer.Font.Size = 16
50        Printer.Print "                   Isolate Report"
60        Printer.Font.Size = 14
70        Printer.Print "Between Dates " & calFromDate & " and " & calToDate
80        Printer.Font.Size = 8

90        Printer.Print "Date     ";    '9
100       Printer.Print "Sample ID ";    '10
110       Printer.Print "Chart    ";    '9
120       Printer.Print "Name            ";    '16
130       Printer.Print "Dob      ";    '9
140       Printer.Print "Address         ";    '16
150       Printer.Print "Ward       ";    '11
160       Printer.Print "Consultant ";    '11
170       Printer.Print "GP         ";    '11
180       Printer.Print "VRE ESBL MRSA Salm Shig Camp E.coli0* EPEC C.diff"
190       For n = 1 To g.Rows - 1
200           Printer.Print Format(g.TextMatrix(n, 0), "dd/MM/yy") & " ";
210           Printer.Print Left$(g.TextMatrix(n, 1) & Space$(9), 9) & " ";
220           Printer.Print Left$(g.TextMatrix(n, 2) & Space$(8), 8) & " ";
230           Printer.Print Left$(g.TextMatrix(n, 3) & Space$(15), 15) & " ";
240           Printer.Print Left$(g.TextMatrix(n, 4) & Space$(8), 8) & " ";
250           Printer.Print Left$(g.TextMatrix(n, 5) & Space$(15), 15) & " ";
260           Printer.Print Left$(g.TextMatrix(n, 6) & Space$(10), 10) & " ";
270           Printer.Print Left$(g.TextMatrix(n, 7) & Space$(10), 10) & " ";
280           Printer.Print Left$(g.TextMatrix(n, 8) & Space$(10), 10) & " ";
290           Printer.Print Left$(g.TextMatrix(n, 9) & Space$(4), 4);    'VRE
300           Printer.Print Left$(g.TextMatrix(n, 10) & Space$(5), 5);    'ESBL
310           Printer.Print Left$(g.TextMatrix(n, 11) & Space$(5), 5);    'MRSA
320           Printer.Print Left$(g.TextMatrix(n, 12) & Space$(5), 5);    'Salm
330           Printer.Print Left$(g.TextMatrix(n, 13) & Space$(5), 5);    'Shig
340           Printer.Print Left$(g.TextMatrix(n, 14) & Space$(5), 5);    'Camp
350           Printer.Print Left$(g.TextMatrix(n, 15) & Space$(9), 9);    'E.coli0*
360           Printer.Print Left$(g.TextMatrix(n, 16) & Space$(5), 5);    'EPEC
370           Printer.Print g.TextMatrix(n, 17)    ' C.diff
380       Next
390       Printer.EndDoc

400       Exit Sub

cmdPrint_Click_Error:

          Dim strES As String
          Dim intEL As Integer

410       intEL = Erl
420       strES = Err.Description
430       LogError "frmIsolateReport", "cmdPrint_Click", intEL, strES

End Sub

Private Sub calfromdate_CloseUp()

10        bGen.Visible = True

End Sub


Private Sub Form_Activate()

10        FillG

End Sub

Private Sub Form_Load()

10        On Error GoTo Form_Load_Error

20        calToDate = Format(Now, "dd/mmm/yyyy")
30        calFromDate = Format(DateAdd("m", -1, Format(Now, "dd/mmm/yyyy")))

40        Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmIsolateReport", "Form_Load", intEL, strES

End Sub

Sub FillG()

          Dim sql As String
          Dim tb As Recordset
          Dim s As String
          Dim dTo As String
          Dim dFrom As String
          Dim n As Integer
          Dim Found As Boolean

10        On Error GoTo FillG_Error

20        Screen.MousePointer = vbHourglass

30        g.Rows = 2
40        g.AddItem ""
50        g.RemoveItem 1

60        dTo = Format(calToDate, "dd/MMM/yyyy")
70        dFrom = Format(calFromDate, "dd/MMM/yyyy")

80        sql = "SELECT D.RunDate, D.SampleID, D.Chart, D.PatName, D.DoB, " & _
                "D.Addr0, D.Ward, D.Clinician, D.GP, I.OrganismName " & _
                "FROM Demographics D, Isolates I WHERE D.SampleID = I.SampleID " & _
                "AND D.RunDate BETWEEN '" & dFrom & "' AND '" & dTo & "' " & _
                "AND ( I.OrganismName LIKE 'VRE%' " & _
                "   OR I.OrganismName LIKE 'ESBL%' " & _
                "   OR I.OrganismName LIKE 'Staphylococcus aureus (MRSA)%' " & _
                "   OR I.OrganismName LIKE 'Salmonella%' " & _
                "   OR I.OrganismName LIKE 'Shigella%' " & _
                "   OR I.OrganismName LIKE 'Campylobacter%' " & _
                "   OR I.OrganismName LIKE 'E.coli0%' " & _
                "   OR I.OrganismName LIKE 'EPEC%' " & _
                "   OR I.OrganismName LIKE 'C.diff%' ) " & _
                "ORDER BY D.RunDate"

90        Set tb = New Recordset
100       RecOpenServer 0, tb, sql
110       Do While Not tb.EOF
120           s = Format(tb!Rundate, "dd/mm/yyyy") & vbTab & _
                  Val(tb!SampleID) - SysOptMicroOffset(0) & vbTab & _
                  tb!Chart & vbTab & _
                  tb!PatName & vbTab
130           If Not IsNull(tb!Dob) Then
140               s = s & Format(tb!Dob, "dd/mm/yyyy")
150           End If
160           s = s & vbTab & _
                  tb!Addr0 & vbTab & _
                  tb!Ward & vbTab & _
                  tb!Clinician & "" & vbTab & _
                  tb!GP & ""
170           g.AddItem s

180           g.Row = g.Rows - 1

190           If Left$(UCase$(tb!OrganismName & ""), 3) = "VRE" Then
200               g.TextMatrix(g.Row, 9) = "X"
210           ElseIf Left$(UCase$(tb!OrganismName & ""), 4) = "ESBL" Then
220               g.TextMatrix(g.Row, 10) = "X"
230           ElseIf Left$(UCase$(tb!OrganismName & ""), 5) = "STAPH" Then
240               g.TextMatrix(g.Row, 11) = "X"
250           ElseIf Left$(UCase$(tb!OrganismName & ""), 4) = "SALM" Then
260               g.TextMatrix(g.Row, 12) = "X"
270           ElseIf Left$(UCase$(tb!OrganismName & ""), 4) = "SHIG" Then
280               g.TextMatrix(g.Row, 13) = "X"
290           ElseIf Left$(UCase$(tb!OrganismName & ""), 4) = "CAMP" Then
300               g.TextMatrix(g.Row, 14) = "X"
310           ElseIf Left$(UCase$(tb!OrganismName & ""), 7) = "E.COLI0" Then
320               g.TextMatrix(g.Row, 15) = "X"
330           ElseIf Left$(UCase$(tb!OrganismName & ""), 10) = "EPEC" Then
340               g.TextMatrix(g.Row, 16) = "X"
350           ElseIf Left$(UCase$(tb!OrganismName & ""), 6) = "C.DIFF" Then
360               g.TextMatrix(g.Row, 17) = "X"
370           End If
380           tb.MoveNext
390       Loop

400       sql = "SELECT D.RunDate, D.SampleID, D.Chart, D.PatName, D.DoB, " & _
                "D.Addr0, D.Ward, D.Clinician, D.GP " & _
                "FROM Demographics D, Faeces F WHERE D.SampleID = F.SampleID " & _
                "AND D.RunDate BETWEEN '" & dFrom & "' AND '" & dTo & "' " & _
                "AND F.ToxinAB = 'P'"
410       Set tb = New Recordset
420       RecOpenServer 0, tb, sql
430       Do While Not tb.EOF
440           Found = False
450           For n = 1 To g.Rows - 1
460               If g.TextMatrix(n, 1) = Format$(Val(tb!SampleID) - SysOptMicroOffset(0)) Then
470                   g.TextMatrix(n, 17) = "X"
480                   Found = True
490                   Exit For
500               End If
510           Next
520           If Not Found Then
530               s = Format(tb!Rundate, "dd/mm/yyyy") & vbTab & _
                      Val(tb!SampleID) - SysOptMicroOffset(0) & vbTab & _
                      tb!Chart & vbTab & _
                      tb!PatName & vbTab
540               If Not IsNull(tb!Dob) Then
550                   s = s & Format(tb!Dob, "dd/mm/yyyy")
560               End If
570               s = s & vbTab & _
                      tb!Addr0 & vbTab & _
                      tb!Ward & vbTab & _
                      tb!Clinician & "" & vbTab & _
                      tb!GP & ""
580               g.AddItem s
590               g.Row = g.Rows - 1
600               g.TextMatrix(n, 17) = "X"
610           End If
620           tb.MoveNext
630       Loop

640       If g.Rows > 2 Then
650           g.RemoveItem 1
660       End If
670       g.Visible = True

680       Screen.MousePointer = vbNormal

690       Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer

700       intEL = Erl
710       strES = Err.Description
720       LogError "frmIsolateReport", "FillG", intEL, strES, sql
730       Screen.MousePointer = vbNormal

End Sub


Private Sub g_Click()
'
'On Error GoTo g_Click_Error
'
'If GraphDrawn Then
'  GraphDrawn = False
'  Exit Sub
'End If
'
'g.Col = 1
'
'Exit Sub
'
'g_Click_Error:
'
'Dim strES As String
'Dim intEL As Integer
'
'intEL = Erl
'strES = Err.Description
'LogError "frmIsolateReport", "g_Click", intEL, strES

End Sub
Sub DisplayGraph(ByVal Analyte As Integer)
'
'Dim Column As Integer
'Dim AnalyteName As String
'Dim Weeks As Integer
'Dim StartDate As Date
'Dim StartY As Integer
'Dim Y As Integer
'Dim gDate As Date
'Dim X As Long
'
'On Error GoTo DisplayGraph_Error
'
'Column = Analyte + 9
'
'AnalyteName = Choose(Analyte + 1, _
  '              "VRE", "ESBL", "MRSA", "Salmonella", "Shigella", "Campylobacter", _
  '              "E Coli 0...", _
  '              "Enteropathogenic E Coli", "C Difficile")
'ChartTitle = AnalyteName
'
'Weeks = HowManyWeeks()
'If Weeks = 0 Then Exit Sub
'
'ReDim Gx(1 To Weeks, 0)
'ReDim temp(1 To Weeks, 0)
'
'For Y = 1 To Weeks
'  Gx(Y, 0) = 0
'  temp(Y, 0) = 0
'Next
'
'StartY = 1
'StartDate = calToDate
'
'For Y = 1 To g.Rows - 1
'  If g.TextMatrix(Y, Column) = "X" Then
'    gDate = g.TextMatrix(Y, 0)
'    X = WeekNumber(gDate)
'    temp(X, 0) = temp(X, 0) + 1
'  End If
'Next
'
'For X = 1 To Weeks
'  Gx(X, 0) = temp(Weeks - X + 1, 0)
'Next
'
'Chart.ChartData = Gx
'
'Exit Sub
'
'DisplayGraph_Error:
'
'Dim strES As String
'Dim intEL As Integer
'
'intEL = Erl
'strES = Err.Description
'LogError "frmIsolateReport", "DisplayGraph", intEL, strES

End Sub
Function WeekNumber(ByVal Given As Date) As Integer

10        On Error GoTo WeekNumber_Error

20        WeekNumber = DateDiff("w", Format(Given, "dd/mmm/yyyy"), Format(calToDate, "dd/mmm/yyyy")) + 1

30        Exit Function

WeekNumber_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmIsolateReport", "WeekNumber", intEL, strES


End Function

Private Sub g_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'
'Dim n As Integer
'Dim c As Integer
'
'On Error GoTo g_MouseUp_Error
'
'lblTotal = 0
'
'If g.Col > 8 Then
'  c = g.Col
'  DisplayGraph Val(g.Col) - 9
'  g.Col = c
'  For n = 1 To g.Rows - 1
'    g.Row = n
'    If g.CellBackColor = vbRed Then
'      lblTotal = lblTotal + 1
'    End If
'  Next
'Else
'  Exit Sub
'End If
'
'GraphDrawn = True
'
'Exit Sub
'
'If Y < 240 Then
'  If X > 7469 And X < 7920 Then
'    DisplayGraph 0 'Salmonella
'  ElseIf X >= 7920 And X < 8325 Then
'    DisplayGraph 1 'Shigella
'  ElseIf X >= 8325 And X < 8820 Then
'    DisplayGraph 2 'Campylobacter
'  ElseIf X >= 8820 And X < 9270 Then
'    DisplayGraph 3 'E Coli 0157
'  ElseIf X > 9270 And X < 9720 Then
'    DisplayGraph 4 'Rota
'  ElseIf X >= 9720 And X < 10215 Then
'    DisplayGraph 5 'Adeno
'  ElseIf X >= 10215 And X < 10710 Then
'    DisplayGraph 6 'EPEC
'  ElseIf X >= 10710 And X < 11205 Then
'    DisplayGraph 7 'Toxin A
'  ElseIf X >= 11205 And X < 11610 Then
'    DisplayGraph 8 'O/P
'  Else
'    Beep
'  End If
'  GraphDrawn = True
'End If
'
'Exit Sub
'
'g_MouseUp_Error:
'
'Dim strES As String
'Dim intEL As Integer
'
'intEL = Erl
'strES = Err.Description
'LogError "frmIsolateReport", "g_MouseUp", intEL, strES

End Sub


