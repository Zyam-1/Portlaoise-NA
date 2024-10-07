VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFullExt 
   Caption         =   "NetAcquire - Full External History"
   ClientHeight    =   7335
   ClientLeft      =   135
   ClientTop       =   435
   ClientWidth     =   11535
   Icon            =   "frmFullExt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   11535
   Begin VB.ComboBox cmbResultCount 
      Height          =   315
      ItemData        =   "frmFullExt.frx":000C
      Left            =   3300
      List            =   "frmFullExt.frx":0022
      TabIndex        =   25
      Text            =   "All"
      Top             =   510
      Width           =   1215
   End
   Begin VB.CheckBox chkChartNumber 
      Caption         =   "Ignore chart number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   540
      Width           =   2535
   End
   Begin VB.CommandButton bPrint 
      Cancel          =   -1  'True
      Caption         =   "Print"
      Height          =   750
      Left            =   6900
      Picture         =   "frmFullExt.frx":0054
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4620
      Width           =   1245
   End
   Begin VB.CommandButton bcancel 
      Caption         =   "Exit"
      Height          =   750
      Left            =   6900
      Picture         =   "frmFullExt.frx":035E
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   6210
      Width           =   1245
   End
   Begin VB.CommandButton cmdExcel 
      Caption         =   "Export"
      Height          =   750
      Left            =   9720
      Picture         =   "frmFullExt.frx":0668
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   6210
      Width           =   1245
   End
   Begin VB.PictureBox pb 
      AutoRedraw      =   -1  'True
      DrawWidth       =   512
      Height          =   2325
      Left            =   1860
      ScaleHeight     =   2265
      ScaleWidth      =   4125
      TabIndex        =   11
      Top             =   4620
      Width           =   4185
      Begin VB.Label lblTest 
         Alignment       =   2  'Center
         Height          =   225
         Left            =   1320
         TabIndex        =   12
         Top             =   0
         Width           =   1275
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   10320
      Top             =   4740
   End
   Begin VB.Frame Frame1 
      Caption         =   "Plot between"
      Height          =   2325
      Left            =   60
      TabIndex        =   8
      Top             =   4620
      Width           =   1785
      Begin VB.CommandButton cmdGo 
         Caption         =   "Start"
         Height          =   900
         Left            =   300
         Picture         =   "frmFullExt.frx":6286
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   1320
         Width           =   1140
      End
      Begin VB.ComboBox cmbPlotTo 
         Height          =   315
         Left            =   150
         TabIndex        =   10
         Text            =   "cmbPlotTo"
         Top             =   870
         Width           =   1455
      End
      Begin VB.ComboBox cmbPlotFrom 
         Height          =   315
         Left            =   150
         TabIndex        =   9
         Text            =   "cmbPlotFrom"
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   2760
      Left            =   3930
      TabIndex        =   1
      Top             =   1350
      Visible         =   0   'False
      Width           =   1635
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   11385
      _ExtentX        =   20082
      _ExtentY        =   6376
      _Version        =   393216
      Rows            =   4
      FixedRows       =   3
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
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
   Begin MSComctlLib.ProgressBar PBar 
      Height          =   195
      Left            =   30
      TabIndex        =   13
      Top             =   7020
      Width           =   11325
      _ExtentX        =   19976
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "Show "
      Height          =   195
      Left            =   2820
      TabIndex        =   27
      Top             =   570
      Width           =   450
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "Results"
      Height          =   195
      Left            =   4560
      TabIndex        =   26
      Top             =   570
      Width           =   705
   End
   Begin VB.Label lblExcelInfo 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exporting..."
      Height          =   285
      Left            =   9720
      TabIndex        =   19
      Top             =   5880
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Label Tn 
      Height          =   225
      Left            =   2940
      TabIndex        =   18
      Top             =   2250
      Width           =   315
   End
   Begin VB.Label lblMaxVal 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   6120
      TabIndex        =   17
      Top             =   4620
      Width           =   480
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   6120
      TabIndex        =   16
      Top             =   6690
      Width           =   480
   End
   Begin VB.Label lblMeanVal 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   6120
      TabIndex        =   15
      Top             =   5610
      Width           =   480
   End
   Begin VB.Label lblNopas 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3390
      TabIndex        =   14
      Top             =   5310
      Width           =   1545
   End
   Begin VB.Label lblChart 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   990
      TabIndex        =   7
      Top             =   120
      Width           =   1545
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Chart"
      Height          =   195
      Left            =   540
      TabIndex        =   6
      Top             =   150
      Width           =   375
   End
   Begin VB.Label lblDoB 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3240
      TabIndex        =   5
      Top             =   120
      Width           =   1545
   End
   Begin VB.Label lblName 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   5790
      TabIndex        =   4
      Top             =   120
      Width           =   3765
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "DoB"
      Height          =   195
      Left            =   2850
      TabIndex        =   3
      Top             =   180
      Width           =   315
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Name"
      Height          =   195
      Left            =   5310
      TabIndex        =   2
      Top             =   180
      Width           =   420
   End
End
Attribute VB_Name = "frmFullExt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type ChartPosition
    xPos As Long
    yPos As Long
    Value As Single
    Date As String
End Type

Private ChartPositions() As ChartPosition

Private NumberOfDays As Long

Private Sub bcancel_Click()

10        Unload Me

End Sub




Private Sub bprint_Click()
          Dim x As Long
          Dim n As Long
          Dim z As Long

10        On Error GoTo bprint_Click_Error

20        x = g.Cols

30        Printer.Orientation = vbPRORLandscape


40        Printer.Font.Size = 16
50        Printer.Print Tab(15); "Cumulative Report from External Tests."
60        Printer.Print

70        Printer.Font.Size = 14
80        Printer.Print Tab(10); "Name : " & lblName;
90        Printer.Print Tab(40); "Dob  : " & lblDoB


100       Printer.Print
110       Printer.Print

120       Printer.Font.Size = 10
130       For n = 0 To g.Rows - 1
140           g.Row = n
150           For z = 0 To x - 1
160               g.Col = z
170               Printer.Print Tab(10 * z); g;
180           Next
190           Printer.Print
200       Next

210       Printer.Print Tab(30); "----End of Report----"

220       Printer.EndDoc

230       Exit Sub

bprint_Click_Error:

          Dim strES As String
          Dim intEL As Integer



240       intEL = Erl
250       strES = Err.Description
260       LogError "frmFullExt", "bPrint_Click", intEL, strES


End Sub

Private Sub chkChartNumber_Click()
10        g.Visible = False
          If cmbResultCount.Text <> "" Then
             FillG (Trim(cmbResultCount.Text))
          End If
          g.Visible = True
20        FillCombos
End Sub

Private Sub cmdExcel_Click()
          Dim strHeading As String
10        On Error GoTo cmdExcel_Click_Error

20        strHeading = "External History" & vbCr
30        strHeading = strHeading & "Patient Name: " & lblName & vbCr
40        strHeading = strHeading & vbCr

50        ExportFlexGrid g, Me, strHeading

60        Exit Sub

cmdExcel_Click_Error:

          Dim strES As String
          Dim intEL As Integer

70        intEL = Erl
80        strES = Err.Description
90        LogError "frmFullExt", "cmdExcel_Click", intEL, strES
End Sub

Private Sub cmdGo_Click()

10        On Error GoTo cmdGo_Click_Error

20        DrawChart

30        Exit Sub

cmdGo_Click_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmFullExt", "cmdGo_Click", intEL, strES


End Sub

Private Sub Form_Activate()

10        On Error GoTo Form_Activate_Error

20        Me.Refresh
30        If GetOptionSetting("EnableiPMSChart", "0") = 0 Then
40            chkChartNumber.Value = 0
50            chkChartNumber.Enabled = Not (lblChart = "")
60        Else
70            chkChartNumber.Value = 1

80        End If
90        g.Visible = False
          If cmbResultCount.Text <> "" Then
             FillG (Trim(cmbResultCount.Text))
          End If
          g.Visible = True
100       FillCombos



110       PBar.Max = LogOffDelaySecs
120       PBar = 0

130       Timer1.Enabled = True

140       Exit Sub

Form_Activate_Error:

          Dim strES As String
          Dim intEL As Integer



150       intEL = Erl
160       strES = Err.Description
170       LogError "frmFullExt", "Form_Activate", intEL, strES


End Sub

Private Sub Form_Deactivate()

10        On Error GoTo Form_Deactivate_Error

20        Timer1.Enabled = False

30        Exit Sub

Form_Deactivate_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmFullExt", "Form_Deactivate", intEL, strES


End Sub


'---------------------------------------------------------------------------------------
' Procedure : Form_MouseMove
' DateTime  : 02/07/2007 11:08
' Author    : Myles
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

10        On Error GoTo Form_MouseMove_Error

20        PBar = 0

30        Exit Sub

Form_MouseMove_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmFullExt", "Form_MouseMove", intEL, strES


End Sub


Private Sub DrawChart()

          Dim n As Long
          Dim Counter As Long
          Dim DaysInterval As Long
          Dim x As Long
          Dim Y As Long
          Dim PixelsPerDay As Single
          Dim PixelsPerPointY As Single
          Dim FirstDayFilled As Boolean
          Dim MaxVal As Single
          Dim cVal As Single
          Dim StartGridX As Long
          Dim StopGridX As Long


10        On Error GoTo DrawChart_Error

20        MaxVal = 0
30        lblMaxVal = ""
40        lblMeanVal = ""

50        pb.Cls
60        pb.Picture = LoadPicture("")
70        lblTest = g.TextMatrix(g.RowSel, 0)

80        NumberOfDays = DateDiff("d", Format(cmbPlotFrom, "dd/mmm/yyyy"), Format(cmbPlotTo, "dd/mmm/yyyy"))
90        If NumberOfDays = 0 Then Exit Sub

100       ReDim ChartPositions(0 To NumberOfDays)

110       For n = 1 To NumberOfDays
120           ChartPositions(n).xPos = 0
130           ChartPositions(n).yPos = 0
140           ChartPositions(n).Value = 0
150           ChartPositions(n).Date = ""
160       Next

170       For n = 1 To g.Cols - 1
180           If Format$(g.TextMatrix(1, n), "dd/mmm/yyyy") = cmbPlotTo Then StartGridX = n
190           If Format$(g.TextMatrix(1, n), "dd/mmm/yyyy") = cmbPlotFrom Then StopGridX = n
200       Next

210       FirstDayFilled = False
220       Counter = 0
230       For x = StartGridX To StopGridX
240           If g.TextMatrix(g.Row, x) <> "" Then
250               If Not FirstDayFilled Then
260                   FirstDayFilled = True
270                   MaxVal = Val(g.TextMatrix(g.Row, x))
280                   ChartPositions(NumberOfDays).Date = Format(g.TextMatrix(1, x), "dd/mmm/yyyy")
                      'LatestDate = Format$(g.TextMatrix(1, x), "dd/mmm/yyyy")
290                   ChartPositions(NumberOfDays).Value = Val(g.TextMatrix(g.Row, x))
300               Else
310                   DaysInterval = Abs(DateDiff("D", cmbPlotTo, Format$(g.TextMatrix(1, x), "dd/mmm/yyyy")))
320                   ChartPositions(NumberOfDays - DaysInterval).Date = g.TextMatrix(1, x)
330                   cVal = Val(g.TextMatrix(g.Row, x))
340                   ChartPositions(NumberOfDays - DaysInterval).Value = cVal
350                   If cVal > MaxVal Then MaxVal = cVal
                      '        EarliestDate = Format$(g.TextMatrix(1, x), "dd/mmm/yyyy")
                      '      Else
                      '        Exit For
                      '      End If
360               End If
                  '    Counter = Counter + 1
                  '    If Counter = 15 Then
                  '      Exit For
                  '    End If
370           End If
380       Next

          'If EarliestDate = "" Or LatestDate = "" Then Exit Sub

          'numberOfDays = Abs(DateDiff("d", EarliestDate, LatestDate))
390       PixelsPerDay = (pb.Width - 1060) / NumberOfDays
400       MaxVal = MaxVal * 1.1
410       If MaxVal = 0 Then Exit Sub
420       PixelsPerPointY = pb.Height / MaxVal

430       x = 580 + (NumberOfDays * PixelsPerDay)
440       Y = pb.Height - (ChartPositions(NumberOfDays).Value * PixelsPerPointY)
450       ChartPositions(NumberOfDays).yPos = Y
460       ChartPositions(NumberOfDays).xPos = x

470       pb.ForeColor = vbBlue
480       pb.Circle (x, Y), 30
490       pb.Line (x - 15, Y - 15)-(x + 15, Y + 15), vbBlue, BF
500       pb.PSet (x, Y)

510       For n = NumberOfDays - 1 To 0 Step -1
520           If ChartPositions(n).Value <> 0 And ChartPositions(n).Date <> "" Then
530               DaysInterval = Abs(DateDiff("d", cmbPlotFrom, Format(ChartPositions(n).Date, "dd/mmm/yyyy")))
540               x = 580 + (DaysInterval * PixelsPerDay)
550               ChartPositions(n).xPos = x
560               Y = pb.Height - (ChartPositions(n).Value * PixelsPerPointY)
570               ChartPositions(n).yPos = Y
580               pb.Line -(x, Y)
590               pb.Line (x - 15, Y - 15)-(x + 15, Y + 15), vbBlue, BF
600               pb.Circle (x, Y), 30
610               pb.PSet (x, Y)
620           End If
630       Next

640       pb.Line (0, pb.Height / 2)-(pb.Width, pb.Height / 2), vbBlack, BF

650       lblMaxVal = Format$(Int(MaxVal + 0.5), "###0")
660       lblMeanVal = Format$(Val(lblMaxVal) / 2, "###0.0")



670       Exit Sub

DrawChart_Error:

          Dim strES As String
          Dim intEL As Integer

680       intEL = Erl
690       strES = Err.Description
700       LogError "frmFullExt", "DrawChart", intEL, strES


End Sub
Private Sub FillCombos()

          Dim x As Long

10        On Error GoTo FillCombos_Error

20        cmbPlotFrom.Clear
30        cmbPlotTo.Clear

40        For x = 3 To g.Cols - 1
50            cmbPlotFrom.AddItem Format$(g.TextMatrix(1, x), "dd/mmm/yyyy")
60            cmbPlotTo.AddItem Format$(g.TextMatrix(1, x), "dd/mmm/yyyy")
70        Next

80        cmbPlotTo = Format$(g.TextMatrix(1, 3), "dd/mmm/yyyy")
90        If Not IsDate(cmbPlotTo) Then Exit Sub

100       For x = g.Cols - 1 To 1 Step -1
110           If DateDiff("d", Format$(g.TextMatrix(1, x), "dd/mmm/yyyy"), cmbPlotTo) < 365 Then
120               cmbPlotFrom = Format$(g.TextMatrix(1, x), "dd/mmm/yyyy")
130               Exit For
140           End If
150       Next

160       Exit Sub

FillCombos_Error:

          Dim strES As String
          Dim intEL As Integer



170       intEL = Erl
180       strES = Err.Description
190       LogError "frmFullExt", "FillCombos", intEL, strES


End Sub


Private Sub FillG(p_RCount As String)

          Dim tb As New Recordset
          Dim snr As Recordset
          Dim sql As String
          Dim gcolumns As Long
          Dim x As Long
          Dim xrun As String
          Dim xdate As String
          Dim Code As String
          Dim sex As String
          Dim n As Integer


10        On Error GoTo FillG_Error
20        p_RCount = UCase(p_RCount)
30        If p_RCount = "First 5" Then
40            sql = "SELECT DISTINCT top 5 (D.SampleID), D.RunDate, D.TimeTaken, D.Sex, D.SampleDate " & _
                  "FROM Demographics D, ExtResults R WHERE ("
50        ElseIf p_RCount = "First 10" Then
60            sql = "SELECT DISTINCT top 10 (D.SampleID), D.RunDate, D.TimeTaken, D.Sex, D.SampleDate " & _
                  "FROM Demographics D, ExtResults R WHERE ("
70        ElseIf p_RCount = "First 20" Then
80            sql = "SELECT DISTINCT top 20 (D.SampleID), D.RunDate, D.TimeTaken, D.Sex, D.SampleDate " & _
                  "FROM Demographics D, ExtResults R WHERE ("
90        ElseIf p_RCount = "First 50" Then
100           sql = "SELECT DISTINCT top 50 (D.SampleID), D.RunDate, D.TimeTaken, D.Sex, D.SampleDate " & _
                  "FROM Demographics D, ExtResults R WHERE ("
110       ElseIf p_RCount = "ALL" Then
120           sql = "SELECT DISTINCT (D.SampleID), D.RunDate, D.TimeTaken, D.Sex, D.SampleDate " & _
                  "FROM Demographics D, ExtResults R WHERE ("
130       End If

140       If Trim(lblChart) <> "" And chkChartNumber.Value = 0 Then
150           sql = sql & "(D.Chart = '" & lblChart & "') and"
160       End If
170       sql = sql & " (D.PatName = '" & AddTicks(lblName) & "' AND D.DoB  = '" & Format(lblDoB, "dd/MMM/yyyy") & "') "

180       sql = sql & ") AND R.SampleID = D.SampleID " & _
              "ORDER BY D.SampleDate desc"

190       Set tb = New Recordset
200       RecOpenServer 0, tb, sql
210       If Not tb.EOF Then
220           sex = tb!sex & ""

230           g.Cols = 3
240           g.ColWidth(0) = 1600
250           g.ColWidth(1) = 500
260           g.ColWidth(2) = 1400
270           g.TextMatrix(0, 0) = "SAMPLE ID"
280           g.TextMatrix(1, 0) = "SAMPLE DATE"
290           g.TextMatrix(2, 0) = "RUN DATE"
300           g.TextMatrix(2, 1) = "S/T"
310           g.TextMatrix(2, 2) = "Ref Ranges"

320           Do While Not tb.EOF
330               g.Cols = g.Cols + 1
340               x = g.Cols - 1
350               gcolumns = x + 1
360               g.ColWidth(x) = 1500
370               g.Col = x
380               xrun = tb!SampleID & ""
390               g.TextMatrix(0, x) = xrun
                  'Sample DateTime
400               If Not IsNull(tb!SampleDate) Then
410                   xdate = Format(tb!SampleDate, "dd/mm/yy")
420                   If Format(tb!SampleDate, "hh:mm") <> "00:00" Then
430                       xdate = xdate & " " & Format(tb!SampleDate, "hh:mm")
440                   End If
450               Else
460                   xdate = ""
470               End If
480               g.TextMatrix(1, x) = xdate
                  'Run DateTime
490               If Not IsNull(tb!Rundate) Then
500                   xdate = Format(tb!Rundate, "dd/mm/yy")
510               Else
520                   xdate = ""
530               End If
540               g.TextMatrix(2, x) = xdate


                  '        If Not IsNull(tb!SampleDate) Then
                  '            xdate = Format$(tb!SampleDate, "dd/mm/yy")
                  '        Else
                  '            xdate = ""
                  '        End If
                  '        g.TextMatrix(1, X) = xdate
                  '        If IsDate(tb!SampleDate) Then
                  '            g.TextMatrix(2, X) = Format(tb!SampleDate, "hh:mm")
                  '        Else
                  '            g.TextMatrix(2, X) = ""
                  '        End If
                  'fill list with test names
550               sql = "SELECT * from extResults " & _
                      "WHERE SampleID = " & xrun & " order by orderlist"
560               Set snr = New Recordset
570               RecOpenServer 0, snr, sql
580               Do While Not snr.EOF
590                   Code = Trim(snr!Analyte)
600                   If Not InList(Code) Then
610                       List1.AddItem Code
620                   End If
630                   snr.MoveNext
640               Loop
650               tb.MoveNext
660           Loop
670           If List1.ListCount = 0 Then Exit Sub
              'fill first col with test names
680           TransferListToGrid

              'fill in results
690           For x = 3 To gcolumns - 1
700               g.Col = x
710               g.Row = 1
720               xdate = Format$(g, "dd/mmm/yyyy")
730               g.Row = 0
740               xrun = g

750               sql = "SELECT * from extResults "
760               sql = sql & "WHERE SampleID = " & xrun & ""
770               Set snr = New Recordset
780               RecOpenServer 0, snr, sql
790               Do While Not snr.EOF
800                   Code = Trim(snr!Analyte)
810                   g.Row = GetRow(Code)
820                   n = Len(Format$(snr!Result & "", "0.00")) * 120
830                   If n > g.ColWidth(x) Then g.ColWidth(x) = n
840                   If Trim(snr!Result & "") = "" Then
850                       g = "Out Standing"
860                   Else
870                       g = Format$(snr!Result & "", "0.00")
880                   End If
890                   snr.MoveNext
900               Loop
910           Next
920       End If

930       g.Visible = True

940       Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer

950       intEL = Erl
960       strES = Err.Description
970       LogError "frmFullExt", "FillG", intEL, strES, sql


End Sub
Private Function GetRow(ByVal testnum As String) As Long

          Dim n As Long

10        On Error GoTo GetRow_Error

20        For n = 0 To List1.ListCount - 1
30            If testnum = List1.List(n) Then
40                GetRow = n + 3
50                Exit For
60            End If
70        Next

80        Exit Function

GetRow_Error:

          Dim strES As String
          Dim intEL As Integer



90        intEL = Erl
100       strES = Err.Description
110       LogError "frmFullExt", "GetRow", intEL, strES


End Function



Private Function InList(ByVal s As String) As Long

          Dim n As Long

10        On Error GoTo InList_Error

20        InList = False
30        If List1.ListCount = 0 Then
40            Exit Function
50        End If

60        For n = 0 To List1.ListCount - 1
70            If s = List1.List(n) Then
80                InList = True
90                Exit For
100           End If
110       Next

120       Exit Function

InList_Error:

          Dim strES As String
          Dim intEL As Integer



130       intEL = Erl
140       strES = Err.Description
150       LogError "frmFullExt", "InList", intEL, strES


End Function

Private Sub TransferListToGrid()

          Dim n As Long
          Dim sql As String

10        On Error GoTo TransferListToGrid_Error

20        If List1.ListCount = 0 Then Exit Sub

30        g.Rows = List1.ListCount + 3

40        g.Col = 0
50        For n = 0 To List1.ListCount - 1
60            g.Row = n + 3
70            g = List1.List(n)
80        Next

90        Exit Sub

TransferListToGrid_Error:

          Dim strES As String
          Dim intEL As Integer

100       intEL = Erl
110       strES = Err.Description
120       LogError "frmFullExt", "TransferListToGrid", intEL, strES, sql

End Sub


Private Sub g_Click()

10        On Error GoTo g_Click_Error

20        DrawChart

30        Exit Sub

g_Click_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmFullExt", "g_Click", intEL, strES


End Sub


Private Sub g_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

10        On Error GoTo g_MouseMove_Error

20        PBar = 0

30        If g.Col > 2 And g.Row > 2 Then
40            Y = g.MouseCol
50            x = g.MouseRow
60            g.ToolTipText = g.TextMatrix(x, Y)
70        End If

80        Exit Sub

g_MouseMove_Error:

          Dim strES As String
          Dim intEL As Integer



90        intEL = Erl
100       strES = Err.Description
110       LogError "frmFullExt", "g_MouseMove", intEL, strES


End Sub



Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

10        On Error GoTo Label1_MouseMove_Error

20        PBar = 0

30        Exit Sub

Label1_MouseMove_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmFullExt", "Label1_MouseMove", intEL, strES


End Sub


Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

10        On Error GoTo Label2_MouseMove_Error

20        PBar = 0

30        Exit Sub

Label2_MouseMove_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmFullExt", "Label2_MouseMove", intEL, strES


End Sub


Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

10        On Error GoTo Label3_MouseMove_Error

20        PBar = 0

30        Exit Sub

Label3_MouseMove_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmFullExt", "Label3_MouseMove", intEL, strES


End Sub



Private Sub lblChart_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

10        On Error GoTo lblChart_MouseMove_Error

20        PBar = 0

30        Exit Sub

lblChart_MouseMove_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmFullExt", "lblChart_MouseMove", intEL, strES


End Sub


Private Sub lblDoB_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

10        On Error GoTo lblDoB_MouseMove_Error

20        PBar = 0

30        Exit Sub

lblDoB_MouseMove_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmFullExt", "lblDoB_MouseMove", intEL, strES


End Sub


Private Sub lblName_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

10        On Error GoTo lblName_MouseMove_Error

20        PBar = 0

30        Exit Sub

lblName_MouseMove_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmFullExt", "lblName_MouseMove", intEL, strES


End Sub




Private Sub pb_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

          Dim i As Long
          Dim CurrentDistance As Long
          Dim BestDistance As Long
          Dim BestIndex As Integer


10        On Error GoTo pb_MouseMove_Error

20        PBar = 0

30        If NumberOfDays = 0 Then Exit Sub

40        BestIndex = -1
50        BestDistance = 99999
60        For i = 0 To NumberOfDays
70            CurrentDistance = ((x - ChartPositions(i).xPos) ^ 2 + (Y - ChartPositions(i).yPos) ^ 2) ^ (1 / 2)
80            If i = 0 Or CurrentDistance < BestDistance Then
90                BestDistance = CurrentDistance
100               BestIndex = i
110           End If
120       Next

130       If BestIndex <> -1 Then
140           pb.ToolTipText = Format$(ChartPositions(BestIndex).Date, "dd/mmm/yyyy") & " " & ChartPositions(BestIndex).Value
150       End If


160       Exit Sub

pb_MouseMove_Error:

          Dim strES As String
          Dim intEL As Integer



170       intEL = Erl
180       strES = Err.Description
190       LogError "frmFullExt", "pb_MouseMove", intEL, strES


End Sub


Private Sub Timer1_Timer()

      'tmrRefresh.Interval set to 1000
10        On Error GoTo Timer1_Timer_Error

20        PBar = PBar + 1

30        If PBar = PBar.Max Then
40            Unload Me
50        End If

60        Exit Sub

Timer1_Timer_Error:

          Dim strES As String
          Dim intEL As Integer



70        intEL = Erl
80        strES = Err.Description
90        LogError "frmFullExt", "Timer1_Timer", intEL, strES


End Sub

Private Sub cmbResultCount_Change()
    On Error GoTo cmbResultCount_Change_Error
    
    g.Visible = False
     
    If cmbResultCount.Text <> "" Then
         FillG (Trim(cmbResultCount.Text))
    End If
    
    g.Visible = True
    
cmbResultCount_Change_Error:

    Dim strES As String
    Dim intEL As Integer
    
    intEL = Erl
    strES = Err.Description
    LogError "frmFullBio", "cmbResultCount_Change", intEL, strES
End Sub

Private Sub cmbResultCount_Click()
    On Error GoTo cmbResultCount_Click_Error
    
    g.Visible = False
     
    If cmbResultCount.Text <> "" Then
         FillG (Trim(cmbResultCount.Text))
    End If
    
    g.Visible = True
    
cmbResultCount_Click_Error:

    Dim strES As String
    Dim intEL As Integer
    
    intEL = Erl
    strES = Err.Description
    LogError "frmFullBio", "cmbResultCount_Click", intEL, strES

End Sub

