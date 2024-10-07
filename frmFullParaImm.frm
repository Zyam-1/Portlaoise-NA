VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmFullParaImm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Full Immunology Paraprotein History"
   ClientHeight    =   6225
   ClientLeft      =   165
   ClientTop       =   465
   ClientWidth     =   11565
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "frmFullParaImm.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6225
   ScaleWidth      =   11565
   Begin VB.ComboBox cmbResultCount 
      Height          =   315
      ItemData        =   "frmFullParaImm.frx":030A
      Left            =   570
      List            =   "frmFullParaImm.frx":0320
      TabIndex        =   26
      Text            =   "First 5"
      Top             =   30
      Width           =   1215
   End
   Begin VB.CommandButton cmdExcel 
      Caption         =   "Export"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   7110
      Picture         =   "frmFullParaImm.frx":0352
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   5040
      Width           =   1245
   End
   Begin VB.CommandButton bPrint 
      Cancel          =   -1  'True
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   8415
      Picture         =   "frmFullParaImm.frx":5F70
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Print Report"
      Top             =   5040
      Width           =   1245
   End
   Begin VB.CommandButton bcancel 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   9720
      Picture         =   "frmFullParaImm.frx":627A
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Exit Report"
      Top             =   5040
      Width           =   1245
   End
   Begin VB.Frame Frame1 
      Caption         =   "Plot between"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   930
      Left            =   6780
      TabIndex        =   11
      Top             =   1245
      Width           =   4725
      Begin VB.CommandButton cmdGo 
         Caption         =   "Go"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   3570
         Picture         =   "frmFullParaImm.frx":6584
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Draw Graph"
         Top             =   180
         Width           =   1095
      End
      Begin VB.ComboBox cmbPlotFrom 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   150
         TabIndex        =   13
         Text            =   "cmbPlotFrom"
         Top             =   240
         Width           =   1680
      End
      Begin VB.ComboBox cmbPlotTo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1890
         TabIndex        =   12
         Text            =   "cmbPlotTo"
         Top             =   240
         Width           =   1635
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   10740
      Top             =   720
   End
   Begin VB.PictureBox pb 
      AutoRedraw      =   -1  'True
      Height          =   2325
      Left            =   6780
      ScaleHeight     =   2265
      ScaleWidth      =   4125
      TabIndex        =   8
      Top             =   2250
      Width           =   4185
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   5385
      Left            =   60
      TabIndex        =   7
      Top             =   405
      Width           =   6585
      _ExtentX        =   11615
      _ExtentY        =   9499
      _Version        =   393216
      Rows            =   4
      FixedRows       =   3
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      ScrollTrack     =   -1  'True
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   2955
      Left            =   -120
      TabIndex        =   6
      Top             =   2010
      Visible         =   0   'False
      Width           =   1635
   End
   Begin MSComctlLib.ProgressBar PBar 
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   5880
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
      Left            =   90
      TabIndex        =   28
      Top             =   90
      Width           =   450
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "Results"
      Height          =   195
      Left            =   1830
      TabIndex        =   27
      Top             =   90
      Width           =   705
   End
   Begin VB.Label lblExcelInfo 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exporting..."
      Height          =   285
      Left            =   7110
      TabIndex        =   25
      Top             =   4800
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Label lblNoRes 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   10230
      TabIndex        =   22
      ToolTipText     =   "Number of Results"
      Top             =   4650
      Width           =   645
   End
   Begin VB.Label lblResInfo 
      Caption         =   "No. of Results"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   8835
      TabIndex        =   23
      Top             =   4650
      Width           =   1680
   End
   Begin VB.Label lblNopas 
      Height          =   255
      Left            =   1080
      TabIndex        =   21
      Top             =   720
      Width           =   1365
   End
   Begin VB.Label Tn 
      Height          =   225
      Left            =   3840
      TabIndex        =   20
      Top             =   1410
      Width           =   315
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   6720
      Picture         =   "frmFullParaImm.frx":688E
      Top             =   780
      Width           =   480
   End
   Begin VB.Label lblMaxVal 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10980
      TabIndex        =   17
      Top             =   2250
      Width           =   480
   End
   Begin VB.Label lblMeanVal 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10980
      TabIndex        =   16
      Top             =   3300
      Width           =   480
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10980
      TabIndex        =   15
      Top             =   4320
      Width           =   480
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Click on Parameter to show Graph"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7230
      TabIndex        =   9
      Top             =   870
      Width           =   2715
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6780
      TabIndex        =   5
      Top             =   450
      Width           =   420
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "DoB"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   9090
      TabIndex        =   4
      Top             =   180
      Width           =   315
   End
   Begin VB.Label lblName 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   7230
      TabIndex        =   3
      ToolTipText     =   "Patients Name"
      Top             =   450
      Width           =   3765
   End
   Begin VB.Label lblDoB 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   9450
      TabIndex        =   2
      ToolTipText     =   "Date of Birth"
      Top             =   150
      Width           =   1545
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Chart"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6810
      TabIndex        =   1
      Top             =   210
      Width           =   375
   End
   Begin VB.Label lblChart 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   7230
      TabIndex        =   0
      ToolTipText     =   "Chart"
      Top             =   150
      Width           =   1545
   End
End
Attribute VB_Name = "frmFullParaImm"
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
Private sex As String

Private NumberOfDays As Long
Private Sub DrawChart()

          Dim n As Long
          Dim Counter As Long
          Dim DaysInterval As Long
          Dim X As Long
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

70        NumberOfDays = DateDiff("d", Format(cmbPlotFrom, "dd/mmm/yyyy"), Format(cmbPlotTo, "dd/mmm/yyyy"))
80        ReDim ChartPositions(0 To NumberOfDays)

90        For n = 1 To NumberOfDays
100           ChartPositions(n).xPos = 0
110           ChartPositions(n).yPos = 0
120           ChartPositions(n).Value = 0
130           ChartPositions(n).Date = ""
140       Next

150       For n = 1 To g.Cols - 1
160           If Format$(g.TextMatrix(1, n), "dd/mmm/yyyy") = cmbPlotTo Then StartGridX = n
170           If Format$(g.TextMatrix(1, n), "dd/mmm/yyyy") = cmbPlotFrom Then StopGridX = n
180       Next

190       FirstDayFilled = False
200       Counter = 0
210       For X = StartGridX To StopGridX
220           If g.TextMatrix(g.row, X) <> "" Then
230               If Not FirstDayFilled Then
240                   FirstDayFilled = True
250                   MaxVal = Val(g.TextMatrix(g.row, X))
260                   ChartPositions(NumberOfDays).Date = Format(g.TextMatrix(1, X), "dd/mmm/yyyy")
270                   ChartPositions(NumberOfDays).Value = Val(g.TextMatrix(g.row, X))
280               Else
290                   DaysInterval = Abs(DateDiff("D", cmbPlotTo, Format$(g.TextMatrix(1, X), "dd/mmm/yyyy")))
300                   ChartPositions(NumberOfDays - DaysInterval).Date = g.TextMatrix(1, X)
310                   cVal = Val(g.TextMatrix(g.row, X))
320                   ChartPositions(NumberOfDays - DaysInterval).Value = cVal
330                   If cVal > MaxVal Then MaxVal = cVal
340               End If
350           End If
360       Next

370       PixelsPerDay = (pb.Width - 1060) / NumberOfDays
380       MaxVal = MaxVal * 1.1
390       If MaxVal = 0 Then Exit Sub
400       PixelsPerPointY = pb.Height / MaxVal

410       X = 580 + (NumberOfDays * PixelsPerDay)
420       Y = pb.Height - (ChartPositions(NumberOfDays).Value * PixelsPerPointY)
430       ChartPositions(NumberOfDays).yPos = Y
440       ChartPositions(NumberOfDays).xPos = X

450       pb.ForeColor = vbBlue
460       pb.Circle (X, Y), 30
470       pb.Line (X - 15, Y - 15)-(X + 15, Y + 15), vbBlue, BF
480       pb.PSet (X, Y)

490       For n = NumberOfDays - 1 To 0 Step -1
500           If ChartPositions(n).Value <> 0 And ChartPositions(n).Date <> "" Then
510               DaysInterval = Abs(DateDiff("d", cmbPlotFrom, Format(ChartPositions(n).Date, "dd/mmm/yyyy")))
520               X = 580 + (DaysInterval * PixelsPerDay)
530               ChartPositions(n).xPos = X
540               Y = pb.Height - (ChartPositions(n).Value * PixelsPerPointY)
550               ChartPositions(n).yPos = Y
560               pb.Line -(X, Y)
570               pb.Line (X - 15, Y - 15)-(X + 15, Y + 15), vbBlue, BF
580               pb.Circle (X, Y), 30
590               pb.PSet (X, Y)
600           End If
610       Next

620       pb.Line (0, pb.Height / 2)-(pb.Width, pb.Height / 2), vbBlack, BF

630       lblMaxVal = Format$(Int(MaxVal + 0.5), "###0")
640       lblMeanVal = Format$(Val(lblMaxVal) / 2, "###0.0")

650       Exit Sub

DrawChart_Error:

          Dim strES As String
          Dim intEL As Integer

660       intEL = Erl
670       strES = Err.Description
680       LogError "frmFullParaImm", "DrawChart", intEL, strES

End Sub
Private Sub FillCombos()

          Dim X As Long

10        On Error GoTo FillCombos_Error

20        cmbPlotFrom.Clear
30        cmbPlotTo.Clear

40        If g.Cols < 3 Then Exit Sub

50        For X = 1 To g.Cols - 1
60            cmbPlotFrom.AddItem Format$(g.TextMatrix(1, X), "dd/mmm/yyyy")
70            cmbPlotTo.AddItem Format$(g.TextMatrix(1, X), "dd/mmm/yyyy")
80        Next

90        cmbPlotTo = Format$(g.TextMatrix(1, 3), "dd/mmm/yyyy")

100       For X = g.Cols - 1 To 3 Step -1
110           If g.TextMatrix(1, X) <> "" Then
120               If DateDiff("d", Format$(g.TextMatrix(1, X), "dd/mmm/yyyy"), cmbPlotTo) < 365 Then
130                   cmbPlotFrom = Format$(g.TextMatrix(1, X), "dd/mmm/yyyy")
140                   Exit For
150               End If
160           End If
170       Next

180       Exit Sub

FillCombos_Error:

          Dim strES As String
          Dim intEL As Integer



190       intEL = Erl
200       strES = Err.Description
210       LogError "frmFullParaImm", "FillCombos", intEL, strES

End Sub

Private Sub bcancel_Click()

10        Unload Me

End Sub

Private Sub FillG(p_RCount As String)

          Dim tb As New Recordset
          Dim snr As Recordset
          Dim sql As String
          Dim X As Long
          Dim xrun As String
          Dim xdate As String
          Dim BRs As New BIEResults
          Dim br As BIEResult
          Dim s As String
          Dim Cat As String

10        On Error GoTo FillG_Error

          If p_RCount = "First 5" Then
20          sql = "SELECT DISTINCT top 5 (D.SampleID), D.RecDate, D.SampleDate, D.Sex, D.Category, D.RunDate, D.TimeTaken " & _
                    "FROM Demographics AS D, ImmResults AS I WHERE ("
          ElseIf p_RCount = "First 10" Then
            sql = "SELECT DISTINCT top 10 (D.SampleID), D.RecDate, D.SampleDate, D.Sex, D.Category, D.RunDate, D.TimeTaken " & _
                    "FROM Demographics AS D, ImmResults AS I WHERE ("
          ElseIf p_RCount = "First 20" Then
            sql = "SELECT DISTINCT top 20 (D.SampleID), D.RecDate, D.SampleDate, D.Sex, D.Category, D.RunDate, D.TimeTaken " & _
                    "FROM Demographics AS D, ImmResults AS I WHERE ("
          ElseIf p_RCount = "First 50" Then
            sql = "SELECT DISTINCT top 50 (D.SampleID), D.RecDate, D.SampleDate, D.Sex, D.Category, D.RunDate, D.TimeTaken " & _
                    "FROM Demographics AS D, ImmResults AS I WHERE ("
          ElseIf p_RCount = "ALL" Then
            sql = "SELECT DISTINCT(D.SampleID), D.RecDate, D.SampleDate, D.Sex, D.Category, D.RunDate, D.TimeTaken " & _
                    "FROM Demographics AS D, ImmResults AS I WHERE ("
          End If

30        If Trim(lblChart) <> "" Then
40            sql = sql & "(D.Chart = '" & lblChart & "') OR "
50        End If
60        sql = sql & "(D.PatName = '" & AddTicks(lblName) & "' " & _
                "AND D.DoB  = '" & Format(lblDoB, "dd/MMM/yyyy") & "') ) " & _
                "AND D.SampleID = I.SampleID " & _
                "AND ( (I.code = '" & SysOptImmCodeForProt(0) & "') " & _
                "   OR (I.code = '" & SysOptImmCodeForAlb(0) & "') " & _
                "   OR (I.code = '" & SysOptImmCodeForPara1(0) & "') " & _
                "   OR (I.code = '" & SysOptImmCodeForPara2(0) & "') " & _
                "   or (I.code = '" & SysOptImmCodeForPara3(0) & "') " & _
                "   or (I.code = '" & SysOptImmCodeForIGG(0) & "') " & _
                "   or (I.code = '" & SysOptImmCodeForIGA(0) & "') " & _
                "   or (I.code = '" & SysOptImmCodeForUPara(0) & "') " & _
                "   or (I.code = '" & SysOptImmCodeForIGM(0) & "') " & _
                "   or (I.code = '" & SysOptImmCodeForB2M(0) & "') " & _
                "   or (I.code = '" & SysOptImmCodeForIGGP(0) & "') " & _
                "   or (I.code = '" & SysOptImmCodeForIGAP(0) & "') " & _
                "   or (I.code = '" & SysOptImmCodeForIGMP(0) & "')) " & _
                "AND D.rundate > '01/May/2005' " & _
                "order by D.sampledate desc"

70        Set tb = New Recordset
80        RecOpenClient 0, tb, sql
90        If Not tb.EOF Then
100           sex = tb!sex & ""
110           g.Visible = False
120           g.Cols = 3
130           g.ColWidth(0) = 1795
140           g.ColWidth(1) = 495
150           g.ColWidth(2) = 1395
160           g.TextMatrix(0, 0) = "SAMPLE ID"
170           g.TextMatrix(1, 0) = "SAMPLE DATE"
180           g.TextMatrix(2, 0) = "SAMPLE TIME"
190           g.TextMatrix(2, 1) = "S/T"
200           g.TextMatrix(0, 2) = "Ref Ranges"

210           Do While Not tb.EOF
                  'SampleID and sampledate across
220               g.Cols = g.Cols + 1
230               X = g.Cols - 1
240               g.ColWidth(X) = 1095
250               g.Col = X
260               xrun = tb!SampleID & ""
270               g.TextMatrix(0, X) = xrun
280               If Not IsNull(tb!SampleDate) Then
290                   xdate = Format$(tb!SampleDate, "dd/mm/yy")
300               ElseIf Not IsNull(tb!RecDate) Then
310                   xdate = Format$(tb!RecDate, "dd/mm/yy")
320               ElseIf Not IsNull(tb!Rundate) Then
330                   xdate = Format$(tb!Rundate, "dd/mm/yy")
340               End If
350               g.TextMatrix(1, X) = xdate
360               If g.TextMatrix(2, X) = "" Then
370                   If Not IsNull(tb!SampleDate) Then g.TextMatrix(2, X) = Format(tb!SampleDate, "hh:mm")
380               End If
                  'fill list with test names
390               sql = "SELECT DISTINCT I.Code FROM ImmResults AS I, ImmTestDefinitions " & _
                        "WHERE SampleID = '" & xrun & "' " & _
                        "and ImmTestDefinitions.Code = I.Code " & _
                        "and ImmTestDefinitions.Hospital = '" & HospName(0) & "' and  " & _
                        "((I.code = '" & SysOptImmCodeForProt(0) & "') or " & _
                        " (I.code = '" & SysOptImmCodeForAlb(0) & "') or " & _
                        " (I.code = '" & SysOptImmCodeForPara1(0) & "') or " & _
                        " (I.code = '" & SysOptImmCodeForPara2(0) & "') or " & _
                        " (I.code = '" & SysOptImmCodeForPara3(0) & "') or " & _
                        " (I.code = '" & SysOptImmCodeForIGG(0) & "') or " & _
                        " (I.code = '" & SysOptImmCodeForIGA(0) & "') or " & _
                        " (I.code = '" & SysOptImmCodeForUPara(0) & "') or " & _
                        " (I.code = '" & SysOptImmCodeForIGM(0) & "') or " & _
                        " (I.code = '" & SysOptImmCodeForB2M(0) & "') or " & _
                        " (I.code = '" & SysOptImmCodeForIGGP(0) & "') or " & _
                        " (I.code = '" & SysOptImmCodeForIGAP(0) & "') or " & _
                        " (I.code = '" & SysOptImmCodeForIGMP(0) & "'))"
400               Set snr = New Recordset
410               RecOpenClient 0, snr, sql
420               Do While Not snr.EOF
430                   If Not InList(snr!Code) Then
440                       List1.AddItem snr!Code
450                   End If
460                   snr.MoveNext
470               Loop
480               tb.MoveNext
490           Loop
500           If List1.ListCount = 0 Then Exit Sub
              'fill first col with test names
510           ChangeListOrder
520           TransferListToGrid

530           For X = 3 To g.Cols - 1
540               g.Col = X
550               g.row = 1
560               xdate = Format$(g, "dd/mmm/yyyy")
570               g.row = 0
580               xrun = g
590               g.row = 3
600               Cat = g

610               If xrun <> "" Then
620                   If UserMemberOf = "Secretarys" Then
630                       Set BRs = BRs.Load("imm", xrun, "Results", gVALID, gDONTCARE, 0, "Default", xdate)
640                   Else
650                       Set BRs = BRs.Load("imm", xrun, "Results", gDONTCARE, gDONTCARE, 0, "Default", xdate)
660                   End If
670                   If BRs.Count = 0 Then g.TextMatrix(4, X) = "Error"
680                   For Each br In BRs
690                       g.row = GetRow(br.Code)
700                       If g.row > 0 Then
710                           If (UserMemberOf = "Secretarys" Or SysOptNoCumShow(0)) And br.Valid <> True Then
720                               g = "NV"
730                           Else
740                               Select Case br.Printformat
                                  Case 0: g = Format$(br.Result, "0")
750                               Case 1: g = Format$(br.Result, "0.0")
760                               Case 2: g = Format$(br.Result, "0.00")
770                               Case 3: g = Format$(br.Result, "0.000")
780                               Case 4: g = Format$(br.Result, "0.0000")
790                               End Select
800                           End If
810                           s = QuickInterpImm(br)
820                           If Trim(s) = "H" Then
830                               g.CellForeColor = SysOptHighFore(0)
840                               g.CellBackColor = SysOptHighBack(0)
850                           ElseIf Trim(s) = "L" Then
860                               g.CellForeColor = SysOptLowFore(0)
870                               g.CellBackColor = SysOptLowBack(0)
880                           Else
890                               g.CellForeColor = vbBlack
900                               g.CellBackColor = vbWhite
910                           End If
920                       End If
930                   Next
940               End If
950           Next
960       End If

970       If g.Cols > 3 Then
980           lblNoRes.Caption = g.Cols - 3
990       Else
1000          lblNoRes.Caption = "0"
1010      End If

1020      g.Visible = True

1030      Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer

1040      intEL = Erl
1050      strES = Err.Description
1060      LogError "frmFullParaImm", "FillG", intEL, strES, sql

End Sub

Private Sub bprint_Click()

          Dim X As Long
          Dim n As Long
          Dim z As Long

10        On Error GoTo bprint_Click_Error

20        X = g.Cols

30        If X > 8 Then X = 8

          'Printer.Orientation = vbPRORLandscape

40        Printer.Font.Size = 14
50        Printer.Print Tab(5); "Cumulative Report from Immunology Department " & initial2upper(HospName(0)) & "  Phone : " & SysOptImmPhone(0)

60        Printer.Print

70        Printer.Font.Size = 12
80        Printer.Print Tab(20); "Name : " & lblName;
90        Printer.Print Tab(60); "Dob  : " & lblDoB


100       Printer.Print
110       Printer.Print

120       Printer.Font.Size = 8
130       For n = 0 To g.Rows - 1
140           If n < 2 Or n > 2 Then
150               g.row = n
160               For z = 0 To X - 1
170                   If z <> 1 Then
180                       g.Col = z
190                       Printer.Print Tab(15 * z); g.TextMatrix(n, z);
200                   End If
210               Next
220               Printer.Print
230           End If
240       Next

250       Printer.Print
260       Printer.Print "IgG, IgA and IgM Ref Ranges changed on the 20th of November 2006"
270       Printer.Print "Results prior to that use the following Ref Ranges"
280       Printer.Print "IgG (8 - 16), IgA (0.55 - 4), IgM (0.4 - 2.4)"


290       Printer.EndDoc

300       Exit Sub

bprint_Click_Error:

          Dim strES As String
          Dim intEL As Integer



310       intEL = Erl
320       strES = Err.Description
330       LogError "frmFullParaImm", "bPrint_Click", intEL, strES


End Sub

Private Sub cmbPlotFrom_KeyPress(KeyAscii As Integer)

10        On Error GoTo cmbPlotFrom_KeyPress_Error

20        KeyAscii = 0

30        Exit Sub

cmbPlotFrom_KeyPress_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmFullParaImm", "cmbPlotFrom_KeyPress", intEL, strES


End Sub


Private Sub cmbPlotTo_KeyPress(KeyAscii As Integer)

10        On Error GoTo cmbPlotTo_KeyPress_Error

20        KeyAscii = 0

30        Exit Sub

cmbPlotTo_KeyPress_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmFullParaImm", "cmbPlotTo_KeyPress", intEL, strES


End Sub


Private Sub cmdExcel_Click()
          Dim strHeading As String
10        On Error GoTo cmdExcel_Click_Error

20        strHeading = "Immunology Paraprotein History" & vbCr
30        strHeading = strHeading & "Patient Name: " & lblName & vbCr
40        strHeading = strHeading & vbCr

50        ExportFlexGrid g, Me, strHeading

60        Exit Sub

cmdExcel_Click_Error:

          Dim strES As String
          Dim intEL As Integer

70        intEL = Erl
80        strES = Err.Description
90        LogError "frmFullParaImm", "cmdExcel_Click", intEL, strES
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
60        LogError "frmFullParaImm", "cmdGo_Click", intEL, strES


End Sub

Private Sub Form_Activate()

10        On Error GoTo Form_Activate_Error

20        g.Visible = False
          If cmbResultCount.Text <> "" Then
            FillG (Trim(cmbResultCount.Text))
          End If
          g.Visible = True
30        FillCombos

40        PBar.Max = LogOffDelaySecs
50        PBar = 0

60        Timer1.Enabled = True

70        Exit Sub

Form_Activate_Error:

          Dim strES As String
          Dim intEL As Integer



80        intEL = Erl
90        strES = Err.Description
100       LogError "frmFullParaImm", "Form_Activate", intEL, strES


End Sub

Private Function GetRow(ByVal testnum As String) As Long

          Dim n As Long

10        On Error GoTo GetRow_Error

20        For n = 0 To List1.ListCount - 1
30            If testnum = List1.List(n) Then
40                GetRow = n + 4
50                Exit For
60            End If
70        Next

80        Exit Function

GetRow_Error:

          Dim strES As String
          Dim intEL As Integer



90        intEL = Erl
100       strES = Err.Description
110       LogError "frmFullParaImm", "GetRow", intEL, strES

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
150       LogError "frmFullParaImm", "InList", intEL, strES

End Function

Private Sub ChangeListOrder()

          Dim tb As Recordset
          Dim sql As String
          Dim TestCodes As String
          Dim i As Integer

10        On Error GoTo ChangeListOrder_Error

20        If List1.ListCount = 0 Then Exit Sub
30        For i = 0 To List1.ListCount - 1
40            TestCodes = TestCodes & "'" & List1.List(i) & "',"
50        Next i
60        If TestCodes <> "" Then
70            TestCodes = Left(TestCodes, Len(TestCodes) - 1)
80        End If
90        sql = "Select Distinct Code,PrintPriority From ImmTestDefinitions Where Code In (" & TestCodes & ") Order By PrintPriority"
100       Set tb = New Recordset
110       RecOpenClient 0, tb, sql
120       If Not tb.EOF Then
130           List1.Clear
140           While Not tb.EOF
150               List1.AddItem Trim(tb!Code)
160               tb.MoveNext
170           Wend
180       End If


190       Exit Sub

ChangeListOrder_Error:

          Dim strES As String
          Dim intEL As Integer

200       intEL = Erl
210       strES = Err.Description
220       LogError "frmFullParaImm", "ChangeListOrder", intEL, strES, sql

End Sub

Private Sub TransferListToGrid()

          Dim n As Long
          Dim DaysOld As Long
          Dim sql As String
          Dim sn As Recordset

10        On Error GoTo TransferListToGrid_Error

20        If List1.ListCount = 0 Then Exit Sub

30        If lblDoB <> "" Then DaysOld = Abs(DateDiff("d", Now, lblDoB)) Else DaysOld = 0

40        g.Rows = List1.ListCount + 4

50        g.Col = 0
60        For n = 0 To List1.ListCount - 1
70            g.row = n + 4
80            sql = "SELECT * FROM ImmTestDefinitions WHERE " & _
                    "Code = '" & List1.List(n) & "' " & _
                    "AND (AgeFromDays <= " & DaysOld & " " & _
                    "AND AgeToDays >= " & DaysOld & ")"
90            Set sn = New Recordset
100           RecOpenServer Tn, sn, sql
110           If Not sn.EOF Then
120               g = sn!ShortName & ""
130               g.TextMatrix(g.row, 1) = sn!SampleType
140               If sn!PrnRR & "" <> False Then
150                   If Left(sex, 1) = "F" Then
160                       g.TextMatrix(g.row, 2) = Trim(sn!FemaleLow) & " - " & Trim(sn!FemaleHigh)
170                   ElseIf Left(sex, 1) = "M" Then
180                       g.TextMatrix(g.row, 2) = Trim(sn!MaleLow) & " - " & Trim(sn!MaleHigh)
190                   Else
200                       g.TextMatrix(g.row, 2) = Trim(sn!FemaleLow) & " - " & Trim(sn!MaleHigh)
210                   End If
220               End If
230           End If
240       Next

250       Exit Sub

TransferListToGrid_Error:

          Dim strES As String
          Dim intEL As Integer



260       intEL = Erl
270       strES = Err.Description
280       LogError "frmFullParaImm", "TransferListToGrid", intEL, strES, sql

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
60        LogError "frmFullParaImm", "Form_Deactivate", intEL, strES


End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        On Error GoTo Form_MouseMove_Error

20        PBar = 0

30        Exit Sub

Form_MouseMove_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmFullParaImm", "Form_MouseMove", intEL, strES


End Sub

Private Sub g_Click()
          Dim X As Long
          Dim Y As Long

10        On Error GoTo g_Click_Error

20        X = g.RowSel
30        Y = g.ColSel

40        If Y > 0 Then
50            If Trim(g.TextMatrix(X, Y)) <> "" Then g.ToolTipText = g.TextMatrix(X, Y)
60        End If

70        DrawChart

80        Exit Sub

g_Click_Error:

          Dim strES As String
          Dim intEL As Integer



90        intEL = Erl
100       strES = Err.Description
110       LogError "frmFullParaImm", "g_Click", intEL, strES


End Sub


Private Sub g_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        On Error GoTo g_MouseMove_Error

20        Y = g.MouseCol
30        X = g.MouseRow

40        g.ToolTipText = ""

50        If Y > 0 And Not IsNumeric(g.TextMatrix(X, Y)) Then
60            g.ToolTipText = g.TextMatrix(X, Y)
70        End If


80        PBar = 0

90        Exit Sub

g_MouseMove_Error:

          Dim strES As String
          Dim intEL As Integer



100       intEL = Erl
110       strES = Err.Description
120       LogError "frmFullParaImm", "g_MouseMove", intEL, strES


End Sub


Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        On Error GoTo Image1_MouseMove_Error

20        PBar = 0

30        Exit Sub

Image1_MouseMove_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmFullParaImm", "Image1_MouseMove", intEL, strES


End Sub


Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        On Error GoTo Label1_MouseMove_Error

20        PBar = 0

30        Exit Sub

Label1_MouseMove_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmFullParaImm", "Label1_MouseMove", intEL, strES


End Sub


Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        On Error GoTo Label2_MouseMove_Error

20        PBar = 0

30        Exit Sub

Label2_MouseMove_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmFullParaImm", "Label2_MouseMove", intEL, strES


End Sub


Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        On Error GoTo Label3_MouseMove_Error

20        PBar = 0

30        Exit Sub

Label3_MouseMove_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmFullParaImm", "Label3_MouseMove", intEL, strES


End Sub


Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        On Error GoTo Label4_MouseMove_Error

20        PBar = 0

30        Exit Sub

Label4_MouseMove_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmFullParaImm", "Label4_MouseMove", intEL, strES


End Sub


Private Sub lblChart_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        On Error GoTo lblChart_MouseMove_Error

20        PBar = 0

30        Exit Sub

lblChart_MouseMove_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmFullParaImm", "lblChart_MouseMove", intEL, strES


End Sub


Private Sub lblDoB_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        On Error GoTo lblDoB_MouseMove_Error

20        PBar = 0

30        Exit Sub

lblDoB_MouseMove_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmFullParaImm", "lblDoB_MouseMove", intEL, strES


End Sub


Private Sub lblName_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        On Error GoTo lblName_MouseMove_Error

20        PBar = 0

30        Exit Sub

lblName_MouseMove_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmFullParaImm", "lblName_MouseMove", intEL, strES


End Sub


Private Sub pb_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

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
70            CurrentDistance = ((X - ChartPositions(i).xPos) ^ 2 + (Y - ChartPositions(i).yPos) ^ 2) ^ (1 / 2)
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
190       LogError "frmFullParaImm", "pb_MouseMove", intEL, strES


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
90        LogError "frmFullParaImm", "Timer1_Timer", intEL, strES


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
