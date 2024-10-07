VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFullBga 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Full Blood Gas History"
   ClientHeight    =   6195
   ClientLeft      =   75
   ClientTop       =   390
   ClientWidth     =   11820
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
   Icon            =   "frmFullBga.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6195
   ScaleWidth      =   11820
   Begin VB.ComboBox cmbResultCount 
      Height          =   315
      ItemData        =   "frmFullBga.frx":030A
      Left            =   3510
      List            =   "frmFullBga.frx":0320
      TabIndex        =   24
      Text            =   "All"
      Top             =   60
      Width           =   1215
   End
   Begin VB.CheckBox chkChartNumber 
      Caption         =   "Ignore chart number"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   120
      Width           =   2535
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
      Left            =   6960
      TabIndex        =   12
      Top             =   1545
      Width           =   4680
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
         Left            =   2025
         TabIndex        =   15
         Text            =   "cmbPlotTo"
         Top             =   315
         Width           =   1635
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
         TabIndex        =   14
         Text            =   "cmbPlotFrom"
         Top             =   315
         Width           =   1815
      End
      Begin VB.CommandButton cmdGo 
         Caption         =   "Start"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   3780
         Picture         =   "frmFullBga.frx":0352
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   135
         Width           =   825
      End
   End
   Begin MSComctlLib.ProgressBar PBar 
      Height          =   195
      Left            =   6960
      TabIndex        =   11
      Top             =   4935
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   10650
      Top             =   5640
   End
   Begin VB.PictureBox pb 
      AutoRedraw      =   -1  'True
      Height          =   2325
      Left            =   6960
      ScaleHeight     =   2265
      ScaleWidth      =   4125
      TabIndex        =   9
      Top             =   2580
      Width           =   4185
      Begin VB.Label lblTest 
         Alignment       =   2  'Center
         Height          =   195
         Left            =   1485
         TabIndex        =   22
         Top             =   0
         Width           =   1095
      End
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   5625
      Left            =   120
      TabIndex        =   8
      Top             =   420
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   9922
      _Version        =   393216
      Rows            =   4
      FixedRows       =   3
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
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
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   2760
      Left            =   0
      TabIndex        =   7
      Top             =   2310
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.CommandButton bcancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
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
      Height          =   825
      Left            =   8670
      Picture         =   "frmFullBga.frx":065C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5250
      Width           =   1245
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "Show "
      Height          =   195
      Left            =   3030
      TabIndex        =   26
      Top             =   120
      Width           =   450
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "Results"
      Height          =   195
      Left            =   4770
      TabIndex        =   25
      Top             =   120
      Width           =   705
   End
   Begin VB.Label Tn 
      Height          =   225
      Left            =   2730
      TabIndex        =   21
      Top             =   1650
      Width           =   315
   End
   Begin VB.Label sex 
      Caption         =   "Sex"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   10560
      TabIndex        =   20
      Top             =   720
      Width           =   405
   End
   Begin VB.Label lblSex 
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
      Left            =   11040
      TabIndex        =   19
      Top             =   690
      Width           =   510
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
      Left            =   11160
      TabIndex        =   18
      Top             =   4620
      Width           =   510
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
      Left            =   11160
      TabIndex        =   17
      Top             =   3600
      Width           =   510
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
      Left            =   11160
      TabIndex        =   16
      Top             =   2550
      Width           =   510
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   6870
      Picture         =   "frmFullBga.frx":0966
      Top             =   1080
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
      Left            =   7380
      TabIndex        =   10
      Top             =   1170
      Width           =   2865
   End
   Begin VB.Label lblChartTitle 
      Alignment       =   1  'Right Justify
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
      Left            =   6870
      TabIndex        =   6
      Top             =   420
      Width           =   555
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
      Left            =   9240
      TabIndex        =   5
      Top             =   420
      Width           =   315
   End
   Begin VB.Label lblName 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   7440
      TabIndex        =   4
      Top             =   690
      Width           =   3045
   End
   Begin VB.Label lblDoB 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   9900
      TabIndex        =   3
      Top             =   390
      Width           =   1545
   End
   Begin VB.Label Label3 
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
      Left            =   6990
      TabIndex        =   2
      Top             =   720
      Width           =   420
   End
   Begin VB.Label lblChart 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   7440
      TabIndex        =   1
      Top             =   390
      Width           =   1545
   End
End
Attribute VB_Name = "frmFullBga"
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
210       For x = StartGridX To StopGridX
220           If g.TextMatrix(g.Row, x) <> "" Then
230               If Not FirstDayFilled Then
240                   FirstDayFilled = True
250                   MaxVal = Val(g.TextMatrix(g.Row, x))
260                   ChartPositions(NumberOfDays).Date = Format(g.TextMatrix(1, x), "dd/mmm/yyyy")
270                   ChartPositions(NumberOfDays).Value = Val(g.TextMatrix(g.Row, x))
280               Else
290                   DaysInterval = Abs(DateDiff("D", cmbPlotTo, Format$(g.TextMatrix(1, x), "dd/mmm/yyyy")))
300                   ChartPositions(NumberOfDays - DaysInterval).Date = g.TextMatrix(1, x)
310                   cVal = Val(g.TextMatrix(g.Row, x))
320                   ChartPositions(NumberOfDays - DaysInterval).Value = cVal
330                   If cVal > MaxVal Then MaxVal = cVal
340               End If
350           End If
360       Next

370       PixelsPerDay = (pb.Width - 1060) / NumberOfDays
380       MaxVal = MaxVal * 1.1
390       If MaxVal = 0 Then Exit Sub
400       PixelsPerPointY = pb.Height / MaxVal

410       x = 580 + (NumberOfDays * PixelsPerDay)
420       Y = pb.Height - (ChartPositions(NumberOfDays).Value * PixelsPerPointY)
430       ChartPositions(NumberOfDays).yPos = Y
440       ChartPositions(NumberOfDays).xPos = x

450       pb.ForeColor = vbBlue
460       pb.Circle (x, Y), 30
470       pb.Line (x - 15, Y - 15)-(x + 15, Y + 15), vbBlue, BF
480       pb.PSet (x, Y)

490       For n = NumberOfDays - 1 To 0 Step -1
500           If ChartPositions(n).Value <> 0 And ChartPositions(n).Date <> "" Then
510               DaysInterval = Abs(DateDiff("d", cmbPlotFrom, Format(ChartPositions(n).Date, "dd/mmm/yyyy")))
520               x = 580 + (DaysInterval * PixelsPerDay)
530               ChartPositions(n).xPos = x
540               Y = pb.Height - (ChartPositions(n).Value * PixelsPerPointY)
550               ChartPositions(n).yPos = Y
560               pb.Line -(x, Y)
570               pb.Line (x - 15, Y - 15)-(x + 15, Y + 15), vbBlue, BF
580               pb.Circle (x, Y), 30
590               pb.PSet (x, Y)
600           End If
610       Next

620       pb.Line (0, pb.Height / 2)-(pb.Width, pb.Height / 2), vbBlack, BF

630       lblMaxVal = Format$(Int(MaxVal + 0.5), "###0")
640       lblMeanVal = Format$(Val(lblMaxVal) / 2, "###0.0")
650       lblTest = g.TextMatrix(g.RowSel, 0)




660       Exit Sub

DrawChart_Error:

          Dim strES As String
          Dim intEL As Integer



670       intEL = Erl
680       strES = Err.Description
690       LogError "frmFullBga", "DrawChart", intEL, strES


End Sub

Private Sub bcancel_Click()

10        Unload Me

End Sub

Private Sub FillG(p_RCount As String)

          Dim tb As New Recordset
          Dim snr As Recordset
          Dim sn As New Recordset
          Dim sql As String
          Dim x As Long
          Dim xrun As String
          Dim xdate As String
          Dim s As String
          Dim BRs As New BIEResults
          Dim br As BIEResult
          Dim Samp As String



10        On Error GoTo FillG_Error

20        ClearFGrid g
30        p_RCount = UCase(p_RCount)
40        If p_RCount = "First 5" Then
50            sql = "SELECT distinct top 5 (demographics.sampleid), demographics.runDate, demographics.TimeTaken, demographics.sex, bgaresults.runtime, demographics.sampledate from demographics, bgaresults WHERE ("
60        ElseIf p_RCount = "First 10" Then
70            sql = "SELECT distinct top 10 (demographics.sampleid), demographics.runDate, demographics.TimeTaken, demographics.sex, bgaresults.runtime, demographics.sampledate from demographics, bgaresults WHERE ("
80        ElseIf p_RCount = "First 20" Then
90            sql = "SELECT distinct top 20 (demographics.sampleid), demographics.runDate, demographics.TimeTaken, demographics.sex, bgaresults.runtime, demographics.sampledate from demographics, bgaresults WHERE ("
100       ElseIf p_RCount = "First 50" Then
110           sql = "SELECT distinct top 50 (demographics.sampleid), demographics.runDate, demographics.TimeTaken, demographics.sex, bgaresults.runtime, demographics.sampledate from demographics, bgaresults WHERE ("
120       ElseIf p_RCount = "FULL" Then
130           sql = "SELECT distinct(demographics.sampleid), demographics.runDate, demographics.TimeTaken, demographics.sex, bgaresults.runtime, demographics.sampledate from demographics, bgaresults WHERE ("
140       End If

150       If Trim(lblChart) <> "" And chkChartNumber.Value = 0 Then
160           sql = sql & "(demographics.chart = '" & lblChart & "') or"
170       End If
180       sql = sql & " (demographics.PatName = '" & AddTicks(lblName) & "' and demographics.dob  = '" & Format(lblDoB, "dd/MMM/yyyy") & "') "

190       sql = sql & ") and demographics.sampleid =  bgaresults.sampleid " & _
              "order by demographics.sampledate desc"

200       Set tb = New Recordset
210       RecOpenServer Tn, tb, sql
220       If Not tb.EOF Then
230           sex = tb!sex & ""
240           g.Cols = 3
250           g.ColWidth(0) = 1795
260           g.ColWidth(1) = 495
270           g.ColWidth(2) = 1395
280           g.TextMatrix(0, 0) = "SAMPLE ID"
290           g.TextMatrix(1, 0) = "SAMPLE DATE"
300           g.TextMatrix(2, 0) = "RUN TIME"
310           g.TextMatrix(2, 1) = "S/T"
320           g.TextMatrix(2, 2) = "Ref Ranges"
              'SampleID and sampledate across
330           Do While Not tb.EOF
340               g.Cols = g.Cols + 1
350               x = g.Cols - 1
360               g.ColWidth(x) = 1095
370               g.Col = x
380               xrun = tb!SampleID & ""
390               g.TextMatrix(0, x) = xrun
400               If Not IsNull(tb!SampleDate) Then
410                   xdate = Format$(tb!SampleDate, "dd/mm/yy")
420               Else
430                   xdate = ""
440               End If
450               g.TextMatrix(1, x) = xdate
460               If IsDate(tb!SampleDate) Then
470                   g.TextMatrix(2, x) = Format(tb!SampleDate, "hh:mm")
480               Else
490                   g.TextMatrix(2, x) = ""
500               End If
                  'fill list with test names
510               If xrun <> "" Then
520                   sql = "SELECT BgaResults.*, PrintPriority " & _
                          "from BgaResults, BgaTestDefinitions " & _
                          "WHERE SampleID = '" & xrun & "' " & _
                          "and BgaTestDefinitions.Code = BgaResults.Code " & _
                          "and BgaTestDefinitions.Hospital = '" & HospName(Tn) & "' " & _
                          "order by PrintPriority"
530                   Set snr = New Recordset
540                   RecOpenServer Tn, snr, sql
550                   Do While Not snr.EOF
560                       If Trim(snr!SampleType) & "" = "" Then Samp = "S" Else Samp = snr!SampleType
570                       If Not InList(snr!Code) Then
580                           List1.AddItem snr!Code & " " & Left(snr!SampleType, 1)
590                       End If
600                       snr.MoveNext
610                   Loop
620                   snr.Close
630                   Set snr = Nothing
640               End If
650               tb.MoveNext
660           Loop



670           If List1.ListCount = 0 Then Exit Sub
              'fill first col with test names
680           TransferListToGrid

              'fill in results
690           For x = 3 To g.Cols - 1
700               g.Col = x
710               g.Row = 1
720               xdate = Format$(g, "dd/mmm/yyyy")
730               g.Row = 0
740               xrun = g
750               If xrun <> "" Then
760                   If UserMemberOf = "Secretarys" Then
770                       Set BRs = BRs.Load("Bga", xrun, "Results", gVALID, gDONTCARE, Tn, "", xdate)
780                   Else
790                       Set BRs = BRs.Load("Bga", xrun, "Results", gDONTCARE, gDONTCARE, Tn, "", xdate)
800                   End If
810                   For Each br In BRs
820                       g.Row = GetRow(br.Code)
830                       If g.Row > 0 Then
                              'Code added 22/08/05
840                           If g.TextMatrix(2, x) = "" Then g.TextMatrix(2, x) = Format(br.RunTime, "hh:mm")
850                           If (UserMemberOf = "Secretarys" Or SysOptNoCumShow(0)) And br.Valid <> True Then
860                               g = "NV"
870                           Else
880                               sql = "SELECT * from bgatestdefinitions WHERE code = '" & br.Code & "'"
890                               If br.SampleType <> "" Then sql = sql & "  and sampletype = '" & br.SampleType & "'"
900                               Set sn = New Recordset
910                               RecOpenServer Tn, sn, sql
920                               Select Case sn!DP
                                      Case 0: g = Format$(br.Result, "0")
930                                   Case 1: g = Format$(br.Result, "0.0")
940                                   Case 2: g = Format$(br.Result, "0.00")
950                                   Case 3: g = Format$(br.Result, "0.000")
960                                   Case 4: g = Format$(br.Result, "0.0000")
970                               End Select
980                           End If
990                           s = QuickInterpBio(br)
1000                          If s = "Low " Then
1010                              g.CellBackColor = vbBlue
1020                              g.CellForeColor = vbYellow
1030                          ElseIf s = "High" Then
1040                              g.CellBackColor = vbRed
1050                              g.CellForeColor = vbYellow
1060                          End If
1070                      End If

1080                  Next
1090              End If
1100          Next
1110      End If

1120      FixG g


1130      Me.Refresh



1140      Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer

1150      intEL = Erl
1160      strES = Err.Description
1170      LogError "frmFullBga", "FillG", intEL, strES, sql


End Sub

Private Sub chkChartNumber_Click()
10        g.Visible = False
          If cmbResultCount.Text <> "" Then
             FillG (Trim(cmbResultCount.Text))
          End If
          g.Visible = True
20        FillCombos
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

Private Sub cmdGo_Click()

10        On Error GoTo cmdGo_Click_Error

20        DrawChart

30        Exit Sub

cmdGo_Click_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmFullBga", "cmdGo_Click", intEL, strES


End Sub

Private Sub Form_Activate()

10        On Error GoTo Form_Activate_Error
          
20        If GetOptionSetting("EnableiPMSChart", "0") = 0 Then
30            chkChartNumber.Value = 0
40            chkChartNumber.Enabled = Not (lblChart = "")
50        Else
60            chkChartNumber.Value = 1

70        End If
80        g.Visible = False
          If cmbResultCount.Text <> "" Then
             FillG (Trim(cmbResultCount.Text))
          End If
          g.Visible = True
90        FillCombos



100       PBar.Max = LogOffDelaySecs
110       PBar = 0

120       Timer1.Enabled = True

130       Exit Sub

Form_Activate_Error:

          Dim strES As String
          Dim intEL As Integer



140       intEL = Erl
150       strES = Err.Description
160       LogError "frmFullBga", "Form_Activate", intEL, strES


End Sub

Private Sub Form_Deactivate()

10        Timer1.Enabled = False

End Sub


Private Sub Form_Load()

10        On Error GoTo Form_Load_Error

20        PBar.Max = LogOffDelaySecs
30        PBar = 0

40        Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer



50        intEL = Erl
60        strES = Err.Description
70        LogError "frmFullBga", "Form_Load", intEL, strES


End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

10        On Error GoTo Form_MouseMove_Error

20        PBar = 0

30        Exit Sub

Form_MouseMove_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmFullBga", "Form_MouseMove", intEL, strES


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
60        LogError "frmFullBga", "g_Click", intEL, strES


End Sub


Private Sub g_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

10        On Error GoTo g_MouseMove_Error

20        PBar = 0

30        Exit Sub

g_MouseMove_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmFullBga", "g_MouseMove", intEL, strES


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
190       LogError "frmFullBga", "pb_MouseMove", intEL, strES


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

90        For x = g.Cols - 1 To 1 Step -1
100           If DateDiff("d", Format$(g.TextMatrix(1, x), "dd/mmm/yyyy"), cmbPlotTo) < 365 Then
110               cmbPlotFrom = Format$(g.TextMatrix(1, x), "dd/mmm/yyyy")
120               Exit For
130           End If
140       Next



150       Exit Sub

FillCombos_Error:

          Dim strES As String
          Dim intEL As Integer



160       intEL = Erl
170       strES = Err.Description
180       LogError "frmFullBga", "FillCombos", intEL, strES


End Sub


Private Sub Timer1_Timer()

      'tmrRefresh.Interval set to 1000
10        On Error GoTo Timer1_Timer_Error

20        PBar = PBar + 1


30        Exit Sub

Timer1_Timer_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmFullBga", "Timer1_Timer", intEL, strES


End Sub

Private Sub TransferListToGrid()

          Dim n As Long
          Dim sql As String
          Dim sn As New Recordset
          Dim DaysOld As Long

10        On Error GoTo TransferListToGrid_Error

20        If List1.ListCount = 0 Then Exit Sub

30        If lblDoB <> "" Then DaysOld = Abs(DateDiff("d", Now, lblDoB)) Else DaysOld = 0

40        g.Rows = List1.ListCount + 3

50        g.Col = 0
60        For n = 0 To List1.ListCount - 1
70            g.Row = n + 3
80            sql = "SELECT * from bgatestdefinitions WHERE  code = '" & Mid((List1.List(n)), 1, Len(List1.List(n)) - 2) & "' and sampletype = '" & Right(List1.List(n), 1) & "' and  (agefromdays <= " & DaysOld & " and agetodays >= " & DaysOld & ")"
90            Set sn = New Recordset
100           RecOpenServer Tn, sn, sql
110           If Not sn.EOF Then
120               g = sn!LongName & ""
130               g.TextMatrix(g.Row, 1) = Right(List1.List(n), 1)
140               If Left(sex, 1) = "F" Then
150                   g.TextMatrix(g.Row, 2) = Trim(sn!FemaleLow) & " - " & Trim(sn!FemaleHigh)
160               ElseIf Left(sex, 1) = "M" Then
170                   g.TextMatrix(g.Row, 2) = Trim(sn!MaleLow) & " - " & Trim(sn!MaleHigh)
180               Else
190                   g.TextMatrix(g.Row, 2) = Trim(sn!FemaleLow) & " - " & Trim(sn!MaleHigh)
200               End If
210           End If
220       Next

230       Exit Sub

TransferListToGrid_Error:

          Dim strES As String
          Dim intEL As Integer



240       intEL = Erl
250       strES = Err.Description
260       LogError "frmFullBga", "TransferListToGrid", intEL, strES


End Sub

Private Function GetRow(ByVal testnum As String) As Long

          Dim n As Long



10        On Error GoTo GetRow_Error

20        For n = 0 To List1.ListCount - 1
30            If testnum = Mid(List1.List(n), 1, Len(List1.List(n)) - 2) Then
40                GetRow = n + 3
50                Exit For
60            End If
70        Next

80        n = GetRow


90        Exit Function

GetRow_Error:

          Dim strES As String
          Dim intEL As Integer



100       intEL = Erl
110       strES = Err.Description
120       LogError "frmFullBga", "GetRow", intEL, strES


End Function

Private Function InList(ByVal s As String) As Long

          Dim n As Long

10        InList = False
20        If List1.ListCount = 0 Then
30            Exit Function
40        End If

50        For n = 0 To List1.ListCount - 1
60            If s = Mid(List1.List(n), 1, Len(List1.List(n)) - 2) Then
70                InList = True
80                Exit For
90            End If
100       Next

End Function

