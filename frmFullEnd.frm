VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFullEnd 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Full Endocrinology History"
   ClientHeight    =   7185
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
   Icon            =   "frmFullEnd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7185
   ScaleWidth      =   11565
   Begin VB.ComboBox cmbResultCount 
      Height          =   315
      ItemData        =   "frmFullEnd.frx":030A
      Left            =   540
      List            =   "frmFullEnd.frx":0320
      TabIndex        =   34
      Text            =   "All"
      Top             =   30
      Width           =   1215
   End
   Begin VB.Frame FrameRange 
      Caption         =   "DateRange"
      Height          =   870
      Left            =   6705
      TabIndex        =   28
      Top             =   1080
      Width           =   4785
      Begin VB.CommandButton cmdRefresh 
         Height          =   615
         Left            =   3870
         Picture         =   "frmFullEnd.frx":0352
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Click to refresh biochemistry history"
         Top             =   180
         Width           =   735
      End
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   330
         Left            =   135
         TabIndex        =   30
         Top             =   450
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MMM/yyyy"
         Format          =   228392963
         CurrentDate     =   38629
      End
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   330
         Left            =   2070
         TabIndex        =   31
         Top             =   450
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MMM/yyyy"
         Format          =   228392963
         CurrentDate     =   38629
      End
      Begin VB.Label Label6 
         Caption         =   "to"
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
         Left            =   2655
         TabIndex        =   33
         Top             =   165
         Width           =   330
      End
      Begin VB.Label lblBetween 
         Caption         =   "Between"
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
         Left            =   495
         TabIndex        =   32
         Top             =   180
         Width           =   960
      End
   End
   Begin VB.CheckBox chkChartNumber 
      Caption         =   "Ignore chart number"
      Height          =   255
      Left            =   6840
      TabIndex        =   26
      Top             =   780
      Width           =   2535
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
      Height          =   720
      Left            =   10215
      Picture         =   "frmFullEnd.frx":0C1C
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   6255
      Width           =   1245
   End
   Begin VB.CommandButton bPrint 
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
      Left            =   8880
      Picture         =   "frmFullEnd.frx":0F26
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   6225
      Width           =   1245
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
      Left            =   7500
      Picture         =   "frmFullEnd.frx":1230
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   6225
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
      Height          =   885
      Left            =   6780
      TabIndex        =   10
      Top             =   2535
      Width           =   4725
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
         Height          =   675
         Left            =   3600
         Picture         =   "frmFullEnd.frx":6E4E
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   135
         Width           =   1050
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
         Left            =   180
         TabIndex        =   12
         Text            =   "cmbPlotFrom"
         Top             =   240
         Width           =   1635
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
         TabIndex        =   11
         Text            =   "cmbPlotTo"
         Top             =   240
         Width           =   1635
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   10350
      Top             =   6060
   End
   Begin VB.PictureBox pb 
      AutoRedraw      =   -1  'True
      Height          =   2325
      Left            =   6780
      ScaleHeight     =   2265
      ScaleWidth      =   4125
      TabIndex        =   7
      Top             =   3495
      Width           =   4185
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   2760
      Left            =   1980
      TabIndex        =   6
      Top             =   2025
      Visible         =   0   'False
      Width           =   1635
   End
   Begin MSComctlLib.ProgressBar PBar 
      Height          =   195
      Left            =   45
      TabIndex        =   9
      Top             =   7020
      Width           =   11385
      _ExtentX        =   20082
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   6615
      Left            =   45
      TabIndex        =   25
      Top             =   345
      Width           =   6585
      _ExtentX        =   11615
      _ExtentY        =   11668
      _Version        =   393216
      Rows            =   5
      FixedRows       =   4
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      ScrollTrack     =   -1  'True
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "Show "
      Height          =   195
      Left            =   60
      TabIndex        =   36
      Top             =   90
      Width           =   450
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "Results"
      Height          =   195
      Left            =   1800
      TabIndex        =   35
      Top             =   90
      Width           =   705
   End
   Begin VB.Label Lbl1 
      Caption         =   "Amended  Results are underline"
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
      Left            =   4005
      TabIndex        =   27
      Top             =   90
      Width           =   2535
   End
   Begin VB.Label lblExcelInfo 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exporting..."
      Height          =   285
      Left            =   7500
      TabIndex        =   22
      Top             =   5970
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
      Left            =   10290
      TabIndex        =   19
      Top             =   5910
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
      Left            =   8895
      TabIndex        =   20
      Top             =   5910
      Width           =   1365
   End
   Begin VB.Label lblNopas 
      Height          =   255
      Left            =   1560
      TabIndex        =   18
      Top             =   1320
      Width           =   1365
   End
   Begin VB.Label Tn 
      Height          =   225
      Left            =   3840
      TabIndex        =   17
      Top             =   1410
      Width           =   315
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   6720
      Picture         =   "frmFullEnd.frx":7158
      Top             =   2025
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
      TabIndex        =   16
      Top             =   3180
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
      TabIndex        =   15
      Top             =   4230
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
      TabIndex        =   14
      Top             =   5250
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
      TabIndex        =   8
      Top             =   2115
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
      Top             =   450
      Width           =   3765
   End
   Begin VB.Label lblDoB 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   9450
      TabIndex        =   2
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
      Top             =   150
      Width           =   1545
   End
End
Attribute VB_Name = "frmFullEnd"
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
80        If NumberOfDays < 3 Then Exit Sub

90        ReDim ChartPositions(0 To NumberOfDays)

100       For n = 1 To NumberOfDays
110           ChartPositions(n).xPos = 0
120           ChartPositions(n).yPos = 0
130           ChartPositions(n).Value = 0
140           ChartPositions(n).Date = ""
150       Next

160       For n = 1 To g.Cols - 1
170           If Format$(g.TextMatrix(1, n), "dd/mmm/yyyy") = cmbPlotTo Then StartGridX = n
180           If Format$(g.TextMatrix(1, n), "dd/mmm/yyyy") = cmbPlotFrom Then StopGridX = n
190       Next

200       FirstDayFilled = False
210       Counter = 0
220       For x = StartGridX To StopGridX
230           If g.TextMatrix(g.Row, x) <> "" Then
240               If Not FirstDayFilled Then
250                   FirstDayFilled = True
260                   MaxVal = Val(g.TextMatrix(g.Row, x))
270                   ChartPositions(NumberOfDays).Date = Format(g.TextMatrix(1, x), "dd/mmm/yyyy")
280                   ChartPositions(NumberOfDays).Value = Val(g.TextMatrix(g.Row, x))
290               Else
300                   DaysInterval = Abs(DateDiff("D", cmbPlotTo, Format$(g.TextMatrix(1, x), "dd/mmm/yyyy")))
310                   ChartPositions(NumberOfDays - DaysInterval).Date = g.TextMatrix(1, x)
320                   cVal = Val(g.TextMatrix(g.Row, x))
330                   ChartPositions(NumberOfDays - DaysInterval).Value = cVal
340                   If cVal > MaxVal Then MaxVal = cVal
350               End If
360           End If
370       Next

380       PixelsPerDay = (pb.Width - 1060) / NumberOfDays
390       MaxVal = MaxVal * 1.1
400       If MaxVal = 0 Then Exit Sub
410       PixelsPerPointY = pb.Height / MaxVal

420       x = 580 + (NumberOfDays * PixelsPerDay)
430       Y = pb.Height - (ChartPositions(NumberOfDays).Value * PixelsPerPointY)
440       ChartPositions(NumberOfDays).yPos = Y
450       ChartPositions(NumberOfDays).xPos = x

460       pb.ForeColor = vbBlue
470       pb.Circle (x, Y), 30
480       pb.Line (x - 15, Y - 15)-(x + 15, Y + 15), vbBlue, BF
490       pb.PSet (x, Y)

500       For n = NumberOfDays - 1 To 0 Step -1
510           If ChartPositions(n).Value <> 0 And ChartPositions(n).Date <> "" Then
520               DaysInterval = Abs(DateDiff("d", cmbPlotFrom, Format(ChartPositions(n).Date, "dd/mmm/yyyy")))
530               x = 580 + (DaysInterval * PixelsPerDay)
540               ChartPositions(n).xPos = x
550               Y = pb.Height - (ChartPositions(n).Value * PixelsPerPointY)
560               ChartPositions(n).yPos = Y
570               pb.Line -(x, Y)
580               pb.Line (x - 15, Y - 15)-(x + 15, Y + 15), vbBlue, BF
590               pb.Circle (x, Y), 30
600               pb.PSet (x, Y)
610           End If
620       Next

630       pb.Line (0, pb.Height / 2)-(pb.Width, pb.Height / 2), vbBlack, BF

640       lblMaxVal = Format$(Int(MaxVal + 0.5), "###0")
650       lblMeanVal = Format$(Val(lblMaxVal) / 2, "###0.0")



660       Exit Sub

DrawChart_Error:

          Dim strES As String
          Dim intEL As Integer

670       intEL = Erl
680       strES = Err.Description
690       LogError "frmFullEnd", "DrawChart", intEL, strES


End Sub
Private Sub FillCombos()

          Dim x As Long


10        On Error GoTo FillCombos_Error

20        cmbPlotFrom.Clear
30        cmbPlotTo.Clear

40        If g.Cols = 2 Then Exit Sub

50        For x = 1 To g.Cols - 1
60            cmbPlotFrom.AddItem Format$(g.TextMatrix(1, x), "dd/mmm/yyyy")
70            cmbPlotTo.AddItem Format$(g.TextMatrix(1, x), "dd/mmm/yyyy")
80        Next

90        cmbPlotTo = Format$(g.TextMatrix(1, 3), "dd/mmm/yyyy")

100       For x = g.Cols - 1 To 3 Step -1
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
190       LogError "frmFullEnd", "FillCombos", intEL, strES


End Sub

Private Sub bcancel_Click()

10        Unload Me

End Sub
Private Sub FillG(p_RCount As String)

          Dim tb As New Recordset
          Dim snr As Recordset
          Dim sql As String
          Dim x As Long
          Dim xrun As String
          Dim xdate As String
          Dim sn As New Recordset
          Dim BRs As New BIEResults
          Dim br As BIEResult
          Dim s As String
          Dim Cat As String


10        On Error GoTo FillG_Error
20        p_RCount = UCase(p_RCount)
30        If p_RCount = "First 5" Then
40            sql = "SELECT DISTINCT top 5 (D.SampleID), D.Sex, D.RunDate, D.TimeTaken, D.Category, D.SampleDate, " & _
                  "(SELECT Top 1 Q.RunTime FROM EndResults Q WHERE Q.SampleID = D.SampleID ORDER BY Q.RunTime) RunDateTime " & _
                  "FROM Demographics D, EndResults R WHERE ("
                
50        ElseIf p_RCount = "First 10" Then
60            sql = "SELECT DISTINCT top 10 (D.SampleID), D.Sex, D.RunDate, D.TimeTaken, D.Category, D.SampleDate, " & _
                  "(SELECT Top 1 Q.RunTime FROM EndResults Q WHERE Q.SampleID = D.SampleID ORDER BY Q.RunTime) RunDateTime " & _
                  "FROM Demographics D, EndResults R WHERE ("

70        ElseIf p_RCount = "First 20" Then
80            sql = "SELECT DISTINCT top 20 (D.SampleID), D.Sex, D.RunDate, D.TimeTaken, D.Category, D.SampleDate, " & _
                  "(SELECT Top 1 Q.RunTime FROM EndResults Q WHERE Q.SampleID = D.SampleID ORDER BY Q.RunTime) RunDateTime " & _
                  "FROM Demographics D, EndResults R WHERE ("

90        ElseIf p_RCount = "First 50" Then
100           sql = "SELECT DISTINCT top 50 (D.SampleID), D.Sex, D.RunDate, D.TimeTaken, D.Category, D.SampleDate, " & _
                  "(SELECT Top 1 Q.RunTime FROM EndResults Q WHERE Q.SampleID = D.SampleID ORDER BY Q.RunTime) RunDateTime " & _
                  "FROM Demographics D, EndResults R WHERE ("

110       ElseIf p_RCount = "ALL" Then
120           sql = "SELECT DISTINCT (D.SampleID), D.Sex, D.RunDate, D.TimeTaken, D.Category, D.SampleDate, " & _
                  "(SELECT Top 1 Q.RunTime FROM EndResults Q WHERE Q.SampleID = D.SampleID ORDER BY Q.RunTime) RunDateTime " & _
                  "FROM Demographics D, EndResults R WHERE ("

130       End If

140       If Trim(lblChart) <> "" And chkChartNumber.Value = 0 Then
150           sql = sql & "(D.Chart = '" & lblChart & "') and"
160       End If
170       sql = sql & " (D.PatName = '" & AddTicks(lblName) & "' AND D.DoB  = '" & Format(lblDoB, "dd/MMM/yyyy") & "') "
180       sql = sql & " AND D.RunDate BETWEEN '" & Format(dtFrom, "dd/MMM/yyyy") & "' AND '" & Format(dtTo + 1, "dd/MMM/yyyy") & "' "
190       sql = sql & ") AND D.SampleID = R.SampleID "
          '+++ Junaid 10-08-2023
          'sql = sql & "ORDER BY RunDateTime desc"
200       sql = sql & "ORDER BY D.SampleDate desc"
          '--- Junaid
210       Set tb = New Recordset
220       RecOpenServer 0, tb, sql


230       g.Cols = 3
240       g.ColWidth(0) = 1600
250       g.ColWidth(1) = 600
260       g.ColWidth(2) = 1400
270       g.TextMatrix(0, 0) = "SAMPLE ID"
280       g.TextMatrix(1, 0) = "SAMPLE DATE"
290       g.TextMatrix(2, 0) = "RUN DATE"
300       g.TextMatrix(2, 1) = "S/T"
310       g.TextMatrix(2, 2) = "Ref Ranges"
320       sex = tb!sex & ""


330       If Not tb.EOF Then
340           g.Visible = False

350           sex = tb!sex & ""

360           Do While Not tb.EOF

                  'SampleID and sampledate across
370               g.Cols = g.Cols + 1
380               x = g.Cols - 1
390               g.ColWidth(x) = 1500
400               g.Col = x
410               xrun = tb!SampleID & ""
420               g.TextMatrix(0, x) = xrun
                  'Sample DateTime
430               If Not IsNull(tb!SampleDate) Then
440                   xdate = Format(tb!SampleDate, "dd/mm/yy")
450                   If Format(tb!SampleDate, "hh:mm") <> "00:00" Then
460                       xdate = xdate & " " & Format(tb!SampleDate, "hh:mm")
470                   End If
480               Else
490                   xdate = ""
500               End If
510               g.TextMatrix(1, x) = xdate
                  'Run DateTime

520               If tb!RunDateTime <> "" Then
530                   xdate = Format(tb!RunDateTime, "dd/mm/yy")
540                   If Format(tb!RunDateTime, "hh:mm") <> "00:00" Then
550                       xdate = xdate & " " & Format(tb!RunDateTime, "hh:mm")
560                   End If
570               Else
580                   xdate = ""
590               End If
600               g.TextMatrix(2, x) = xdate
                  '        If Not IsNull(tb!SampleDate) Then
                  '            xdate = Format$(tb!SampleDate, "dd/mm/yy")
                  '        Else
                  '            xdate = ""
                  '        End If
                  '        g.TextMatrix(1, X) = xdate
                  '        If IsDate(tb!SampleDate) Then g.TextMatrix(2, X) = Format$(tb!SampleDate, "hh:mm")
610               If Trim(tb!Category) & "" = "" Then g.TextMatrix(3, x) = "Default" Else g.TextMatrix(3, x) = tb!Category
                  'fill list with test names
620               sql = "SELECT endResults.*, PrintPriority " & _
                      "from endResults, endTestDefinitions " & _
                      "WHERE SampleID = '" & xrun & "' " & _
                      "and endTestDefinitions.Code = endResults.Code "
630               If UCase(HospName(0)) <> UCase("Portlaoise") Then
640                   sql = sql & " and endTestDefinitions.Hospital = '" & HospName(0) & "' "
650               End If
660               sql = sql & " order by PrintPriority"

670               Set snr = New Recordset
680               RecOpenServer 0, snr, sql
690               Do While Not snr.EOF
700                   If Not InList(snr!Code) Then
710                       List1.AddItem snr!Code

720                   End If
730                   snr.MoveNext
740               Loop
750               tb.MoveNext
760           Loop
770           If List1.ListCount = 0 Then Exit Sub
              'fill first col with test names
780           TransferListToGrid

              'fill in results
790           For x = 3 To g.Cols - 1
800               g.Col = x
810               g.Row = 1
820               xdate = Format$(g, "dd/mmm/yyyy")
830               g.Row = 0
840               xrun = g
850               g.Row = 3
860               Cat = g

870               If xrun <> "" Then
880                   If UserMemberOf = "Secretarys" Then
890                       Set BRs = BRs.Load("end", xrun, "Results", gVALID, gDONTCARE, Tn, Cat, xdate)
900                   Else
910                       Set BRs = BRs.Load("end", xrun, "Results", gDONTCARE, gDONTCARE, Tn, Cat, xdate)
920                   End If
930                   If BRs.Count = 0 Then g.TextMatrix(4, x) = "Error"
940                   For Each br In BRs
950                       g.Row = GetRow(br.Code)
960                       If g.Row > 0 Then
970                           If g.TextMatrix(2, x) = "" Then g.TextMatrix(2, x) = Format(br.RunTime, "hh:mm")

980                           If (UserMemberOf = "Secretarys" Or SysOptNoCumShow(0)) And br.Valid <> True Then
990                               g = "NV"
1000                          Else

1010                              If UCase(br.Analyser) = "VIROLOGY" Then
1020                                  br.Result = TranslateEndResultVirology(br.Code, br.Result)
1030                                  g = br.Result
1040                                  g.TextMatrix(g.Row, 0) = br.ShortName
1050                              Else
1060                                  sql = "SELECT * from endtestdefinitions WHERE code = '" & br.Code & "'"
1070                                  If br.SampleType <> "" Then sql = sql & "  and sampletype = '" & br.SampleType & "'"
1080                                  Set sn = New Recordset
1090                                  RecOpenServer Tn, sn, sql
1100                                  Select Case sn!DP
                                          Case 0: g = Format$(br.Result, "0")
1110                                      Case 1: g = Format$(br.Result, "0.0")
1120                                      Case 2: g = Format$(br.Result, "0.00")
1130                                      Case 3: g = Format$(br.Result, "0.000")
1140                                      Case 4: g = Format$(br.Result, "0.0000")
1150                                  End Select
1160                              End If
1170                              If Not br.Valid Then g = g & "(NV)"
1180                          End If
                              '----------------------
1190                          If IsResultAmended("End", g.TextMatrix(0, x), br.Code, br.Result) = True Then
1200                              g.CellFontUnderline = True
1210                          End If
                              '======================
1220                          s = QuickInterpEnd(br)
1230                          If UCase(br.Analyser) <> "VIROLOGY" Then
1240                              If Trim(s) = "H" Then
1250                                  g.CellForeColor = SysOptHighFore(0)
1260                                  g.CellBackColor = SysOptHighBack(0)
1270                              ElseIf Trim(s) = "L" Then
1280                                  g.CellForeColor = SysOptLowFore(0)
1290                                  g.CellBackColor = SysOptLowBack(0)
1300                              Else
1310                                  g.CellForeColor = vbBlack
1320                                  g.CellBackColor = vbWhite
1330                              End If
1340                          End If
1350                      End If
1360                  Next
1370              End If
1380          Next
1390      End If

1400      If g.Cols > 2 Then lblNoRes.Caption = g.Cols - 3 Else lblNoRes.Caption = 0

1410      g.Visible = True



1420      Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer

1430      intEL = Erl
1440      strES = Err.Description
1450      LogError "frmFullEnd", "FillG", intEL, strES, sql


End Sub


Private Sub bprint_Click()

      'Dim x As Long
      'Dim n As Long
      'Dim z As Long

          Dim Px As Printer
          Dim x As Integer
          Dim Y As Integer
          Dim z As Integer
          Dim PageCounter As Integer
          Dim CurrentPage As Integer
          Dim TabPos As Integer
          Dim ArrayStart As Integer
          Dim ArrayStop As Integer
          Dim MaxCols As Integer
          Dim TotalLines As Integer
          Dim LinesToPrint As Integer
          Dim Start As Integer
          Dim Last As Integer
          Dim cWidth As Integer

10        On Error GoTo bprint_Click_Error



20        Printer.Font = "Courier New"
30        Printer.Orientation = vbPRORLandscape


40        TotalLines = 27
50        LinesToPrint = g.Rows - 4
60        Start = 4
70        PageCounter = ((g.Rows - 4) \ TotalLines)
80        If (g.Rows - 4) Mod TotalLines > 0 Then PageCounter = PageCounter + 1
90        CurrentPage = 1

100       If LinesToPrint > TotalLines Then
110           Last = TotalLines
120           LinesToPrint = LinesToPrint - TotalLines
130       Else
140           Last = LinesToPrint + 3
150           LinesToPrint = 0
160       End If

170       Do While CurrentPage <= PageCounter
180           Printer.Print
190           PrintText FormatString("Cumulative Endocrinology Report", 70, , AlignCenter), 14, True, , , , True
200           PrintText FormatString("Page " & Format$(CurrentPage) & " of " & PageCounter, 100, , AlignCenter), 10, , , , , True

210           PrintText "  Patient Name: " & lblName, 14, True, , , , True
220           PrintText " Date of Birth: " & Format$(lblDoB, "dd/mm/yyyy"), 14, True, , , , True

230           Printer.Print

240           With g
250               If .Cols > 8 Then
260                   MaxCols = 7
270               Else
280                   MaxCols = .Cols - 1
290               End If

                  'Add seperator
300               PrintText String(217, "-") & vbCrLf, 4, True
                  'Print SampleID row
310               PrintText FormatString(.TextMatrix(0, 0), 16, "|"), 10, True
320               For z = 1 To MaxCols
330                   PrintText FormatString(.TextMatrix(0, z), 9, "|", AlignCenter), 10
340               Next z
350               PrintText vbCrLf
                  'Print Sample Date Row
360               PrintText FormatString(.TextMatrix(1, 0), 16, "|"), 10, True
370               For z = 1 To MaxCols
380                   PrintText FormatString(Format(.TextMatrix(1, z), "dd/MM/yy"), 9, "|", AlignCenter), 10
390               Next z
400               PrintText vbCrLf
                  'Print Sample Time Row

410               PrintText FormatString("SAMPLE TIME", 16, "|"), 10, True

420               For z = 1 To MaxCols
430                   If Format(.TextMatrix(1, z), "hh:mm") = "00:00" Then
440                       PrintText FormatString("", 9, "|", AlignCenter), 10
450                   Else
460                       PrintText FormatString(Format(.TextMatrix(1, z), "hh:mm"), 9, "|", AlignCenter), 10
470                   End If
480               Next z
490               PrintText vbCrLf

                  'Print Run Date Row
500               PrintText FormatString(.TextMatrix(2, 0), 16, "|"), 10, True
510               PrintText FormatString("", 9, "|"), 10
520               PrintText FormatString("", 9, "|"), 10
530               For z = 3 To MaxCols
540                   PrintText FormatString(Format(.TextMatrix(2, z), "dd/MM/yy"), 9, "|", AlignCenter), 10
550               Next z
560               PrintText vbCrLf
                  'Print Run Time Row is exists

570               PrintText FormatString("RUN TIME", 16, "|"), 10, True
580               For z = 1 To MaxCols
590                   If Format(.TextMatrix(2, z), "hh:mm") = "00:00" Then
600                       PrintText FormatString("", 9, "|", AlignCenter), 10
610                   Else
620                       PrintText FormatString(Format(.TextMatrix(2, z), "hh:mm"), 9, "|", AlignCenter), 10
630                   End If
640               Next z
650               PrintText vbCrLf
                  'Add seperator
660               PrintText String(217, "-") & vbCrLf, 4, True
                  'Print results
670               For Y = Start To Last

680                   PrintText FormatString(.TextMatrix(Y, 0), 16, "|"), 10
690                   For z = 1 To MaxCols
700                       PrintText FormatString(.TextMatrix(Y, z), 9, "|", IIf(z = 2, AlignLeft, AlignCenter)), 10
710                   Next z
720                   PrintText vbCrLf
730               Next Y

740           End With


              'End of Page Line
750           PrintText String(217, "-"), 4, True
760           If CurrentPage < PageCounter Then Printer.NewPage
770           CurrentPage = CurrentPage + 1
780           Start = Last + 1
790           If LinesToPrint > TotalLines Then
800               Last = Last + TotalLines
810               LinesToPrint = LinesToPrint - TotalLines
820           Else
830               Last = Last + LinesToPrint + 3
840               LinesToPrint = 0
850           End If

860       Loop

870       Printer.EndDoc



          'x = g.Cols
          '
          '
          'Printer.Orientation = vbPRORLandscape
          '
          'Printer.Font.Size = 16
          'Printer.Print Tab(15); "Cumulative Report from Endocrinology Dept."
          'Printer.Print
          '
          'Printer.Font.Size = 14
          'Printer.Print Tab(10); "Name : " & lblName;
          'Printer.Print Tab(40); "Dob  : " & lblDoB
          '
          '
          'Printer.Print
          'Printer.Print
          '
          'Printer.Font.Size = 10
          'For n = 0 To g.Rows - 1
          '    g.Row = n
          '    For z = 0 To x - 1
          '        g.Col = z
          '        Printer.Print Tab(12 * z); g;
          '    Next
          '    Printer.Print
          'Next
          '
          'Printer.Print Tab(30); "----End of Report----"
          '
          'Printer.EndDoc

880       Exit Sub

bprint_Click_Error:

          Dim strES As String
          Dim intEL As Integer



890       intEL = Erl
900       strES = Err.Description
910       LogError "frmFullEnd", "bPrint_Click", intEL, strES


End Sub

Private Sub chkChartNumber_Click()
10        g.Visible = False
          If cmbResultCount.Text <> "" Then
             FillG (Trim(cmbResultCount.Text))
          End If
          g.Visible = True
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
60        LogError "frmFullEnd", "cmbPlotFrom_KeyPress", intEL, strES


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
60        LogError "frmFullEnd", "cmbPlotTo_KeyPress", intEL, strES


End Sub


Private Sub cmdExcel_Click()
          Dim strHeading As String
10        On Error GoTo cmdExcel_Click_Error

20        strHeading = "Endocrinology History" & vbCr
30        strHeading = strHeading & "Patient Name: " & lblName & vbCr
40        strHeading = strHeading & vbCr

50        ExportFlexGrid g, Me, strHeading

60        Exit Sub

cmdExcel_Click_Error:

          Dim strES As String
          Dim intEL As Integer

70        intEL = Erl
80        strES = Err.Description
90        LogError "frmFullEnd", "cmdExcel_Click", intEL, strES
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
60        LogError "frmFullEnd", "cmdGo_Click", intEL, strES


End Sub



Private Sub cmdRefresh_Click()
10    On Error GoTo cmdRefresh_Click_Error

20      g.Visible = False
        If cmbResultCount.Text <> "" Then
           FillG (Trim(cmbResultCount.Text))
        End If
        g.Visible = True
FillCombos

30    Exit Sub

cmdRefresh_Click_Error:
      Dim strES As String
      Dim intEL As Integer

40    intEL = Erl
50    strES = Err.Description
60    LogError "frmFullEnd", "cmdRefresh_Click", intEL, strES

End Sub



Private Sub Form_Activate()

10        On Error GoTo Form_Activate_Error

20        dtFrom = Format(Now - SysOptWardDate(0), "dd/MMM/yyyy")
30        dtTo = Format(Now, "dd/MMM/yyyy")

40        If TestSys = True Then Me.Caption = Me.Caption & " - TEST SYSTEM"


50        If GetOptionSetting("EnableiPMSChart", "0") = 0 Then
60            chkChartNumber.Value = 0
70            chkChartNumber.Enabled = Not (lblChart = "")
80        Else
90            chkChartNumber.Value = 1

100       End If

110       g.Visible = False
          If cmbResultCount.Text <> "" Then
             FillG (Trim(cmbResultCount.Text))
          End If
          g.Visible = True
120       FillCombos


130       PBar.Max = LogOffDelaySecs
140       PBar = 0

150       Timer1.Enabled = True

160       Exit Sub

Form_Activate_Error:

          Dim strES As String
          Dim intEL As Integer



170       intEL = Erl
180       strES = Err.Description
190       LogError "frmFullEnd", "Form_Activate", intEL, strES


End Sub

Private Function GetRow(ByVal testnum As String) As Long

          Dim n As Long

10        On Error GoTo GetRow_Error

20        For n = 0 To List1.ListCount - 1
30            If UCase(testnum) = UCase(List1.List(n)) Then
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
110       LogError "frmFullEnd", "GetRow", intEL, strES


End Function

Private Function InList(ByVal s As String) As Long

          Dim n As Long

10        On Error GoTo InList_Error

20        InList = False
30        If List1.ListCount = 0 Then
40            Exit Function
50        End If

60        For n = 0 To List1.ListCount - 1
70            If UCase(s) = UCase(List1.List(n)) Then
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
150       LogError "frmFullEnd", "InList", intEL, strES


End Function

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
70            g.Row = n + 4
80            sql = "SELECT * from endtestdefinitions WHERE  code = '" & List1.List(n) & "' and  (agefromdays <= " & DaysOld & " and agetodays >= " & DaysOld & ")"
90            Set sn = New Recordset
100           RecOpenServer Tn, sn, sql
110           If Not sn.EOF Then
120               g = sn!LongName & ""
130               g.TextMatrix(g.Row, 1) = sn!SampleType
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
260       LogError "frmFullEnd", "TransferListToGrid", intEL, strES, sql


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
60        LogError "frmFullEnd", "Form_Deactivate", intEL, strES


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
60        LogError "frmFullEnd", "Form_MouseMove", intEL, strES


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
60        LogError "frmFullEnd", "g_Click", intEL, strES


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
60        LogError "frmFullEnd", "g_MouseMove", intEL, strES


End Sub


Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

10        On Error GoTo Image1_MouseMove_Error

20        PBar = 0

30        Exit Sub

Image1_MouseMove_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmFullEnd", "Image1_MouseMove", intEL, strES


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
60        LogError "frmFullEnd", "Label1_MouseMove", intEL, strES


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
60        LogError "frmFullEnd", "Label2_MouseMove", intEL, strES


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
60        LogError "frmFullEnd", "Label3_MouseMove", intEL, strES


End Sub


Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

10        On Error GoTo Label4_MouseMove_Error

20        PBar = 0

30        Exit Sub

Label4_MouseMove_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmFullEnd", "Label4_MouseMove", intEL, strES


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
60        LogError "frmFullEnd", "lblChart_MouseMove", intEL, strES


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
60        LogError "frmFullEnd", "lblDoB_MouseMove", intEL, strES


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
60        LogError "frmFullEnd", "lblName_MouseMove", intEL, strES


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
190       LogError "frmFullEnd", "pb_MouseMove", intEL, strES


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
90        LogError "frmFullEnd", "Timer1_Timer", intEL, strES


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

