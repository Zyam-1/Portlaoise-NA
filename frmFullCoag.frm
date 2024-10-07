VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmFullCoag 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Full Coagulation History"
   ClientHeight    =   8505
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   11535
   Icon            =   "frmFullCoag.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8505
   ScaleWidth      =   11535
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbResultCount 
      Height          =   315
      ItemData        =   "frmFullCoag.frx":030A
      Left            =   990
      List            =   "frmFullCoag.frx":0320
      TabIndex        =   34
      Text            =   "All"
      Top             =   420
      Width           =   1215
   End
   Begin VB.Frame FrameRange 
      Caption         =   "DateRange ====> by sample date"
      Height          =   870
      Left            =   6750
      TabIndex        =   28
      Top             =   4545
      Width           =   4650
      Begin VB.CommandButton cmdRefresh 
         Height          =   615
         Left            =   3780
         Picture         =   "frmFullCoag.frx":0352
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
         Format          =   189792259
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
         Format          =   189792259
         CurrentDate     =   38629
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
         TabIndex        =   33
         Top             =   180
         Width           =   960
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
         TabIndex        =   32
         Top             =   165
         Width           =   330
      End
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
      Left            =   3000
      TabIndex        =   26
      Top             =   4500
      Width           =   2535
   End
   Begin VB.CommandButton cmdExcel 
      Caption         =   "Export"
      Height          =   750
      Left            =   9540
      Picture         =   "frmFullCoag.frx":0C1C
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   6420
      Width           =   1245
   End
   Begin VB.PictureBox pb 
      AutoRedraw      =   -1  'True
      Height          =   2325
      Left            =   1860
      ScaleHeight     =   2265
      ScaleWidth      =   4125
      TabIndex        =   14
      Top             =   4905
      Width           =   4185
      Begin VB.Label lblTest 
         Alignment       =   2  'Center
         Height          =   225
         Left            =   1320
         TabIndex        =   15
         Top             =   0
         Width           =   1275
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   10440
      Top             =   5490
   End
   Begin VB.Frame Frame1 
      Caption         =   "Plot between"
      Height          =   2325
      Left            =   60
      TabIndex        =   10
      Top             =   4905
      Width           =   1785
      Begin VB.ComboBox cmbPlotTo 
         Height          =   315
         Left            =   150
         TabIndex        =   13
         Text            =   "cmbPlotTo"
         Top             =   870
         Width           =   1455
      End
      Begin VB.ComboBox cmbPlotFrom 
         Height          =   315
         Left            =   150
         TabIndex        =   12
         Text            =   "cmbPlotFrom"
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdGo 
         Caption         =   "Start"
         Height          =   900
         Left            =   270
         Picture         =   "frmFullCoag.frx":683A
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1290
         Width           =   1140
      End
   End
   Begin VB.CommandButton bcancel 
      Caption         =   "Exit"
      Height          =   750
      Left            =   6885
      Picture         =   "frmFullCoag.frx":6B44
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6495
      Width           =   1245
   End
   Begin VB.CommandButton bPrint 
      Cancel          =   -1  'True
      Caption         =   "Print"
      Height          =   750
      Left            =   6900
      Picture         =   "frmFullCoag.frx":6E4E
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5625
      Width           =   1245
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   2760
      Left            =   3930
      TabIndex        =   1
      Top             =   1410
      Visible         =   0   'False
      Width           =   1635
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   3615
      Left            =   45
      TabIndex        =   0
      Top             =   810
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
      Left            =   135
      TabIndex        =   16
      Top             =   7290
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
      Left            =   510
      TabIndex        =   36
      Top             =   480
      Width           =   450
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "Results"
      Height          =   195
      Left            =   2250
      TabIndex        =   35
      Top             =   480
      Width           =   525
   End
   Begin VB.Label Lbl1 
      Caption         =   "Amended  Results are underline"
      Height          =   195
      Left            =   8820
      TabIndex        =   27
      Top             =   585
      Width           =   2535
   End
   Begin VB.Label lblExcelInfo 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exporting..."
      Height          =   285
      Left            =   9540
      TabIndex        =   25
      Top             =   6120
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
      Height          =   330
      Left            =   1575
      TabIndex        =   22
      Top             =   4500
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
      Left            =   180
      TabIndex        =   23
      Top             =   4545
      Width           =   1680
   End
   Begin VB.Label Tn 
      Height          =   225
      Left            =   2940
      TabIndex        =   21
      Top             =   2310
      Width           =   315
   End
   Begin VB.Label lblMaxVal 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   6120
      TabIndex        =   20
      Top             =   4905
      Width           =   480
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   6120
      TabIndex        =   19
      Top             =   6975
      Width           =   480
   End
   Begin VB.Label lblMeanVal 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   6120
      TabIndex        =   18
      Top             =   5895
      Width           =   480
   End
   Begin VB.Label lblNopas 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3390
      TabIndex        =   17
      Top             =   5595
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
Attribute VB_Name = "frmFullCoag"
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
      'Dim X As Long
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
50        LinesToPrint = g.Rows - 3
60        Start = 3
70        PageCounter = ((g.Rows - 3) \ TotalLines)
80        If (g.Rows - 3) Mod TotalLines > 0 Then PageCounter = PageCounter + 1
90        CurrentPage = 1

100       If LinesToPrint > TotalLines Then
110           Last = TotalLines
120           LinesToPrint = LinesToPrint - TotalLines
130       Else
140           Last = LinesToPrint + 2
150           LinesToPrint = 0
160       End If



170       Do While CurrentPage <= PageCounter
180           Printer.Print
190           PrintText FormatString("Cumulative Coagulation Report", 70, , AlignCenter), 14, True, , , , True
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
330                   Select Case z
                      Case 1: cWidth = 5
340                   Case 2: cWidth = 13
350                   Case Else: cWidth = 9
360                   End Select
370                   PrintText FormatString(.TextMatrix(0, z), cWidth, "|", AlignCenter), 10

380               Next z
390               PrintText vbCrLf
                  'Print Sample Date Row
400               PrintText FormatString(.TextMatrix(1, 0), 16, "|"), 10, True
410               For z = 1 To MaxCols
420                   Select Case z
                      Case 1: cWidth = 5
430                   Case 2: cWidth = 13
440                   Case Else: cWidth = 9
450                   End Select
460                   PrintText FormatString(Format(.TextMatrix(1, z), "dd/MM/yy"), cWidth, "|", AlignCenter), 10
470               Next z
480               PrintText vbCrLf
                  'Print Sample Time Row

490               PrintText FormatString("SAMPLE TIME", 16, "|"), 10, True

500               For z = 1 To MaxCols
510                   Select Case z
                      Case 1: cWidth = 5
520                   Case 2: cWidth = 13
530                   Case Else: cWidth = 9
540                   End Select
550                   If Format(.TextMatrix(1, z), "hh:mm") = "00:00" Then
560                       PrintText FormatString("", cWidth, "|", AlignCenter), 10
570                   Else
580                       PrintText FormatString(Format(.TextMatrix(1, z), "hh:mm"), cWidth, "|", AlignCenter), 10
590                   End If
600               Next z
610               PrintText vbCrLf

                  'Print Run Date Row
620               PrintText FormatString(.TextMatrix(2, 0), 16, "|"), 10, True
630               PrintText FormatString("", 5, "|"), 10
640               PrintText FormatString("", 13, "|"), 10
650               For z = 3 To MaxCols

660                   PrintText FormatString(Format(.TextMatrix(2, z), "dd/MM/yy"), 9, "|", AlignCenter), 10
670               Next z
680               PrintText vbCrLf
                  'Print Run Time Row is exists

690               PrintText FormatString("RUN TIME", 16, "|"), 10, True
700               For z = 1 To MaxCols
710                   Select Case z
                      Case 1: cWidth = 5
720                   Case 2: cWidth = 13
730                   Case Else: cWidth = 9
740                   End Select
750                   If Format(.TextMatrix(2, z), "hh:mm") = "00:00" Then
760                       PrintText FormatString("", cWidth, "|", AlignCenter), 10
770                   Else
780                       PrintText FormatString(Format(.TextMatrix(2, z), "hh:mm"), cWidth, "|", AlignCenter), 10
790                   End If
800               Next z
810               PrintText vbCrLf
                  'Add seperator
820               PrintText String(217, "-") & vbCrLf, 4, True
                  'Print results
830               For Y = Start To Last
840                   PrintText FormatString(.TextMatrix(Y, 0), 16, "|"), 10
850                   For z = 1 To MaxCols
860                       Select Case z
                          Case 1: cWidth = 5
870                       Case 2: cWidth = 13
880                       Case Else: cWidth = 9
890                       End Select
900                       PrintText FormatString(.TextMatrix(Y, z), cWidth, "|", IIf(z = 2, AlignLeft, AlignCenter)), 10
910                   Next z
920                   PrintText vbCrLf
930               Next Y

940           End With


              'End of Page Line
950           PrintText String(217, "-"), 4, True
960           If CurrentPage < PageCounter Then Printer.NewPage
970           CurrentPage = CurrentPage + 1
980           Start = Last + 1
990           If LinesToPrint > TotalLines Then
1000              Last = Last + TotalLines
1010              LinesToPrint = LinesToPrint - TotalLines
1020          Else
1030              Last = Last + LinesToPrint + 3
1040              LinesToPrint = 0
1050          End If

1060      Loop

1070      Printer.EndDoc



          'X = g.Cols
          '
          'Printer.Orientation = vbPRORLandscape
          '
          '
          'Printer.Font.Size = 16
          'Printer.Print Tab(15); "Cumulative Report from Coagulation Dept."
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
          '    For z = 0 To X - 1
          '        g.Col = z
          '        Printer.Print Tab(15 * z); g;
          '    Next
          '    Printer.Print
          'Next
          '
          'Printer.Print Tab(30); "----End of Report----"
          '
          'Printer.EndDoc

1080      Exit Sub

bprint_Click_Error:

          Dim strES As String
          Dim intEL As Integer



1090      intEL = Erl
1100      strES = Err.Description
1110      LogError "frmFullCoag", "bPrint_Click", intEL, strES


End Sub

Private Sub chkChartNumber_Click()
10        If cmbResultCount.Text <> "" Then
             FillG (Trim(cmbResultCount.Text))
          End If
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

Private Sub cmdExcel_Click()
          Dim strHeading As String
10        On Error GoTo cmdExcel_Click_Error

20        strHeading = "Coagulation History" & vbCr
30        strHeading = strHeading & "Patient Name: " & lblName & vbCr
40        strHeading = strHeading & vbCr

50        ExportFlexGrid g, Me, strHeading

60        Exit Sub

cmdExcel_Click_Error:

          Dim strES As String
          Dim intEL As Integer

70        intEL = Erl
80        strES = Err.Description
90        LogError "frmFullCoag", "cmdExcel_Click", intEL, strES
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
60        LogError "frmFullCoag", "cmdGo_Click", intEL, strES


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
90        If NumberOfDays < 1 Then Exit Sub

100       ReDim ChartPositions(0 To NumberOfDays)

110       For n = 1 To NumberOfDays
120           ChartPositions(n).xPos = 0
130           ChartPositions(n).yPos = 0
140           ChartPositions(n).Value = 0
150           ChartPositions(n).Date = ""
160       Next

170       For n = 3 To g.Cols - 1
180           If Format$(g.TextMatrix(1, n), "dd/mmm/yyyy") = Format(cmbPlotTo, "dd/mmm/yyyy") Then StartGridX = n
190           If Format$(g.TextMatrix(1, n), "dd/mmm/yyyy") = Format(cmbPlotFrom, "dd/mmm/yyyy") Then StopGridX = n
200       Next

210       FirstDayFilled = False
220       Counter = 0
230       For x = StartGridX To StopGridX
240           If g.TextMatrix(g.Row, x) <> "" Then
250               If Not FirstDayFilled Then
260                   FirstDayFilled = True
270                   MaxVal = Val(g.TextMatrix(g.Row, x))
280                   ChartPositions(NumberOfDays).Date = Format(g.TextMatrix(1, x), "dd/mmm/yyyy")
290                   ChartPositions(NumberOfDays).Value = Val(g.TextMatrix(g.Row, x))
300               Else
310                   DaysInterval = Abs(DateDiff("D", cmbPlotTo, Format$(g.TextMatrix(1, x), "dd/mmm/yyyy")))
320                   ChartPositions(NumberOfDays - DaysInterval).Date = g.TextMatrix(1, x)
330                   cVal = Val(g.TextMatrix(g.Row, x))
340                   ChartPositions(NumberOfDays - DaysInterval).Value = cVal
350                   If cVal > MaxVal Then MaxVal = cVal
360               End If
370           End If
380       Next

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
700       LogError "frmFullCoag", "DrawChart", intEL, strES


End Sub
Private Sub FillCombos()

          Dim x As Long


10        On Error GoTo FillCombos_Error

20        cmbPlotFrom.Clear
30        cmbPlotTo.Clear

40        If g.Cols = 2 Then Exit Sub

50        For x = 3 To g.Cols - 1
60            cmbPlotFrom.AddItem Format$(g.TextMatrix(1, x), "dd/mmm/yyyy") & " " & Format(g.TextMatrix(2, x), "hh:mm")
70            cmbPlotTo.AddItem Format$(g.TextMatrix(1, x), "dd/mmm/yyyy") & " " & Format(g.TextMatrix(2, x), "hh:mm")
80        Next

90        cmbPlotTo = Format$(g.TextMatrix(1, 3), "dd/mmm/yyyy") & " " & Format(g.TextMatrix(2, 3), "hh:mm")
100       If Not IsDate(cmbPlotTo) Then Exit Sub

110       For x = g.Cols - 1 To 4 Step -1
120           If DateDiff("d", Format$(g.TextMatrix(1, x), "dd/mmm/yyyy") & " " & Format(g.TextMatrix(2, x), "hh:mm"), cmbPlotTo) < 365 Then
130               cmbPlotFrom = Format$(g.TextMatrix(1, x), "dd/mmm/yyyy") & " " & Format(g.TextMatrix(2, x), "hh:mm")
140               Exit For
150           End If
160       Next

170       Exit Sub

FillCombos_Error:

          Dim strES As String
          Dim intEL As Integer



180       intEL = Erl
190       strES = Err.Description
200       LogError "frmFullCoag", "FillCombos", intEL, strES


End Sub

Private Sub FillG(p_RCount As String)

          Dim tb As New Recordset
          Dim snr As Recordset
          Dim sql As String
          Dim x As Long
          Dim xrun As String
          Dim xdate As String
          Dim Flag As String
          Dim Code As String
          Dim sex As String
          Dim zx As Long
          Dim Unit As String
          Dim DaysOld As String
          Dim resultFlag As Boolean
          resultFlag = False

10        On Error GoTo FillG_Error

20        ClearFGrid g
30        p_RCount = UCase(p_RCount)
40        If p_RCount = "First 5" Then
50            sql = "SELECT distinct top 5 (demographics.sampleid), demographics.runDate, demographics.TimeTaken, demographics.sex, demographics.dob, demographics.sampledate, " & _
                  "(SELECT Top 1 Q.RunTime FROM CoagResults Q WHERE Q.SampleID = Demographics.SampleID ORDER BY Q.RunTime) RunDateTime " & _
                  "from demographics, coagresults WHERE ("
60        ElseIf p_RCount = "First 10" Then
70            sql = "SELECT distinct top 10 (demographics.sampleid), demographics.runDate, demographics.TimeTaken, demographics.sex, demographics.dob, demographics.sampledate, " & _
                  "(SELECT Top 1 Q.RunTime FROM CoagResults Q WHERE Q.SampleID = Demographics.SampleID ORDER BY Q.RunTime) RunDateTime " & _
                  "from demographics, coagresults WHERE ("
80        ElseIf p_RCount = "First 20" Then
90            sql = "SELECT distinct top 20 (demographics.sampleid), demographics.runDate, demographics.TimeTaken, demographics.sex, demographics.dob, demographics.sampledate, " & _
                  "(SELECT Top 1 Q.RunTime FROM CoagResults Q WHERE Q.SampleID = Demographics.SampleID ORDER BY Q.RunTime) RunDateTime " & _
                  "from demographics, coagresults WHERE ("
100       ElseIf p_RCount = "First 50" Then
110           sql = "SELECT distinct top 50 (demographics.sampleid), demographics.runDate, demographics.TimeTaken, demographics.sex, demographics.dob, demographics.sampledate, " & _
                  "(SELECT Top 1 Q.RunTime FROM CoagResults Q WHERE Q.SampleID = Demographics.SampleID ORDER BY Q.RunTime) RunDateTime " & _
                  "from demographics, coagresults WHERE ("
120       ElseIf p_RCount = "ALL" Then
130           sql = "SELECT distinct (demographics.sampleid), demographics.runDate, demographics.TimeTaken, demographics.sex, demographics.dob, demographics.sampledate, " & _
                  "(SELECT Top 1 Q.RunTime FROM CoagResults Q WHERE Q.SampleID = Demographics.SampleID ORDER BY Q.RunTime) RunDateTime " & _
                  "from demographics, coagresults WHERE ("
140       End If
150       If Trim(lblChart) <> "" And chkChartNumber.Value = 0 Then
160           sql = sql & "(demographics.chart = '" & lblChart & "') and"
170       End If
180       sql = sql & " ((demographics.PatName = '" & AddTicks(lblName) & "' and demographics.dob  = '" & Format(lblDoB, "dd/MMM/yyyy") & "') "
          '+++ Junaid 10-08-2023
          '80        sql = sql & "))  and demographics.rundate between'" & Format(dtFrom, "dd/MMM/yyyy") & "' and '" & Format(dtTo.Value + 1, "dd/MMM/yyyy") & "' " & _
          '              " and coagresults.sampleid = demographics.sampleid" & _
          '              " order by RunDateTime desc"
190       sql = sql & "))  and demographics.rundate between'" & Format(dtFrom, "dd/MMM/yyyy") & "' and '" & Format(dtTo.Value + 1, "dd/MMM/yyyy") & "' " & _
              " and coagresults.sampleid = demographics.sampleid" & _
              " order by demographics.sampledate desc"
          '--- Junaid
200       Set tb = New Recordset
210       RecOpenServer Tn, tb, sql
220       If Not tb.EOF Then
230           sex = tb!sex
240           g.Cols = 3
250           g.ColWidth(0) = 1600
260           g.ColWidth(1) = 1000
270           g.ColWidth(2) = 1400
280           g.TextMatrix(0, 0) = "SAMPLE ID"
290           g.TextMatrix(1, 0) = "SAMPLE DATE"
300           g.TextMatrix(2, 0) = "RUN DATE"
310           g.TextMatrix(2, 1) = "Units"
320           g.TextMatrix(2, 2) = "Ref Ranges"

330           Do While Not tb.EOF
340               g.Cols = g.Cols + 1
350               x = g.Cols - 1
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
                  'RunDateTime = GetRunDateTime(tb!SampleID, "Coag", Latest)
490               If tb!RunDateTime <> "" Then
500                   xdate = Format(tb!RunDateTime, "dd/mm/yy")
510                   If Format(tb!RunDateTime, "hh:mm") <> "00:00" Then
520                       xdate = xdate & " " & Format(tb!RunDateTime, "hh:mm")
530                   End If
540               Else
550                   xdate = ""
560               End If
570               g.TextMatrix(2, x) = xdate
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
580               If Not IsNull(tb!Dob) Then
590                   DaysOld = Abs(DateDiff("d", tb!Rundate, tb!Dob))
600               Else
610                   DaysOld = 12783
620               End If

                  'fill list with test names
630               sql = "SELECT CoagResults.*, PrintPriority, dp, coagtestdefinitions.units " & _
                      "from CoagResults, CoagTestDefinitions " & _
                      "WHERE SampleID = '" & xrun & "' " & _
                      "and coagTestDefinitions.Code = CoagResults.Code and " & _
                      " CoagTestDefinitions.Hospital = '" & HospName(Tn) & "' AND CoagTestDefinitions.Code <> '4' "
640               If SysOptExp(0) = True Then
650                   sql = sql & " and (coagresults.units = coagtestdefinitions.units)"
660               End If
670               sql = sql & "order by PrintPriority"
680               Set snr = New Recordset
690               RecOpenServer Tn, snr, sql
700               Do While Not snr.EOF

                      'If Trim(snr!Units) = "INR" Then
                      'Code = "INR"
                      'Else
                      'Code = Trim(snr!Code)
                      'End If
710                   Code = Trim(snr!Code)
720                   If Not InList(Code, snr!Units) Then
730                       List1.AddItem Code & " " & Trim(snr!Units)
740                   End If
750                   snr.MoveNext
760               Loop
770               tb.MoveNext
780           Loop
790           If List1.ListCount = 0 Then Exit Sub
              'fill first col with test names
800           TransferListToGrid DaysOld

              'fill in results
810           For x = 3 To g.Cols - 1
820               g.Col = x
830               g.Row = 1
840               xdate = Format$(g, "dd/mmm/yyyy")
850               g.Row = 0
860               xrun = g

870               sql = "SELECT * from CoagResults "
880               sql = sql & "WHERE SampleID = '" & xrun & "'"
890               Set snr = New Recordset
900               RecOpenServer Tn, snr, sql
910               Do While Not snr.EOF
                      'Zyam
                      If InStr(1, snr!Result, ">") Then
                        resultFlag = True
                      End If
                      'Zyam
920                   If Trim(snr!RunTime) <> "" And g.TextMatrix(2, x) = "" Then
930                       g.TextMatrix(2, x) = Format(snr!RunTime, "hh:mm")
940                   End If
                      '            If Trim(snr!Units) = "INR" And CoagNameFor(snr!Code) = "INR" Then
                      '                Code = "INR"
                      '                Unit = "INR"
                      '            Else
                      '                If UCase(CoagNameFor(snr!Code)) = "INR" Then
                      '                    Code = "INR"
                      '                    Unit = ""
                      '                Else
                      '                    Code = Trim(snr!Code)
                      '                    Unit = UnitConv(snr!Units & "")
                      '                End If
                      '            End If
950                   Code = Trim(snr!Code)
960                   Unit = UnitConv(snr!Units & "")

970                   zx = GetRow(Code, Unit)
980                   g.Row = zx
990                   If g.Row <> 0 Then
1000                      If (UserMemberOf = "Secretarys" Or SysOptNoCumShow(0) = True) And snr!Valid = False Then
1010                          g = "NV"
1020                      Else

1030                          Select Case CoagPrintFormat(Trim(snr!Code) & "")
                                  Case 0: g = Format$(snr!Result, "0")
1040                              Case 1: g = Format$(snr!Result, "0.0")
1050                              Case 2: g = Format$(snr!Result, "0.00")
1060                          End Select
1070                          If snr!Valid = False Then
1080                              g = g & " (NV)"
1090                          End If
1100                      End If
                          '----------------------
1110                      If IsResultAmended("Coag", snr!SampleID, snr!Code, snr!Result) = True Then
1120                          g.CellFontUnderline = True
1130                      End If
                          '======================
1140                  End If
1150                  Flag = InterpCoag(sex, snr!Code, snr!Result, DaysOld, resultFlag)
1160                  g.Col = x
1170                  g.Row = zx
1180                  If zx <> 0 Then
1190                      If Flag = "X" Then
1200                          g.CellBackColor = SysOptPlasBack(0)
1210                          g.CellForeColor = SysOptPlasFore(0)
1220                      ElseIf Flag = "H" Then
1230                          g.CellForeColor = SysOptHighFore(0)
1240                          g.CellBackColor = SysOptHighBack(0)
1250                      ElseIf Flag = "L" Then
1260                          g.CellForeColor = SysOptLowFore(0)
1270                          g.CellBackColor = SysOptLowBack(0)
1280                      Else
1290                          g.CellForeColor = vbBlack
1300                          g.CellBackColor = vbWhite
1310                      End If
1320                  End If
1330                  Flag = ""
1340                  snr.MoveNext
1350              Loop
1360          Next
1370      End If

1380      If g.Cols > 2 Then lblNoRes.Caption = g.Cols - 3 Else lblNoRes.Caption = 0


          'If g.Rows > 4 Then
          '  g.RemoveItem 3
          'End If
1390      g.Visible = True


1400      Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer

1410      intEL = Erl
1420      strES = Err.Description
1430      LogError "frmFullCoag", "FillG", intEL, strES, sql


End Sub





Private Sub cmdRefresh_Click()
10    On Error GoTo cmdRefresh_Click_Error

          If cmbResultCount.Text <> "" Then
20          FillG (Trim(cmbResultCount.Text))
          End If
30    FillCombos


40    Exit Sub

cmdRefresh_Click_Error:
      Dim strES As String
      Dim intEL As Integer

50    intEL = Erl
60    strES = Err.Description
70    LogError "frmFullCoag", "cmdRefresh_Click", intEL, strES

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

90        dtFrom = Format(Now - SysOptWardDate(0), "dd/MMM/yyyy")
100       dtTo = Format(Now, "dd/MMM/yyyy")

110       If TestSys = True Then Me.Caption = Me.Caption & " - TEST SYSTEM"
            
          g.Visible = False
          If cmbResultCount.Text <> "" Then
120         FillG UCase((Trim(cmbResultCount.Text)))
          End If
          g.Visible = True
130       FillCombos

          

140       pBar.Max = LogOffDelaySecs
150       pBar = 0

160       Timer1.Enabled = True

170       Exit Sub

Form_Activate_Error:

          Dim strES As String
          Dim intEL As Integer



180       intEL = Erl
190       strES = Err.Description
200       LogError "frmFullCoag", "Form_Activate", intEL, strES


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
60        LogError "frmFullCoag", "Form_Deactivate", intEL, strES


End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

10        On Error GoTo Form_MouseMove_Error

20        pBar = 0

30        Exit Sub

Form_MouseMove_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmFullCoag", "Form_MouseMove", intEL, strES


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
60        LogError "frmFullCoag", "g_Click", intEL, strES


End Sub

Private Sub g_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

10        On Error GoTo g_MouseMove_Error

20        pBar = 0

30        Exit Sub

g_MouseMove_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmFullCoag", "g_MouseMove", intEL, strES


End Sub

Private Function GetRow(ByVal testnum As String, ByVal Unit As String) As Long

          Dim n As Long
          Dim Units As String



10        On Error GoTo GetRow_Error

20        For n = 0 To List1.ListCount - 1
30            If testnum = Trim(Left(List1.List(n), InStr(1, List1.List(n), " "))) Then    'And UCase(Trim(Unit)) = UCase(Units) Then
40                GetRow = n + 3: Debug.Print List1.List(n)
50                Exit For
60            End If
70        Next

80        Exit Function

GetRow_Error:

          Dim strES As String
          Dim intEL As Integer



90        intEL = Erl
100       strES = Err.Description
110       LogError "frmFullCoag", "GetRow", intEL, strES


End Function

Private Function InList(ByVal s As String, ByVal U As String) As Long

          Dim n As Long

10        On Error GoTo InList_Error

20        InList = False
30        If List1.ListCount = 0 Then
40            Exit Function
50        End If

60        For n = 0 To List1.ListCount - 1
70            If s & " " & Trim(U) = List1.List(n) Then
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
150       LogError "frmFullCoag", "InList", intEL, strES


End Function

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

10        On Error GoTo Label1_MouseMove_Error

20        pBar = 0

30        Exit Sub

Label1_MouseMove_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmFullCoag", "Label1_MouseMove", intEL, strES


End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

10        On Error GoTo Label2_MouseMove_Error

20        pBar = 0

30        Exit Sub

Label2_MouseMove_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmFullCoag", "Label2_MouseMove", intEL, strES


End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

10        On Error GoTo Label3_MouseMove_Error

20        pBar = 0

30        Exit Sub

Label3_MouseMove_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmFullCoag", "Label3_MouseMove", intEL, strES


End Sub

Private Sub lblChart_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

10        On Error GoTo lblChart_MouseMove_Error

20        pBar = 0

30        Exit Sub

lblChart_MouseMove_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmFullCoag", "lblChart_MouseMove", intEL, strES


End Sub

Private Sub lblDoB_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

10        On Error GoTo lblDoB_MouseMove_Error

20        pBar = 0

30        Exit Sub

lblDoB_MouseMove_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmFullCoag", "lblDoB_MouseMove", intEL, strES


End Sub

Private Sub lblName_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

10        On Error GoTo lblName_MouseMove_Error

20        pBar = 0

30        Exit Sub

lblName_MouseMove_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmFullCoag", "lblName_MouseMove", intEL, strES


End Sub





Private Sub pb_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

          Dim i As Long
          Dim CurrentDistance As Long
          Dim BestDistance As Long
          Dim BestIndex As Integer


10        On Error GoTo pb_MouseMove_Error

20        pBar = 0
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
190       LogError "frmFullCoag", "pb_MouseMove", intEL, strES


End Sub

Private Sub Timer1_Timer()

      'tmrRefresh.Interval set to 1000
10        On Error GoTo Timer1_Timer_Error

20        pBar = pBar + 1

30        If pBar = pBar.Max Then
40            Unload Me
50        End If

60        Exit Sub

Timer1_Timer_Error:

          Dim strES As String
          Dim intEL As Integer



70        intEL = Erl
80        strES = Err.Description
90        LogError "frmFullCoag", "Timer1_Timer", intEL, strES


End Sub

Private Sub TransferListToGrid(DaysOld As String)

      Dim n          As Long
      Dim sql        As String
      Dim tb         As New Recordset
      Dim Code       As String
      Dim Unit       As String
      Dim m          As Long
      Dim CodeUnit() As String

10    On Error GoTo TransferListToGrid_Error
      'Debug.Print CalcAge(CDate(lblDoB), CDate(g.TextMatrix(2, 4)))
20    If List1.ListCount = 0 Then Exit Sub

30    g.Rows = List1.ListCount + 3

40    g.Col = 0
50    For n = 0 To List1.ListCount - 1
60        g.Row = n + 3
70        Code = ""
80        Unit = ""
          
90        CodeUnit = Split(List1.List(n), " ")
100       Code = CodeUnit(0)
110       If UBound(CodeUnit) > 0 Then Unit = CodeUnit(1)
120       sql = "SELECT * from coagtestdefinitions WHERE agefromdays <= '" & DaysOld & "' AND AgeToDays >= '" & DaysOld & "' AND  code = '" & Code & "' and units = '" & Unit & "'"
130       Set tb = New Recordset
140       Set tb = Cnxn(Tn).Execute(sql)
150       If Not tb.EOF Then
160           g = tb!TestName
170           g.TextMatrix(g.Row, 1) = UnitConv(Unit)
              'Zyam 11-3-24
              If Code = "8001" Or Code = "9999" Or Code = "1" Or Code = "13" Or Code = "94" Then
181           g.TextMatrix(g.Row, 2) = ""
              Else
180           g.TextMatrix(g.Row, 2) = "( " & tb!FemaleLow & " - " & tb!MaleHigh & " ) "
              End If

              'Zyam 11-3-24
190       End If

          '        If Left(List1.List(n), 3) <> "INR" Then
          '            For m = 1 To Len(List1.List(n))
          '                If Mid(List1.List(n), m, 1) = " " Then
          '                    Unit = Mid(List1.List(n), m + 1, Len(List1.List(n)) - m)
          '                    Exit For
          '                End If
          '                Code = Code & Mid(List1.List(n), m, 1)
          '            Next
          '            sql = "SELECT * from coagtestdefinitions WHERE agefromdays <= '" & DaysOld & "' AND AgeToDays >= '" & DaysOld & "' AND  code = '" & Code & "' and units = '" & Unit & "'"
          '            Set tb = New Recordset
          '            Set tb = Cnxn(Tn).Execute(sql)
          '            If Not tb.EOF Then
          '                g = tb!TestName
          '                If Unit <> "INR" Then g.TextMatrix(g.row, 1) = UnitConv(Unit)
          '                g.TextMatrix(g.row, 2) = "( " & tb!FemaleLow & " - " & tb!MaleHigh & " ) "
          '            End If
          '        Else
          '            g = "INR"
          '        End If
200   Next

210   Exit Sub

TransferListToGrid_Error:

      Dim strES      As String
      Dim intEL      As Integer



220   intEL = Erl
230   strES = Err.Description
240   LogError "frmFullCoag", "TransferListToGrid", intEL, strES, sql


End Sub
