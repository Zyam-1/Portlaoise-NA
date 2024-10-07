VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFullImm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Full Immunology History"
   ClientHeight    =   7095
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
   Icon            =   "frmFullImm.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7095
   ScaleWidth      =   11565
   Begin VB.ComboBox cmbResultCount 
      Height          =   315
      ItemData        =   "frmFullImm.frx":030A
      Left            =   540
      List            =   "frmFullImm.frx":0320
      TabIndex        =   44
      Text            =   "All"
      Top             =   30
      Width           =   1215
   End
   Begin VB.CheckBox chkChartNumber 
      Caption         =   "Ignore chart number"
      Height          =   255
      Left            =   2490
      TabIndex        =   42
      Top             =   45
      Width           =   2115
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
      Left            =   8910
      Picture         =   "frmFullImm.frx":0352
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   4980
      Width           =   1245
   End
   Begin Threed.SSPanel pnlFetching 
      Height          =   285
      Left            =   60
      TabIndex        =   33
      Top             =   6690
      Width           =   6585
      _Version        =   65536
      _ExtentX        =   11615
      _ExtentY        =   503
      _StockProps     =   15
      ForeColor       =   -2147483630
      BackColor       =   65535
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FloodType       =   1
   End
   Begin VB.Frame fraHL 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   1920
      TabIndex        =   36
      Top             =   6720
      Visible         =   0   'False
      Width           =   3285
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Abnormal Results shown "
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
         Left            =   150
         TabIndex        =   39
         Top             =   30
         Width           =   1785
      End
      Begin VB.Label lblHigh 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "HIGH"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   2040
         TabIndex        =   38
         Top             =   30
         Width           =   495
      End
      Begin VB.Label lblLow 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "LOW"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   2670
         TabIndex        =   37
         Top             =   30
         Width           =   435
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Print Cumulative Report"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   90
      TabIndex        =   26
      Top             =   8220
      Width           =   6615
      Begin VB.CommandButton cmdPrint 
         Height          =   435
         Left            =   5910
         Picture         =   "frmFullImm.frx":5F70
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Print Cumulative Report"
         Top             =   180
         Width           =   555
      End
      Begin VB.ComboBox cmbPrinter 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   960
         TabIndex        =   27
         Text            =   "cmbPrinter"
         Top             =   240
         Width           =   3195
      End
      Begin MSComCtl2.UpDown udPrevious 
         Height          =   315
         Left            =   4800
         TabIndex        =   29
         Top             =   240
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   556
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "lblPrevious"
         BuddyDispid     =   196620
         OrigLeft        =   3510
         OrigTop         =   930
         OrigRight       =   4050
         OrigBottom      =   1215
         Max             =   1
         Min             =   1
         Orientation     =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label lblPrevious 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "6"
         Height          =   315
         Left            =   4350
         TabIndex        =   32
         Top             =   240
         Width           =   435
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Current and Previous"
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
         Left            =   4170
         TabIndex        =   31
         Top             =   0
         Width           =   1515
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Use Printer"
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
         Left            =   120
         TabIndex        =   30
         Top             =   300
         Width           =   780
      End
   End
   Begin VB.CommandButton cmdPara 
      Caption         =   "Paraprotein History"
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
      Left            =   8910
      TabIndex        =   24
      ToolTipText     =   "Paraprotein History"
      Top             =   5790
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
      Left            =   10260
      Picture         =   "frmFullImm.frx":65DA
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Print"
      Top             =   4980
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
      Left            =   10260
      Picture         =   "frmFullImm.frx":68E4
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Exit"
      Top             =   5790
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
      Top             =   1545
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
         Picture         =   "frmFullImm.frx":6BEE
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
      Left            =   10200
      Top             =   1140
   End
   Begin VB.PictureBox pb 
      AutoRedraw      =   -1  'True
      Height          =   2325
      Left            =   6780
      ScaleHeight     =   2265
      ScaleWidth      =   4125
      TabIndex        =   8
      Top             =   2550
      Width           =   4185
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   6300
      Left            =   60
      TabIndex        =   7
      Top             =   390
      Width           =   6585
      _ExtentX        =   11615
      _ExtentY        =   11113
      _Version        =   393216
      Rows            =   4
      Cols            =   3
      FixedRows       =   3
      FixedCols       =   2
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      ScrollTrack     =   -1  'True
      FormatString    =   "<Code  |<         |<         "
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   2760
      Left            =   10770
      TabIndex        =   6
      Top             =   7230
      Visible         =   0   'False
      Width           =   1635
   End
   Begin MSComctlLib.ProgressBar PBar 
      Height          =   195
      Left            =   6750
      TabIndex        =   10
      Top             =   6750
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSFlexGridLib.MSFlexGrid gDem 
      Height          =   2745
      Left            =   6510
      TabIndex        =   25
      Top             =   8370
      Visible         =   0   'False
      Width           =   4035
      _ExtentX        =   7117
      _ExtentY        =   4842
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      FormatString    =   "<SampleID    |<SampleDate   |<SampleTime |<Cnxn "
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "Show "
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
      Left            =   60
      TabIndex        =   46
      Top             =   90
      Width           =   450
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "Results"
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
      Left            =   1800
      TabIndex        =   45
      Top             =   90
      Width           =   525
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
      Left            =   4650
      TabIndex        =   43
      Top             =   60
      Width           =   2535
   End
   Begin VB.Label lblExcelInfo 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exporting..."
      Height          =   285
      Left            =   7620
      TabIndex        =   41
      Top             =   4980
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Label lblSex 
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
      Height          =   285
      Left            =   10530
      TabIndex        =   35
      Top             =   450
      Width           =   975
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
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
      Height          =   195
      Left            =   10170
      TabIndex        =   34
      Top             =   510
      Width           =   270
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
      Left            =   8130
      TabIndex        =   22
      ToolTipText     =   "Number of Results"
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
      Left            =   6735
      TabIndex        =   23
      Top             =   5910
      Width           =   1680
   End
   Begin VB.Label lblNopas 
      Height          =   255
      Left            =   1080
      TabIndex        =   21
      Top             =   1020
      Width           =   1365
   End
   Begin VB.Label Tn 
      Height          =   225
      Left            =   3840
      TabIndex        =   20
      Top             =   1710
      Width           =   315
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   6720
      Picture         =   "frmFullImm.frx":6EF8
      Top             =   1080
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
      Top             =   2550
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
      Top             =   3600
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
      Top             =   4620
      Width           =   720
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
      Top             =   1170
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
      Left            =   6720
      TabIndex        =   5
      Top             =   810
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
      Left            =   8550
      TabIndex        =   4
      Top             =   510
      Width           =   315
   End
   Begin VB.Label lblName 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   7170
      TabIndex        =   3
      ToolTipText     =   "Patients Name"
      Top             =   780
      Width           =   4335
   End
   Begin VB.Label lblDoB 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   8880
      TabIndex        =   2
      ToolTipText     =   "Date of Birth"
      Top             =   450
      Width           =   1155
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
      Left            =   6750
      TabIndex        =   1
      Top             =   510
      Width           =   375
   End
   Begin VB.Label lblChart 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   7170
      TabIndex        =   0
      ToolTipText     =   "Chart"
      Top             =   450
      Width           =   1245
   End
End
Attribute VB_Name = "frmFullImm"
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
Dim MaxPrevious As Integer

Dim gArray() As String

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
80        If NumberOfDays < 1 Then Exit Sub
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
690       LogError "frmFullImm", "DrawChart", intEL, strES


End Sub
Private Sub FillCombos()

          Dim x As Long
          Dim Px As Printer

10        On Error GoTo FillCombos_Error

20        cmbPlotFrom.Clear
30        cmbPlotTo.Clear

40        For x = 4 To g.Cols - 1
50            cmbPlotFrom.AddItem Format$(g.TextMatrix(1, x), "dd/mmm/yyyy")
60            cmbPlotTo.AddItem Format$(g.TextMatrix(1, x), "dd/mmm/yyyy")
70        Next

80        cmbPlotTo = Format$(g.TextMatrix(1, 4), "dd/mmm/yyyy")

90        For x = g.Cols - 1 To 5 Step -1
100           If g.TextMatrix(1, x) <> "" Then
110               If DateDiff("d", Format$(g.TextMatrix(1, x), "dd/mmm/yyyy"), cmbPlotTo) < 365 Then
120                   cmbPlotFrom = Format$(g.TextMatrix(1, x), "dd/mmm/yyyy")
130                   Exit For
140               End If
150           End If
160       Next

170       cmbPrinter.Clear
180       For Each Px In Printers
190           cmbPrinter.AddItem Px.DeviceName
200       Next
210       For x = 0 To cmbPrinter.ListCount - 1
220           If Printer.DeviceName = cmbPrinter.List(x) Then
230               cmbPrinter.ListIndex = x
240               Exit For
250           End If
260       Next

270       Exit Sub

FillCombos_Error:

          Dim strES As String
          Dim intEL As Integer



280       intEL = Erl
290       strES = Err.Description
300       LogError "frmFullImm", "FillCombos", intEL, strES

End Sub

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

20        For Each Px In Printers
30            If Px.DeviceName = cmbPrinter.Text Then
40                Set Printer = Px
50                Exit For
60            End If
70        Next

80        Printer.Font = "Courier New"
90        Printer.Orientation = vbPRORLandscape


100       TotalLines = 27
110       LinesToPrint = g.Rows - 4
120       Start = 4
130       PageCounter = ((g.Rows - 4) \ TotalLines)
140       If (g.Rows - 4) Mod TotalLines > 0 Then PageCounter = PageCounter + 1
150       CurrentPage = 1

160       If LinesToPrint > TotalLines Then
170           Last = TotalLines
180           LinesToPrint = LinesToPrint - TotalLines
190       Else
200           Last = LinesToPrint + 3
210           LinesToPrint = 0
220       End If

230       Do While CurrentPage <= PageCounter
240           Printer.Print
250           PrintText FormatString("Cumulative Immunology Report", 70, , AlignCenter), 14, True, , , , True
260           PrintText FormatString("Page " & Format$(CurrentPage) & " of " & PageCounter, 100, , AlignCenter), 10, , , , , True

270           PrintText "  Patient Name: " & lblName, 14, True, , , , True
280           PrintText " Date of Birth: " & Format$(lblDoB, "dd/mm/yyyy"), 14, True, , , , True
290           PrintText "         Chart: " & lblChart, 14, True, , , , True

300           Printer.Print

310           With g
320               If .Cols > 9 Then
330                   MaxCols = 8
340               Else
350                   MaxCols = .Cols - 1
360               End If

                  'Add seperator
370               PrintText String(217, "-") & vbCrLf, 4, True
                  'Print SampleID row
380               PrintText FormatString(.TextMatrix(0, 1), 16, "|"), 10, True
390               For z = 2 To MaxCols
400                   Select Case z
                      Case 2: cWidth = 5
410                   Case 3: cWidth = 13
420                   Case Else: cWidth = 9
430                   End Select
440                   PrintText FormatString(.TextMatrix(0, z), cWidth, "|", AlignCenter), 10
450               Next z
460               PrintText vbCrLf
                  'Print Sample Date Row
470               PrintText FormatString(.TextMatrix(1, 1), 16, "|"), 10, True
480               For z = 2 To MaxCols
490                   Select Case z
                      Case 2: cWidth = 5
500                   Case 3: cWidth = 13
510                   Case Else: cWidth = 9
520                   End Select
530                   PrintText FormatString(Format(.TextMatrix(1, z), "dd/MM/yy"), cWidth, "|", AlignCenter), 10
540               Next z
550               PrintText vbCrLf
                  'Print Sample Time Row is exists

560               PrintText FormatString("SAMPLE TIME", 16, "|"), 10, True
570               For z = 2 To MaxCols
580                   Select Case z
                      Case 2: cWidth = 5
590                   Case 3: cWidth = 13
600                   Case Else: cWidth = 9
610                   End Select
620                   If Format(.TextMatrix(1, z), "hh:mm") = "00:00" Then
630                       PrintText FormatString("", cWidth, "|", AlignCenter), 10
640                   Else
650                       PrintText FormatString(Format(.TextMatrix(1, z), "hh:mm"), cWidth, "|", AlignCenter), 10
660                   End If
670               Next z
680               PrintText vbCrLf

                  'Print Run Date Row
690               PrintText FormatString(.TextMatrix(2, 1), 16, "|"), 10, True
700               PrintText FormatString("", 5, "|"), 10
710               PrintText FormatString("", 13, "|"), 10
720               For z = 4 To MaxCols
730                   PrintText FormatString(Format(.TextMatrix(2, z), "dd/MM/yy"), 9, "|", AlignCenter), 10
740               Next z
750               PrintText vbCrLf
                  'Print Run Time Row is exists

760               PrintText FormatString("RUN TIME", 16, "|"), 10, True
770               For z = 2 To MaxCols
780                   Select Case z
                      Case 2: cWidth = 5
790                   Case 3: cWidth = 13
800                   Case Else: cWidth = 9
810                   End Select
820                   If Format(.TextMatrix(2, z), "hh:mm") = "00:00" Then
830                       PrintText FormatString("", cWidth, "|", AlignCenter), 10
840                   Else
850                       PrintText FormatString(Format(.TextMatrix(2, z), "hh:mm"), cWidth, "|", AlignCenter), 10
860                   End If
870               Next z
880               PrintText vbCrLf
                  'Add seperator
890               PrintText String(217, "-") & vbCrLf, 4, True
                  'Print results
900               For Y = Start To Last
910                   PrintText FormatString(ImmShortNameFor(.TextMatrix(Y, 0)), 16, "|"), 10
920                   For z = 2 To MaxCols
930                       Select Case z
                          Case 2: cWidth = 5
940                       Case 3: cWidth = 13
950                       Case Else: cWidth = 9
960                       End Select
970                       PrintText FormatString(.TextMatrix(Y, z), cWidth, "|", IIf(z = 2, AlignLeft, AlignCenter)), 10
980                   Next z
990                   PrintText vbCrLf
1000              Next Y

1010          End With


              'End of Page Line
1020          PrintText String(217, "-"), 4, True
1030          If CurrentPage < PageCounter Then Printer.NewPage
1040          CurrentPage = CurrentPage + 1
1050          Start = Last + 1
1060          If LinesToPrint > TotalLines Then
1070              Last = Last + TotalLines
1080              LinesToPrint = LinesToPrint - TotalLines
1090          Else
1100              Last = Last + LinesToPrint + 3
1110              LinesToPrint = 0
1120          End If

1130      Loop

1140      Printer.EndDoc



          'X = g.Cols
          '
          'Printer.Orientation = vbPRORLandscape
          '
          'Printer.Font.Size = 16
          'Printer.Print Tab(15); "Cumulative Report from Immumology Department"
          '
          'Printer.Font.Size = 14
          'Printer.Print Tab(10); "Name : " & lblName;
          'Printer.Print Tab(40); "Dob  : " & lblDoB
          '
          'Printer.Print
          'Printer.Print
          '
          'Printer.Font.Name = "Courier New"
          'Printer.Font.Size = 6
          'For n = 0 To g.Rows - 1
          '    g.Row = n
          '    For z = 1 To X - 1
          '        g.Col = z
          '        Printer.Print Tab(20 * (z - 1)); Left$(g.TextMatrix(n, z) & Space(19), 19);
          '    Next
          '    Printer.Print
          'Next
          '
          'Printer.Print Tab(30); "----End of Report----"
          '
          'Printer.EndDoc

1150      Exit Sub

bprint_Click_Error:

          Dim strES As String
          Dim intEL As Integer



1160      intEL = Erl
1170      strES = Err.Description
1180      LogError "frmFullImm", "bPrint_Click", intEL, strES

End Sub

Private Sub chkChartNumber_Click()
10        g.Visible = False
          If cmbResultCount.Text <> "" Then
              FillgDem (Trim(cmbResultCount.Text))
          End If
          g.Visible = True
20        FillResultGrid
30        FillCodeGrid
40        TransferDems

50        g.ColWidth(0) = 0
60        g.ColWidth(1) = 1600
70        g.ColWidth(2) = 495
80        g.ColWidth(3) = 1400
90        g.ColAlignment(3) = flexAlignLeftCenter
100       g.TextMatrix(0, 1) = "SAMPLE ID"
110       g.TextMatrix(1, 1) = "SAMPLE DATE"
120       g.TextMatrix(2, 1) = "RUN DATE"
130       g.TextMatrix(2, 2) = "S/T"
140       g.TextMatrix(2, 3) = "Ref Ranges"

150       Me.Refresh
160       FillAllResults
170       RemoveLines

          'FillG
180       FillCombos

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
60        LogError "frmFullImm", "cmbPlotFrom_KeyPress", intEL, strES


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
60        LogError "frmFullImm", "cmbPlotTo_KeyPress", intEL, strES


End Sub


Private Sub cmbResultCount_Change()
    On Error GoTo cmbResultCount_Change_Error
    
    g.Visible = False
     
    If cmbResultCount.Text <> "" Then
         FillgDem (Trim(cmbResultCount.Text))
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
         FillgDem (Trim(cmbResultCount.Text))
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

20        strHeading = "Immunology History" & vbCr
30        strHeading = strHeading & "Patient Name: " & lblName & vbCr
40        strHeading = strHeading & vbCr

50        ExportFlexGrid g, Me, strHeading

60        Exit Sub

cmdExcel_Click_Error:

          Dim strES As String
          Dim intEL As Integer

70        intEL = Erl
80        strES = Err.Description
90        LogError "frmFullImm", "cmdExcel_Click", intEL, strES
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
60        LogError "frmFullImm", "cmdGo_Click", intEL, strES


End Sub

Private Sub cmdPara_Click()

10        On Error GoTo cmdPara_Click_Error

20        With frmFullParaImm
30            .lblChart = lblChart
40            .lblName = lblName
50            .lblDoB = lblDoB
60            .Tn = "0"
70            .Show 1
80        End With

90        Exit Sub

cmdPara_Click_Error:

          Dim strES As String
          Dim intEL As Integer



100       intEL = Erl
110       strES = Err.Description
120       LogError "frmFullImm", "cmdPara_Click", intEL, strES


End Sub

Private Sub cmdPrint_Click()
          Dim Px As Printer
          Dim x As Integer
          Dim Y As Integer
          Dim PageCounter As Integer
          Dim CurrentPage As Integer
          Dim TabPos As Integer
          Dim ArrayStart As Integer
          Dim ArrayStop As Integer

10        For Each Px In Printers
20            If Px.DeviceName = cmbPrinter.Text Then
30                Set Printer = Px
40                Exit For
50            End If
60        Next

70        PageCounter = (Val(lblPrevious) \ 7) + 1
80        CurrentPage = 1

90        Do While CurrentPage <= PageCounter
100           Printer.Font.Name = "Courier New"
110           Printer.Font.Size = 14
120           Printer.Font.Bold = True
130           Printer.Print
140           Printer.Print "Cumulative Biochemistry Report              ";
150           Printer.Font.Size = 10
160           Printer.Font.Bold = False
170           Printer.Print "Page "; Format$(CurrentPage); " of "; PageCounter
180           Printer.Print

190           Printer.Font.Size = 14
200           Printer.Font.Bold = True
210           Printer.Print " Patient Name:"; lblName
220           If lblDoB <> "" Then
230               Printer.Print "Date of Birth:"; Format$(lblDoB, "dd/mm/yyyy")
240           End If
250           Printer.Print "        Chart:"; lblChart
260           Printer.Print
270           Printer.Print

280           Printer.Font.Size = 10
290           Printer.Font.Bold = False

300           For Y = 0 To UBound(gArray, 1) - 1
310               Printer.Print gArray(Y, 0);
320               TabPos = 10

330               ArrayStart = ((CurrentPage - 1) * 7) + 3
340               ArrayStop = ArrayStart + 6

                  'MaxCol = ((CurrentPage - 1) * 7) + 7
350               If ArrayStop > Val(lblPrevious) + 3 Then
360                   ArrayStop = Val(lblPrevious) + 3
370               End If
380               For x = ArrayStart To ArrayStop
390                   If x <= UBound(gArray, 2) Then
400                       Printer.Print Tab(TabPos); gArray(Y, x);
410                   End If
420                   TabPos = TabPos + 10
430               Next
440               Printer.Print
450           Next

460           Printer.EndDoc

470           CurrentPage = CurrentPage + 1

480       Loop

End Sub

Private Sub Form_Activate()

10        On Error GoTo Form_Activate_Error

20        If TestSys = True Then Me.Caption = Me.Caption & " - TEST SYSTEM"

30        If GetOptionSetting("EnableiPMSChart", "0") = 0 Then
40            chkChartNumber.Value = 0
50            chkChartNumber.Enabled = Not (lblChart = "")
60        Else
70            chkChartNumber.Value = 1

80        End If

          g.Visible = False
          If cmbResultCount.Text <> "" Then
90          FillgDem (Trim(cmbResultCount.Text))
          End If
          g.Visible = True
100       FillResultGrid
110       FillCodeGrid
120       TransferDems

130       g.ColWidth(0) = 0
140       g.ColWidth(1) = 1600
150       g.ColWidth(2) = 600
160       g.ColWidth(3) = 1400
170       g.ColAlignment(3) = flexAlignLeftCenter
180       g.TextMatrix(0, 1) = "SAMPLE ID"
190       g.TextMatrix(1, 1) = "SAMPLE DATE"
200       g.TextMatrix(2, 1) = "RUN DATE"
210       g.TextMatrix(2, 2) = "S/T"
220       g.TextMatrix(2, 3) = "Ref Ranges"

230       Me.Refresh
240       FillAllResults
250       RemoveLines

          'FillG
260       FillCombos



270       pnlFetching.Visible = False

280       PBar.Max = LogOffDelaySecs
290       PBar = 0

300       lblNoRes = g.Cols - 4

310       Timer1.Enabled = True

320       Exit Sub

Form_Activate_Error:

          Dim strES As String
          Dim intEL As Integer



330       intEL = Erl
340       strES = Err.Description
350       LogError "frmFullImm", "Form_Activate", intEL, strES

End Sub
Private Sub RemoveLines()

          Dim x As Integer
          Dim Y As Integer
          Dim Found As Boolean

10        For Y = g.Rows - 1 To 4 Step -1
20            Found = False
30            For x = 4 To g.Cols - 1
40                If g.TextMatrix(Y, x) <> "" Then
50                    Found = True
60                    Exit For
70                End If
80            Next
90            If Not Found Then
100               g.RemoveItem Y
110           End If
120       Next

End Sub

Private Sub FillAllResults()

      Dim tb As Recordset
      Dim sql As String
      Dim x As Integer
      Dim Y As Integer
      Dim SampleID As Long
      Dim strSex As String
      Dim PerCent As Integer
      Dim DaysOld As Long

10    On Error GoTo FillAllResults_Error

20    If Left$(lblSex, 1) = "M" Then
30        strSex = "CT.MaleLow as Low, CT.MaleHigh as High "
40    ElseIf Left$(lblSex, 1) = "F" Then
50        strSex = "CT.FemaleLow as Low, CT.FemaleHigh as High "
60    Else
70        strSex = "CT.FemaleLow as Low, CT.MaleHigh as High "
80    End If

90    If Not IsDate(lblDoB) Then
100       DaysOld = 25 * 365.25
110   Else
120       DaysOld = DateDiff("d", lblDoB, Now)
130   End If

140   PerCent = 0

150   For x = 4 To g.Cols - 1

160       If x = 6 Then
170           Me.Refresh
180       End If

190       g.ColWidth(x) = 1500

200       PerCent = (x / g.Cols) * 100
210       pnlFetching.FloodPercent = PerCent

220       SampleID = Val(g.TextMatrix(0, x))

230       sql = "SELECT " & _
                "CASE WHEN ISNUMERIC(CR.Result) = 1 AND CR.Result <> '.' THEN " & _
                "  STR(CONVERT(FLOAT,CR.Result) + 0.00000001, 6, CT.DP)  " & _
                "ELSE " & _
                "  CR.Result " & _
                "END AS Result, CR.Valid, CT.LongName, CR.Code, CR.SampleType, " & _
                "CT.PlausibleLow, CT.PlausibleHigh, " & _
                strSex & _
                "FROM ImmResults as CR, ImmTestDefinitions as CT " & _
                "WHERE SampleID = '" & SampleID & "' " & _
                "AND CR.Code = CT.Code " & _
                "AND CR.SampleType = CT.SampleType " & _
                "AND CT.AgeToDays >= " & DaysOld & " " & _
                "AND CT.AgeFromDays <= " & DaysOld & " "
240       Set tb = New Recordset
250       RecOpenClient 0, tb, sql
260       Do While Not tb.EOF

270           For Y = 4 To g.Rows - 1
280               If tb!Code = g.TextMatrix(Y, 0) Then    'And tb!SampleType = g.TextMatrix(y, 2) Then

290                   g.TextMatrix(Y, x) = tb!Result & ""
                      '----------------------------------------------------
300                   If IsResultAmended("Imm", SampleID, tb!Code, tb!Result) = True Then
310                       g.Col = x
320                       g.Row = Y
330                       g.CellFontUnderline = True
340                   End If
                      '====================================================
350                   If IsNumeric(tb!Result & "") Then
360                       If Val(tb!Result & "") > tb!PlausibleHigh Then
370                           g.Col = x
380                           g.Row = Y
390                           g.CellBackColor = vbBlue
400                           g.CellForeColor = vbYellow
410                           g.TextMatrix(Y, x) = "*****"
420                       ElseIf Val(tb!Result) < tb!PlausibleLow Then
430                           g.Col = x
440                           g.Row = Y
450                           g.CellBackColor = vbBlack
460                           g.CellForeColor = vbYellow
470                           g.TextMatrix(Y, x) = "*****"
480                       ElseIf Val(tb!Result) > tb!High Then
490                           fraHL.Visible = True
500                           g.Col = x
510                           g.Row = Y
520                           g.CellBackColor = SysOptHighBack(0)
530                           g.CellForeColor = SysOptHighFore(0)
540                       ElseIf Val(tb!Result) < tb!Low Then
550                           fraHL.Visible = True
560                           g.Col = x
570                           g.Row = Y
580                           g.CellBackColor = SysOptLowBack(0)
590                           g.CellForeColor = SysOptLowFore(0)
600                       End If
610                   End If
620                   If Not tb!Valid = 1 Then
630                       g.TextMatrix(Y, x) = g.TextMatrix(Y, x) & " NV"
640                   End If
650                   Exit For
660               End If
670           Next
680           tb.MoveNext
690       Loop
700   Next

710   Exit Sub

FillAllResults_Error:

      Dim strES As String
      Dim intEL As Integer

720   intEL = Erl
730   strES = Err.Description
740   LogError "frmFullImm", "FillAllResults", intEL, strES, sql

End Sub

Private Sub TransferDems()

          Dim gDemY As Integer
          Dim gResX As Integer

10        For gDemY = 1 To gDem.Rows - 1
20            gResX = gDemY + 3
30            g.TextMatrix(0, gResX) = gDem.TextMatrix(gDemY, 0)    'SampleID
40            g.TextMatrix(1, gResX) = gDem.TextMatrix(gDemY, 1)    'SampleDate
50            g.TextMatrix(2, gResX) = gDem.TextMatrix(gDemY, 2)    'SampleTime
60        Next

End Sub

Private Sub FillCodeGrid()

          Dim tb As Recordset
          Dim tbref As Recordset
          Dim sql As String
          Dim s As String
          Dim strSex As String
          Dim DaysOld As Long

10        On Error GoTo FillCodeGrid_Error

20        With g
30            .Rows = 5
40            .AddItem ""
50            .RemoveItem 4
60        End With

70        If Not IsDate(lblDoB) Then
80            DaysOld = 25 * 365.25
90        Else
100           DaysOld = DateDiff("d", lblDoB, Now)
110       End If

          'Get Codes and LongNames
120       sql = "SELECT D.PrintPriority, D.LongName, D.Code, D.SampleType " & _
                "FROM ImmResults AS R, ImmTestDefinitions AS D, Demographics AS G WHERE " & _
                "R.Code = D.Code " & _
                "AND R.SampleID = G.SampleID " & _
                "AND PatName = '" & AddTicks(lblName) & "' "
130       If lblChart <> "" And chkChartNumber = 0 Then
140           sql = sql & "AND Chart = '" & lblChart & "' "
150       End If


160       If IsDate(lblDoB) Then
170           sql = sql & "AND DoB = '" & Format(lblDoB, "dd/mmm/yyyy") & "' "
180       Else
190           sql = sql & "AND (DoB IS NULL OR DoB = '') "
200       End If
210       sql = sql & "GROUP BY D.PrintPriority, D.LongName, D.Code, D.SampleType"

220       Set tb = New Recordset
230       RecOpenClient 0, tb, sql
240       Do While Not tb.EOF
250           If Left$(lblSex, 1) = "M" Then
260               strSex = "MaleLow as Low, MaleHigh as High "
270           ElseIf Left$(lblSex, 1) = "F" Then
280               strSex = "FemaleLow as Low, FemaleHigh as High "
290           Else
300               strSex = "FemaleLow as Low, MaleHigh as High "
310           End If
320           sql = "SELECT  " & strSex & " FROM  ImmTestDefinitions WHERE " & _
                    "Code = '" & tb!Code & "' " & _
                    "AND SampleType = '" & tb!SampleType & "' " & _
                    "AND AgeToDays >= " & DaysOld & " " & _
                    "AND AgeFromDays <= " & DaysOld & " "

              '"AND InUse = 1"
330           Set tbref = New Recordset
340           RecOpenClient 0, tbref, sql

350           s = Trim$(tb!Code & "") & vbTab & _
                  tb!LongName & vbTab & _
                  tb!SampleType & vbTab & _
                  tbref!Low & "-" & tbref!High & ""
360           g.AddItem s
370           tb.MoveNext
380       Loop

390       If g.Rows > 5 Then
400           g.RemoveItem 4
410       End If

420       Exit Sub

FillCodeGrid_Error:

          Dim strES As String
          Dim intEL As Integer

430       intEL = Erl
440       strES = Err.Description
450       LogError "frmFullImm", "FillCodeGrid", intEL, strES, sql

End Sub


Private Sub FillgDem(p_RCount As String)

          Dim n As Integer
          Dim tb As Recordset
          Dim sql As String
          Dim s As String
          Dim xdate As String

10        On Error GoTo FillgDem_Error

20        With gDem
30            .Visible = False
40            .Rows = 2
50            .AddItem ""
60            .RemoveItem 1
70        End With
80        p_RCount = UCase(p_RCount)
90        If p_RCount = "First 5" Then
100           sql = "SELECT DISTINCT top 5 (D.SampleID), D.SampleDate, D.RunDate, " & _
                  "(SELECT Top 1 Q.RunTime FROM ImmResults Q WHERE Q.SampleID = D.SampleID ORDER BY Q.RunTime) RunDateTime " & _
                  "FROM Demographics AS D, ImmResults WHERE ("
110       ElseIf p_RCount = "First 10" Then
120           sql = "SELECT DISTINCT top 10 (D.SampleID), D.SampleDate, D.RunDate, " & _
                  "(SELECT Top 1 Q.RunTime FROM ImmResults Q WHERE Q.SampleID = D.SampleID ORDER BY Q.RunTime) RunDateTime " & _
                  "FROM Demographics AS D, ImmResults WHERE ("
130       ElseIf p_RCount = "First 20" Then
140           sql = "SELECT DISTINCT top 20 (D.SampleID), D.SampleDate, D.RunDate, " & _
                  "(SELECT Top 1 Q.RunTime FROM ImmResults Q WHERE Q.SampleID = D.SampleID ORDER BY Q.RunTime) RunDateTime " & _
                  "FROM Demographics AS D, ImmResults WHERE ("
150       ElseIf p_RCount = "First 50" Then
160           sql = "SELECT DISTINCT top 50 (D.SampleID), D.SampleDate, D.RunDate, " & _
                  "(SELECT Top 1 Q.RunTime FROM ImmResults Q WHERE Q.SampleID = D.SampleID ORDER BY Q.RunTime) RunDateTime " & _
                  "FROM Demographics AS D, ImmResults WHERE ("
170       ElseIf p_RCount = "ALL" Then
180           sql = "SELECT DISTINCT (D.SampleID), D.SampleDate, D.RunDate, " & _
                  "(SELECT Top 1 Q.RunTime FROM ImmResults Q WHERE Q.SampleID = D.SampleID ORDER BY Q.RunTime) RunDateTime " & _
                  "FROM Demographics AS D, ImmResults WHERE ("
190       End If
200       If Trim(lblChart) <> "" And chkChartNumber.Value = 0 Then
210           sql = sql & "(D.Chart = '" & lblChart & "') AND"
220       End If
          '+++ Junaid 10-08-2023
          '120       sql = sql & " (D.PatName = '" & AddTicks(lblName) & "' " & _
          '                "AND D.DoB  = '" & Format(lblDoB, "dd/MMM/yyyy") & "') ) " & _
          '                "AND D.SampleID = ImmResults.SampleID order by RunDateTime Desc"
230       sql = sql & " (D.PatName = '" & AddTicks(lblName) & "' " & _
              "AND D.DoB  = '" & Format(lblDoB, "dd/MMM/yyyy") & "') ) " & _
              "AND D.SampleID = ImmResults.SampleID order by D.SampleDate Desc"
          '--- Junaid
          'sql = "SELECT  D.SampleID, D.Chart, D.dob, D.SampleDate FROM demographics D " & _
          '      "JOIN ImmResults I ON D.SampleID = I.SampleID " & _
          '      "WHERE D.dob='" & Format(lblDoB, "dd/MMM/yyyy") & "' AND D.patname = '" & AddTicks(lblName) & "' " & _
          '      "UNION " & _
          '      "SELECT  D.SampleID, D.Chart, D.dob, D.SampleDate FROM demographics D " & _
          '      "JOIN ImmResults I ON D.SampleID = I.SampleID " & _
          '      "WHERE D.chart in " & _
          '      "(SELECT distinct D.chart FROM demographics WHERE dob='" & Format(lblDoB, "dd/MMM/yyyy") & "' " & _
          '      "AND D.patname = '" & AddTicks(lblName) & "' AND COALESCE(D.chart,'')<>'') " & _
          '      "ORDER BY D.sampledate DESC"

240       Set tb = New Recordset
250       RecOpenClient 0, tb, sql
260       Do While Not tb.EOF
270           s = tb!SampleID & vbTab
              'Sample DateTime
280           If Not IsNull(tb!SampleDate) Then
290               xdate = Format(tb!SampleDate, "dd/mm/yy")
300               If Format(tb!SampleDate, "hh:mm") <> "00:00" Then
310                   xdate = xdate & " " & Format(tb!SampleDate, "hh:mm")
320               End If
330           Else
340               xdate = ""
350           End If
360           s = s & xdate & vbTab
              'Run DateTime
370           If tb!RunDateTime <> "" Then
380               xdate = Format(tb!RunDateTime, "dd/mm/yy")
390               If Format(tb!RunDateTime, "hh:mm") <> "00:00" Then
400                   xdate = xdate & " " & Format(tb!RunDateTime, "hh:mm")
410               End If
420           Else
430               xdate = ""
440           End If
450           s = s & xdate & vbTab
460           s = s & vbTab & Format$(n)
470           gDem.AddItem s
480           tb.MoveNext
490       Loop

500       With gDem
510           If .Rows > 2 Then
520               .RemoveItem 1
                  '.Col = 1
                  '.Sort = 9
530               .Visible = True
540           End If
550       End With

560       Exit Sub

FillgDem_Error:

          Dim strES As String
          Dim intEL As Integer

570       intEL = Erl
580       strES = Err.Description
590       LogError "frmFullImm", "FillgDem", intEL, strES, sql

End Sub

Private Sub FillResultGrid()

          Dim sql As String
          Dim intCols As Integer

          'Get number of columns

10        On Error GoTo FillResultGrid_Error

20        intCols = gDem.Rows + 3

30        MaxPrevious = intCols - 3
40        udPrevious.Max = MaxPrevious
50        If MaxPrevious < 7 Then
60            lblPrevious = Format$(MaxPrevious)
70            udPrevious.Value = MaxPrevious
80        Else
90            lblPrevious = "6"
100           udPrevious.Value = 6
110       End If

120       g.Cols = intCols

130       Exit Sub

FillResultGrid_Error:

          Dim strES As String
          Dim intEL As Integer

140       intEL = Erl
150       strES = Err.Description
160       LogError "frmFullImm", "FillResultGrid", intEL, strES, sql

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
110       LogError "frmFullImm", "GetRow", intEL, strES


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
150       LogError "frmFullImm", "InList", intEL, strES


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
80            sql = "SELECT * from immtestdefinitions WHERE  code = '" & List1.List(n) & "' and  (agefromdays <= " & DaysOld & " and agetodays >= " & DaysOld & ")"
90            Set sn = New Recordset
100           RecOpenServer Tn, sn, sql
110           If Not sn.EOF Then

120               g = sn!LongName & ""
130               g.TextMatrix(g.Row, 1) = sn!SampleType
140               If sn!PrnRR & "" <> False Then
150                   If Left(sex, 1) = "F" Then
160                       g.TextMatrix(g.Row, 2) = Trim(sn!FemaleLow) & " - " & Trim(sn!FemaleHigh)
170                   ElseIf Left(sex, 1) = "M" Then
180                       g.TextMatrix(g.Row, 2) = Trim(sn!MaleLow) & " - " & Trim(sn!MaleHigh)
190                   Else
200                       g.TextMatrix(g.Row, 2) = Trim(sn!FemaleLow) & " - " & Trim(sn!MaleHigh)
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
280       LogError "frmFullImm", "TransferListToGrid", intEL, strES, sql


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
60        LogError "frmFullImm", "Form_Deactivate", intEL, strES


End Sub

Private Sub Form_Load()

10        lblHigh.BackColor = SysOptHighBack(0)
20        lblHigh.ForeColor = SysOptHighFore(0)
30        lblLow.BackColor = SysOptLowBack(0)
40        lblLow.ForeColor = SysOptLowFore(0)

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
60        LogError "frmFullImm", "Form_MouseMove", intEL, strES


End Sub

Private Sub g_Click()
          Dim x As Long
          Dim Y As Long

10        On Error GoTo g_Click_Error

20        x = g.RowSel
30        Y = g.ColSel

40        If Y > 0 Then
50            If Trim(g.TextMatrix(x, Y)) <> "" Then g.ToolTipText = g.TextMatrix(x, Y)
60        End If

70        DrawChart

80        Exit Sub

g_Click_Error:

          Dim strES As String
          Dim intEL As Integer



90        intEL = Erl
100       strES = Err.Description
110       LogError "frmFullImm", "g_Click", intEL, strES


End Sub


Private Sub g_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)



10        On Error GoTo g_MouseMove_Error

20        Y = g.MouseCol
30        x = g.MouseRow

40        g.ToolTipText = ""

50        If Y > 0 And Not IsNumeric(g.TextMatrix(x, Y)) Then
60            g.ToolTipText = g.TextMatrix(x, Y)
70        End If


80        PBar = 0

90        Exit Sub

g_MouseMove_Error:

          Dim strES As String
          Dim intEL As Integer



100       intEL = Erl
110       strES = Err.Description
120       LogError "frmFullImm", "g_MouseMove", intEL, strES


End Sub


Private Sub gDem_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)

          Dim d1 As String
          Dim d2 As String

10        If Not IsDate(gDem.TextMatrix(Row1, 1)) Then
20            Cmp = 0
30            Exit Sub
40        End If

50        If Not IsDate(gDem.TextMatrix(Row2, 1)) Then
60            Cmp = 0
70            Exit Sub
80        End If

90        d1 = gDem.TextMatrix(Row1, 1)
100       d2 = gDem.TextMatrix(Row2, 1)

110       Cmp = Sgn(DateDiff("s", d1, d2))

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
60        LogError "frmFullImm", "Image1_MouseMove", intEL, strES


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
60        LogError "frmFullImm", "Label1_MouseMove", intEL, strES


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
60        LogError "frmFullImm", "Label2_MouseMove", intEL, strES


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
60        LogError "frmFullImm", "Label3_MouseMove", intEL, strES


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
60        LogError "frmFullImm", "Label4_MouseMove", intEL, strES


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
60        LogError "frmFullImm", "lblChart_MouseMove", intEL, strES


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
60        LogError "frmFullImm", "lblDoB_MouseMove", intEL, strES


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
60        LogError "frmFullImm", "lblName_MouseMove", intEL, strES


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
190       LogError "frmFullImm", "pb_MouseMove", intEL, strES


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
90        LogError "frmFullImm", "Timer1_Timer", intEL, strES


End Sub


