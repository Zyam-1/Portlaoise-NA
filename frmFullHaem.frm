VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFullHaem 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Full Haematology History"
   ClientHeight    =   7095
   ClientLeft      =   75
   ClientTop       =   390
   ClientWidth     =   11685
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
   Icon            =   "frmFullHaem.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7095
   ScaleWidth      =   11685
   Begin VB.ComboBox cmbResultCount 
      Height          =   315
      ItemData        =   "frmFullHaem.frx":000C
      Left            =   600
      List            =   "frmFullHaem.frx":0022
      TabIndex        =   37
      Text            =   "ALL"
      Top             =   30
      Width           =   1215
   End
   Begin VB.Frame FrameRange 
      Caption         =   "DateRange"
      Height          =   870
      Left            =   6885
      TabIndex        =   31
      Top             =   1080
      Width           =   4785
      Begin VB.CommandButton cmdRefresh 
         Height          =   615
         Left            =   3870
         Picture         =   "frmFullHaem.frx":0054
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Click to refresh biochemistry history"
         Top             =   180
         Width           =   735
      End
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   330
         Left            =   135
         TabIndex        =   32
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
         Format          =   232062979
         CurrentDate     =   38629
      End
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   330
         Left            =   2070
         TabIndex        =   33
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
         Format          =   232062979
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
         TabIndex        =   35
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
         TabIndex        =   34
         Top             =   165
         Width           =   330
      End
   End
   Begin VB.CommandButton cmdCancel 
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
      Left            =   10320
      Picture         =   "frmFullHaem.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Exit"
      Top             =   6045
      Width           =   1245
   End
   Begin VB.CheckBox chkChartNumber 
      Caption         =   "Ignore chart number"
      Height          =   255
      Left            =   7260
      TabIndex        =   27
      Top             =   720
      Width           =   2535
   End
   Begin VB.CommandButton cmdExcel 
      Caption         =   "Export"
      Height          =   750
      Left            =   7560
      Picture         =   "frmFullHaem.frx":0C28
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   6045
      Width           =   1245
   End
   Begin VB.CommandButton cmdPrint 
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
      Left            =   8940
      Picture         =   "frmFullHaem.frx":6846
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   6045
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
      Left            =   6885
      TabIndex        =   11
      Top             =   2355
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
         Picture         =   "frmFullHaem.frx":6B50
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Draw Graph"
         Top             =   120
         Width           =   1035
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
         Left            =   2010
         TabIndex        =   13
         Top             =   240
         Width           =   1455
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
         TabIndex        =   12
         Top             =   240
         Width           =   1725
      End
   End
   Begin MSComctlLib.ProgressBar PBar 
      Height          =   195
      Left            =   6750
      TabIndex        =   10
      Top             =   6930
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   7020
      Top             =   6375
   End
   Begin VB.PictureBox pb 
      AutoRedraw      =   -1  'True
      Height          =   2325
      Left            =   6885
      ScaleHeight     =   2265
      ScaleWidth      =   4125
      TabIndex        =   8
      Top             =   3315
      Width           =   4185
      Begin VB.Label lblTest 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1395
         TabIndex        =   21
         Top             =   0
         Width           =   1320
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdHaem 
      Height          =   6675
      Left            =   45
      TabIndex        =   7
      Top             =   360
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   11774
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
      Height          =   2565
      Left            =   2700
      TabIndex        =   6
      Top             =   1980
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "Show "
      Height          =   195
      Left            =   60
      TabIndex        =   39
      Top             =   90
      Width           =   450
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "Results"
      Height          =   195
      Left            =   1920
      TabIndex        =   38
      Top             =   90
      Width           =   645
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
      Left            =   4095
      TabIndex        =   30
      Top             =   45
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   6780
      Picture         =   "frmFullHaem.frx":6E5A
      Top             =   1965
      Width           =   480
   End
   Begin VB.Label lblExcelInfo 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exporting..."
      Height          =   285
      Left            =   7560
      TabIndex        =   24
      Top             =   5745
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
      Left            =   10320
      TabIndex        =   22
      Top             =   5715
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
      Left            =   8925
      TabIndex        =   23
      Top             =   5715
      Width           =   1380
   End
   Begin VB.Label Tn 
      Height          =   225
      Left            =   10365
      TabIndex        =   20
      Top             =   4725
      Width           =   315
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
      Left            =   10800
      TabIndex        =   19
      Top             =   390
      Width           =   510
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
      Left            =   10320
      TabIndex        =   18
      Top             =   420
      Width           =   405
   End
   Begin VB.Label lblNopas 
      Height          =   255
      Left            =   10005
      TabIndex        =   17
      Top             =   3795
      Width           =   825
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
      Left            =   11085
      TabIndex        =   16
      Top             =   5385
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
      Left            =   11085
      TabIndex        =   15
      Top             =   4365
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
      Left            =   11085
      TabIndex        =   14
      Top             =   3315
      Width           =   510
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
      Left            =   7305
      TabIndex        =   9
      Top             =   2055
      Width           =   2865
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
      Left            =   6810
      TabIndex        =   5
      Top             =   390
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
      Left            =   9390
      TabIndex        =   4
      Top             =   120
      Width           =   315
   End
   Begin VB.Label lblName 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   7260
      TabIndex        =   3
      Top             =   390
      Width           =   2985
   End
   Begin VB.Label lblDoB 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   9750
      TabIndex        =   2
      Top             =   90
      Width           =   1815
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
      Left            =   6840
      TabIndex        =   1
      Top             =   150
      Width           =   375
   End
   Begin VB.Label lblChart 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   7260
      TabIndex        =   0
      Top             =   90
      Width           =   1545
   End
End
Attribute VB_Name = "frmFullHaem"
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


Private Sub chkChartNumber_Click()
On Error GoTo chkChartNumber_Click_Error

grdHaem.Visible = False
If cmbResultCount.Text <> "" Then
    FillG (Trim(cmbResultCount.Text))
End If
grdHaem.Visible = True

Exit Sub
chkChartNumber_Click_Error:
grdHaem.Visible = True
LogError "frmFullHaem", "chkChartNumber_Click", Erl, Err.Description
End Sub

Private Sub cmbResultCount_Change()
    On Error GoTo cmbResultCount_Change_Error
    
    grdHaem.Visible = False
     
    If cmbResultCount.Text <> "" Then
         FillG (Trim(cmbResultCount.Text))
    End If
    
    grdHaem.Visible = True
    
cmbResultCount_Change_Error:

    Dim strES As String
    Dim intEL As Integer
    
    intEL = Erl
    strES = Err.Description
    LogError "frmFullBio", "cmbResultCount_Change", intEL, strES
End Sub

Private Sub cmbResultCount_Click()
    On Error GoTo cmbResultCount_Click_Error
    
    grdHaem.Visible = False
     
    If cmbResultCount.Text <> "" Then
         FillG (Trim(cmbResultCount.Text))
    End If
    
    grdHaem.Visible = True
    
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

20        strHeading = "Haematology History" & vbCr
30        strHeading = strHeading & "Patient Name: " & lblName & vbCr
40        strHeading = strHeading & vbCr

50        ExportFlexGrid grdHaem, Me, strHeading

60        Exit Sub

cmdExcel_Click_Error:

          Dim strES As String
          Dim intEL As Integer

70        intEL = Erl
80        strES = Err.Description
90        LogError "frmFullHaem", "cmdExcel_Click", intEL, strES

End Sub

Private Sub cmdPrint_Click()
          Dim Y As Integer
          Dim z As Long
          Dim MaxCols As Integer

10        On Error GoTo cmdPrint_Click_Error




20        Printer.Orientation = vbPRORLandscape
30        Printer.Font = "Courier New"
40        Printer.Font.Size = 16
50        Printer.Print Tab(10); "Cumulative Report from Haematology Dept."
60        Printer.Print

70        Printer.Font.Size = 14
80        Printer.Print Tab(10); "Name : " & lblName;
90        Printer.Print Tab(40); "Dob  : " & lblDoB

100       Printer.Print

110       With grdHaem
120           If .Cols > 8 Then
130               MaxCols = 7
140           Else
150               MaxCols = .Cols - 1
160           End If

              'Add seperator
170           PrintText String(217, "-") & vbCrLf, 4, True
              'Print SampleID row
180           PrintText FormatString(.TextMatrix(0, 0), 16, "|"), 10, True
190           For z = 1 To MaxCols
200               PrintText FormatString(.TextMatrix(0, z), 9, "|", AlignCenter), 10
210           Next z
220           PrintText vbCrLf
              'Print Sample Date Row
230           PrintText FormatString(.TextMatrix(1, 0), 16, "|"), 10, True
240           For z = 1 To MaxCols
250               PrintText FormatString(Format(.TextMatrix(1, z), "dd/MM/yy"), 9, "|", AlignCenter), 10
260           Next z
270           PrintText vbCrLf
              'Print Sample Time Row is exists
280           PrintText FormatString("SAMPLE TIME", 16, "|"), 10, True
290           For z = 1 To MaxCols
300               If Format(.TextMatrix(1, z), "hh:mm") = "00:00" Then
310                   PrintText FormatString("", 9, "|", AlignCenter), 10
320               Else
330                   PrintText FormatString(Format(.TextMatrix(1, z), "hh:mm"), 9, "|", AlignCenter), 10
340               End If
350           Next z
360           PrintText vbCrLf

              'Print Run Date Row
370           PrintText FormatString(.TextMatrix(2, 0), 16, "|"), 10, True
380           For z = 1 To MaxCols
390               PrintText FormatString(Format(.TextMatrix(2, z), "dd/MM/yy"), 9, "|", AlignCenter), 10
400           Next z
410           PrintText vbCrLf
              'Print Run Time Row is exists

420           PrintText FormatString("RUN TIME", 16, "|"), 10, True
430           For z = 1 To MaxCols
440               If Format(.TextMatrix(2, z), "hh:mm") = "00:00" Then
450                   PrintText FormatString("", 9, "|", AlignCenter), 10
460               Else
470                   PrintText FormatString(Format(.TextMatrix(2, z), "hh:mm"), 9, "|", AlignCenter), 10
480               End If
490           Next z
500           PrintText vbCrLf
              'Add seperator
510           PrintText String(217, "-") & vbCrLf, 4, True
              'Print results
520           For Y = 3 To .Rows - 1
530               PrintText FormatString(.TextMatrix(Y, 0), 16, "|"), 10
540               For z = 1 To MaxCols
550                   PrintText FormatString(.TextMatrix(Y, z), 9, "|", AlignCenter), 10
560               Next z
570               PrintText vbCrLf
580           Next Y
              'Add seperator
590           PrintText String(217, "-"), 4, True
600       End With
610       Printer.EndDoc

          'Printer.Font.Size = 10
          'For n = 0 To grdHaem.Rows - 1
          '    grdHaem.Row = n
          '    For z = 0 To x - 1
          '        grdHaem.Col = z
          '        Printer.Print Tab(10 * z); grdHaem;
          '    Next
          '    Printer.Print
          'Next

          'Printer.Print Tab(30); "End of Report"







620       Exit Sub

cmdPrint_Click_Error:

          Dim strES As String
          Dim intEL As Integer



630       intEL = Erl
640       strES = Err.Description
650       LogError "frmFullHaem", "cmdPrint_Click", intEL, strES


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

70        NumberOfDays = DateDiff("d", Format(cmbPlotFrom, "dd/mmm/yyyy"), Format(cmbPlotTo, "dd/mmm/yyyy"))
80        If NumberOfDays < 1 Then Exit Sub
90        ReDim ChartPositions(0 To NumberOfDays)

100       For n = 1 To NumberOfDays
110           ChartPositions(n).xPos = 0
120           ChartPositions(n).yPos = 0
130           ChartPositions(n).Value = 0
140           ChartPositions(n).Date = ""
150       Next

160       For n = 1 To grdHaem.Cols - 1
170           If Format$(grdHaem.TextMatrix(1, n), "dd/mmm/yyyy") = cmbPlotTo Then StartGridX = n
180           If Format$(grdHaem.TextMatrix(1, n), "dd/mmm/yyyy") = cmbPlotFrom Then StopGridX = n
190       Next

200       FirstDayFilled = False
210       Counter = 0
220       For x = StartGridX To StopGridX
230           If grdHaem.TextMatrix(grdHaem.Row, x) <> "" Then
240               If Not FirstDayFilled Then
250                   FirstDayFilled = True
260                   MaxVal = Val(grdHaem.TextMatrix(grdHaem.Row, x))
270                   ChartPositions(NumberOfDays).Date = Format(grdHaem.TextMatrix(1, x), "dd/mmm/yyyy")
280                   ChartPositions(NumberOfDays).Value = Val(grdHaem.TextMatrix(grdHaem.Row, x))
290               Else
300                   DaysInterval = Abs(DateDiff("D", cmbPlotTo, Format$(grdHaem.TextMatrix(1, x), "dd/mmm/yyyy")))
310                   ChartPositions(NumberOfDays - DaysInterval).Date = grdHaem.TextMatrix(1, x)
320                   cVal = Val(grdHaem.TextMatrix(grdHaem.Row, x))
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

660       lblTest = grdHaem.TextMatrix(grdHaem.RowSel, 0)

670       Exit Sub

DrawChart_Error:

          Dim strES As String
          Dim intEL As Integer

680       intEL = Erl
690       strES = Err.Description
700       LogError "frmFullHaem", "DrawChart", intEL, strES

End Sub

Private Sub cmdCancel_Click()

10        Unload Me

End Sub

Private Sub FillG(p_RCount As String)

          Dim sn As New Recordset
          Dim tb As New Recordset
          Dim sql As String
          Dim gcolumns As Long
          Dim x As Long
          Dim xrun As String
          Dim xdate As String
          Dim n As Integer



10        On Error GoTo FillG_Error

20        With grdHaem
30            .Rows = 4
40            .AddItem ""
50            .RemoveItem 3
60            .Visible = False
70        End With
80        p_RCount = UCase(p_RCount)
90        If p_RCount = "First 5" Then
100           sql = "SELECT DISTINCT top 5 (D.SampleID), D.RunDate, D.TimeTaken, D.SampleDate, " & _
                  "(SELECT Top 1 Q.RunDateTime FROM HaemResults Q WHERE Q.SampleID = D.SampleID ORDER BY Q.RunDateTime) RunDateTime " & _
                  "FROM Demographics D, HaemResults R WHERE ("
110       ElseIf p_RCount = "First 10" Then
120           sql = "SELECT DISTINCT top 10 (D.SampleID), D.RunDate, D.TimeTaken, D.SampleDate, " & _
                  "(SELECT Top 1 Q.RunDateTime FROM HaemResults Q WHERE Q.SampleID = D.SampleID ORDER BY Q.RunDateTime) RunDateTime " & _
                  "FROM Demographics D, HaemResults R WHERE ("
130       ElseIf p_RCount = "First 20" Then
140           sql = "SELECT DISTINCT top 20 (D.SampleID), D.RunDate, D.TimeTaken, D.SampleDate, " & _
                  "(SELECT Top 1 Q.RunDateTime FROM HaemResults Q WHERE Q.SampleID = D.SampleID ORDER BY Q.RunDateTime) RunDateTime " & _
                  "FROM Demographics D, HaemResults R WHERE ("
150       ElseIf p_RCount = "First 50" Then
160           sql = "SELECT DISTINCT top 50 (D.SampleID), D.RunDate, D.TimeTaken, D.SampleDate, " & _
                  "(SELECT Top 1 Q.RunDateTime FROM HaemResults Q WHERE Q.SampleID = D.SampleID ORDER BY Q.RunDateTime) RunDateTime " & _
                  "FROM Demographics D, HaemResults R WHERE ("
170       ElseIf p_RCount = "ALL" Then
180           sql = "SELECT DISTINCT (D.SampleID), D.RunDate, D.TimeTaken, D.SampleDate, " & _
                  "(SELECT Top 1 Q.RunDateTime FROM HaemResults Q WHERE Q.SampleID = D.SampleID ORDER BY Q.RunDateTime) RunDateTime " & _
                  "FROM Demographics D, HaemResults R WHERE ("
190       End If

200       If Trim(lblChart) <> "" And chkChartNumber.Value = 0 Then
210           sql = sql & "(D.Chart = '" & lblChart & "') AND "
220       End If
230       sql = sql & "(D.PatName = '" & AddTicks(lblName) & "' " & _
              "AND D.DoB  = '" & Format$(lblDoB, "dd/MMM/yyyy") & "') " & _
              "AND D.RunDate BETWEEN'" & Format$(dtFrom, "dd/MMM/yyyy") & "' " & _
              "                AND '" & Format(dtTo + 1, "dd/MMM/yyyy") & "') " & _
              "AND D.SampleID = R.SampleID " & _
              "ORDER BY D.SampleDate DESC, D.SampleID Desc"

240       Set sn = New Recordset
250       RecOpenServer Val(Tn), sn, sql
260       If Not sn.EOF Then
270           grdHaem.Visible = False
280           grdHaem.Cols = gcolumns + 1
290           grdHaem.ColWidth(0) = 1600
300           grdHaem.TextMatrix(0, 0) = "SAMPLE ID"
310           grdHaem.TextMatrix(1, 0) = "SAMPLE DATE"
320           grdHaem.TextMatrix(2, 0) = "RUN DATE"
330           grdHaem.Cols = 1
              'SampleID and sampledate across
340           Do While Not sn.EOF
350               grdHaem.Cols = grdHaem.Cols + 1
360               x = grdHaem.Cols - 1
370               grdHaem.ColWidth(x) = 1500
380               grdHaem.Col = x
390               xrun = sn!SampleID & ""
400               grdHaem.TextMatrix(0, x) = xrun
                  'Sample DateTime
410               If Not IsNull(sn!SampleDate) Then
420                   xdate = Format(sn!SampleDate, "dd/mm/yy")
430                   If Format(sn!SampleDate, "hh:mm") <> "00:00" Then
440                       xdate = xdate & " " & Format(sn!SampleDate, "hh:mm")
450                   End If
460               Else
470                   xdate = ""
480               End If
490               grdHaem.TextMatrix(1, x) = xdate
                  'Run DateTime
500               If sn!RunDateTime <> "" Then
510                   xdate = Format(sn!RunDateTime, "dd/mm/yy")
520                   If Format(sn!RunDateTime, "hh:mm") <> "00:00" Then
530                       xdate = xdate & " " & Format(sn!RunDateTime, "hh:mm")
540                   End If
550               Else
560                   xdate = ""
570               End If
580               grdHaem.TextMatrix(2, x) = xdate

590               sn.MoveNext
600           Loop

610           grdHaem.AddItem "WBC"
620           grdHaem.AddItem "RBC"
630           grdHaem.AddItem "Hgb"
640           grdHaem.AddItem "Hct"
650           grdHaem.AddItem "MCV"
660           grdHaem.AddItem "MCH"
670           grdHaem.AddItem "MCHC"
680           grdHaem.AddItem "CHCM"
690           grdHaem.AddItem "RDW"
700           grdHaem.AddItem "Plt"
710           grdHaem.AddItem "Neut A"
720           grdHaem.AddItem "Lymp A"
730           grdHaem.AddItem "Mono A"
740           grdHaem.AddItem "Eos A"
750           grdHaem.AddItem "Bas A"
760           grdHaem.AddItem "Luc"
770           grdHaem.AddItem "ESR"
780           grdHaem.AddItem "Monospot"
790           grdHaem.AddItem "Retic"
800           grdHaem.AddItem "RF"
810           grdHaem.AddItem "Film"
820           grdHaem.AddItem "Malaria Screen"    'QMS Ref #817977
830           grdHaem.AddItem "Sickle Screen"     'QMS Ref #817977


840           grdHaem.Row = 4
850           grdHaem.Col = 0
860           grdHaem.CellFontSize = 10
870           grdHaem.Row = 6
880           grdHaem.Col = 0
890           grdHaem.CellFontSize = 10
900           grdHaem.Row = 12
910           grdHaem.Col = 0
920           grdHaem.CellFontSize = 10


930           If Tn = "" Then Tn = 0

940           For x = 1 To grdHaem.Cols - 1
950               sql = "SELECT * from HaemResults WHERE " & _
                      "SampleID = '" & grdHaem.TextMatrix(0, x) & "'"
960               Set tb = New Recordset
970               RecOpenServer 0, tb, sql
980               If Not tb.EOF Then
990                   If grdHaem.TextMatrix(2, x) = "" Then grdHaem.TextMatrix(2, x) = Format(tb!RunDateTime, "hh:mm")
1000                  If UserMemberOf = "Secretarys" Or UserMemberOf = "LookUp" Or SysOptNoCumShow(0) Then
1010                      If Not IsNull(tb!Valid) And tb!Valid Then
1020                          grdHaem.TextMatrix(4, x) = Trim(tb!wbc & "")
1030                          ColouriseG "WBC", grdHaem, 4, x, Trim(tb!wbc & ""), lblSex, lblDoB
1040                          grdHaem.Row = 4
1050                          grdHaem.Col = x
1060                          grdHaem.CellFontSize = 12
1070                          grdHaem.TextMatrix(5, x) = Trim(tb!rbc & "")
1080                          ColouriseG "RBC", grdHaem, 5, x, Trim(tb!rbc & ""), lblSex, lblDoB
1090                          grdHaem.TextMatrix(6, x) = Trim(tb!Hgb & "")
1100                          ColouriseG "HGB", grdHaem, 6, x, Trim(tb!Hgb & ""), lblSex, lblDoB
1110                          grdHaem.Row = 6
1120                          grdHaem.Col = x
1130                          grdHaem.CellFontSize = 12

1140                          grdHaem.TextMatrix(7, x) = Trim(tb!hct & "")
1150                          ColouriseG "HCT", grdHaem, 7, x, Trim(tb!hct & ""), lblSex, lblDoB
1160                          grdHaem.TextMatrix(8, x) = Trim(tb!MCV & "")
1170                          ColouriseG "MCV", grdHaem, 8, x, Trim(tb!MCV & ""), lblSex, lblDoB
1180                          grdHaem.TextMatrix(9, x) = Trim(tb!mch & "")
1190                          ColouriseG "MC", grdHaem, 9, x, Trim(tb!mch & ""), lblSex, lblDoB, "MCH"
1200                          grdHaem.TextMatrix(10, x) = Trim(tb!mchc & "")
1210                          ColouriseG "MCHC", grdHaem, 10, x, Trim(tb!mchc & ""), lblSex, lblDoB, "MCHC"
1220                          grdHaem.TextMatrix(11, x) = Trim(tb!cH & "")
1230                          ColouriseG "CHCM", grdHaem, 11, x, Trim(tb!cH & ""), lblSex, lblDoB, "CH"
1240                          If Trim(tb!RDWCV & "") <> "" Then
1250                              grdHaem.TextMatrix(12, x) = Trim(tb!RDWCV & "")
1260                              ColouriseG "RDW", grdHaem, 12, x, Trim(tb!RDWCV & ""), lblSex, lblDoB, "RDWCV"
1270                          End If
1280                          grdHaem.TextMatrix(13, x) = Trim(tb!Plt & "")
1290                          ColouriseG "PLT", grdHaem, 13, x, Trim(tb!Plt & ""), lblSex, lblDoB
1300                          grdHaem.Row = 13
1310                          grdHaem.Col = x
1320                          grdHaem.CellFontSize = 12
1330                          grdHaem.TextMatrix(14, x) = Trim(tb!NeutA & "")
1340                          ColouriseG "NEUTA", grdHaem, 14, x, Trim(tb!NeutA & ""), lblSex, lblDoB
1350                          grdHaem.TextMatrix(15, x) = Trim(tb!LymA & "")
1360                          ColouriseG "LYMA", grdHaem, 15, x, Trim(tb!LymA & ""), lblSex, lblDoB
1370                          grdHaem.TextMatrix(16, x) = Trim(tb!MonoA & "")
1380                          ColouriseG "MONOA", grdHaem, 16, x, Trim(tb!MonoA & ""), lblSex, lblDoB
1390                          grdHaem.TextMatrix(17, x) = Trim(tb!EosA & "")
1400                          ColouriseG "EOSA", grdHaem, 17, x, Trim(tb!EosA & ""), lblSex, lblDoB
1410                          grdHaem.TextMatrix(18, x) = Trim(tb!BasA & "")
1420                          ColouriseG "BASA", grdHaem, 18, x, Trim(tb!BasA & ""), lblSex, lblDoB
1430                          grdHaem.TextMatrix(19, x) = Trim(tb!luca & "")
                              '--------------------------------------------------
1440                          If IsHaemResultAmended(grdHaem.TextMatrix(0, x), "LUCA") = True Then
1450                              grdHaem.Col = x
1460                              grdHaem.Row = 19
1470                              grdHaem.CellFontUnderline = True
1480                          End If
                              '=================================================
1490                          grdHaem.TextMatrix(20, x) = Trim(tb!esr & "")
                              '--------------------------------------------------
1500                          If IsHaemResultAmended(grdHaem.TextMatrix(0, x), "ESR") = True Then
1510                              grdHaem.Col = x
1520                              grdHaem.Row = 20
1530                              grdHaem.CellFontUnderline = True
1540                          End If
                              '=================================================
1550                          If tb!Monospot & "" <> "" Then
1560                              If tb!Monospot & "" = "N" Then
1570                                  grdHaem.TextMatrix(21, x) = "Negative"
1580                              ElseIf tb!Monospot = "P" Then
1590                                  grdHaem.TextMatrix(21, x) = "Positive"
1600                              ElseIf tb!Monospot = "I" Then
1610                                  grdHaem.TextMatrix(21, x) = "Inconclusive"
1620                              End If
1630                          End If
                              '--------------------------------------------------
1640                          If IsHaemResultAmended(grdHaem.TextMatrix(0, x), "MONOSPOT") = True Then
1650                              grdHaem.Col = x
1660                              grdHaem.Row = 21
1670                              grdHaem.CellFontUnderline = True
1680                          End If
                              '=================================================
                              'grdHaem.TextMatrix(21, X) = Trim(tb!Monospot & "")
1690                          grdHaem.TextMatrix(22, x) = Trim(tb!retics & "")
                              '--------------------------------------------------
1700                          If IsHaemResultAmended(grdHaem.TextMatrix(0, x), "RETICS") = True Then
1710                              grdHaem.Col = x
1720                              grdHaem.Row = 22
1730                              grdHaem.CellFontUnderline = True
1740                          End If
                              '=================================================
1750                          grdHaem.TextMatrix(23, x) = Trim(tb!tRa & "")
                              '--------------------------------------------------
1760                          If IsHaemResultAmended(grdHaem.TextMatrix(0, x), "TRA") = True Then
1770                              grdHaem.Col = x
1780                              grdHaem.Row = 23
1790                              grdHaem.CellFontUnderline = True
1800                          End If
                              '=================================================
1810                          If tb!cFilm = True Then grdHaem.TextMatrix(24, x) = "Yes"
                              '--------------------------------------------------
1820                          If IsHaemResultAmended(grdHaem.TextMatrix(0, x), "cFilm") = True Then
1830                              grdHaem.Col = x
1840                              grdHaem.Row = 24
1850                              grdHaem.CellFontUnderline = True
1860                          End If
                              '=================================================
1870                      Else
1880                          grdHaem.TextMatrix(4, x) = "Not"
1890                          grdHaem.TextMatrix(5, x) = "Valid"
1900                      End If
1910                  Else
1920                      grdHaem.TextMatrix(4, x) = tb!wbc & ""
1930                      ColouriseG "WBC", grdHaem, 4, x, Trim(tb!wbc & ""), lblSex, lblDoB
1940                      grdHaem.Row = 4
1950                      grdHaem.Col = x
1960                      grdHaem.CellFontSize = 12
1970                      grdHaem.TextMatrix(5, x) = tb!rbc & ""
1980                      ColouriseG "RBC", grdHaem, 5, x, Trim(tb!rbc & ""), lblSex, lblDoB
1990                      grdHaem.TextMatrix(6, x) = tb!Hgb & ""
2000                      ColouriseG "HGB", grdHaem, 6, x, Trim(tb!Hgb & ""), lblSex, lblDoB
2010                      grdHaem.Row = 6
2020                      grdHaem.Col = x
2030                      grdHaem.CellFontSize = 12
2040                      grdHaem.TextMatrix(7, x) = tb!hct & ""
2050                      ColouriseG "HCT", grdHaem, 7, x, Trim(tb!hct & ""), lblSex, lblDoB
2060                      grdHaem.TextMatrix(8, x) = tb!MCV & ""
2070                      ColouriseG "MCV", grdHaem, 8, x, Trim(tb!MCV & ""), lblSex, lblDoB
2080                      grdHaem.TextMatrix(9, x) = tb!mch & ""
2090                      ColouriseG "MC", grdHaem, 9, x, Trim(tb!mch & ""), lblSex, lblDoB, "MCH"
2100                      grdHaem.TextMatrix(10, x) = tb!mchc & ""
2110                      ColouriseG "MCHC", grdHaem, 10, x, Trim(tb!mchc & ""), lblSex, lblDoB
2120                      grdHaem.TextMatrix(11, x) = tb!cH & ""
2130                      ColouriseG "CHCM", grdHaem, 11, x, Trim(tb!cH & ""), lblSex, lblDoB, "CH"
2140                      If Trim(tb!RDWCV & "") <> "" Then
2150                          grdHaem.TextMatrix(12, x) = Trim(tb!RDWCV & "")
2160                          ColouriseG "RDW", grdHaem, 12, x, Trim(tb!RDWCV & ""), lblSex, lblDoB, "RDWCV"
2170                      End If
2180                      grdHaem.TextMatrix(13, x) = tb!Plt & ""
2190                      ColouriseG "PLT", grdHaem, 13, x, Trim(tb!Plt & ""), lblSex, lblDoB
2200                      grdHaem.Row = 13
2210                      grdHaem.Col = x
2220                      grdHaem.CellFontSize = 12
2230                      grdHaem.TextMatrix(14, x) = tb!NeutA & ""
2240                      ColouriseG "NEUTA", grdHaem, 14, x, Trim(tb!NeutA & ""), lblSex, lblDoB
2250                      grdHaem.TextMatrix(15, x) = tb!LymA & ""
2260                      ColouriseG "LYMA", grdHaem, 15, x, Trim(tb!LymA & ""), lblSex, lblDoB
2270                      grdHaem.TextMatrix(16, x) = tb!MonoA & ""
2280                      ColouriseG "MONOA", grdHaem, 16, x, Trim(tb!MonoA & ""), lblSex, lblDoB
2290                      grdHaem.TextMatrix(17, x) = tb!EosA & ""
2300                      ColouriseG "EOSA", grdHaem, 17, x, Trim(tb!EosA & ""), lblSex, lblDoB
2310                      grdHaem.TextMatrix(18, x) = tb!BasA & ""
2320                      ColouriseG "BASA", grdHaem, 18, x, Trim(tb!BasA & ""), lblSex, lblDoB
2330                      grdHaem.TextMatrix(19, x) = Trim(tb!luca & "")
                          '--------------------------------------------------
2340                      If IsHaemResultAmended(grdHaem.TextMatrix(0, x), "LUCA") = True Then
2350                          grdHaem.Col = x
2360                          grdHaem.Row = 19
2370                          grdHaem.CellFontUnderline = True
2380                      End If
                          '=================================================
2390                      grdHaem.TextMatrix(20, x) = tb!esr & ""
2400                      If tb!Monospot & "" <> "" Then
2410                          If tb!Monospot & "" = "N" Then
2420                              grdHaem.TextMatrix(21, x) = "Negative"
2430                          ElseIf tb!Monospot = "P" Then
2440                              grdHaem.TextMatrix(21, x) = "Positive"
2450                          ElseIf tb!Monospot = "I" Then
2460                              grdHaem.TextMatrix(21, x) = "Inconclusive"
2470                          End If
2480                      End If
                          '--------------------------------------------------
2490                      If IsHaemResultAmended(grdHaem.TextMatrix(0, x), "ESR") = True Then
2500                          grdHaem.Col = x
2510                          grdHaem.Row = 21
2520                          grdHaem.CellFontUnderline = True
2530                      End If
                          '=================================================
                          'grdHaem.TextMatrix(21, X) = tb!Monospot & ""
2540                      grdHaem.TextMatrix(22, x) = tb!reta & ""
                          '--------------------------------------------------
2550                      If IsHaemResultAmended(grdHaem.TextMatrix(0, x), "RETA") = True Then
2560                          grdHaem.Col = x
2570                          grdHaem.Row = 22
2580                          grdHaem.CellFontUnderline = True
2590                      End If
                          '=================================================
2600                      grdHaem.TextMatrix(23, x) = tb!tRa & ""
                          '--------------------------------------------------
2610                      If IsHaemResultAmended(grdHaem.TextMatrix(0, x), "tRa") = True Then
2620                          grdHaem.Col = x
2630                          grdHaem.Row = 23
2640                          grdHaem.CellFontUnderline = True
2650                      End If
                          '=================================================
2660                      If tb!cFilm = True Then grdHaem.TextMatrix(24, x) = "Yes"
                          '--------------------------------------------------
2670                      If IsHaemResultAmended(grdHaem.TextMatrix(0, x), "cFilm") = True Then
2680                          grdHaem.Col = x
2690                          grdHaem.Row = 24
2700                          grdHaem.CellFontUnderline = True
2710                      End If
                          '=================================================
2720                      grdHaem.TextMatrix(25, x) = tb!Malaria & ""
                          '--------------------------------------------------
2730                      If IsHaemResultAmended(grdHaem.TextMatrix(0, x), "MALARIA") = True Then
2740                          grdHaem.Col = x
2750                          grdHaem.Row = 25
2760                          grdHaem.CellFontUnderline = True
2770                      End If
                          '=================================================
2780                      grdHaem.TextMatrix(26, x) = tb!Sickledex & ""
                          '--------------------------------------------------
2790                      If IsHaemResultAmended(grdHaem.TextMatrix(0, x), "SICKLEDEX") = True Then
2800                          grdHaem.Col = x
2810                          grdHaem.Row = 26
2820                          grdHaem.CellFontUnderline = True
2830                      End If
                          '=================================================

2840                      If Not IsNull(tb!Valid) And tb!Valid = 1 Then

2850                      Else
2860                          grdHaem.TextMatrix(0, x) = grdHaem.TextMatrix(0, x) & " NV"
2870                      End If
2880                  End If
2890              End If
2900          Next
2910      End If

2920      lblNoRes = grdHaem.Cols - 1


2930      If grdHaem.Cols > 2 Then lblNoRes.Caption = grdHaem.Cols - 1 Else lblNoRes.Caption = 0

2940      If grdHaem.Rows > 4 Then
2950          grdHaem.RemoveItem 3
2960      End If
2970      grdHaem.Visible = True

2980      If SysOptNoSeeRF(0) Then
2990          grdHaem.RowHeight(22) = 0
3000      End If

3010      For n = 1 To grdHaem.Cols - 1
3020          If InStr(grdHaem.TextMatrix(0, n), "NV") > 1 Then
3030              grdHaem.Row = 0
3040              grdHaem.Col = n
3050              grdHaem.CellBackColor = vbYellow
3060              grdHaem.CellForeColor = vbBlue
3070          Else
3080              grdHaem.Row = 0
3090              grdHaem.Col = n
3100              grdHaem.CellBackColor = 0
3110              grdHaem.CellForeColor = 0
3120          End If
3130      Next

3140      Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer

3150      intEL = Erl
3160      strES = Err.Description
3170      LogError "frmFullHaem", "FillG", intEL, strES, sql


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
60        LogError "frmFullHaem", "cmdGo_Click", intEL, strES


End Sub



Private Sub cmdRefresh_Click()
10    On Error GoTo cmdRefresh_Click_Error

20    grdHaem.Visible = False
          If cmbResultCount.Text <> "" Then
             FillG (Trim(cmbResultCount.Text))
          End If
          grdHaem.Visible = True
30    FillCombos

40    Exit Sub

cmdRefresh_Click_Error:
Dim strES As String
Dim intEL As Integer

50    intEL = Erl
60    strES = Err.Description
70    LogError "frmFullHaem", "cmdRefresh_Click", intEL, strES

End Sub




Private Sub Form_Activate()

10        On Error GoTo Form_Activate_Error

20        Me.Refresh

30        dtFrom = Format(Now - SysOptWardDate(0), "dd/MMM/yyyy")
40        dtTo = Format(Now, "dd/MMM/yyyy")

50        If TestSys = True Then Me.Caption = Me.Caption & " - TEST SYSTEM"

60        If GetOptionSetting("EnableiPMSChart", "0") = 0 Then
70            chkChartNumber.Value = 0
80            chkChartNumber.Enabled = Not (lblChart = "")
90        Else
100           chkChartNumber.Value = 1

110       End If

          grdHaem.Visible = False
          If cmbResultCount.Text <> "" Then
120         FillG (Trim(cmbResultCount.Text))
          End If
          grdHaem.Visible = True
130       FillCombos



140       PBar.Max = LogOffDelaySecs
150       PBar = 0


160       Timer1.Enabled = True

170       Exit Sub

Form_Activate_Error:

          Dim strES As String
          Dim intEL As Integer



180       intEL = Erl
190       strES = Err.Description
200       LogError "frmFullHaem", "Form_Activate", intEL, strES


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
60        LogError "frmFullHaem", "Form_Deactivate", intEL, strES


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
70        LogError "frmFullHaem", "Form_Load", intEL, strES


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
60        LogError "frmFullHaem", "Form_MouseMove", intEL, strES


End Sub

Private Sub grdHaem_Click()

10        On Error GoTo grdHaem_Click_Error

20        DrawChart

30        Exit Sub

grdHaem_Click_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmFullHaem", "grdHaem_Click", intEL, strES


End Sub


Private Sub grdHaem_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

          Dim Obs As Observations
          Dim SampleID As Long

10        On Error GoTo grdHaem_MouseMove_Error

20        PBar = 0

30        SampleID = Val(grdHaem.TextMatrix(0, grdHaem.MouseCol))
40        If SampleID = 0 Then Exit Sub

50        If grdHaem.MouseRow = 23 Then
60            If Trim(grdHaem.TextMatrix(23, grdHaem.MouseCol)) = "Yes" Then
70                Set Obs = New Observations
80                Set Obs = Obs.Load(SampleID, "Haematology")
90                If Not Obs Is Nothing Then
100                   grdHaem.ToolTipText = Obs.Item(1).Comment
110               Else
120                   grdHaem.ToolTipText = ""

130               End If
140           Else
150               grdHaem.ToolTipText = ""
160           End If
170       Else
180           grdHaem.ToolTipText = ""
190       End If

200       Exit Sub

grdHaem_MouseMove_Error:

          Dim strES As String
          Dim intEL As Integer

210       intEL = Erl
220       strES = Err.Description
230       LogError "frmFullHaem", "grdHaem_MouseMove", intEL, strES

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
190       LogError "frmFullHaem", "pb_MouseMove", intEL, strES


End Sub

Private Sub FillCombos()

          Dim x As Long

10        On Error GoTo FillCombos_Error

20        cmbPlotFrom.Clear
30        cmbPlotTo.Clear

40        If grdHaem.Cols < 2 Or grdHaem.TextMatrix(0, 1) = "" Then Exit Sub

50        For x = 1 To grdHaem.Cols - 1
60            cmbPlotFrom.AddItem Format$(grdHaem.TextMatrix(1, x), "dd/mmm/yyyy")
70            cmbPlotTo.AddItem Format$(grdHaem.TextMatrix(1, x), "dd/mmm/yyyy")
80        Next

90        cmbPlotTo = Format$(grdHaem.TextMatrix(1, 1), "dd/mmm/yyyy")

100       For x = grdHaem.Cols - 1 To 1 Step -1
110           If DateDiff("d", Format$(grdHaem.TextMatrix(1, x), "dd/mmm/yyyy"), cmbPlotTo) < 365 Then
120               cmbPlotFrom = Format$(grdHaem.TextMatrix(1, x), "dd/mmm/yyyy")
130               Exit For
140           End If
150       Next

160       Exit Sub

FillCombos_Error:

          Dim strES As String
          Dim intEL As Integer



170       intEL = Erl
180       strES = Err.Description
190       LogError "frmFullHaem", "FillCombos", intEL, strES


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
90        LogError "frmFullHaem", "Timer1_Timer", intEL, strES


End Sub

Private Sub ColouriseG(ByVal Analyte As String, _
                       ByVal Destination As MSFlexGrid, _
                       ByVal x As Long, _
                       ByVal Y As Long, _
                       ByVal strValue As String, _
                       ByVal sex As String, _
                       ByVal Dob As String, _
                       Optional FeildName As String)

      Dim Value As Single
      Dim z As Long

10    On Error GoTo ColouriseG_Error

20    strValue = Trim(strValue)

30    Value = Val(strValue)

40    If InStr(strValue, ">") Then
50        z = InStr(strValue, ">")
60        Value = Mid(strValue, z + 1)
70    End If




80    Destination.TextMatrix(x, Y) = strValue

90    Destination.Col = Y
100   Destination.Row = x

110   If Trim$(strValue) = "" Then


120       Destination.CellBackColor = &HFFFFFF
130       Destination.CellForeColor = &H0&
140       Exit Sub
150   End If

160   Select Case InterpH(Value, Analyte, sex, Dob, Tn, Format(Destination.TextMatrix(1, Y), "dd/MMM/yyyy"))
      Case "X":
170       Destination.CellBackColor = SysOptPlasBack(0)
180       Destination.CellForeColor = SysOptPlasFore(0)
190   Case "H":
200       Destination.CellBackColor = SysOptHighBack(0)
210       Destination.CellForeColor = SysOptHighFore(0)
220   Case "L"
230       Destination.CellBackColor = SysOptLowBack(0)
240       Destination.CellForeColor = SysOptLowFore(0)
250   Case Else
260       Destination.CellBackColor = &HFFFFFF
270       Destination.CellForeColor = &H0&
280   End Select
      '--------------------------------------------------
290   If FeildName = "" Then FeildName = Analyte
      'Debug.Print Val(Destination.TextMatrix(0, Y)) & "    " & FeildName
300   If IsHaemResultAmended(Val(Destination.TextMatrix(0, Y)), FeildName) = True Then

310       Destination.CellFontUnderline = True
320   End If
      '=================================================
330   Exit Sub

ColouriseG_Error:

      Dim strES As String
      Dim intEL As Integer



340   intEL = Erl
350   strES = Err.Description
360   LogError "frmFullHaem", "ColouriseG", intEL, strES


End Sub


