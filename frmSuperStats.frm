VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frmSuperStats 
   Caption         =   "NetAcquire - Statistics "
   ClientHeight    =   8415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15075
   Icon            =   "frmSuperStats.frx":0000
   LinkTopic       =   "frmBigStats"
   ScaleHeight     =   8415
   ScaleWidth      =   15075
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CheckBox chkHideZeroCount 
      Caption         =   "Hide tests with zero count"
      Height          =   375
      Left            =   1140
      TabIndex        =   52
      Top             =   2040
      Width           =   2715
   End
   Begin MSFlexGridLib.MSFlexGrid grdStats 
      Height          =   5775
      Left            =   120
      TabIndex        =   22
      Top             =   2460
      Width           =   14910
      _ExtentX        =   26300
      _ExtentY        =   10186
      _Version        =   393216
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
   Begin VB.PictureBox fraProgress 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   4800
      ScaleHeight     =   525
      ScaleWidth      =   3825
      TabIndex        =   44
      Top             =   4740
      Width           =   3855
      Begin MSComctlLib.ProgressBar pbProgress 
         Height          =   255
         Left            =   0
         TabIndex        =   45
         Top             =   240
         Width           =   3840
         _ExtentX        =   6773
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         Min             =   1
      End
      Begin VB.Label lblProgress 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Fetching results ..."
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         TabIndex        =   46
         Top             =   0
         Width           =   3840
      End
   End
   Begin Threed.SSPanel sspanOOH 
      Height          =   1260
      Left            =   12600
      TabIndex        =   32
      Top             =   0
      Width           =   1545
      _Version        =   65536
      _ExtentX        =   2725
      _ExtentY        =   2222
      _StockProps     =   15
      Caption         =   "When"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   0
      Begin VB.OptionButton cRooH 
         Alignment       =   1  'Right Justify
         Caption         =   "All"
         Height          =   195
         Index           =   2
         Left            =   420
         TabIndex        =   35
         Top             =   285
         Value           =   -1  'True
         Width           =   885
      End
      Begin VB.OptionButton cRooH 
         Alignment       =   1  'Right Justify
         Caption         =   "Routine"
         Height          =   195
         Index           =   0
         Left            =   420
         TabIndex        =   34
         Top             =   600
         Width           =   885
      End
      Begin VB.OptionButton cRooH 
         Alignment       =   1  'Right Justify
         Caption         =   "Out of Hours"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   33
         Top             =   915
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdExcel 
      Caption         =   "Export to Excel"
      Height          =   870
      Left            =   1530
      Picture         =   "frmSuperStats.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   405
      Width           =   1230
   End
   Begin Threed.SSCommand cmdStart 
      Height          =   870
      Left            =   90
      TabIndex        =   19
      Top             =   405
      Width           =   1320
      _Version        =   65536
      _ExtentX        =   2328
      _ExtentY        =   1535
      _StockProps     =   78
      Caption         =   "Start"
      Picture         =   "frmSuperStats.frx":0614
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   1260
      Left            =   4260
      TabIndex        =   4
      Top             =   0
      Width           =   3435
      _Version        =   65536
      _ExtentX        =   6059
      _ExtentY        =   2222
      _StockProps     =   15
      Caption         =   "Discipline"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   0
      Begin VB.OptionButton optDisp 
         Caption         =   "Microbiology"
         Height          =   195
         Index           =   6
         Left            =   1845
         TabIndex        =   51
         Tag             =   "Micro"
         Top             =   750
         Width           =   1320
      End
      Begin VB.OptionButton optDisp 
         Caption         =   "Blood Gas"
         Height          =   195
         Index           =   5
         Left            =   1845
         TabIndex        =   18
         Tag             =   "BGA"
         Top             =   240
         Width           =   1275
      End
      Begin VB.OptionButton optDisp 
         Caption         =   "Immunology"
         Height          =   195
         Index           =   4
         Left            =   1845
         TabIndex        =   17
         Tag             =   "Imm"
         Top             =   495
         Width           =   1320
      End
      Begin VB.OptionButton optDisp 
         Caption         =   "Endocrinology"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   16
         Tag             =   "End"
         Top             =   1005
         Width           =   1365
      End
      Begin VB.OptionButton optDisp 
         Caption         =   "Haematology"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Tag             =   "Haem"
         Top             =   240
         Value           =   -1  'True
         Width           =   1275
      End
      Begin VB.OptionButton optDisp 
         Caption         =   "Coagulation"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Tag             =   "Coag"
         Top             =   750
         Width           =   1185
      End
      Begin VB.OptionButton optDisp 
         Caption         =   "Biochemistry"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Tag             =   "Bio"
         Top             =   495
         Width           =   1275
      End
   End
   Begin MSComCtl2.DTPicker dtFrom 
      Height          =   330
      Left            =   450
      TabIndex        =   0
      Top             =   0
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   582
      _Version        =   393216
      Format          =   16515073
      CurrentDate     =   38631
   End
   Begin MSComCtl2.DTPicker dtTo 
      Height          =   330
      Left            =   2475
      TabIndex        =   1
      Top             =   0
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   582
      _Version        =   393216
      Format          =   16515073
      CurrentDate     =   38631
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   1260
      Left            =   7740
      TabIndex        =   5
      Top             =   0
      Width           =   1275
      _Version        =   65536
      _ExtentX        =   2249
      _ExtentY        =   2222
      _StockProps     =   15
      Caption         =   "Group By"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   0
      Begin VB.OptionButton optDoc 
         Caption         =   "Ward"
         Height          =   195
         Index           =   2
         Left            =   270
         TabIndex        =   8
         Top             =   900
         Width           =   690
      End
      Begin VB.OptionButton optDoc 
         Caption         =   "Clinician"
         Height          =   195
         Index           =   1
         Left            =   270
         TabIndex        =   7
         Top             =   585
         Width           =   915
      End
      Begin VB.OptionButton optDoc 
         Caption         =   "Gp"
         Height          =   195
         Index           =   0
         Left            =   270
         TabIndex        =   6
         Top             =   270
         Value           =   -1  'True
         Width           =   690
      End
   End
   Begin Threed.SSPanel SSPanel3 
      Height          =   960
      Left            =   4260
      TabIndex        =   9
      Top             =   1290
      Width           =   9900
      _Version        =   65536
      _ExtentX        =   17462
      _ExtentY        =   1693
      _StockProps     =   15
      Caption         =   "Hospital"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   0
      Begin VB.OptionButton optHosp 
         Height          =   195
         Index           =   14
         Left            =   7560
         TabIndex        =   50
         Top             =   720
         Visible         =   0   'False
         Width           =   2200
      End
      Begin VB.OptionButton optHosp 
         Height          =   195
         Index           =   13
         Left            =   7560
         TabIndex        =   49
         Top             =   495
         Visible         =   0   'False
         Width           =   2200
      End
      Begin VB.OptionButton optHosp 
         Height          =   195
         Index           =   12
         Left            =   7560
         TabIndex        =   48
         Top             =   270
         Visible         =   0   'False
         Width           =   2200
      End
      Begin VB.OptionButton optHosp 
         Height          =   195
         Index           =   11
         Left            =   7560
         TabIndex        =   47
         Top             =   45
         Visible         =   0   'False
         Width           =   2200
      End
      Begin VB.OptionButton optHosp 
         Height          =   195
         Index           =   10
         Left            =   5175
         TabIndex        =   30
         Top             =   720
         Visible         =   0   'False
         Width           =   2200
      End
      Begin VB.OptionButton optHosp 
         Height          =   195
         Index           =   9
         Left            =   5175
         TabIndex        =   29
         Top             =   495
         Visible         =   0   'False
         Width           =   2200
      End
      Begin VB.OptionButton optHosp 
         Height          =   195
         Index           =   8
         Left            =   5175
         TabIndex        =   28
         Top             =   270
         Visible         =   0   'False
         Width           =   2200
      End
      Begin VB.OptionButton optHosp 
         Height          =   195
         Index           =   7
         Left            =   5175
         TabIndex        =   27
         Top             =   45
         Visible         =   0   'False
         Width           =   2200
      End
      Begin VB.OptionButton optHosp 
         Height          =   195
         Index           =   6
         Left            =   2430
         TabIndex        =   26
         Top             =   720
         Visible         =   0   'False
         Width           =   2200
      End
      Begin VB.OptionButton optHosp 
         Height          =   195
         Index           =   5
         Left            =   2430
         TabIndex        =   25
         Top             =   495
         Visible         =   0   'False
         Width           =   2200
      End
      Begin VB.OptionButton optHosp 
         Height          =   195
         Index           =   4
         Left            =   2430
         TabIndex        =   24
         Top             =   270
         Visible         =   0   'False
         Width           =   2200
      End
      Begin VB.OptionButton optHosp 
         Height          =   195
         Index           =   3
         Left            =   2430
         TabIndex        =   23
         Top             =   45
         Visible         =   0   'False
         Width           =   2200
      End
      Begin VB.OptionButton optHosp 
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   12
         Top             =   720
         Visible         =   0   'False
         Width           =   2200
      End
      Begin VB.OptionButton optHosp 
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   11
         Top             =   495
         Visible         =   0   'False
         Width           =   2200
      End
      Begin VB.OptionButton optHosp 
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   10
         Top             =   270
         Value           =   -1  'True
         Width           =   2200
      End
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   870
      Left            =   2880
      TabIndex        =   20
      Top             =   405
      Width           =   1140
      _Version        =   65536
      _ExtentX        =   2011
      _ExtentY        =   1535
      _StockProps     =   78
      Caption         =   "Exit"
      Picture         =   "frmSuperStats.frx":092E
   End
   Begin Threed.SSPanel SSPanel4 
      Height          =   1260
      Left            =   9060
      TabIndex        =   37
      Top             =   0
      Width           =   3495
      _Version        =   65536
      _ExtentX        =   6165
      _ExtentY        =   2222
      _StockProps     =   15
      Caption         =   "Filter By"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   0
      Begin VB.ComboBox cmbList 
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   1200
         TabIndex        =   43
         Top             =   240
         Width           =   1935
      End
      Begin VB.ComboBox cmbList 
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   1200
         TabIndex        =   42
         Top             =   555
         Width           =   1935
      End
      Begin VB.ComboBox cmbList 
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   1200
         TabIndex        =   41
         Top             =   870
         Width           =   1935
      End
      Begin VB.CheckBox chkCriteria 
         Caption         =   "Ward"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   40
         Top             =   930
         Width           =   975
      End
      Begin VB.CheckBox chkCriteria 
         Caption         =   "Clinician"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   39
         Top             =   615
         Width           =   975
      End
      Begin VB.CheckBox chkCriteria 
         Caption         =   "GP"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   38
         Top             =   300
         Width           =   975
      End
   End
   Begin VB.Label lblExcelInfo 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exporting..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1500
      TabIndex        =   36
      Top             =   1350
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label3 
      Caption         =   "The Report is being Generated.              Please Wait."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      Left            =   3375
      TabIndex        =   21
      Top             =   3480
      Width           =   6765
   End
   Begin VB.Label Label2 
      Caption         =   "To"
      Height          =   240
      Left            =   2205
      TabIndex        =   3
      Top             =   45
      Width           =   825
   End
   Begin VB.Label Label1 
      Caption         =   "From"
      Height          =   285
      Left            =   45
      TabIndex        =   2
      Top             =   45
      Width           =   1140
   End
End
Attribute VB_Name = "frmSuperStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Dept As String
Dim Choice As String
Dim Hosp As String

Private Sub chkCriteria_Click(Index As Integer)
10        cmbList(Index).Enabled = chkCriteria(Index).Value
20        If cmbList(Index).Enabled = False Then
30            cmbList(Index) = ""
40        End If
End Sub

Private Sub chkHideZeroCount_Click()

10        On Error GoTo chkHideZeroCount_Click_Error

20        If chkHideZeroCount.Value = 1 Then
30            HideTestsWithZeroCount True
40        Else
50            HideTestsWithZeroCount False
60        End If
70        Exit Sub

chkHideZeroCount_Click_Error:

          Dim strES As String
          Dim intEL As Integer

80        intEL = Erl
90        strES = Err.Description
100       LogError "frmSuperStats", "chkHideZeroCount_Click", intEL, strES

End Sub

Private Sub cmdExcel_Click()


10        On Error GoTo cmdExcel_Click_Error

20        ExportFlexGrid grdStats, Me

30        Exit Sub

cmdExcel_Click_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmSuperStats", "cmdExcel_Click", intEL, strES


End Sub

Private Sub cmdExit_Click()

10        Unload Me

End Sub

Private Sub FillHosps()
          Dim n As Long
          Dim sql As String
          Dim tb As Recordset

10        On Error GoTo FillHosps_Error

20        n = 0

          'sql = "SELECT hospital , Count(hospital) AS NumberOfDups From demographics " & _
           '      "GROUP BY hospital " & _
           '      "Having (Count(Hospital) > 1) order by numberofdups desc"

30        sql = "SELECT DISTINCT Text Hospital FROM Lists WHERE ListType = 'HO' AND COALESCE(Text, '') <> '' " & _
                "ORDER BY Text"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        Do While Not tb.EOF
70            If Trim(tb!Hospital & "") <> "" Then
80                optHosp(n).Visible = True
90                optHosp(n).Caption = Left(initial2upper(Trim(tb!Hospital)), 25)
100               n = n + 1
110               If n = 15 Then Exit Do
120           End If
130           tb.MoveNext
140       Loop

150       Exit Sub

FillHosps_Error:

          Dim strES As String
          Dim intEL As Integer



160       intEL = Erl
170       strES = Err.Description
180       LogError "frmSuperStats", "FillHosps", intEL, strES, sql


End Sub
Private Sub cmdStart_Click()

          Dim tb As Recordset
          Dim sql As String
          Dim n As Long
          Dim z As Long
          Dim TestTot As Long
          Dim s As String
          Dim Sel As String
          Dim Y As Integer
          Dim X As Integer
          Dim xFilter As String
          Dim FromDate As String
          Dim ToDate As String
          Dim SourceTable As String
          Dim MonthIndex As Integer
          Dim StartDate As Date
          Dim EndDate As Date
          Dim SrcUpdated As Boolean
          Dim DiffMonths As Integer

10        On Error GoTo cmdStart_Click_Error

20        If chkCriteria(0).Value = True And cmbList(0) = "" Then
30            iMsg "Please select GP first"
40            Exit Sub
50        End If
60        If chkCriteria(1) = True And cmbList(1) = "" Then
70            iMsg "Please select Clinician first"
80            Exit Sub
90        End If
100       If chkCriteria(2) = True And cmbList(2) = "" Then
110           iMsg "Please select Ward first"
120           Exit Sub
130       End If

140       pbProgress.Max = 2
150       StartDate = dtFrom
160       pbProgress.Value = 1
170       lblProgress = "Fetching results ... (0%)"
180       lblProgress.Refresh
190       DiffMonths = DateDiff("m", dtFrom, dtTo)
200       If DiffMonths = 0 Then
210           EndDate = dtTo
220           pbProgress.Max = 2
230           DiffMonths = DiffMonths + 1
240       Else
250           EndDate = DateAdd("m", 1, dtFrom)
260           EndDate = "01/" & Month(EndDate) & "/" & Year(EndDate)
270           EndDate = DateAdd("d", -1, EndDate)
280           DiffMonths = DiffMonths + 1
290           pbProgress.Max = DiffMonths + 1
300       End If

310       grdStats.Visible = False
320       Me.Refresh

330       FillRows
340       FillCols
350       xFilter = ""
360       If chkCriteria(0).Value = 1 Then
370           xFilter = "And Gp = '" & cmbList(0) & "' "
380       End If
390       If chkCriteria(1).Value = 1 Then
400           xFilter = xFilter & "And Clinician = '" & cmbList(1) & "' "
410       End If
420       If chkCriteria(2).Value = 1 Then
430           xFilter = xFilter & "And Ward = '" & cmbList(2) & "' "
440       End If

450       grdStats.TextMatrix(0, 0) = Choice & " Name"

460       Sel = "xxx"



470       If Dept = "Haem" Then
480           For MonthIndex = 1 To DiffMonths
490               pbProgress.Value = pbProgress.Value + 1
500               lblProgress = "Fetching results ... (" & Int(pbProgress.Value * 100 / pbProgress.Max) & " %)"
510               lblProgress.Refresh

520               sql = "IF EXISTS (SELECT * FROM dbo.sysobjects WHERE name = 'CustomTable') " & _
                        "  DROP TABLE [CustomTable] " & _
                        "SELECT DISTINCT D." & Choice & " Choice, D.SampleID, " & _
                        "CASE WBC WHEN '' THEN 0 ELSE COUNT (WBC) END WBC, " & _
                        "CASE RetA WHEN '' THEN 0 ELSE COUNT (RetA) END RetA, " & _
                        "CASE ESR WHEN '' THEN 0 ELSE COUNT (ESR) END ESR, " & _
                        "CASE tRA WHEN '' THEN 0 ELSE COUNT (tRA) END tRA, " & _
                        "CASE Malaria WHEN '' THEN 0 ELSE COUNT (Malaria) END Malaria, " & _
                        "CASE Monospot WHEN '' THEN 0 ELSE COUNT (Monospot) END Monospot, " & _
                        "CASE Sickledex WHEN '' THEN 0 ELSE COUNT (Sickledex) END Sickledex, " & _
                        "CASE tASOT WHEN '' THEN 0 ELSE COUNT (tASOT) END tASOT " & _
                        "INTO CustomTable " & _
                        "FROM HaemResults R, Demographics D WHERE " & _
                        "D.SampleID = R.SampleID " & _
                        "AND D.Rundate BETWEEN '" & Format(StartDate, "dd/MMM/yyyy 00:00:00") & "' AND '" & _
                        Format(EndDate, "dd/MMM/yyyy 23:59:59") & "' " & _
                        "AND (D.Hospital = '" & Hosp & "') " & xFilter
530               If cRooH(1) Then
540                   sql = sql & " AND D.RooH = 0 "
550               ElseIf cRooH(0) Then
560                   sql = sql & " AND D.RooH = 1 "
570               End If
580               sql = sql & "GROUP BY D." & Choice & ", WBC, RetA, ESR, tRA, Malaria, Monospot, Sickledex, tASOT, D.SampleID " & _
                        "ORDER BY " & Choice & ", D.SampleID "
590               Cnxn(0).Execute sql
600               sql = "SELECT Choice, " & _
                        "SUM(WBC) WBC, " & _
                        "SUM(RetA) RetA, " & _
                        "SUM(ESR) ESR, " & _
                        "sum(tRA) tRA, " & _
                        "SUM(Malaria) Malaria, " & _
                        "SUM(Monospot) Monospot, " & _
                        "SUM(Sickledex) Sickledex, " & _
                        "SUM(tASOT) tASOT " & _
                        "FROM CustomTable " & _
                        "GROUP BY Choice " & _
                        "ORDER By Choice "

610               Set tb = New Recordset
620               RecOpenServer 0, tb, sql
630               Do While Not tb.EOF
640                   SrcUpdated = False
650                   For Y = 1 To grdStats.Rows - 1
660                       If UCase(Trim$(grdStats.TextMatrix(Y, 0))) = UCase(Trim$(tb!Choice & "")) Then
670                           grdStats.TextMatrix(Y, 1) = Val(grdStats.TextMatrix(Y, 1)) + tb!wbc
680                           grdStats.TextMatrix(Y, 2) = Val(grdStats.TextMatrix(Y, 2)) + tb!reta
690                           grdStats.TextMatrix(Y, 3) = Val(grdStats.TextMatrix(Y, 3)) + tb!esr
700                           grdStats.TextMatrix(Y, 4) = Val(grdStats.TextMatrix(Y, 4)) + tb!tRa
710                           grdStats.TextMatrix(Y, 5) = Val(grdStats.TextMatrix(Y, 5)) + tb!Malaria
720                           grdStats.TextMatrix(Y, 6) = Val(grdStats.TextMatrix(Y, 6)) + tb!Monospot
730                           grdStats.TextMatrix(Y, 7) = Val(grdStats.TextMatrix(Y, 7)) + tb!Sickledex
740                           grdStats.TextMatrix(Y, 8) = Val(grdStats.TextMatrix(Y, 8)) + tb!tASOt
750                           SrcUpdated = True
760                           Exit For


770                       End If
780                   Next
790                   If Not SrcUpdated Then
800                       grdStats.AddItem tb!Choice & "" & vbTab & _
                                           tb!wbc & "" & vbTab & _
                                           tb!reta & "" & vbTab & _
                                           tb!esr & "" & vbTab & _
                                           tb!tRa & "" & vbTab & _
                                           tb!Malaria & "" & vbTab & _
                                           tb!Monospot & "" & vbTab & _
                                           tb!Sickledex & "" & vbTab & _
                                           tb!tASOt
810                   End If

                      '            If UCase(Sel) <> UCase(Trim$(tb!Choice & "")) Then
                      '                Sel = UCase(Trim$(tb!Choice & ""))
                      '                For y = 1 To grdStats.Rows - 1
                      '                    If UCase(Trim$(grdStats.TextMatrix(y, 0))) = UCase(Trim$(tb!Choice & "")) Then
                      '                        Exit For
                      '                    End If
                      '                Next
                      '            End If
                      '            If y <> grdStats.Rows And y <> 0 Then
                      '                grdStats.TextMatrix(y, 1) = Val(grdStats.TextMatrix(y, 1)) + tb!wbc
                      '                grdStats.TextMatrix(y, 2) = Val(grdStats.TextMatrix(y, 2)) + tb!reta
                      '                grdStats.TextMatrix(y, 3) = Val(grdStats.TextMatrix(y, 3)) + tb!esr
                      '                grdStats.TextMatrix(y, 4) = Val(grdStats.TextMatrix(y, 4)) + tb!tRa
                      '                grdStats.TextMatrix(y, 5) = Val(grdStats.TextMatrix(y, 5)) + tb!Malaria
                      '                grdStats.TextMatrix(y, 6) = Val(grdStats.TextMatrix(y, 6)) + tb!Monospot
                      '                grdStats.TextMatrix(y, 7) = Val(grdStats.TextMatrix(y, 7)) + tb!Sickledex
                      '                grdStats.TextMatrix(y, 8) = Val(grdStats.TextMatrix(y, 8)) + tb!tASOt
                      '            End If
820                   tb.MoveNext
830               Loop


840               StartDate = DateAdd("d", 1, EndDate)
850               EndDate = DateAdd("m", 1, StartDate)
860               EndDate = "01/" & Month(EndDate) & "/" & Year(EndDate)
870               EndDate = DateAdd("d", -1, EndDate)
880               If MonthIndex = DiffMonths And DateDiff("d", dtTo, EndDate) > 0 Then
890                   EndDate = dtTo
900               End If
910           Next MonthIndex


920       Else
930           For MonthIndex = 1 To DiffMonths
940               pbProgress.Value = pbProgress.Value + 1
950               lblProgress = "Fetching results ... (" & Int(pbProgress.Value * 100 / pbProgress.Max) & " %)"
960               lblProgress.Refresh
970               If Dept = "Micro" Then
980                   sql = "SELECT DISTINCT LTRIM(RTRIM(R.Site)) AS Code, D." & Choice & " Choice, COUNT(R.Site) X " & _
                            "FROM " & Dept & "SiteDetails R, Demographics D WHERE " & _
                            "D.SampleID = R.SampleID " & _
                            "AND D.Hospital = '" & Hosp & "' " & _
                            "AND D.Rundate BETWEEN '" & Format(StartDate, "dd/MMM/yyyy 00:00:00") & "' AND '" & _
                            Format(EndDate, "dd/MMM/yyyy 23:50:59") & "' " & _
                            "AND D.SampleID > " & SysOptMicroOffset(0) & " " & xFilter
990                   If cRooH(1) Then
1000                      sql = sql & "AND D.Rooh = 0 "
1010                  ElseIf cRooH(0) Then
1020                      sql = sql & "AND D.Rooh <> 0 "
1030                  End If

1040                  sql = sql & "GROUP BY D." & Choice & " ,R.Site " & _
                            "ORDER BY " & Choice
1050              Else
1060                  sql = "SELECT DISTINCT LTRIM(RTRIM(R.Code)) AS Code, D." & Choice & " Choice, COUNT(R.Code) X " & _
                            "FROM " & Dept & "Results R, Demographics D WHERE " & _
                            "D.SampleID = R.SampleID " & _
                            "AND D.Hospital = '" & Hosp & "' " & _
                            "AND D.Rundate BETWEEN '" & Format(StartDate, "dd/MMM/yyyy 00:00:00") & "' AND '" & _
                            Format(EndDate, "dd/MMM/yyyy 23:50:59") & "' " & xFilter
1070                  If cRooH(1) Then
1080                      sql = sql & "AND D.Rooh = 0 "
1090                  ElseIf cRooH(0) Then
1100                      sql = sql & "AND D.Rooh <> 0 "
1110                  End If

1120                  sql = sql & "GROUP BY D." & Choice & " ,R.Code " & _
                            "ORDER BY " & Choice
1130              End If


1140              Set tb = New Recordset
1150              RecOpenServer 0, tb, sql
1160              Do While Not tb.EOF
1170                  SrcUpdated = False
1180                  For Y = 1 To grdStats.Rows - 1
1190                      If UCase(Trim$(grdStats.TextMatrix(Y, 0))) = UCase(Trim$(tb!Choice & "")) Then
1200                          For X = 1 To grdStats.Cols - 1
1210                              If tb!Code = Trim$(grdStats.TextMatrix(0, X)) Then
1220                                  grdStats.TextMatrix(Y, X) = Val(grdStats.TextMatrix(Y, X)) + tb!X
1230                                  SrcUpdated = True
1240                                  Exit For
1250                              End If
1260                          Next
1270                      End If
1280                  Next
1290                  If Not SrcUpdated Then

1300                      For X = 1 To grdStats.Cols - 1
1310                          If tb!Code = Trim$(grdStats.TextMatrix(0, X)) Then
1320                              grdStats.AddItem tb!Choice & ""
1330                              grdStats.Row = grdStats.Rows - 1
1340                              grdStats.TextMatrix(Y, X) = tb!X
1350                              Exit For
1360                          End If
1370                      Next

1380                  End If
                      '            If UCase(Sel) <> UCase(Trim$(tb!Choice & "")) Then
                      '                Sel = UCase(Trim$(tb!Choice & ""))
                      '                For y = 1 To grdStats.Rows - 1
                      '                    If UCase(Trim$(grdStats.TextMatrix(y, 0))) = UCase(Trim$(tb!Choice & "")) Then
                      '                        Exit For
                      '                    End If
                      '                Next
                      '            End If
                      '            If y <> grdStats.Rows And y <> 0 Then
                      '                For X = 1 To grdStats.Cols - 1
                      '                    If tb!Code = Trim$(grdStats.TextMatrix(0, X)) Then
                      '                        grdStats.TextMatrix(y, X) = Val(grdStats.TextMatrix(y, X)) + tb!X
                      '                        Exit For
                      '                    End If
                      '                Next
                      '            End If
1390                  tb.MoveNext
1400              Loop
1410              StartDate = DateAdd("d", 1, EndDate)
1420              EndDate = DateAdd("m", 1, StartDate)
1430              EndDate = "01/" & Month(EndDate) & "/" & Year(EndDate)
1440              EndDate = DateAdd("d", -1, EndDate)
1450              If MonthIndex = DiffMonths And DateDiff("d", dtTo, EndDate) > 0 Then
1460                  EndDate = dtTo
1470              End If
1480          Next MonthIndex

1490      End If

1500      grdStats.Cols = grdStats.Cols + 1
1510      grdStats.TextMatrix(0, grdStats.Cols - 1) = "Totals"
1520      For n = 1 To grdStats.Rows - 1
1530          For z = 1 To grdStats.Cols - 2
1540              TestTot = TestTot + Val(grdStats.TextMatrix(n, z))
1550          Next
1560          grdStats.TextMatrix(n, grdStats.Cols - 1) = TestTot
1570          TestTot = 0
1580      Next

1590      s = "TOTAL" & vbTab

1600      For n = 1 To grdStats.Cols - 1
1610          For z = 1 To grdStats.Rows - 1
1620              TestTot = TestTot + Val(grdStats.TextMatrix(z, n))
1630          Next
1640          s = s & TestTot & vbTab
              '        If TestTot = 0 Then
              '            grdStats.ColWidth(n) = IIf(Dept = "Micro", 1500, 1000)
              '        Else
              '
              '            grdStats.ColWidth(n) = IIf(Dept = "Micro", 1500, 1000)
              '
              '        End If
1650          TestTot = 0
1660      Next

1670      TestTot = 0

1680      If Dept <> "Haem" And Dept <> "Micro" Then
1690          For n = 1 To grdStats.Cols - 1
1700              grdStats.TextMatrix(0, n) = GetSCode(grdStats.TextMatrix(0, n), Dept)
1710          Next
1720      End If

1730      grdStats.AddItem ""
1740      grdStats.AddItem s

1750      grdStats.TextMatrix(1, 0) = "Not Assigned"
1760      grdStats.Visible = True

1770      If chkHideZeroCount.Value = 1 Then
1780          HideTestsWithZeroCount True
1790      Else
1800          HideTestsWithZeroCount False
1810      End If

1820      Exit Sub

cmdStart_Click_Error:

          Dim strES As String
          Dim intEL As Integer



1830      intEL = Erl
1840      strES = Err.Description
1850      LogError "frmSuperStats", "cmdStart_Click", intEL, strES, sql

End Sub
Private Function GetSCode(ByVal Code As String, ByVal Dept As String) As String
          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo GetSCode_Error

20        If Dept = "Coag" Then
30            sql = "Select * from " & Dept & "testdefinitions where code = '" & Code & "' and inuse = 1"
40            Set tb = New Recordset
50            RecOpenServer 0, tb, sql
60            If Not tb.EOF Then
70                GetSCode = tb!TestName
80            End If

90        Else
100           sql = "Select * from " & Dept & "testdefinitions where code = '" & Code & "' and inuse = 1"
110           Set tb = New Recordset
120           RecOpenServer 0, tb, sql
130           If Not tb.EOF Then
140               GetSCode = tb!ShortName
150           End If
160       End If

170       Exit Function

GetSCode_Error:

          Dim strES As String
          Dim intEL As Integer



180       intEL = Erl
190       strES = Err.Description
200       LogError "frmSuperStats", "GetSCode", intEL, strES, sql

End Function






Private Sub Form_Load()

10        On Error GoTo Form_Load_Error

20        FillHosps
30        LoadCombos

40        dtFrom = Format(Now, "dd/MMM/yyyy")
50        dtTo = Format(Now, "dd/MMM/yyyy")




          'For n = 0 To intOtherHospitalsInGroup
          '  optHosp(n).Visible = True
          '  optHosp(n).Caption = initial2upper(Hospname(n))
          'Next




60        Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer



70        intEL = Erl
80        strES = Err.Description
90        LogError "frmSuperStats", "Form_Load", intEL, strES


End Sub


Private Sub FillRows()
          Dim tb As Recordset
          Dim sql As String
          Dim n As Long

10        On Error GoTo FillRows_Error

20        grdStats.ColWidth(0) = 2500

30        With grdStats
40            .Rows = 2
50            .Cols = 2
60            .TextMatrix(0, 1) = ""
70            .AddItem ""
80            .RemoveItem 1
90            .AddItem ""
100       End With

110       Hosp = ""

120       For n = 0 To 14
130           If optHosp(n) Then
140               Hosp = optHosp(n).Caption
150           End If
160       Next

170       For n = 0 To 2
180           If optDoc(n) Then
190               Choice = optDoc(n).Caption
200           End If
210       Next


220       For n = 0 To 6
230           If optDisp(n) Then
240               Dept = optDisp(n).Tag
250           End If
260       Next

          'sql = "select distinct(" & Choice & ")  from " & _
           '      "demographics " & _
           '      "WHERE rundate between '" & _
           '      Format(dtFrom, "dd/MMM/yyyy") & "' and '" & _
           '      Format(dtTo, "dd/MMM/yyyy") & "' and " & Choice & " <> '' " & _
           '      "AND Hospital = '" & Hosp & "' "
          'If optDoc(0).Value = True And chkCriteria(0).Value = 1 Then
          '    sql = sql & " AND " & Choice & " = '" & cmbList(0) & "' "
          'ElseIf optDoc(1).Value = True And chkCriteria(1).Value = 1 Then
          '    sql = sql & " AND " & Choice & " = '" & cmbList(1) & "' "
          'ElseIf optDoc(2).Value = True And chkCriteria(2).Value = 1 Then
          '    sql = sql & " AND " & Choice & " = '" & cmbList(2) & "' "
          'End If
          'sql = sql & "order by " & Choice & ""
          'Set tb = New Recordset
          'RecOpenServer 0, tb, sql
          'Do While Not tb.EOF
          '    If Trim(tb(Choice) & "") <> "" Then grdStats.AddItem tb(Choice)
          '    tb.MoveNext
          'Loop

270       If grdStats.Rows > 2 And grdStats.TextMatrix(1, 0) = "" Then
280           grdStats.RemoveItem 1
290       End If


300       Exit Sub

FillRows_Error:

          Dim strES As String
          Dim intEL As Integer



310       intEL = Erl
320       strES = Err.Description
330       LogError "frmSuperStats", "FillRows", intEL, strES, sql


End Sub


Private Sub FillCols()
          Dim tb As Recordset
          Dim sql As String
          Dim n As Long


10        On Error GoTo FillCols_Error

20        If Dept <> "Haem" And Dept <> "Coag" And Dept <> "Micro" Then
30            sql = "SELECT DISTINCT Code, PrintPriority FROM " & Dept & "TestDefinitions WHERE " & _
                    "InUse = 1 " & _
                    "ORDER BY PrintPriority"
              '  sql = "SELECT DISTINCT " & Dept & "TestDefinitions.Code " & _
                 '        "FROM Demographics INNER JOIN " & _
                 '        "" & Dept & "Results ON Demographics.SampleID = " & Dept & "Results.SampleId INNER JOIN " & _
                 '        "" & Dept & "TestDefinitions ON " & Dept & "Results.Code = " & Dept & "TestDefinitions.Code " & _
                 '        "WHERE (" & Dept & "TestDefinitions.InUse = 1) AND " & _
                 '        "(Demographics.RunDate BETWEEN '" & Format(dtFrom, "dd/MMM/yyyy") & "' and '" & _
                 '        Format(dtTo, "dd/MMM/yyyy") & "') and Demographics." & Choice & " <> '' " & _
                 '        "ORDER BY " & Dept & "TestDefinitions.Code"
40        ElseIf Dept = "Coag" Then
50            sql = "SELECT DISTINCT " & Dept & "TestDefinitions.Code, " & Dept & "TestDefinitions.Testname " & _
                    "FROM Demographics INNER JOIN " & _
                    "" & Dept & "Results ON Demographics.SampleID = " & Dept & "Results.SampleId INNER JOIN " & _
                    "" & Dept & "TestDefinitions ON " & Dept & "Results.Code = " & Dept & "TestDefinitions.Code " & _
                    "WHERE (" & Dept & "TestDefinitions.InUse = 1) AND " & _
                    "(Demographics.RunDate BETWEEN '" & Format(dtFrom, "dd/MMM/yyyy") & "' and '" & _
                    Format(dtTo, "dd/MMM/yyyy") & "') and Demographics." & Choice & " <> '' " & _
                    "ORDER BY " & Dept & "TestDefinitions.code"
60        ElseIf Dept = "Micro" Then
70            sql = "SELECT DISTINCT Text Code FROM Lists WHERE ListType = 'SI' ORDER BY Text"
80        Else
90            grdStats.Cols = 9
100           grdStats.TextMatrix(0, 1) = "FBC"
110           grdStats.TextMatrix(0, 2) = "Retics"
120           grdStats.TextMatrix(0, 3) = "Esr"
130           grdStats.TextMatrix(0, 4) = "RF"
140           grdStats.TextMatrix(0, 5) = "Malaria"
150           grdStats.TextMatrix(0, 6) = "Monospot/IM"
160           grdStats.TextMatrix(0, 7) = "Sickledex"
170           grdStats.TextMatrix(0, 8) = "Asot"
180           Exit Sub
190       End If

200       Set tb = New Recordset
210       RecOpenServer 0, tb, sql
220       Do While Not tb.EOF
230           n = grdStats.Cols - 1
240           grdStats.TextMatrix(0, n) = tb!Code & ""
250           grdStats.Cols = grdStats.Cols + 1
260           tb.MoveNext
270       Loop
280       If grdStats.Cols > 2 Then grdStats.Cols = grdStats.Cols - 1

290       Exit Sub

FillCols_Error:

          Dim strES As String
          Dim intEL As Integer



300       intEL = Erl
310       strES = Err.Description
320       LogError "frmSuperStats", "FillCols", intEL, strES, sql


End Sub

Private Sub grdStats_Click()

          Static SortOrder As Boolean

10        On Error GoTo grdStats_Click_Error

20        If grdStats.MouseRow = 0 Then
30            If SortOrder Then
40                grdStats.Sort = flexSortGenericAscending
50            Else
60                grdStats.Sort = flexSortGenericDescending
70            End If
80            SortOrder = Not SortOrder
90        End If

100       Exit Sub

grdStats_Click_Error:

          Dim strES As String
          Dim intEL As Integer



110       intEL = Erl
120       strES = Err.Description
130       LogError "frmSuperStats", "grdStats_Click", intEL, strES


End Sub



Private Sub LoadCombos()

          Dim sql As String
          Dim tb As Recordset

10        On Error GoTo LoadCombos_Error

20        sql = "Select Text From Gps Order By Text"
30        Set tb = New Recordset
40        RecOpenClient 0, tb, sql
50        cmbList(0).Clear
60        If Not tb.EOF Then
70            While Not tb.EOF
80                cmbList(0).AddItem tb!Text
90                tb.MoveNext
100           Wend
110       End If
120       FixComboWidth cmbList(0)

130       sql = "Select Text From Clinicians Order By Text"
140       Set tb = New Recordset
150       RecOpenClient 0, tb, sql
160       cmbList(1).Clear
170       If Not tb.EOF Then
180           While Not tb.EOF
190               cmbList(1).AddItem tb!Text
200               tb.MoveNext
210           Wend
220       End If
230       FixComboWidth cmbList(1)

240       sql = "Select Text From Wards Order By Text"
250       Set tb = New Recordset
260       RecOpenClient 0, tb, sql
270       cmbList(2).Clear
280       If Not tb.EOF Then
290           While Not tb.EOF
300               cmbList(2).AddItem tb!Text
310               tb.MoveNext
320           Wend
330       End If
340       FixComboWidth cmbList(2)

350       Exit Sub

LoadCombos_Error:

          Dim strES As String
          Dim intEL As Integer

360       intEL = Erl
370       strES = Err.Description
380       LogError "frmSuperStats", "LoadCombos", intEL, strES, sql

End Sub

Private Sub HideTestsWithZeroCount(HideFlag As Boolean)

10        On Error GoTo HideTestsWithZeroCount_Error

          Dim i As Integer
20        With grdStats
30            For i = 1 To .Cols - 1
40                If .TextMatrix(.Rows - 1, i) = 0 Then
50                    If HideFlag Then
60                        .ColWidth(i) = 0
70                    Else
80                        If optDisp(6) Then
90                            .ColWidth(i) = 1500
100                       Else
110                           .ColWidth(i) = 1000
120                       End If
130                   End If
140               Else
150                   If optDisp(6) Then
160                       .ColWidth(i) = 1500
170                   Else
180                       .ColWidth(i) = 1000
190                   End If
200               End If
210           Next i
220       End With
230       Exit Sub

HideTestsWithZeroCount_Error:

          Dim strES As String
          Dim intEL As Integer

240       intEL = Erl
250       strES = Err.Description
260       LogError "frmSuperStats", "HideTestsWithZeroCount", intEL, strES

End Sub
