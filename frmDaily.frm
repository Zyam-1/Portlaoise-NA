VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDaily 
   Appearance      =   0  'Flat
   Caption         =   "NetAcquire - Daily Report"
   ClientHeight    =   7875
   ClientLeft      =   810
   ClientTop       =   1395
   ClientWidth     =   13140
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
   Icon            =   "frmDaily.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7875
   ScaleWidth      =   13140
   Begin VB.CommandButton cmdRefresh 
      Appearance      =   0  'Flat
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
      Height          =   975
      Left            =   9210
      Picture         =   "frmDaily.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   270
      Width           =   1245
   End
   Begin VB.Frame Frame1 
      Caption         =   "Discipline"
      Height          =   1380
      Left            =   90
      TabIndex        =   3
      Top             =   60
      Width           =   9090
      Begin VB.OptionButton optSemen 
         Caption         =   "Semen Analysis"
         Height          =   255
         Left            =   7350
         TabIndex        =   16
         Top             =   765
         Width           =   1665
      End
      Begin VB.OptionButton optMicro 
         Caption         =   "Microbiology"
         Enabled         =   0   'False
         Height          =   255
         Left            =   7350
         TabIndex        =   15
         Top             =   495
         Width           =   1575
      End
      Begin VB.OptionButton optCyto 
         Caption         =   "Cytology"
         Height          =   255
         Left            =   7350
         TabIndex        =   14
         Top             =   225
         Width           =   1575
      End
      Begin VB.OptionButton optHisto 
         Caption         =   "Histology"
         Height          =   255
         Left            =   5730
         TabIndex        =   13
         Top             =   1035
         Width           =   1575
      End
      Begin VB.OptionButton optBG 
         Caption         =   "Blood Gas"
         Height          =   255
         Left            =   5730
         TabIndex        =   11
         Top             =   765
         Width           =   1575
      End
      Begin VB.OptionButton optImm 
         Caption         =   "Immunology"
         Height          =   255
         Left            =   5730
         TabIndex        =   10
         Top             =   225
         Width           =   1575
      End
      Begin VB.OptionButton optEnd 
         Caption         =   "Endocrinology"
         Height          =   255
         Left            =   3870
         TabIndex        =   9
         Top             =   1035
         Width           =   1575
      End
      Begin VB.OptionButton optExt 
         Caption         =   "External"
         Height          =   255
         Left            =   5730
         TabIndex        =   7
         Top             =   495
         Width           =   1485
      End
      Begin VB.OptionButton optCoag 
         Caption         =   "Coagulation"
         Height          =   255
         Left            =   3870
         TabIndex        =   6
         Top             =   765
         Width           =   1485
      End
      Begin VB.OptionButton optHaem 
         Caption         =   "Haematology"
         Height          =   255
         Left            =   3870
         TabIndex        =   5
         Top             =   495
         Width           =   1485
      End
      Begin VB.OptionButton optBio 
         Caption         =   "Biochemistry"
         Height          =   255
         Left            =   3870
         TabIndex        =   4
         Top             =   225
         Width           =   1485
      End
      Begin MSComCtl2.DTPicker dt 
         Height          =   315
         Left            =   945
         TabIndex        =   12
         Top             =   315
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   556
         _Version        =   393216
         Format          =   161808385
         CurrentDate     =   37096
      End
      Begin MSMask.MaskEdBox tFromTime 
         Height          =   315
         Left            =   945
         TabIndex        =   18
         ToolTipText     =   "Time of Sample"
         Top             =   945
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox tToTime 
         Height          =   315
         Left            =   2355
         TabIndex        =   19
         ToolTipText     =   "Time of Sample"
         Top             =   945
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label 
         Caption         =   "Date"
         Height          =   195
         Left            =   315
         TabIndex        =   23
         Top             =   360
         Width           =   510
      End
      Begin VB.Label Label1 
         Caption         =   "Time"
         Height          =   195
         Left            =   315
         TabIndex        =   22
         Top             =   1005
         Width           =   510
      End
      Begin VB.Label lblFrom 
         Caption         =   "From"
         Height          =   195
         Left            =   1170
         TabIndex        =   21
         Top             =   720
         Width           =   510
      End
      Begin VB.Label Label2 
         Caption         =   "To"
         Height          =   195
         Left            =   2670
         TabIndex        =   20
         Top             =   720
         Width           =   330
      End
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   5955
      Left            =   90
      TabIndex        =   2
      Top             =   1545
      Width           =   13005
      _ExtentX        =   22939
      _ExtentY        =   10504
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      GridLines       =   3
      GridLinesFixed  =   3
      AllowUserResizing=   1
      FormatString    =   $"frmDaily.frx":0614
   End
   Begin VB.CommandButton bprint 
      Appearance      =   0  'Flat
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   10530
      Picture         =   "frmDaily.frx":06DE
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   270
      Width           =   1245
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
      Height          =   975
      Left            =   11850
      Picture         =   "frmDaily.frx":09E8
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   270
      Width           =   1245
   End
   Begin VB.Label lblTotal 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
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
      Left            =   90
      TabIndex        =   17
      Top             =   7560
      Width           =   480
   End
End
Attribute VB_Name = "frmDaily"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub bcancel_Click()

10  Unload Me

End Sub

Private Sub bprint_Click()

    Dim Y As Long
    Dim fs As Long

10  On Error GoTo bprint_Click_Error


20  If g.Rows = 2 Then
30      iMsg "Nothing to print", vbInformation
40      Exit Sub
50  End If

60  fs = Printer.FontSize

70  Printer.Orientation = vbPRORLandscape
80  Printer.FontSize = 12
90  Printer.Print
100 Printer.Print

110 Printer.Print "Daily worklist for "; Format$(dt, "dd/mmm/yyyy")
120 Printer.Print

130 Printer.FontSize = 8

140 For Y = 0 To g.Rows - 1
150     g.row = Y
160     g.Col = 0
170     Printer.Print g;
180     g.Col = 1    'name
190     Printer.Print Tab(17); Left$(g, 17);
200     g.Col = 2    'chart
210     Printer.Print Tab(35); Trim(g);
220     g.Col = 3    'gp
230     Printer.Print Tab(43); Left$(g, 16);
240     g.Col = 4    'ward
250     Printer.Print Tab(65); Left$(g, 17);
260     g.Col = 5    'clinician
270     Printer.Print Tab(85); g;
280     g.Col = 6    'clinician
290     Printer.Print Tab(110); g
300 Next

310 Printer.FontSize = fs

320 Printer.EndDoc

330 Exit Sub

bprint_Click_Error:

    Dim strES As String
    Dim intEL As Integer

340 intEL = Erl
350 strES = Err.Description
360 LogError "frmDaily", "bPrint_Click", intEL, strES

End Sub

Private Sub cmdRefresh_Click()

10  FillG

End Sub

Private Sub FillG()

    Dim tb As New Recordset
    Dim s As String
    Dim sql As String
    Dim Asql As String
    Dim Bsql As String
    Dim OldSampleID As String
    Dim NewSampleID As String
    Dim Disc As String
    Dim TestColumn As String
    Dim ResultColumn As String


10  On Error GoTo FillG_Error

20  ClearFGrid g

30  If optBio Then
40      Disc = "Bio"
50      TestColumn = "ShortName"
60  ElseIf optHaem Then
70      Disc = "Haem"
80      TestColumn = "AnalyteName"
90      ResultColumn = "RBC"
100 ElseIf optCoag Then
110     Disc = "Coag"
120     TestColumn = "TestName"
130 ElseIf optExt Then
140     Disc = "Ext"
150     TestColumn = "Analyte"
160 ElseIf optEnd Then
170     Disc = "End"
180     TestColumn = "ShortName"
190 ElseIf optImm Then
200     Disc = "Imm"
210     TestColumn = "ShortName"
220 ElseIf optBG Then
230     Disc = "Bga"
240     TestColumn = "ShortName"
250 ElseIf optCyto Then
260     Disc = "Cyto"
270     TestColumn = ""
280 ElseIf optHisto Then
290     Disc = "Histo"
300     TestColumn = ""
310 ElseIf optMicro Then
320     sql = sql & ", Urine "
330     Asql = " AND D.SampleID = Urine.SampleID "
340 ElseIf optSemen Then
350     Disc = "Semen"
360     TestColumn = ""
370     ResultColumn = "Motility"
380 Else
390     g.Visible = True
400     iMsg "No Discipline Choosen!"
410     Exit Sub
420 End If



430 If optBio Or optCoag Or optEnd Or optImm Or optBG Then
440     Asql = "SELECT DISTINCT(D.SampleID), D.PatName, D.hYear, D.Chart, D.GP, D.Ward, D.Clinician, T.TestResult " & _
               "FROM Demographics D "
450     Bsql = "INNER JOIN (Select * From " & Disc & "Requests WHERE COALESCE(Programmed, 0) <> 2) R ON D.SampleID = R.SampleID " & _
               "INNER JOIN (Select Code, " & TestColumn & " AS TestResult FROM " & Disc & "TestDefinitions) T ON R.Code = T.Code " & _
               "WHERE D.rundate BETWEEN CONVERT(DATETIME,'" & Format(dt & " " & tFromTime.Text, "dd/mmm/yyyy hh:mm:ss") & "',102) and CONVERT(DATETIME,'" & Format(dt & " " & tToTime.Text, "dd/mmm/yyyy hh:mm:ss") & "',102)"
460     sql = Asql & Bsql
        'Bsql = "INNER JOIN " & Disc & "Results R ON D.SampleID = R.SampleID " & _
         "INNER JOIN (Select Code, " & TestColumn & " AS TestResult FROM " & Disc & "TestDefinitions) T ON R.Code = T.Code " & _
         "WHERE D.RunDate = '" & Format$(dt, "dd/mmm/yyyy") & "' "
        'sql = sql & " UNION " & Asql & Bsql
470     sql = sql & "ORDER BY D.SampleID"
480 ElseIf optExt Then
490     Asql = "SELECT DISTINCT(D.SampleID), D.PatName, D.hYear, D.Chart, D.GP, D.Ward, D.Clinician, R.TestResult " & _
               "FROM Demographics D "
500     Bsql = "INNER JOIN (SELECT SampleID, " & TestColumn & " AS TestResult, Result, SentDate FROM " & Disc & "Results) R ON D.SampleID = R.SampleID " & _
               "WHERE R.SentDate BETWEEN '" & Format$(dt, "dd/mmm/yyyy 00:00:01") & _
               "' AND '" & Format$(dt, "dd/mmm/yyyy 23:59:59") & "' AND COALESCE(R.Result, '') = '' "
510     sql = Asql & Bsql & "ORDER BY D.SampleID"
520 ElseIf optSemen Or optHaem Then
530     Asql = "SELECT DISTINCT(D.SampleID), D.PatName, D.hYear, D.Chart, D.GP, D.Ward, D.Clinician, '' AS TestResult " & _
               "FROM Demographics D "
540     Bsql = "INNER JOIN " & Disc & "Results R ON D.SampleID = R.SampleID " & _
               "WHERE D.rundate BETWEEN CONVERT(DATETIME,'" & Format(dt & " " & tFromTime.Text, "dd/mmm/yyyy hh:mm:ss") & "',102) and CONVERT(DATETIME,'" & Format(dt & " " & tToTime.Text, "dd/mmm/yyyy hh:mm:ss") & "',102) AND COALESCE(" & ResultColumn & ", '') = '' "
550     sql = Asql & Bsql & "ORDER BY D.SampleID"
560 ElseIf optMicro Then
570     Asql = "SELECT DISTINCT(D.SampleID), D.PatName, D.hYear, D.Chart, D.GP, D.Ward, D.Clinician, '' AS TestResult " & _
               "FROM Demographics D "
580     Bsql = "INNER JOIN Urine R ON D.SampleID = R.SampleID " & _
               "WHERE D.rundate BETWEEN CONVERT(DATETIME,'" & Format(dt & " " & tFromTime.Text, "dd/mmm/yyyy hh:mm:ss") & "',102) and CONVERT(DATETIME,'" & Format(dt & " " & tToTime.Text, "dd/mmm/yyyy hh:mm:ss") & "',102)"
590     sql = Asql & Bsql & "ORDER BY D.SampleID"
600 End If



610 OldSampleID = ""
620 NewSampleID = ""
630 Set tb = New Recordset
640 RecOpenServer 0, tb, sql
650 Do While Not tb.EOF
660     OldSampleID = NewSampleID
670     NewSampleID = tb!SampleID
680     If OldSampleID <> NewSampleID Then
            'convert sample ids for histo,cyto,micro and semen
690         If optHisto Then
700             s = Val(tb!Hyear) & "/" & Val(tb!SampleID) - (SysOptHistoOffset(0) + (Val(Swap_Year(Trim(tb!Hyear))) * 1000)) & "H"
710         ElseIf optCyto Then
720             s = Val(tb!Hyear) & "/" & Val(tb!SampleID) - (SysOptCytoOffset(0) + (Val(Swap_Year(Trim(tb!Hyear))) * 1000)) & "C"
730         ElseIf optMicro Then
740             s = Val(tb!SampleID) - SysOptMicroOffset(0)
750         ElseIf optSemen Then
760             s = Val(tb!SampleID) - SysOptSemenOffset(0)
770         Else
780             s = tb!SampleID
790         End If

800         s = s & vbTab & _
                tb!PatName & vbTab & _
                tb!Chart & vbTab & _
                tb!GP & vbTab & _
                tb!Ward & vbTab & _
                tb!Clinician & "" & vbTab & _
                tb!TestResult & ""

810         g.AddItem s
820     Else
830         g.TextMatrix(g.Rows - 1, 6) = g.TextMatrix(g.Rows - 1, 6) & ", " & tb!TestResult & ""
840     End If


850     tb.MoveNext


860 Loop

870 FixG g

880 If g.Rows > 2 Then
890     lblTotal = "Total samples : " & g.Rows - 1
900 Else
910     lblTotal = ""
920 End If

930 Exit Sub

FillG_Error:

    Dim strES As String
    Dim intEL As Integer

940 intEL = Erl
950 strES = Err.Description
960 LogError "frmDaily", "FillG", intEL, strES, sql

End Sub

'Private Sub FillG()
'
'Dim tb As New Recordset
'Dim s As String
'Dim sql As String
'Dim T As String
'Dim sn As New Recordset
'Dim atb As Recordset
'Dim Asql As String
'
'On Error GoTo FillG_Error
'
'ClearFGrid g
'
'sql = "SELECT DISTINCT(D.SampleID), D.PatName, D.hYear, D.Chart, D.GP, D.Ward, D.Clinician FROM Demographics D "
'If optBio Then
'    sql = sql & ", BioRequests "
'    Asql = " AND D.SampleID = BioRequests.SampleID "
'ElseIf optHaem Then
'    sql = sql & ", HaemResults "
'    Asql = " AND D.SampleID = HaemResults.SampleID "
'ElseIf optCoag Then
'    sql = sql & ", CoagRequests "
'    Asql = " AND D.SampleID = CoagRequests.SampleID "
'ElseIf optExt Then
'    sql = sql & ", ExtResults "
'    Asql = " AND D.SampleID = ExtResults.SampleID "
'ElseIf optEnd Then
'    sql = sql & ", EndRequests "
'    Asql = " AND D.SampleID = EndRequests.SampleID "
'ElseIf optImm Then
'    sql = sql & ", ImmRequests "
'    Asql = " AND D.SampleID = ImmRequests.SampleID "
'ElseIf optBG Then
'    sql = sql & ", BGAResults "
'    Asql = " AND D.SampleID = BGAResults.SampleID "
'ElseIf optCyto Then
'    sql = sql & " , CytoResults "
'    Asql = " AND D.SampleID = CytoResults.SampleID "
'ElseIf optHisto Then
'    sql = sql & ", HistoResults "
'    Asql = " AND D.Sampleid = HistoResults.SampleID "
'ElseIf optMicro Then
'    sql = sql & ", Urine "
'    Asql = " AND D.SampleID = Urine.SampleID "
'ElseIf optSemen Then
'    sql = sql & ", SemenResults "
'    Asql = " AND D.SampleID = SemenResults.SampleID "
'Else
'    g.Visible = True
'    iMsg "No Discipline Choosen!"
'    Exit Sub
'End If
'
'sql = sql & "WHERE D.RunDate = '" & Format$(dt, "dd/mmm/yyyy") & "' "
'sql = sql & Asql
'sql = sql & "ORDER BY D.SampleID"
'
'Set tb = New Recordset
'RecOpenServer 0, tb, sql
'Do While Not tb.EOF
'    If optHisto Then
'        s = Val(tb!Hyear) & "/" & Val(tb!SampleID) - (SysOptHistoOffset(0) + (Val(Swap_Year(Trim(tb!Hyear))) * 1000)) & "H"
'    ElseIf optCyto Then
'        s = Val(tb!Hyear) & "/" & Val(tb!SampleID) - (SysOptCytoOffset(0) + (Val(Swap_Year(Trim(tb!Hyear))) * 1000)) & "C"
'    ElseIf optMicro Then
'        s = Val(tb!SampleID) - SysOptMicroOffset(0)
'    ElseIf optSemen Then
'        s = Val(tb!SampleID) - SysOptSemenOffset(0)
'    Else
'        s = tb!SampleID
'    End If
'    s = s & vbTab & _
     '        tb!PatName & vbTab & _
     '        tb!Chart & vbTab & _
     '        tb!GP & vbTab & _
     '        tb!Ward & vbTab & _
     '        tb!Clinician & ""
'
'    If optExt Then
'        Set sn = New Recordset
'        RecOpenServer 0, sn, "SELECT * from extresults WHERE sampleid = '" & tb!SampleID & "'"
'        Do While Not sn.EOF
'            T = T & sn!Analyte & ", "
'            sn.MoveNext
'        Loop
'        T = Trim(T)
'        If T <> "" Then T = Left(T, Len(T) - 1)
'        s = s & vbTab & T
'        T = ""
'    ElseIf optBio Then
'        Set sn = New Recordset
'        RecOpenServer 0, sn, "SELECT * from BIOREquests WHERE sampleid = '" & tb!SampleID & "'"
'        Do While Not sn.EOF
'            sql = "SELECT shortname from biotestdefinitions   WHERE code = '" & sn!Code & "'"
'            Set atb = New Recordset
'            RecOpenServer 0, atb, sql
'            If Not atb.EOF Then T = T & Trim(atb!ShortName & "") & ", "
'            sn.MoveNext
'        Loop
'        If Len(T) > 1 Then T = Mid(T, 1, Len(T) - 1)
'        s = s & vbTab & T
'        T = ""
'    ElseIf optImm Then
'        Set sn = New Recordset
'        RecOpenServer 0, sn, "SELECT * from ImmRequests WHERE sampleid = '" & tb!SampleID & "'"
'        Do While Not sn.EOF
'            sql = "SELECT shortname from ImmTestDefinitions WHERE code = '" & sn!Code & "'"
'            Set atb = New Recordset
'            RecOpenServer 0, atb, sql
'            If Not atb.EOF Then T = T & Trim(atb!ShortName & "") & ", "
'            sn.MoveNext
'        Loop
'        If Len(T) > 1 Then T = Mid(T, 1, Len(T) - 1)
'        s = s & vbTab & T
'        T = ""
'    ElseIf optEnd Then
'        Set sn = New Recordset
'        RecOpenServer 0, sn, "SELECT * from endRequests WHERE sampleid = '" & tb!SampleID & "'"
'        Do While Not sn.EOF
'            sql = "SELECT shortname from endTestDefinitions WHERE code = '" & sn!Code & "'"
'            Set atb = New Recordset
'            RecOpenServer 0, atb, sql
'            If Not atb.EOF Then T = T & Trim(atb!ShortName & "") & ", "
'            sn.MoveNext
'        Loop
'        If Len(T) > 1 Then T = Mid(T, 1, Len(T) - 1)
'        s = s & vbTab & T
'        T = ""
'    ElseIf optCoag Then
'
'        Set sn = New Recordset
'        RecOpenServer 0, sn, "SELECT * from CoagRequests WHERE sampleid = '" & tb!SampleID & "' AND COALESCE(Programmed,0) <> 2"
'        Do While Not sn.EOF
'            sql = "SELECT testname from CoagTestDefinitions WHERE code = '" & sn!Code & "'"
'            Set atb = New Recordset
'            RecOpenServer 0, atb, sql
'            If Not atb.EOF Then T = T & Trim(atb!TestName & "") & ", "
'            sn.MoveNext
'        Loop
'
'
'        T = Trim(T)
'        If Len(T) > 1 Then T = Mid(T, 1, Len(T) - 1)
'        s = s & vbTab & T
'        T = ""
'
'    End If
'    g.AddItem s
'    tb.MoveNext
'Loop
'
'FixG g
'
'Exit Sub
'
'FillG_Error:
'
'Dim strES As String
'Dim intEL As Integer
'
'intEL = Erl
'strES = Err.Description
'LogError "frmDaily", "FillG", intEL, strES, sql
'
'End Sub



Private Sub Form_Load()

10  On Error GoTo Form_Load_Error



20  dt = Format$(Now, "dd/mm/yyyy")
30  tFromTime.Text = "00:00:00"
40  tToTime.Text = "23:59:59"
50  optEnd.Enabled = SysOptDeptEnd(0)
60  optImm.Enabled = SysOptDeptImm(0)
70  optBG.Enabled = SysOptDeptBga(0)
80  optMicro.Enabled = False    'SysOptDeptMicro(0)
90  optSemen.Enabled = SysOptDeptSemen(0)
100 optHisto.Enabled = SysOptDeptHisto(0)
110 optCyto.Enabled = SysOptDeptCyto(0)

120 SetFormCaption

130 Exit Sub

Form_Load_Error:

    Dim strES As String
    Dim intEL As Integer

140 intEL = Erl
150 strES = Err.Description
160 LogError "frmDaily", "Form_Load", intEL, strES

End Sub

Private Sub g_Click()

    Static SortOrder As Boolean

10  On Error GoTo g_Click_Error

20  If g.MouseRow = 0 Then
30      If SortOrder Then
40          g.Sort = flexSortGenericAscending
50      Else
60          g.Sort = flexSortGenericDescending
70      End If
80      SortOrder = Not SortOrder
90  End If

100 Exit Sub

g_Click_Error:

    Dim strES As String
    Dim intEL As Integer

110 intEL = Erl
120 strES = Err.Description
130 LogError "frmDaily", "g_Click", intEL, strES

End Sub

Private Sub optBG_Click()
10  On Error GoTo optBG_Click_Error

20  SetFormCaption
30  FillG
40  g.ColWidth(6) = 2500

50  Exit Sub

optBG_Click_Error:

    Dim strES As String
    Dim intEL As Integer

60  intEL = Erl
70  strES = Err.Description
80  LogError "frmDaily", "optBG_Click", intEL, strES

End Sub

Private Sub optBio_Click()
10  On Error GoTo optBio_Click_Error

20  SetFormCaption
30  FillG
40  g.ColWidth(6) = 2500

50  Exit Sub

optBio_Click_Error:

    Dim strES As String
    Dim intEL As Integer

60  intEL = Erl
70  strES = Err.Description
80  LogError "frmDaily", "optBio_Click", intEL, strES

End Sub

Private Sub optCoag_Click()
10  On Error GoTo optCoag_Click_Error

20  SetFormCaption
30  FillG
40  g.ColWidth(6) = 2500

50  Exit Sub

optCoag_Click_Error:

    Dim strES As String
    Dim intEL As Integer

60  intEL = Erl
70  strES = Err.Description
80  LogError "frmDaily", "optCoag_Click", intEL, strES

End Sub

Private Sub optCyto_Click()
10  On Error GoTo optCyto_Click_Error

20  SetFormCaption
30  FillG
40  g.ColWidth(6) = 2500

50  Exit Sub

optCyto_Click_Error:

    Dim strES As String
    Dim intEL As Integer

60  intEL = Erl
70  strES = Err.Description
80  LogError "frmDaily", "optCyto_Click", intEL, strES

End Sub

Private Sub optEnd_Click()
10  On Error GoTo optEnd_Click_Error

20  SetFormCaption
30  FillG
40  g.ColWidth(6) = 2500

50  Exit Sub

optEnd_Click_Error:

    Dim strES As String
    Dim intEL As Integer

60  intEL = Erl
70  strES = Err.Description
80  LogError "frmDaily", "optEnd_Click", intEL, strES

End Sub

Private Sub optExt_Click()

10  On Error GoTo optExt_Click_Error

20  SetFormCaption
30  FillG
40  g.ColWidth(6) = 2500

50  Exit Sub

optExt_Click_Error:

    Dim strES As String
    Dim intEL As Integer

60  intEL = Erl
70  strES = Err.Description
80  LogError "frmDaily", "optExt_Click", intEL, strES

End Sub

Private Sub optHaem_Click()
10  On Error GoTo optHaem_Click_Error

20  SetFormCaption
30  FillG
40  g.ColWidth(6) = 0

50  Exit Sub

optHaem_Click_Error:

    Dim strES As String
    Dim intEL As Integer

60  intEL = Erl
70  strES = Err.Description
80  LogError "frmDaily", "optHaem_Click", intEL, strES

End Sub

Private Sub optHisto_Click()
10  On Error GoTo optHisto_Click_Error

20  SetFormCaption
30  FillG
40  g.ColWidth(6) = 2500

50  Exit Sub

optHisto_Click_Error:

    Dim strES As String
    Dim intEL As Integer

60  intEL = Erl
70  strES = Err.Description
80  LogError "frmDaily", "optHisto_Click", intEL, strES

End Sub

Private Sub optImm_Click()
10  On Error GoTo optImm_Click_Error

20  SetFormCaption
30  FillG
40  g.ColWidth(6) = 2500

50  Exit Sub

optImm_Click_Error:

    Dim strES As String
    Dim intEL As Integer

60  intEL = Erl
70  strES = Err.Description
80  LogError "frmDaily", "optImm_Click", intEL, strES

End Sub

Private Sub optMicro_Click()
10  On Error GoTo optMicro_Click_Error

20  SetFormCaption
30  FillG
40  g.ColWidth(6) = 2500

50  Exit Sub

optMicro_Click_Error:

    Dim strES As String
    Dim intEL As Integer

60  intEL = Erl
70  strES = Err.Description
80  LogError "frmDaily", "optMicro_Click", intEL, strES

End Sub

Private Sub optSemen_Click()
10  On Error GoTo optSemen_Click_Error

20  SetFormCaption
30  FillG
40  g.ColWidth(6) = 2500

50  Exit Sub

optSemen_Click_Error:

    Dim strES As String
    Dim intEL As Integer

60  intEL = Erl
70  strES = Err.Description
80  LogError "frmDaily", "optSemen_Click", intEL, strES

End Sub


Private Sub SetFormCaption()

10  On Error GoTo SetFormCaption_Error

20  If optBio.Value = True Then
30      Me.Caption = "NetAcquire - Daily " & optBio.Caption & " Worklist"
40  ElseIf optHaem.Value = True Then
50      Me.Caption = "NetAcquire - Daily " & optHaem.Caption & " Worklist"
60  ElseIf optCoag.Value = True Then
70      Me.Caption = "NetAcquire - Daily " & optCoag.Caption & " Worklist"
80  ElseIf optEnd.Value = True Then
90      Me.Caption = "NetAcquire - Daily " & optEnd.Caption & " Worklist"
100 ElseIf optImm.Value = True Then
110     Me.Caption = "NetAcquire - Daily " & optImm.Caption & " Worklist"
120 ElseIf optExt.Value = True Then
130     Me.Caption = "NetAcquire - Daily " & optExt.Caption & " Outstanding"
140 ElseIf optBG.Value = True Then
150     Me.Caption = "NetAcquire - Daily " & optBG.Caption & " Worklist"
160 ElseIf optHisto.Value = True Then
170     Me.Caption = "NetAcquire - Daily " & optHisto.Caption & " Worklist"
180 ElseIf optCyto.Value = True Then
190     Me.Caption = "NetAcquire - Daily " & optCyto.Caption & " Worklist"
200 ElseIf optMicro.Value = True Then
210     Me.Caption = "NetAcquire - Daily " & optMicro.Caption & " Worklist"
220 ElseIf optSemen.Value = True Then
230     Me.Caption = "NetAcquire - Daily " & optSemen.Caption & " Worklist"
240 End If
250 Exit Sub

SetFormCaption_Error:

    Dim strES As String
    Dim intEL As Integer

260 intEL = Erl
270 strES = Err.Description
280 LogError "frmDaily", "SetFormCaption", intEL, strES

End Sub

Private Sub tFromTime_Validate(Cancel As Boolean)
10  On Error GoTo tFromTime_Validate_Error

20  If IsDate(tFromTime.Text) = False Then
30      iMsg "Please Enter a Valid Time"
40      Cancel = True
50  End If


70  Exit Sub

tFromTime_Validate_Error:
    Dim strES As String
    Dim intEL As Integer
80  intEL = Erl
90  strES = Err.Description
100 LogError "frmDaily", "tFromTime_Validate", intEL, strES
End Sub

Private Sub tToTime_Validate(Cancel As Boolean)
10  On Error GoTo ttoTime_Validate_Error
20  If IsDate(tToTime.Text) = False Then
30      iMsg "Please Enter a Valid Time"
40      Cancel = True
50  Else
60      If tFromTime > tToTime Then
70          iMsg "Start Time should less than End Time"
80          Cancel = True
90      End If
100 End If
110 Exit Sub
ttoTime_Validate_Error:
    Dim strES As String
    Dim intEL As Integer
120 intEL = Erl
130 strES = Err.Description
140 LogError "frmDaily", "ttoTime_Validate", intEL, strES
End Sub
