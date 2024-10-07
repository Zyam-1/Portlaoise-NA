VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmGlucose 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Glucose Tolerance Test"
   ClientHeight    =   6390
   ClientLeft      =   510
   ClientTop       =   1695
   ClientWidth     =   11805
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
   Icon            =   "frmGlucose.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6390
   ScaleWidth      =   11805
   Begin VB.CommandButton bPrintSeries 
      Caption         =   "Print as Glucose Series"
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
      Left            =   7410
      Picture         =   "frmGlucose.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   5040
      Width           =   1875
   End
   Begin VB.Frame Frame2 
      Height          =   2985
      Left            =   6570
      TabIndex        =   5
      Top             =   90
      Width           =   4995
      Begin VB.Label Label6 
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
         Left            =   765
         TabIndex        =   29
         Top             =   570
         Width           =   420
      End
      Begin VB.Label lblName 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1305
         TabIndex        =   28
         Top             =   540
         Width           =   3465
      End
      Begin VB.Label lsex 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4365
         TabIndex        =   23
         Top             =   870
         Width           =   405
      End
      Begin VB.Label lage 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3225
         TabIndex        =   22
         Top             =   870
         Width           =   555
      End
      Begin VB.Label lgp 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1305
         TabIndex        =   21
         Top             =   2370
         Width           =   3465
      End
      Begin VB.Label lward 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1305
         TabIndex        =   20
         Top             =   2070
         Width           =   3465
      End
      Begin VB.Label lclinician 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1305
         TabIndex        =   19
         Top             =   1770
         Width           =   3465
      End
      Begin VB.Label laddr1 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1305
         TabIndex        =   18
         Top             =   1470
         Width           =   3465
      End
      Begin VB.Label laddr0 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1305
         TabIndex        =   17
         Top             =   1170
         Width           =   3465
      End
      Begin VB.Label ldob 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1305
         TabIndex        =   16
         Top             =   870
         Width           =   1245
      End
      Begin VB.Label lchart 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1305
         TabIndex        =   15
         Top             =   210
         Width           =   1245
      End
      Begin VB.Label label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ward"
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
         Index           =   10
         Left            =   735
         TabIndex        =   14
         Top             =   2100
         Width           =   390
      End
      Begin VB.Label label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Consultant"
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
         Index           =   9
         Left            =   375
         TabIndex        =   13
         Top             =   1800
         Width           =   750
      End
      Begin VB.Label label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Age"
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
         Index           =   8
         Left            =   2880
         TabIndex        =   12
         Top             =   900
         Width           =   285
      End
      Begin VB.Label label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "D.o.B."
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
         Index           =   7
         Left            =   675
         TabIndex        =   11
         Top             =   900
         Width           =   450
      End
      Begin VB.Label label1 
         Alignment       =   1  'Right Justify
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
         Index           =   6
         Left            =   4065
         TabIndex        =   10
         Top             =   855
         Width           =   270
      End
      Begin VB.Label label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Addr1"
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
         Index           =   5
         Left            =   705
         TabIndex        =   9
         Top             =   1200
         Width           =   420
      End
      Begin VB.Label label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Chart Number"
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
         Index           =   2
         Left            =   150
         TabIndex        =   8
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Addr2"
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
         Left            =   705
         TabIndex        =   7
         Top             =   1500
         Width           =   420
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "G. P."
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
         Left            =   765
         TabIndex        =   6
         Top             =   2400
         Width           =   360
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2985
      Left            =   210
      TabIndex        =   3
      Top             =   90
      Width           =   6255
      Begin MSFlexGridLib.MSFlexGrid gNames 
         Height          =   2535
         Left            =   60
         TabIndex        =   26
         Top             =   210
         Width           =   3915
         _ExtentX        =   6906
         _ExtentY        =   4471
         _Version        =   393216
         FixedCols       =   0
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   0
         GridLines       =   3
         GridLinesFixed  =   3
         ScrollBars      =   2
         FormatString    =   "<Name                             |<Date of Birth  "
      End
      Begin MSComCtl2.DTPicker dtRun 
         Height          =   315
         Left            =   4500
         TabIndex        =   24
         Top             =   1020
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         _Version        =   393216
         Format          =   59768833
         CurrentDate     =   37501
      End
      Begin MSComctlLib.ProgressBar pb 
         Height          =   165
         Left            =   60
         TabIndex        =   4
         Top             =   2760
         Width           =   3885
         _ExtentX        =   6853
         _ExtentY        =   291
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Run Date"
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
         Left            =   4770
         TabIndex        =   25
         Top             =   810
         Width           =   690
      End
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   2985
      Left            =   300
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3240
      Width           =   5685
      _ExtentX        =   10028
      _ExtentY        =   5265
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      HighLight       =   0
      GridLines       =   2
      ScrollBars      =   2
      FormatString    =   "<Run #          |<Date/Time                |<Serum  |<H/L|<Urine  |<H/L"
   End
   Begin VB.CommandButton bPrintGTT 
      Caption         =   "&Print as GTT Report"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   7410
      Picture         =   "frmGlucose.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4005
      Width           =   1875
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
      Height          =   870
      Left            =   9540
      Picture         =   "frmGlucose.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4005
      Width           =   1380
   End
End
Attribute VB_Name = "frmGlucose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private Sub FillNames()

          Dim tb1 As Recordset
          Dim sn As New Recordset
          Dim sql As String
          Dim n As Long
          Dim Name2Find As String
          Dim Found As Long
          Dim X As Long
          Dim s As String



10        g.Rows = 2
20        g.AddItem ""
30        g.RemoveItem 1

40        sql = "SELECT distinct demographics.patname, demographics.DoB  from demographics, bioresults WHERE " & _
                "demographics.RunDate = '" & Format$(dtRun, "dd/mmm/yyyy") & "' and demographics.sampleid = bioresults.sampleid  " & _
                "and (bioresults.Code = '" & SysOptBioCodeForFastGlucose(0) & "' " & _
                "or bioresults.Code = '" & SysOptBioCodeForGlucose1(0) & _
                "' or bioresults.Code = '" & SysOptBioCodeForGlucose2(0) & _
                "' or bioresults.Code = '" & SysOptBioCodeForGlucose3(0) & _
                "' or bioresults.Code = '" & SysOptBioCodeForFastGlucoseP(0) & _
                "' or bioresults.Code = '" & SysOptBioCodeForGlucose1P(0) & _
                "' or bioresults.Code = '" & SysOptBioCodeForGlucose2P(0) & _
                "' or bioresults.Code = '" & SysOptBioCodeForGlucose3P(0) & "')"


50        Set sn = New Recordset
60        RecOpenServer 0, sn, sql

70        gNames.Visible = False
80        gNames.Rows = 2
90        gNames.AddItem ""
100       gNames.RemoveItem 1

110       If sn.EOF Then
120           gNames.AddItem "None Found"
130           gNames.RemoveItem 1
140           gNames.Visible = True
150           Exit Sub
160       End If

170       Do While Not sn.EOF
180           s = sn!PatName & vbTab
190           If IsDate(sn!Dob) Then
200               s = s & Format$(sn!Dob, "dd/mm/yyyy")
210           End If
220           gNames.AddItem s
230           sn.MoveNext
240       Loop

250       pb.Visible = True
260       pb.Max = gNames.Rows



270       For n = gNames.Rows - 1 To 2 Step -1
280           pb = pb.Max - n
290           Name2Find = gNames.TextMatrix(n, 0)
300           For X = Len(Name2Find) - 1 To 1 Step -1
310               If Mid$(Name2Find, X, 1) = "'" Then
320                   Name2Find = Left$(Name2Find, X) & Mid$(Name2Find, X)
330               End If
340           Next
350           Found = 0
360           sql = "SELECT * from BioResults, Demographics WHERE " & _
                    "patname = '" & Name2Find & "' " & _
                    "and BioResults.SampleID = Demographics.SampleID " & _
                    "and (BioResults.Code = '" & SysOptBioCodeForFastGlucose(0) & "' " & _
                    " or BioResults.Code = '" & SysOptBioCodeForGlucose1(0) & "'" & _
                    " or BioResults.Code = '" & SysOptBioCodeForGlucose2(0) & "'" & _
                    " or BioResults.Code = '" & SysOptBioCodeForGlucose3(0) & "'" & _
                    " or BioResults.Code = '" & SysOptBioCodeForFastGlucoseP(0) & "' " & _
                    " or BioResults.Code = '" & SysOptBioCodeForGlucose1P(0) & "'" & _
                    " or BioResults.Code = '" & SysOptBioCodeForGlucose2P(0) & "'" & _
                    " or BioResults.Code = '" & SysOptBioCodeForGlucose3P(0) & "') " & _
                    "and demographics.rundate = '" & Format$(dtRun, "dd/mmm/yyyy") & "'"

370           If IsDate(gNames.TextMatrix(n, 1)) Then
380               sql = sql & " and DoB = '" & Format$(gNames.TextMatrix(n, 1), "dd/mmm/yyyy") & "'"
390           End If

400           Set tb1 = New Recordset
410           RecOpenClient 0, tb1, sql
420           If Not tb1.EOF Then
430               If tb1.RecordCount < 2 Then
440                   gNames.RemoveItem n
450               End If
460           Else
470               gNames.RemoveItem n
480           End If
490       Next
500       pb.Visible = False

510       If gNames.Rows > 2 Then
520           gNames.RemoveItem 1
530       End If
540       gNames.Visible = True



550       Exit Sub

FillNames_Error:

          Dim strES As String
          Dim intEL As Integer

560       intEL = Erl
570       strES = Err.Description
580       LogError "frmGlucose", "FillNames", intEL, strES, sql


End Sub


Private Sub bcancel_Click()

10        Unload Me

End Sub

Private Sub bPrintGTT_Click()

          Dim tb As New Recordset
          Dim sql As String
          Dim SampleID As String
          Dim strWard As String
          Dim strGp As String
          Dim strClin As String

10        On Error GoTo bPrintGTT_Click_Error

20        SampleID = g.TextMatrix(1, 0)
30        If Trim$(SampleID) = "" Then
40            iMsg "Nothing to do!" & vbCrLf & "SELECT a Name to Print.", vbExclamation
50            Exit Sub
60        End If

70        strWard = ""
80        strGp = ""
90        strClin = ""

100       sql = "SELECT * FROM Demographics WHERE " & _
                "SampleID = '" & SampleID & "'"
110       Set tb = New Recordset
120       RecOpenServer 0, tb, sql
130       If Not tb.EOF Then
140           strWard = tb!Ward & ""
150           strClin = tb!Clinician & ""
160           strGp = tb!GP & ""
170       End If

180       sql = "SELECT * FROM PrintPending WHERE " & _
                "Department = 'G' " & _
                "AND SampleID = '" & SampleID & "'"
190       Set tb = New Recordset
200       RecOpenClient 0, tb, sql
210       If tb.EOF Then
220           tb.AddNew
230       End If
240       tb!SampleID = SampleID
250       tb!Department = "G"
260       tb!Initiator = Username
270       tb!Ward = strWard
280       tb!Clinician = strClin
290       tb!GP = strGp
300       tb.Update

310       Exit Sub

bPrintGTT_Click_Error:

          Dim strES As String
          Dim intEL As Integer



320       intEL = Erl
330       strES = Err.Description
340       LogError "frmGlucose", "bPrintGTT_Click", intEL, strES, sql


End Sub


Private Sub bPrintSeries_Click()

          Dim tb As New Recordset
          Dim sql As String
          Dim SampleID As String
          Dim strWard As String
          Dim strGp As String
          Dim strClin As String

10        On Error GoTo bPrintSeries_Click_Error

20        SampleID = g.TextMatrix(1, 0)
30        If Trim$(SampleID) = "" Then
40            iMsg "Nothing to do!" & vbCrLf & "SELECT a Name to Print.", vbExclamation
50            Exit Sub
60        End If

70        strWard = ""
80        strGp = ""
90        strClin = ""

100       sql = "SELECT * FROM Demographics WHERE " & _
                "SampleID = '" & SampleID & "'"

110       Set tb = New Recordset
120       RecOpenServer 0, tb, sql
130       If Not tb.EOF Then
140           strWard = tb!Ward & ""
150           strClin = tb!Clinician & ""
160           strGp = tb!GP & ""
170       End If

180       sql = "SELECT * FROM PrintPending WHERE " & _
                "Department = 'S' " & _
                "AND SampleID = '" & SampleID & "'"
190       Set tb = New Recordset
200       RecOpenClient 0, tb, sql
210       If tb.EOF Then
220           tb.AddNew
230       End If
240       tb!SampleID = SampleID
250       tb!Department = "S"
260       tb!Initiator = Username
270       tb!Initiator = Username
280       tb!Ward = strWard
290       tb!Clinician = strClin
300       tb!GP = strGp
310       tb.Update

320       Exit Sub

bPrintSeries_Click_Error:

          Dim strES As String
          Dim intEL As Integer



330       intEL = Erl
340       strES = Err.Description
350       LogError "frmGlucose", "bPrintSeries_Click", intEL, strES, sql


End Sub

Private Sub dtRun_CloseUp()

10        On Error GoTo dtRun_CloseUp_Error

20        FillNames

30        lChart = ""
40        lDoB = ""
50        lAge = ""
60        lSex = ""
70        laddr0 = ""
80        laddr1 = ""
90        lclinician = ""
100       lward = ""
110       lgp = ""

120       Exit Sub

dtRun_CloseUp_Error:

          Dim strES As String
          Dim intEL As Integer



130       intEL = Erl
140       strES = Err.Description
150       LogError "frmGlucose", "dtRun_CloseUp", intEL, strES


End Sub


Private Sub Form_Activate()

10        On Error GoTo Form_Activate_Error

20        FillNames

30        Exit Sub

Form_Activate_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmGlucose", "Form_Activate", intEL, strES


End Sub

Private Sub Form_Load()


10        On Error GoTo Form_Load_Error

20        dtRun = Format$(Now, "dd/mm/yyyy")

30        Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmGlucose", "Form_Load", intEL, strES


End Sub

Private Sub gNames_Click()

          Dim tb As New Recordset
          Dim sql As String
          Dim s As String
          Dim Found As Long
          Dim Name2Find As String
          Dim X As Long
          Dim BRs As New BIEResults
          Dim BRres As BIEResults
          Dim br As BIEResult
          Dim n As Long
          Dim i As Integer
          Dim CodeSerum As String
          Dim CodePlasma As String

10        On Error GoTo gNames_Click_Error

20        If gNames.MouseRow = 0 Then
30            Exit Sub
40        End If

50        Name2Find = gNames.TextMatrix(gNames.Row, 0)
60        For X = Len(Name2Find) - 1 To 1 Step -1
70            If Mid$(Name2Find, X, 1) = "'" Then
80                Name2Find = Left$(Name2Find, X) & Mid$(Name2Find, X)
90            End If
100       Next

110       g.Rows = 2
120       g.AddItem ""
130       g.RemoveItem 1




140       For i = 1 To 4
150           CodeSerum = Choose(i, SysOptBioCodeForFastGlucose(0), SysOptBioCodeForGlucose1(0), SysOptBioCodeForGlucose2(0), SysOptBioCodeForGlucose3(0))
160           CodePlasma = Choose(i, SysOptBioCodeForFastGlucoseP(0), SysOptBioCodeForGlucose1P(0), SysOptBioCodeForGlucose2P(0), SysOptBioCodeForGlucose3P(0))
170           sql = "SELECT d.* , b.Result, b.Code, b.Units, b.Operator from demographics d " & _
                    "left join BioResults b on d.SampleID = b.SampleId " & _
                    "WHERE d.patname = '" & Name2Find & "' and d.RunDate = '" & Format$(dtRun, "dd/mmm/yyyy") & "' " & _
                    "and (Code = '" & CodeSerum & "' or Code = '" & CodePlasma & "') "

180           If IsDate(gNames.TextMatrix(gNames.Row, 1)) Then
190               sql = sql & "and d.DoB = '" & Format$(gNames.TextMatrix(gNames.Row, 1), "dd/mmm/yyyy") & "' "
200           End If
210           Set tb = New Recordset
220           RecOpenClient 0, tb, sql

230           If Not tb.EOF Then
240               lChart = tb!Chart & ""
250               If Not IsNull(tb!Dob) Then
260                   lDoB = tb!Dob
270               Else
280                   lDoB = ""
290               End If
300               lAge = tb!Age & ""
310               lSex = tb!sex & ""
320               laddr0 = tb!Addr0 & ""
330               laddr1 = tb!Addr1 & ""
340               lclinician = tb!Clinician & ""
350               lward = tb!Ward & ""
360               lgp = tb!GP & ""
370               lblName = tb!PatName & ""

380               Set BRres = BRs.Load("Bio", tb!SampleID, "Results", gDONTCARE, gDONTCARE, 0, "", dtRun)
390               If Not BRres Is Nothing Then
400                   For Each br In BRres
410                       If br.Code = CodeSerum Or br.Code = CodePlasma Then
420                           If br.SampleType = "S" Or br.SampleType = "PL" Then
430                               Found = True
440                               s = br.SampleID & vbTab & Format$(tb!SampleDate & "", "dd/mm/yyyy hh:mm") & vbTab
450                               s = s & Format$(br.Result, "0.0") & vbTab & Left(QuickInterpBio(br), 1)
460                               Found = True
470                           End If
480                       End If
490                       If Found Then
500                           g.AddItem s
510                           Found = False
520                           s = ""
530                       End If

540                   Next

550               End If
560           End If
570       Next i



580       If g.Rows > 2 Then
590           g.RemoveItem 1
600       End If

610       For n = 1 To g.Rows - 1
620           If g.TextMatrix(n, 3) = "L" Then
630               g.Row = n
640               g.Col = 3
650               g.CellBackColor = vbBlue
660               g.CellForeColor = vbYellow
670           ElseIf g.TextMatrix(n, 3) = "H" Then
680               g.Row = n
690               g.Col = 3
700               g.CellBackColor = vbRed
710               g.CellForeColor = vbYellow
720           End If
730           If g.TextMatrix(n, 5) = "L" Then
740               g.Row = n
750               g.Col = 5
760               g.CellBackColor = vbBlue
770               g.CellForeColor = vbYellow
780           ElseIf g.TextMatrix(n, 5) = "H" Then
790               g.Row = n
800               g.Col = 5
810               g.CellBackColor = vbRed
820               g.CellForeColor = vbYellow
830           End If
840       Next

850       Exit Sub

gNames_Click_Error:

          Dim strES As String
          Dim intEL As Integer

860       intEL = Erl
870       strES = Err.Description
880       LogError "frmGlucose", "gNames_Click", intEL, strES, sql

End Sub


'Private Sub gNames_Click()
'
'Dim tb As New Recordset
'Dim sql As String
'Dim s As String
'Dim Found As Long
'Dim Name2Find As String
'Dim X As Long
'Dim BRs As New BIEResults
'Dim BRres As BIEResults
'Dim br As BIEResult
'Dim n As Long
'
'On Error GoTo gNames_Click_Error
'
'If gNames.MouseRow = 0 Then
'    Exit Sub
'End If
'
'Name2Find = gNames.TextMatrix(gNames.Row, 0)
'For X = Len(Name2Find) - 1 To 1 Step -1
'    If Mid$(Name2Find, X, 1) = "'" Then
'        Name2Find = Left$(Name2Find, X) & Mid$(Name2Find, X)
'    End If
'Next
'
'sql = "SELECT * from demographics WHERE " & _
 '        "patname = '" & Name2Find & "' " & _
 '        "and rundate = '" & Format$(dtRun, "dd/mmm/yyyy") & "' "
'If IsDate(gNames.TextMatrix(gNames.Row, 1)) Then
'    sql = sql & "and DoB = '" & Format$(gNames.TextMatrix(gNames.Row, 1), "dd/mmm/yyyy") & "' "
'End If
'sql = sql & "order by SampleDate asc"
'Set tb = New Recordset
'RecOpenClient 0, tb, sql
'
'If tb.EOF Then
'    iMsg "No details found", vbInformation
'    Exit Sub
'End If
'
'lchart = tb!Chart & ""
'If Not IsNull(tb!Dob) Then
'    ldob = tb!Dob
'Else
'    ldob = ""
'End If
'lage = tb!Age & ""
'lsex = tb!sex & ""
'laddr0 = tb!Addr0 & ""
'laddr1 = tb!Addr1 & ""
'lclinician = tb!Clinician & ""
'lward = tb!Ward & ""
'lgp = tb!GP & ""
'lblName = tb!PatName & ""
'g.Rows = 2
'g.AddItem ""
'g.RemoveItem 1
'
'Do While Not tb.EOF
'
'    Found = False
'    Set BRres = BRs.Load("Bio", tb!SampleID, "Results", gDONTCARE, gDONTCARE, 0, "", dtRun)
'
'    If BRres Is Nothing Then
'    Else
'        For Each br In BRres
'
'            If br.Code = SysOptBioCodeForFastGlucose(0) Or br.Code = SysOptBioCodeForGlucose1(0) Or br.Code = SysOptBioCodeForGlucose2(0) Or br.Code = SysOptBioCodeForGlucose3(0) Or _
             '                    br.Code = SysOptBioCodeForFastGlucoseP(0) Or br.Code = SysOptBioCodeForGlucose1P(0) Or br.Code = SysOptBioCodeForGlucose2P(0) Or br.Code = SysOptBioCodeForGlucose3P(0) Then
'                If br.SampleType = "S" Or br.SampleType = "PL" Then
'                    s = br.SampleID & vbTab & Format$(tb!SampleDate & "", "dd/mm/yyyy hh:mm") & vbTab
'                    s = s & Format$(br.Result, "0.0") & vbTab & Left(QuickInterpBio(br), 1)
'                    Found = True
'                End If
'            End If
'            If Found Then
'                g.AddItem s, ListIndex
'                Found = False
'                s = ""
'            End If
'        Next
'    End If
'    tb.MoveNext
'Loop
'
'If g.Rows > 2 Then
'    g.RemoveItem 1
'End If
'
'For n = 1 To g.Rows - 1
'    If g.TextMatrix(n, 3) = "L" Then
'        g.Row = n
'        g.Col = 3
'        g.CellBackColor = vbBlue
'        g.CellForeColor = vbYellow
'    ElseIf g.TextMatrix(n, 3) = "H" Then
'        g.Row = n
'        g.Col = 3
'        g.CellBackColor = vbRed
'        g.CellForeColor = vbYellow
'    End If
'    If g.TextMatrix(n, 5) = "L" Then
'        g.Row = n
'        g.Col = 5
'        g.CellBackColor = vbBlue
'        g.CellForeColor = vbYellow
'    ElseIf g.TextMatrix(n, 5) = "H" Then
'        g.Row = n
'        g.Col = 5
'        g.CellBackColor = vbRed
'        g.CellForeColor = vbYellow
'    End If
'Next
'
'Exit Sub
'
'gNames_Click_Error:
'
'Dim strES As String
'Dim intEL As Integer
'
'intEL = Erl
'strES = Err.Description
'LogError "frmGlucose", "gNames_Click", intEL, strES, sql
'
'End Sub


