VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGluByName 
   Caption         =   "NetAcquire - Glucose Tolerance"
   ClientHeight    =   6870
   ClientLeft      =   240
   ClientTop       =   645
   ClientWidth     =   8970
   Icon            =   "frmGluByName.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   8970
   Begin MSComctlLib.ProgressBar pb 
      Height          =   225
      Left            =   150
      TabIndex        =   29
      Top             =   1230
      Visible         =   0   'False
      Width           =   3045
      _ExtentX        =   5371
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.CommandButton bPrintGTT 
      Caption         =   "&Print as GTT Report"
      Height          =   1005
      Left            =   7320
      Picture         =   "frmGluByName.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   3000
      Width           =   1305
   End
   Begin VB.CommandButton bPrintSeries 
      Caption         =   "Print as  Glucose Series"
      Height          =   1005
      Left            =   7320
      Picture         =   "frmGluByName.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   4170
      Width           =   1305
   End
   Begin VB.CommandButton bcancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   840
      Left            =   7335
      Picture         =   "frmGluByName.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   5760
      Width           =   1260
   End
   Begin VB.PictureBox SSPanel2 
      Height          =   2475
      Left            =   3540
      ScaleHeight     =   2415
      ScaleWidth      =   4995
      TabIndex        =   6
      Top             =   240
      Width           =   5055
      Begin VB.Label lsex 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4395
         TabIndex        =   24
         Top             =   480
         Width           =   405
      End
      Begin VB.Label lage 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3255
         TabIndex        =   23
         Top             =   480
         Width           =   555
      End
      Begin VB.Label lgp 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1335
         TabIndex        =   22
         Top             =   1980
         Width           =   3465
      End
      Begin VB.Label lward 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1335
         TabIndex        =   21
         Top             =   1680
         Width           =   3465
      End
      Begin VB.Label lclinician 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1335
         TabIndex        =   20
         Top             =   1380
         Width           =   3465
      End
      Begin VB.Label laddr1 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1335
         TabIndex        =   19
         Top             =   1080
         Width           =   3465
      End
      Begin VB.Label laddr0 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1335
         TabIndex        =   18
         Top             =   780
         Width           =   3465
      End
      Begin VB.Label ldob 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1335
         TabIndex        =   17
         Top             =   480
         Width           =   1245
      End
      Begin VB.Label lchart 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1335
         TabIndex        =   16
         Top             =   180
         Width           =   1245
      End
      Begin VB.Label label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ward"
         Height          =   195
         Index           =   10
         Left            =   765
         TabIndex        =   15
         Top             =   1710
         Width           =   390
      End
      Begin VB.Label label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Consultant"
         Height          =   195
         Index           =   9
         Left            =   405
         TabIndex        =   14
         Top             =   1410
         Width           =   750
      End
      Begin VB.Label label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Age"
         Height          =   195
         Index           =   8
         Left            =   2925
         TabIndex        =   13
         Top             =   510
         Width           =   285
      End
      Begin VB.Label label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "D.o.B."
         Height          =   195
         Index           =   7
         Left            =   705
         TabIndex        =   12
         Top             =   510
         Width           =   450
      End
      Begin VB.Label label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Sex"
         Height          =   195
         Index           =   6
         Left            =   4095
         TabIndex        =   11
         Top             =   510
         Width           =   270
      End
      Begin VB.Label label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Addr1"
         Height          =   195
         Index           =   5
         Left            =   735
         TabIndex        =   10
         Top             =   810
         Width           =   420
      End
      Begin VB.Label label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Chart Number"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   9
         Top             =   210
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Addr2"
         Height          =   195
         Left            =   735
         TabIndex        =   8
         Top             =   1110
         Width           =   420
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "G. P."
         Height          =   195
         Left            =   795
         TabIndex        =   7
         Top             =   2010
         Width           =   360
      End
   End
   Begin VB.TextBox tName 
      Height          =   285
      Left            =   600
      MaxLength       =   30
      TabIndex        =   2
      Top             =   930
      Width           =   2445
   End
   Begin MSFlexGridLib.MSFlexGrid gName 
      Height          =   5265
      Left            =   150
      TabIndex        =   4
      Top             =   1410
      Width           =   3045
      _ExtentX        =   5371
      _ExtentY        =   9287
      _Version        =   393216
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      GridLines       =   3
      GridLinesFixed  =   3
      FormatString    =   "<Date         |<Name                               "
   End
   Begin VB.Frame Frame1 
      Caption         =   "Between Dates"
      Height          =   675
      Left            =   180
      TabIndex        =   3
      Top             =   150
      Width           =   2865
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   315
         Left            =   1470
         TabIndex        =   1
         Top             =   240
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Format          =   271974401
         CurrentDate     =   37585
      End
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   315
         Left            =   90
         TabIndex        =   0
         Top             =   240
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Format          =   271974401
         CurrentDate     =   37585
      End
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   3855
      Left            =   3510
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   2820
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   6800
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
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
      ScrollBars      =   2
      FormatString    =   "<Run #        |<Time     |<Serum mmol/L "
   End
   Begin VB.Image iFind 
      Height          =   480
      Left            =   3060
      Picture         =   "frmGluByName.frx":0C28
      Top             =   900
      Width           =   480
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      Caption         =   "Name"
      Height          =   195
      Index           =   0
      Left            =   150
      TabIndex        =   5
      Top             =   960
      Width           =   420
   End
End
Attribute VB_Name = "frmGluByName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub FillgName()

          Dim tb As New Recordset
          Dim sn As New Recordset
          Dim tf As Recordset
          Dim sql As String
          Dim n As Long
          Dim NameToFind As String
          Dim Found As Long
          Dim Code As String

10        On Error GoTo FillgName_Error

20        Code = SysOptBioCodeForFastGlucose(0)

30        gName.Rows = 2
40        gName.AddItem ""
50        gName.RemoveItem 1

60        g.Rows = 2
70        g.AddItem ""
80        g.RemoveItem 1

90        pb.Visible = True

100       sql = "SELECT Distinct PatName, RunDate " & _
                "from Demographics WHERE " & _
                "(RunDate between '" & Format$(dtFrom, "dd/mmm/yyyy") & "' " & _
                "and '" & Format$(dtTo, "dd/mmm/yyyy") & "') " & _
                "and PatName like '" & AddTicks(tName) & "%' " & _
                " order by rundate desc"

110       Set tb = New Recordset
120       RecOpenClient 0, tb, sql

130       If tb.EOF Then Exit Sub

140       pb.Min = 0
150       pb = 0
          'pb.max = tb.RecordCount

160       Do While Not tb.EOF
              '  pb = pb + 1
170           gName.AddItem Format$(tb!Rundate, "dd/mm/yy") & vbTab & tb!PatName & ""
180           tb.MoveNext
190       Loop

200       pb = 0
210       pb.Max = gName.Rows

220       For n = gName.Rows - 1 To 2 Step -1
230           pb = pb + 1
240           NameToFind = AddTicks(gName.TextMatrix(n, 1))
250           Found = 0
260           sql = "SELECT * from demographics WHERE " & _
                    "patname = '" & NameToFind & "' " & _
                    "and rundate = '" & Format$(gName.TextMatrix(n, 0), "dd/mmm/yyyy") & "' "
270           Set sn = New Recordset
280           RecOpenServer 0, sn, sql
290           Do While Not sn.EOF
300               sql = "SELECT * from BioResults WHERE " & _
                        "SampleID = '" & sn!SampleID & "' " & _
                        "and RunDate = '" & Format$(gName.TextMatrix(n, 0), "dd/mmm/yyyy") & "' " & _
                        "and (Code = '" & SysOptBioCodeForFastGlucose(0) & "' " & _
                        " or Code = '" & SysOptBioCodeForGlucose1(0) & "'" & _
                        " or Code = '" & SysOptBioCodeForGlucose2(0) & "'" & _
                        " or Code = '" & SysOptBioCodeForGlucose3(0) & "'" & _
                        " or Code = '" & SysOptBioCodeForFastGlucoseP(0) & "' " & _
                        " or Code = '" & SysOptBioCodeForGlucose1P(0) & "'" & _
                        " or Code = '" & SysOptBioCodeForGlucose2P(0) & "'" & _
                        " or Code = '" & SysOptBioCodeForGlucose3P(0) & "')"
310               Set tf = New Recordset
320               RecOpenServer 0, tf, sql

330               If Not tf.EOF Then
340                   Found = Found + 1
350               End If
360               sn.MoveNext
370           Loop
380           If Found < 2 Then
390               gName.RemoveItem n
400           End If
410       Next

420       If gName.Rows > 2 Then
430           gName.RemoveItem 1
440       End If

450       pb.Visible = False

460       Exit Sub

FillgName_Error:

          Dim strES As String
          Dim intEL As Integer

470       intEL = Erl
480       strES = Err.Description
490       LogError "frmGluByName", "FillgName", intEL, strES, sql

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
260       tb!Initiator = UserName
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
340       LogError "frmGluByName", "bPrintGTT_Click", intEL, strES, sql


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
260       tb!Initiator = UserName
270       tb!Ward = strWard
280       tb!Clinician = strClin

290       tb!GP = strGp
300       tb.Update

310       Exit Sub

bPrintSeries_Click_Error:

          Dim strES As String
          Dim intEL As Integer



320       intEL = Erl
330       strES = Err.Description
340       LogError "frmGluByName", "bPrintSeries_Click", intEL, strES, sql


End Sub


Private Sub Form_Load()

10        On Error GoTo Form_Load_Error

20        dtFrom = Format$(Now - 60, "dd/mm/yyyy")
30        dtTo = Format$(Now, "dd/mm/yyyy")

40        Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer



50        intEL = Erl
60        strES = Err.Description
70        LogError "frmGluByName", "Form_Load", intEL, strES


End Sub


Private Sub gName_Click()

          Dim sn As New Recordset
          Dim tb As New Recordset
          Dim sql As String
          Dim s As String
          Dim Found As Long
          Dim Code As String

10        On Error GoTo gName_Click_Error

20        Code = SysOptBioCodeForGlucose(0)

30        g.Rows = 2
40        g.AddItem ""
50        g.RemoveItem 1

60        If gName.MouseRow = 0 Then
70            Exit Sub
80        End If

90        sql = "SELECT * from demographics WHERE " & _
                "patname = '" & gName.TextMatrix(gName.Row, 1) & "' " & _
                "and rundate = '" & Format$(gName.TextMatrix(gName.Row, 0), "dd/mmm/yyyy") & "' " & _
                "order by TimeTaken"
100       Set sn = New Recordset
110       RecOpenClient 0, sn, sql

120       If sn.EOF Then
130           iMsg "No details found", vbExclamation
140           Exit Sub
150       End If

160       lChart = sn!Chart & ""
170       If Not IsNull(sn!Dob) Then
180           lDoB = sn!Dob
190       Else
200           lDoB = ""
210       End If
220       lAge = sn!Age & ""
230       lSex = sn!sex & ""
240       laddr0 = sn!Addr0 & ""
250       laddr1 = sn!Addr1 & ""
260       lclinician = sn!Clinician & ""
270       lward = sn!Ward & ""
280       lgp = sn!GP & ""

290       Do While Not sn.EOF
300           Found = False
310           s = Format$(sn!TimeTaken & "", "hh:mm") & vbTab
320           sql = "SELECT * from BioResults WHERE " & _
                    "SampleID = '" & sn!SampleID & "' " & _
                    "and RunDate = '" & Format$(gName.TextMatrix(gName.Row, 0), "dd/mmm/yyyy") & "' " & _
                    "and (Code = '" & SysOptBioCodeForFastGlucose(0) & "' " & _
                    " or Code = '" & SysOptBioCodeForGlucose1(0) & "'" & _
                    " or Code = '" & SysOptBioCodeForGlucose2(0) & "'" & _
                    " or Code = '" & SysOptBioCodeForGlucose3(0) & "'" & _
                    " or Code = '" & SysOptBioCodeForFastGlucoseP(0) & "' " & _
                    " or Code = '" & SysOptBioCodeForGlucose1P(0) & "'" & _
                    " or Code = '" & SysOptBioCodeForGlucose2P(0) & "'" & _
                    " or Code = '" & SysOptBioCodeForGlucose3P(0) & "')"
330           Set tb = New Recordset
340           RecOpenClient 0, tb, sql
350           If Not tb.EOF Then
360               s = s & Format$(tb!Result, "0.0")
370               Found = True
380           End If
390           If Found Then g.AddItem sn!SampleID & vbTab & s
400           sn.MoveNext
410       Loop

420       If g.Rows > 2 Then
430           g.RemoveItem 1
440       End If



450       Exit Sub

gName_Click_Error:

          Dim strES As String
          Dim intEL As Integer

460       intEL = Erl
470       strES = Err.Description
480       LogError "frmGluByName", "gName_Click", intEL, strES, sql


End Sub

Private Sub iFind_Click()

10        On Error GoTo iFind_Click_Error

20        If Len(Trim$(tName)) > 1 Then
30            FillgName
40        End If

50        Exit Sub

iFind_Click_Error:

          Dim strES As String
          Dim intEL As Integer



60        intEL = Erl
70        strES = Err.Description
80        LogError "frmGluByName", "iFind_Click", intEL, strES


End Sub

Private Sub tName_Change()

10        On Error GoTo tName_Change_Error

20        If Len(Trim$(tName)) > 4 Then
30            FillgName
40        End If

50        Exit Sub

tName_Change_Error:

          Dim strES As String
          Dim intEL As Integer



60        intEL = Erl
70        strES = Err.Description
80        LogError "frmGluByName", "tName_Change", intEL, strES


End Sub


