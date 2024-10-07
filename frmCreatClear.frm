VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmCreatClear 
   Caption         =   "NetAcquire - Creatinine Clearance"
   ClientHeight    =   6015
   ClientLeft      =   840
   ClientTop       =   1965
   ClientWidth     =   9420
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
   Icon            =   "frmCreatClear.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6015
   ScaleWidth      =   9420
   Begin VB.CommandButton cmdSave 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Save"
      Enabled         =   0   'False
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
      Left            =   495
      Picture         =   "frmCreatClear.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   5220
      Width           =   1410
   End
   Begin VB.CommandButton bRefresh 
      Caption         =   "Refresh Numbers"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   3360
      Picture         =   "frmCreatClear.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   90
      Visible         =   0   'False
      Width           =   1500
   End
   Begin MSComCtl2.DTPicker dtTo 
      Height          =   315
      Left            =   1800
      TabIndex        =   1
      Top             =   300
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
      _Version        =   393216
      Format          =   59834369
      CurrentDate     =   37722
   End
   Begin MSComCtl2.DTPicker dtFrom 
      Height          =   315
      Left            =   210
      TabIndex        =   0
      Top             =   300
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
      _Version        =   393216
      Format          =   59834369
      CurrentDate     =   37505
   End
   Begin VB.PictureBox Panel3D1 
      Height          =   2025
      Index           =   1
      Left            =   4500
      ScaleHeight     =   1965
      ScaleWidth      =   4575
      TabIndex        =   27
      Top             =   3090
      Width           =   4635
      Begin VB.CommandButton bprinturine 
         Appearance      =   0  'Flat
         Height          =   600
         Left            =   3420
         Picture         =   "frmCreatClear.frx":0A56
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   855
         Width           =   825
      End
      Begin VB.Label lurinedate 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   2520
         TabIndex        =   46
         Top             =   90
         Width           =   1995
      End
      Begin VB.Label lcomment 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   44
         Top             =   1470
         Width           =   3525
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Name"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   33
         Top             =   570
         Width           =   495
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Chart"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   210
         TabIndex        =   32
         Top             =   870
         Width           =   465
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "DoB"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   300
         TabIndex        =   31
         Top             =   1170
         Width           =   375
      End
      Begin VB.Label ldob 
         BorderStyle     =   1  'Fixed Single
         Height          =   225
         Index           =   1
         Left            =   720
         TabIndex        =   30
         Top             =   1170
         Width           =   1695
      End
      Begin VB.Label lchart 
         BorderStyle     =   1  'Fixed Single
         Height          =   225
         Index           =   1
         Left            =   720
         TabIndex        =   29
         Top             =   870
         Width           =   1695
      End
      Begin VB.Label lname 
         BorderStyle     =   1  'Fixed Single
         Height          =   225
         Index           =   1
         Left            =   720
         TabIndex        =   28
         Top             =   570
         Width           =   3525
      End
   End
   Begin VB.PictureBox Panel3D3 
      Height          =   2025
      Left            =   210
      ScaleHeight     =   1965
      ScaleWidth      =   4155
      TabIndex        =   23
      Top             =   3090
      Width           =   4215
      Begin VB.Label lblCreatClear 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3120
         TabIndex        =   40
         Top             =   1500
         Width           =   660
      End
      Begin VB.Label lblUrnProt24 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3120
         TabIndex        =   39
         Top             =   1230
         Width           =   840
      End
      Begin VB.Label lblUrnProt 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3120
         TabIndex        =   38
         Top             =   960
         Width           =   525
      End
      Begin VB.Label lblUrnCreat 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3120
         TabIndex        =   37
         Top             =   690
         Width           =   615
      End
      Begin VB.Label lblSerCreat 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3120
         TabIndex        =   36
         Top             =   420
         Width           =   660
      End
      Begin VB.Label lupc 
         BorderStyle     =   1  'Fixed Single
         Height          =   225
         Left            =   2070
         TabIndex        =   35
         Top             =   1200
         Width           =   1020
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Urinary Protein"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   780
         TabIndex        =   34
         Top             =   1200
         Width           =   1275
      End
      Begin VB.Label lcc 
         BorderStyle     =   1  'Fixed Single
         Height          =   225
         Left            =   2070
         TabIndex        =   7
         Top             =   1470
         Width           =   1020
      End
      Begin VB.Label lup 
         BorderStyle     =   1  'Fixed Single
         Height          =   225
         Left            =   2070
         TabIndex        =   8
         Top             =   945
         Width           =   1020
      End
      Begin VB.Label luc 
         BorderStyle     =   1  'Fixed Single
         Height          =   225
         Left            =   2070
         TabIndex        =   9
         Top             =   660
         Width           =   1020
      End
      Begin VB.Label lsc 
         BorderStyle     =   1  'Fixed Single
         Height          =   225
         Left            =   2070
         TabIndex        =   10
         Top             =   390
         Width           =   1020
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Creatinine Clearance"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   270
         TabIndex        =   11
         Top             =   1470
         Width           =   1785
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Urinary Protein"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   780
         TabIndex        =   12
         Top             =   930
         Width           =   1275
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Urine Creatinine"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   690
         TabIndex        =   13
         Top             =   660
         Width           =   1380
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Serum Creatinine"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   600
         TabIndex        =   14
         Top             =   390
         Width           =   1455
      End
   End
   Begin VB.PictureBox Panel3D2 
      Height          =   2025
      Left            =   210
      ScaleHeight     =   1965
      ScaleWidth      =   4155
      TabIndex        =   22
      Top             =   945
      Width           =   4215
      Begin VB.ComboBox cUrine 
         Height          =   315
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   810
         Width           =   1515
      End
      Begin VB.ComboBox cSerum 
         Height          =   315
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   450
         Width           =   1515
      End
      Begin VB.TextBox tvolume 
         Height          =   285
         Left            =   2400
         TabIndex        =   5
         Top             =   1170
         Width           =   1245
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "mL"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3660
         TabIndex        =   41
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Total Urinary Volume"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   600
         TabIndex        =   26
         Top             =   1200
         Width           =   1785
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Urine Sample Number"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   540
         TabIndex        =   25
         Top             =   840
         Width           =   1845
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Serum Sample Number"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   480
         TabIndex        =   24
         Top             =   480
         Width           =   1920
      End
   End
   Begin VB.PictureBox Panel3D1 
      Height          =   2025
      Index           =   0
      Left            =   4500
      ScaleHeight     =   1965
      ScaleWidth      =   4575
      TabIndex        =   21
      Top             =   960
      Width           =   4635
      Begin VB.CommandButton bprintserum 
         Height          =   555
         Left            =   3420
         Picture         =   "frmCreatClear.frx":0D60
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   840
         Width           =   825
      End
      Begin VB.Label lserumdate 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2580
         TabIndex        =   45
         Top             =   90
         Width           =   1935
      End
      Begin VB.Label lcomment 
         BorderStyle     =   1  'Fixed Single
         Height          =   225
         Index           =   0
         Left            =   720
         TabIndex        =   43
         Top             =   1470
         Width           =   3525
      End
      Begin VB.Label lname 
         BorderStyle     =   1  'Fixed Single
         Height          =   225
         Index           =   0
         Left            =   720
         TabIndex        =   15
         Top             =   570
         Width           =   3525
      End
      Begin VB.Label lchart 
         BorderStyle     =   1  'Fixed Single
         Height          =   225
         Index           =   0
         Left            =   720
         TabIndex        =   16
         Top             =   870
         Width           =   1695
      End
      Begin VB.Label ldob 
         BorderStyle     =   1  'Fixed Single
         Height          =   225
         Index           =   0
         Left            =   720
         TabIndex        =   17
         Top             =   1170
         Width           =   1695
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "DoB"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   18
         Top             =   1170
         Width           =   375
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Chart"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   210
         TabIndex        =   19
         Top             =   870
         Width           =   465
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Name"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   20
         Top             =   570
         Width           =   495
      End
   End
   Begin VB.CommandButton bcancel 
      Appearance      =   0  'Flat
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
      Height          =   705
      Left            =   7695
      Picture         =   "frmCreatClear.frx":106A
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5220
      Width           =   1410
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "Run Dates Between"
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
      Left            =   960
      TabIndex        =   47
      Top             =   60
      Width           =   1440
   End
End
Attribute VB_Name = "frmCreatClear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Activated As Boolean

Private Sub bcancel_Click()

10        Unload Me

End Sub

Private Sub bPrintSerum_Click()

          Dim tb As New Recordset
          Dim sql As String
          Dim strWard As String
          Dim strClin As String
          Dim strGp As String

10        On Error GoTo bPrintSerum_Click_Error

20        SaveCreat

30        strWard = ""
40        strGp = ""
50        strClin = ""

60        sql = "SELECT * FROM Demographics WHERE " & _
                "SampleID = '" & cSerum & "'"
70        Set tb = New Recordset
80        RecOpenServer 0, tb, sql
90        If Not tb.EOF Then
100           strWard = tb!Ward & ""
110           strClin = tb!Clinician & ""
120           strGp = tb!GP & ""
130       End If

140       sql = "SELECT * FROM PrintPending WHERE " & _
                "Department = 'T' " & _
                "AND SampleID = '" & cSerum & "'"
150       Set tb = New Recordset
160       RecOpenClient 0, tb, sql
170       If tb.EOF Then
180           tb.AddNew
190       End If
200       tb!SampleID = cSerum
210       tb!Initiator = Username
220       tb!Ward = strWard
230       tb!Clinician = strClin
240       tb!GP = strGp
250       tb!pTime = Now
260       tb!Department = "T"
270       tb!Initiator = Username
280       tb.Update

290       Exit Sub

bPrintSerum_Click_Error:

          Dim strES As String
          Dim intEL As Integer

300       intEL = Erl
310       strES = Err.Description
320       LogError "frmCreatClear", "bPrintSerum_Click", intEL, strES

End Sub

Private Sub bprinturine_Click()

          Dim tb As New Recordset
          Dim sql As String
          Dim strWard As String
          Dim strClin As String
          Dim strGp As String

10        On Error GoTo bprinturine_Click_Error

20        SaveCreat

30        strWard = ""
40        strGp = ""
50        strClin = ""

60        sql = "SELECT * FROM Demographics WHERE " & _
                "SampleID = '" & cUrine & "'"

70        Set tb = New Recordset
80        RecOpenServer 0, tb, sql
90        If Not tb.EOF Then
100           strWard = tb!Ward & ""
110           strClin = tb!Clinician & ""
120           strGp = tb!GP & ""
130       End If

140       sql = "SELECT * FROM PrintPending WHERE " & _
                "Department = 'R' " & _
                "AND SampleID = '" & cUrine & "'"
150       Set tb = New Recordset
160       RecOpenClient 0, tb, sql
170       If tb.EOF Then
180           tb.AddNew
190       End If
200       tb!SampleID = cUrine
210       tb!Department = "R"
220       tb!Initiator = Username
230       tb!Ward = strWard
240       tb!Clinician = strClin
250       tb!GP = strGp
260       tb!pTime = Now
270       tb.Update

280       Exit Sub

bprinturine_Click_Error:

          Dim strES As String
          Dim intEL As Integer

290       intEL = Erl
300       strES = Err.Description
310       LogError "frmCreatClear", "bprinturine_Click", intEL, strES

End Sub

Private Sub bRefresh_Click()

10        On Error GoTo bRefresh_Click_Error

20        SuggestNumbers

30        bRefresh.Visible = False

40        Exit Sub

bRefresh_Click_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmCreatClear", "bRefresh_Click", intEL, strES

End Sub

Private Sub Calculate()

10        On Error GoTo Calculate_Error

20        If Val(lup) <> 0 And Val(tvolume) <> 0 Then
30            lupc = Format$(Val(lup) * Val(tvolume) / 1000, "0.000")
40        End If

50        If Val(luc) <> 0 And Val(lsc) <> 0 And Val(tvolume) <> 0 Then
60            lcc = (((Val(luc) * SysOptUCVal(0)) * Val(tvolume)) / ((Val(lsc) * 1440)))
70            lcc = (Format$(lcc, "##0"))
80        End If

90        cmdSave.Enabled = True

100       Exit Sub

Calculate_Error:

          Dim strES As String
          Dim intEL As Integer

110       intEL = Erl
120       strES = Err.Description
130       LogError "frmCreatClear", "calculate", intEL, strES

End Sub

Private Sub cmdSave_Click()

          Dim sql As String
          Dim tb As Recordset

10        On Error GoTo cmdSave_Click_Error

20        If SysOptBioCodeForCreatClear(0) <> "" And Val(lcc) <> 0 Then
30            sql = "SELECT * FROM BioResults WHERE " & _
                    "SampleID = " & cUrine & " " & _
                    "AND Code = '" & SysOptBioCodeForCreatClear(0) & "'"
40            Set tb = New Recordset
50            RecOpenClient 0, tb, sql
60            If tb.EOF Then tb.AddNew
70            tb!SampleID = cUrine
80            tb!Rundate = Format$(Now, "dd/mmm/yyyy")
90            tb!Code = SysOptBioCodeForCreatClear(0)
100           tb!Result = lcc
110           tb!RunTime = Format$(Now, "dd/mmm/yyyy hh:mm")
120           tb!Units = lblCreatClear.Caption
130           tb!SampleType = "U"
140           tb!Valid = 0
150           tb!Printed = 0
160           tb.Update
170       End If
180       If SysOptBioCodeFor24UProt(0) <> "" And Val(lupc) <> 0 Then
190           sql = "select * from bioresults where sampleid = " & cUrine & " and code = '" & SysOptBioCodeFor24UProt(0) & "'"
200           Set tb = New Recordset
210           RecOpenClient 0, tb, sql
220           If tb.EOF Then tb.AddNew
230           tb!SampleID = cUrine
240           tb!Rundate = Format$(Now, "dd/mmm/yyyy")
250           tb!Code = SysOptBioCodeFor24UProt(0)
260           tb!Result = lupc
270           tb!RunTime = Format$(Now, "dd/mmm/yyyy hh:mm")
280           tb!Units = lblUrnProt24.Caption
290           tb!SampleType = "U"
300           tb!Valid = 0
310           tb!Printed = 0
320           tb.Update
330       End If

340       cmdSave.Enabled = False

350       Exit Sub

cmdSave_Click_Error:

          Dim strES As String
          Dim intEL As Integer

360       intEL = Erl
370       strES = Err.Description
380       LogError "frmCreatClear", "cmdsave_Click", intEL, strES, sql

End Sub

Private Sub cSerum_Click()

          Dim tb As New Recordset
          Dim sql As String

10        On Error GoTo cSerum_Click_Error

20        If Len(cSerum) = 0 Then Exit Sub

30        sql = "SELECT * FROM Demographics WHERE " & _
                "SampleID = '" & cSerum & "'"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        If tb.EOF Then
70            lserumdate = ""
80            lName(0) = ""
90            lChart(0) = ""
100           lDoB(0) = ""
110           lcomment(0) = ""
120       Else
130           tb.MoveLast
140           lserumdate = tb!Rundate
150           lChart(0) = tb!Chart & ""
160           lName(0) = tb!PatName & ""
170           lDoB(0) = tb!Dob & ""
              '  lcomment(0) = tb!biocomment0 & ""
180       End If

190       sql = "SELECT * from bioresults WHERE " & _
                "SampleID = '" & cSerum & "' " & _
                "and Code = '" & SysOptBioCodeForCreat(0) & "'"
200       Set tb = New Recordset
210       RecOpenServer 0, tb, sql
220       If Not tb.EOF Then
230           If Not IsNull(tb!Result) Then
240               lsc = Format$(tb!Result, "####.0")
250           Else
260               lsc = ""
270           End If
280       Else
290           lsc = ""
300       End If

310       If lName(0) <> lName(1) Then

320           cUrine.Clear

330           sql = "SELECT DISTINCT BioResults.SampleId " & _
                    "FROM BioResults INNER JOIN " & _
                    "Demographics ON BioResults.SampleId = Demographics.SampleID " & _
                    "WHERE (BioResults.RunTime BETWEEN '" & Format$(dtFrom, "dd/mmm/yyyy") & " 00:00:00' " & _
                    "AND '" & Format$(dtTo, "dd/mmm/yyyy") & " 23:59:59') " & _
                    "AND (BioResults.Code = '" & SysOptBioCodeForUProt(0) & "' OR " & _
                    "BioResults.Code = '" & SysOptBioCodeForUCreat(0) & "') AND " & _
                    "(Demographics.PatName = '" & lName(0) & "') " & _
                    "AND (Demographics.DoB = '" & Format(lDoB(0), "dd/MMM/yyyy") & "')"

340           Set tb = New Recordset
350           RecOpenServer 0, tb, sql
360           Do While Not tb.EOF
370               cUrine.AddItem tb!SampleID
380               tb.MoveNext    '
390           Loop

400       End If
410       Calculate

420       Exit Sub

cSerum_Click_Error:

          Dim strES As String
          Dim intEL As Integer

430       intEL = Erl
440       strES = Err.Description
450       LogError "frmCreatClear", "cSerum_Click", intEL, strES, sql

End Sub

Private Sub curine_Click()

          Dim tb As New Recordset
          Dim sql As String

10        On Error GoTo curine_Click_Error

20        If Len(cUrine) = 0 Then Exit Sub

30        sql = "SELECT * FROM Demographics WHERE " & _
                "SampleID = '" & cUrine & "'"

40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        If tb.EOF Then
70            lurinedate = ""
80            lName(1) = ""
90            lChart(1) = ""
100           lDoB(1) = ""
110           lcomment(1) = ""
120       Else
130           tb.MoveLast
140           lurinedate = tb!Rundate
150           lChart(1) = tb!Chart & ""
160           lName(1) = tb!PatName & ""
170           lDoB(1) = tb!Dob & ""
              '  lcomment(1) = tb!biocomment0 & ""
180       End If

190       sql = "SELECT * from bioresults WHERE " & _
                "SampleID = '" & cUrine & "' " & _
                "and Code = '" & SysOptBioCodeForUCreat(0) & "'"
200       Set tb = New Recordset
210       RecOpenServer 0, tb, sql
220       If Not tb.EOF Then
230           If Not IsNull(tb!Result) Then
240               luc = Format$(tb!Result, "#0.0")
250           Else
260               luc = ""
270           End If
280       Else
290           luc = ""
300       End If

310       sql = "SELECT * from bioresults WHERE " & _
                "SampleID = '" & cUrine & "' " & _
                "and code = '" & SysOptBioCodeForUProt(0) & "'"
320       Set tb = New Recordset
330       RecOpenServer 0, tb, sql
340       If Not tb.EOF Then
350           If Not IsNull(tb!Result) Then
360               lup = Format$(tb!Result, "#0.00")
370           Else
380               lup = ""
390           End If
400       Else
410           lup = ""
420       End If

430       If lName(0) <> lName(1) Then

440           cSerum.Clear

450           sql = "SELECT DISTINCT BioResults.SampleId " & _
                    "FROM BioResults INNER JOIN " & _
                    "Demographics ON BioResults.SampleId = Demographics.SampleID " & _
                    "WHERE (BioResults.RunTime BETWEEN '" & Format$(dtFrom, "dd/mmm/yyyy") & " 00:00:00' " & _
                    "AND '" & Format$(dtTo, "dd/mmm/yyyy") & " 23:59:59') " & _
                    "AND (BioResults.Code = '" & SysOptBioCodeForCreat(0) & "') AND " & _
                    "(Demographics.PatName = '" & lName(1) & "') " & _
                    "AND (Demographics.DoB = '" & Format(lDoB(1), "dd/MMM/yyyy") & "')"

460           Set tb = New Recordset
470           RecOpenServer 0, tb, sql
480           Do While Not tb.EOF
490               cSerum.AddItem tb!SampleID
500               tb.MoveNext    '
510           Loop

520       End If

530       Calculate

540       Exit Sub

curine_Click_Error:

          Dim strES As String
          Dim intEL As Integer

550       intEL = Erl
560       strES = Err.Description
570       LogError "frmCreatClear", "curine_Click", intEL, strES, sql

End Sub

Private Sub dtFrom_CloseUp()

10        bRefresh.Visible = True

End Sub

Private Sub dtTo_CloseUp()

10        bRefresh.Visible = True

End Sub

Private Sub Form_Activate()

10        On Error GoTo Form_Activate_Error

20        If Activated Then Exit Sub

30        SuggestNumbers

40        Activated = True

50        Exit Sub

Form_Activate_Error:

          Dim strES As String
          Dim intEL As Integer

60        intEL = Erl
70        strES = Err.Description
80        LogError "frmCreatClear", "Form_Activate", intEL, strES

End Sub

Private Sub Form_Load()
          Dim sql As String
          Dim tb As New ADODB.Recordset

10        On Error GoTo Form_Load_Error

20        dtFrom = Format$(Now, "dd/mm/yyyy")
30        dtTo = Format$(Now, "dd/mm/yyyy")

40        Activated = False

50        sql = "Select * from biotestdefinitions"
60        Set tb = New Recordset
70        RecOpenServer 0, tb, sql
80        Do While Not tb.EOF
90            If SysOptBioCodeForCreat(0) = tb!Code Then
100               lblSerCreat.Caption = tb!Units
110           End If
120           If SysOptBioCodeForUCreat(0) = tb!Code Then
130               lblUrnCreat.Caption = tb!Units
140           End If
150           If SysOptBioCodeForUProt(0) = tb!Code Then
160               lblUrnProt.Caption = tb!Units
170           End If
180           If SysOptBioCodeForCreatClear(0) = tb!Code Then
190               lblCreatClear.Caption = tb!Units
200           End If
210           If SysOptBioCodeFor24UProt(0) = tb!Code Then
220               lblUrnProt24.Caption = tb!Units
230           End If
240           tb.MoveNext
250       Loop

260       Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

270       intEL = Erl
280       strES = Err.Description
290       LogError "frmCreatClear", "Form_Load", intEL, strES, sql

End Sub

Private Sub Form_Unload(Cancel As Integer)

10        Activated = False

End Sub

Private Sub SaveCreat()

          Dim tb As New Recordset
          Dim sql As String

10        On Error GoTo SaveCreat_Error

20        sql = "SELECT * from Creatinine WHERE " & _
                "SerumNumber = '" & cSerum & "' " & _
                "and UrineNumber = '" & cUrine & "'"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql

50        If tb.EOF Then
60            tb.AddNew
70        End If
80        tb!serumnumber = cSerum
90        tb!urinenumber = cUrine
100       tb!urinevolume = tvolume
110       tb!serumcreat = lsc
120       tb!urinecreat = luc
130       tb!urineprol = lup
140       tb!urinepro24hr = lupc
150       tb!ccl = lcc
160       tb!Name = lName(1)
170       tb!Chart = lChart(1)
180       If IsDate(lDoB(1)) Then tb!Dob = lDoB(1)
190       tb!Comment = Left$(lcomment(1), 30)
200       tb.Update

210       Exit Sub

SaveCreat_Error:

          Dim strES As String
          Dim intEL As Integer

220       intEL = Erl
230       strES = Err.Description
240       LogError "frmCreatClear", "SaveCreat", intEL, strES, sql

End Sub

Private Sub SuggestNumbers()

          Dim tb As New Recordset
          Dim sql As String
          Dim SCreaCode As String
          Dim UCreaCode As String
          Dim UProCode As String

10        On Error GoTo SuggestNumbers_Error

20        SCreaCode = SysOptBioCodeForCreat(0)
30        UCreaCode = SysOptBioCodeForUCreat(0)
40        UProCode = SysOptBioCodeForUProt(0)

50        cSerum.Clear
60        cUrine.Clear

70        If Abs(DateDiff("d", dtFrom, dtTo)) > 6 Then
80            iMsg "Maximum Seven Days!", vbExclamation
90            Exit Sub
100       End If

110       sql = "SELECT distinct SampleID from bioresults WHERE " & _
                "Runtime between '" & Format$(dtFrom, "dd/mmm/yyyy") & "' " & _
                "and '" & Format$(dtTo, "dd/mmm/yyyy") & " 23:59:59' " & _
                "and code = '" & SCreaCode & "'"
120       Set tb = New Recordset
130       RecOpenServer 0, tb, sql
140       Do While Not tb.EOF
150           cSerum.AddItem tb!SampleID & ""
160           tb.MoveNext
170       Loop

180       sql = "SELECT distinct SampleID from bioresults WHERE " & _
                "Runtime between '" & Format$(dtFrom, "dd/mmm/yyyy") & "' " & _
                "and '" & Format$(dtTo, "dd/mmm/yyyy") & " 23:59:59' " & _
                "and (code = '" & UCreaCode & "' or code = '" & UProCode & "')"
190       Set tb = New Recordset
200       RecOpenServer 0, tb, sql
210       Do While Not tb.EOF
220           cUrine.AddItem tb!SampleID & ""
230           tb.MoveNext
240       Loop

250       Exit Sub

SuggestNumbers_Error:

          Dim strES As String
          Dim intEL As Integer

260       intEL = Erl
270       strES = Err.Description
280       LogError "frmCreatClear", "SuggestNumbers", intEL, strES, sql

End Sub

Private Sub tVolume_LostFocus()

10        Calculate

End Sub
