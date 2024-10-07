VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form frm24hrUrine 
   Caption         =   "NetAcquire - 24 Hr Urine Excretion"
   ClientHeight    =   6465
   ClientLeft      =   2445
   ClientTop       =   600
   ClientWidth     =   5430
   Icon            =   "frm24hrUrine.frx":0000
   LinkTopic       =   "Form1"
   Palette         =   "frm24hrUrine.frx":030A
   PaletteMode     =   2  'Custom
   ScaleHeight     =   6465
   ScaleWidth      =   5430
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   750
      Left            =   4057
      Picture         =   "frm24hrUrine.frx":521C
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   3705
      Width           =   1110
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   840
      Left            =   4035
      Picture         =   "frm24hrUrine.frx":5526
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   5415
      Width           =   1155
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   750
      Left            =   4057
      Picture         =   "frm24hrUrine.frx":5830
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   4560
      Width           =   1110
   End
   Begin VB.Frame Frame4 
      Caption         =   "Units"
      Height          =   825
      Left            =   3990
      TabIndex        =   40
      Top             =   2700
      Visible         =   0   'False
      Width           =   1245
      Begin VB.TextBox tUnits 
         Height          =   285
         Left            =   600
         MaxLength       =   2
         TabIndex        =   41
         Text            =   "24"
         Top             =   180
         Width           =   300
      End
      Begin ComCtl2.UpDown udHours 
         Height          =   195
         Left            =   150
         TabIndex        =   44
         Top             =   510
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   344
         _Version        =   327681
         Value           =   24
         BuddyControl    =   "tUnits"
         BuddyDispid     =   196613
         OrigLeft        =   3180
         OrigTop         =   3600
         OrigRight       =   3420
         OrigBottom      =   4365
         Max             =   24
         Min             =   1
         Orientation     =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Hrs"
         Height          =   195
         Left            =   930
         TabIndex        =   43
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "mmol/"
         Height          =   195
         Left            =   150
         TabIndex        =   42
         Top             =   240
         Width           =   435
      End
   End
   Begin VB.Frame Frame3 
      Height          =   4425
      Left            =   180
      TabIndex        =   10
      Top             =   1800
      Width           =   3405
      Begin VB.TextBox tper24 
         BackColor       =   &H80000011&
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   9
         Left            =   2190
         Locked          =   -1  'True
         TabIndex        =   52
         Top             =   3930
         Width           =   795
      End
      Begin VB.TextBox tmmol 
         Height          =   285
         Index           =   9
         Left            =   1140
         TabIndex        =   51
         Top             =   3930
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.TextBox tper24 
         BackColor       =   &H80000011&
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   8
         Left            =   2190
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   3300
         Width           =   795
      End
      Begin VB.TextBox tper24 
         BackColor       =   &H80000011&
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   7
         Left            =   2190
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   2940
         Width           =   795
      End
      Begin VB.TextBox tper24 
         BackColor       =   &H80000011&
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   6
         Left            =   2190
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   2580
         Width           =   795
      End
      Begin VB.TextBox tper24 
         BackColor       =   &H80000011&
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   5
         Left            =   2190
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   2220
         Width           =   795
      End
      Begin VB.TextBox tper24 
         BackColor       =   &H80000011&
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   4
         Left            =   2190
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   1860
         Width           =   795
      End
      Begin VB.TextBox tper24 
         BackColor       =   &H80000011&
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   3
         Left            =   2190
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   1500
         Width           =   795
      End
      Begin VB.TextBox tper24 
         BackColor       =   &H80000011&
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   2
         Left            =   2190
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   1140
         Width           =   795
      End
      Begin VB.TextBox tper24 
         BackColor       =   &H80000011&
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   1
         Left            =   2190
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   780
         Width           =   795
      End
      Begin VB.TextBox tper24 
         BackColor       =   &H80000011&
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   2190
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   420
         Width           =   795
      End
      Begin VB.TextBox tmmol 
         Height          =   285
         Index           =   8
         Left            =   1140
         TabIndex        =   28
         Top             =   3300
         Width           =   795
      End
      Begin VB.TextBox tmmol 
         Height          =   285
         Index           =   7
         Left            =   1140
         TabIndex        =   27
         Top             =   2940
         Width           =   795
      End
      Begin VB.TextBox tmmol 
         Height          =   285
         Index           =   6
         Left            =   1140
         TabIndex        =   26
         Top             =   2580
         Width           =   795
      End
      Begin VB.TextBox tmmol 
         Height          =   285
         Index           =   5
         Left            =   1140
         TabIndex        =   25
         Top             =   2220
         Width           =   795
      End
      Begin VB.TextBox tmmol 
         Height          =   285
         Index           =   4
         Left            =   1140
         TabIndex        =   24
         Top             =   1860
         Width           =   795
      End
      Begin VB.TextBox tmmol 
         Height          =   285
         Index           =   3
         Left            =   1140
         TabIndex        =   23
         Top             =   1500
         Width           =   795
      End
      Begin VB.TextBox tmmol 
         Height          =   285
         Index           =   2
         Left            =   1140
         TabIndex        =   22
         Top             =   1140
         Width           =   795
      End
      Begin VB.TextBox tmmol 
         Height          =   285
         Index           =   1
         Left            =   1140
         TabIndex        =   21
         Top             =   780
         Width           =   795
      End
      Begin VB.TextBox tmmol 
         Height          =   285
         Index           =   0
         Left            =   1140
         TabIndex        =   20
         Top             =   420
         Width           =   795
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Urine"
         Height          =   195
         Left            =   1350
         TabIndex        =   55
         Top             =   180
         Width           =   375
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "24 Hr"
         Height          =   195
         Left            =   2392
         TabIndex        =   54
         Top             =   180
         Width           =   390
      End
      Begin VB.Label l 
         AutoSize        =   -1  'True
         Caption         =   "Nitrogen"
         Height          =   195
         Index           =   9
         Left            =   450
         TabIndex        =   50
         Top             =   3990
         Width           =   600
      End
      Begin VB.Label l 
         AutoSize        =   -1  'True
         Caption         =   "Potassium"
         Height          =   195
         Index           =   3
         Left            =   330
         TabIndex        =   19
         Top             =   1530
         Width           =   720
      End
      Begin VB.Label l 
         AutoSize        =   -1  'True
         Caption         =   "Sodium"
         Height          =   195
         Index           =   2
         Left            =   525
         TabIndex        =   18
         Top             =   1170
         Width           =   525
      End
      Begin VB.Label l 
         AutoSize        =   -1  'True
         Caption         =   "Urea"
         Height          =   195
         Index           =   1
         Left            =   705
         TabIndex        =   17
         Top             =   825
         Width           =   345
      End
      Begin VB.Label l 
         AutoSize        =   -1  'True
         Caption         =   "Calcium"
         Height          =   195
         Index           =   5
         Left            =   495
         TabIndex        =   16
         Top             =   2280
         Width           =   555
      End
      Begin VB.Label l 
         AutoSize        =   -1  'True
         Caption         =   "Phosphorus"
         Height          =   195
         Index           =   6
         Left            =   210
         TabIndex        =   15
         Top             =   2610
         Width           =   840
      End
      Begin VB.Label l 
         AutoSize        =   -1  'True
         Caption         =   "T. Prot"
         Height          =   195
         Index           =   8
         Left            =   570
         TabIndex        =   14
         Top             =   3360
         Width           =   480
      End
      Begin VB.Label l 
         AutoSize        =   -1  'True
         Caption         =   "Magnesium"
         Height          =   195
         Index           =   7
         Left            =   240
         TabIndex        =   13
         Top             =   2970
         Width           =   810
      End
      Begin VB.Label l 
         AutoSize        =   -1  'True
         Caption         =   "Chloride"
         Height          =   195
         Index           =   4
         Left            =   480
         TabIndex        =   12
         Top             =   1920
         Width           =   570
      End
      Begin VB.Label l 
         AutoSize        =   -1  'True
         Caption         =   "Creatinine"
         Height          =   195
         Index           =   0
         Left            =   345
         TabIndex        =   11
         Top             =   480
         Width           =   705
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Volume"
      Height          =   675
      Left            =   3990
      TabIndex        =   9
      Top             =   1800
      Width           =   1245
      Begin VB.TextBox txtVolume 
         Height          =   285
         Left            =   105
         TabIndex        =   38
         Top             =   270
         Width           =   795
      End
      Begin VB.Label Label16 
         Caption         =   "ml"
         Height          =   195
         Left            =   930
         TabIndex        =   39
         Top             =   330
         Width           =   165
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1545
      Left            =   150
      TabIndex        =   0
      Top             =   120
      Width           =   5085
      Begin VB.TextBox tDoB 
         Height          =   285
         Left            =   3690
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   180
         Width           =   1245
      End
      Begin VB.TextBox tChart 
         Height          =   285
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   840
         Width           =   1245
      End
      Begin VB.TextBox tName 
         Height          =   285
         Left            =   900
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   6
         Top             =   510
         Width           =   4035
      End
      Begin VB.TextBox txtSampleID 
         Height          =   285
         Left            =   900
         TabIndex        =   2
         Top             =   180
         Width           =   1245
      End
      Begin VB.Label lRunTime 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3720
         TabIndex        =   49
         Top             =   1140
         Width           =   1215
      End
      Begin VB.Label lRunDate 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3720
         TabIndex        =   48
         Top             =   870
         Width           =   1215
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Run Date/Time"
         Height          =   195
         Left            =   2550
         TabIndex        =   47
         Top             =   900
         Width           =   1110
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         Height          =   195
         Left            =   420
         TabIndex        =   5
         Top             =   540
         Width           =   420
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "DoB"
         Height          =   195
         Left            =   3330
         TabIndex        =   4
         Top             =   240
         Width           =   315
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Chart"
         Height          =   195
         Left            =   480
         TabIndex        =   3
         Top             =   870
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Sample ID"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   735
      End
   End
End
Attribute VB_Name = "frm24hrUrine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strmMol(0 To 9) As String
Dim strPer24(0 To 9) As String
Dim strDP(0 To 9) As String

Private Sub Calculate()

          Dim n As Long

10        On Error GoTo Calculate_Error

20        For n = 0 To 8
30            tper24(n) = ""
40        Next

50        If Val(txtVolume) = 0 Then Exit Sub
60        If Val(tUnits) = 0 Then Exit Sub

70        For n = 0 To 8
80            If Val(tmmol(n)) > 0 Then
90                strPer24(n) = (Val(txtVolume) / 1000) * Val(tmmol(n))
100           End If
110       Next

120       If Val(strPer24(1)) > 0 Then
130           strPer24(9) = Format(Val(strPer24(1)) * 0.028, "0.0")
140       End If

150       For n = 0 To 8
160           tper24(n) = Format$(strPer24(n), strDP(n))
              'tmmol(n) = strmMol(n)
170       Next

180       Exit Sub

Calculate_Error:

          Dim strES As String
          Dim intEL As Integer

190       intEL = Erl
200       strES = Err.Description
210       LogError "frm24hrUrine", "Calculate", intEL, strES

End Sub

Private Sub cmdCancel_Click()

10        Unload Me

End Sub

Private Sub cmdPrint_Click()

          Dim sql As String
          Dim tb As New Recordset

10        On Error GoTo cmdPrint_Click_Error

20        sql = "SELECT * FROM PrintPending WHERE " & _
                "Department = 'B' " & _
                "AND SampleID = '" & txtSampleID & "'"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        If tb.EOF Then
60            tb.AddNew
70        End If
80        tb!SampleID = txtSampleID
90        tb!Department = "B"
100       tb!Initiator = UserName
110       tb!UsePrinter = ""
120       tb.Update

130       Exit Sub

cmdPrint_Click_Error:

          Dim strES As String
          Dim intEL As Integer

140       intEL = Erl
150       strES = Err.Description
160       LogError "frm24hrUrine", "cmdPrint_Click", intEL, strES, sql

End Sub

Private Sub cmdSave_Click()

          Dim tb As Recordset
          Dim sql As String
          Dim n As Long
          Dim BioCode As String

10        On Error GoTo cmdSave_Click_Error

          'SAVE URINE RESULTS IF ENTERED MANUALLY
20        For n = 0 To 8
30            If tmmol(n) <> "" And tmmol(n).BackColor = &H80000005 Then
40                BioCode = GetCodeUrine(n)
50                If BioCode <> "" Then
60                    sql = "SELECT * FROM BioResults WHERE " & _
                            "SampleID = '" & txtSampleID & "' " & _
                            "AND Code = '" & BioCode & "'"
70                    Set tb = New Recordset
80                    RecOpenServer 0, tb, sql
90                    If tb.EOF Then
100                       tb.AddNew
110                       tb!SampleID = txtSampleID
120                       tb!SampleType = "S"
130                       tb!Code = BioCode
140                       tb!Result = tmmol(n)
150                       tb!Units = BioUnitsFor(BioCode)
160                       tb!Rundate = Format(Now, "dd/MMM/yyyy")
170                       tb!RunTime = Format(Now, "dd/MMM/yyyy hh:mm:ss")
180                       tb!Operator = UserCode
190                       tb.Update
200                   End If
210               End If
220               tmmol(n).Locked = True
230               tmmol(n).BackColor = &H80000011
240           End If
250       Next



          'SAVE 24 HOUR URINE CALCUATION
260       For n = 0 To 8
270           If tper24(n) <> "" Then
280               BioCode = GetCode(n)
290               If BioCode <> "" Then
300                   sql = "SELECT * FROM BioResults WHERE " & _
                            "SampleID = '" & txtSampleID & "' " & _
                            "AND Code = '" & BioCode & "'"
310                   Set tb = New Recordset
320                   RecOpenServer 0, tb, sql
330                   If Not tb.EOF Then
340                       sql = "SELECT * FROM BioRepeats WHERE 0 = 1"
350                       Set tb = New Recordset
360                       RecOpenServer 0, tb, sql
370                   End If
380                   tb.AddNew
390                   tb!SampleID = txtSampleID
400                   tb!SampleType = "U"
410                   tb!Code = BioCode
420                   tb!Result = tper24(n)
430                   tb!Units = BioUnitsFor(BioCode)
440                   tb!Rundate = Format(Now, "dd/MMM/yyyy")
450                   tb!RunTime = Format(Now, "dd/MMM/yyyy hh:mm:ss")
460                   tb!Operator = UserCode
470                   tb.Update
480               End If
490           End If

500       Next


          'SAVE 24HR VOLUME
510       sql = "SELECT * FROM BioResults WHERE " & _
                "SampleID = '" & Val(txtSampleID) & "' " & _
                "AND Code = '" & SysOptBioCodeFor24Vol(0) & "'"
520       Set tb = New Recordset
530       RecOpenServer 0, tb, sql
540       If Not tb.EOF Then
550           sql = "SELECT * FROM BioRepeats WHERE 0 = 1"
560           Set tb = New Recordset
570           RecOpenServer 0, tb, sql
580       End If
590       tb.AddNew
600       tb!SampleID = txtSampleID
610       tb!SampleType = "U"
620       tb!Code = SysOptBioCodeFor24Vol(0)
630       tb!Result = txtVolume
640       tb!Units = BioUnitsFor(SysOptBioCodeFor24Vol(0))
650       tb!Rundate = Format(Now, "dd/MMM/yyyy")
660       tb!RunTime = Format(Now, "dd/MMM/yyyy hh:mm:ss")
670       tb!Operator = UserCode
680       tb.Update



690       cmdSave.Enabled = False

700       Exit Sub

cmdSave_Click_Error:

          Dim strES As String
          Dim intEL As Integer

710       intEL = Erl
720       strES = Err.Description
730       LogError "frm24hrUrine", "cmdsave_Click", intEL, strES, sql

End Sub

Private Function GetCode(ByVal s As String) As String

10        On Error GoTo GetCode_Error

20        Select Case Val(s)
          Case 0
30            GetCode = SysOptBioCodeFor24UCreat(0)
40        Case 1
50            GetCode = SysOptBioCodeFor24UUrea(0)
60        Case 2
70            GetCode = SysOptBioCodeFor24UNa(0)
80        Case 3
90            GetCode = SysOptBioCodeFor24UK(0)
100       Case 4
110           GetCode = SysOptBioCodeFor24UChol(0)
120       Case 5
130           GetCode = SysOptBioCodeFor24UCA(0)
140       Case 6
150           GetCode = SysOptBioCodeFor24UPhos(0)
160       Case 7
170           GetCode = SysOptBioCodeFor24UMag(0)
180       Case 8
190           GetCode = SysOptBioCodeFor24UProt(0)
200       End Select

210       Exit Function

GetCode_Error:

          Dim strES As String
          Dim intEL As Integer

220       intEL = Erl
230       strES = Err.Description
240       LogError "frm24hrUrine", "GetCode", intEL, strES

End Function

Private Function GetCodeUrine(ByVal s As String) As String

10        On Error GoTo GetCodeUrine_Error

20        Select Case Val(s)
          Case 0
30            GetCodeUrine = SysOptBioCodeFor24UCreat(0)
40        Case 1
50            GetCodeUrine = SysOptBioCodeForUUrea(0)
60        Case 2
70            GetCodeUrine = SysOptBioCodeForUNa(0)
80        Case 3
90            GetCodeUrine = SysOptBioCodeForUK(0)
100       Case 4
110           GetCodeUrine = SysOptBioCodeForUChol(0)
120       Case 5
130           GetCodeUrine = SysOptBioCodeForUCA(0)
140       Case 6
150           GetCodeUrine = SysOptBioCodeForUPhos(0)
160       Case 7
170           GetCodeUrine = SysOptBioCodeForUMag(0)
180       Case 8
190           GetCodeUrine = SysOptBioCodeForUProt(0)
200       End Select

210       Exit Function

GetCodeUrine_Error:

          Dim strES As String
          Dim intEL As Integer

220       intEL = Erl
230       strES = Err.Description
240       LogError "frm24hrUrine", "GetCodeUrine", intEL, strES

End Function



Private Function BioUnitsFor(ByVal Code As String) As String

          Dim sql As String
          Dim tb As Recordset

10        On Error GoTo BioUnitsFor_Error

20        sql = "SELECT Units FROM BioTestDefinitions WHERE " & _
                "Code = '" & Code & "'"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        If Not tb.EOF Then
60            BioUnitsFor = tb!Units
70        Else
80            BioUnitsFor = ""
90        End If

100       Exit Function

BioUnitsFor_Error:

          Dim strES As String
          Dim intEL As Integer

110       intEL = Erl
120       strES = Err.Description
130       LogError "frm24hrUrine", "BioUnitsFor", intEL, strES, sql

End Function
Private Sub Form_Load()

10        If TestSys = True Then Me.Caption = Me.Caption & " - TEST SYSTEM"

End Sub


Private Sub tmmol_KeyPress(Index As Integer, KeyAscii As Integer)
10        KeyAscii = VI(KeyAscii, Numericfullstopdash)
End Sub

Private Sub txtSampleID_LostFocus()

          Dim sn As New Recordset
          Dim sql As String
          Dim br As BIEResult
          Dim BRs As New BIEResults
          Dim n As Long
          Dim OffSet As Long
          Dim CurrentPrintFormat As String


10        On Error GoTo txtSampleID_LostFocus_Error

20        If Trim(txtSampleID) = "" Then Exit Sub

30        txtSampleID = Format(Val(txtSampleID))

40        sql = "SELECT * FROM Demographics WHERE " & _
                "SampleID = '" & txtSampleID & "'"

50        Set sn = New Recordset
60        RecOpenClient 0, sn, sql
70        If sn.EOF Then
80            tName = ""
90            tChart = ""
100           tDoB = ""
110           lRunDate = ""
120       Else
130           tChart = sn!Chart & ""
140           tName = sn!PatName & ""
150           tDoB = sn!Dob & ""
160           lRunDate = Format(sn!Rundate, "dd/mm/yyyy")
170       End If

180       lRunTime = ""

190       For n = 0 To 9
200           strmMol(n) = ""
210           strPer24(n) = ""
220       Next

230       Set BRs = BRs.Load("Bio", txtSampleID, "Results", gDONTCARE, gDONTCARE, 0, "", lRunDate)
240       For Each br In BRs
250           Select Case br.Code
              Case SysOptBioCodeFor24Vol(0): txtVolume = br.Result: OffSet = -1      'Urine Volume
260           Case 887: tUnits = br.Result: OffSet = -1    'Urine Collection Hours
270           Case SysOptBioCodeForUCreat(0):    'Creatinine
280               OffSet = 0
290           Case SysOptBioCodeForUUrea(0):    'Urea
300               OffSet = 1
310           Case SysOptBioCodeForUNa(0):    'Sodium
320               OffSet = 2
330           Case SysOptBioCodeForUK(0):    'Potassium
340               OffSet = 3
350           Case SysOptBioCodeForUChol(0):    'Chloride
360               OffSet = 4
370           Case SysOptBioCodeForUCA(0):    'Calcium
380               OffSet = 5
390           Case SysOptBioCodeForUPhos(0):    'Phosphorus
400               OffSet = 6
410           Case SysOptBioCodeForUMag(0):    'Magnesium
420               OffSet = 7
430           Case SysOptBioCodeForUProt(0):    'TProt
440               OffSet = 8
450           Case Else:
460               OffSet = -1
470           End Select
480           If OffSet > -1 Then
490               Select Case br.Printformat
                  Case 0: CurrentPrintFormat = "####0   "
500               Case 1: CurrentPrintFormat = "####0.0 "
510               Case 2: CurrentPrintFormat = "####0.00"
520               Case 3: CurrentPrintFormat = "###0.000"
530               End Select
540               lRunTime = Format(br.RunTime, "dd/mm/yy hh:mm")
550               strmMol(OffSet) = Format(br.Result, CurrentPrintFormat)
560               tmmol(OffSet) = Format(br.Result, CurrentPrintFormat)
570               tmmol(OffSet).Locked = True
580               tmmol(OffSet).BackColor = &H80000011
590               strDP(OffSet) = CurrentPrintFormat
600           End If
610       Next

620       Calculate




630       Exit Sub

txtSampleID_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

640       intEL = Erl
650       strES = Err.Description
660       LogError "frm24hrUrine", "txtSampleID_LostFocus", intEL, strES, sql


End Sub

Private Sub txtVolume_LostFocus()

10        Calculate
20        cmdSave.Enabled = True

End Sub


