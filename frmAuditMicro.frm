VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmAuditMicro 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Microbiology Audit"
   ClientHeight    =   8370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13095
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8370
   ScaleWidth      =   13095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   1100
      Left            =   11700
      Picture         =   "frmAuditMicro.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   5160
      Width           =   1200
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Default         =   -1  'True
      Height          =   690
      Left            =   2070
      Picture         =   "frmAuditMicro.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   90
      Width           =   960
   End
   Begin VB.Frame Frame1 
      Caption         =   "Department"
      Height          =   795
      Left            =   3210
      TabIndex        =   4
      Top             =   0
      Width           =   8265
      Begin VB.OptionButton optDept 
         Caption         =   "Faeces"
         Enabled         =   0   'False
         Height          =   195
         Index           =   0
         Left            =   1260
         TabIndex        =   16
         Top             =   510
         Width           =   825
      End
      Begin VB.OptionButton optDept 
         Alignment       =   1  'Right Justify
         Caption         =   "Site Details"
         Enabled         =   0   'False
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   15
         Top             =   510
         Width           =   1095
      End
      Begin VB.OptionButton optDept 
         Caption         =   "Urine"
         Enabled         =   0   'False
         Height          =   195
         Index           =   2
         Left            =   1260
         TabIndex        =   14
         Top             =   270
         Width           =   675
      End
      Begin VB.OptionButton optDept 
         Caption         =   "Generic"
         Enabled         =   0   'False
         Height          =   195
         Index           =   3
         Left            =   3870
         TabIndex        =   13
         Top             =   270
         Width           =   915
      End
      Begin VB.OptionButton optDept 
         Alignment       =   1  'Right Justify
         Caption         =   "Isolates"
         Enabled         =   0   'False
         Height          =   195
         Index           =   4
         Left            =   360
         TabIndex        =   12
         Top             =   270
         Width           =   885
      End
      Begin VB.OptionButton optDept 
         Alignment       =   1  'Right Justify
         Caption         =   "Identification"
         Enabled         =   0   'False
         Height          =   195
         Index           =   8
         Left            =   5250
         TabIndex        =   11
         Top             =   270
         Width           =   1245
      End
      Begin VB.OptionButton optDept 
         Caption         =   "Sensitivities"
         Enabled         =   0   'False
         Height          =   195
         Index           =   5
         Left            =   3870
         TabIndex        =   10
         Top             =   510
         Width           =   1125
      End
      Begin VB.OptionButton optDept 
         Alignment       =   1  'Right Justify
         Caption         =   "Urine Requests"
         Enabled         =   0   'False
         Height          =   195
         Index           =   6
         Left            =   2400
         TabIndex        =   9
         Top             =   270
         Width           =   1425
      End
      Begin VB.OptionButton optDept 
         Alignment       =   1  'Right Justify
         Caption         =   "Faecal Requests"
         Enabled         =   0   'False
         Height          =   195
         Index           =   7
         Left            =   2280
         TabIndex        =   8
         Top             =   510
         Width           =   1545
      End
      Begin VB.OptionButton optDept 
         Caption         =   "Semen"
         Enabled         =   0   'False
         Height          =   195
         Index           =   11
         Left            =   6540
         TabIndex        =   7
         Top             =   510
         Width           =   795
      End
      Begin VB.OptionButton optDept 
         Caption         =   "Demographics"
         Enabled         =   0   'False
         Height          =   195
         Index           =   9
         Left            =   6540
         TabIndex        =   6
         Top             =   270
         Width           =   1335
      End
      Begin VB.OptionButton optDept 
         Alignment       =   1  'Right Justify
         Caption         =   "Comments"
         Enabled         =   0   'False
         Height          =   195
         Index           =   10
         Left            =   5400
         TabIndex        =   5
         Top             =   510
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   1100
      Left            =   11700
      Picture         =   "frmAuditMicro.frx":284C
      Style           =   1  'Graphical
      TabIndex        =   2
      Tag             =   "101"
      ToolTipText     =   "Exit Screen"
      Top             =   7080
      Width           =   1200
   End
   Begin VB.TextBox txtSampleID 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   180
      TabIndex        =   0
      Top             =   330
      Width           =   1770
   End
   Begin RichTextLib.RichTextBox rtb 
      Height          =   7065
      Left            =   150
      TabIndex        =   1
      Top             =   1110
      Width           =   11325
      _ExtentX        =   19976
      _ExtentY        =   12462
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmAuditMicro.frx":3716
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   180
      TabIndex        =   3
      Top             =   870
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Default Printer"
      ForeColor       =   &H00000040&
      Height          =   195
      Left            =   11700
      TabIndex        =   20
      Top             =   4860
      Width           =   1200
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Sample ID"
      Height          =   195
      Left            =   600
      TabIndex        =   18
      Top             =   60
      Width           =   735
   End
End
Attribute VB_Name = "frmAuditMicro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pTableName As String
Private pTableNameAudit As String
Private pWhereClause As String

Public Property Let SampleID(ByVal sNewValue As String)

10        txtSampleID = sNewValue
20        FillOptions

End Property
Public Property Let TableName(ByVal sNewValue As String)

10        pTableName = sNewValue
20        pTableNameAudit = sNewValue & "Audit"

End Property

Public Property Let WhereClause(ByVal sNewValue As String)
10        pWhereClause = sNewValue
End Property

Private Sub LoadAudit()

          Dim sql As String
          Dim tb As Recordset
          Dim tbArc As Recordset
          Dim n As Integer
          Dim CurrentName() As String
          Dim current() As String
          Dim NameDisplayed As Boolean
          Dim NameSuffix As String
          Dim SID As Double
          Dim AuditFound As Boolean

10        On Error GoTo LoadAudit_Error

20        SID = Val(txtSampleID) + SysOptMicroOffset(0)
30        If UCase$(pTableName) = "SEMENRESULTS" Then
40            SID = Val(txtSampleID) + SysOptSemenOffset(0)
50        End If

60        AuditFound = False
70        rtb.Text = ""
80        rtb.SelFontSize = 12

90        If Trim$(txtSampleID) = "" Then Exit Sub

100       PrintTextRTB rtb, "SampleID: " & txtSampleID & vbCrLf, 16, True
110       PrintTextRTB rtb, "Audit Trail for ", 16, True, , True
120       PrintTextRTB rtb, pTableName & vbCrLf & vbCrLf, 16, True, , True, vbRed

130       rtb.SelUnderline = False
140       rtb.SelFontSize = 12

150       sql = "SELECT * FROM " & pTableName & " WHERE " & _
                "SampleID = " & SID
160       Set tb = New Recordset
170       RecOpenServer 0, tb, sql
180       If tb.EOF Then
190           rtb.SelText = "No Current Record found." & vbCrLf
200           Exit Sub
210       End If

220       ReDim current(0 To tb.Fields.Count - 1)
230       ReDim CurrentName(0 To tb.Fields.Count - 1)

240       While Not tb.EOF
250           For n = 0 To tb.Fields.Count - 1
260               If Not tb.EOF Then
270                   current(n) = tb.Fields(n).Value & ""
280               Else
290                   current(n) = ""
300               End If
310               CurrentName(n) = tb.Fields(n).Name
320               NameSuffix = ""
330               If UCase(tb.Fields(n).Name) <> "ROWGUID" And UCase(tb.Fields(n).Name) <> "DATETIMEOFARCHIVE" _
                     And UCase(tb.Fields(n).Name) <> "ARCHIVEOPERATOR" Then

340                   sql = "SELECT ArchivedBy, ArchiveDateTime, [" & CurrentName(n) & "] FROM " & pTableNameAudit & " " & _
                            pWhereClause & _
                            " ORDER BY ArchiveDateTime DESC"

350                   Select Case pTableName
                      Case "Faeces", "Urine", "UrineRequests", "FaecalRequests", "UrineIdent", "Demographics", "SemenResults":
360                       sql = Replace(sql, "%sampleid", tb!SampleID)
370                       NameSuffix = ""
380                   Case "MicroSiteDetails":
390                       sql = Replace(sql, "%sampleid", tb!SampleID)
400                       sql = Replace(sql, "%site", tb!Site)
410                       NameSuffix = "(" & tb!Site & ")"
420                   Case "GenericResults":
430                       sql = Replace(sql, "%sampleid", tb!SampleID)
440                       sql = Replace(sql, "%testname", tb!TestName)
450                       NameSuffix = "(" & tb!TestName & ")"
460                   Case "Isolates":
470                       sql = Replace(sql, "%sampleid", tb!SampleID)
480                       sql = Replace(sql, "%isolatenumber", tb!IsolateNumber)
490                       NameSuffix = "(" & tb!OrganismName & ")"
500                   Case "Sensitivities":
510                       sql = Replace(sql, "%sampleid", tb!SampleID)
520                       sql = Replace(sql, "%isolatenumber", tb!IsolateNumber)
530                       sql = Replace(sql, "%antibioticcode", tb!AntibioticCode)
540                       NameSuffix = tb!Antibiotic & "  (" & tb!IsolateNumber & ")"
550                   Case "Observations":
560                       sql = Replace(sql, "%sampleid", tb!SampleID)
570                       sql = Replace(sql, "%disicipline", tb!Discipline)
580                       NameSuffix = "(" & tb!Discipline & ")"
590                   End Select

600                   Set tbArc = New Recordset
610                   RecOpenServer 0, tbArc, sql
620                   If tbArc.EOF Then
                          '                rtb.SelText = "No Changes Made." & vbCrLf
                          '                If Trim$(tb!UserName & "") <> "" Then
                          '                    rtb.SelText = "Original entry by " & tb!UserName
                          '                End If
                          '                Exit For
630                   Else
640                       AuditFound = True
650                       NameDisplayed = False
660                       Do While Not tbArc.EOF
670                           If Trim$(current(n)) <> Trim$(tbArc.Fields(CurrentName(n)) & "") Then
680                               If Not NameDisplayed Then
690                                   PrintTextRTB rtb, CurrentName(n) & Space(4) & NameSuffix & vbCrLf, 12, True, , , vbBlue
700                                   NameDisplayed = True
710                               End If
720                               PrintTextRTB rtb, tbArc!ArchiveDateTime & " ", 12
730                               PrintTextRTB rtb, tbArc!ArchivedBy & "", 12, , , , vbRed
740                               PrintTextRTB rtb, " Changed ", 12
750                               If Trim$(tbArc.Fields(CurrentName(n)) & "") = "" Then
760                                   PrintTextRTB rtb, "<Blank> ", 12, True, , , vbGreen
770                               Else
780                                   PrintTextRTB rtb, LeftOfBar(Trim$(tbArc.Fields(CurrentName(n)))), 12, True, , , vbGreen
790                               End If
800                               PrintTextRTB rtb, " to ", 12
810                               If Trim$(current(n)) = "" Then
820                                   PrintTextRTB rtb, "<Blank> ", 12, True
830                               Else
840                                   PrintTextRTB rtb, LeftOfBar(Trim$(current(n))) & vbCrLf, 12, True
850                               End If


860                               current(n) = Trim$(tbArc.Fields(CurrentName(n)) & "")
870                           End If
880                           tbArc.MoveNext
890                       Loop
900                   End If
910                   If NameDisplayed Then
920                       rtb.SelText = vbCrLf
930                   End If
940               End If
950           Next
960           tb.MoveNext
970       Wend

980       If Not AuditFound Then
990           rtb.SelText = "No Changes Made." & vbCrLf
1000      End If

1010      Exit Sub

LoadAudit_Error:

          Dim strES As String
          Dim intEL As Integer

1020      intEL = Erl
1030      strES = Err.Description
1040      LogError "frmAuditMicro", "LoadAudit", intEL, strES, sql

End Sub


Private Sub cmdCancel_Click()

10        Unload Me

End Sub


Private Sub FillOptions()

          Dim tb As Recordset
          Dim sql As String
          Dim SID As Double
          Dim Index As Integer
          Dim Dept As String

10        On Error GoTo FillOptions_Error

20        For Index = 0 To 11
30            optDept(Index).Enabled = False
40            optDept(Index).ForeColor = vbBlack
50        Next

60        rtb.Text = ""

70        SID = Val(txtSampleID) + SysOptMicroOffset(0)

80        For Index = 0 To 11
90            Dept = Choose(Index + 1, "FaecesAudit", "MicroSiteDetailsAudit", _
                            "UrineAudit", "GenericResultsAudit", "IsolatesAudit", "SensitivitiesAudit", _
                            "UrineRequestsAudit", "FaecalRequestsAudit", "UrineIdentAudit", _
                            "DemographicsAudit", "ObservationsAudit", "SemenResultsAudit")
100           If Index = 11 Then
110               SID = Val(txtSampleID) + SysOptSemenOffset(0)
120           End If
130           sql = "SELECT COUNT(*) AS Tot FROM " & Dept & " WHERE " & _
                    "SampleID = '" & SID & "'"
140           Set tb = New Recordset
150           RecOpenServer 0, tb, sql
160           If tb!Tot > 0 Then
170               optDept(Index).Enabled = True
180               optDept(Index).ForeColor = vbRed
190           Else
200               Dept = Choose(Index + 1, "Faeces", "MicroSiteDetails", _
                                "Urine", "GenericResults", "Isolates", "Sensitivities", _
                                "UrineRequests", "FaecalRequests", "UrineIdent", "Demographics", "Observations", "SemenResults")
210               sql = "SELECT COUNT(*) AS Tot FROM " & Dept & " WHERE " & _
                        "SampleID = '" & SID & "'"
220               Set tb = New Recordset
230               RecOpenServer 0, tb, sql
240               If tb!Tot > 0 Then
250                   optDept(Index).Enabled = True
260                   optDept(Index).ForeColor = vbBlack
270               End If
280           End If
290       Next

300       Exit Sub

FillOptions_Error:

          Dim strES As String
          Dim intEL As Integer

310       intEL = Erl
320       strES = Err.Description
330       LogError "frmAuditMicro", "FillOptions", intEL, strES, sql

End Sub

Private Sub cmdPrint_Click()

10        On Error GoTo cmdPrint_Click_Error


20        rtb.SelStart = 0
30        rtb.SelLength = 10000000#
40        rtb.SelPrint Printer.hDC

50        Exit Sub

cmdPrint_Click_Error:

          Dim strES As String
          Dim intEL As Integer

60        intEL = Erl
70        strES = Err.Description
80        LogError "frmAuditMicro", "cmdPrint_Click", intEL, strES

End Sub

Private Sub cmdStart_Click()

10        FillOptions

End Sub

Private Sub optDept_Click(Index As Integer)

10        rtb.Text = ""

20        Select Case Index
          Case 0: pTableName = "Faeces"
30            pWhereClause = "WHERE SampleID = %sampleid"
40        Case 1: pTableName = "MicroSiteDetails"
50            pWhereClause = "WHERE SampleID = %sampleid AND Site = '%site'"
60        Case 2: pTableName = "Urine"
70            pWhereClause = "WHERE SampleID = %sampleid"
80        Case 3: pTableName = "GenericResults"
90            pWhereClause = "WHERE SampleID = %sampleid AND TestName = '%testname'"
100       Case 4: pTableName = "Isolates"
110           pWhereClause = "WHERE SampleID = %sampleid AND IsolateNumber = %isolatenumber"
120       Case 5: pTableName = "Sensitivities"
130           pWhereClause = "WHERE SampleID = %sampleid AND IsolateNumber = %isolatenumber AND AntibioticCode = '%antibioticcode'"
140       Case 6: pTableName = "UrineRequests"
150           pWhereClause = "WHERE SampleID = %sampleid"
160       Case 7: pTableName = "FaecalRequests"
170           pWhereClause = "WHERE SampleID = %sampleid"
180       Case 8: pTableName = "UrineIdent"
190           pWhereClause = "WHERE SampleID = %sampleid"
200       Case 9: pTableName = "Demographics"
210           pWhereClause = "WHERE SampleID = %sampleid"
220       Case 10: pTableName = "Observations"
230           pWhereClause = "WHERE SampleID = %sampleid AND Discipline = '%discipline'"
240       Case 11: pTableName = "SemenResults"
250           pWhereClause = "WHERE SampleID = %sampleid"
260       End Select

270       pTableNameAudit = pTableName & "Audit"

280       LoadAudit

End Sub


