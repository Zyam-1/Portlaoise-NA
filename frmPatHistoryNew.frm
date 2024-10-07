VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form frmPatHistoryNew 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Search"
   ClientHeight    =   10410
   ClientLeft      =   105
   ClientTop       =   600
   ClientWidth     =   15015
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
   Icon            =   "frmPatHistoryNew.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   10410
   ScaleWidth      =   15015
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame5 
      Caption         =   "Search By"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   180
      TabIndex        =   25
      Top             =   765
      Width           =   3105
      Begin VB.OptionButton oSB 
         Alignment       =   1  'Right Justify
         Caption         =   "Date"
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
         Index           =   0
         Left            =   60
         TabIndex        =   27
         Top             =   210
         Value           =   -1  'True
         Width           =   1305
      End
      Begin VB.OptionButton oSB 
         Alignment       =   1  'Right Justify
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
         Height          =   225
         Index           =   1
         Left            =   1575
         TabIndex        =   26
         Top             =   225
         Width           =   1275
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Search For"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8595
      TabIndex        =   18
      Top             =   90
      Width           =   2775
      Begin VB.OptionButton oFor 
         Alignment       =   1  'Right Justify
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
         Index           =   0
         Left            =   75
         TabIndex        =   23
         Top             =   240
         Value           =   -1  'True
         Width           =   930
      End
      Begin VB.OptionButton oFor 
         Alignment       =   1  'Right Justify
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
         Index           =   1
         Left            =   45
         TabIndex        =   22
         Top             =   480
         Width           =   960
      End
      Begin VB.OptionButton oFor 
         Alignment       =   1  'Right Justify
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
         Index           =   2
         Left            =   45
         TabIndex        =   21
         Top             =   720
         Width           =   960
      End
      Begin VB.CheckBox chkSoundex 
         Caption         =   "Use Soundex"
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
         Left            =   1050
         TabIndex        =   20
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton oFor 
         Caption         =   "Name+DoB"
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
         Index           =   3
         Left            =   1050
         TabIndex        =   19
         Top             =   720
         Width           =   1575
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Records"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   13590
      TabIndex        =   15
      Top             =   90
      Width           =   1275
      Begin VB.TextBox txtRecords 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   16
         Text            =   "frmPatHistoryNew.frx":030A
         Top             =   270
         Width           =   765
      End
      Begin ComCtl2.UpDown udRecords 
         Height          =   225
         Left            =   240
         TabIndex        =   17
         Top             =   570
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   397
         _Version        =   327681
         Value           =   20
         BuddyControl    =   "txtRecords"
         BuddyDispid     =   196615
         OrigLeft        =   150
         OrigTop         =   450
         OrigRight       =   915
         OrigBottom      =   690
         Increment       =   20
         Max             =   9999
         Min             =   20
         Orientation     =   1
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
   End
   Begin VB.TextBox txtDoB 
      BackColor       =   &H00FFFF00&
      Enabled         =   0   'False
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   2970
      TabIndex        =   14
      Text            =   "Date of Birth"
      Top             =   1320
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Frame fraSearch 
      Caption         =   "How"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   11340
      TabIndex        =   10
      Top             =   90
      Width           =   2175
      Begin VB.OptionButton optExact 
         Caption         =   "Exact Match"
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
         TabIndex        =   13
         Top             =   240
         Width           =   1905
      End
      Begin VB.OptionButton optLeading 
         Caption         =   "Leading Characters"
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
         TabIndex        =   12
         Top             =   480
         Value           =   -1  'True
         Width           =   1950
      End
      Begin VB.OptionButton optTrailing 
         Caption         =   "Trailing Characters"
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
         TabIndex        =   11
         Top             =   720
         Width           =   1890
      End
   End
   Begin VB.Frame Frame2 
      Height          =   495
      Left            =   3540
      TabIndex        =   9
      Top             =   30
      Width           =   3585
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   90
         MaxLength       =   20
         TabIndex        =   0
         Top             =   150
         Width           =   3375
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   180
      TabIndex        =   5
      Top             =   60
      Width           =   3105
      Begin VB.CheckBox chkType 
         Caption         =   "As You Type"
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
         Left            =   1395
         TabIndex        =   28
         Top             =   495
         Width           =   1650
      End
      Begin VB.CheckBox chkOtherHosps 
         Caption         =   "Other Hospitals"
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
         Left            =   1395
         TabIndex        =   24
         Top             =   225
         Width           =   1650
      End
      Begin VB.OptionButton oHD 
         Alignment       =   1  'Right Justify
         Caption         =   "Historic"
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
         Index           =   1
         Left            =   90
         TabIndex        =   7
         Top             =   450
         Value           =   -1  'True
         Width           =   1275
      End
      Begin VB.OptionButton oHD 
         Alignment       =   1  'Right Justify
         Caption         =   "Download"
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
         Index           =   0
         Left            =   60
         TabIndex        =   6
         Top             =   210
         Width           =   1305
      End
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   9045
      Left            =   60
      TabIndex        =   4
      Top             =   1305
      Width           =   14805
      _ExtentX        =   26114
      _ExtentY        =   15954
      _Version        =   393216
      Cols            =   23
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      BackColorSel    =   16711680
      ForeColorSel    =   65280
      ScrollTrack     =   -1  'True
      GridLines       =   3
      GridLinesFixed  =   3
      AllowUserResizing=   1
      FormatString    =   $"frmPatHistoryNew.frx":030F
   End
   Begin VB.CommandButton bcopy 
      Appearance      =   0  'Flat
      Caption         =   "Copy to &Edit"
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
      Height          =   735
      Left            =   5940
      Picture         =   "frmPatHistoryNew.frx":03C7
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   540
      Width           =   1095
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
      Height          =   735
      Left            =   4830
      Picture         =   "frmPatHistoryNew.frx":06D1
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   540
      Width           =   1035
   End
   Begin VB.CommandButton bsearch 
      Caption         =   "&Search"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3660
      Picture         =   "frmPatHistoryNew.frx":09DB
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   540
      Width           =   1095
   End
   Begin VB.Label lNoPrevious 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "No Previous Details"
      ForeColor       =   &H000000FF&
      Height          =   465
      Left            =   7170
      TabIndex        =   8
      Top             =   750
      Visible         =   0   'False
      Width           =   1335
   End
End
Attribute VB_Name = "frmPatHistoryNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private NoPrevious As Boolean
Private mFromEdit As Boolean
Private mEditScreen As Form
Private Activated As Boolean
Public mDemoVal As Boolean
Private mFromLookup As Boolean

Private pWithin As Long    'Used for fuzzy DoB search

Private Sub FillG()

10        On Error GoTo FillG_Error

20        lNoPrevious.Visible = False

30        If Trim$(txtName) = "" Then
40            iMsg "No criteria entered!"
50            Exit Sub
60        End If

70        If oFor(2).Value = True Then
80            txtName = Convert62Date(txtName, BACKWARD)
90            If Not IsDate(txtName) Then
100               iMsg "Please provide valid date of birth", vbInformation
110               Exit Sub
120           End If
'Wrike #638831592 commented out 130 - 160 below
'130           If DateDiff("yyyy", txtName, Now) > 120 Then
'140               iMsg "Please provide valid date of birth", vbInformation
'150               Exit Sub
'160           End If
170       End If

180       ClearFGrid g

190       LocalFillG

200       With g
210           If .Rows > 2 Then
220               .RemoveItem 1
230               .Row = 1
240               If oHD(0) Then
250                   .Col = 0
260               Else
270                   .Col = 11
280               End If
290               .ColSel = .Cols - 1
300               .RowSel = 1
310               .HighLight = flexHighlightAlways
320           End If
330       End With

340       If mFromEdit = True Then
350           If mDemoVal = True Then
360               bcopy.Enabled = False
370           Else
380               bcopy.Enabled = mFromEdit
390           End If
400       End If

410       g.Visible = True

420       Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer

430       intEL = Erl
440       strES = Err.Description
450       LogError "frmPatHistoryNew", "FillG", intEL, strES

End Sub

Public Property Let EditScreen(ByVal f As Form)

10        On Error GoTo EditScreen_Error

20        Set mEditScreen = f

30        Exit Property

EditScreen_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmPatHistoryNew", "EditScreen", intEL, strES

End Property




Private Sub LocalFillG()

      Dim s As String
      Dim tb As New Recordset
      Dim tbCount As Recordset
      Dim sql As String
      Dim n As Long
      Dim Cn As Long
      Dim Obs As Observations
      Dim Ob As Observation
      Dim TabName As String
      Dim Criteria As String
      Dim SemenComment As Boolean
      Dim HistologyComment As Boolean
      Dim CytologyComment As Boolean
      Dim CoagulationComment As Boolean
      Dim HaematologyComment As Boolean
      Dim BiochemistryComment As Boolean
      Dim EndocrinologyComment As Boolean
      Dim ImmunologyComment As Boolean
      Dim BloodGasComment As String
      Dim MicroGeneralComment As Boolean

10    On Error GoTo LocalFillG_Error

20    Criteria = ""

30    If InStr(txtName, "%") > 0 Or InStr(txtName, "_") > 0 Or txtName = "'" Then
40        iMsg "Invalid Search Criteria!"
50        Exit Sub
60    End If

70    If oHD(0) Then
80        TabName = "PatientIFs."
90    Else
100       TabName = "Demographics."
110   End If

120   If oFor(0) Then
130       If chkSoundex = 1 Then
140           Criteria = Criteria & "  REPLACE(" & TabName & "PatName,'''','') = soundex('" & AddTicks(Replace(txtName, "'", " ")) & "') "
150       Else
160           If optExact Then
                  'Criteria = Criteria & TabName & "PatName = '" & AddTicks(txtName) & "' "
170               Criteria = Criteria & SearchForIrishNames(txtName, ExactMatch)

180           ElseIf optLeading Then
                  'Criteria = Criteria & TabName & "PatName like '" & AddTicks(txtName) & "%' "
190               Criteria = Criteria & SearchForIrishNames(txtName, LeadingCharacters)
200           Else
                  'Criteria = Criteria & TabName & "PatName like '%" & AddTicks(txtName) & "' "
210               Criteria = Criteria & SearchForIrishNames(txtName, TrailingCharacters)
220           End If
230       End If
240   ElseIf oFor(1) Then
250       Criteria = Criteria & TabName & "chart = '" & UCase(AddTicks(txtName)) & "' "
260   ElseIf oFor(2) Then
270       txtName = Convert62Date(txtName, BACKWARD)
280       If Not IsDate(txtName) Then
290           iMsg "Invalid Date", vbExclamation, "Date of Birth Search"
300           Exit Sub
310       End If
320       Criteria = Criteria & TabName & "DoB = '" & Format$(txtName, "dd/mmm/yyyy") & "'"
330   Else    'Name+DoB
340       If chkSoundex = 1 Then
350           Criteria = Criteria & "  REPLACE(" & TabName & "PatName,'''','') = soundex('" & AddTicks(Replace(txtName, "'", " ")) & "') "
360       Else
370           If optExact Then
                  'Criteria = Criteria & TabName & "PatName = '" & AddTicks(txtName) & "' "
380               Criteria = Criteria & SearchForIrishNames(txtName, ExactMatch)

390           ElseIf optLeading Then
                  'Criteria = Criteria & TabName & "PatName like '" & AddTicks(txtName) & "%' "
400               Criteria = Criteria & SearchForIrishNames(txtName, LeadingCharacters)
410           Else
                  'Criteria = Criteria & TabName & "PatName like '%" & AddTicks(txtName) & "' "
420               Criteria = Criteria & SearchForIrishNames(txtName, TrailingCharacters)
430           End If
440       End If
450       If pWithin = 0 Then
460           Criteria = Criteria & "and " & TabName & "DoB = '" & Format$(txtDoB, "dd/mmm/yyyy") & "' "
470       Else
480           Criteria = Criteria & "and " & TabName & "DoB between '" & Format$(DateAdd("yyyy", -pWithin, txtDoB), "dd/mmm/yyyy") & "' " & _
                         "and '" & Format$(DateAdd("yyyy", pWithin, txtDoB), "dd/mmm/yyyy") & "' "
490       End If
500   End If

510   If oHD(0) Then
520       Criteria = Criteria & " order by "
530       Criteria = Criteria & TabName & "DateTimeAmended desc"
540   End If

550   If oHD(0) Then
560       sql = "SELECT top " & Format$(Val(txtRecords)) & " * from " & _
                "PatientIFs WHERE " & Criteria
570       TabName = "Patientifs."
580   Else
590       sql = "SELECT DISTINCT " & _
                "TOP " & Val(txtRecords) & " SampleID, " & _
                "AandE, PatName, Chart, RunDate, SampleDate, GP, hyear, Clinician, " & _
                "Ward, Addr0, Addr1, Age, Sex, Hospital, DoB " & _
                "FROM Demographics WHERE " & Criteria & " " & _
                "ORDER BY SampleDate DESC, SampleID DESC"
600       TabName = "Demographics."
610   End If

620   NoPrevious = True

630   If chkOtherHosps = 0 Then
640       Cn = 0
650   Else
660       Cn = intOtherHospitalsInGroup
670   End If

680   For n = 0 To Cn
690       Set tb = New Recordset

700       RecOpenClient n, tb, sql

710       With tb
720           If Not .EOF Then
730               NoPrevious = False
740           End If
750           g.Visible = False
760           If oHD(1) Then

770               Do While Not .EOF
780                   s = vbTab & vbTab & vbTab & vbTab & vbTab & vbTab
790                   s = s & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & _
                          Format$(!SampleDate, "dd/mm/yy") & vbTab & _
                          Trim$(!SampleID & "") & vbTab & _
                          !AandE & vbTab
800                   If Trim(!Chart) & "" <> "" Then
810                       If UCase(tb!Hospital) = UCase(HospName(0)) Then
820                           s = s & !Chart & vbTab
830                       Else
840                           If Not IsNumeric(Left(!Chart, 1)) Then
850                               s = s & !Chart & vbTab
860                           Else
870                               s = s & Left(tb!Hospital, 1) & !Chart & vbTab
880                           End If
890                       End If
900                   Else
910                       s = s & vbTab
920                   End If
930                   s = s & initial2upper(!PatName) & vbTab & _
                          Format$(!Dob, "dd/mm/yyyy") & vbTab

940                   If !Age & "" = "" Then
950                       If Trim(!Dob) & "" <> "" Then s = s & CalcOldAge(!Dob, !Rundate)
960                   Else
970                       s = s & !Age
980                   End If

990                   s = s & vbTab & _
                          !sex & vbTab & _
                          !Addr0 & vbTab & _
                          !Addr1 & vbTab & _
                          Trim$(!Ward & "") & vbTab & _
                          initial2upper(!Clinician & "") & vbTab & _
                          !GP & vbTab & _
                          HospName(n) & vbTab & vbTab & Trim(!Hyear & "")

1000                  g.AddItem s
1010                  g.Row = g.Rows - 1

1020                  SemenComment = False
1030                  HistologyComment = False
1040                  CytologyComment = False
1050                  CoagulationComment = False
1060                  HaematologyComment = False
1070                  BiochemistryComment = False
1080                  EndocrinologyComment = False
1090                  ImmunologyComment = False
1100                  BloodGasComment = False
1110                  MicroGeneralComment = False

1120                  Set Obs = New Observations
1130                  Set Obs = Obs.Load(!SampleID, "MicroGeneral", "BloodGas", "Immunology", "Endocrinology", _
                                         "Biochemistry", "Haematology", "Semen", "Histology", _
                                         "Cytology", "Coagulation")
1140                  If Not Obs Is Nothing Then
1150                      For Each Ob In Obs
1160                          Select Case UCase$(Ob.Discipline)
                                  Case "MICROGENERAL": MicroGeneralComment = True
1170                              Case "BLOODGAS": BloodGasComment = True
1180                              Case "IMMUNOLOGY": ImmunologyComment = True
1190                              Case "ENDOCRINOLOGY": EndocrinologyComment = True
1200                              Case "BIOCHEMISTRY": BiochemistryComment = True
1210                              Case "HAEMATOLOGY": HaematologyComment = True
1220                              Case "SEMEN": SemenComment = True
1230                              Case "HISTOLOGY": HistologyComment = True
1240                              Case "CYTOLOGY": CytologyComment = True
1250                              Case "COAGULATION": CoagulationComment = True
1260                          End Select
1270                      Next
1280                  End If

1290                  sql = "SELECT " & _
                            "HaemCount = (SELECT COUNT(*) FROM HaemResults WHERE sampleid = '" & !SampleID & "'), " & _
                            "BioCount = (SELECT count(*) FROM BioResults WHERE sampleid  = '" & !SampleID & "'), " & _
                            "CoagCount = (SELECT count(*) FROM CoagResults WHERE sampleid = '" & !SampleID & "'), " & _
                            "EndCount = (SELECT count(*) FROM EndResults WHERE sampleid  = '" & !SampleID & "'), " & _
                            "ImmCount = (SELECT count(*) FROM ImmResults WHERE sampleid  = '" & !SampleID & "'), " & _
                            "BGACount = (SELECT count(*) FROM BGAResults WHERE sampleid  = '" & !SampleID & "'), " & _
                            "ExtCount = (SELECT count(*) FROM ExtResults WHERE sampleid  = '" & !SampleID & "')"
1300                  Set tbCount = New Recordset
1310                  RecOpenClient n, tbCount, sql

                      'Check Semen
1320                  If SysOptDeptSemen(0) Then
1330                      If !SampleID > SysOptSemenOffset(0) And !SampleID < SysOptMicroOffset(0) Then
1340                          g.TextMatrix(g.Row, 13) = Format$(Val(g.TextMatrix(g.Row, 13)) - SysOptSemenOffset(0))
1350                          g.Col = 9
1360                          g.CellBackColor = vbRed
1370                          If SemenComment Then
1380                              g = "C"
1390                          End If
1400                      End If
1410                  End If

                      'Check Histology
1420                  If SysOptDeptHisto(0) Then
1430                      If tb!SampleID < SysOptCytoOffset(0) And tb!SampleID > SysOptHistoOffset(0) Then
1440                          g.TextMatrix(g.Row, 26) = g.TextMatrix(g.Row, 13)
1450                          g.TextMatrix(g.Row, 27) = Trim$(tb!Hyear & "")
1460                          g.TextMatrix(g.Row, 13) = Trim(tb!Hyear) & "/" & Format$(Val(g.TextMatrix(g.Row, 13)) - (SysOptHistoOffset(0) + (Val(Swap_Year(Trim(tb!Hyear))) * 1000))) & "/H"
1470                          g.Col = 10
1480                          g.CellBackColor = vbRed
1490                          If HistologyComment Then
1500                              g = "C"
1510                          End If
1520                      End If
1530                  End If

                      'Check Cytology
1540                  If SysOptDeptCyto(0) Then
1550                      If tb!SampleID < 49999999 And tb!SampleID > SysOptCytoOffset(0) Then
1560                          g.TextMatrix(g.Row, 26) = g.TextMatrix(g.Row, 13)
1570                          g.TextMatrix(g.Row, 27) = tb!Hyear & ""
1580                          g.TextMatrix(g.Row, 13) = Trim(tb!Hyear & "") & "/" & Format$(Val(g.TextMatrix(g.Row, 13)) - (SysOptCytoOffset(0) + (Val(Swap_Year(Trim(tb!Hyear & ""))) * 1000))) & "/C"
1590                          g.Col = 11
1600                          g.CellBackColor = vbRed
1610                          If CytologyComment Then
1620                              g = "C"
1630                          End If
1640                      End If
1650                  End If

                      'Check Coagulation
1660                  If tbCount!CoagCount > 0 Then
1670                      g.Col = 2
1680                      g.CellBackColor = vbRed
1690                      If CoagulationComment Then
1700                          g = "C"
1710                      End If
1720                  ElseIf CoagulationComment Then
1730                      g.Col = 2
1740                      g.CellBackColor = vbRed
1750                      g = "C"
1760                  End If

                      'Check Haematology
1770                  If tbCount!HaemCount > 0 Then
1780                      g.Col = 1
1790                      g.CellBackColor = vbRed
1800                      If HaematologyComment Then
1810                          g = "C"
1820                      End If
1830                  ElseIf HaematologyComment Then
1840                      g.Col = 1
                          'g.CellBackColor = vbRed
1850                      g = "C"
1860                  End If

                      'Check Biochemistry
1870                  If tbCount!BioCount > 0 Then
1880                      g.Col = 0
1890                      g.CellBackColor = vbRed
1900                      If BiochemistryComment Then
1910                          g = "C"
1920                      End If
1930                  ElseIf BiochemistryComment Then
1940                      g.Col = 0
                          'g.CellBackColor = vbRed
1950                      g = "C"
1960                  End If

                      'Check Endocrinology
1970                  If SysOptDeptEnd(0) Then
1980                      If tbCount!EndCount > 0 Then
1990                          g.Col = 3
2000                          g.CellBackColor = vbRed
2010                          If EndocrinologyComment Then
2020                              g = "C"
2030                          End If
2040                      ElseIf EndocrinologyComment Then
2050                          g.Col = 3
                              'g.CellBackColor = vbRed
2060                          g = "C"
2070                      End If
2080                  End If

                      'Check Immunology
2090                  If SysOptDeptImm(0) Then
2100                      If tbCount!ImmCount > 0 Then
2110                          g.Col = 4
2120                          g.CellBackColor = vbRed
2130                          If ImmunologyComment Then
2140                              g = "C"
2150                          End If
2160                      ElseIf ImmunologyComment Then
2170                          g.Col = 4
                              'g.CellBackColor = vbRed
2180                          g = "C"
2190                      End If
2200                  End If

                      'Check Blood Gas
2210                  If SysOptDeptBga(0) Then
2220                      If tbCount!BGACount > 0 Then
2230                          g.Col = 5
2240                          g.CellBackColor = vbRed
2250                          If BloodGasComment Then
2260                              g = "C"
2270                          End If
2280                      ElseIf BloodGasComment Then
2290                          g.Col = 5
                              'g.CellBackColor = vbRed
2300                          g = "C"
2310                      End If
2320                  End If

                      'Check Externals
2330                  If SysOptDeptExt(0) Then
2340                      If tbCount!ExtCount > 0 Then
2350                          g.Col = 6
2360                          g.CellBackColor = vbRed
2370                      End If
2380                  End If

                      'Check Microbiology
'2390                  If SysOptDeptMicro(0) Then
'2400                      If Val(!SampleID & "") > SysOptMicroOffset(0) Then    'And Val(!SampleID) < SysOptHistoOffset(0) Then
'2410                          g.TextMatrix(g.Row, 26) = g.TextMatrix(g.Row, 13)
'2420                          g.TextMatrix(g.Row, 13) = Format$(Val(g.TextMatrix(g.Row, 13)) - SysOptMicroOffset(0))
'2430                          g.Col = 8
'2440                          g.CellBackColor = vbRed
'2450                          If MicroGeneralComment Then
'2460                              g = "C"
'2470                          End If
'2480                      End If
'2490                  End If
                       'Zyam changed the conditon on which micro should be be higlighteed 12-22-23
2401                   If Mid(CStr(!SampleID), 1, 1) = 2 Then  'And Val(!SampleID) < SysOptHistoOffset(0) Then
2412                       g.TextMatrix(g.Row, 26) = g.TextMatrix(g.Row, 13)
2423                       g.TextMatrix(g.Row, 13) = Mid(CStr(!SampleID), 2)
2434                       g.Col = 8
2445                       g.CellBackColor = vbRed
2456                       If MicroGeneralComment Then
2467                          g = "C"
2478                       End If
2490                   End If
                        'Zyam
2500                  .MoveNext
2510              Loop
2520          Else
2530              Do While Not .EOF
2540                  s = !Chart & vbTab & _
                          initial2upper(!PatName) & vbTab & _
                          Format$(!Dob, "dd/mm/yyyy") & vbTab & _
                          !sex & vbTab & _
                          initial2upper(!Address0 & "") & vbTab & _
                          initial2upper(!Address1 & "") & vbTab & _
                          initial2upper(!Ward & "") & vbTab & _
                          initial2upper(!Clinician & "") & vbTab & HospName(Cn)
2550                  g.AddItem s
2560                  .MoveNext
2570              Loop
2580          End If
2590      End With
2600  Next

2610  If NoPrevious Then
2620      lNoPrevious.Visible = True
2630  End If

2640  Exit Sub

LocalFillG_Error:

      Dim strES As String
      Dim intEL As Integer

2650  intEL = Erl
2660  strES = Err.Description
2670  LogError "frmPatHistoryNew", "LocalFillG", intEL, strES, sql

End Sub

Public Property Get NoPreviousDetails() As Variant

10        On Error GoTo NoPreviousDetails_Error

20        NoPreviousDetails = NoPrevious

30        Exit Property

NoPreviousDetails_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmPatHistoryNew", "NoPreviousDetails", intEL, strES


End Property

Private Sub bcancel_Click()

10        Unload Me

End Sub

Private Sub bCopy_Click()

          Dim gRow As Long
          Dim strWard As String
          Dim strGp As String
          Dim strSex As String
          Dim strName As String

10        On Error GoTo bCopy_Click_Error

20        gRow = g.Row

30        With mEditScreen
40            If oHD(1) Then
50                .txtAandE = g.TextMatrix(g.Row, 14)
60                If GetOptionSetting("EnableiPMSChart", "0") = 0 Then
70                .txtChart = g.TextMatrix(gRow, 15)
80                End If
90                strName = initial2upper(g.TextMatrix(gRow, 16))
100               .txtName = strName
110               .txtDoB = g.TextMatrix(gRow, 17)
120               .txtAge = CalcAge(.txtDoB, .dtSampleDate)
130               strSex = g.TextMatrix(gRow, 19)
140               If strSex = "" Then
150                   NameLostFocus strName, strSex
160               End If
170               If strSex = "M" Then
180                   .txtSex = "Male"
190               ElseIf strSex = "F" Then
200                   .txtSex = "Female"
210               End If
220               .taddress(0) = initial2upper(g.TextMatrix(gRow, 20))
230               .taddress(1) = initial2upper(g.TextMatrix(gRow, 21))
240               If SysOptDemo(0) = True Then
250                   strWard = initial2upper(g.TextMatrix(gRow, 22))
260                   strGp = initial2upper(g.TextMatrix(gRow, 24))
270                   If strWard = "" And strGp <> "" Then
280                       strWard = "GP"
290                   End If
300                   .cmbWard = strWard
310                   .cmbGP = strGp
320                   .cmbClinician = initial2upper(g.TextMatrix(gRow, 23))
330               End If
340           Else
350               .txtChart = g.TextMatrix(gRow, 0)
360               .txtName = initial2upper(g.TextMatrix(gRow, 1))
370               .txtDoB = g.TextMatrix(gRow, 2)
380               .txtAge = CalcAge(.txtDoB, .dtSampleDate)
390               .txtSex = g.TextMatrix(gRow, 3)
400               .taddress(0) = initial2upper(g.TextMatrix(gRow, 4))
410               .taddress(1) = initial2upper(g.TextMatrix(gRow, 5))
420               If SysOptDemo(0) = True Then
430                   .cmbWard = initial2upper(g.TextMatrix(gRow, 6))
440                   .cmbClinician = initial2upper(g.TextMatrix(gRow, 7))
450               End If
460           End If
470       End With

480       Unload Me

490       Exit Sub

bCopy_Click_Error:

          Dim strES As String
          Dim intEL As Integer

500       intEL = Erl
510       strES = Err.Description
520       LogError "frmPatHistoryNew", "bCopy_Click", intEL, strES

End Sub

Private Sub bsearch_Click()

10        On Error GoTo bsearch_Click_Error

20        FillG

30        Exit Sub

bsearch_Click_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmPatHistoryNew", "bsearch_Click", intEL, strES

End Sub

Private Sub chkSoundex_Click()

10        On Error GoTo chkSoundex_Click_Error

20        If chkSoundex = 1 Then
30            fraSearch.Visible = False
40        Else
50            fraSearch.Visible = True
60        End If

70        FillG

80        Exit Sub

chkSoundex_Click_Error:

          Dim strES As String
          Dim intEL As Integer



90        intEL = Erl
100       strES = Err.Description
110       LogError "frmPatHistoryNew", "chkSoundex_Click", intEL, strES


End Sub

Private Sub Form_Activate()

10        On Error GoTo Form_Activate_Error

20        If Activated Then Exit Sub

30        If TestSys = True Then Me.Caption = Me.Caption & " - TEST SYSTEM"

40        Set_Font Me

50        Activated = True

60        txtName.SetFocus

70        Exit Sub

Form_Activate_Error:

          Dim strES As String
          Dim intEL As Integer

80        intEL = Erl
90        strES = Err.Description
100       LogError "frmPatHistoryNew", "Form_Activate", intEL, strES

End Sub

Private Sub Form_Load()

10        On Error GoTo Form_Load_Error

20        Activated = False

30        LoadHeading IIf(oHD(0), 0, 1)

40        Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmPatHistoryNew", "Form_Load", intEL, strES

End Sub


Private Sub LoadHeading(ByVal Index As Integer)

          Dim n As Long

10        On Error GoTo LoadHeading_Error

20        For n = 0 To 8
30            g.ColWidth(n) = 250
40        Next

50        If Index = 0 Then

60            g.Cols = 9
70            g.FormatString = "<Chart     |<Name                         |<D.o.B.          " & _
                               "|<Sex|<Address                     |<                    " & _
                               "|<Ward                 |<Clinician              |<Hospital         "
80            g.Col = 2
90            g.Row = 0
100           txtDoB.Left = g.Left + g.CellLeft
110           txtDoB.Width = g.CellWidth

120       Else
130           g.Cols = 28
140           g.FormatString = "B|H|C|E|I|G|X|R|M|S|HI|CY|<Sample Date |<Run #      |<A + E           |<Chart      |" & _
                               "<Name                             |<D.o.B.           |<Age|^Sex|" & _
                               "<Address        |<  |<Ward           |<Clinician           |<GP                |<Hospital   |Samp"

150           g.ColWidth(26) = 0
160           g.ColWidth(27) = 0
170           If Not SysOptDeptExt(0) Then
180               g.ColWidth(6) = 0
190           End If
200           If Not SysOptDeptMedibridge(0) Then
210               g.ColWidth(7) = 0
220           End If
230           If Not SysOptDeptMicro(0) Then
240               g.ColWidth(8) = 0
250           End If
260           If Not SysOptDeptSemen(0) Then
270               g.ColWidth(9) = 0
280           End If
290           If Not SysOptDeptHisto(0) Then
300               g.ColWidth(10) = 0
310           End If
320           If Not SysOptDeptCyto(0) Then
330               g.ColWidth(11) = 0
340           End If
350           If Not SysOptDeptBga(0) Then
360               g.ColWidth(5) = 0
370           End If
380           If Not SysOptDeptEnd(0) Then
390               g.ColWidth(3) = 0
400           End If
410           If Not SysOptDeptImm(0) Then
420               g.ColWidth(4) = 0
430           End If

440           g.Col = 17
450           g.Row = 0
460           txtDoB.Left = g.Left + g.CellLeft
470           txtDoB.Width = g.CellWidth

480       End If

490       Exit Sub

LoadHeading_Error:

          Dim strES As String
          Dim intEL As Integer

500       intEL = Erl
510       strES = Err.Description
520       LogError "frmPatHistoryNew", "LoadHeading", intEL, strES

End Sub

Private Sub Form_Unload(Cancel As Integer)

10        On Error GoTo Form_Unload_Error

20        Activated = False

30        Exit Sub

Form_Unload_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmPatHistoryNew", "Form_Unload", intEL, strES


End Sub


Private Sub g_Click()

          Dim tb As New Recordset
          Dim sql As String
          Dim NewChart As String
          Dim PatName As String
          Dim Dob As String

10        On Error GoTo g_Click_Error

20        If g.MouseRow = 0 Then Exit Sub

30        If oHD(0) Then
40            If g.Col > 1 And mFromEdit Then
50                g.Col = 0
60                g.ColSel = g.Cols - 1
70                g.RowSel = g.Row
80                g.HighLight = flexHighlightAlways
90                bcopy.Enabled = True
100           ElseIf g.Col = 0 Then
110               If Trim$(g.TextMatrix(g.Row, 0)) = "" Then Exit Sub
120               PatName = g.TextMatrix(g.Row, 1)
130               If Trim$(PatName) = "" Then Exit Sub
140               Dob = g.TextMatrix(g.Row, 2)
150               If IsDate(Dob) Then
160                   Dob = Format$(Dob, "dd/mmm/yyyy")
170               Else
180                   Exit Sub
190               End If
200               If iMsg("Do you want to change the Chart Number?", vbQuestion + vbYesNo) = vbYes Then
210                   NewChart = iBOX("New Chart Number", , g.TextMatrix(g.Row, 0))
220                   sql = "UPDATE Demographics " & _
                            "set Chart = '" & NewChart & "' WHERE " & _
                            "PatName = '" & AddTicks(PatName) & "' " & _
                            "and dob = '" & Dob & "'"
230                   Set tb = New Recordset
240                   RecOpenClient 0, tb, sql
250                   FillG
260               End If
270           End If
280       Else
290           If g.Col = 15 Then    'chart
300               If Trim$(g.TextMatrix(g.Row, 15)) = "" Then Exit Sub
310               PatName = g.TextMatrix(g.Row, 16)
320               If Trim$(PatName) = "" Then Exit Sub
330               Dob = g.TextMatrix(g.Row, 15)
340               If IsDate(Dob) Then
350                   Dob = Format$(Dob, "dd/mmm/yyyy")
360               Else
370                   Exit Sub
380               End If
390               If iMsg("Do you want to change the Chart Number?", vbQuestion + vbYesNo) = vbYes Then
400                   NewChart = iBOX("New Chart Number", , g.TextMatrix(g.Row, 15))
410                   sql = "UPDATE Demographics " & _
                            "set Chart = '" & NewChart & "' WHERE " & _
                            "PatName = '" & AddTicks(PatName) & "' " & _
                            "and dob = '" & Dob & "'"
420                   Set tb = New Recordset
430                   RecOpenClient 0, tb, sql
440                   FillG
450               End If
460           ElseIf g.Col > 11 Then
470               g.Col = 12
480               g.ColSel = g.Cols - 1
490               g.RowSel = g.Row
500               g.HighLight = flexHighlightAlways
510               If mFromEdit Then
520                   If mDemoVal = True Then
530                       bcopy.Enabled = False
540                   Else
550                       bcopy.Enabled = True
560                   End If
570               End If
580           Else
590               bcopy.Enabled = False

600               If g.CellBackColor <> vbRed And g.TextMatrix(g.Row, g.Col) <> "C" Then Exit Sub

610               If g.Col = 7 Then
620                   With frmViewMedibridge
630                       .SampleID = g.TextMatrix(g.Row, 13)
640                       .Show 1
650                   End With
660               ElseIf g.Col = 8 Then    'Micro
                      'do not display form histolookup users    QMS Ref:#817961
670                   If UCase(UserMemberOf) = "HISTOLOOKUP" Then
680                       iMsg "You are not allowed to view micro results"
690                   Else
700                       With frmMicroReport
710                           .PatChart = g.TextMatrix(g.Row, 15)
720                           .PatName = g.TextMatrix(g.Row, 16)
730                           .PatDoB = g.TextMatrix(g.Row, 17)
740                           .PatSex = g.TextMatrix(g.Row, 19)
750                           .Show 1
760                       End With
770                   End If
780               ElseIf g.Col = 10 Then    'Histo
790                   If Val(g.TextMatrix(g.Row, 13)) <> 0 Then
800                       With frmHiCyReport
810                           .Load_Report "H", g.TextMatrix(g.Row, 26), g.TextMatrix(g.Row, 27)
820                           .lblSampleID = g.TextMatrix(g.Row, 13)
830                           .lblChart = g.TextMatrix(g.Row, 15)
840                           .lblName = g.TextMatrix(g.Row, 16)
850                           .lblDoB = g.TextMatrix(g.Row, 17)
860                           .Show 1
870                       End With
880                   End If
890               ElseIf g.Col = 11 Then    'Cyto
900                   If Val(g.TextMatrix(g.Row, 13)) <> 0 Then
910                       With frmHiCyReport
920                           .lblSampleID = g.TextMatrix(g.Row, 13)
930                           .lblChart = g.TextMatrix(g.Row, 15)
940                           .lblName = g.TextMatrix(g.Row, 16)
950                           .lblDoB = g.TextMatrix(g.Row, 17)
960                           .SampleID = g.TextMatrix(g.Row, 26)
970                           .Load_Report "C", g.TextMatrix(g.Row, 26), g.TextMatrix(g.Row, 27)
980                           .Show 1
990                       End With
1000                  End If
1010              ElseIf g.Col = 9 Then    'Semen Analysis
1020                  With frmSemenReport
1030                      .PatChart = g.TextMatrix(g.Row, 15)
1040                      .PatName = g.TextMatrix(g.Row, 16)
1050                      .PatDoB = g.TextMatrix(g.Row, 17)
1060                      .PatSex = g.TextMatrix(g.Row, 19)
1070                      .Show 1
1080                  End With
                      '            With frmSemenHistory
                      '                .lblName = g.TextMatrix(g.Row, 16)
                      '                .Show 1
                      '            End With
1090              ElseIf g.Col = 6 Then    'Ext
1100                  With frmFullExt
1110                      .lblChart = g.TextMatrix(g.Row, 15)
1120                      .lblName = g.TextMatrix(g.Row, 16)
1130                      .lblDoB = g.TextMatrix(g.Row, 17)
1140                      .Show 1
1150                  End With
1160              Else
1170                  With frmViewResults

1180                      .lblSampleID = g.TextMatrix(g.Row, 13)
1190                      .lblChart = g.TextMatrix(g.Row, 15)
1200                      .lblName = g.TextMatrix(g.Row, 16)
1210                      .lblDoB = g.TextMatrix(g.Row, 17)
1220                      .lblHosp = g.TextMatrix(g.Row, 25)
                          
1230                      .Show 1
1240                  End With
1250              End If
1260          End If
1270      End If

1280      Exit Sub

g_Click_Error:

          Dim strES As String
          Dim intEL As Integer

1290      intEL = Erl
1300      strES = Err.Description
1310      LogError "frmPatHistoryNew", "g_Click", intEL, strES, sql

End Sub

Private Sub ofor_Click(Index As Integer)

          Dim f As Form

10        On Error GoTo ofor_Click_Error

20        Select Case Index

          Case 0: optLeading = True
30            chkSoundex.Enabled = True
40            txtDoB.Visible = False

50        Case 1, 2: optExact = True
60            chkSoundex.Enabled = False
70            chkSoundex = 0
80            txtDoB.Visible = False

90        Case 3: optLeading = True
100           chkSoundex.Enabled = True

110           Set f = New frmGetDoB
120           f.Show 1
130           txtDoB = f.txtDoB
140           If f.lblWithin.Enabled Then
150               pWithin = f.lblWithin
160           Else
170               pWithin = 0
180           End If
190           Unload f
200           Set f = Nothing

210           If Not IsDate(txtDoB) Then
220               oFor(0).Value = True
230               Exit Sub
240           End If

250           txtDoB.Visible = True

260           If Me.Visible Then txtName.SetFocus

270       End Select


280       g.Rows = 2
290       g.AddItem ""
300       g.RemoveItem 1
310       txtName = ""

320       Exit Sub

ofor_Click_Error:

          Dim strES As String
          Dim intEL As Integer



330       intEL = Erl
340       strES = Err.Description
350       LogError "frmPatHistoryNew", "ofor_Click", intEL, strES


End Sub

Private Sub oHD_Click(Index As Integer)

10        On Error GoTo oHD_Click_Error

20        g.Rows = 2
30        g.AddItem ""
40        g.RemoveItem 1

50        LoadHeading Index

60        If Not Activated Then Exit Sub

70        FillG

80        Exit Sub

oHD_Click_Error:

          Dim strES As String
          Dim intEL As Integer

90        intEL = Erl
100       strES = Err.Description
110       LogError "frmPatHistoryNew", "oHD_Click", intEL, strES

End Sub

Private Sub txtName_KeyUp(KeyCode As Integer, Shift As Integer)

10        On Error GoTo txtName_KeyUp_Error

20        If chkType.Value = 1 Then
30            If oFor(0) Or oFor(3) Then
40                If Len(Trim$(txtName)) > 3 Then
50                    FillG
60                End If
70            End If
80        End If

90        Exit Sub

txtName_KeyUp_Error:

          Dim strES As String
          Dim intEL As Integer



100       intEL = Erl
110       strES = Err.Description
120       LogError "frmPatHistoryNew", "txtName_KeyUp", intEL, strES


End Sub





Private Sub udRecords_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

10        On Error GoTo udRecords_MouseUp_Error

20        If txtRecords = "9999" Then txtRecords = "20"

30        If Trim(txtName) = "" Then Exit Sub

40        FillG

50        Exit Sub

udRecords_MouseUp_Error:

          Dim strES As String
          Dim intEL As Integer



60        intEL = Erl
70        strES = Err.Description
80        LogError "frmPatHistoryNew", "udRecords_MouseUp", intEL, strES


End Sub



Public Property Let FromEdit(ByVal x As Boolean)

10        On Error GoTo FromEdit_Error

20        mFromEdit = x

30        Exit Property

FromEdit_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmPatHistoryNew", "FromEdit", intEL, strES

End Property

Public Property Let FromLookUp(ByVal bNewValue As Boolean)

10        On Error GoTo FromLookUp_Error

20        mFromLookup = bNewValue

30        Exit Property

FromLookUp_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmPatHistoryNew", "FromLookUp", intEL, strES

End Property


