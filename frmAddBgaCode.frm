VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmAddBgaCode 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire 6.9 - Add Blood Gas Analyte"
   ClientHeight    =   5700
   ClientLeft      =   510
   ClientTop       =   2715
   ClientWidth     =   6570
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
   Icon            =   "frmAddBgaCode.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5700
   ScaleWidth      =   6570
   Begin VB.Frame Frame1 
      Height          =   3300
      Left            =   90
      TabIndex        =   8
      Top             =   135
      Width           =   3255
      Begin VB.ComboBox cmbSampleType 
         Height          =   315
         Left            =   1095
         TabIndex        =   0
         Top             =   360
         Width           =   1965
      End
      Begin VB.ComboBox cmbUnits 
         Height          =   315
         Left            =   1095
         TabIndex        =   4
         Top             =   2340
         Width           =   1965
      End
      Begin VB.TextBox txtCode 
         Height          =   285
         Left            =   1095
         MaxLength       =   7
         TabIndex        =   1
         Top             =   870
         Width           =   1965
      End
      Begin VB.TextBox txtShortName 
         Height          =   285
         Left            =   1095
         MaxLength       =   6
         TabIndex        =   2
         Top             =   1410
         Width           =   1965
      End
      Begin VB.TextBox txtLongName 
         Height          =   285
         Left            =   1095
         MaxLength       =   40
         TabIndex        =   3
         Top             =   1890
         Width           =   1965
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "SampleType"
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
         Left            =   45
         TabIndex        =   13
         Top             =   420
         Width           =   885
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Units"
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
         Left            =   570
         TabIndex        =   12
         Top             =   2400
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Code"
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
         Left            =   555
         TabIndex        =   11
         Top             =   900
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Long Name"
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
         Left            =   105
         TabIndex        =   10
         Top             =   1920
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Short Name"
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
         TabIndex        =   9
         Top             =   1440
         Width           =   840
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdTest 
      Height          =   5415
      Left            =   3465
      TabIndex        =   7
      Top             =   135
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   9551
      _Version        =   393216
      FixedCols       =   0
      FormatString    =   "Name                       |Code    "
   End
   Begin VB.CommandButton cmdSave 
      Appearance      =   0  'Flat
      Caption         =   "&Save"
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
      Left            =   765
      Picture         =   "frmAddBgaCode.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3600
      Width           =   1395
   End
   Begin VB.CommandButton cmdCancel 
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
      Height          =   750
      Left            =   765
      Picture         =   "frmAddBgaCode.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4545
      Width           =   1395
   End
End
Attribute VB_Name = "frmAddBgaCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

10        Unload Me

End Sub

Private Sub cmdSave_Click()

          Dim tb As New Recordset
          Dim sql As String
          Dim SampleType As String



10        On Error GoTo cmdSave_Click_Error

20        If Trim$(cmbSampleType) = "" Then
30            iMsg "SELECT Sample Type.", vbCritical
40            Exit Sub
50        End If

60        If Trim$(txtCode) = "" Then
70            iMsg "Enter Code.", vbCritical
80            Exit Sub
90        End If

100       If Trim$(txtShortName) = "" Then
110           iMsg "Enter Short Name.", vbCritical
120           Exit Sub
130       End If

140       If Trim$(txtLongName) = "" Then
150           iMsg "Enter Long Name.", vbCritical
160           Exit Sub
170       End If

180       If Trim$(cmbUnits) = "" Then
190           If iMsg("No SELECT Units. Ok ?", vbYesNo) = vbNo Then
200               Exit Sub
210           End If
220       End If

230       SampleType = ListCodeFor("ST", cmbSampleType)

240       sql = "SELECT * from BgaTestDefinitions WHERE " & _
                "Code = '" & Trim$(txtCode) & "'"
250       Set tb = New Recordset
260       RecOpenServer 0, tb, sql
270       If Not tb.EOF Then
280           iMsg "Code already used.", vbCritical
290           Exit Sub
300       End If

310       With tb
320           .AddNew
330           !Code = UCase(txtCode)
340           !ShortName = txtShortName
350           !LongName = txtLongName
360           !DoDelta = False
370           !DeltaLimit = 0
380           !PrintPriority = 999
390           !DP = 1
400           !BarCode = ""
410           !Units = cmbUnits
420           !h = False
430           !s = False
440           !l = False
450           !o = False
460           !g = False
470           !J = False
480           !MaleLow = 0
490           !MaleHigh = 999
500           !FemaleLow = 0
510           !FemaleHigh = 999
520           !FlagMaleLow = 0
530           !FlagMaleHigh = 999
540           !FlagFemaleLow = 0
550           !FlagFemaleHigh = 999
560           !SampleType = SampleType
570           !LControlLow = 0
580           !LControlHigh = 999
590           !NControlLow = 0
600           !NControlHigh = 999
610           !HControlLow = 0
620           !HControlHigh = 999
630           !Printable = False
640           !PlausibleLow = 0
650           !PlausibleHigh = 9999
660           !InUse = True
670           !AgeFromDays = 0
680           !AgeToDays = MaxAgeToDays
690           !Hospital = HospName(0)
700           !KnownToAnalyser = 1
710           !Category = ""
720           .Update

730       End With

740       txtCode = ""
750       txtShortName = ""
760       txtLongName = ""

770       FillTest



780       Exit Sub

cmdSave_Click_Error:

          Dim strES As String
          Dim intEL As Integer

790       intEL = Erl
800       strES = Err.Description
810       LogError "frmAddBgaCode", "cmdsave_Click", intEL, strES, sql


End Sub

Private Sub FillLists()

          Dim tb As New Recordset
          Dim sql As String


10        On Error GoTo FillLists_Error

20        cmbUnits.Clear
30        cmbSampleType.Clear

40        sql = "SELECT * from Lists WHERE " & _
                "ListType = 'ST' or ListType = 'UN' " & _
                "Order by ListOrder"
50        Set tb = New Recordset
60        RecOpenServer 0, tb, sql
70        Do While Not tb.EOF
80            Select Case tb!ListType & ""
              Case "ST": cmbSampleType.AddItem tb!Text & ""
90            Case "UN": cmbUnits.AddItem tb!Text & ""
100           End Select
110           tb.MoveNext
120       Loop






130       Exit Sub

FillLists_Error:

          Dim strES As String
          Dim intEL As Integer

140       intEL = Erl
150       strES = Err.Description
160       LogError "frmAddBgaCode", "FillLists", intEL, strES, sql


End Sub

Private Sub FillTest()
          Dim sql As String
          Dim tb As New Recordset


10        On Error GoTo FillTest_Error

20        ClearFGrid grdTest

30        sql = "SELECT distinct(longname), code, inuse from bgatestdefinitions"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        Do While Not tb.EOF
70            grdTest.AddItem Trim(tb!LongName) & vbTab & tb!Code
80            grdTest.Row = grdTest.Rows - 1
90            If tb!InUse = False Then
100               grdTest.Col = 0
110               grdTest.CellBackColor = vbGreen
120               grdTest.Col = 1
130               grdTest.CellBackColor = vbGreen
140           Else
150               grdTest.Col = 0
160               grdTest.CellBackColor = vbWhite
170               grdTest.Col = 1
180               grdTest.CellBackColor = vbWhite
190           End If
200           tb.MoveNext
210       Loop

220       FixG grdTest




230       Exit Sub

FillTest_Error:

          Dim strES As String
          Dim intEL As Integer

240       intEL = Erl
250       strES = Err.Description
260       LogError "frmAddBgaCode", "FillTest", intEL, strES, sql


End Sub

Private Sub Form_Load()

10        On Error GoTo Form_Load_Error

20        FillLists

30        FillTest

40        Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmAddBgaCode", "Form_Load", intEL, strES


End Sub
