VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmAddNewTest 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Add / Edit Endocrinology Test"
   ClientHeight    =   3960
   ClientLeft      =   1665
   ClientTop       =   1455
   ClientWidth     =   10845
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
   Icon            =   "frmEndAddCode.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3960
   ScaleWidth      =   10845
   Begin VB.TextBox txtHost 
      Height          =   285
      Left            =   1620
      MaxLength       =   10
      TabIndex        =   19
      Top             =   1575
      Width           =   1965
   End
   Begin MSFlexGridLib.MSFlexGrid grdEnd 
      Height          =   3570
      Left            =   7560
      TabIndex        =   18
      Top             =   180
      Width           =   3120
      _ExtentX        =   5503
      _ExtentY        =   6297
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      AllowUserResizing=   1
      FormatString    =   "<Test Code  |< Test Name          "
   End
   Begin VB.Frame Frame1 
      Caption         =   "Analysers"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1140
      Left            =   4590
      TabIndex        =   14
      Top             =   180
      Width           =   2670
      Begin VB.OptionButton optAnal 
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
         Left            =   225
         TabIndex        =   17
         Top             =   825
         Width           =   2085
      End
      Begin VB.OptionButton optAnal 
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
         Left            =   225
         TabIndex        =   16
         Top             =   540
         Width           =   2085
      End
      Begin VB.OptionButton optAnal 
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
         Left            =   225
         TabIndex        =   15
         Top             =   270
         Value           =   -1  'True
         Width           =   2085
      End
   End
   Begin VB.ComboBox cCat 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1620
      TabIndex        =   1
      Top             =   600
      Width           =   1965
   End
   Begin VB.ComboBox cSampleType 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1620
      TabIndex        =   0
      Top             =   150
      Width           =   1965
   End
   Begin VB.ComboBox cunits 
      Height          =   315
      Left            =   1620
      TabIndex        =   5
      Top             =   3165
      Width           =   1965
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
      Height          =   795
      Left            =   5760
      Picture         =   "frmEndAddCode.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2880
      Width           =   1275
   End
   Begin VB.CommandButton bsave 
      Appearance      =   0  'Flat
      Caption         =   "Save"
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
      Height          =   795
      Left            =   4140
      Picture         =   "frmEndAddCode.frx":1C8C
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2880
      Width           =   1305
   End
   Begin VB.TextBox tcode 
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
      Left            =   1620
      MaxLength       =   10
      TabIndex        =   2
      Top             =   1110
      Width           =   1965
   End
   Begin VB.TextBox tShortName 
      Height          =   285
      Left            =   1620
      MaxLength       =   6
      TabIndex        =   4
      Top             =   2700
      Width           =   825
   End
   Begin VB.TextBox tLongName 
      Height          =   285
      Left            =   1620
      MaxLength       =   40
      TabIndex        =   3
      Top             =   2160
      Width           =   5505
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Host Code"
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
      Left            =   765
      TabIndex        =   20
      Top             =   1620
      Width           =   750
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Category"
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
      Left            =   840
      TabIndex        =   13
      Top             =   660
      Width           =   630
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
      Left            =   570
      TabIndex        =   12
      Top             =   210
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
      Left            =   1095
      TabIndex        =   11
      Top             =   3225
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
      Index           =   0
      Left            =   1080
      TabIndex        =   10
      Top             =   1140
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
      Left            =   630
      TabIndex        =   9
      Top             =   2205
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
      Left            =   585
      TabIndex        =   8
      Top             =   2745
      Width           =   840
   End
End
Attribute VB_Name = "frmAddNewTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mDepartment As String

Private Sub bcancel_Click()

10        Unload Me

End Sub

Private Sub bSave_Click()

          Dim tb As New Recordset
          Dim sql As String
          Dim SampleType As String

10        On Error GoTo bSave_Click_Error

20        If Trim$(cSampleType) = "" Then
30            iMsg "SELECT Sample Type.", vbCritical
40            Exit Sub
50        End If

60        If Trim$(tCode) = "" Then
70            iMsg "Enter Code.", vbCritical
80            Exit Sub
90        End If

100       If Trim$(tShortName) = "" Then
110           iMsg "Enter Short Name.", vbCritical
120           Exit Sub
130       End If

140       If Trim$(tLongName) = "" Then
150           iMsg "Enter Long Name.", vbCritical
160           Exit Sub
170       End If

180       If Trim$(cUnits) = "" Then
190           If iMsg("No Units Selected." & vbCrLf & "Is this correct?", vbQuestion + vbYesNo) = vbNo Then
200               Exit Sub
210           End If
220       End If

230       SampleType = ListCodeFor("ST", cSampleType)

240       sql = "SELECT * FROM " & Left$(mDepartment, 3) & "TestDefinitions WHERE " & _
                "Code = '" & tCode & "' " & _
                "AND Category = '" & cCat & "'"
250       Set tb = New Recordset
260       RecOpenServer 0, tb, sql
270       If Not tb.EOF Then
280           iMsg "Code already used.", vbCritical
290           Exit Sub
300       Else
310           With tb
320               .AddNew
330               !Code = tCode
340               !ShortName = tShortName
350               !LongName = tLongName
360               !DoDelta = False
370               !DeltaLimit = 0
380               !PrintPriority = 999
390               !DP = 1
400               !BarCode = ""
410               !KnownToAnalyser = 1
420               !Units = cUnits
430               !h = False
440               !s = False
450               !l = False
460               !o = False
470               !g = False
480               !J = False
490               !MaleLow = 0
500               !MaleHigh = 999
510               !FemaleLow = 0
520               !FemaleHigh = 999
530               !FlagMaleLow = 0
540               !FlagMaleHigh = 999
550               !FlagFemaleLow = 0
560               !FlagFemaleHigh = 999
570               !SampleType = SampleType
580               !Category = cCat
590               !LControlLow = 0
600               !LControlHigh = 999
610               !NControlLow = 0
620               !NControlHigh = 999
630               !HControlLow = 0
640               !HControlHigh = 999
650               !Printable = 1
660               !PlausibleLow = 0
670               !PlausibleHigh = 9999
680               !InUse = True
690               !AgeFromDays = 0
700               !AgeToDays = MaxAgeToDays
710               If Trim(txtHost) <> "" Then !immunocode = txtHost Else !immunocode = tCode
720               !Hospital = HospName(0)
730               If optAnal(0) Then
740                   !Analyser = Left(optAnal(0).Caption, 1)
750               ElseIf optAnal(1) Then
760                   !Analyser = Left(optAnal(1).Caption, 1)
770               ElseIf optAnal(2) Then
780                   !Analyser = optAnal(2).Caption
790               End If
800               .Update
810           End With
820       End If

830       tCode = ""
840       tShortName = ""
850       tLongName = ""
860       txtHost = ""

870       Exit Sub

bSave_Click_Error:

          Dim strES As String
          Dim intEL As Integer

880       intEL = Erl
890       strES = Err.Description
900       LogError "frmAddNewTest", "bSave_Click", intEL, strES, sql

End Sub

Private Sub cCat_KeyPress(KeyAscii As Integer)

10        On Error GoTo cCat_KeyPress_Error

20        KeyAscii = 0

30        Exit Sub

cCat_KeyPress_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmAddNewTest", "cCat_KeyPress", intEL, strES

End Sub

Private Sub FillLists()

          Dim tb As New Recordset
          Dim sql As String
          Dim Y As Integer

10        On Error GoTo FillLists_Error

20        cUnits.Clear
30        cSampleType.Clear

40        sql = "SELECT ListType, Text FROM Lists WHERE " & _
                "ListType = 'ST' " & _
                "OR ListType = 'UN'"
50        Set tb = New Recordset
60        RecOpenServer 0, tb, sql

70        Do While Not tb.EOF
80            If Trim(tb!ListType) = "ST" Then
90                cSampleType.AddItem Trim(tb!Text)
100           ElseIf Trim(tb!ListType) = "UN" Then
110               cUnits.AddItem Trim(tb!Text)
120           End If
130           tb.MoveNext
140       Loop

150       For Y = 0 To cSampleType.ListCount - 1
160           If UCase$(cSampleType.List(Y)) = "SERUM" Then
170               cSampleType.ListIndex = Y
180               Exit For
190           End If
200       Next

210       cCat.Clear

220       sql = "SELECT * FROM Categorys"
230       Set tb = New Recordset
240       RecOpenServer 0, tb, sql
250       Do While Not tb.EOF
260           cCat.AddItem Trim(tb!Cat)
270           tb.MoveNext
280       Loop

290       For Y = 0 To cCat.ListCount - 1
300           If UCase$(cCat.List(Y)) = "DEFAULT" Then
310               cCat.ListIndex = Y
320               Exit For
330           End If
340       Next

350       Exit Sub

FillLists_Error:

          Dim strES As String
          Dim intEL As Integer

360       intEL = Erl
370       strES = Err.Description
380       LogError "frmAddNewTest", "FillLists", intEL, strES, sql

End Sub

Private Sub Form_Load()

10        On Error GoTo Form_Load_Error

20        Me.Caption = "NetAcquire - Add " & mDepartment & " Test"

30        FillLists
40        FillGrid

50        Select Case mDepartment
          Case "Endocrinology"
60            optAnal(0).Caption = GetOptionSetting("EndAn1", "")
70            optAnal(1).Caption = GetOptionSetting("EndAn2", "")
80            optAnal(2).Caption = GetOptionSetting("EndAn3", "")
90        Case "Immunology"
100           optAnal(0).Caption = GetOptionSetting("ImmAn1", "")
110           optAnal(1).Caption = GetOptionSetting("ImmAn2", "")
120           optAnal(2).Caption = GetOptionSetting("ImmAn3", "")
130       Case "Biochemistry"
140           optAnal(0).Caption = GetOptionSetting("BioAn1", "")
150           optAnal(1).Caption = GetOptionSetting("BioAn2", "")
160           optAnal(2).Caption = GetOptionSetting("BioAn3", "")
170       End Select

180       Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

190       intEL = Erl
200       strES = Err.Description
210       LogError "frmAddNewTest", "Form_Load", intEL, strES

End Sub

Private Sub FillGrid()

          Dim sql As String
          Dim tb As Recordset

10        On Error GoTo FillGrid_Error

20        With grdEnd
30            .Rows = 2
40            .AddItem ""
50            .RemoveItem 1
60        End With

70        sql = "SELECT DISTINCT Code, LongName " & _
                "FROM " & Left$(mDepartment, 3) & "TestDefinitions"
80        Set tb = New Recordset
90        RecOpenServer 0, tb, sql
100       Do While Not tb.EOF
110           grdEnd.AddItem tb!Code & vbTab & Trim(tb!LongName)
120           tb.MoveNext
130       Loop

140       If grdEnd.Rows > 2 And grdEnd.TextMatrix(1, 0) = "" Then grdEnd.RemoveItem 1

150       Exit Sub

FillGrid_Error:

          Dim strES As String
          Dim intEL As Integer

160       intEL = Erl
170       strES = Err.Description
180       LogError "frmAddNewTest", "FillGrid", intEL, strES, sql

End Sub


Public Property Let Department(ByVal strNewValue As String)

10        mDepartment = strNewValue

End Property
