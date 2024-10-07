VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmAges 
   Caption         =   "NetAcquire - Ages"
   ClientHeight    =   8040
   ClientLeft      =   2895
   ClientTop       =   750
   ClientWidth     =   4905
   Icon            =   "frmAges.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   4905
   Begin VB.ComboBox cmbMonth 
      Height          =   315
      Left            =   3060
      TabIndex        =   10
      Top             =   960
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.ComboBox cmbYear 
      Height          =   315
      Left            =   2310
      TabIndex        =   9
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox cmbDay 
      Height          =   315
      Left            =   3660
      TabIndex        =   8
      Top             =   960
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove Age Range"
      Height          =   735
      Left            =   2955
      Picture         =   "frmAges.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5775
      Width           =   1605
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   825
      Left            =   585
      Picture         =   "frmAges.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7020
      Width           =   1245
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add Age Range"
      Height          =   735
      Left            =   285
      Picture         =   "frmAges.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5775
      Width           =   1605
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H0000FFFF&
      Caption         =   "&Save"
      Height          =   735
      Left            =   1995
      Picture         =   "frmAges.frx":0C28
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5775
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   825
      Left            =   2925
      Picture         =   "frmAges.frx":0F32
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7020
      Width           =   1245
   End
   Begin MSFlexGridLib.MSFlexGrid grdAge 
      Height          =   5115
      Left            =   330
      TabIndex        =   0
      Top             =   570
      Width           =   4275
      _ExtentX        =   7541
      _ExtentY        =   9022
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   315
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      FormatString    =   "<Age From (YMD)       |<Age To (YMD)           "
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
   Begin VB.Label lblSampleType 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Serum"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   330
      TabIndex        =   7
      Top             =   120
      Width           =   1725
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000001&
      BorderWidth     =   2
      Height          =   75
      Left            =   285
      Top             =   6765
      Width           =   4335
   End
   Begin VB.Label lblParameter 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Creatinine"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2280
      TabIndex        =   3
      Top             =   120
      Width           =   2025
   End
End
Attribute VB_Name = "frmAges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mAnalyte As String
Private mSampleType As String
Private mDiscipline As String
Private mCat As String
Private Activated As Boolean
Private mAnalyser As String

Private FromDays() As Long
Private ToDays() As Long

Private Sub AddBga()

          Dim tb As New Recordset
          Dim tbNew As Recordset
          Dim sql As String
          Dim fld As Field


10        On Error GoTo AddBga_Error

20        sql = "SELECT top 1 * from BgaTestDefinitions WHERE " & _
                "LongName = '" & mAnalyte & "' " & _
                "and SampleType = '" & mSampleType & "' " & _
                "order by AgeToDays desc"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql

50        If tb.EOF Then Exit Sub

60        Set tbNew = New Recordset
70        RecOpenServer 0, tbNew, sql

80        tbNew.AddNew
90        For Each fld In tb.Fields
100           If fld.Name = "AgeToDays" Then
110               tbNew!AgeToDays = tb!AgeToDays + 1
120           ElseIf fld.Name = "AgeFromDays" Then
130               tbNew!AgeFromDays = tb!AgeToDays + 1
140           Else
150               tbNew(fld.Name) = tb(fld.Name)
160           End If
170       Next
180       tbNew.Update

190       FillBgaAges




200       Exit Sub

AddBga_Error:

          Dim strES As String
          Dim intEL As Integer

210       intEL = Erl
220       strES = Err.Description
230       LogError "frmAges", "AddBga", intEL, strES


End Sub

Private Sub AddBio()

          Dim tb As New Recordset
          Dim tbNew As Recordset
          Dim sql As String
          Dim fld As Field


10        On Error GoTo AddBio_Error

20        sql = "SELECT top 1 * from BioTestDefinitions WHERE " & _
                "LongName = '" & mAnalyte & "' " & _
                "and SampleType = '" & mSampleType & "' " & _
                "order by AgeToDays desc"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql

50        Set tbNew = New Recordset
60        RecOpenServer 0, tbNew, sql

70        tbNew.AddNew
80        For Each fld In tb.Fields
90            If fld.Name = "AgeToDays" Then
100               tbNew!AgeToDays = tb!AgeToDays + 1
110           ElseIf fld.Name = "AgeFromDays" Then
120               tbNew!AgeFromDays = tb!AgeToDays + 1
130           Else
140               If fld.Name <> "rowguid" Then
150                   tbNew(fld.Name) = tb(fld.Name)
160               End If
170           End If
180       Next
190       tbNew.Update

200       FillBioAges




210       Exit Sub

AddBio_Error:

          Dim strES As String
          Dim intEL As Integer

220       intEL = Erl
230       strES = Err.Description
240       LogError "frmAges", "AddBio", intEL, strES, sql


End Sub

Private Sub AddCoag()

          Dim tb As New Recordset
          Dim tbNew As Recordset
          Dim sql As String
          Dim fld As Field


10        On Error GoTo AddCoag_Error

20        sql = "SELECT top 1 * from CoagTestDefinitions WHERE " & _
                "TestName = '" & mAnalyte & "' " & _
                "order by AgeToDays desc"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql

50        Set tbNew = New Recordset
60        RecOpenServer 0, tbNew, sql

70        tbNew.AddNew
80        For Each fld In tb.Fields
90            If fld.Name = "AgeToDays" Then
100               tbNew!AgeToDays = tb!AgeToDays + 1
110           ElseIf fld.Name = "AgeFromDays" Then
120               tbNew!AgeFromDays = tb!AgeToDays + 1
130           Else
140               If fld.Name <> "rowguid" Then
150                   tbNew(fld.Name) = tb(fld.Name)
160               End If
170           End If
180       Next
190       tbNew.Update

200       FillCoagAges



210       Exit Sub

AddCoag_Error:

          Dim strES As String
          Dim intEL As Integer

220       intEL = Erl
230       strES = Err.Description
240       LogError "frmAges", "AddCoag", intEL, strES, sql


End Sub

Private Sub AddEnd()

          Dim tb As New Recordset
          Dim tbNew As Recordset
          Dim sql As String
          Dim fld As Field




10        On Error GoTo AddEnd_Error

20        sql = "SELECT top 1 * from EndTestDefinitions WHERE " & _
                "LongName = '" & mAnalyte & "' and category = '" & mCat & "' " & _
                "and SampleType = '" & mSampleType & "' " & _
                "order by AgeToDays desc"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql

50        Set tbNew = New Recordset
60        RecOpenServer 0, tbNew, sql

70        tbNew.AddNew
80        For Each fld In tb.Fields
90            If fld.Name = "AgeToDays" Then
100               tbNew!AgeToDays = tb!AgeToDays + 1
110           ElseIf fld.Name = "AgeFromDays" Then
120               tbNew!AgeFromDays = tb!AgeToDays + 1
130           Else
140               If fld.Name <> "rowguid" Then
150                   tbNew(fld.Name) = tb(fld.Name)
160               End If
170           End If
180       Next
190       tbNew.Update

200       FillEndAges




210       Exit Sub

AddEnd_Error:

          Dim strES As String
          Dim intEL As Integer

220       intEL = Erl
230       strES = Err.Description
240       LogError "frmAges", "AddEnd", intEL, strES


End Sub

Private Sub AddHaem()

          Dim tb As New Recordset
          Dim tbNew As Recordset
          Dim sql As String
          Dim fld As Field



10        On Error GoTo AddHaem_Error

20        sql = "SELECT top 1 * from HaemTestDefinitions WHERE " & _
                "AnalyteName = '" & mAnalyte & "' " & _
                "order by AgeToDays desc"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql

50        Set tbNew = New Recordset
60        RecOpenServer 0, tbNew, sql

70        tbNew.AddNew
80        For Each fld In tb.Fields
90            If fld.Name = "AgeToDays" Then
100               tbNew!AgeToDays = tb!AgeToDays + 1
110           ElseIf fld.Name = "AgeFromDays" Then
120               tbNew!AgeFromDays = tb!AgeToDays + 1
130           Else
140               If fld.Name <> "rowguid" Then
150                   tbNew(fld.Name) = tb(fld.Name)
160               End If
170           End If
180       Next
190       tbNew.Update

200       FillHaemAges




210       Exit Sub

AddHaem_Error:

          Dim strES As String
          Dim intEL As Integer

220       intEL = Erl
230       strES = Err.Description
240       LogError "frmAges", "AddHaem", intEL, strES


End Sub

Private Sub AddImm()

          Dim tb As New Recordset
          Dim tbNew As Recordset
          Dim sql As String
          Dim fld As Field



10        On Error GoTo AddImm_Error

20        sql = "SELECT top 1 * from ImmTestDefinitions WHERE " & _
                "LongName = '" & mAnalyte & "' " & _
                "and SampleType = '" & mSampleType & "' and analyser = '" & mAnalyser & "' " & _
                "order by AgeToDays desc"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql

50        If tb.EOF Then Exit Sub

60        Set tbNew = New Recordset
70        RecOpenServer 0, tbNew, sql

80        tbNew.AddNew
90        For Each fld In tb.Fields
100           If fld.Name = "AgeToDays" Then
110               tbNew!AgeToDays = tb!AgeToDays + 1
120           ElseIf fld.Name = "AgeFromDays" Then
130               tbNew!AgeFromDays = tb!AgeToDays + 1
140           Else
150               If fld.Name <> "rowguid" Then
160                   tbNew(fld.Name) = tb(fld.Name)
170               End If
180           End If
190       Next
200       tbNew.Update

210       FillImmAges



220       Exit Sub

AddImm_Error:

          Dim strES As String
          Dim intEL As Integer

230       intEL = Erl
240       strES = Err.Description
250       LogError "frmAges", "AddImm", intEL, strES, sql


End Sub

Private Sub AdjustG()

          Dim Num As Long



10        On Error GoTo AdjustG_Error

20        For Num = 0 To UBound(FromDays)
30            grdAge.TextMatrix(Num + 1, 0) = dmyFromCount(FromDays(Num))
40            grdAge.TextMatrix(Num + 1, 1) = dmyFromCount(ToDays(Num))
50        Next




60        Exit Sub

AdjustG_Error:

          Dim strES As String
          Dim intEL As Integer

70        intEL = Erl
80        strES = Err.Description
90        LogError "frmAges", "AdjustG", intEL, strES


End Sub

Public Property Let Analyte(ByVal Analyte As String)

10        On Error GoTo Analyte_Error

20        mAnalyte = Analyte

30        Exit Property

Analyte_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmAges", "Analyte", intEL, strES


End Property

Public Property Let Cat(ByVal Cat As String)

10        On Error GoTo Cat_Error

20        mCat = Cat

30        Exit Property

Cat_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmAges", "Cat", intEL, strES


End Property

Private Sub cmbDay_Click()

10        cmdSave.Visible = True

End Sub

Private Sub cmbDay_KeyPress(KeyAscii As Integer)

10        On Error GoTo cmbDay_KeyPress_Error

20        KeyAscii = 0

30        Exit Sub

cmbDay_KeyPress_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmAges", "cmbDay_KeyPress", intEL, strES


End Sub

Private Sub cmbMonth_Click()

10        On Error GoTo cmbMonth_Click_Error

20        cmdSave.Visible = True

30        Exit Sub

cmbMonth_Click_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmAges", "cmbMonth_Click", intEL, strES


End Sub

Private Sub cmbMonth_KeyPress(KeyAscii As Integer)

10        KeyAscii = 0

End Sub

Private Sub cmbYear_Click()

10        On Error GoTo cmbYear_Click_Error

20        If Val(cmbYear) > 10 Then
30            cmbMonth = "0"
40            cmbDay = "0"
50        End If

60        cmdSave.Visible = True

70        Exit Sub

cmbYear_Click_Error:

          Dim strES As String
          Dim intEL As Integer

80        intEL = Erl
90        strES = Err.Description
100       LogError "frmAges", "cmbYear_Click", intEL, strES


End Sub

Private Sub cmbYear_KeyPress(KeyAscii As Integer)

10        KeyAscii = 0

End Sub

Private Sub cmdadd_Click()

10        On Error GoTo cmdadd_Click_Error

20        Select Case mDiscipline
          Case "Biochemistry": AddBio
30        Case "Haematology": AddHaem
40        Case "Coagulation": AddCoag
50        Case "Immunology": AddImm
60        Case "Endocrinology": AddEnd
70        Case "Blood Gas": AddBga
80        End Select

90        Exit Sub

cmdadd_Click_Error:

          Dim strES As String
          Dim intEL As Integer

100       intEL = Erl
110       strES = Err.Description
120       LogError "frmAges", "cmdAdd_Click", intEL, strES


End Sub

Private Sub cmdCancel_Click()

10        Unload Me

End Sub

Private Sub cmdRemove_Click()



10        On Error GoTo cmdRemove_Click_Error

20        Select Case mDiscipline
          Case "Biochemistry": RemoveBio
30        Case "Haematology": RemoveHaem
40        Case "Coagulation": RemoveCoag
50        Case "Immunology": RemoveImm
60        Case "Endocrinology": RemoveEnd
70        Case "Blood Gas": RemoveBga
80        End Select

90        cmbYear.Visible = False
100       cmbMonth.Visible = False
110       cmbDay.Visible = False
120       cmdSave.Visible = False



130       Exit Sub

cmdRemove_Click_Error:

          Dim strES As String
          Dim intEL As Integer

140       intEL = Erl
150       strES = Err.Description
160       LogError "frmAges", "cmdRemove_Click", intEL, strES


End Sub

Private Sub cmdSave_Click()

10        On Error GoTo cmdSave_Click_Error

20        Select Case mDiscipline
          Case "Biochemistry": SaveBio
30        Case "Haematology": SaveHaem
40        Case "Coagulation": SaveCoag
50        Case "Immunology": SaveImm
60        Case "Endocrinology": SaveEnd
70        Case "Blood Gas": SaveBga
80        End Select


90        Exit Sub

cmdSave_Click_Error:

          Dim strES As String
          Dim intEL As Integer

100       intEL = Erl
110       strES = Err.Description
120       LogError "frmAges", "cmdsave_Click", intEL, strES


End Sub
Public Property Let Analyser(ByVal Analyser As String)

10        mAnalyser = Analyser

End Property
Public Property Let Discipline(ByVal Discipline As String)

10        mDiscipline = Discipline

End Property

Private Sub FillBgaAges()

          Dim tb As New Recordset
          Dim sql As String
          Dim Num As Long
          Dim Str As String



10        On Error GoTo FillBgaAges_Error

20        ClearFGrid grdAge

30        sql = "SELECT AgeFromDays, AgeToDays from BgaTestDefinitions WHERE " & _
                "LongName = '" & mAnalyte & "' " & _
                "and SampleType = '" & mSampleType & "' and category = '" & mCat & "'" & _
                "Order by AgeFromDays"
40        Set tb = New Recordset
50        RecOpenClient 0, tb, sql

60        ReDim FromDays(0 To tb.recordCount - 1)
70        ReDim ToDays(0 To tb.recordCount - 1)
80        Num = 0
90        Do While Not tb.EOF
100           If tb!AgeFromDays & "" = "" Then FromDays(Num) = 0 Else FromDays(Num) = tb!AgeFromDays
110           If tb!AgeToDays & "" = "" Then ToDays(Num) = 43870 Else ToDays(Num) = tb!AgeToDays
120           Str = dmyFromCount(FromDays(Num)) & vbTab & _
                    dmyFromCount(ToDays(Num))
130           grdAge.AddItem Str
140           Num = Num + 1
150           tb.MoveNext
160       Loop

170       FixG grdAge

180       cmdRemove.Enabled = grdAge.Rows > 2


190       Exit Sub

FillBgaAges_Error:

          Dim strES As String
          Dim intEL As Integer

200       intEL = Erl
210       strES = Err.Description
220       LogError "frmAges", "FillBgaAges", intEL, strES


End Sub

Private Sub FillBioAges()

          Dim tb As New Recordset
          Dim sql As String
          Dim Num As Long
          Dim Str As String

10        On Error GoTo FillBioAges_Error

20        ClearFGrid grdAge

30        sql = "SELECT AgeFromDays, AgeToDays from BioTestDefinitions WHERE " & _
                "LongName = '" & mAnalyte & "' " & _
                "and SampleType = '" & mSampleType & "' and COALESCE(category, '') = '" & mCat & "'" & _
                "Order by AgeFromDays"
40        Set tb = New Recordset
50        RecOpenClient 0, tb, sql

60        ReDim FromDays(0 To tb.recordCount - 1)
70        ReDim ToDays(0 To tb.recordCount - 1)
80        Num = 0
90        Do While Not tb.EOF
100           If tb!AgeFromDays & "" = "" Then FromDays(Num) = 0 Else FromDays(Num) = tb!AgeFromDays
110           If tb!AgeToDays & "" = "" Then ToDays(Num) = 43870 Else ToDays(Num) = tb!AgeToDays
120           Str = dmyFromCount(FromDays(Num)) & vbTab & _
                    dmyFromCount(ToDays(Num))
130           grdAge.AddItem Str
140           Num = Num + 1
150           tb.MoveNext
160       Loop

170       FixG grdAge

180       cmdRemove.Enabled = grdAge.Rows > 2



190       Exit Sub

FillBioAges_Error:

          Dim strES As String
          Dim intEL As Integer

200       intEL = Erl
210       strES = Err.Description
220       LogError "frmAges", "FillBioAges", intEL, strES


End Sub

Private Sub FillCoagAges()

          Dim tb As New Recordset
          Dim sql As String
          Dim Num As Long
          Dim Str As String

10        On Error GoTo FillCoagAges_Error

20        ClearFGrid grdAge
30        sql = "SELECT AgeFromDays, AgeToDays from CoagTestDefinitions WHERE " & _
                "TestName = '" & mAnalyte & "' " & _
                "Order by AgeFromDays"
40        Set tb = New Recordset
50        RecOpenClient 0, tb, sql

60        If tb.EOF Then Exit Sub

70        ReDim FromDays(0 To tb.recordCount - 1)
80        ReDim ToDays(0 To tb.recordCount - 1)
90        Num = 0
100       Do While Not tb.EOF
110           FromDays(Num) = tb!AgeFromDays
120           ToDays(Num) = tb!AgeToDays
130           Str = dmyFromCount(FromDays(Num)) & vbTab & _
                    dmyFromCount(ToDays(Num))
140           grdAge.AddItem Str
150           Num = Num + 1
160           tb.MoveNext
170       Loop

180       FixG grdAge


190       cmdRemove.Enabled = grdAge.Rows > 2



200       Exit Sub

FillCoagAges_Error:

          Dim strES As String
          Dim intEL As Integer

210       intEL = Erl
220       strES = Err.Description
230       LogError "frmAges", "FillCoagAges", intEL, strES


End Sub

Private Sub FillDMY(ByVal Days As Long, _
                    ByRef Year As Long, _
                    ByRef Month As Long, _
                    ByRef Day As Long)


10        On Error GoTo FillDMY_Error

20        Year = Days \ 365

30        Days = Days - (Year * 365)

40        Month = Days \ 30.42

50        Day = Days - (Month * 30.42)



60        Exit Sub

FillDMY_Error:

          Dim strES As String
          Dim intEL As Integer

70        intEL = Erl
80        strES = Err.Description
90        LogError "frmAges", "FillDMY", intEL, strES


End Sub

Private Sub FillEndAges()

          Dim tb As New Recordset
          Dim sql As String
          Dim Num As Long
          Dim Str As String


10        On Error GoTo FillEndAges_Error

20        ClearFGrid grdAge

30        sql = "SELECT AgeFromDays, AgeToDays from EndTestDefinitions WHERE " & _
                "LongName = '" & mAnalyte & "' " & _
                "and SampleType = '" & mSampleType & "' and category = '" & mCat & "' " & _
                "Order by AgeFromDays"
40        Set tb = New Recordset
50        RecOpenClient 0, tb, sql

60        ReDim FromDays(0 To tb.recordCount - 1)
70        ReDim ToDays(0 To tb.recordCount - 1)
80        Num = 0
90        Do While Not tb.EOF
100           FromDays(Num) = tb!AgeFromDays
110           ToDays(Num) = tb!AgeToDays
120           Str = dmyFromCount(FromDays(Num)) & vbTab & _
                    dmyFromCount(ToDays(Num))
130           grdAge.AddItem Str
140           Num = Num + 1
150           tb.MoveNext
160       Loop

170       FixG grdAge


180       cmdRemove.Enabled = grdAge.Rows > 2



190       Exit Sub

FillEndAges_Error:

          Dim strES As String
          Dim intEL As Integer

200       intEL = Erl
210       strES = Err.Description
220       LogError "frmAges", "FillEndAges", intEL, strES


End Sub

Private Sub FillHaemAges()

          Dim tb As New Recordset
          Dim sql As String
          Dim Num As Long
          Dim Str As String



10        On Error GoTo FillHaemAges_Error

20        ClearFGrid grdAge

30        sql = "SELECT AgeFromDays, AgeToDays from HaemTestDefinitions WHERE " & _
                "AnalyteName = '" & mAnalyte & "' " & _
                "Order by AgeFromDays"
40        Set tb = New Recordset
50        RecOpenClient 0, tb, sql

60        ReDim FromDays(0 To tb.recordCount - 1)
70        ReDim ToDays(0 To tb.recordCount - 1)
80        Num = 0
90        Do While Not tb.EOF
100           FromDays(Num) = tb!AgeFromDays & ""
110           ToDays(Num) = tb!AgeToDays & ""
120           Str = dmyFromCount(FromDays(Num)) & vbTab & _
                    dmyFromCount(ToDays(Num))
130           grdAge.AddItem Str
140           Num = Num + 1
150           tb.MoveNext
160       Loop

170       FixG grdAge

180       cmdRemove.Enabled = grdAge.Rows > 2



190       Exit Sub

FillHaemAges_Error:

          Dim strES As String
          Dim intEL As Integer

200       intEL = Erl
210       strES = Err.Description
220       LogError "frmAges", "FillHaemAges", intEL, strES


End Sub

Private Sub FillImmAges()

          Dim tb As New Recordset
          Dim sql As String
          Dim Num As Long
          Dim Str As String



10        On Error GoTo FillImmAges_Error

20        ClearFGrid grdAge

30        sql = "SELECT AgeFromDays, AgeToDays from IMmTestDefinitions WHERE " & _
                "LongName = '" & mAnalyte & "' " & _
                "and SampleType = '" & mSampleType & "' and analyser = '" & mAnalyser & "' " & _
                "Order by AgeFromDays"
40        Set tb = New Recordset
50        RecOpenClient 0, tb, sql

60        If Not tb.EOF Then
70            ReDim FromDays(0 To tb.recordCount - 1)
80            ReDim ToDays(0 To tb.recordCount - 1)
90            Num = 0
100           Do While Not tb.EOF
110               FromDays(Num) = tb!AgeFromDays
120               ToDays(Num) = tb!AgeToDays
130               Str = dmyFromCount(FromDays(Num)) & vbTab & _
                        dmyFromCount(ToDays(Num))
140               grdAge.AddItem Str
150               Num = Num + 1
160               tb.MoveNext
170           Loop
180       End If
190       FixG grdAge

200       cmdRemove.Enabled = grdAge.Rows > 2



210       Exit Sub

FillImmAges_Error:

          Dim strES As String
          Dim intEL As Integer

220       intEL = Erl
230       strES = Err.Description
240       LogError "frmAges", "FillImmAges", intEL, strES


End Sub

Private Sub Form_Activate()

10        On Error GoTo Form_Activate_Error

20        If Activated Then Exit Sub

30        Select Case mDiscipline
          Case "Biochemistry": FillBioAges
40        Case "Haematology": FillHaemAges
50        Case "Coagulation": FillCoagAges
60        Case "Immunology": FillImmAges
70        Case "Endocrinology": FillEndAges
80        Case "Blood Gas": FillBgaAges
90        End Select

100       AdjustG

110       grdAge.Col = 0
120       grdAge.row = 1
130       grdAge.CellBackColor = vbYellow

140       Activated = True

150       Exit Sub

Form_Activate_Error:

          Dim strES As String
          Dim intEL As Integer

160       intEL = Erl
170       strES = Err.Description
180       LogError "frmAges", "Form_Activate", intEL, strES


End Sub

Private Sub Form_Load()

          Dim Num As Long

10        On Error GoTo Form_Load_Error

20        lblParameter = mAnalyte

30        Select Case mSampleType
          Case "S": lblSampleType = ListText("ST", "S")
40        Case "U": lblSampleType = ListText("ST", "U")
50        Case "C": lblSampleType = ListText("ST", "C")
60        Case Else: lblSampleType = mSampleType
70        End Select

80        Activated = False

90        cmbYear.Clear
100       cmbMonth.Clear
110       cmbDay.Clear

120       For Num = 0 To 120
130           cmbYear.AddItem Format$(Num)
140       Next
150       For Num = 0 To 11
160           cmbMonth.AddItem Format$(Num)
170       Next
180       For Num = 0 To 30
190           cmbDay.AddItem Format$(Num)
200       Next

210       Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

220       intEL = Erl
230       strES = Err.Description
240       LogError "frmAges", "Form_Load", intEL, strES


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
60        LogError "frmAges", "Form_Unload", intEL, strES


End Sub

Private Sub grdAge_Click()

          Dim Num As Long
          Dim OrigNum As Long

          Dim Days As Long
          Dim Months As Long
          Dim Years As Long



10        On Error GoTo grdAge_Click_Error

20        If grdAge.MouseRow = 0 Then Exit Sub
          'If grdAge.MouseRow = grdAge.Rows - 1 Then Exit Sub

30        If grdAge.Rows = 2 And grdAge.MouseRow = 1 Then Exit Sub


40        OrigNum = grdAge.row

50        grdAge.Col = 0
60        For Num = 1 To grdAge.Rows - 1
70            grdAge.row = Num
80            grdAge.CellBackColor = 0
90        Next

100       grdAge.row = OrigNum
110       grdAge.CellBackColor = vbYellow

120       FillDMY ToDays(grdAge.row - 1), Years, Months, Days
130       cmbYear = Format$(Years)
140       cmbMonth = Format$(Months)
150       cmbDay = Format$(Days)

160       cmbYear.Top = grdAge.Top + 50 + (grdAge.row * 315)
170       cmbMonth.Top = grdAge.Top + 50 + (grdAge.row * 315)
180       cmbDay.Top = grdAge.Top + 50 + (grdAge.row * 315)

190       cmbYear.Visible = True
200       cmbMonth.Visible = True
210       cmbDay.Visible = True
220       cmdSave.Visible = True



230       Exit Sub

grdAge_Click_Error:

          Dim strES As String
          Dim intEL As Integer

240       intEL = Erl
250       strES = Err.Description
260       LogError "frmAges", "grdAge_Click", intEL, strES


End Sub

Private Sub RemoveBga()

          Dim Num As Long
          Dim rFrom As Long
          Dim rTo As Long
          Dim sql As String



10        On Error GoTo RemoveBga_Error

20        grdAge.Col = 0
30        For Num = 1 To grdAge.Rows - 1
40            grdAge.row = Num
50            If grdAge.CellBackColor = vbYellow Then
60                Exit For
70            End If
80        Next
90        Num = Num - 1

100       If Num = 0 Then Exit Sub

110       rFrom = FromDays(Num)
120       rTo = ToDays(Num)

130       sql = "DELETE from BgaTestDefinitions WHERE longname = '" & lblParameter & "' and category = '" & mCat & "' and" & _
                " AgeFromDays = '" & rFrom & "' " & _
                "and AgeToDays = '" & rTo & "'"
140       Cnxn(0).Execute sql

150       sql = "UPDATE BgaTestDefinitions " & _
                "Set AgeToDays = '" & rTo + 1 & "' " & _
                "WHERE AgeToDays = '" & rFrom - 1 & "' and longname = '" & lblParameter & "' and category = '" & mCat & "'"
160       Cnxn(0).Execute sql

170       FillBgaAges



180       Exit Sub

RemoveBga_Error:

          Dim strES As String
          Dim intEL As Integer

190       intEL = Erl
200       strES = Err.Description
210       LogError "frmAges", "RemoveBga", intEL, strES


End Sub

Private Sub RemoveBio()

          Dim Num As Long
          Dim rFrom As Long
          Dim rTo As Long
          Dim sql As String
          Dim Rfound As Boolean



10        On Error GoTo RemoveBio_Error

20        Rfound = False

30        grdAge.Col = 0
40        For Num = 1 To grdAge.Rows - 1
50            grdAge.row = Num
60            If grdAge.CellBackColor = vbYellow Then
70                Rfound = True
80                Exit For
90            End If
100       Next

110       If Rfound = False Then
120           iMsg "No row select for deletion!"
130           Exit Sub
140       End If

150       Num = Num - 1

160       If Num = 0 Then Exit Sub

170       rFrom = FromDays(Num)
180       rTo = ToDays(Num)


190       sql = "DELETE from BioTestDefinitions WHERE longname = '" & lblParameter & "' and category = '" & mCat & "' and" & _
                " AgeFromDays = '" & rFrom & "' " & _
                "and AgeToDays = '" & rTo & "'"
200       Cnxn(0).Execute sql


210       sql = "UPDATE BioTestDefinitions " & _
                "Set AgeToDays = '" & rTo + 1 & "' " & _
                "WHERE AgeToDays = '" & rFrom - 1 & "' and longname = '" & lblParameter & "' and category = '" & mCat & "'"
220       Cnxn(0).Execute sql

230       FillBioAges


240       Exit Sub

RemoveBio_Error:

          Dim strES As String
          Dim intEL As Integer

250       intEL = Erl
260       strES = Err.Description
270       LogError "frmAges", "RemoveBio", intEL, strES


End Sub

Private Sub RemoveCoag()

          Dim Num As Long
          Dim rFrom As Long
          Dim rTo As Long
          Dim sql As String


10        On Error GoTo RemoveCoag_Error

20        grdAge.Col = 0
30        For Num = 1 To grdAge.Rows - 1
40            grdAge.row = Num
50            If grdAge.CellBackColor = vbYellow Then
60                Exit For
70            End If
80        Next
90        Num = Num - 1

100       If Num = 0 Then Exit Sub

110       rFrom = FromDays(Num)
120       rTo = ToDays(Num)

130       sql = "DELETE from CoagTestDefinitions WHERE testname = '" & mAnalyte & "'" & _
                " and units = '" & mCat & "' and AgeFromDays = '" & rFrom & "' " & _
                "and AgeToDays = '" & rTo & "'"
140       Cnxn(0).Execute sql

150       sql = "UPDATE CoagTestDefinitions " & _
                "Set AgeToDays = '" & rTo + 1 & "' " & _
                "WHERE AgeToDays = '" & rFrom - 1 & "'  and units = '" & mCat & "' and testname = '" & mAnalyte & "'"
160       Cnxn(0).Execute sql

170       FillCoagAges




180       Exit Sub

RemoveCoag_Error:

          Dim strES As String
          Dim intEL As Integer

190       intEL = Erl
200       strES = Err.Description
210       LogError "frmAges", "RemoveCoag", intEL, strES


End Sub

Private Sub RemoveEnd()

          Dim Num As Long
          Dim rFrom As Long
          Dim rTo As Long
          Dim sql As String



10        On Error GoTo RemoveEnd_Error

20        grdAge.Col = 0
30        For Num = 1 To grdAge.Rows - 1
40            grdAge.row = Num
50            If grdAge.CellBackColor = vbYellow Then
60                Exit For
70            End If
80        Next
90        Num = Num - 1

100       If Num = 0 Then Exit Sub

110       rFrom = FromDays(Num)
120       rTo = ToDays(Num)

130       sql = "DELETE from EndTestDefinitions WHERE longname = '" & lblParameter & "' and" & _
                " AgeFromDays = '" & rFrom & "' " & _
                "and AgeToDays = '" & rTo & "'"
140       Cnxn(0).Execute sql

150       sql = "UPDATE EndTestDefinitions " & _
                "SET AgeToDays = '" & rTo + 1 & "' " & _
                "WHERE AgeToDays = '" & rFrom - 1 & "' AND longname = '" & lblParameter & "'"
160       Cnxn(0).Execute sql

170       FillEndAges



180       Exit Sub

RemoveEnd_Error:

          Dim strES As String
          Dim intEL As Integer

190       intEL = Erl
200       strES = Err.Description
210       LogError "frmAges", "RemoveEnd", intEL, strES


End Sub

Private Sub RemoveHaem()

          Dim Num As Long
          Dim rFrom As Long
          Dim rTo As Long
          Dim sql As String



10        On Error GoTo RemoveHaem_Error

20        grdAge.Col = 0
30        For Num = 1 To grdAge.Rows - 1
40            grdAge.row = Num
50            If grdAge.CellBackColor = vbYellow Then
60                Exit For
70            End If
80        Next
90        Num = Num - 1

100       If Num = 0 Then Exit Sub

110       rFrom = FromDays(Num)
120       rTo = ToDays(Num)

130       sql = "DELETE from HaemTestDefinitions WHERE " & _
                "AgeFromDays = '" & rFrom & "' " & _
                "and AgeToDays = '" & rTo & "' and AnalyteName = '" & lblParameter & "'"
140       Cnxn(0).Execute sql

150       sql = "UPDATE HaemTestDefinitions " & _
                "Set AgeToDays = '" & rTo + 1 & "' " & _
                "WHERE AgeToDays = '" & rFrom - 1 & "' and AnalyteName = '" & lblParameter & "'"
160       Cnxn(0).Execute sql

170       FillHaemAges




180       Exit Sub

RemoveHaem_Error:

          Dim strES As String
          Dim intEL As Integer

190       intEL = Erl
200       strES = Err.Description
210       LogError "frmAges", "RemoveHaem", intEL, strES


End Sub

Private Sub RemoveImm()

          Dim Num As Long
          Dim rFrom As Long
          Dim rTo As Long
          Dim sql As String



10        On Error GoTo RemoveImm_Error

20        grdAge.Col = 0
30        For Num = 1 To grdAge.Rows - 1
40            grdAge.row = Num
50            If grdAge.CellBackColor = vbYellow Then
60                Exit For
70            End If
80        Next
90        Num = Num - 1

100       If Num = 0 Then Exit Sub

110       rFrom = FromDays(Num)
120       rTo = ToDays(Num)

130       sql = "DELETE from ImmTestDefinitions WHERE longname = '" & lblParameter & "' and" & _
                " AgeFromDays = '" & rFrom & "' " & _
                "and AgeToDays = '" & rTo & "' and analyser = '" & mAnalyser & "'"
140       Cnxn(0).Execute sql

150       sql = "UPDATE IMMTestDefinitions " & _
                "Set AgeToDays = '" & rTo + 1 & "' " & _
                "WHERE AgeToDays = '" & rFrom - 1 & "' and longname = '" & lblParameter & "' and analyser = '" & mAnalyser & "'"
160       Cnxn(0).Execute sql

170       FillImmAges




180       Exit Sub

RemoveImm_Error:

          Dim strES As String
          Dim intEL As Integer

190       intEL = Erl
200       strES = Err.Description
210       LogError "frmAges", "RemoveImm", intEL, strES


End Sub

Public Property Let SampleType(ByVal SampleType As String)

10        On Error GoTo SampleType_Error

20        mSampleType = SampleType

30        Exit Property

SampleType_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmAges", "SampleType", intEL, strES


End Property

Private Sub SaveBga()

          Dim sql As String
          Dim Days As Long
          Dim Num As Long


10        On Error GoTo SaveBga_Error

20        ReDim WasFrom(0 To UBound(FromDays)) As Long
30        ReDim WasTo(0 To UBound(ToDays)) As Long
          Dim Active As Long

40        grdAge.Col = 0
50        For Num = 1 To grdAge.Rows - 1
60            grdAge.row = Num
70            If grdAge.CellBackColor = vbYellow Then
80                Active = Num - 1
90                Exit For
100           End If
110       Next

120       Days = (Val(cmbYear) * 365.25) + (Val(cmbMonth) * 30.42) + Val(cmbDay)
130       If Days = 0 Then Exit Sub

140       For Num = 0 To UBound(FromDays)
150           WasFrom(Num) = FromDays(Num)
160           WasTo(Num) = ToDays(Num)
170       Next

180       ToDays(Active) = Days

190       For Num = 0 To UBound(FromDays) - 1
200           FromDays(Num + 1) = ToDays(Num) + 1
210           If ToDays(Num + 1) < FromDays(Num + 1) Then
220               ToDays(Num + 1) = FromDays(Num + 1)
230           End If
240       Next

250       For Num = 0 To UBound(WasFrom)
260           sql = "UPDATE BgaTestDefinitions " & _
                    "Set AgeFromDays = '" & FromDays(Num) & "', " & _
                    "AgeToDays = '" & ToDays(Num) & "' WHERE " & _
                    "AgeFromDays = '" & WasFrom(Num) & "' " & _
                    "and AgeToDays = '" & WasTo(Num) & "' " & _
                    "and longName = '" & lblParameter & "' " & _
                    "and SampleType = '" & mSampleType & "'"
270           Cnxn(0).Execute sql
280       Next

290       FillBgaAges

300       AdjustG

310       cmbYear.Visible = False
320       cmbMonth.Visible = False
330       cmbDay.Visible = False
340       cmdSave.Visible = False



350       Exit Sub

SaveBga_Error:

          Dim strES As String
          Dim intEL As Integer

360       intEL = Erl
370       strES = Err.Description
380       LogError "frmAges", "SaveBga", intEL, strES


End Sub

Private Sub SaveBio()

          Dim sql As String
          Dim Days As Long
          Dim Num As Long


10        On Error GoTo SaveBio_Error

20        ReDim WasFrom(0 To UBound(FromDays)) As Long
30        ReDim WasTo(0 To UBound(ToDays)) As Long
          Dim Active As Long

40        grdAge.Col = 0
50        For Num = 1 To grdAge.Rows - 1
60            grdAge.row = Num
70            If grdAge.CellBackColor = vbYellow Then
80                Active = Num - 1
90                Exit For
100           End If
110       Next

120       Days = (Val(cmbYear) * 365.25) + (Val(cmbMonth) * 30.42) + Val(cmbDay)
130       If Days = 0 Then Exit Sub

140       For Num = 0 To UBound(FromDays)
150           WasFrom(Num) = FromDays(Num)
160           WasTo(Num) = ToDays(Num)
170       Next

180       ToDays(Active) = Days

190       For Num = 0 To UBound(FromDays) - 1
200           FromDays(Num + 1) = ToDays(Num) + 1
210           If ToDays(Num + 1) < FromDays(Num + 1) Then
220               ToDays(Num + 1) = FromDays(Num + 1)
230           End If
240       Next

250       For Num = 0 To UBound(WasFrom)
260           sql = "UPDATE BioTestDefinitions " & _
                    "Set AgeFromDays = '" & FromDays(Num) & "', " & _
                    "AgeToDays = '" & ToDays(Num) & "' WHERE " & _
                    "AgeFromDays = '" & WasFrom(Num) & "' " & _
                    "and AgeToDays = '" & WasTo(Num) & "' " & _
                    "and longName = '" & lblParameter & "' " & _
                    "and SampleType = '" & mSampleType & "' and category = '" & mCat & "'"
270           Cnxn(0).Execute sql
280       Next

290       AdjustG

300       cmbYear.Visible = False
310       cmbMonth.Visible = False
320       cmbDay.Visible = False
330       cmdSave.Visible = False



340       Exit Sub

SaveBio_Error:

          Dim strES As String
          Dim intEL As Integer

350       intEL = Erl
360       strES = Err.Description
370       LogError "frmAges", "SaveBio", intEL, strES


End Sub

Private Sub SaveCoag()

          Dim sql As String
          Dim Days As Long
          Dim Num As Long


10        On Error GoTo SaveCoag_Error

20        ReDim WasFrom(0 To UBound(FromDays)) As Long
30        ReDim WasTo(0 To UBound(ToDays)) As Long
          Dim Active As Long

40        grdAge.Col = 0
50        For Num = 1 To grdAge.Rows - 1
60            grdAge.row = Num
70            If grdAge.CellBackColor = vbYellow Then
80                Active = Num - 1
90                Exit For
100           End If
110       Next

120       Days = (Val(cmbYear) * 365.25) + (Val(cmbMonth) * 30.42) + Val(cmbDay)
130       If Days = 0 Then Exit Sub

140       For Num = 0 To UBound(FromDays)
150           WasFrom(Num) = FromDays(Num)
160           WasTo(Num) = ToDays(Num)
170       Next

180       ToDays(Active) = Days

190       For Num = 0 To UBound(FromDays) - 1
200           FromDays(Num + 1) = ToDays(Num) + 1
210           If ToDays(Num + 1) < FromDays(Num + 1) Then
220               ToDays(Num + 1) = FromDays(Num + 1)
230           End If
240       Next

250       For Num = 0 To UBound(WasFrom)
260           sql = "UPDATE CoagTestDefinitions " & _
                    "Set AgeFromDays = '" & FromDays(Num) & "', " & _
                    "AgeToDays = '" & ToDays(Num) & "' WHERE " & _
                    "AgeFromDays = '" & WasFrom(Num) & "' " & _
                    "and AgeToDays = '" & WasTo(Num) & "' " & _
                    "and TestName = '" & lblParameter & "'"
270           Cnxn(0).Execute sql
280       Next

290       AdjustG

300       cmbYear.Visible = False
310       cmbMonth.Visible = False
320       cmbDay.Visible = False
330       cmdSave.Visible = False


340       Exit Sub

SaveCoag_Error:

          Dim strES As String
          Dim intEL As Integer

350       intEL = Erl
360       strES = Err.Description
370       LogError "frmAges", "SaveCoag", intEL, strES, sql


End Sub

Private Sub SaveEnd()

          Dim sql As String
          Dim Days As Long
          Dim Num As Long


10        On Error GoTo SaveEnd_Error

20        ReDim WasFrom(0 To UBound(FromDays)) As Long
30        ReDim WasTo(0 To UBound(ToDays)) As Long
          Dim Active As Long

40        grdAge.Col = 0
50        For Num = 1 To grdAge.Rows - 1
60            grdAge.row = Num
70            If grdAge.CellBackColor = vbYellow Then
80                Active = Num - 1
90                Exit For
100           End If
110       Next

120       Days = (Val(cmbYear) * 365.25) + (Val(cmbMonth) * 30.42) + Val(cmbDay)
130       If Days = 0 Then Exit Sub

140       For Num = 0 To UBound(FromDays)
150           WasFrom(Num) = FromDays(Num)
160           WasTo(Num) = ToDays(Num)
170       Next

180       ToDays(Active) = Days

190       For Num = 0 To UBound(FromDays) - 1
200           FromDays(Num + 1) = ToDays(Num) + 1
210           If ToDays(Num + 1) < FromDays(Num + 1) Then
220               ToDays(Num + 1) = FromDays(Num + 1)
230           End If
240       Next

250       For Num = 0 To UBound(WasFrom)
260           sql = "UPDATE EndTestDefinitions " & _
                    "Set AgeFromDays = '" & FromDays(Num) & "', " & _
                    "AgeToDays = '" & ToDays(Num) & "' WHERE " & _
                    "AgeFromDays = '" & WasFrom(Num) & "' " & _
                    "and AgeToDays = '" & WasTo(Num) & "' " & _
                    "and longName = '" & lblParameter & "' " & _
                    "and SampleType = '" & mSampleType & "'"
270           Cnxn(0).Execute sql
280       Next

290       FillEndAges

300       AdjustG

310       cmbYear.Visible = False
320       cmbMonth.Visible = False
330       cmbDay.Visible = False
340       cmdSave.Visible = False



350       Exit Sub

SaveEnd_Error:

          Dim strES As String
          Dim intEL As Integer

360       intEL = Erl
370       strES = Err.Description
380       LogError "frmAges", "SaveEnd", intEL, strES, sql


End Sub

Private Sub SaveHaem()

          Dim sql As String
          Dim Days As Long
          Dim Num As Long


10        On Error GoTo SaveHaem_Error

20        ReDim WasFrom(0 To UBound(FromDays)) As Long
30        ReDim WasTo(0 To UBound(ToDays)) As Long
          Dim Active As Long

40        grdAge.Col = 0
50        For Num = 1 To grdAge.Rows - 1
60            grdAge.row = Num
70            If grdAge.CellBackColor = vbYellow Then
80                Active = Num - 1
90                Exit For
100           End If
110       Next

120       Days = Fix((Val(cmbYear) * 365.25) + (Val(cmbMonth) * 30.42) + Val(cmbDay))
130       If Days = 0 Then Exit Sub

140       For Num = 0 To UBound(FromDays)
150           WasFrom(Num) = FromDays(Num)
160           WasTo(Num) = ToDays(Num)
170       Next

180       ToDays(Active) = Days

190       For Num = 0 To UBound(FromDays) - 1
200           FromDays(Num + 1) = ToDays(Num) + 1
210           If ToDays(Num + 1) < FromDays(Num + 1) Then
220               ToDays(Num + 1) = FromDays(Num + 1)
230           End If
240       Next

250       For Num = 0 To UBound(WasFrom)
260           sql = "UPDATE HaemTestDefinitions " & _
                    "Set AgeFromDays = '" & FromDays(Num) & "', " & _
                    "AgeToDays = '" & ToDays(Num) & "' WHERE " & _
                    "AgeFromDays = '" & WasFrom(Num) & "' " & _
                    "and AgeToDays = '" & WasTo(Num) & "' " & _
                    "and AnalyteName = '" & lblParameter & "'"
270           Cnxn(0).Execute sql
280       Next

290       AdjustG

300       cmbYear.Visible = False
310       cmbMonth.Visible = False
320       cmbDay.Visible = False
330       cmdSave.Visible = False



340       Exit Sub

SaveHaem_Error:

          Dim strES As String
          Dim intEL As Integer

350       intEL = Erl
360       strES = Err.Description
370       LogError "frmAges", "SaveHaem", intEL, strES, sql


End Sub

Private Sub SaveImm()

          Dim sql As String
          Dim Days As Long
          Dim Num As Long


10        On Error GoTo SaveImm_Error

20        ReDim WasFrom(0 To UBound(FromDays)) As Long
30        ReDim WasTo(0 To UBound(ToDays)) As Long
          Dim Active As Long

40        grdAge.Col = 0
50        For Num = 1 To grdAge.Rows - 1
60            grdAge.row = Num
70            If grdAge.CellBackColor = vbYellow Then
80                Active = Num - 1
90                Exit For
100           End If
110       Next

120       Days = (Val(cmbYear) * 365.25) + (Val(cmbMonth) * 30.42) + Val(cmbDay)
130       If Days = 0 Then Exit Sub

140       For Num = 0 To UBound(FromDays)
150           WasFrom(Num) = FromDays(Num)
160           WasTo(Num) = ToDays(Num)
170       Next

180       ToDays(Active) = Days

190       For Num = 0 To UBound(FromDays) - 1
200           FromDays(Num + 1) = ToDays(Num) + 1
210           If ToDays(Num + 1) < FromDays(Num + 1) Then
220               ToDays(Num + 1) = FromDays(Num + 1)
230           End If
240       Next

250       For Num = 0 To UBound(WasFrom)
260           sql = "UPDATE ImmTestDefinitions " & _
                    "Set AgeFromDays = '" & FromDays(Num) & "', " & _
                    "AgeToDays = '" & ToDays(Num) & "' WHERE " & _
                    "AgeFromDays = '" & WasFrom(Num) & "' " & _
                    "and AgeToDays = '" & WasTo(Num) & "' " & _
                    "and longName = '" & lblParameter & "' " & _
                    "and SampleType = '" & mSampleType & "' and analyser = '" & mAnalyser & "'"
270           Cnxn(0).Execute sql
280       Next

290       FillImmAges

300       AdjustG

310       cmbYear.Visible = False
320       cmbMonth.Visible = False
330       cmbDay.Visible = False
340       cmdSave.Visible = False



350       Exit Sub

SaveImm_Error:

          Dim strES As String
          Dim intEL As Integer

360       intEL = Erl
370       strES = Err.Description
380       LogError "frmAges", "SaveImm", intEL, strES, sql


End Sub

