VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmBioMultiAddAge 
   Caption         =   "NetAcquire - Ages"
   ClientHeight    =   6360
   ClientLeft      =   870
   ClientTop       =   420
   ClientWidth     =   4905
   LinkTopic       =   "Form1"
   ScaleHeight     =   6360
   ScaleWidth      =   4905
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   855
      HelpContextID   =   10026
      Left            =   3150
      Picture         =   "frmBioMultiAddAge.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5310
      Width           =   1455
   End
   Begin VB.ComboBox cM 
      Height          =   315
      Left            =   3030
      TabIndex        =   7
      Text            =   "cM"
      Top             =   1410
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.ComboBox cY 
      Height          =   315
      Left            =   2310
      TabIndex        =   6
      Text            =   "cY"
      Top             =   1410
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox cD 
      Height          =   315
      Left            =   3600
      TabIndex        =   5
      Text            =   "cD"
      Top             =   1410
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove Age Range"
      Height          =   735
      Left            =   3030
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4080
      Width           =   1605
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add Age Range"
      Height          =   735
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4080
      Width           =   1605
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H0000FFFF&
      Caption         =   "&Save"
      Height          =   735
      Left            =   2070
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4080
      Visible         =   0   'False
      Width           =   825
   End
   Begin MSFlexGridLib.MSFlexGrid grdAges 
      Height          =   2985
      Left            =   330
      TabIndex        =   0
      Top             =   960
      Width           =   4275
      _ExtentX        =   7541
      _ExtentY        =   5265
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
   Begin VB.Label lblDiscipline 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Biochemistry"
      Height          =   285
      Left            =   1440
      TabIndex        =   9
      Top             =   90
      Width           =   1965
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000001&
      BorderWidth     =   2
      Height          =   75
      Left            =   360
      Top             =   5070
      Width           =   4335
   End
   Begin VB.Label lblParameter 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblParameter"
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
      Left            =   1410
      TabIndex        =   2
      Top             =   450
      Width           =   2025
   End
End
Attribute VB_Name = "frmBioMultiAddAge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mAnalyte As String
Private mSampleType As String
Private mHospital As String
Private mCategory As String

Private Activated As Boolean

Private FromDays() As Long
Private ToDays() As Long
Private Sub AdjustG()

          Dim Y As Integer

10        On Error GoTo AdjustG_Error

20        For Y = 0 To UBound(FromDays)
30            grdAges.TextMatrix(Y + 1, 0) = dmyFromCount(FromDays(Y))
40            grdAges.TextMatrix(Y + 1, 1) = dmyFromCount(ToDays(Y))
50        Next

60        Exit Sub

AdjustG_Error:

          Dim strES As String
          Dim intEL As Integer

70        intEL = Erl
80        strES = Err.Description
90        LogError "frmBioMultiAddAge", "AdjustG", intEL, strES


End Sub

Private Sub FillAges()

          Dim tb As Recordset
          Dim sql As String
          Dim n As Integer
          Dim s As String
          Dim Discipline As String

10        On Error GoTo FillAges_Error

20        If Trim$(mSampleType) = "" Then
30            iMsg "Define Sample Type"
40            Exit Sub
50        End If
60        If Trim$(mHospital) = "" Then
70            iMsg "Define Hospital"
80            Exit Sub
90        End If
100       If Trim$(mCategory) = "" Then
110           iMsg "Define Category"
120           Exit Sub
130       End If

140       grdAges.Rows = 2
150       grdAges.AddItem ""
160       grdAges.RemoveItem 1

170       Discipline = Left$(lblDiscipline, 3)

180       sql = "Select AgeFromDays, AgeToDays from " & Discipline & "TestDefinitions where " & _
                "LongName = '" & mAnalyte & "' " & _
                "and Hospital = '" & mHospital & "' " & _
                "and Category = '" & mCategory & "' " & _
                "and SampleType = '" & mSampleType & "' " & _
                "and ActiveToDate = '" & Format$(Now, "dd/mmm/yyyy") & "' " & _
                "Order by AgeFromDays"
190       Set tb = New Recordset
200       RecOpenClient 0, tb, sql

210       ReDim FromDays(0 To tb.RecordCount - 1)
220       ReDim ToDays(0 To tb.RecordCount - 1)
230       n = 0
240       Do While Not tb.EOF
250           FromDays(n) = tb!AgeFromDays
260           ToDays(n) = tb!AgeToDays
270           s = dmyFromCount(FromDays(n)) & vbTab & _
                  dmyFromCount(ToDays(n))
280           grdAges.AddItem s
290           n = n + 1
300           tb.MoveNext
310       Loop

320       If grdAges.Rows > 2 Then
330           grdAges.RemoveItem 1
340       End If

350       cmdRemove.Enabled = grdAges.Rows > 2

360       Exit Sub

FillAges_Error:

          Dim strES As String
          Dim intEL As Integer

370       intEL = Erl
380       strES = Err.Description
390       LogError "frmBioMultiAddAge", "FillAges", intEL, strES


End Sub

Private Sub FillDMY(ByVal Days As Long, _
                    ByRef Y As Long, _
                    ByRef m As Long, _
                    ByRef D As Long)

10        On Error GoTo FillDMY_Error

20        Y = Days \ 365

30        Days = Days - (Y * 365)

40        m = Days \ 30

50        D = Days - (m * 30)

60        Exit Sub

FillDMY_Error:

          Dim strES As String
          Dim intEL As Integer

70        intEL = Erl
80        strES = Err.Description
90        LogError "frmBioMultiAddAge", "FillDMY", intEL, strES


End Sub

Private Sub cmdadd_Click()

          Dim tb As Recordset
          Dim tbNew As Recordset
          Dim sql As String
          Dim fld As Field

10        On Error GoTo cmdadd_Click_Error

20        If Trim$(mSampleType) = "" Then
30            iMsg "Define Sample Type"
40            Exit Sub
50        End If
60        If Trim$(mHospital) = "" Then
70            iMsg "Define Hospital"
80            Exit Sub
90        End If
100       If Trim$(mCategory) = "" Then
110           iMsg "Define Category"
120           Exit Sub
130       End If

140       sql = "Select top 1 * from BioTestDefinitions where " & _
                "LongName = '" & mAnalyte & "' " & _
                "and Hospital = '" & mHospital & "' " & _
                "and Category = '" & mCategory & "' " & _
                "and SampleType = '" & mSampleType & "' " & _
                "and ActiveToDate = '" & Format$(Now, "dd/mmm/yyyy") & "' " & _
                "order by AgeToDays desc"
150       Set tb = New Recordset
160       RecOpenClient 0, tb, sql

170       Set tbNew = New Recordset
180       RecOpenClient 0, tbNew, sql

190       tbNew.AddNew
200       For Each fld In tb.Fields
210           If fld.Name = "AgeToDays" Then
220               tbNew!AgeToDays = tb!AgeToDays + 1
230           ElseIf fld.Name = "AgeFromDays" Then
240               tbNew!AgeFromDays = tb!AgeToDays + 1
250           Else
260               tbNew(fld.Name) = tb(fld.Name)
270           End If
280       Next
290       tbNew.Update

300       FillAges

310       Exit Sub

cmdadd_Click_Error:

          Dim strES As String
          Dim intEL As Integer

320       intEL = Erl
330       strES = Err.Description
340       LogError "frmBioMultiAddAge", "cmdAdd_Click", intEL, strES


End Sub


Private Sub cmdCancel_Click()

10        Unload Me

End Sub

Private Sub cmdRemove_Click()

          Dim Y As Integer
          Dim rFrom As Long
          Dim rTo As Long
          Dim sql As String
          Dim Discipline As String

10        On Error GoTo cmdRemove_Click_Error

20        If Trim$(mSampleType) = "" Then
30            iMsg "Define Sample Type"
40            Exit Sub
50        End If
60        If Trim$(mHospital) = "" Then
70            iMsg "Define Hospital"
80            Exit Sub
90        End If
100       If Trim$(mCategory) = "" Then
110           iMsg "Define Category"
120           Exit Sub
130       End If

140       grdAges.Col = 0
150       For Y = 1 To grdAges.Rows - 2
160           grdAges.Row = Y
170           If grdAges.CellBackColor = vbYellow Then
180               Exit For
190           End If
200       Next
210       Y = Y - 1

220       If Y = 0 Then Exit Sub

230       rFrom = FromDays(Y)
240       rTo = ToDays(Y)

250       Discipline = Left$(lblDiscipline, 3)

260       sql = "Delete from " & Discipline & "TestDefinitions where " & _
                "AgeFromDays = '" & rFrom & "' " & _
                "and AgeToDays = '" & rTo & "' " & _
                "and LongName = '" & mAnalyte & "' " & _
                "and Hospital = '" & mHospital & "' " & _
                "and Category = '" & mCategory & "' " & _
                "and SampleType = '" & mSampleType & "' " & _
                "and ActiveToDate = '" & Format$(Now, "dd/mmm/yyyy") & "'"
270       Cnxn(0).Execute sql

280       sql = "Update " & Discipline & "TestDefinitions " & _
                "Set AgeToDays = '" & rTo + 1 & "' " & _
                "where AgeToDays = '" & rFrom - 1 & "' " & _
                "and LongName = '" & mAnalyte & "' " & _
                "and Hospital = '" & mHospital & "' " & _
                "and Category = '" & mCategory & "' " & _
                "and SampleType = '" & mSampleType & "' " & _
                "and ActiveToDate = '" & Format$(Now, "dd/mmm/yyyy") & "'"
290       Cnxn(0).Execute sql

300       FillAges

310       Exit Sub

cmdRemove_Click_Error:

          Dim strES As String
          Dim intEL As Integer

320       intEL = Erl
330       strES = Err.Description
340       LogError "frmBioMultiAddAge", "cmdRemove_Click", intEL, strES


End Sub

Private Sub cmdSave_Click()

          Dim sql As String
          Dim Days As Long
          Dim n As Integer
10        On Error GoTo cmdSave_Click_Error

20        ReDim WasFrom(0 To UBound(FromDays)) As Long
30        ReDim WasTo(0 To UBound(ToDays)) As Long
          Dim Active As Integer
          Dim Discipline As String

40        If Trim$(mSampleType) = "" Then
50            iMsg "Define Sample Type"
60            Exit Sub
70        End If
80        If Trim$(mHospital) = "" Then
90            iMsg "Define Hospital"
100           Exit Sub
110       End If
          'If Trim$(mCategory) = "" Then
          '  iMsg "Define Category"
          '  Exit Sub
          'End If

120       grdAges.Col = 0
130       For n = 1 To grdAges.Rows - 2
140           grdAges.Row = n
150           If grdAges.CellBackColor = vbYellow Then
160               Active = n - 1
170               Exit For
180           End If
190       Next

200       Days = (Val(cY) * 365) + (Val(cM) * 30) + Val(cD)
210       If Days = 0 Then Exit Sub

220       For n = 0 To UBound(FromDays)
230           WasFrom(n) = FromDays(n)
240           WasTo(n) = ToDays(n)
250       Next

260       ToDays(Active) = Days

270       For n = 0 To UBound(FromDays) - 1
280           FromDays(n + 1) = ToDays(n) + 1
290           If ToDays(n + 1) < FromDays(n + 1) Then
300               ToDays(n + 1) = FromDays(n + 1)
310           End If
320       Next

330       Discipline = Left$(lblDiscipline, 3)

340       For n = 0 To UBound(WasFrom)
350           sql = "Update " & Discipline & "TestDefinitions " & _
                    "Set AgeFromDays = '" & FromDays(n) & "', " & _
                    "AgeToDays = '" & ToDays(n) & "' where " & _
                    "AgeFromDays = '" & WasFrom(n) & "' " & _
                    "and AgeToDays = '" & WasTo(n) & "' " & _
                    "and LongName = '" & mAnalyte & "' " & _
                    "and Hospital = '" & mHospital & "' " & _
                    "and Category = '" & mCategory & "' " & _
                    "and SampleType = '" & mSampleType & "' " & _
                    "and ActiveToDate = '" & Format$(Now, "dd/mmm/yyyy") & "'"
360           Cnxn(0).Execute sql
370       Next

380       AdjustG

390       cY.Visible = False
400       cM.Visible = False
410       cD.Visible = False
420       cmdSave.Visible = False
430       cmdAdd.Visible = True

440       Exit Sub

cmdSave_Click_Error:

          Dim strES As String
          Dim intEL As Integer

450       intEL = Erl
460       strES = Err.Description
470       LogError "frmBioMultiAddAge", "cmdsave_Click", intEL, strES


End Sub

Private Sub cD_Click()

10        On Error GoTo cD_Click_Error

20        cmdSave.Visible = True
30        cmdAdd.Visible = False

40        Exit Sub

cD_Click_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmBioMultiAddAge", "cD_Click", intEL, strES


End Sub

Private Sub cD_KeyPress(KeyAscii As Integer)

10        On Error GoTo cD_KeyPress_Error

20        KeyAscii = 0

30        Exit Sub

cD_KeyPress_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmBioMultiAddAge", "cD_KeyPress", intEL, strES


End Sub


Private Sub cM_Click()

10        On Error GoTo cM_Click_Error

20        cmdSave.Visible = True
30        cmdAdd.Visible = False

40        Exit Sub

cM_Click_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmBioMultiAddAge", "cM_Click", intEL, strES


End Sub

Private Sub cM_KeyPress(KeyAscii As Integer)

10        On Error GoTo cM_KeyPress_Error

20        KeyAscii = 0

30        Exit Sub

cM_KeyPress_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmBioMultiAddAge", "cM_KeyPress", intEL, strES


End Sub


Private Sub cY_Click()

10        On Error GoTo cY_Click_Error

20        cmdSave.Visible = True
30        cmdAdd.Visible = False

40        Exit Sub

cY_Click_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmBioMultiAddAge", "cY_Click", intEL, strES


End Sub

Private Sub cY_KeyPress(KeyAscii As Integer)

10        On Error GoTo cY_KeyPress_Error

20        KeyAscii = 0

30        Exit Sub

cY_KeyPress_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmBioMultiAddAge", "cY_KeyPress", intEL, strES


End Sub


Private Sub Form_Activate()

10        On Error GoTo Form_Activate_Error

20        If Activated Then Exit Sub


30        FillAges

40        AdjustG

50        grdAges.Col = 0
60        grdAges.Row = 1
70        grdAges.CellBackColor = vbYellow

80        Activated = True

90        Exit Sub

Form_Activate_Error:

          Dim strES As String
          Dim intEL As Integer

100       intEL = Erl
110       strES = Err.Description
120       LogError "frmBioMultiAddAge", "Form_Activate", intEL, strES


End Sub

Private Sub Form_Load()

          Dim n As Integer

10        On Error GoTo Form_Load_Error

20        lblParameter = mAnalyte

30        Activated = False

40        cY.Clear
50        cM.Clear
60        cD.Clear

70        For n = 0 To 120
80            cY.AddItem Format$(n)
90        Next
100       For n = 0 To 11
110           cM.AddItem Format$(n)
120       Next
130       For n = 0 To 30
140           cD.AddItem Format$(n)
150       Next

160       Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

170       intEL = Erl
180       strES = Err.Description
190       LogError "frmBioMultiAddAge", "Form_Load", intEL, strES


End Sub


Private Sub Form_Unload(Cancel As Integer)

10        On Error GoTo Form_Unload_Error

20        If cmdSave.Visible Then
30            If iMsg("Cancel without Saving?", vbQuestion + vbYesNo) = vbNo Then
40                Cancel = True
50                Exit Sub
60            End If
70        End If

80        Activated = False

90        Exit Sub

Form_Unload_Error:

          Dim strES As String
          Dim intEL As Integer

100       intEL = Erl
110       strES = Err.Description
120       LogError "frmBioMultiAddAge", "Form_Unload", intEL, strES


End Sub


Private Sub grdAges_Click()

          Dim Y As Integer
          Dim OrigY As Integer

          Dim Days As Long
          Dim Months As Long
          Dim Years As Long

10        On Error GoTo grdAges_Click_Error

20        If grdAges.MouseRow = 0 Then Exit Sub
30        If grdAges.MouseRow = grdAges.Rows - 1 Then Exit Sub

40        OrigY = grdAges.Row

50        grdAges.Col = 0
60        For Y = 1 To grdAges.Rows - 1
70            grdAges.Row = Y
80            grdAges.CellBackColor = 0
90        Next

100       grdAges.Row = OrigY
110       grdAges.CellBackColor = vbYellow

120       FillDMY ToDays(grdAges.Row - 1), Years, Months, Days
130       cY = Format$(Years)
140       cM = Format$(Months)
150       cD = Format$(Days)

160       cY.Top = grdAges.Top + 50 + (grdAges.Row * 315)
170       cM.Top = grdAges.Top + 50 + (grdAges.Row * 315)
180       cD.Top = grdAges.Top + 50 + (grdAges.Row * 315)

190       cY.Visible = True
200       cM.Visible = True
210       cD.Visible = True
220       cmdSave.Visible = True

230       Exit Sub

grdAges_Click_Error:

          Dim strES As String
          Dim intEL As Integer

240       intEL = Erl
250       strES = Err.Description
260       LogError "frmBioMultiAddAge", "grdAges_Click", intEL, strES


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
60        LogError "frmBioMultiAddAge", "Analyte", intEL, strES


End Property


Public Property Let Category(ByVal Category As String)

10        On Error GoTo Category_Error

20        mCategory = Category

30        Exit Property

Category_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmBioMultiAddAge", "Category", intEL, strES


End Property



Public Property Let Discipline(ByVal Discipline As String)

10        On Error GoTo Discipline_Error

20        lblDiscipline = Discipline

30        Exit Property

Discipline_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmBioMultiAddAge", "Discipline", intEL, strES


End Property




Public Property Let Hospital(ByVal Hospital As String)

10        On Error GoTo Hospital_Error

20        mHospital = Hospital

30        Exit Property

Hospital_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmBioMultiAddAge", "Hospital", intEL, strES


End Property
Public Property Let SampleType(ByVal SampleType As String)

10        On Error GoTo SampleType_Error

20        mSampleType = SampleType

30        Exit Property

SampleType_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmBioMultiAddAge", "SampleType", intEL, strES


End Property

