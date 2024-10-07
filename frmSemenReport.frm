VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.Ocx"
Begin VB.Form frmSemenReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NetAcquire --- Semen Analysis History"
   ClientHeight    =   9645
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12075
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9645
   ScaleWidth      =   12075
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   795
      Left            =   120
      TabIndex        =   4
      Top             =   405
      Width           =   11775
      Begin VB.Label lblName 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   855
         TabIndex        =   14
         Top             =   225
         Width           =   3960
      End
      Begin VB.Label lblChart 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8505
         TabIndex        =   13
         Top             =   225
         Width           =   1200
      End
      Begin VB.Label lblDoB 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6180
         TabIndex        =   12
         Top             =   225
         Width           =   1470
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         Height          =   195
         Index           =   1
         Left            =   195
         TabIndex        =   11
         Top             =   315
         Width           =   420
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Chart"
         Height          =   195
         Index           =   1
         Left            =   7890
         TabIndex        =   10
         Top             =   315
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Date of Birth"
         Height          =   195
         Index           =   1
         Left            =   5055
         TabIndex        =   9
         Top             =   315
         Width           =   885
      End
      Begin VB.Label lblSex 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10455
         TabIndex        =   8
         Top             =   225
         Width           =   1020
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Sex"
         Height          =   195
         Index           =   1
         Left            =   9945
         TabIndex        =   7
         Top             =   315
         Width           =   270
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Exit"
      Height          =   1000
      Left            =   10830
      Picture         =   "frmSemenReport.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8490
      Width           =   1100
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   7500
      Top             =   8460
   End
   Begin VB.Frame Frame3 
      Caption         =   "Report"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5145
      Left            =   120
      TabIndex        =   1
      Top             =   3225
      Width           =   11775
      Begin RichTextLib.RichTextBox txtReport 
         Height          =   4695
         Left            =   90
         TabIndex        =   2
         Top             =   315
         Width           =   11475
         _ExtentX        =   20241
         _ExtentY        =   8281
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"frmSemenReport.frx":030A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.CommandButton cmdPrint 
      Appearance      =   0  'Flat
      Caption         =   "&Print"
      Height          =   1000
      Left            =   9615
      Picture         =   "frmSemenReport.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8490
      Width           =   1100
   End
   Begin MSComctlLib.ProgressBar PBar 
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSFlexGridLib.MSFlexGrid grdSID 
      Height          =   1935
      Left            =   120
      TabIndex        =   6
      Top             =   1260
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   3413
      _Version        =   393216
      Cols            =   5
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      FormatString    =   $"frmSemenReport.frx":0694
   End
End
Attribute VB_Name = "frmSemenReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Activated As Boolean

Private Type ABResult
    Antibiotic As String
    ReportName As String
    Result(1 To 8) As String
    Report(1 To 8) As Boolean
    Group(1 To 8) As String
    Qualifier(1 To 8) As String
End Type


Private pPatName As String
Private pPatChart As String
Private pPatDoB As String
Private pPatSex As String
Private pPatWard As String
Private pPatClinician As String
Private pPatGP As String

Private SortOrder As Boolean

Private Type LineText
    LineType As String
    LineText As String
End Type
Private udtPL() As LineText
Private udtPR() As LineText
Private ABExists As Boolean



Private Sub FillGrid()

          Dim sql As String
          Dim tb As Recordset
          Dim s As String


10        On Error GoTo FillGrid_Error



20        With grdSID
30            .Rows = 2
40            .AddItem ""
50            .RemoveItem 1
60        End With

70        sql = "SELECT * from Demographics where "
80        sql = sql & "PatName = '" & AddTicks(pPatName) & "' and "
90        If Trim$(pPatChart) <> "" Then
100           sql = sql & "Chart = '" & pPatChart & "' and "
110       Else
120           sql = sql & "(Chart = '' OR Chart IS NULL) and "
130       End If
140       If Trim$(pPatSex) <> "" Then
150           sql = sql & "Sex = '" & pPatSex & "' and "
160       Else
170           sql = sql & "(Sex = '' OR Sex IS NULL) and "
180       End If
190       If IsDate(pPatDoB) Then
200           sql = sql & "DoB = '" & Format(pPatDoB, "yyyy/mm/dd") & "' "
210       Else
220           sql = sql & "(COALESCE(DoB, '')  = '') "
230       End If
240       sql = sql & " AND SampleID BETWEEN " & SysOptSemenOffset(0) & " AND " & SysOptMicroOffset(0)
250       sql = sql & "ORDER BY RunDate desc"


260       Set tb = New Recordset
270       RecOpenClient 0, tb, sql


280       Do While Not tb.EOF


290           If Trim(tb!Hyear & "") = "" Then
300               If Val(tb!SampleID) > SysOptSemenOffset(0) Then

310                   s = Format$(Val(tb!SampleID) - SysOptSemenOffset( _
                                  0)) & vbTab & tb!Rundate & vbTab
320                   If IsDate(tb!SampleDate) Then
330                       If Format(tb!SampleDate, "hh:mm") <> "00:00" Then
340                           s = s & Format(tb!SampleDate, "dd/MM/yyyy hh:mm")
350                       Else
360                           s = s & Format(tb!SampleDate, "dd/MM/yyyy")
370                       End If
380                   Else
390                       s = s & "Not Specified"
400                   End If

410                   s = s & vbTab

420                   If IsDate(tb!SampleDate) Then
430                       s = s & CalcAge(pPatDoB, tb!SampleDate)
440                   End If
450                   s = s & vbTab
460                   s = s & tb!Addr0 & " " & tb!Addr1 & ""
470                   grdSID.AddItem s


480               End If
490           End If
500           tb.MoveNext
510       Loop

520       If grdSID.Rows > 2 Then
530           grdSID.RemoveItem 1
540       End If

550       Exit Sub

FillGrid_Error:

          Dim strES As String
          Dim intEL As Integer



560       intEL = Erl
570       strES = Err.Description
580       LogError "frmMicroReport", "FillGrid", intEL, strES, sql


End Sub

Private Sub cmdCancel_Click()

10        Unload Me

End Sub

Private Sub cmdPrint_Click()
10        PrintThis
End Sub

Private Sub Form_Activate()

10        On Error GoTo Form_Activate_Error


20        PBar.Max = LogOffDelaySecs
30        PBar = 0

40        Timer1.Enabled = True

50        If Not Activated Then
60            Activated = True
70            FillGrid
80        End If

90        Exit Sub

Form_Activate_Error:

          Dim strES As String
          Dim intEL As Integer



100       intEL = Erl
110       strES = Err.Description
120       LogError "frmMicroReport", "Form_Activate", intEL, strES


End Sub

Private Sub Form_Deactivate()

10        On Error GoTo Form_Deactivate_Error

20        Timer1.Enabled = False

30        Exit Sub

Form_Deactivate_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmMicroReport", "Form_Deactivate", intEL, strES


End Sub

Private Sub Form_Load()

10        On Error GoTo Form_Load_Error

20        Activated = False

30        PBar.Max = LogOffDelaySecs
40        PBar = 0

50        Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer



60        intEL = Erl
70        strES = Err.Description
80        LogError "frmMicroReport", "Form_Load", intEL, strES


End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, _
                           Y As Single)

10        On Error GoTo Form_MouseMove_Error

20        PBar = 0

30        Exit Sub

Form_MouseMove_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmMicroReport", "Form_MouseMove", intEL, strES


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
60        LogError "frmMicroReport", "Form_Unload", intEL, strES


End Sub

Private Sub grdSID_Click()

          Dim X As Long
          Dim Y As Long

10        On Error GoTo grdSID_Click_Error

20        txtReport = ""

30        If grdSID.MouseRow = 0 Then
40            If grdSID.MouseCol = 1 Or grdSID.MouseCol = 2 Then
50                grdSID.Col = grdSID.MouseCol
60                grdSID.Sort = 9
70            Else
80                If SortOrder Then
90                    grdSID.Sort = flexSortGenericAscending
100               Else
110                   grdSID.Sort = flexSortGenericDescending
120               End If
130           End If
140           SortOrder = Not SortOrder
150           Exit Sub
160       End If

170       For Y = 1 To grdSID.Rows - 1
180           grdSID.Row = Y
190           For X = 1 To grdSID.Cols - 1
200               grdSID.Col = X
210               grdSID.CellBackColor = 0
220           Next
230       Next

240       grdSID.Row = grdSID.MouseRow
250       For X = 1 To grdSID.Cols - 1
260           grdSID.Col = X
270           grdSID.CellBackColor = vbYellow
280       Next

290       FillResultSemen Val(grdSID.TextMatrix(grdSID.Row, _
                                                0)) + SysOptSemenOffset(0)

300       Exit Sub

grdSID_Click_Error:

          Dim strES As String
          Dim intEL As Integer

310       intEL = Erl
320       strES = Err.Description
330       LogError "frmMicroReport", "grdSID_Click", intEL, strES

End Sub

Private Sub grdSID_Compare(ByVal Row1 As Long, ByVal Row2 As Long, _
                           Cmp As Integer)

          Dim d1 As String
          Dim d2 As String

10        If Not IsDate(grdSID.TextMatrix(Row1, grdSID.Col)) Then
20            Cmp = 0
30            Exit Sub
40        End If

50        If Not IsDate(grdSID.TextMatrix(Row2, grdSID.Col)) Then
60            Cmp = 0
70            Exit Sub
80        End If

90        d1 = Format(grdSID.TextMatrix(Row1, grdSID.Col), "dd/mmm/yyyy hh:mm:ss")
100       d2 = Format(grdSID.TextMatrix(Row2, grdSID.Col), "dd/mmm/yyyy hh:mm:ss")

110       If SortOrder Then
120           Cmp = Sgn(DateDiff("s", d1, d2))
130       Else
140           Cmp = Sgn(DateDiff("s", d2, d1))
150       End If

End Sub

Private Sub FillResultSemen(ByVal SampleIDWithOffset As Double)

          Dim sql As String
          Dim tb As Recordset
          Dim s As String
          Dim pDefault As Long

10        On Error GoTo FillResultSemen_Error


20        sql = "SELECT D.Valid, D.ClDetails FROM Demographics D " & _
                "WHERE D.SampleID = " & SampleIDWithOffset
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        pDefault = 3

60        If tb.EOF Then Exit Sub


70        If IsNull(tb!Valid) Or tb!Valid = 0 Then
80            PrintTextRTB txtReport, "DEMOGRAPHICS NOT VALIDATED" & vbCrLf, 10, True, , True, vbRed
90        End If

100       If Trim$(tb!ClDetails & "") <> "" Then
110           PrintTextRTB txtReport, "Clinical Details: ", 10
120           PrintTextRTB txtReport, tb!ClDetails & "" & vbCrLf, 10, True
130       End If

140       GetPrintLineSemen SampleIDWithOffset
150       PrintTextRTB txtReport, vbCrLf
160       GetPrintLineComments SampleIDWithOffset, "Demographic Comment: ", "Demographic"
170       GetPrintLineComments SampleIDWithOffset, "Semen Analysis Comment: ", "Semen"
180       PrintTextRTB txtReport, ".________________________End of Report________________________." & vbCrLf

190       Exit Sub

FillResultSemen_Error:

          Dim strES As String
          Dim intEL As Integer

200       intEL = Erl
210       strES = Err.Description
220       LogError "frmMicroReport", "FillResultSemen", intEL, strES, sql

End Sub








Private Sub Timer1_Timer()

      'tmrRefresh.Interval set to 1000
10        On Error GoTo Timer1_Timer_Error

20        PBar = PBar + 1

30        If PBar = PBar.Max Then
40            Unload Me
50        End If

60        Exit Sub

Timer1_Timer_Error:

          Dim strES As String
          Dim intEL As Integer



70        intEL = Erl
80        strES = Err.Description
90        LogError "frmMicroReport", "Timer1_Timer", intEL, strES


End Sub

Public Property Let PatName(ByVal NewValue As String)

10        pPatName = NewValue
20        lblName = NewValue

End Property

Public Property Let PatChart(ByVal NewValue As String)

10        pPatChart = NewValue
20        lblChart = NewValue

End Property

Public Property Let PatDoB(ByVal NewValue As String)

10        pPatDoB = NewValue
20        lblDoB = NewValue

End Property

Public Property Let PatSex(ByVal NewValue As String)

10        pPatSex = NewValue
20        lblSex = NewValue

End Property

Public Property Let PatWard(ByVal NewValue As String)
10        pPatWard = NewValue
End Property

Public Property Get PatWard() As String
10        PatWard = pPatWard
End Property

Public Property Let PatClinician(ByVal NewValue As String)
10        pPatClinician = NewValue
End Property

Public Property Get PatClinician() As String
10        PatClinician = pPatClinician
End Property

Public Property Let PatGP(ByVal NewValue As String)
10        pPatGP = NewValue
End Property

Public Property Get PatGP() As String
10        PatGP = pPatGP
End Property





Private Sub GetPrintLineComments(SampleIDWithOffset As Double, ByVal CommentTitle As String, _
                                 ByVal FieldName As String)

10        On Error GoTo GetPrintLineComments_Error

20        ReDim Comments(1 To 8) As String
          Dim sql As String
          Dim tb As Recordset
          Dim n As Integer
          Dim lpc As Integer
          Dim Ob As Observation
          Dim Obs As New Observations

30        Set Obs = Obs.Load(SampleIDWithOffset, FieldName)

40        If Not Obs Is Nothing Then
50            For Each Ob In Obs
60                PrintTextRTB txtReport, CommentTitle, 9, True
70                FillCommentLines Ob.Comment, 8, Comments(), 80
80                For n = 1 To 8
90                    If Trim(Comments(n) & "") <> "" Then
100                       PrintTextRTB txtReport, FormatString(Comments(n), 80, , AlignLeft) & vbCrLf, 9
110                   End If
120               Next
130           Next
140       End If

150       Exit Sub

GetPrintLineComments_Error:

          Dim strES As String
          Dim intEL As Integer

160       intEL = Erl
170       strES = Err.Description
180       LogError "modNewMicro", "GetPrintLineComments", intEL, strES, sql

End Sub

Public Sub FillCommentLines(ByVal FullComment As String, _
                            ByVal NumberOfLines As Integer, _
                            ByRef Comments() As String, _
                            Optional ByVal MaxLen As Integer = 80)

          Dim n As Integer
          Dim CurrentLine As Integer
          Dim X As Integer
          Dim ThisLine As String
          Dim SpaceFound As Boolean

10        On Error GoTo FillCommentLines_Error

20        For n = 1 To UBound(Comments)
30            Comments(n) = ""
40        Next

50        CurrentLine = 0
60        FullComment = Trim(FullComment)
70        n = Len(FullComment)

80        For X = n - 1 To 1 Step -1
90            If Mid(FullComment, X, 1) = vbCr Or Mid(FullComment, X, 1) = vbLf Or Mid(FullComment, X, 1) = vbTab Then
100               Mid(FullComment, X, 1) = " "
110           End If
120       Next

130       For X = n - 3 To 1 Step -1
140           If Mid(FullComment, X, 2) = "  " Then
150               FullComment = Left(FullComment, X) & Mid(FullComment, X + 2)
160           End If
170       Next
180       n = Len(FullComment)

190       Do While n > MaxLen
200           SpaceFound = False
210           For X = MaxLen To 1 Step -1
220               If Mid(FullComment, X, 1) = " " Then
230                   ThisLine = Left(FullComment, X - 1)
240                   FullComment = Mid(FullComment, X + 1)

250                   CurrentLine = CurrentLine + 1
260                   If CurrentLine <= NumberOfLines Then
270                       Comments(CurrentLine) = ThisLine
280                   End If
290                   SpaceFound = True
300                   Exit For
310               End If
320           Next
330           If Not SpaceFound Then
340               ThisLine = Left(FullComment, MaxLen)
350               FullComment = Mid(FullComment, MaxLen + 1)

360               CurrentLine = CurrentLine + 1
370               If CurrentLine <= NumberOfLines Then
380                   Comments(CurrentLine) = ThisLine
390               End If
400           End If
410           n = Len(FullComment)
420       Loop

430       CurrentLine = CurrentLine + 1
440       If CurrentLine <= NumberOfLines Then
450           Comments(CurrentLine) = FullComment
460       End If

470       Exit Sub

FillCommentLines_Error:

          Dim strES As String
          Dim intEL As Integer

480       intEL = Erl
490       strES = Err.Description
500       LogError "Other", "FillCommentLines", intEL, strES

End Sub

Private Sub PrintThis()

          Dim tb As Recordset
          Dim sql As String
          Dim SampleIDWithOffset As Double

10        On Error GoTo PrintThis_Error

20        PBar = 0

30        If grdSID.TextMatrix(grdSID.Row, 0) = "" Then Exit Sub

40        SampleIDWithOffset = grdSID.TextMatrix(grdSID.Row, 0) + SysOptSemenOffset(0)

50        sql = "Select * from PrintPending where " & _
                "Department = 'N' " & _
                "and SampleID = " & SampleIDWithOffset
60        Set tb = New Recordset
70        RecOpenClient 0, tb, sql
80        If tb.EOF Then
90            tb.AddNew
100       End If
110       tb!SampleID = SampleIDWithOffset
120       tb!Ward = PatWard
130       tb!Clinician = PatClinician
140       tb!GP = PatGP
150       tb!Department = "Z"
160       tb!Initiator = UserName
170       tb!UsePrinter = ""   'pPrintToPrinter
180       tb!NoOfCopies = 1
190       tb!FinalInterim = "F"
200       tb!pTime = Now
210       tb.Update

220       Exit Sub

PrintThis_Error:

          Dim strES As String
          Dim intEL As Integer

230       intEL = Erl
240       strES = Err.Description
250       LogError "frmMicroReport", "PrintThis", intEL, strES, sql

End Sub

Public Sub GetPrintLineSemen(ByVal SampleIDWithOffset As String)

          Dim tb As Recordset
          Dim sql As String
          Dim n As Integer
          Dim lpc As Integer
          Dim ColName As String
          Dim ShortName As String
          Dim Units As String
          Dim TestNameLength As Integer
          Dim ResultLength As Integer
          Dim SemenUsername As String


10        On Error GoTo GetPrintLineSemen_Error

20        TestNameLength = 20
30        ResultLength = 22

40        sql = "SELECT * FROM SemenResults WHERE " & _
                "SampleID = '" & SampleIDWithOffset & "'"
50        Set tb = New Recordset
60        RecOpenServer 0, tb, sql

70        If Not tb.EOF Then
80            If Trim$(tb!UserName & "") <> "" Then
90                SemenUsername = tb!UserName
100           End If
110           PrintTextRTB txtReport, "Semen Analysis" & vbCrLf & vbCrLf, 10, True, , True
120           For n = 1 To 8
130               ColName = Choose(n, "Volume", "Consistency", "SemenCount", _
                                   "Motility", "MotilityPro", "MotilityNonPro", _
                                   "MotilitySlow", "MotilityNonMotile")
140               ShortName = Choose(n, "Volume", "Consistency", "Spermatozoa Count", _
                                     "Motile", "Progressive", "Non Progressive", _
                                     "Slow Progressive", "Non Motile")
150               Units = Choose(n, "mL", "", "Million per mL", _
                                 "%", "%", "%", "%", "%")
160               If n = 3 And InStr(UCase(tb!SemenCount & ""), "SEEN") Then
170                   Units = ""
180               End If


190               If Trim$(tb(ColName) & "") <> "" Then
200                   PrintTextRTB txtReport, FormatString(ShortName, TestNameLength, ":"), 10
210                   PrintTextRTB txtReport, FormatString(tb(ColName) & " " & Units, ResultLength), 10, True
220                   PrintTextRTB txtReport, vbCrLf
230               End If
240           Next
250       End If

260       sql = "SELECT " & _
                "R = ( SELECT  Result FROM GenericResults WHERE " & _
                "      TestName = 'SemenMorphResult' " & _
                "      AND (SampleID = '" & SampleIDWithOffset & "' )), " & _
                "D = ( SELECT  Result FROM GenericResults WHERE " & _
                "      TestName = 'SemenMorphDescription' " & _
                "      AND (SampleID = '" & SampleIDWithOffset & "' )) "
270       Set tb = New Recordset
280       RecOpenServer 0, tb, sql
290       If Not tb.EOF Then
300           If Not IsNull(tb!R) And Not IsNull(tb!D) Then
310               PrintTextRTB txtReport, FormatString("Morphology", TestNameLength, ":"), 10
320               PrintTextRTB txtReport, FormatString(tb!R & " " & tb!D, ResultLength), 10, True
330               PrintTextRTB txtReport, vbCrLf
340           End If
350       End If



360       Exit Sub

GetPrintLineSemen_Error:

          Dim strES As String
          Dim intEL As Integer

370       intEL = Erl
380       strES = Err.Description
390       LogError "frmSemenReport", "GetPrintLineSemen", intEL, strES, sql

End Sub

