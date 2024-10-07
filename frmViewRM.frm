VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "mschrt20.ocx"
Begin VB.Form frmViewRM 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - View Running Means"
   ClientHeight    =   5295
   ClientLeft      =   3360
   ClientTop       =   2385
   ClientWidth     =   6630
   Icon            =   "frmViewRM.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5295
   ScaleWidth      =   6630
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Caption         =   "Parameter"
      Height          =   720
      Left            =   1770
      TabIndex        =   8
      Top             =   75
      Width           =   2085
      Begin VB.ComboBox lstParameter 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   135
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   270
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Department"
      Height          =   1455
      Left            =   90
      TabIndex        =   5
      Top             =   75
      Width           =   1620
      Begin VB.OptionButton optEnd 
         Caption         =   "Endocrinology"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   1215
         Width           =   1395
      End
      Begin VB.OptionButton optCoag 
         Caption         =   "Coagulation"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton optImm 
         Caption         =   "Immunology"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton optBio 
         Caption         =   "Biochemistry"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton optHaem 
         Caption         =   "Haematology"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   225
         Value           =   -1  'True
         Width           =   1305
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Data Points"
      Height          =   1095
      Left            =   3915
      TabIndex        =   2
      Top             =   45
      Width           =   1155
      Begin ComCtl2.UpDown UpDown1 
         Height          =   405
         Left            =   810
         TabIndex        =   3
         Top             =   450
         Width           =   240
         _ExtentX        =   450
         _ExtentY        =   714
         _Version        =   327681
         Value           =   100
         BuddyControl    =   "lblDataPoints"
         BuddyDispid     =   196618
         OrigLeft        =   240
         OrigTop         =   660
         OrigRight       =   960
         OrigBottom      =   915
         Increment       =   10
         Max             =   500
         Min             =   50
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65537
         Enabled         =   -1  'True
      End
      Begin VB.Label lblDataPoints 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "100"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   540
         Width           =   510
      End
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   645
      Left            =   5190
      Picture         =   "frmViewRM.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   330
      Width           =   1245
   End
   Begin MSChart20Lib.MSChart g 
      Height          =   3615
      Left            =   120
      OleObjectBlob   =   "frmViewRM.frx":0614
      TabIndex        =   1
      Top             =   1575
      Visible         =   0   'False
      Width           =   6435
   End
End
Attribute VB_Name = "frmViewRM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub DrawHaemGraph()

          Dim rs As New Recordset
          Dim sql As String
          Dim X As Long

10        On Error GoTo DrawHaemGraph_Error

20        ReDim D(1 To lblDataPoints, 1 To 1) As Variant
          Dim maxy As Single
          Dim miny As Single
          Dim rm As Single
          Dim fld As Field
          Dim Run(500) As Single
          Dim varTot(500) As Single
          Dim varMax As Single
          Dim varMin As Single
          Dim varRm(500) As Single
          Dim varNo As Long
          Dim n As Long


30        If lstParameter = "" Then Exit Sub

40        If Val(lblDataPoints) < 50 Then
50            lblDataPoints = "50"
60        ElseIf Val(lblDataPoints) > 500 Then
70            lblDataPoints = "500"
80        End If

90        g.Visible = False

100       sql = "SELECT TOP " & lblDataPoints & " " & lstParameter & _
                " FROM HaemResults  WHERE " & lstParameter & " <> '' " & _
                "ORDER BY rundatetime DESC"

110       Set rs = New Recordset
120       RecOpenClient 0, rs, sql

130       maxy = 0
140       miny = 999
150       n = 1

160       Do While Not rs.EOF
170           Run(n) = Val(rs("" & lstParameter & ""))
180           n = n + 1
190           rs.MoveNext
200       Loop
210       varMin = 10000
220       varMax = 0
230       For n = 1 To (Val(lblDataPoints) - 21)
240           For X = n To n + 21
250               If varMin > Run(X) Then varMin = Run(X)
260               If varMax < Run(X) Then varMax = Run(X)
270               varTot(n) = varTot(n) + Run(X)
280           Next
290           varTot(n) = varTot(n) - (varMin + varMax)
300           varTot(n) = varTot(n) / 20
310           If n = 1 Then
320               varRm(n) = varTot(n)
330           Else
340               varRm(n) = (varTot(n) * 0.05) + (varRm(n - 1) * 0.95)
350           End If
360           varMin = 10000
370           varMax = 0
380           varNo = n
390       Next

400       miny = 9999
410       maxy = 0
420       X = varNo - 1
430       For n = 1 To varNo
440           X = X - 1
450           If X = 0 Then Exit For
460           D(X, 1) = varRm(n)
470           If varRm(n) < miny Then
480               miny = varRm(n)
490           End If
500           If varRm(n) > maxy Then
510               maxy = varRm(n)
520           End If
530       Next

540       With g.Plot.Axis(VtChAxisIdY).ValueScale
550           .Auto = False
560           .Maximum = Int(maxy * 1.05 + 0.5)
570           If .Maximum = 0 Then
580               .Maximum = 1
590           End If
600           If maxy > .Maximum Then
610               .Maximum = .Maximum + 1
620           End If
630           .minimum = Int(miny * 0.95 - 0.5)
640           If .minimum < 0 Then
650               .minimum = 0
660           End If
670       End With

680       g.ChartData = D
690       g.Visible = True





700       Exit Sub

DrawHaemGraph_Error:

          Dim strES As String
          Dim intEL As Integer


710       intEL = Erl
720       strES = Err.Description
730       LogError "frmViewRM", "DrawHaemGraph", intEL, strES, sql


End Sub



Private Sub DrawBioGraph()

          Dim rs As New Recordset
          Dim sql As String
          Dim X As Long
10        On Error GoTo DrawBioGraph_Error

20        ReDim D(1 To lblDataPoints, 1 To 1) As Variant
          Dim maxy As Single
          Dim miny As Single
          Dim rm As Single
          Dim fld As Field
          Dim Run(500) As Single
          Dim varTot(500) As Single
          Dim varMax As Single
          Dim varMin As Single
          Dim varRm(500) As Single
          Dim varNo As Long
          Dim n As Long
          Dim ParameterCode As String



30        If lstParameter = "" Then Exit Sub

40        ParameterCode = CodeForLongName(lstParameter.Text)

50        If Val(lblDataPoints) < 50 Then
60            lblDataPoints = "50"
70        ElseIf Val(lblDataPoints) > 500 Then
80            lblDataPoints = "500"
90        End If

100       g.Visible = False

110       sql = "SELECT TOP " & lblDataPoints & " Result " & _
                "FROM BioResults WHERE Code = '" & ParameterCode & "' and Result <> '' " & _
                "ORDER BY runtime DESC"

120       Set rs = New Recordset
130       RecOpenClient 0, rs, sql

140       maxy = 0
150       miny = 999
160       n = 1

170       Do While Not rs.EOF
180           Run(n) = Val(rs!Result)
190           n = n + 1
200           rs.MoveNext
210       Loop
220       varMin = 10000
230       varMax = 0
240       For n = 1 To (Val(lblDataPoints) - 21)
250           For X = n To n + 21
260               If varMin > Run(X) Then varMin = Run(X)
270               If varMax < Run(X) Then varMax = Run(X)
280               varTot(n) = varTot(n) + Run(X)
290           Next
300           varTot(n) = varTot(n) - (varMin + varMax)
310           varTot(n) = varTot(n) / 20
320           If n = 1 Then
330               varRm(n) = varTot(n)
340           Else
350               varRm(n) = (varTot(n) * 0.05) + (varRm(n - 1) * 0.95)
360           End If
370           varMin = 10000
380           varMax = 0
390           varNo = n
400       Next

410       miny = 9999
420       maxy = 0
430       X = varNo - 1
440       For n = 1 To varNo
450           X = X - 1
460           If X = 0 Then Exit For
470           D(X, 1) = varRm(n)
480           If varRm(n) < miny Then
490               miny = varRm(n)
500           End If
510           If varRm(n) > maxy Then
520               maxy = varRm(n)
530           End If
540       Next

550       With g.Plot.Axis(VtChAxisIdY).ValueScale
560           .Auto = False
570           .Maximum = Int(maxy * 1.05 + 0.5)
580           If .Maximum = 0 Then
590               .Maximum = 1
600           End If
610           If maxy > .Maximum Then
620               .Maximum = .Maximum + 1
630           End If
640           .minimum = Int(miny * 0.95 - 0.5)
650           If .minimum < 0 Then
660               .minimum = 0
670           End If
680       End With

690       g.ChartData = D
700       g.Visible = True




710       Exit Sub

DrawBioGraph_Error:

          Dim strES As String
          Dim intEL As Integer


720       intEL = Erl
730       strES = Err.Description
740       LogError "frmViewRM", "DrawBioGraph", intEL, strES, sql


End Sub

Private Sub FillParameterList()

10        On Error GoTo FillParameterList_Error

20        If optHaem Then
30            FillWithHaem
40        ElseIf optBio Then
50            FillWithBio
60        ElseIf optEnd Then
70            FillWithEnd
80        ElseIf optCoag Then
90            FillWithCoag
100       ElseIf optImm Then
110           FillWithImm
120       End If

130       Exit Sub

FillParameterList_Error:

          Dim strES As String
          Dim intEL As Integer


140       intEL = Erl
150       strES = Err.Description
160       LogError "frmViewRM", "FillParameterList", intEL, strES


End Sub
Private Sub DrawImmGraph()

          Dim rs As New Recordset
          Dim sql As String
          Dim X As Long

10        On Error GoTo DrawImmGraph_Error

20        ReDim D(1 To lblDataPoints, 1 To 1) As Variant
          Dim maxy As Single
          Dim miny As Single
          Dim rm As Single
          Dim fld As Field
          Dim Run(500) As Single
          Dim varTot(500) As Single
          Dim varMax As Single
          Dim varMin As Single
          Dim varRm(500) As Single
          Dim varNo As Long
          Dim n As Long
          Dim ParameterCode As String


30        If lstParameter = "" Then Exit Sub

40        ParameterCode = iCodeForLongName(lstParameter.Text)

50        If Val(lblDataPoints) < 50 Then
60            lblDataPoints = "50"
70        ElseIf Val(lblDataPoints) > 500 Then
80            lblDataPoints = "500"
90        End If

100       g.Visible = False

110       sql = "SELECT TOP " & lblDataPoints & " Result " & _
                "FROM ImmResults WHERE Code = '" & ParameterCode & "' " & _
                "ORDER BY runtime DESC"

120       Set rs = New Recordset
130       RecOpenClient 0, rs, sql

140       maxy = 0
150       miny = 999
160       n = 1

170       Do While Not rs.EOF
180           Run(n) = Val(rs!Result)
190           n = n + 1
200           rs.MoveNext
210       Loop
220       varMin = 10000
230       varMax = 0
240       For n = 1 To (Val(lblDataPoints) - 21)
250           For X = n To n + 21
260               If varMin > Run(X) Then varMin = Run(X)
270               If varMax < Run(X) Then varMax = Run(X)
280               varTot(n) = varTot(n) + Run(X)
290           Next
300           varTot(n) = varTot(n) - (varMin + varMax)
310           varTot(n) = varTot(n) / 20
320           If n = 1 Then
330               varRm(n) = varTot(n)
340           Else
350               varRm(n) = (varTot(n) * 0.05) + (varRm(n - 1) * 0.95)
360           End If
370           varMin = 10000
380           varMax = 0
390           varNo = n
400       Next

410       miny = 9999
420       maxy = 0
430       X = varNo - 1
440       For n = 1 To varNo
450           X = X - 1
460           If X = 0 Then Exit For
470           D(X, 1) = varRm(n)
480           If varRm(n) < miny Then
490               miny = varRm(n)
500           End If
510           If varRm(n) > maxy Then
520               maxy = varRm(n)
530           End If
540       Next

550       With g.Plot.Axis(VtChAxisIdY).ValueScale
560           .Auto = False
570           .Maximum = Int(maxy * 1.05 + 0.5)
580           If .Maximum = 0 Then
590               .Maximum = 1
600           End If
610           If maxy > .Maximum Then
620               .Maximum = .Maximum + 1
630           End If
640           .minimum = Int(miny * 0.95 - 0.5)
650           If .minimum < 0 Then
660               .minimum = 0
670           End If
680       End With

690       g.ChartData = D
700       g.Visible = True



710       Exit Sub

DrawImmGraph_Error:

          Dim strES As String
          Dim intEL As Integer


720       intEL = Erl
730       strES = Err.Description
740       LogError "frmViewRM", "DrawImmGraph", intEL, strES, sql


End Sub




Private Sub FillWithHaem()

          Dim sql As String
          Dim tb As Recordset

10        On Error GoTo FillWithHaem_Error

20        lstParameter.Clear

30        sql = "select distinct analytename from haemtestdefinitions"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        Do While Not tb.EOF
70            lstParameter.AddItem tb!AnalyteName
80            tb.MoveNext
90        Loop

100       Exit Sub

FillWithHaem_Error:

          Dim strES As String
          Dim intEL As Integer

110       intEL = Erl
120       strES = Err.Description
130       LogError "frmViewRM", "FillWithHaem", intEL, strES, sql

End Sub

Private Sub FillWithBio()

          Dim InList As Boolean
          Dim n As Long
          Dim sql As String
          Dim tb As New Recordset

10        On Error GoTo FillWithBio_Error

20        lstParameter.Clear

30        sql = "SELECT * from biotestdefinitions"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        Do While Not tb.EOF
70            InList = False
80            For n = 0 To lstParameter.ListCount - 1
90                If lstParameter.List(n) = tb!LongName Then
100                   InList = True
110                   Exit For
120               End If
130           Next
140           If Not InList Then
150               lstParameter.AddItem tb!LongName
160           End If
170           tb.MoveNext
180       Loop

190       Exit Sub

FillWithBio_Error:

          Dim strES As String
          Dim intEL As Integer


200       intEL = Erl
210       strES = Err.Description
220       LogError "frmViewRM", "FillWithBio", intEL, strES, sql

End Sub

Private Sub FillWithEnd()

          Dim InList As Boolean
          Dim n As Long
          Dim sql As String
          Dim tb As New Recordset

10        On Error GoTo FillWithEnd_Error

20        lstParameter.Clear

30        sql = "SELECT * from Endtestdefinitions"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        Do While Not tb.EOF
70            InList = False
80            For n = 0 To lstParameter.ListCount - 1
90                If lstParameter.List(n) = tb!LongName Then
100                   InList = True
110                   Exit For
120               End If
130           Next
140           If Not InList Then
150               lstParameter.AddItem tb!LongName
160           End If
170           tb.MoveNext
180       Loop

190       Exit Sub

FillWithEnd_Error:

          Dim strES As String
          Dim intEL As Integer


200       intEL = Erl
210       strES = Err.Description
220       LogError "frmViewRM", "FillWithEnd", intEL, strES, sql

End Sub
Private Sub FillWithImm()

          Dim tb As New Recordset
          Dim sql As String
          Dim InList As Boolean
          Dim n As Long

10        On Error GoTo FillWithImm_Error

20        lstParameter.Clear

30        sql = "SELECT * from ImmTestDefinitions " & _
                "Order by PrintPriority"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        Do While Not tb.EOF
70            InList = False
80            For n = 0 To lstParameter.ListCount - 1
90                If lstParameter.List(n) = tb!LongName Then
100                   InList = True
110                   Exit For
120               End If
130           Next
140           If Not InList Then
150               lstParameter.AddItem tb!LongName
160           End If
170           tb.MoveNext
180       Loop

190       Exit Sub

FillWithImm_Error:

          Dim strES As String
          Dim intEL As Integer


200       intEL = Erl
210       strES = Err.Description
220       LogError "frmViewRM", "FillWithImm", intEL, strES, sql


End Sub

Private Sub FillWithCoag()

          Dim tb As New Recordset
          Dim InList As Boolean
          Dim n As Long
          Dim sql As String

10        On Error GoTo FillWithCoag_Error

20        lstParameter.Clear

30        sql = "SELECT * from coagtestdefinitions"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql

60        Do While Not tb.EOF
70            InList = False
80            For n = 0 To lstParameter.ListCount - 1
90                If lstParameter.List(n) = tb!TestName Then
100                   InList = True
110                   Exit For
120               End If
130           Next
140           If Not InList Then
150               lstParameter.AddItem tb!TestName
160           End If
170           tb.MoveNext
180       Loop

190       Exit Sub

FillWithCoag_Error:

          Dim strES As String
          Dim intEL As Integer


200       intEL = Erl
210       strES = Err.Description
220       LogError "frmViewRM", "FillWithCoag", intEL, strES, sql


End Sub

Private Sub cmdExit_Click()

10        Unload Me

End Sub

Private Sub Form_Load()

10        On Error GoTo Form_Load_Error

20        FillParameterList

30        Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer


40        intEL = Erl
50        strES = Err.Description
60        LogError "frmViewRM", "Form_Load", intEL, strES


End Sub

Private Sub lstParameter_Click()

10        On Error GoTo lstParameter_Click_Error

20        If optHaem Then
30            DrawHaemGraph
40        ElseIf optBio Then
50            DrawBioGraph
60        ElseIf optEnd Then
70            DrawEndGraph
80        ElseIf optCoag Then
90            DrawCoagGraph
100       ElseIf optImm Then
110           DrawImmGraph
120       End If

130       Exit Sub

lstParameter_Click_Error:

          Dim strES As String
          Dim intEL As Integer


140       intEL = Erl
150       strES = Err.Description
160       LogError "frmViewRM", "lstParameter_Click", intEL, strES


End Sub

Private Sub optBio_Click()

10        On Error GoTo optBio_Click_Error

20        FillParameterList

30        Exit Sub

optBio_Click_Error:

          Dim strES As String
          Dim intEL As Integer


40        intEL = Erl
50        strES = Err.Description
60        LogError "frmViewRM", "optBio_Click", intEL, strES


End Sub

Private Sub optCoag_Click()

10        On Error GoTo optCoag_Click_Error

20        FillParameterList

30        Exit Sub

optCoag_Click_Error:

          Dim strES As String
          Dim intEL As Integer


40        intEL = Erl
50        strES = Err.Description
60        LogError "frmViewRM", "optCoag_Click", intEL, strES


End Sub

Private Sub optEnd_Click()

10        On Error GoTo optEnd_Click_Error

20        FillParameterList

30        Exit Sub

optEnd_Click_Error:

          Dim strES As String
          Dim intEL As Integer


40        intEL = Erl
50        strES = Err.Description
60        LogError "frmViewRM", "optEnd_Click", intEL, strES


End Sub

Private Sub optHaem_Click()

10        On Error GoTo optHaem_Click_Error

20        FillParameterList

30        Exit Sub

optHaem_Click_Error:

          Dim strES As String
          Dim intEL As Integer


40        intEL = Erl
50        strES = Err.Description
60        LogError "frmViewRM", "optHaem_Click", intEL, strES


End Sub



Private Sub optImm_Click()

10        On Error GoTo optImm_Click_Error

20        FillParameterList

30        Exit Sub

optImm_Click_Error:

          Dim strES As String
          Dim intEL As Integer


40        intEL = Erl
50        strES = Err.Description
60        LogError "frmViewRM", "optImm_Click", intEL, strES


End Sub



Private Sub UpDown1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)


10        On Error GoTo UpDown1_MouseUp_Error

20        If optHaem Then
30            DrawHaemGraph
40        ElseIf optBio Then
50            DrawBioGraph
60        ElseIf optCoag Then
70            DrawCoagGraph
80        ElseIf optImm Then
90            DrawImmGraph
100       End If

110       Exit Sub

UpDown1_MouseUp_Error:

          Dim strES As String
          Dim intEL As Integer


120       intEL = Erl
130       strES = Err.Description
140       LogError "frmViewRM", "UpDown1_MouseUp", intEL, strES


End Sub



Private Sub DrawCoagGraph()

          Dim rs As New Recordset
          Dim sql As String
          Dim X As Long

10        On Error GoTo DrawCoagGraph_Error

20        ReDim D(1 To lblDataPoints, 1 To 1) As Variant
          Dim maxy As Single
          Dim miny As Single
          Dim rm As Single
          Dim fld As Field
          Dim Run(500) As Single
          Dim varTot(500) As Single
          Dim varMax As Single
          Dim varMin As Single
          Dim varRm(500) As Single
          Dim varNo As Long
          Dim n As Long
          Dim ParameterCode As String


30        If lstParameter = "" Then Exit Sub

40        ParameterCode = CoagCodeFor(lstParameter.Text)

50        If Val(lblDataPoints) < 50 Then
60            lblDataPoints = "50"
70        ElseIf Val(lblDataPoints) > 500 Then
80            lblDataPoints = "500"
90        End If

100       g.Visible = False

110       sql = "SELECT TOP " & lblDataPoints & " Result " & _
                "FROM CoagResults WHERE Code = '" & ParameterCode & "' and Result <> '' " & _
                "ORDER BY runtime DESC"

120       Set rs = New Recordset
130       RecOpenClient 0, rs, sql

140       maxy = 0
150       miny = 999
160       n = 1

170       Do While Not rs.EOF
180           Run(n) = Val(rs!Result)
190           n = n + 1
200           rs.MoveNext
210       Loop
220       varMin = 10000
230       varMax = 0
240       For n = 1 To (Val(lblDataPoints) - 21)
250           For X = n To n + 21
260               If varMin > Run(X) Then varMin = Run(X)
270               If varMax < Run(X) Then varMax = Run(X)
280               varTot(n) = varTot(n) + Run(X)
290           Next
300           varTot(n) = varTot(n) - (varMin + varMax)
310           varTot(n) = varTot(n) / 20
320           If n = 1 Then
330               varRm(n) = varTot(n)
340           Else
350               varRm(n) = (varTot(n) * 0.05) + (varRm(n - 1) * 0.95)
360           End If
370           varMin = 10000
380           varMax = 0
390           varNo = n
400       Next

410       miny = 9999
420       maxy = 0
430       X = varNo - 1
440       For n = 1 To varNo
450           X = X - 1
460           If X = 0 Then Exit For
470           D(X, 1) = varRm(n)
480           If varRm(n) < miny Then
490               miny = varRm(n)
500           End If
510           If varRm(n) > maxy Then
520               maxy = varRm(n)
530           End If
540       Next

550       With g.Plot.Axis(VtChAxisIdY).ValueScale
560           .Auto = False
570           .Maximum = Int(maxy * 1.05 + 0.5)
580           If .Maximum = 0 Then
590               .Maximum = 1
600           End If
610           If maxy > .Maximum Then
620               .Maximum = .Maximum + 1
630           End If
640           .minimum = Int(miny * 0.95 - 0.5)
650           If .minimum < 0 Then
660               .minimum = 0
670           End If
680       End With

690       g.ChartData = D
700       g.Visible = True





710       Exit Sub

DrawCoagGraph_Error:

          Dim strES As String
          Dim intEL As Integer


720       intEL = Erl
730       strES = Err.Description
740       LogError "frmViewRM", "DrawCoagGraph", intEL, strES, sql


End Sub


Private Sub DrawEndGraph()

          Dim rs As New Recordset
          Dim sql As String
          Dim X As Long

10        On Error GoTo DrawEndGraph_Error

20        ReDim D(1 To lblDataPoints, 1 To 1) As Variant
          Dim maxy As Single
          Dim miny As Single
          Dim rm As Single
          Dim fld As Field
          Dim Run(500) As Single
          Dim varTot(500) As Single
          Dim varMax As Single
          Dim varMin As Single
          Dim varRm(500) As Single
          Dim varNo As Long
          Dim n As Long
          Dim ParameterCode As String


30        If lstParameter = "" Then Exit Sub

40        ParameterCode = CodeForLongName(lstParameter.Text)

50        If Val(lblDataPoints) < 50 Then
60            lblDataPoints = "50"
70        ElseIf Val(lblDataPoints) > 500 Then
80            lblDataPoints = "500"
90        End If

100       g.Visible = False

110       sql = "SELECT TOP " & lblDataPoints & " Result " & _
                "FROM EndResults WHERE Code = '" & ParameterCode & "' and Result <> '' " & _
                "ORDER BY runtime DESC"

120       Set rs = New Recordset
130       RecOpenClient 0, rs, sql

140       maxy = 0
150       miny = 999
160       n = 1

170       Do While Not rs.EOF
180           Run(n) = Val(rs!Result)
190           n = n + 1
200           rs.MoveNext
210       Loop
220       varMin = 10000
230       varMax = 0
240       For n = 1 To (Val(lblDataPoints) - 21)
250           For X = n To n + 21
260               If varMin > Run(X) Then varMin = Run(X)
270               If varMax < Run(X) Then varMax = Run(X)
280               varTot(n) = varTot(n) + Run(X)
290           Next
300           varTot(n) = varTot(n) - (varMin + varMax)
310           varTot(n) = varTot(n) / 20
320           If n = 1 Then
330               varRm(n) = varTot(n)
340           Else
350               varRm(n) = (varTot(n) * 0.05) + (varRm(n - 1) * 0.95)
360           End If
370           varMin = 10000
380           varMax = 0
390           varNo = n
400       Next

410       miny = 9999
420       maxy = 0
430       X = varNo - 1
440       For n = 1 To varNo
450           X = X - 1
460           If X = 0 Then Exit For
470           D(X, 1) = varRm(n)
480           If varRm(n) < miny Then
490               miny = varRm(n)
500           End If
510           If varRm(n) > maxy Then
520               maxy = varRm(n)
530           End If
540       Next

550       With g.Plot.Axis(VtChAxisIdY).ValueScale
560           .Auto = False
570           .Maximum = Int(maxy * 1.05 + 0.5)
580           If .Maximum = 0 Then
590               .Maximum = 1
600           End If
610           If maxy > .Maximum Then
620               .Maximum = .Maximum + 1
630           End If
640           .minimum = Int(miny * 0.95 - 0.5)
650           If .minimum < 0 Then
660               .minimum = 0
670           End If
680       End With

690       g.ChartData = D
700       g.Visible = True



710       Exit Sub

DrawEndGraph_Error:

          Dim strES As String
          Dim intEL As Integer


720       intEL = Erl
730       strES = Err.Description
740       LogError "frmViewRM", "DrawEndGraph", intEL, strES, sql


End Sub


