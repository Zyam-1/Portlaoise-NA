VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "mschrt20.ocx"
Begin VB.Form frcControl 
   AutoRedraw      =   -1  'True
   ClientHeight    =   6345
   ClientLeft      =   960
   ClientTop       =   1980
   ClientWidth     =   10140
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
   Icon            =   "frcControl.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   423
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   676
   Begin MSChart20Lib.MSChart mscGraph 
      Height          =   5775
      Left            =   2970
      OleObjectBlob   =   "frcControl.frx":030A
      TabIndex        =   6
      Top             =   450
      Width           =   7035
   End
   Begin VB.CommandButton cmdPrint 
      Appearance      =   0  'Flat
      Caption         =   "&Print"
      Height          =   690
      Left            =   1530
      Picture         =   "frcControl.frx":1D78
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4680
      Width           =   1320
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   600
      Left            =   90
      Picture         =   "frcControl.frx":2082
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5400
      Width           =   2730
   End
   Begin VB.CommandButton cmdRemove 
      Appearance      =   0  'Flat
      Caption         =   "&Remove"
      Height          =   690
      Left            =   90
      Picture         =   "frcControl.frx":238C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4680
      Width           =   1365
   End
   Begin VB.ListBox lstResults 
      BackColor       =   &H8000000A&
      ForeColor       =   &H8000000D&
      Height          =   4350
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label lblStats 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Minimum:888  Maximum:888  Mean:888.888 SD:888.88"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   2970
      TabIndex        =   5
      Top             =   120
      Width           =   7005
   End
   Begin VB.Label lblName 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Label1"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   810
      TabIndex        =   3
      Top             =   6030
      Visible         =   0   'False
      Width           =   1545
   End
End
Attribute VB_Name = "frcControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdCancel_Click()

10        Unload Me

End Sub

Private Sub cmdPrint_Click()

10        On Error GoTo cmdPrint_Click_Error

20        Me.PrintForm

30        Exit Sub

cmdPrint_Click_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frcControl", "cmdPrint_Click", intEL, strES


End Sub

Private Sub cmdRemove_Click()

          Dim s As String
          Dim D As String
          Dim n As String
          Dim Code As String
          Dim sql As String

10        On Error GoTo cmdRemove_Click_Error

20        If lstResults.ListIndex = -1 Then Exit Sub

30        s = lstResults.List(lstResults.ListIndex)
40        D = Format(Left(s, 10), "dd/mmm/yyyy")
50        n = Trim(Mid(s, 15, 6))

60        If frmQCparent.optBio Then
70            Code = CodeForLongName(mscGraph.TitleText)
80            sql = "DELETE FROM BioResults WHERE " & _
                    "SampleID = '" & n & "' " & _
                    "AND Code = '" & Code & "'"
90            Cnxn(0).Execute (sql)
100           drawgrafbio
110       Else
120           Code = CoagCodeFor(mscGraph.TitleText)
130           sql = "DELETE FROM CoagResults WHERE " & _
                    "SampleID = '" & n & "' " & _
                    "AND Code = '" & Code & "'"

140           Cnxn(0).Execute (sql)
150           DrawGrafCoag
160       End If

170       Exit Sub

cmdRemove_Click_Error:

          Dim strES As String
          Dim intEL As Integer

180       intEL = Erl
190       strES = Err.Description
200       LogError "frcControl", "cmdRemove_Click", intEL, strES

End Sub

Private Sub drawgrafbio()

          Dim sql As String
          Dim rsn As Recordset
          Dim sn As New Recordset
          Dim Parameter As String
          Dim s As String
          Dim l As String
          Dim Counter As Long
          Dim v As Single
          Dim Maximum As Single
          Dim minimum As Single
          Dim sd As Single
          Dim mean As Single
          Dim n As Long

10        On Error GoTo drawgrafbio_Error

20        ReDim X(0 To 0) As Variant
30        ReDim All(0 To 1) As Single



40        Parameter = CodeForLongName(mscGraph.TitleText)


50        sql = "SELECT RunDate, SampleID FROM Demographics " & _
                "WHERE Chart = '" & lblName.Caption & "' " & _
                "AND RunDate BETWEEN '" & _
                Format(frmQCparent.calFromDate, "dd/mmm/yyyy") & "' AND '" & _
                Format(frmQCparent.calToDate, "dd/mmm/yyyy") & "' " & _
                "ORDER BY RunDate"
60        Set sn = New Recordset
70        RecOpenClient 0, sn, sql

80        If sn.EOF Then
90            n = iMsg("No data.", 64, "NetAcquire")
100           Unload Me
110           Exit Sub
120       End If
130       sn.MoveLast
140       If sn.RecordCount < 5 Then
150           n = iMsg("Less than 5 Data points." & Chr(10) & "Unable to draw graph", 64, "NetAcquire")
160           sn.Close
170           Unload Me
180           Exit Sub
190       End If

200       s = ""
210       Counter = 0
220       lstResults.Clear
230       Maximum = 0
240       minimum = 99999
250       sn.MoveFirst
260       Do While Not sn.EOF
270           sql = "SELECT * FROM Bioresults WHERE " & _
                    "SampleID = '" & sn!SampleID & "' " & _
                    "AND Code = '" & Parameter & "'"

280           Set rsn = New Recordset
290           RecOpenServer 0, rsn, sql
300           If Not rsn.EOF Then
310               ReDim Preserve All(0 To Counter)
320               ReDim Preserve X(0 To Counter)
330               v = rsn!Result
340               All(Counter) = v
350               X(Counter) = v
360               s = s & v & vbTab
370               l = Format(sn!Rundate, "dd/mm/yyyy") & ":" & "   " & Trim(rsn!SampleID) & " " & rsn!Result
380               lstResults.AddItem l, 0
390               Counter = Counter + 1
400               If v < minimum Then minimum = v
410               If v > Maximum Then Maximum = v
420           End If
430           sn.MoveNext
440       Loop

450       If Counter < 5 Then
460           n = iMsg("Less than 5 Data points." & Chr(10) & "Unable to draw graph", 64, "NetAcquire")
470           Unload Me
480           Exit Sub
490       End If
500       s = Left(s, Len(s) - 1)
510       sd = calcsd(All()) + 0.00001
520       mean = calcmean(All())

530       lblStats = "Minimum:" & minimum & "   Maximum:" & Maximum & "   Mean:" & mean & "   SD:" & Format(sd, "0.0##") & "   CV:" & Format(sd / mean * 100, "#0.00") & "%"

540       ReDim gX(1 To Counter, 0)
550       For n = 1 To Counter
560           gX(n, 0) = X(n - 1)
570       Next

580       mscGraph.ChartData = gX


          'mscGraph.Plot.Axis(VtChAxisIdY).ValueScale.Auto = True
590       mscGraph.Plot.Axis(VtChAxisIdY2).ValueScale.minimum = minimum - 1
600       mscGraph.Plot.Axis(VtChAxisIdY2).ValueScale.Maximum = Maximum + 1    'mean + sd * 3
610       mscGraph.Plot.Axis(VtChAxisIdY).ValueScale.minimum = minimum - 1    'mean - sd * 3
620       mscGraph.Plot.Axis(VtChAxisIdY).ValueScale.Maximum = Maximum + 1    'mean + sd * 3

630       sn.Close




640       Exit Sub

drawgrafbio_Error:

          Dim strES As String
          Dim intEL As Integer

650       intEL = Erl
660       strES = Err.Description
670       LogError "frcControl", "drawgrafbio", intEL, strES


End Sub

Private Sub DrawGrafCoag()

          Dim sql As String
          Dim rsn As Recordset
          Dim sn As New Recordset
          Dim Parameter As String
          Dim s As String
          Dim l As String
          Dim Counter As Long
          Dim v As Single
          Dim Maximum As Single
          Dim minimum As Single
          Dim sd As Single
          Dim mean As Single
          Dim n As Long

10        On Error GoTo drawgrafCoag_Error

20        ReDim X(0 To 0) As Variant
30        ReDim All(0 To 1) As Single

40        Parameter = CoagCodeFor(mscGraph.TitleText)

50        sql = "SELECT RunDate, SampleID FROM Demographics " & _
                "WHERE Chart = '" & lblName.Caption & "' " & _
                "AND RunDate BETWEEN '" & _
                Format(frmQCparent.calFromDate, "dd/mmm/yyyy") & "' AND '" & _
                Format(frmQCparent.calToDate, "dd/mmm/yyyy") & "' " & _
                "ORDER BY RunDate"
60        Set sn = New Recordset
70        RecOpenClient 0, sn, sql

80        If sn.EOF Then
90            n = iMsg("No data.", 64, "NetAcquire")
100           Unload Me
110           Exit Sub
120       End If
130       sn.MoveLast
140       If sn.RecordCount < 5 Then
150           n = iMsg("Less than 5 Data points." & Chr(10) & "Unable to draw graph", 64, "NetAcquire")
160           sn.Close
170           Unload Me
180           Exit Sub
190       End If

200       s = ""
210       Counter = 0
220       lstResults.Clear
230       Maximum = 0
240       minimum = 99999
250       sn.MoveFirst
260       Do While Not sn.EOF
270           sql = "SELECT * FROM CoagResults AS R, CoagTestDefinitions AS T WHERE " & _
                    "R.SampleID = '" & sn!SampleID & "' " & _
                    "AND R.Code = '" & Parameter & "' " & _
                    "AND R.Code = T.Code " & _
                    "AND R.Units = T.Units"

280           Set rsn = New Recordset
290           RecOpenServer 0, rsn, sql
300           If Not rsn.EOF Then
310               ReDim Preserve All(0 To Counter)
320               ReDim Preserve X(0 To Counter)
330               v = rsn!Result
340               All(Counter) = v
350               X(Counter) = v
360               s = s & v & vbTab
370               l = Format(rsn!Rundate, "dd/mm/yyyy") & ":" & "   " & rsn!SampleID & " " & rsn!Result
380               lstResults.AddItem l, 0
390               Counter = Counter + 1
400               If v < minimum Then minimum = v
410               If v > Maximum Then Maximum = v
420           End If
430           sn.MoveNext
440       Loop

450       If Counter < 5 Then
460           n = iMsg("Less than 5 Data points." & Chr(10) & "Unable to draw graph", 64, "NetAcquire")
470           Unload Me
480           Exit Sub
490       End If
500       s = Left(s, Len(s) - 1)
510       sd = calcsd(All()) + 0.00001
520       mean = calcmean(All())

530       lblStats = "Minimum:" & minimum & "   Maximum:" & Maximum & "   Mean:" & mean & "   SD:" & Format(sd, "0.0##") & "   CV:" & Format(sd / mean * 100, "#0.00") & "%"

540       mscGraph.Plot.Axis(VtChAxisIdY2).ValueScale.minimum = mean - sd * 3
550       mscGraph.Plot.Axis(VtChAxisIdY2).ValueScale.Maximum = mean + sd * 3
560       mscGraph.Plot.Axis(VtChAxisIdY).ValueScale.minimum = mean - sd * 3
570       mscGraph.Plot.Axis(VtChAxisIdY).ValueScale.Maximum = mean + sd * 3
580       ReDim gX(1 To Counter, 0)
590       For n = 1 To Counter
600           gX(n, 0) = X(n - 1)
610       Next

620       mscGraph.ChartData = gX

630       sn.Close

640       Exit Sub

drawgrafCoag_Error:

          Dim strES As String
          Dim intEL As Integer

650       intEL = Erl
660       strES = Err.Description
670       LogError "frcControl", "drawgrafCoag", intEL, strES, sql

End Sub

Private Sub Form_Activate()

10        On Error GoTo Form_Activate_Error

20        Me.Height = 6810

30        If frmQCparent.optBio Then
40            drawgrafbio
50        Else
60            DrawGrafCoag
70        End If

80        Exit Sub

Form_Activate_Error:

          Dim strES As String
          Dim intEL As Integer

90        intEL = Erl
100       strES = Err.Description
110       LogError "frcControl", "Form_Activate", intEL, strES


End Sub

Private Sub Form_Load()


10        On Error GoTo Form_Load_Error

20        lblName.Caption = frmQCparent.tName

30        mscGraph.TitleText = frmQCparent.lstpara

40        Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frcControl", "Form_Load", intEL, strES


End Sub

Private Sub mscGraph_PointSelected(Series As Integer, DataPoint As Integer, MouseFlags As Integer, Cancel As Integer)
          Dim dps As Long

10        On Error GoTo mscGraph_PointSelected_Error

20        dps = mscGraph.Plot.SeriesCollection(1).DataPoints.Count

30        lstResults.Selected(dps - DataPoint) = True

40        Exit Sub

mscGraph_PointSelected_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frcControl", "mscGraph_PointSelected", intEL, strES

End Sub
