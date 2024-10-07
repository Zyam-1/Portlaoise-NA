VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.ocx"
Begin VB.Form frmDayEndCommon 
   Caption         =   "NetAcquire 6 - Biochemistry End of Day Report (Common Parameters)"
   ClientHeight    =   8430
   ClientLeft      =   90
   ClientTop       =   420
   ClientWidth     =   13110
   Icon            =   "frmDayEndCommon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   13110
   Begin VB.CommandButton cmdExcel 
      Caption         =   "Export to Excel"
      Height          =   960
      Left            =   7470
      Picture         =   "frmDayEndCommon.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   60
      Width           =   1230
   End
   Begin VB.CommandButton cmdStart 
      Appearance      =   0  'Flat
      Caption         =   "Start Report"
      Height          =   960
      Left            =   2070
      Picture         =   "frmDayEndCommon.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   60
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.CommandButton bcancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   960
      Left            =   11745
      Picture         =   "frmDayEndCommon.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   45
      Width           =   1245
   End
   Begin VB.CommandButton bprint 
      Appearance      =   0  'Flat
      Caption         =   "&Print Listing"
      Height          =   960
      Left            =   10395
      Picture         =   "frmDayEndCommon.frx":0C28
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   45
      Width           =   1245
   End
   Begin VB.CommandButton breport 
      Appearance      =   0  'Flat
      Caption         =   "Print &Report"
      Height          =   960
      Left            =   10395
      Picture         =   "frmDayEndCommon.frx":0F32
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   45
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.CommandButton bBlockPrint 
      Caption         =   "Block Print"
      Height          =   960
      Left            =   10395
      Picture         =   "frmDayEndCommon.frx":123C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   45
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.CommandButton bBlockValidate 
      Caption         =   "Block Validate"
      Height          =   960
      Left            =   10440
      Picture         =   "frmDayEndCommon.frx":1546
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   45
      Visible         =   0   'False
      Width           =   1155
   End
   Begin MSComCtl2.DTPicker dtRunDate 
      Height          =   315
      Left            =   330
      TabIndex        =   0
      Top             =   180
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd/MMM/yyyy"
      Format          =   109117443
      CurrentDate     =   36966
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   7035
      Left            =   180
      TabIndex        =   3
      Top             =   1260
      Width           =   12825
      _ExtentX        =   22622
      _ExtentY        =   12409
      _Version        =   393216
      Cols            =   200
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
      FormatString    =   $"frmDayEndCommon.frx":1850
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComctlLib.ProgressBar pb 
      Height          =   165
      Left            =   180
      TabIndex        =   7
      Top             =   1080
      Width           =   12780
      _ExtentX        =   22543
      _ExtentY        =   291
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin VB.Label lblExcelInfo 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exporting..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   8850
      TabIndex        =   10
      Top             =   360
      Visible         =   0   'False
      Width           =   1125
   End
End
Attribute VB_Name = "frmDayEndCommon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub bBlockPrint_Click()

          Dim Y As Long
          Dim gStart As Long
          Dim gStop As Long

10        On Error GoTo bBlockPrint_Click_Error

20        If g.Row > g.RowSel Then
30            gStart = g.RowSel
40            gStop = g.Row
50        ElseIf g.Row < g.RowSel Then
60            gStart = g.Row
70            gStop = g.RowSel
80        Else
90            gStart = g.Row
100           gStop = g.Row
110       End If

120       For Y = gStart To gStop
130           g.Row = Y
140           g.Col = 0

150       Next

160       Exit Sub

bBlockPrint_Click_Error:

          Dim strES As String
          Dim intEL As Integer

170       intEL = Erl
180       strES = Err.Description
190       LogError "frmDayEndCommon", "bBlockPrint_Click", intEL, strES

End Sub

Private Sub bBlockValidate_Click()

          Dim sql As String
          Dim Y As Long
          Dim gStart As Long
          Dim gStop As Long

10        On Error GoTo bBlockValidate_Click_Error

20        If g.Rows = 2 And g.TextMatrix(1, 0) = "" Then Exit Sub

30        If g.Row > g.RowSel Then
40            gStart = g.RowSel
50            gStop = g.Row
60        ElseIf g.Row < g.RowSel Then
70            gStart = g.Row
80            gStop = g.RowSel
90        Else
100           gStart = g.Row
110           gStop = g.Row
120       End If

130       For Y = gStart To gStop
140           sql = "UPDATE Bioresults " & _
                    "set Valid = 1, Operator = '" & AddTicks(UserCode) & "' " & _
                    "WHERE RunNumber = '" & g.TextMatrix(Y, 0) & "'"
150           Cnxn(0).Execute (sql)
160       Next

170       Exit Sub

bBlockValidate_Click_Error:

          Dim strES As String
          Dim intEL As Integer

180       intEL = Erl
190       strES = Err.Description
200       LogError "frmDayEndCommon", "bBlockValidate_Click", intEL, strES

End Sub

Private Sub bcancel_Click()

10        Unload Me

End Sub

Private Sub bprint_Click()

          Dim Y As Long
          Dim X As Long
          Dim sn As New Recordset
          Dim sql As String

10        On Error GoTo bprint_Click_Error

20        pb.Visible = True
30        pb.Max = g.Rows - 1

40        Printer.Orientation = vbPRORLandscape
50        Printer.Font.Name = "Courier New"
60        Printer.Font.Size = 8
70        Printer.Print "End of day report for " & Format(dtRunDate, "dd/mmm/yyyy")

80        For Y = 0 To g.Rows - 1
90            pb = Y
100           Printer.Print g.TextMatrix(Y, 0);    'RunNumber
110           Printer.Print Tab(10); g.TextMatrix(Y, 1);    'Name
120           Printer.Print Tab(30); g.TextMatrix(Y, 2);    'DoB
130           If Y = 0 Then
140               Printer.Print Tab(41); "Ward";
150           Else
160               sql = "SELECT * from wards WHERE text = '" & Trim(UCase(g.TextMatrix(Y, 3))) & "' or code = '" & Trim(UCase(g.TextMatrix(Y, 3))) & "'"
170               Set sn = New Recordset
180               RecOpenServer 0, sn, sql
190               If Not sn.EOF Then
200                   Printer.Print Tab(41); Left(sn!Code & "      ", 6);
210               End If
220           End If
230           If Y = 0 Then
240               Printer.Print Tab(48); "Clin";
250           Else
260               sql = "SELECT * from gps WHERE text =  '" & AddTicks(Trim(UCase(g.TextMatrix(Y, 4)))) & "'"
270               Set sn = New Recordset
280               RecOpenServer 0, sn, sql
290               If Not sn.EOF Then
300                   Printer.Print Tab(48); Left(sn!Code & "      ", 6);
310               End If
320           End If
330           Printer.Print Tab(54);
340           For X = 5 To g.Cols - 1
350               Printer.Print Left(g.TextMatrix(Y, X) & "     ", 5);
360           Next
370           Printer.Print
380           If Printer.CurrentY = 5000 Then
390               Printer.NewPage
400               Printer.Print g.TextMatrix(0, 0);    'RunNumber
410               Printer.Print Tab(10); g.TextMatrix(0, 1);    'Name
420               Printer.Print Tab(30); g.TextMatrix(0, 2);    'DoB
430               Printer.Print Tab(41); "Ward";
440               Printer.Print Tab(48); "Clin";
450               For X = 5 To g.Cols - 1
460                   Printer.Print Left(g.TextMatrix(0, X) & "     ", 5);
470               Next

480           End If
490       Next

500       Printer.Font.Size = 10
510       Printer.EndDoc

520       pb.Visible = False

530       Exit Sub

bprint_Click_Error:

          Dim strES As String
          Dim intEL As Integer

540       intEL = Erl
550       strES = Err.Description
560       LogError "frmDayEndCommon", "bPrint_Click", intEL, strES, sql

End Sub

Private Sub breport_Click()

10        g.Col = 0

End Sub

Private Sub cmdExcel_Click()

10        ExportFlexGrid g, Me

End Sub

Private Sub cmdStart_Click()

10        FillG

End Sub

Private Sub dtRunDate_Change()

10        cmdStart.Visible = True

End Sub

Private Sub dtRunDate_CloseUp()

10        FillG

End Sub

Private Sub FillG()

          Dim tb As New Recordset
          Dim sql As String
          Dim s As String
          Dim Y As Long
          Dim Counter As Long
          Dim vaRunNumber As String
          Dim tsql As String
          Dim Test(200) As String
          Dim X As Long
          Dim n As Long
          Dim strin As String
          Dim InList As Boolean

10        On Error GoTo FillG_Error

20        ClearFGrid g

30        For n = 6 To (g.Cols - 1)
40            g.ColWidth(n) = 0
50        Next

60        Counter = 0
70        vaRunNumber = ""
80        n = 0

90        sql = "SELECT distinct(code), shortname from biotestdefinitions WHERE eod = 1"
100       Set tb = New Recordset
110       RecOpenServer 0, tb, sql
120       Do While Not tb.EOF
130           InList = False

140           For Counter = 0 To n
150               If Test(Counter) = tb!ShortName Then
160                   InList = True
170                   Exit For
180               End If
190           Next

200           If InList = False Then
210               tsql = tsql & "R.Code = '" & tb!Code & "' or "
220               Test(n) = tb!ShortName
230               n = n + 1
240               g.ColWidth(n + 6) = 500
250           End If
260           tb.MoveNext
270       Loop

280       n = n - 1

290       For X = 0 To n
300           g.TextMatrix(0, X + 6) = Test(X)
310       Next

320       If Len(Trim(tsql)) = 0 Then
330           iMsg "No Tests Chosen!"
340           Exit Sub
350       End If

360       tsql = Mid(tsql, 1, Len(tsql) - 3)

370       sql = "SELECT  D.*, R.* FROM Demographics D, BioResults R WHERE " & _
                "D.SampleID = R.SampleID " & _
                "AND (R.RunTime BETWEEN  '" & Format(dtRunDate, "dd/mmm/yyyy") & " 00:00:00" & "' and '" & Format(dtRunDate, "dd/mmm/yyyy") & " 23:59:59" & "') " & _
                "AND (" & tsql & ") " & _
                "ORDER BY R.SampleID"
380       Set tb = New Recordset
390       RecOpenClient 0, tb, sql
400       If Not tb.EOF Then
410           pb.Max = 300 + 1
420           pb = 0
430           pb.Visible = True
440           g.Visible = False

450           With tb
460               Do While Not .EOF
                      '  PB = PB + 1
470                   If vaRunNumber <> Trim(tb!SampleID) Then
480                       s = tb!SampleID & vbTab & tb!PatName & vbTab & _
                              tb!Dob & vbTab & _
                              tb!Ward & vbTab & _
                              tb!Clinician & "" & vbTab & tb!GP & ""
490                       vaRunNumber = Trim(tb!SampleID)
500                       g.AddItem s
510                   End If
520                   Y = g.Rows - 1

530                   strin = ShortNameFor(!Code & "")

540                   For X = 0 To n
550                       If strin = Test(X) Then
560                           g.TextMatrix(Y, X + 6) = !Result
570                       End If
580                   Next
590                   .MoveNext
600               Loop
610           End With
620       End If

630       FixG g

640       cmdStart.Visible = False

650       pb.Visible = False
660       g.Visible = True

670       Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer

680       intEL = Erl
690       strES = Err.Description
700       LogError "frmDayEndCommon", "FillG", intEL, strES, sql

End Sub

Private Sub Form_Load()

10        On Error GoTo Form_Load_Error

20        dtRunDate = Format(Now, "dd/mmm/yyyy")

30        FillG

40        Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmDayEndCommon", "Form_Load", intEL, strES

End Sub

Private Sub g_Click()

          Static SortOrder As Boolean

10        On Error GoTo g_Click_Error

20        If g.MouseRow = 0 Then
30            If SortOrder Then
40                g.Sort = flexSortStringAscending
50            Else
60                g.Sort = flexSortStringDescending
70            End If
80            SortOrder = Not SortOrder
90            breport.Visible = False
100           Exit Sub
110       End If

120       If g.Row = g.RowSel Then
130           breport.Visible = True
140           bBlockValidate.Visible = False
150           bBlockPrint.Visible = False
160       Else
170           breport.Visible = False
180           bBlockValidate.Visible = True
190           bBlockPrint.Visible = True
200       End If

210       Exit Sub

g_Click_Error:

          Dim strES As String
          Dim intEL As Integer

220       intEL = Erl
230       strES = Err.Description
240       LogError "frmDayEndCommon", "g_Click", intEL, strES

End Sub
