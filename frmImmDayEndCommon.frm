VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.ocx"
Begin VB.Form frmEndDayImmCommon 
   Caption         =   "NetAcquire 6 - Immunology End of Day Report (Common Parameters)"
   ClientHeight    =   8430
   ClientLeft      =   90
   ClientTop       =   420
   ClientWidth     =   11520
   Icon            =   "frmImmDayEndCommon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   11520
   Begin VB.CommandButton bBlockValidate 
      Caption         =   "Block Validate"
      Height          =   960
      Left            =   5655
      Picture         =   "frmImmDayEndCommon.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   45
      Width           =   1245
   End
   Begin VB.CommandButton bBlockPrint 
      Caption         =   "Block Print"
      Height          =   960
      Left            =   6960
      Picture         =   "frmImmDayEndCommon.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   45
      Width           =   1245
   End
   Begin VB.CommandButton breport 
      Appearance      =   0  'Flat
      Caption         =   "Print &Report"
      Height          =   960
      Left            =   4350
      Picture         =   "frmImmDayEndCommon.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   45
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.CommandButton bprint 
      Appearance      =   0  'Flat
      Caption         =   "&Print Listing"
      Height          =   960
      Left            =   8265
      Picture         =   "frmImmDayEndCommon.frx":0C28
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   45
      Width           =   1245
   End
   Begin VB.CommandButton bcancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   960
      Left            =   9570
      Picture         =   "frmImmDayEndCommon.frx":0F32
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   45
      Width           =   1245
   End
   Begin VB.CommandButton cmdStart 
      Appearance      =   0  'Flat
      Caption         =   "Start Report"
      Height          =   960
      Left            =   3060
      Picture         =   "frmImmDayEndCommon.frx":123C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   45
      Visible         =   0   'False
      Width           =   1245
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   6945
      Left            =   180
      TabIndex        =   0
      Top             =   1350
      Width           =   11205
      _ExtentX        =   19764
      _ExtentY        =   12250
      _Version        =   393216
      Cols            =   25
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
      FormatString    =   $"frmImmDayEndCommon.frx":1546
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
      TabIndex        =   1
      Top             =   1125
      Width           =   11205
      _ExtentX        =   19764
      _ExtentY        =   291
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin MSComCtl2.DTPicker dtRunDate 
      Height          =   405
      Left            =   720
      TabIndex        =   8
      Top             =   360
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MMM/yyyy"
      Format          =   59047939
      CurrentDate     =   36966
   End
End
Attribute VB_Name = "frmEndDayImmCommon"
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
190       LogError "frmEndDayImmCommon", "bBlockPrint_Click", intEL, strES


End Sub

Private Sub bBlockValidate_Click()

          Dim sql As String
          Dim Y As Long
          Dim gStart As Long
          Dim gStop As Long


10        On Error GoTo bBlockValidate_Click_Error

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
130           sql = "UPDATE Immresults " & _
                    "set Valid = 1 , Operator = '" & AddTicks(UserCode) & "' " & _
                    "WHERE sampleid = '" & g.TextMatrix(Y, 0) & "'"
140           Cnxn(0).Execute (sql)
150       Next




160       Exit Sub

bBlockValidate_Click_Error:

          Dim strES As String
          Dim intEL As Integer



170       intEL = Erl
180       strES = Err.Description
190       LogError "frmEndDayImmCommon", "bBlockValidate_Click", intEL, strES, sql


End Sub

Private Sub bcancel_Click()

10        Unload Me

End Sub

Private Sub bprint_Click()

          Dim Y As Long
          Dim X As Long
          Dim sql As String
          Dim sn As New Recordset


10        On Error GoTo bprint_Click_Error

20        pb.Visible = True
30        pb.Max = g.Rows - 1

40        Printer.Orientation = vbPRORLandscape
50        Printer.Font.Name = "Courier New"
60        Printer.Font.Size = 8
70        Printer.Print "Immunology end of day report for " & Format(dtRunDate, "dd/mmm/yyyy")

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
260               sql = "SELECT * from gps WHERE text =  '" & Trim(UCase(g.TextMatrix(Y, 4))) & "'"
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
380       Next

390       Printer.Font.Size = 10
400       Printer.EndDoc

410       pb.Visible = False







420       Exit Sub

bprint_Click_Error:

          Dim strES As String
          Dim intEL As Integer



430       intEL = Erl
440       strES = Err.Description
450       LogError "frmEndDayImmCommon", "bPrint_Click", intEL, strES, sql


End Sub

Private Sub breport_Click()

10        On Error GoTo breport_Click_Error

20        g.Col = 0

30        Exit Sub

breport_Click_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmEndDayImmCommon", "breport_Click", intEL, strES


End Sub

Private Sub cmdStart_Click()

10        On Error GoTo cmdStart_Click_Error

20        FillG

30        Exit Sub

cmdStart_Click_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmEndDayImmCommon", "cmdStart_Click", intEL, strES


End Sub

Private Sub dtRunDate_CloseUp()

10        On Error GoTo dtRunDate_CloseUp_Error

20        FillG

30        Exit Sub

dtRunDate_CloseUp_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmEndDayImmCommon", "dtRunDate_CloseUp", intEL, strES


End Sub

Private Sub FillG()

          Dim tb As New Recordset
          Dim sql As String
          Dim s As String
          Dim Y As Long
          Dim Counter As Long
          Dim vaRunNumber As String
          Dim tsql As String
          Dim Test(30) As String
          Dim X As Long
          Dim n As Long
          Dim strin As String


10        On Error GoTo FillG_Error

20        ClearFGrid g


30        Counter = 0
40        vaRunNumber = ""
50        n = 0

60        sql = "SELECT distinct shortname, code from immtestdefinitions WHERE eod = '1'"
70        Set tb = New Recordset
80        RecOpenServer 0, tb, sql
90        Do While Not tb.EOF
100           tsql = tsql & "immresults.code = '" & Trim(tb!Code) & "' or "
110           Test(n) = Trim(tb!ShortName)
120           n = n + 1
130           tb.MoveNext
140       Loop

150       n = n - 1

160       For X = 0 To n
170           g.TextMatrix(0, X + 5) = Test(X)
180       Next

190       For X = 5 To g.Cols - 1
200           g.ColWidth(X) = 0
210       Next

220       For X = 5 To 5 + n
230           g.ColWidth(X) = 800
240       Next

250       If Len(Trim(tsql)) = 0 Then
260           iMsg "No Tests Chosen!"
270           Exit Sub
280       End If

290       tsql = Mid(tsql, 1, Len(tsql) - 3)

300       sql = "SELECT * from demographics, immresults  WHERE " & _
                "demographics.sampleid = immresults.sampleid " & _
                "and ( immresults.rundate =  '" & Format(dtRunDate, "dd/mmm/yyyy") & "') and (" & tsql & _
                ") order by immresults.sampleid"
310       Set tb = New Recordset
320       RecOpenClient 0, tb, sql

330       pb.Max = tb.RecordCount + 1
340       pb = 0
350       pb.Visible = True
360       g.Visible = False

370       With tb
380           Do While Not .EOF
390               pb = pb + 1
400               If vaRunNumber <> Trim(tb!SampleID) Then
410                   s = tb!SampleID & vbTab & tb!PatName & vbTab & _
                          tb!Dob & vbTab & _
                          tb!Ward & vbTab & _
                          tb!Clinician & ""
420                   vaRunNumber = Trim(tb!SampleID)
430                   g.AddItem s
440               End If
450               Y = g.Rows - 1

460               strin = ImmShortNameFor(!Code & "")

470               For X = 0 To n
480                   If strin = Test(X) Then
490                       g.TextMatrix(Y, X + 5) = !Result
500                   End If
510               Next
520               .MoveNext
530           Loop
540       End With

550       FixG g

560       pb.Visible = False





570       Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer



580       intEL = Erl
590       strES = Err.Description
600       LogError "frmEndDayImmCommon", "FillG", intEL, strES, sql


End Sub

Private Sub Form_Load()

10        On Error GoTo Form_Load_Error

20        dtRunDate = Format(Now, "dd/mmm/yyyy")

30        FillG

40        Set_Font Me

50        Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer



60        intEL = Erl
70        strES = Err.Description
80        LogError "frmEndDayImmCommon", "Form_Load", intEL, strES


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
240       LogError "frmEndDayImmCommon", "g_Click", intEL, strES


End Sub
