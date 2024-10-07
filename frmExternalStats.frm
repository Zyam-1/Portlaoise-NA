VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmExternalStats 
   Caption         =   "NetAcquire - External Tests"
   ClientHeight    =   6285
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   10935
   Icon            =   "frmExternalStats.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6285
   ScaleWidth      =   10935
   Begin VB.CommandButton cmdXL 
      Caption         =   "Export to Excel"
      Height          =   825
      Left            =   9600
      Picture         =   "frmExternalStats.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   180
      Width           =   825
   End
   Begin MSFlexGridLib.MSFlexGrid grdAnalyte 
      Height          =   4545
      Left            =   7860
      TabIndex        =   20
      Top             =   1560
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   8017
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   2
      FormatString    =   "<Analyte                           |^Total     "
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   165
      Left            =   4080
      TabIndex        =   19
      Top             =   1380
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   810
      Left            =   6510
      Picture         =   "frmExternalStats.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   195
      Width           =   1245
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   810
      Left            =   7860
      Picture         =   "frmExternalStats.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   195
      Width           =   1245
   End
   Begin VB.ListBox List1 
      Height          =   3255
      IntegralHeight  =   0   'False
      Left            =   300
      TabIndex        =   15
      Top             =   2820
      Width           =   3615
   End
   Begin VB.PictureBox SSPanel1 
      Height          =   2565
      Index           =   1
      Left            =   300
      ScaleHeight     =   2505
      ScaleWidth      =   3555
      TabIndex        =   4
      Top             =   180
      Width           =   3615
      Begin VB.OptionButton oBetween 
         Alignment       =   1  'Right Justify
         Caption         =   "Last Week"
         Height          =   225
         Index           =   0
         Left            =   600
         TabIndex        =   12
         Top             =   780
         Width           =   1095
      End
      Begin VB.OptionButton oBetween 
         Alignment       =   1  'Right Justify
         Caption         =   "Last Month"
         Height          =   195
         Index           =   1
         Left            =   540
         TabIndex        =   11
         Top             =   1050
         Width           =   1155
      End
      Begin VB.OptionButton oBetween 
         Alignment       =   1  'Right Justify
         Caption         =   "Last Full Month"
         Height          =   195
         Index           =   2
         Left            =   300
         TabIndex        =   10
         Top             =   1320
         Width           =   1395
      End
      Begin VB.OptionButton oBetween 
         Alignment       =   1  'Right Justify
         Caption         =   "Last Quarter"
         Height          =   195
         Index           =   3
         Left            =   510
         TabIndex        =   9
         Top             =   1590
         Width           =   1185
      End
      Begin VB.OptionButton oBetween 
         Alignment       =   1  'Right Justify
         Caption         =   "Last Full Quarter"
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   8
         Top             =   1860
         Width           =   1515
      End
      Begin VB.OptionButton oBetween 
         Alignment       =   1  'Right Justify
         Caption         =   "Year To Date"
         Height          =   195
         Index           =   5
         Left            =   390
         TabIndex        =   7
         Top             =   2130
         Width           =   1305
      End
      Begin VB.OptionButton oBetween 
         Alignment       =   1  'Right Justify
         Caption         =   "Today"
         Height          =   195
         Index           =   6
         Left            =   900
         TabIndex        =   6
         Top             =   540
         Value           =   -1  'True
         Width           =   795
      End
      Begin VB.CommandButton breCalc 
         Caption         =   "Start"
         Height          =   945
         Left            =   1980
         Picture         =   "frmExternalStats.frx":0C28
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   930
         Width           =   1185
      End
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   315
         Left            =   1830
         TabIndex        =   13
         Top             =   150
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   65601537
         CurrentDate     =   38126
      End
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   315
         Left            =   240
         TabIndex        =   14
         Top             =   150
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   65601537
         CurrentDate     =   38126
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Source"
      Height          =   915
      Left            =   4140
      TabIndex        =   0
      Top             =   90
      Width           =   1185
      Begin VB.OptionButton oSource 
         Caption         =   "G.P.s"
         Height          =   195
         Index           =   2
         Left            =   150
         TabIndex        =   3
         Top             =   630
         Width           =   825
      End
      Begin VB.OptionButton oSource 
         Caption         =   "Wards"
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   2
         Top             =   420
         Width           =   855
      End
      Begin VB.OptionButton oSource 
         Caption         =   "Clinicians"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   1
         Top             =   210
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   4545
      Left            =   4080
      TabIndex        =   18
      Top             =   1560
      Width           =   3705
      _ExtentX        =   6535
      _ExtentY        =   8017
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      GridLines       =   2
      FormatString    =   "<Source               |<Samples |<Tests      |<T/S      "
   End
   Begin VB.Label lblExcelInfo 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exporting..."
      Height          =   345
      Left            =   9360
      TabIndex        =   22
      Top             =   1020
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   4410
      Picture         =   "frmExternalStats.frx":0F32
      Stretch         =   -1  'True
      Top             =   1020
      Width           =   660
   End
End
Attribute VB_Name = "frmExternalStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub FillGrid()

          Dim tb As Recordset
          Dim sql As String
          Dim Total As Long
          Dim Samples As Long
          Dim tests As Long
          Dim n As Integer
          Dim tps As String
          Dim Y As Long
          Dim StartSID As Long
          Dim StopSID As Double


10        On Error GoTo FillGrid_Error

20        With g
30            .Rows = 2
40            .AddItem ""
50            .RemoveItem 1
60        End With

          '      "SampleID > 800000 and SampleID < 900000 " & _

70               If HospName(0) = "Bantry" Then
80        sql = "Select top 1 SampleID from Demographics where " & _
                "SampleID > 800000 and RunDate between '" & Format$(dtFrom, "dd/mmm/yyyy") & " 00:00:00' and '" & _
                Format$(dtTo, "dd/mmm/yyyy") & " 23:59:59' " & _
                "order by SampleID asc"
90    Else
100       sql = "Select top 1 SampleID from Demographics where " & _
                "RunDate between '" & Format$(dtFrom, "dd/mmm/yyyy") & " 00:00:00' and '" & _
                Format$(dtTo, "dd/mmm/yyyy") & " 23:59:59' " & _
                "order by SampleID asc"
110   End If
120   Set tb = New Recordset
130   RecOpenServer 0, tb, sql
140   If Not tb.EOF Then
150       StartSID = tb!SampleID
160   End If

      '      "SampleID > 800000 and SampleID < 900000 "
170   If HospName(0) = "Bantry" Then
180       sql = "Select top 1 SampleID from Demographics where " & _
                "SampleID > 800000 and RunDate between '" & Format$(dtFrom, "dd/mmm/yyyy") & " 00:00:00' and '" & _
                Format$(dtTo, "dd/mmm/yyyy") & " 23:59:59' " & _
                "order by SampleID desc"
190   Else
200       sql = "Select top 1 SampleID from Demographics where " & _
                "RunDate between '" & Format$(dtFrom, "dd/mmm/yyyy") & " 00:00:00' and '" & _
                Format$(dtTo, "dd/mmm/yyyy") & " 23:59:59' " & _
                "order by SampleID desc"
210   End If
220   Set tb = New Recordset
230   RecOpenServer 0, tb, sql
240   If Not tb.EOF Then
250       StopSID = tb!SampleID
260   End If

270   If StartSID = 0 Or StopSID = 0 Then Exit Sub

280   pb.Visible = True

290   For n = 0 To List1.ListCount - 1
300       ReDim SampleIDs(0 To 0) As Double
310       List1.Selected(n) = True
320       sql = "Select distinct D.SampleID from Demographics as D, ExtResults as E where " & _
                "(D.SampleID between " & StartSID & " and " & StopSID & " ) and D.SampleID = E.SampleID and "

330       If oSource(0) Then
340           sql = sql & "clinician = '"
350       ElseIf oSource(1) Then
360           sql = sql & "ward = '"
370       Else
380           sql = sql & "gp = '"
390       End If
400       sql = sql & AddTicks(List1.List(n)) & "'"
410       Set tb = New Recordset
420       RecOpenClient 0, tb, sql
430       If Not tb.EOF Then
440           Do While Not tb.EOF
450               ReDim Preserve SampleIDs(0 To UBound(SampleIDs) + 1)
460               SampleIDs(UBound(SampleIDs)) = tb!SampleID
470               tb.MoveNext
480           Loop

490           pb = 0
500           pb.Max = UBound(SampleIDs)

510           Samples = UBound(SampleIDs)
520           tests = 0
530           For Y = 1 To UBound(SampleIDs)
540               pb = Y
550               sql = "Select count(SampleID) as Tests from ExtResults where " & _
                        "SampleID = '" & SampleIDs(Y) & "'"
560               Set tb = New Recordset
570               RecOpenServer 0, tb, sql
580               tests = tests + tb!tests
590           Next
600           If tests <> 0 And Samples <> 0 Then
610               tps = Format$(tests / Samples, "##.00")
620               g.AddItem List1.List(n) & vbTab & Samples & vbTab & tests & vbTab & tps
630               g.Refresh
640           End If
650       End If
660   Next
670   pb.Visible = False
680   g.AddItem ""

690   If g.Rows = 2 Then Exit Sub

700   g.Col = 1
710   Total = 0
720   For n = 2 To g.Rows - 1
730       g.Row = n
740       Total = Total + Val(g)
750   Next

760   g.Col = 2
770   tests = 0
780   For n = 2 To g.Rows - 1
790       g.Row = n
800       tests = tests + Val(g)
810   Next

820   If Total <> 0 And tests <> 0 Then
830       g.AddItem "Total" & vbTab & Total & vbTab & tests & vbTab & Format$(tests / Total, ".00")
840   Else
850       g.AddItem "Total"
860   End If
870   g.Refresh

880   g.RemoveItem 1


890   Exit Sub

FillGrid_Error:

      Dim strES As String
      Dim intEL As Integer

900   intEL = Erl
910   strES = Err.Description
920   LogError "frmExternalStats", "FillGrid", intEL, strES, sql


End Sub
Private Sub FillGridAnalyte()

          Dim tb As Recordset
          Dim tbRes As Recordset
          Dim sql As String
          Dim tests As Long
          Dim n As Long
          Dim Y As Long
          Dim StartSID As Double
          Dim StopSID As Double
          Dim Found As Boolean


10        On Error GoTo FillGridAnalyte_Error

20        With grdAnalyte
30            .Rows = 2
40            .AddItem ""
50            .RemoveItem 1
60        End With

          '      "SampleID > 800000 and SampleID < 900000 "
70        If HospName(0) = "Bantry" Then
80            sql = "Select top 1 SampleID from Demographics where " & _
                    "SampleID > 800000 and RunDate between '" & Format$(dtFrom, "dd/mmm/yyyy") & " 00:00:00' and '" & _
                    Format$(dtTo, "dd/mmm/yyyy") & " 23:59:59' " & _
                    "order by SampleID asc"
90        Else
100           sql = "Select top 1 SampleID from Demographics where " & _
                    "RunDate between '" & Format$(dtFrom, "dd/mmm/yyyy") & " 00:00:00' and '" & _
                    Format$(dtTo, "dd/mmm/yyyy") & " 23:59:59' " & _
                    "order by SampleID asc"
110       End If
120       Set tb = New Recordset
130       RecOpenServer 0, tb, sql
140       If Not tb.EOF Then
150           StartSID = tb!SampleID
160       End If

          '      "SampleID > 800000 and SampleID < 900000 "
170       If HospName(0) = "Bantry" Then
180           sql = "Select top 1 SampleID from Demographics where " & _
                    "SampleID > 800000 and RunDate between '" & Format$(dtFrom, "dd/mmm/yyyy") & " 00:00:00' and '" & _
                    Format$(dtTo, "dd/mmm/yyyy") & " 23:59:59' " & _
                    "order by SampleID desc"
190       Else
200           sql = "Select top 1 SampleID from Demographics where " & _
                    "RunDate between '" & Format$(dtFrom, "dd/mmm/yyyy") & " 00:00:00' and '" & _
                    Format$(dtTo, "dd/mmm/yyyy") & " 23:59:59' " & _
                    "order by SampleID desc"
210       End If
220       Set tb = New Recordset
230       RecOpenServer 0, tb, sql
240       If Not tb.EOF Then
250           StopSID = tb!SampleID
260       End If

270       If StartSID = 0 Or StopSID = 0 Then Exit Sub

280       pb.Visible = True

290       ReDim SampleIDs(0 To 0) As Double
300       sql = "Select SampleID from demographics where " & _
                "SampleID between " & StartSID & " and " & StopSID
310       Set tb = New Recordset
320       RecOpenClient 0, tb, sql
330       Do While Not tb.EOF
340           ReDim Preserve SampleIDs(0 To UBound(SampleIDs) + 1)
350           SampleIDs(UBound(SampleIDs)) = tb!SampleID
360           tb.MoveNext
370       Loop

380       pb = 0
390       pb.Max = UBound(SampleIDs)

400       For Y = 1 To UBound(SampleIDs)
410           pb = Y
420           sql = "Select * from ExtResults where " & _
                    "SampleID = '" & SampleIDs(Y) & "'"
430           Set tbRes = New Recordset
440           RecOpenServer 0, tbRes, sql
450           Do While Not tbRes.EOF
460               Found = False
470               For n = 1 To grdAnalyte.Rows - 1
480                   If Trim$(tbRes!Analyte & "") = grdAnalyte.TextMatrix(n, 0) Then
490                       grdAnalyte.TextMatrix(n, 1) = Format$(Val(grdAnalyte.TextMatrix(n, 1)) + 1)
500                       Found = True
510                       Exit For
520                   End If
530               Next
540               If Not Found Then
550                   grdAnalyte.AddItem Trim$(tbRes!Analyte & "") & vbTab & "1"
560               End If
570               tbRes.MoveNext
580           Loop
590       Next

600       pb.Visible = False

610       tests = 0
620       With grdAnalyte
630           For n = 1 To .Rows - 1
640               tests = tests + Val(.TextMatrix(n, 1))
650           Next
660           .AddItem "Total" & vbTab & tests

670           If .Rows > 2 Then
680               .RemoveItem 1
690           End If
700       End With



710       Exit Sub

FillGridAnalyte_Error:

          Dim strES As String
          Dim intEL As Integer

720       intEL = Erl
730       strES = Err.Description
740       LogError "frmExternalStats", "FillGridAnalyte", intEL, strES, sql


End Sub

Private Sub FillList()

          Dim sql As String
          Dim tb As Recordset
          Dim strSource As String
          Dim Y As Integer
          Dim Found As Boolean


10        On Error GoTo FillList_Error

20        List1.Clear

30        strSource = Switch(oSource(0), "Clinician", _
                             oSource(1), "Ward", _
                             oSource(2), "GP")

40        sql = "Select distinct " & strSource & " as Source " & _
                "from Demographics where " & _
                "RunDate between '" & Format(dtFrom, "dd/mmm/yyyy") & "' " & _
                "and '" & Format(dtTo, "dd/mmm/yyyy") & "' " & _
                "Order by " & strSource
50        Set tb = New Recordset
60        RecOpenServer 0, tb, sql
70        Do While Not tb.EOF
80            Debug.Print tb!Source & ""
90            Found = False
100           For Y = 0 To List1.ListCount - 1
110               If UCase$(Trim$(tb!Source & "")) = UCase$(List1.List(Y)) Then
120                   Found = True
130                   Exit For
140               End If
150           Next
160           If Not Found Then
170               List1.AddItem Trim$(tb!Source & "")
180           End If
190           tb.MoveNext
200       Loop

210       Exit Sub

FillList_Error:

          Dim strES As String
          Dim intEL As Integer



220       intEL = Erl
230       strES = Err.Description
240       LogError "frmExternalStats", "FillList", intEL, strES, sql


End Sub


Private Sub brecalc_Click()

10        On Error GoTo brecalc_Click_Error

20        FillList
30        FillGrid
40        FillGridAnalyte

50        Exit Sub

brecalc_Click_Error:

          Dim strES As String
          Dim intEL As Integer



60        intEL = Erl
70        strES = Err.Description
80        LogError "frmExternalStats", "brecalc_Click", intEL, strES


End Sub


Private Sub cmdCancel_Click()

10        Unload Me

End Sub


Private Sub cmdPrint_Click()

          Dim n As Integer
          Dim X As Integer

10        On Error GoTo cmdPrint_Click_Error

20        Printer.Print "Totals:"; dtFrom; " to "; dtTo

30        Printer.Print
40        For n = 0 To g.Rows - 1
50            g.Row = n
60            For X = 0 To 3
70                g.Col = X
80                Printer.Print Tab(Choose(X + 1, 1, 40, 50, 60)); g;
90            Next
100           Printer.Print
110       Next

120       Printer.Print
130       For n = 0 To grdAnalyte.Rows - 1
140           Printer.Print grdAnalyte.TextMatrix(n, 0); Tab(25); grdAnalyte.TextMatrix(n, 1)
150       Next

160       Printer.EndDoc

170       Exit Sub

cmdPrint_Click_Error:

          Dim strES As String
          Dim intEL As Integer



180       intEL = Erl
190       strES = Err.Description
200       LogError "frmExternalStats", "cmdPrint_Click", intEL, strES


End Sub


Private Sub cmdXL_Click()

10        On Error GoTo cmdXL_Click_Error

20        ExportFlexGrid g, Me

30        Exit Sub

cmdXL_Click_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmExternalStats", "cmdXL_Click", intEL, strES


End Sub

Private Sub dtFrom_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)

10        On Error GoTo dtFrom_CallbackKeyDown_Error

20        FillList

30        Exit Sub

dtFrom_CallbackKeyDown_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmExternalStats", "dtFrom_CallbackKeyDown", intEL, strES


End Sub


Private Sub dtTo_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)

10        On Error GoTo dtTo_CallbackKeyDown_Error

20        FillList

30        Exit Sub

dtTo_CallbackKeyDown_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmExternalStats", "dtTo_CallbackKeyDown", intEL, strES


End Sub


Private Sub Form_Load()

10        On Error GoTo Form_Load_Error

20        dtFrom = Format$(Now, "dd/mmm/yyyy")
30        dtTo = dtFrom

40        Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer



50        intEL = Erl
60        strES = Err.Description
70        LogError "frmExternalStats", "Form_Load", intEL, strES


End Sub


Private Sub oBetween_Click(Index As Integer)

          Dim upto As String

10        On Error GoTo oBetween_Click_Error

20        dtFrom = BetweenDates(Index, upto)
30        dtTo = upto

40        FillList

50        Exit Sub

oBetween_Click_Error:

          Dim strES As String
          Dim intEL As Integer



60        intEL = Erl
70        strES = Err.Description
80        LogError "frmExternalStats", "oBetween_Click", intEL, strES


End Sub


Private Sub osource_Click(Index As Integer)

10        On Error GoTo osource_Click_Error

20        FillList

30        Exit Sub

osource_Click_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmExternalStats", "osource_Click", intEL, strES


End Sub


