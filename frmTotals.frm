VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmTotals 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Total Statistics"
   ClientHeight    =   7695
   ClientLeft      =   930
   ClientTop       =   810
   ClientWidth     =   10305
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
   Icon            =   "frmTotals.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7695
   ScaleWidth      =   10305
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox fraProgress 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   1500
      ScaleHeight     =   525
      ScaleWidth      =   3825
      TabIndex        =   24
      Top             =   3480
      Visible         =   0   'False
      Width           =   3855
      Begin MSComctlLib.ProgressBar pbProgress 
         Height          =   255
         Left            =   0
         TabIndex        =   25
         Top             =   240
         Width           =   3840
         _ExtentX        =   6773
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         Min             =   1
      End
      Begin VB.Label lblProgress 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Fetching results ..."
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         TabIndex        =   26
         Top             =   0
         Width           =   3840
      End
   End
   Begin VB.CommandButton cmdExcel 
      Caption         =   "Export to Excel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5490
      Picture         =   "frmTotals.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   6570
      Width           =   1305
   End
   Begin VB.CommandButton bReCalc 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6930
      Picture         =   "frmTotals.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   90
      Width           =   1305
   End
   Begin VB.ComboBox cmbHosp 
      Height          =   315
      Left            =   7140
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   1380
      Width           =   2715
   End
   Begin VB.OptionButton o 
      Caption         =   "Clinicians"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   8730
      TabIndex        =   14
      Top             =   1935
      Value           =   -1  'True
      Width           =   1245
   End
   Begin VB.OptionButton o 
      Caption         =   "Wards"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   7830
      TabIndex        =   13
      Top             =   1935
      Width           =   885
   End
   Begin VB.OptionButton o 
      Caption         =   "G.P.s"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   7065
      TabIndex        =   12
      Top             =   1935
      Width           =   885
   End
   Begin VB.PictureBox SSPanel1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Index           =   1
      Left            =   300
      ScaleHeight     =   1095
      ScaleWidth      =   5625
      TabIndex        =   4
      Top             =   30
      Width           =   5685
      Begin MSComCtl2.DTPicker calFrom 
         Height          =   375
         Left            =   585
         TabIndex        =   15
         Top             =   30
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   661
         _Version        =   393216
         Format          =   16515073
         CurrentDate     =   38037
      End
      Begin VB.OptionButton oBetween 
         Caption         =   "Today"
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
         Index           =   6
         Left            =   120
         TabIndex        =   11
         Top             =   450
         Value           =   -1  'True
         Width           =   795
      End
      Begin VB.OptionButton oBetween 
         Caption         =   "Year To Date"
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
         Index           =   5
         Left            =   4230
         TabIndex        =   10
         Top             =   450
         Width           =   1305
      End
      Begin VB.OptionButton oBetween 
         Caption         =   "Last Full Quarter"
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
         Index           =   4
         Left            =   2850
         TabIndex        =   9
         Top             =   720
         Width           =   1515
      End
      Begin VB.OptionButton oBetween 
         Caption         =   "Last Quarter"
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
         Index           =   3
         Left            =   2850
         TabIndex        =   8
         Top             =   450
         Width           =   1185
      End
      Begin VB.OptionButton oBetween 
         Caption         =   "Last Full Month"
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
         Left            =   1350
         TabIndex        =   7
         Top             =   690
         Width           =   1395
      End
      Begin VB.OptionButton oBetween 
         Caption         =   "Last Month"
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
         Left            =   1350
         TabIndex        =   6
         Top             =   420
         Width           =   1155
      End
      Begin VB.OptionButton oBetween 
         Caption         =   "Last Week"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   690
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker calTo 
         Height          =   375
         Left            =   2610
         TabIndex        =   16
         Top             =   30
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   661
         _Version        =   393216
         Format          =   16515073
         CurrentDate     =   38037
      End
      Begin VB.Label Label2 
         Caption         =   "To"
         Height          =   285
         Left            =   2340
         TabIndex        =   21
         Top             =   90
         Width           =   285
      End
      Begin VB.Label Label1 
         Caption         =   "From"
         Height          =   330
         Left            =   135
         TabIndex        =   20
         Top             =   90
         Width           =   465
      End
   End
   Begin VB.ListBox List1 
      BackColor       =   &H8000000F&
      ForeColor       =   &H8000000D&
      Height          =   4155
      Left            =   7065
      TabIndex        =   2
      Top             =   2205
      Width           =   2895
   End
   Begin VB.CommandButton bprint 
      Appearance      =   0  'Flat
      Caption         =   "&Print"
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
      Left            =   11580
      Picture         =   "frmTotals.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5700
      Width           =   1245
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
      Height          =   855
      Left            =   8670
      Picture         =   "frmTotals.frx":0C28
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6570
      Width           =   1305
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   5115
      Left            =   180
      TabIndex        =   3
      Top             =   1260
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   9022
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      GridLines       =   2
      FormatString    =   "<Source               |<Samples   |<% Sample |<Tests          |<%Test  |<T/S        "
   End
   Begin VB.CommandButton bGraph 
      Caption         =   "Graph"
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
      Left            =   11610
      Picture         =   "frmTotals.frx":0F32
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6690
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label lblExcelInfo 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exporting..."
      Height          =   345
      Left            =   4140
      TabIndex        =   23
      Top             =   6810
      Visible         =   0   'False
      Width           =   1275
   End
End
Attribute VB_Name = "frmTotals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public TotDept As String

Private Sub bcancel_Click()

10        Unload Me

End Sub

Private Sub bGraph_Click()

10        On Error GoTo bGraph_Click_Error

20        If g.TextMatrix(1, 1) = "" Then Exit Sub

30        With frmGraph
40            .DrawGraph Me, g
50            .Show 1
60        End With

70        Exit Sub

bGraph_Click_Error:

          Dim strES As String
          Dim intEL As Integer



80        intEL = Erl
90        strES = Err.Description
100       LogError "frmTotals", "bGraph_Click", intEL, strES


End Sub

Private Sub bprint_Click()

          Dim n As Long
          Dim X As Long

10        On Error GoTo bprint_Click_Error

20        Printer.Print "Totals: "; calFrom; " to "; calTo

30        Printer.Print
40        For n = 0 To g.Rows - 1
50            g.Row = n
60            For X = 0 To 4
70                g.Col = X
80                Printer.Print Tab(Choose(X + 1, 1, 40, 50, 60, 80)); g;
90            Next
100           Printer.Print
110       Next

120       Printer.EndDoc

130       Exit Sub

bprint_Click_Error:

          Dim strES As String
          Dim intEL As Integer



140       intEL = Erl
150       strES = Err.Description
160       LogError "frmTotals", "bPrint_Click", intEL, strES


End Sub

Private Sub brecalc_Click()

10        On Error GoTo brecalc_Click_Error

20        bReCalc.Visible = False

30        FillGrid

40        Exit Sub

brecalc_Click_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmTotals", "brecalc_Click", intEL, strES

End Sub

Private Sub FillList(ByVal Source As Long)

          Dim sql As String
          Dim tb As New Recordset

10        On Error GoTo FillList_Error

20        List1.Clear
30        Select Case Source
          Case 0:
40            sql = "SELECT DISTINCT(C.Text) FROM Clinicians C JOIN Lists L " & _
                    "ON C.HospitalCode = L.Code " & _
                    "WHERE L.Text = '" & AddTicks(cmbHosp) & "' " & _
                    "AND L.ListType = 'HO' " & _
                    "ORDER BY C.Text"
50        Case 1:
60            sql = "SELECT DISTINCT(W.Text) FROM Wards W JOIN Lists L " & _
                    "ON W.HospitalCode = L.Code " & _
                    "WHERE L.Text = '" & AddTicks(cmbHosp) & "' " & _
                    "AND L.ListType = 'HO' " & _
                    "ORDER BY W.Text"
70        Case 2:
80            sql = "SELECT DISTINCT(G.Text) FROM GPs G JOIN Lists L " & _
                    "ON G.HospitalCode = L.Code " & _
                    "WHERE L.Text = '" & AddTicks(cmbHosp) & "' " & _
                    "AND L.ListType = 'HO' " & _
                    "ORDER BY G.Text"
90        End Select

100       Set tb = New Recordset
110       RecOpenServer 0, tb, sql
120       Do While Not tb.EOF
130           List1.AddItem tb!Text & ""
140           tb.MoveNext
150       Loop

160       Exit Sub

FillList_Error:

          Dim strES As String
          Dim intEL As Integer

170       intEL = Erl
180       strES = Err.Description
190       LogError "frmTotals", "FillList", intEL, strES, sql

End Sub

Private Sub FillGrid()

          Dim tb As New Recordset
          Dim sql As String
          Dim Total As Long
          Dim Samples As Long
          Dim tests As Long
          Dim n As Long
          Dim At As Long
          Dim An As Long
          Dim tps As String
          Dim sn As New Recordset
          Dim Source As String
          Dim FromDate As String
          Dim ToDate As String
          Dim SourceTable As String
          Dim MonthIndex As Integer
          Dim StartDate As Date
          Dim EndDate As Date
          Dim SrcUpdated As Boolean
          Dim DiffMonths As Integer

10        On Error GoTo FillGrid_Error

20        ClearFGrid g

30        FromDate = Format$(calFrom, "dd/mmm/yyyy") & " 00:00:00"
40        ToDate = Format$(calTo, "dd/mmm/yyyy") & " 23:59:59"

50        If o(0) Then
60            Source = "Clinician"
70            SourceTable = "Clinicians"
80        ElseIf o(1) Then
90            Source = "Ward"
100           SourceTable = "Wards"
110       ElseIf o(2) Then
120           Source = "GP"
130           SourceTable = "GPs"
140       End If

150       StartDate = FromDate
160       pbProgress.Value = 1
170       DiffMonths = DateDiff("m", FromDate, ToDate)
180       If DiffMonths = 0 Then
190           EndDate = ToDate
200           pbProgress.Max = 2
210           DiffMonths = DiffMonths + 1
220       Else
230           EndDate = DateAdd("m", 1, calFrom)
240           EndDate = "01/" & Month(EndDate) & "/" & Year(EndDate)
250           EndDate = DateAdd("d", -1, EndDate)
260           pbProgress.Max = DiffMonths + 1
270       End If



280       fraProgress.Visible = True
290       For MonthIndex = 1 To DiffMonths
300           pbProgress.Value = pbProgress.Value + 1
310           lblProgress = "Fetching results ... (" & Int(pbProgress.Value * 100 / pbProgress.Max) & " %)"
320           lblProgress.Refresh
330           sql = "SELECT " & Source & " AS Src, Count(D.SampleID) AS TotTests, Count(DISTINCT D.SampleID) AS TotSamples " & _
                    "FROM Demographics D INNER JOIN " & TotDept & "Results R ON D.SampleID = R.SampleID " & _
                    "WHERE D." & Source & " IN " & _
                    "(SELECT DISTINCT(C.Text) FROM " & SourceTable & " C JOIN Lists L ON C.HospitalCode = L.Code " & _
                    "WHERE L.Text = '" & cmbHosp & "' AND L.ListType = 'HO') " & _
                    "AND D.RunDate BETWEEN '" & Format(StartDate, "dd/mmm/yyyy 00:00:00") & _
                    "' AND '" & Format(EndDate, "dd/mmm/yyyy 23:59:59") & "' " & _
                    "AND D.Hospital = '" & cmbHosp & "' " & _
                    "GROUP BY " & Source

340           Set tb = New Recordset
350           RecOpenClient 0, tb, sql
360           If Not tb.EOF Then
370               While Not tb.EOF
380                   SrcUpdated = False
390                   For n = 1 To g.Rows - 1
400                       If g.TextMatrix(n, 0) = tb!Src Then
410                           g.TextMatrix(n, 1) = Val(g.TextMatrix(n, 1)) + tb!TotSamples
420                           g.TextMatrix(n, 3) = Val(g.TextMatrix(n, 3)) + tb!TotTests
430                           SrcUpdated = True
440                       End If
450                   Next n
460                   If Not SrcUpdated Then
470                       g.AddItem tb!Src & vbTab & tb!TotSamples & vbTab & vbTab & tb!TotTests
480                   End If

490                   tb.MoveNext
500               Wend
510           End If
520           StartDate = DateAdd("d", 1, EndDate)
530           EndDate = DateAdd("m", 1, StartDate)
540           EndDate = "01/" & Month(EndDate) & "/" & Year(EndDate)
550           EndDate = DateAdd("d", -1, EndDate)
560           If MonthIndex = DiffMonths And DateDiff("d", ToDate, EndDate) > 0 Then
570               EndDate = ToDate
580           End If
590       Next MonthIndex

600       For n = 1 To g.Rows - 1
610           If g.TextMatrix(n, 0) <> "" And Val(g.TextMatrix(n, 1)) > 0 And Val(g.TextMatrix(n, 3)) > 0 Then
620               g.TextMatrix(n, 5) = Format$(g.TextMatrix(n, 3) / g.TextMatrix(n, 1), "##.00")
630           End If
640       Next n



          'For n = 0 To List1.ListCount - 1
          '    List1.Selected(n) = True
          '
          '
          '
          '    sql = "SELECT TotTests = (SELECT COUNT(*) FROM " & TotDept & "Results " & _
               '                         "WHERE SampleID IN (SELECT DISTINCT D.SampleID FROM Demographics D INNER JOIN " & TotDept & "Results R " & _
               '                         "ON D.SampleID = R.SampleID " & _
               '                         "WHERE D.RunDate BETWEEN '" & FromDate & "' AND '" & ToDate & "' " & _
               '                         "AND " & Source & " = '" & AddTicks(List1.List(n)) & "' " & _
               '                         "AND D.Hospital = '" & cmbHosp & "')), " & _
               '                         "TotSamples = (SELECT  COUNT(DISTINCT D.SampleID) " & _
               '                         "FROM Demographics D INNER JOIN " & TotDept & "Results R " & _
               '                         "ON D.SampleID = R.SampleID " & _
               '                         "WHERE D.RunDate BETWEEN '" & FromDate & "' AND '" & ToDate & "' " & _
               '                         "AND " & Source & " = '" & AddTicks(List1.List(n)) & "' " & _
               '                         "AND D.Hospital = '" & cmbHosp & "')"
          '
          '    Set tb = New Recordset
          '    RecOpenClient 0, tb, sql
          '    Samples = tb!TotSamples
          '    tests = tb!TotTests
          '    If tests > 0 And Samples > 0 Then
          '        tps = Format$(tests / Samples, "##.00")
          '        g.AddItem List1.List(n) & vbTab & Samples & vbTab & vbTab & tests & vbTab & vbTab & tps
          '        g.Refresh
          '    End If
          'Next
650       g.AddItem ""

660       If g.Rows = 2 Then Exit Sub

670       g.Col = 1
680       Total = 0
690       For n = 2 To g.Rows - 1
700           g.Row = n
710           Total = Total + Val(g)
720       Next

730       g.Col = 3
740       tests = 0
750       For n = 2 To g.Rows - 1
760           g.Row = n
770           tests = tests + Val(g)
780       Next

790       If Total <> 0 And tests <> 0 Then
800           g.AddItem "Sub Total" & vbTab & Total & vbTab & vbTab & tests & vbTab & vbTab & Format$(tests / Total, ".00")
810       Else
820           g.AddItem "Sub Total"
830       End If
840       g.Refresh

850       sql = "SELECT count(DISTINCT sampleid) as tot " & _
                "FROM " & TotDept & "results WHERE " & _
                "runtime between '" & FromDate & "' " & _
                "and '" & ToDate & "'"
860       Set sn = New Recordset
870       RecOpenServer 0, sn, sql
880       If Not sn.EOF Then
890           sn.MoveLast
900           An = Format(sn!Tot)
910       End If

920       sql = "SELECT count(*) as tot from " & TotDept & "results WHERE " & _
                "(Runtime between '" & FromDate & "' and '" & ToDate & "') "
930       Set tb = New Recordset
940       RecOpenServer 0, tb, sql
950       Do While Not tb.EOF
960           At = At + Val(tb!Tot)
970           tb.MoveNext
980       Loop

990       If An + At <> 0 Then
1000          g.AddItem ""
1010          g.AddItem "Total" & vbTab & An & vbTab & vbTab & At & vbTab & vbTab & Format$(At / An, ".00")
1020      Else
1030          g.AddItem ""
1040          g.AddItem "Total"
1050      End If

1060      FixG g

1070      For n = 1 To g.Rows - 5
1080          g.TextMatrix(n, 2) = Format((g.TextMatrix(n, 1) / (An / 100)), "0.00")
1090          g.TextMatrix(n, 4) = Format((g.TextMatrix(n, 3) / (At / 100)), "0.00")
1100      Next

1110      If g.Rows > 1 And g.TextMatrix(1, 0) <> "" Then
1120          g.TextMatrix(g.Rows - 3, 2) = Format((g.TextMatrix(g.Rows - 3, 1) / (An / 100)), "0.00")
1130          g.TextMatrix(g.Rows - 3, 4) = Format((g.TextMatrix(g.Rows - 3, 3) / (At / 100)), "0.00")
1140      End If

1150      fraProgress.Visible = False

1160      Exit Sub

FillGrid_Error:

          Dim strES As String
          Dim intEL As Integer


1170      intEL = Erl
1180      strES = Err.Description
1190      LogError "frmTotals", "FillGrid", intEL, strES, sql
1200      fraProgress.Visible = False

End Sub

Private Sub calFrom_DateClick(ByVal DateClicked As Date)

10        On Error GoTo calFrom_DateClick_Error

20        bReCalc.Visible = True

30        Exit Sub

calFrom_DateClick_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmTotals", "calFrom_DateClick", intEL, strES


End Sub


Private Sub calTo_DateClick(ByVal DateClicked As Date)

10        On Error GoTo calTo_DateClick_Error

20        bReCalc.Visible = True

30        Exit Sub

calTo_DateClick_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmTotals", "calTo_DateClick", intEL, strES


End Sub


Private Sub calFrom_CloseUp()

10        On Error GoTo calFrom_CloseUp_Error

20        bReCalc.Visible = True

30        Exit Sub

calFrom_CloseUp_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmTotals", "calFrom_CloseUp", intEL, strES


End Sub



Private Sub cmbHosp_Click()
          Dim n As Long

10        On Error GoTo cmbHosp_Click_Error

20        For n = 0 To 2
30            If o(n).Value = True Then
40                FillList n
50            End If
60        Next


70        bReCalc.Visible = True

80        Exit Sub

cmbHosp_Click_Error:

          Dim strES As String
          Dim intEL As Integer



90        intEL = Erl
100       strES = Err.Description
110       LogError "frmTotals", "cmbHosp_Click", intEL, strES


End Sub

Private Sub cmdExcel_Click()

10        ExportFlexGrid g, Me

End Sub

Private Sub Form_Activate()

10        On Error GoTo Form_Activate_Error

20        FillHosp
30        Set_Font Me

40        Exit Sub

Form_Activate_Error:

          Dim strES As String
          Dim intEL As Integer



50        intEL = Erl
60        strES = Err.Description
70        LogError "frmTotals", "Form_Activate", intEL, strES


End Sub


Private Sub FillHosp()
          Dim sql As String
          Dim tb As Recordset

10        On Error GoTo FillHosp_Error

20        sql = "SELECT * from lists WHERE listtype = 'HO'"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        Do While Not tb.EOF
60            If UCase(tb!Text) = HospName(0) Then
70                cmbHosp.AddItem tb!Text, 0
80            Else
90                cmbHosp.AddItem tb!Text
100           End If
110           tb.MoveNext
120       Loop

130       cmbHosp.ListIndex = 0

140       Exit Sub

FillHosp_Error:

          Dim strES As String
          Dim intEL As Integer

150       intEL = Erl
160       strES = Err.Description
170       LogError "frmTotals", "FillHosp", intEL, strES, sql

End Sub
Private Sub Form_Load()

10        On Error GoTo Form_Load_Error

20        calFrom = Format$(Now, "dd/mmm/yyyy")
30        calTo = calFrom

40        Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmTotals", "Form_Load", intEL, strES

End Sub


Private Sub o_Click(Index As Integer)

10        On Error GoTo o_Click_Error

20        FillList Index

30        g.Refresh

40        bReCalc.Visible = True

50        Exit Sub

o_Click_Error:

          Dim strES As String
          Dim intEL As Integer



60        intEL = Erl
70        strES = Err.Description
80        LogError "frmTotals", "o_Click", intEL, strES


End Sub

Private Sub oBetween_Click(Index As Integer)

          Dim upto As String

10        On Error GoTo oBetween_Click_Error

20        calFrom = BetweenDates(Index, upto)
30        calTo = upto

40        bReCalc.Visible = True

50        Exit Sub

oBetween_Click_Error:

          Dim strES As String
          Dim intEL As Integer



60        intEL = Erl
70        strES = Err.Description
80        LogError "frmTotals", "oBetween_Click", intEL, strES


End Sub

