VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.ocx"
Begin VB.Form frmCoagSourceTotals 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire 6 - Coagulation - Totals"
   ClientHeight    =   7695
   ClientLeft      =   1905
   ClientTop       =   1260
   ClientWidth     =   7935
   ForeColor       =   &H80000008&
   Icon            =   "frmCoagSourceTotals.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7695
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton bGraph 
      Caption         =   "Graph"
      Height          =   885
      Left            =   6390
      Picture         =   "frmCoagSourceTotals.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2790
      Width           =   1275
   End
   Begin VB.CommandButton bcancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   795
      Left            =   6390
      Picture         =   "frmCoagSourceTotals.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6750
      Width           =   1245
   End
   Begin VB.CommandButton bprint 
      Appearance      =   0  'Flat
      Caption         =   "&Print"
      Height          =   930
      Left            =   6390
      Picture         =   "frmCoagSourceTotals.frx":0A56
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5715
      Width           =   1245
   End
   Begin VB.Frame Frame2 
      Caption         =   "Between Dates"
      Height          =   2505
      Left            =   60
      TabIndex        =   1
      Top             =   0
      Width           =   7635
      Begin VB.CommandButton bRecalc 
         Caption         =   "Start"
         Height          =   795
         Left            =   5100
         Picture         =   "frmCoagSourceTotals.frx":0D60
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1275
         Width           =   1545
      End
      Begin MSComCtl2.DTPicker calfromdate 
         Height          =   375
         Left            =   2070
         TabIndex        =   11
         Top             =   420
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   661
         _Version        =   393216
         Format          =   16515073
         CurrentDate     =   37868
      End
      Begin MSComCtl2.DTPicker caltodate 
         Height          =   375
         Left            =   4470
         TabIndex        =   10
         Top             =   420
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   661
         _Version        =   393216
         Format          =   16515073
         CurrentDate     =   37868
      End
      Begin ComctlLib.ProgressBar PB 
         Height          =   195
         Left            =   150
         TabIndex        =   9
         Top             =   2190
         Visible         =   0   'False
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   344
         _Version        =   327682
         Appearance      =   1
      End
      Begin Threed.SSOption obetween 
         Height          =   225
         Index           =   6
         Left            =   930
         TabIndex        =   2
         Top             =   900
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Today"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Value           =   -1  'True
      End
      Begin Threed.SSOption obetween 
         Height          =   225
         Index           =   3
         Left            =   1950
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   930
         Width           =   1155
         _Version        =   65536
         _ExtentX        =   2037
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Last Quarter"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin Threed.SSOption obetween 
         Height          =   225
         Index           =   4
         Left            =   5010
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   990
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Last Full Quarter"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin Threed.SSOption obetween 
         Height          =   225
         Index           =   2
         Left            =   3390
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1230
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Last Full Month"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin Threed.SSOption obetween 
         Height          =   225
         Index           =   5
         Left            =   1890
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Year to Date"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin Threed.SSOption obetween 
         Height          =   225
         Index           =   1
         Left            =   3720
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   960
         Width           =   1065
         _Version        =   65536
         _ExtentX        =   1879
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Last Month"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin Threed.SSOption obetween 
         Height          =   225
         Index           =   0
         Left            =   600
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1170
         Width           =   1065
         _Version        =   65536
         _ExtentX        =   1879
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Last Week"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   4935
      Left            =   1410
      TabIndex        =   0
      Top             =   2610
      Width           =   4635
      _ExtentX        =   8176
      _ExtentY        =   8705
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      FormatString    =   "<Source                    |^Samples |^Tests    |^T/S    "
   End
   Begin Threed.SSOption o 
      Height          =   195
      Index           =   1
      Left            =   90
      TabIndex        =   12
      Top             =   2880
      Width           =   1125
      _Version        =   65536
      _ExtentX        =   1984
      _ExtentY        =   344
      _StockProps     =   78
      Caption         =   "Ward"
      ForeColor       =   0
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
   Begin Threed.SSOption o 
      Height          =   195
      Index           =   0
      Left            =   90
      TabIndex        =   13
      Top             =   2610
      Width           =   1155
      _Version        =   65536
      _ExtentX        =   2037
      _ExtentY        =   344
      _StockProps     =   78
      Caption         =   "Clinician"
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Value           =   -1  'True
   End
   Begin Threed.SSOption o 
      Height          =   195
      Index           =   2
      Left            =   90
      TabIndex        =   18
      Top             =   3150
      Width           =   1125
      _Version        =   65536
      _ExtentX        =   1984
      _ExtentY        =   344
      _StockProps     =   78
      Caption         =   "Gp"
      ForeColor       =   0
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
   Begin ComctlLib.ImageList ImageList1 
      Left            =   2670
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   3
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCoagSourceTotals.frx":106A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCoagSourceTotals.frx":1384
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCoagSourceTotals.frx":169E
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmCoagSourceTotals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit  '© Custom Software 2001

Private Sub bcancel_Click()

10        Unload Me

End Sub

Private Sub bGraph_Click()

10        On Error GoTo bGraph_Click_Error

20        With frmGraph
30            .DrawGraph Me, g
40            .Show 1
50        End With

60        Exit Sub

bGraph_Click_Error:

          Dim strES As String
          Dim intEL As Integer

70        intEL = Erl
80        strES = Err.Description
90        LogError "frmCoagSourceTotals", "bGraph_Click", intEL, strES


End Sub

Private Sub bprint_Click()

          Dim n As Long
          Dim X As Long


10        On Error GoTo bprint_Click_Error

20        Printer.Print "Totals between "; calFromDate; " and "; calToDate

30        Printer.Print
40        For n = 0 To g.Rows - 1
50            g.Row = n
60            For X = 0 To 3
70                g.Col = X
80                Printer.Print Tab(Choose(X + 1, 1, 40, 50, 60)); g;
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
160       LogError "frmCoagSourceTotals", "bPrint_Click", intEL, strES


End Sub

Private Sub brecalc_Click()

10        On Error GoTo brecalc_Click_Error

20        FillGrid

30        Exit Sub

brecalc_Click_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmCoagSourceTotals", "brecalc_Click", intEL, strES


End Sub

Private Sub calfromdate_Click()

10        On Error GoTo calfromdate_Click_Error

20        bReCalc.Visible = True

30        Exit Sub

calfromdate_Click_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmCoagSourceTotals", "calfromdate_Click", intEL, strES


End Sub

Private Sub calfromdate_CloseUp()

10        On Error GoTo calfromdate_CloseUp_Error

20        bReCalc.Visible = True

30        Exit Sub

calfromdate_CloseUp_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmCoagSourceTotals", "calfromdate_CloseUp", intEL, strES


End Sub

Private Sub caltodate_Click()

10        On Error GoTo caltodate_Click_Error

20        bReCalc.Visible = True

30        Exit Sub

caltodate_Click_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmCoagSourceTotals", "caltodate_Click", intEL, strES


End Sub

Private Sub caltodate_CloseUp()

10        On Error GoTo caltodate_CloseUp_Error

20        bReCalc.Visible = True

30        Exit Sub

caltodate_CloseUp_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmCoagSourceTotals", "caltodate_CloseUp", intEL, strES


End Sub

Private Sub FillGrid()

          Dim snr As Recordset
          Dim snc As Recordset
          Dim sql As String
          Dim Total As Long
          Dim tests As Long
          Dim n As Long
          Dim tps As String
          Dim colCounters As New Counters
          Dim c As Counter
          Dim Found As Boolean
          Dim s As String

10        On Error GoTo FillGrid_Error

20        ClearFGrid g

30        sql = "SELECT DISTINCT(D.SampleID), D.* FROM Demographics D, CoagResults R WHERE  " & _
                "D.RunDate BETWEEN '" & Format(calFromDate, "dd/mmm/yyyy") & "' AND '" & Format(calToDate, _
                                                                                                "dd/mmm/yyyy") & "' " & _
                                                                                                "AND D.SampleID = R.SampleID"

40        Set snr = New Recordset
50        RecOpenServer 0, snr, sql

60        If snr.EOF Then
70            Exit Sub
80        End If

90        pb.Visible = True

100       Do While Not snr.EOF
110           If IsDate(Format(snr!Rundate, "dd/MMM/yyyy")) Then
120               If o(0) And Trim(snr!Clinician & "") <> "" Then
130                   Found = False
140                   For Each c In colCounters
150                       If Trim(snr!Clinician & "") <> "" Then
160                           If c.Name = Trim(snr!Clinician & "") Then
170                               Found = True
180                               Exit For
190                           End If
200                       Else
210                           If c.Name = Trim(snr!GP & "") Then
220                               Found = True
230                               Exit For
240                           End If
250                       End If
260                   Next
270                   If Not Found Then
280                       Set c = New Counter
290                       If Trim(snr!Clinician & "") <> "" Then c.Name = Trim( _
                             snr!Clinician & "") Else c.Name = Trim(snr!GP & "") & "/GP"
300                   End If
310                   c.SampleCount = c.SampleCount + 1
320                   sql = "SELECT count('result') as tot from coagresults, coagtestdefinitions WHERE " & _
                            "coagresults.sampleid = '" & snr!SampleID & "' and coagresults.rundate = '" & Format(snr!Rundate, _
                                                                                                                 "dd/MMM/yyyy hh:mm:ss") & "' and coagtestdefinitions.code = coagresults.code"
330                   Set snc = New Recordset
340                   RecOpenServer 0, snc, sql
350                   c.TestCount = c.TestCount + snc!Tot
360                   If Not Found Then
370                       colCounters.Add c
380                   End If
390               ElseIf o(1) And Trim(snr!Ward & "") <> "" Then
400                   Found = False
410                   For Each c In colCounters
420                       If c.Name = Trim(snr!Ward & "") Then
430                           Found = True
440                           Exit For
450                       End If
460                   Next
470                   If Not Found Then
480                       Set c = New Counter
490                       c.Name = Trim(snr!Ward & "")
500                   End If
510                   c.SampleCount = c.SampleCount + 1
520                   sql = "SELECT count('result') as tot from coagresults, coagtestdefinitions WHERE " & _
                            "sampleid = '" & snr!SampleID & "' and rundate = '" & Format(snr!Rundate, _
                                                                                         "dd/MMM/yyyy hh:mm:ss") & "' and coagtestdefinitions.code = coagresults.code"
530                   Set snc = New Recordset
540                   RecOpenServer 0, snc, sql
550                   c.TestCount = c.TestCount + snc!Tot
560                   If Not Found Then
570                       colCounters.Add c
580                   End If
590               ElseIf o(2) And Trim(snr!GP & "") <> "" Then
600                   Found = False
610                   For Each c In colCounters
620                       If c.Name = Trim(snr!GP & "") Then
630                           Found = True
640                           Exit For
650                       End If
660                   Next
670                   If Not Found Then
680                       Set c = New Counter
690                       c.Name = Trim(snr!GP & "")
700                   End If
710                   c.SampleCount = c.SampleCount + 1
720                   sql = "SELECT count('result') as tot from coagresults, coagtestdefinitions WHERE " & _
                            "sampleid = '" & snr!SampleID & "' and rundate = '" & Format(snr!Rundate, _
                                                                                         "dd/MMM/yyyy hh:mm:ss") & "' and coagtestdefinitions.code = coagresults.code"
730                   Set snc = New Recordset
740                   RecOpenServer 0, snc, sql
750                   c.TestCount = c.TestCount + snc!Tot
760                   If Not Found Then
770                       colCounters.Add c
780                   End If
790               End If
800           End If
810           snr.MoveNext
820       Loop

830       For Each c In colCounters
840           With c
850               s = .Name & vbTab & .SampleCount & vbTab & .TestCount & vbTab
860               tps = Format(.TestCount / .SampleCount, "##.00")
870               s = s & tps
880               g.AddItem s
890           End With
900       Next
910       g.AddItem ""

920       If g.Rows = 2 Then
930           pb.Visible = False
940           Exit Sub
950       End If

960       g.Col = 1
970       Total = 0
980       For n = 2 To g.Rows - 1
990           g.Row = n
1000          Total = Total + Val(g)
1010      Next

1020      g.Col = 2
1030      tests = 0
1040      For n = 2 To g.Rows - 1
1050          g.Row = n
1060          tests = tests + Val(g)
1070      Next

1080      If Total * tests <> 0 Then
1090          g.AddItem "Total above" & vbTab & Total & vbTab & tests & vbTab & Format( _
                        tests / Total, ".00")
1100      Else
1110          g.AddItem "Total above"
1120      End If
1130      g.AddItem ""

1140      pb = 25

1150      sql = "SELECT count(distinct sampleid) as tot from coagresults WHERE " & _
                "runDate between '" & _
                Format(calFromDate, "dd/mmm/yyyy") & _
                "' and '" & _
                Format(calToDate, "dd/mmm/yyyy") & "'"
1160      Set snr = New Recordset
1170      RecOpenServer 0, snr, sql


1180      g.AddItem "Total Samples" & vbTab & snr!Tot

1190      pb = 50

1200      sql = "SELECT count(sampleid) as tot from coagresults, coagtestdefinitions WHERE " & _
                "runDate between '" & Format(calFromDate, _
                                             "dd/mmm/yyyy") & "' and '" & Format(calToDate, "dd/mmm/yyyy") & "' and coagtestdefinitions.code = coagresults.code"
1210      Set snr = New Recordset
1220      RecOpenServer 0, snr, sql

1230      n = 0


1240      pb = 100

1250      g.AddItem "Total Tests" & vbTab & vbTab & snr!Tot

1260      g.Refresh

1270      FixG g

1280      pb.Visible = False

1290      Exit Sub

FillGrid_Error:

          Dim strES As String
          Dim intEL As Integer

1300      intEL = Erl
1310      strES = Err.Description
1320      LogError "frmCoagSourceTotals", "FillGrid", intEL, strES

End Sub

Private Sub Form_Load()

10        On Error GoTo Form_Load_Error

20        calFromDate = Format(Now, "dd/mmm/yyyy")
30        calToDate = calFromDate
40        Set_Font Me

50        Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

60        intEL = Erl
70        strES = Err.Description
80        LogError "frmCoagSourceTotals", "Form_Load", intEL, strES


End Sub

Private Sub o_Click(Index As Integer, Value As Integer)

10        On Error GoTo o_Click_Error

20        bReCalc.Visible = True

30        Exit Sub

o_Click_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmCoagSourceTotals", "o_Click", intEL, strES


End Sub

Private Sub oBetween_Click(Index As Integer, Value As Integer)

          Dim upto As String

10        On Error GoTo oBetween_Click_Error

20        calFromDate = BetweenDates(Index, upto)
30        calToDate = upto

40        FillGrid

50        Exit Sub

oBetween_Click_Error:

          Dim strES As String
          Dim intEL As Integer

60        intEL = Erl
70        strES = Err.Description
80        LogError "frmCoagSourceTotals", "oBetween_Click", intEL, strES


End Sub
