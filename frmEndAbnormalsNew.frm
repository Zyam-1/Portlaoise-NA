VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.ocx"
Begin VB.Form frmedf 
   Caption         =   "NetAcquire - Endocrinology Abnormals"
   ClientHeight    =   6135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9060
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEndAbnormals.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   9060
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Between Dates"
      Height          =   1230
      Left            =   4005
      TabIndex        =   39
      Top             =   90
      Width           =   4995
      Begin MSComCtl2.DTPicker calTo 
         Height          =   375
         Left            =   2520
         TabIndex        =   40
         Top             =   405
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   661
         _Version        =   393216
         Format          =   173342721
         CurrentDate     =   37951
      End
      Begin MSComCtl2.DTPicker calFrom 
         Height          =   375
         Left            =   240
         TabIndex        =   41
         Top             =   420
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         Format          =   173342721
         CurrentDate     =   37951
      End
      Begin Threed.SSOption o 
         Height          =   255
         Index           =   0
         Left            =   1710
         TabIndex        =   42
         Top             =   900
         Width           =   525
         _Version        =   65536
         _ExtentX        =   926
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "All"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin Threed.SSOption o 
         Height          =   255
         Index           =   1
         Left            =   2295
         TabIndex        =   43
         Top             =   900
         Width           =   2070
         _Version        =   65536
         _ExtentX        =   3651
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "Only Abnormals"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   -1  'True
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3555
      Left            =   4005
      TabIndex        =   4
      Top             =   1305
      Width           =   4965
      Begin VB.ComboBox lAnalyte 
         Height          =   360
         Left            =   495
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   225
         Width           =   2415
      End
      Begin VB.TextBox tLow 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3180
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   1350
         Width           =   825
      End
      Begin VB.TextBox tHigh 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   4050
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   1350
         Width           =   825
      End
      Begin VB.CommandButton bRecalc 
         Caption         =   "Start"
         Height          =   720
         Left            =   3330
         Picture         =   "frmEndAbnormals.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   360
         Width           =   1455
      End
      Begin ComctlLib.ProgressBar PB 
         Height          =   195
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Visible         =   0   'False
         Width           =   4965
         _ExtentX        =   8758
         _ExtentY        =   344
         _Version        =   327682
         Appearance      =   1
      End
      Begin Threed.SSRibbon b 
         Height          =   285
         Index           =   5
         Left            =   2250
         TabIndex        =   9
         Top             =   2490
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   503
         _StockProps     =   65
         BackColor       =   12632256
         Value           =   -1  'True
         PictureDnChange =   2
         PictureUp       =   "frmEndAbnormals.frx":0614
      End
      Begin Threed.SSRibbon b 
         Height          =   285
         Index           =   4
         Left            =   1395
         TabIndex        =   10
         Top             =   2490
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   503
         _StockProps     =   65
         BackColor       =   12632256
         GroupAllowAllUp =   -1  'True
         PictureUp       =   "frmEndAbnormals.frx":0BD2
      End
      Begin Threed.SSRibbon b 
         Height          =   285
         Index           =   3
         Left            =   540
         TabIndex        =   11
         Top             =   2490
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   503
         _StockProps     =   65
         BackColor       =   12632256
         GroupAllowAllUp =   -1  'True
         PictureUp       =   "frmEndAbnormals.frx":1204
      End
      Begin Threed.SSRibbon b 
         Height          =   285
         Index           =   2
         Left            =   2220
         TabIndex        =   12
         Top             =   900
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   503
         _StockProps     =   65
         BackColor       =   12632256
         GroupAllowAllUp =   -1  'True
         PictureUp       =   "frmEndAbnormals.frx":17C2
      End
      Begin Threed.SSRibbon b 
         Height          =   285
         Index           =   1
         Left            =   1395
         TabIndex        =   13
         Top             =   900
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   503
         _StockProps     =   65
         BackColor       =   12632256
         GroupAllowAllUp =   -1  'True
         PictureUp       =   "frmEndAbnormals.frx":1D80
      End
      Begin Threed.SSRibbon b 
         Height          =   285
         Index           =   0
         Left            =   540
         TabIndex        =   14
         Top             =   900
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   503
         _StockProps     =   65
         BackColor       =   12632256
         GroupAllowAllUp =   -1  'True
         PictureUp       =   "frmEndAbnormals.frx":23B2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Normal Ranges"
         Height          =   195
         Left            =   1170
         TabIndex        =   38
         Top             =   690
         Width           =   1095
      End
      Begin VB.Label l 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   0
         Left            =   540
         TabIndex        =   37
         Top             =   1230
         Width           =   720
      End
      Begin VB.Label l 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   2
         Left            =   1410
         TabIndex        =   36
         Top             =   1230
         Width           =   720
      End
      Begin VB.Label l 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   4
         Left            =   2220
         TabIndex        =   35
         Top             =   1230
         Width           =   720
      End
      Begin VB.Label l 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   1
         Left            =   540
         TabIndex        =   34
         Top             =   1500
         Width           =   720
      End
      Begin VB.Label l 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   3
         Left            =   1410
         TabIndex        =   33
         Top             =   1500
         Width           =   720
      End
      Begin VB.Label l 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   5
         Left            =   2220
         TabIndex        =   32
         Top             =   1500
         Width           =   720
      End
      Begin VB.Label l 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   8
         Left            =   1410
         TabIndex        =   31
         Top             =   2820
         Width           =   720
      End
      Begin VB.Label l 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   6
         Left            =   540
         TabIndex        =   30
         Top             =   2820
         Width           =   720
      End
      Begin VB.Label l 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   10
         Left            =   2250
         TabIndex        =   29
         Top             =   2820
         Width           =   720
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Low"
         Height          =   195
         Left            =   180
         TabIndex        =   28
         Top             =   3090
         Width           =   300
      End
      Begin VB.Label l 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   7
         Left            =   540
         TabIndex        =   27
         Top             =   3090
         Width           =   720
      End
      Begin VB.Label l 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   9
         Left            =   1410
         TabIndex        =   26
         Top             =   3090
         Width           =   720
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "High"
         Height          =   195
         Left            =   150
         TabIndex        =   25
         Top             =   1260
         Width           =   330
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Low"
         Height          =   195
         Left            =   180
         TabIndex        =   24
         Top             =   1530
         Width           =   300
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "High"
         Height          =   195
         Left            =   150
         TabIndex        =   23
         Top             =   2850
         Width           =   330
      End
      Begin VB.Label l 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   11
         Left            =   2250
         TabIndex        =   22
         Top             =   3090
         Width           =   720
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Flag Ranges"
         Height          =   195
         Left            =   1260
         TabIndex        =   21
         Top             =   2280
         Width           =   900
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "De-Select Range"
         Height          =   195
         Left            =   3390
         TabIndex        =   20
         Top             =   1140
         Width           =   1230
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Over Range"
         Height          =   195
         Left            =   3630
         TabIndex        =   19
         Top             =   1680
         Width           =   885
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Under Range"
         Height          =   195
         Left            =   3600
         TabIndex        =   18
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label lOver 
         Alignment       =   2  'Center
         BackColor       =   &H008080FF&
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
         Height          =   315
         Left            =   3630
         TabIndex        =   17
         Top             =   1890
         Width           =   900
      End
      Begin VB.Label lUnder 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
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
         Height          =   315
         Left            =   3630
         TabIndex        =   16
         Top             =   2610
         Width           =   900
      End
   End
   Begin VB.CommandButton bCancel 
      BackColor       =   &H80000004&
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   795
      Left            =   7470
      Picture         =   "frmEndAbnormals.frx":2970
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CommandButton bPrintList 
      Caption         =   "Print List"
      Height          =   795
      Left            =   4095
      Picture         =   "frmEndAbnormals.frx":2C7A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5040
      Width           =   1245
   End
   Begin VB.CommandButton bPrintReports 
      Caption         =   "Print Reports"
      Height          =   795
      Left            =   5715
      Picture         =   "frmEndAbnormals.frx":2F84
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5040
      Width           =   1380
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   5940
      Left            =   45
      TabIndex        =   3
      Top             =   45
      Width           =   3915
      _ExtentX        =   6906
      _ExtentY        =   10478
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      ScrollBars      =   2
      FormatString    =   "^Date                  |^Run #         |^Result   "
   End
End
Attribute VB_Name = "frmedf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit  '� Custom Software 2001

Private Sub b_Click(Index As Integer, Value As Integer)


10        On Error GoTo b_Click_Error

20        If Value = False Then
30            tLow.Enabled = True
40            tHigh.Enabled = True
50            tLow.SelStart = 0
60            tLow.SelLength = Len(tLow)
70            tLow.SetFocus
80        Else
90            tLow.Enabled = False
100           tHigh.Enabled = False
110           tLow = l((Index * 2) + 1)
120           tHigh = l(Index * 2)
130           FillG
140       End If



150       Exit Sub

b_Click_Error:

          Dim strES As String
          Dim intEL As Integer



160       intEL = Erl
170       strES = Err.Description
180       LogError "frmEndAbnormals", "b_Click", intEL, strES


End Sub

Private Sub bcancel_Click()

10        Unload Me

End Sub

Private Sub bPrintList_Click()

          Dim n As Long
          Dim s As String

10        On Error GoTo bPrintList_Click_Error

20        For n = 0 To g.Rows - 1
30            g.Row = n
40            g.Col = 0
50            s = g & vbTab
60            g.Col = 1
70            s = s & g & vbTab
80            g.Col = 2
90            s = s & g
100           Printer.Print s
110       Next

120       Printer.EndDoc

130       Exit Sub

bPrintList_Click_Error:

          Dim strES As String
          Dim intEL As Integer



140       intEL = Erl
150       strES = Err.Description
160       LogError "frmEndAbnormals", "bPrintList_Click", intEL, strES


End Sub

Private Sub bPrintReports_Click()

          Dim n As Long


10        On Error GoTo bPrintReports_Click_Error

20        If g.Rows < 2 Then Exit Sub
30        If g.Rows > 20 Then
40            If iMsg("Print " & g.Rows & " results?", vbQuestion + vbYesNo) = vbNo Then
50                Exit Sub
60            End If
70        End If

80        g.Col = 1
90        For n = 1 To g.Rows - 1
100           g.Row = n
110           PrintResultEndWin g.TextMatrix(n, 1)
120       Next



130       Exit Sub

bPrintReports_Click_Error:

          Dim strES As String
          Dim intEL As Integer



140       intEL = Erl
150       strES = Err.Description
160       LogError "frmEndAbnormals", "bPrintReports_Click", intEL, strES


End Sub

Private Sub brecalc_Click()

10        On Error GoTo brecalc_Click_Error

20        FillG

30        Exit Sub

brecalc_Click_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmEndAbnormals", "brecalc_Click", intEL, strES


End Sub

Private Sub CalcOutsideRange()

          Dim n As Long
          Dim l As Single
          Dim h As Single
          Dim LCounter As Double
          Dim HCounter As Double


10        On Error GoTo CalcOutsideRange_Error

20        If g.Rows < 2 Then
30            lUnder = "0"
40            lOver = "0"
50            Exit Sub
60        End If

70        l = Val(tLow)
80        h = Val(tHigh)
90        LCounter = 0
100       HCounter = 0

110       g.Col = 2

120       g.Visible = False

130       For n = g.Rows - 1 To 1 Step -1
140           g.Row = n
150           If Val(g) > h Then
160               HCounter = HCounter + 1
170               g.CellBackColor = &H8080FF    'red
                  'set red
180           ElseIf Val(g) < l Then
190               LCounter = LCounter + 1
200               g.CellBackColor = &HFFFF80    'blue
210           ElseIf InStr(1, g, ">") Then
220               HCounter = HCounter + 1
230               g.CellBackColor = &H8080FF    'red
240           Else
250               If o(0) Then
260                   g.CellBackColor = 0
270               Else
280                   If g.Rows = 2 Then g.AddItem "", g.Row + 1
290                   g.RemoveItem g.Row

300               End If
310           End If
320       Next

330       g.Visible = True

340       lUnder = LCounter
350       lOver = HCounter



360       Exit Sub

CalcOutsideRange_Error:

          Dim strES As String
          Dim intEL As Integer



370       intEL = Erl
380       strES = Err.Description
390       LogError "frmEndAbnormals", "CalcOutsideRange", intEL, strES


End Sub

Private Sub calFrom_Click()

10        On Error GoTo calFrom_Click_Error

20        FillG

30        Exit Sub

calFrom_Click_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmEndAbnormals", "calFrom_Click", intEL, strES


End Sub

Private Sub calTo_Click()

10        On Error GoTo calTo_Click_Error

20        FillG

30        Exit Sub

calTo_Click_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmEndAbnormals", "calTo_Click", intEL, strES


End Sub

Private Sub FillG()

          Dim sn As New Recordset
          Dim sql As String
          Dim TestNumber As String
          Dim s As String
          Dim DP As Long
          Dim RecCount, pCount As Long
          Dim n As Long
          Dim snr As Recordset

10        On Error GoTo FillG_Error

20        ClearFGrid g

30        If lAnalyte = "" Then Exit Sub

40        TestNumber = ""

50        sql = "SELECT * from endtestdefinitions WHERE longname = '" & lAnalyte & "'"
60        Set sn = New Recordset
70        RecOpenServer 0, sn, sql
80        If Not sn.EOF Then
90            TestNumber = sn!Code
100           DP = sn!DP
110       End If

120       If TestNumber = "" Then
130           g.Visible = True
140           Exit Sub
150       End If

160       sql = "SELECT * from endresults WHERE " & _
                "(runtime between '" & _
                Format(calFrom, "dd/mmm/yyyy") & " 00:00:00' and '" & _
                Format(calTo, "dd/mmm/yyyy") & " 23:59:59') " & _
                "and code = '" & TestNumber & _
                "' order by runTime"
170       Set sn = New Recordset
180       RecOpenServer 0, sn, sql

190       If sn.EOF Then

200           g.Visible = True
210           Exit Sub
220       End If

230       g.Visible = False

240       If Not sn.EOF Then

250           sql = "SELECT count(sampleid) as tot from endresults WHERE " & _
                    "(runtime between '" & _
                    Format(calFrom, "dd/mmm/yyyy") & " 00:00:00' and '" & _
                    Format(calTo, "dd/mmm/yyyy") & " 23:59:59') " & _
                    "and code = '" & TestNumber & _
                    "'"
260           Set snr = New Recordset
270           RecOpenServer 0, snr, sql
280           n = Val(snr!Tot)
290       End If

300       pb.Visible = True
310       pCount = n / 100
320       RecCount = 0
330       Do While Not sn.EOF
340           If RecCount = Int(pCount) Then
350               If pb < pb.Max Then pb = pb + 1
360               RecCount = 0
370           Else
380               RecCount = RecCount + 1
390           End If
400           s = Format(sn!RunTime, "dd/mm/yyyy") & vbTab & _
                  sn!SampleID & vbTab
410           Select Case DP
              Case 0: s = s & Format(sn!Result, "######")
420           Case 1: s = s & Format(sn!Result, "####.0")
430           Case 2: s = s & Format(sn!Result, "###.00")
440           Case 3: s = s & Format(sn!Result, "##.000")
450           End Select
460           g.AddItem s
470           sn.MoveNext
480       Loop

490       pb.Visible = False

500       FixG g

510       CalcOutsideRange

520       Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer

530       intEL = Erl
540       strES = Err.Description
550       LogError "frmEndAbnormals", "FillG", intEL, strES, sql

End Sub

Private Sub Form_Load()

          Dim sql As String
          Dim tb As New Recordset


10        On Error GoTo Form_Load_Error

20        calFrom = DateAdd("m", -1, Now)
30        calTo = Now

40        lAnalyte.Clear

50        Set_Font Me

60        sql = "SELECT distinct(longname) from endtestdefinitions WHERE inuse = 1"
70        Set tb = New Recordset
80        RecOpenServer 0, tb, sql
90        Do While Not tb.EOF
100           lAnalyte.AddItem Trim(tb!LongName)
110           tb.MoveNext
120       Loop




130       Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer



140       intEL = Erl
150       strES = Err.Description
160       LogError "frmEndAbnormals", "Form_Load", intEL, strES


End Sub

Private Sub g_Click()

          Dim tb As New Recordset
          Dim s As String
          Dim sql As String

10        On Error GoTo g_Click_Error

20        If g.Row < 1 Then Exit Sub

30        g.Col = 1

40        sql = "SELECT * FROM Demographics WHERE " & _
                "SampleID = '" & g & "'"

50        Set tb = New Recordset
60        RecOpenServer 0, tb, sql

70        If Not tb.EOF Then
80            s = "   Run Number : " & g & vbCrLf & _
                  "      Patient : " & tb!PatName & vbCrLf & _
                  "        Chart : " & tb!Chart & vbCrLf & _
                  "Date of Birth : " & Format(tb!Dob, "dd/mm/yyyy") & ""
90            tb.Close
100           iMsg s, vbInformation + vbOKOnly
110       End If

120       Exit Sub

g_Click_Error:

          Dim strES As String
          Dim intEL As Integer



130       intEL = Erl
140       strES = Err.Description
150       LogError "frmEndAbnormals", "g_Click", intEL, strES

End Sub

Private Sub lAnalyte_Click()

          Dim n As Long
          Dim sql As String
          Dim tb As New Recordset



10        On Error GoTo lAnalyte_Click_Error

20        sql = "SELECT * from endtestdefinitions WHERE longname = '" & lAnalyte & "'"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql

50        If Not tb.EOF Then
60            l(0) = tb!MaleHigh
70            l(1) = tb!MaleLow
80            l(2) = tb!FemaleHigh
90            l(3) = tb!FemaleLow
100           l(4) = IIf(Val(tb!MaleHigh) > Val(tb!FemaleHigh), tb!MaleHigh, tb!FemaleHigh)
110           l(5) = IIf(Val(tb!MaleLow) < Val(tb!FemaleLow), tb!MaleLow, tb!FemaleLow)

120           l(6) = tb!FlagMaleHigh
130           l(7) = tb!FlagMaleLow
140           l(8) = tb!FlagFemaleHigh
150           l(9) = tb!FlagFemaleLow
160           l(10) = IIf(Val(tb!FlagMaleHigh) > Val(tb!FlagFemaleHigh), tb!FlagMaleHigh, tb!FlagFemaleHigh)
170           l(11) = IIf(Val(tb!FlagMaleLow) < Val(tb!FlagFemaleLow), tb!FlagMaleLow, tb!FlagFemaleLow)
180       End If

190       For n = 0 To 5
200           If b(n) Then
210               tLow = l((n * 2) + 1)
220               tHigh = l(n * 2)
230           End If
240       Next

250       FillG




260       Exit Sub

lAnalyte_Click_Error:

          Dim strES As String
          Dim intEL As Integer



270       intEL = Erl
280       strES = Err.Description
290       LogError "frmEndAbnormals", "lAnalyte_Click", intEL, strES, sql


End Sub

Private Sub o_Click(Index As Integer, Value As Integer)

10        On Error GoTo o_Click_Error

20        If Index = 0 Then
30            FillG
40        Else
50            RemoveNormals
60        End If

70        Exit Sub

o_Click_Error:

          Dim strES As String
          Dim intEL As Integer



80        intEL = Erl
90        strES = Err.Description
100       LogError "frmEndAbnormals", "o_Click", intEL, strES


End Sub

Private Sub RemoveNormals()

          Dim n As Long



10        On Error GoTo RemoveNormals_Error

20        If g.Rows < 3 Then Exit Sub

30        g.Col = 2

40        g.Visible = False

50        For n = g.Rows - 1 To 2 Step -1
60            g.Row = n
70            If g.CellBackColor <> &H8080FF And g.CellBackColor <> &HFFFF80 Then
80                g.RemoveItem g.Row
90            End If
100       Next

110       g.Visible = True


120       Exit Sub

RemoveNormals_Error:

          Dim strES As String
          Dim intEL As Integer



130       intEL = Erl
140       strES = Err.Description
150       LogError "frmEndAbnormals", "RemoveNormals", intEL, strES


End Sub


