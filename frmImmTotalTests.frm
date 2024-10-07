VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.ocx"
Begin VB.Form frmImmTotalTests 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire 6 - Immunology Total Tests"
   ClientHeight    =   5895
   ClientLeft      =   2010
   ClientTop       =   3075
   ClientWidth     =   8745
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
   Icon            =   "frmImmTotalTests.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5895
   ScaleWidth      =   8745
   Begin VB.Frame Frame2 
      Caption         =   "Analysers"
      Height          =   1725
      Left            =   270
      TabIndex        =   21
      Top             =   2475
      Visible         =   0   'False
      Width           =   1725
      Begin Threed.SSOption SSOption1 
         Height          =   240
         Index           =   0
         Left            =   135
         TabIndex        =   22
         Top             =   315
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   423
         _StockProps     =   78
         Caption         =   "All"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   -1  'True
      End
      Begin Threed.SSOption SSOption1 
         Height          =   240
         Index           =   1
         Left            =   135
         TabIndex        =   23
         Top             =   630
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   423
         _StockProps     =   78
         Caption         =   "Image"
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
      Begin Threed.SSOption SSOption1 
         Height          =   240
         Index           =   2
         Left            =   135
         TabIndex        =   24
         Top             =   945
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   423
         _StockProps     =   78
         Caption         =   "Best"
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
      Begin Threed.SSOption SSOption1 
         Height          =   240
         Index           =   3
         Left            =   135
         TabIndex        =   25
         Top             =   1260
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   423
         _StockProps     =   78
         Caption         =   "Prospec"
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
   End
   Begin VB.Frame Frame1 
      Caption         =   "Between Dates"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1875
      Left            =   180
      TabIndex        =   10
      Top             =   90
      Width           =   5175
      Begin VB.CommandButton brecalc 
         Caption         =   "&Start"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   3630
         MaskColor       =   &H8000000F&
         Picture         =   "frmImmTotalTests.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   300
         Visible         =   0   'False
         Width           =   1275
      End
      Begin Threed.SSOption obetween 
         Height          =   225
         Index           =   6
         Left            =   720
         TabIndex        =   12
         Top             =   870
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
         Left            =   300
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1470
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
         Left            =   1530
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1470
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
      End
      Begin Threed.SSOption obetween 
         Height          =   225
         Index           =   2
         Left            =   1530
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   1170
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
      End
      Begin Threed.SSOption obetween 
         Height          =   225
         Index           =   5
         Left            =   3300
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1470
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
         Left            =   390
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1170
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
         Left            =   3450
         TabIndex        =   18
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
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   285
         Left            =   300
         TabIndex        =   19
         Top             =   300
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   503
         _Version        =   393216
         Format          =   59834369
         CurrentDate     =   37019
      End
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   315
         Left            =   2010
         TabIndex        =   20
         Top             =   300
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         _Version        =   393216
         Format          =   59834369
         CurrentDate     =   37019
      End
   End
   Begin ComctlLib.ProgressBar pb 
      Height          =   135
      Left            =   240
      TabIndex        =   9
      Top             =   2100
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   238
      _Version        =   327682
      Appearance      =   1
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   5505
      Left            =   5430
      TabIndex        =   8
      Top             =   180
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   9710
      _Version        =   393216
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
      ScrollBars      =   2
      FormatString    =   "<Parameter                 |<Tests    "
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
      Height          =   750
      Left            =   585
      Picture         =   "frmImmTotalTests.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4815
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
      Height          =   750
      Left            =   3780
      Picture         =   "frmImmTotalTests.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4815
      Width           =   1245
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Tests/Sample"
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
      Left            =   2205
      TabIndex        =   7
      Top             =   3330
      Width           =   990
   End
   Begin VB.Label ltps 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   3270
      TabIndex        =   6
      Top             =   3270
      Width           =   1245
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Total Samples"
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
      Left            =   2190
      TabIndex        =   5
      Top             =   3000
      Width           =   1005
   End
   Begin VB.Label lsamples 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3270
      TabIndex        =   4
      Top             =   2940
      Width           =   1245
   End
   Begin VB.Label ltotal 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3270
      TabIndex        =   3
      Top             =   2610
      Width           =   1245
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Total Tests"
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
      Left            =   2400
      TabIndex        =   2
      Top             =   2670
      Width           =   795
   End
End
Attribute VB_Name = "frmImmTotalTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit  '© Custom Software 2001

Private Sub bcancel_Click()

10        On Error GoTo bCancel_Click_Error


20        Unload Me

30        Exit Sub

bCancel_Click_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmImmTotalTests", "bcancel_Click", intEL, strES


End Sub

Private Sub bprint_Click()

          Dim pleft As Long
          Dim P As Long

10        On Error GoTo bprint_Click_Error

20        Printer.Print "Total Number of Tests"
30        Printer.Print "Between "; dtFrom; " and "; dtTo
40        Printer.Print
50        Printer.Print "Test Name"; Tab(25); "Number"; Tab(40); "Test Name"; Tab(65); "Number"
60        P = 0
70        pleft = True
80        Do While P <= g.Rows - 1
90            g.Row = P
100           g.Col = 0
110           Printer.Print g; Tab(IIf(pleft, 25, 65));
120           g.Col = 1
130           Printer.Print g; Tab(IIf(pleft, 40, 80));
140           pleft = Not pleft
150           If pleft Then Printer.Print
160           P = P + 1
170       Loop

180       Printer.Print
190       Printer.Print
200       Printer.Print "     Total Tests: "; ltotal
210       Printer.Print "   Total Samples: "; lsamples
220       Printer.Print "Tests per Sample: "; ltps

230       Printer.EndDoc

240       Exit Sub

bprint_Click_Error:

          Dim strES As String
          Dim intEL As Integer



250       intEL = Erl
260       strES = Err.Description
270       LogError "frmImmTotalTests", "bPrint_Click", intEL, strES


End Sub

Private Sub brecalc_Click()

10        On Error GoTo brecalc_Click_Error

20        bReCalc.Visible = False
30        DoEvents

40        FillG

50        Exit Sub

brecalc_Click_Error:

          Dim strES As String
          Dim intEL As Integer



60        intEL = Erl
70        strES = Err.Description
80        LogError "frmImmTotalTests", "brecalc_Click", intEL, strES


End Sub

Private Sub dtFrom_CloseUp()

10        On Error GoTo dtFrom_CloseUp_Error

20        bReCalc.Visible = True

30        Exit Sub

dtFrom_CloseUp_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmImmTotalTests", "dtFrom_CloseUp", intEL, strES


End Sub

Private Sub dtTo_CloseUp()

10        On Error GoTo dtTo_CloseUp_Error

20        bReCalc.Visible = True

30        Exit Sub

dtTo_CloseUp_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmImmTotalTests", "dtTo_CloseUp", intEL, strES


End Sub

Sub FillG()

          Dim sn As New Recordset
          Dim sql As String
          Dim n As Long
          Dim s As String
          Dim Total As Long
          Dim rs As Recordset


10        On Error GoTo FillG_Error

20        ReDim parameterCount(0 To 999, 0 To 1) As String
          '0 serum, 1 urine
          Dim parcnt As Long

30        pb.Visible = True

40        g.ColWidth(0) = 2040

50        ClearFGrid g
60        sql = "SELECT distinct(code) from Immtestdefinitions WHERE inuse = 1"
70        Set rs = New Recordset
80        RecOpenServer 0, rs, sql
90        Do While Not rs.EOF
100           If SSOption1(2) Then
110               sql = "SELECT sampleid, code, sampletype from Immresults WHERE " & _
                        "analyser = '2' and code = '" & rs!Code & "' and rundtime between '" & Format(dtFrom & " 00:00:00", "dd/mmm/yyyy hh:mm:ss") & "' " & _
                        "and '" & Format(dtTo & " 23:59:59", "dd/mmm/yyyy hh:mm:ss") & "'"
120           ElseIf SSOption1(1) Then
130               sql = "SELECT sampleid, code, sampletype from Immresults WHERE " & _
                        "analyser = '1' and code = '" & rs!Code & "' and runtime between '" & Format(dtFrom & " 00:00:00", "dd/mmm/yyyy hh:mm:ss") & "' " & _
                        "and '" & Format(dtTo & " 23:59:59", "dd/mmm/yyyy hh:mm:ss") & "'"
140           ElseIf SSOption1(3) Then
150               sql = "SELECT sampleid, code, sampletype from Immresults WHERE " & _
                        "analyser = '3' and code = '" & rs!Code & "' and runtime between '" & Format(dtFrom & " 00:00:00", "dd/mmm/yyyy hh:mm:ss") & "' " & _
                        "and '" & Format(dtTo & " 23:59:59", "dd/mmm/yyyy hh:mm:ss") & "'"
160           Else
170               sql = "SELECT count(code) as tot " & _
                        "FROM IMmresults WHERE code = '" & rs!Code & "' and " & _
                        "runtime between '" & Format(dtFrom & " 00:00:00", "dd/mmm/yyyy hh:mm:ss") & "' " & _
                        "and '" & Format(dtTo & " 23:59:59", "dd/mmm/yyyy hh:mm:ss") & "'"
180           End If
190           Set sn = New Recordset
200           RecOpenServer 0, sn, sql

210           If sn!Tot > 0 Then
220               s = ImmLongNameFor(rs!Code) & vbTab & sn!Tot
230               g.AddItem s
240           End If
250           rs.MoveNext
260       Loop



270       sql = "SELECT DISTINCT sampleid, rundate " & _
                "FROM immresults WHERE " & _
                "rundate between '" & Format(dtFrom & " 00:00:00", "dd/mmm/yyyy hh:mm:ss") & "' " & _
                "and '" & Format(dtTo & " 23:59:59", "dd/mmm/yyyy hh:mm:ss") & "'"
280       Set sn = New Recordset
290       RecOpenClient 0, sn, sql
300       If Not sn.EOF Then
310           sn.MoveLast
320           lsamples = Format(sn.RecordCount)
330       End If

340       g.Col = 1
350       Total = 0
360       For n = 1 To g.Rows - 1
370           g.Row = n
380           Total = Total + Val(g)
390       Next
400       ltotal = Format(Total)

410       If Val(lsamples) <> 0 Then
420           ltps = Format(Val(ltotal) / Val(lsamples), ".00")
430       Else
440           ltps = "0"
450       End If

460       pb.Visible = False



470       FixG g




480       Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer



490       intEL = Erl
500       strES = Err.Description
510       LogError "frmImmTotalTests", "FillG", intEL, strES, sql


End Sub

Private Sub Form_Load()

          Dim A2 As String

10        On Error GoTo Form_Load_Error

20        dtFrom = Format(Now, "dd/mmm/yyyy")
30        dtTo = dtFrom

40        SSOption1(1).Caption = GetOptionSetting("ImmAn1", "")
50        A2 = GetOptionSetting("ImmAn2", "")
60        If A2 <> "" Then
70            SSOption1(2).Caption = A2
80        Else
90            SSOption1(2).Visible = False
100       End If

110       SSOption1(0).Value = True
120       FillG

130       Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer



140       intEL = Erl
150       strES = Err.Description
160       LogError "frmImmTotalTests", "Form_Load", intEL, strES


End Sub

Private Sub oBetween_Click(Index As Integer, Value As Integer)

          Dim upto As String

10        On Error GoTo oBetween_Click_Error

20        dtFrom = BetweenDates(Index, upto)
30        dtTo = upto

40        FillG

50        Exit Sub

oBetween_Click_Error:

          Dim strES As String
          Dim intEL As Integer



60        intEL = Erl
70        strES = Err.Description
80        LogError "frmImmTotalTests", "oBetween_Click", intEL, strES


End Sub

Private Sub SSOption1_Click(Index As Integer, Value As Integer)
10        On Error GoTo SSOption1_Click_Error

20        bReCalc.Visible = True

30        Exit Sub

SSOption1_Click_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmImmTotalTests", "SSOption1_Click", intEL, strES

End Sub


