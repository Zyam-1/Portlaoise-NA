VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmCoagTotalTests 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire 6 - Coagulation Total Tests"
   ClientHeight    =   5685
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
   Icon            =   "frmCoagTotalTests.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5685
   ScaleWidth      =   8745
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   900
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   8745
      _ExtentX        =   15425
      _ExtentY        =   1588
      ButtonWidth     =   1032
      ButtonHeight    =   1376
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   5
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Start"
            Key             =   ""
            Object.ToolTipText     =   "Recalculate"
            Object.Tag             =   "Rec"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Print"
            Key             =   ""
            Object.ToolTipText     =   "Print"
            Object.Tag             =   "Print"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Exit"
            Key             =   ""
            Object.ToolTipText     =   "Exit/Cancel"
            Object.Tag             =   "Exit"
            ImageIndex      =   2
         EndProperty
      EndProperty
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
      Left            =   150
      TabIndex        =   8
      Top             =   1215
      Width           =   5175
      Begin Threed.SSOption obetween 
         Height          =   225
         Index           =   6
         Left            =   720
         TabIndex        =   9
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
         TabIndex        =   10
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
         TabIndex        =   11
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
         TabIndex        =   12
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
         TabIndex        =   13
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
         TabIndex        =   14
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
         TabIndex        =   15
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
         TabIndex        =   16
         Top             =   300
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   503
         _Version        =   393216
         Format          =   272236545
         CurrentDate     =   37019
      End
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   315
         Left            =   2010
         TabIndex        =   17
         Top             =   300
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         _Version        =   393216
         Format          =   272236545
         CurrentDate     =   37019
      End
   End
   Begin ComctlLib.ProgressBar pb 
      Height          =   135
      Left            =   210
      TabIndex        =   7
      Top             =   3225
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   238
      _Version        =   327682
      Appearance      =   1
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   4455
      Left            =   5370
      TabIndex        =   6
      Top             =   1155
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   7858
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
   Begin ComctlLib.ImageList ImageList1 
      Left            =   4380
      Top             =   -120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   4
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCoagTotalTests.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCoagTotalTests.frx":0624
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCoagTotalTests.frx":093E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCoagTotalTests.frx":0C58
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Tests/Sample"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   960
      TabIndex        =   5
      Top             =   4965
      Width           =   1455
   End
   Begin VB.Label ltps 
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
      Height          =   405
      Left            =   2430
      TabIndex        =   4
      Top             =   4935
      Width           =   1605
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Total Samples"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   870
      TabIndex        =   3
      Top             =   4515
      Width           =   1515
   End
   Begin VB.Label lsamples 
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
      Height          =   405
      Left            =   2430
      TabIndex        =   2
      Top             =   4455
      Width           =   1605
   End
   Begin VB.Label ltotal 
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
      Height          =   465
      Left            =   2430
      TabIndex        =   1
      Top             =   3915
      Width           =   1605
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Total Tests"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1230
      TabIndex        =   0
      Top             =   4005
      Width           =   1170
   End
End
Attribute VB_Name = "frmCoagTotalTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit  '© Custom Software 2001

Private Sub Cancel_Click()

10        Unload Me

End Sub

Private Sub dtFrom_CloseUp()

10        On Error GoTo dtFrom_CloseUp_Error

20        Toolbar1.Buttons(1).Enabled = True

30        Exit Sub

dtFrom_CloseUp_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmCoagTotalTests", "dtFrom_CloseUp", intEL, strES


End Sub

Private Sub dtTo_CloseUp()

10        On Error GoTo dtTo_CloseUp_Error

20        Toolbar1.Buttons(1).Enabled = True

30        Exit Sub

dtTo_CloseUp_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmCoagTotalTests", "dtTo_CloseUp", intEL, strES


End Sub

Sub FillG()

          Dim sn As New Recordset
          Dim sql As String
          Dim n As Long
          Dim s As String
          Dim Total As Long
          Dim pCount As Long
          Dim RecCount As Long

10        On Error GoTo FillG_Error

20        ReDim parameterCount(0 To 9999, 0 To 1) As Long
          '0 serum, 1 urine
          Dim parcnt As Long

30        pb.Visible = True

40        g.ColWidth(0) = 2040

50        ClearFGrid g

60        sql = "SELECT distinct  sampleid, code " & _
                "FROM coagresults WHERE " & _
                "rundate between '" & Format(dtFrom & " 00:00:00", "dd/mmm/yyyy hh:mm:ss") & "' " & _
                "and '" & Format(dtTo & " 23:59:59", "dd/mmm/yyyy hh:mm:ss") & "'"
70        Set sn = New Recordset
80        RecOpenServer 0, sn, sql
90        If sn.EOF Then
100           pb.Visible = False
110           g.Visible = True
120           Exit Sub
130       End If
140       pb.Max = 100
150       parcnt = 0
160       Do While Not sn.EOF
170           n = n + 1
180           sn.MoveNext
190       Loop
200       sn.MoveFirst
210       pCount = n / 100
220       RecCount = 0
230       Do While Not sn.EOF
240           If RecCount = Int(pCount) Then
250               If pb < pb.Max Then pb = pb + 1
260               RecCount = 0
270           Else
280               RecCount = RecCount + 1
290           End If
300           parameterCount(Val(sn!Code), 1) = parameterCount(Val(sn!Code), 1) + 1
310           parcnt = parcnt + 1
320           sn.MoveNext
330       Loop

340       pb.Max = 9999
350       For n = 1 To 9999
360           pb = n
370           If parameterCount(n, 0) <> 0 Then
380               s = CoagNameFor(n) & vbTab
390               s = s & parameterCount(n, 0)
400               g.AddItem s
410               g.Refresh
420           End If
430           If parameterCount(n, 1) <> 0 Then
440               s = CoagNameFor(n) & vbTab
450               s = s & parameterCount(n, 1)
460               g.AddItem s
470               g.Refresh
480           End If
490       Next

500       sql = "SELECT count(DISTINCT sampleid) as tot " & _
                "FROM coagresults WHERE " & _
                "rundate between '" & Format(dtFrom & " 00:00:00", "dd/mmm/yyyy hh:mm:ss") & "' " & _
                "and '" & Format(dtTo & " 23:59:59", "dd/mmm/yyyy hh:mm:ss") & "'"
510       Set sn = New Recordset
520       RecOpenServer 0, sn, sql
530       If Not sn.EOF Then
540           lsamples = sn!Tot
550       End If

560       g.Col = 1
570       Total = 0
580       For n = 1 To g.Rows - 1
590           If IsNumeric(g.TextMatrix(n, 0)) Then
                  '    g.RemoveItem n
600           Else
610               g.Row = n
620               Total = Total + Val(g)
630           End If
640       Next
650       ltotal = Format(Total)

660       For n = g.Rows - 1 To 1 Step -1
670           If IsNumeric(g.TextMatrix(n, 0)) Then
680               g.RemoveItem n
690           End If
700       Next

710       If Val(lsamples) <> 0 Then
720           ltps = Format(Val(ltotal) / Val(lsamples), ".00")
730       Else
740           ltps = "0"
750       End If

760       pb.Visible = False

770       FixG g

780       Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer

790       intEL = Erl
800       strES = Err.Description
810       LogError "frmCoagTotalTests", "FillG", intEL, strES

End Sub

Private Sub Form_Load()

10        On Error GoTo Form_Load_Error

20        dtFrom = Format(Now, "dd/mmm/yyyy")
30        dtTo = dtFrom
          'FillG
40        Set_Font Me

50        Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

60        intEL = Erl
70        strES = Err.Description
80        LogError "frmCoagTotalTests", "Form_Load", intEL, strES


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
80        LogError "frmCoagTotalTests", "oBetween_Click", intEL, strES


End Sub

Private Sub Print_Click()

          Dim pleft As Long
          Dim P As Long


10        On Error GoTo Print_Click_Error

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

Print_Click_Error:

          Dim strES As String
          Dim intEL As Integer

250       intEL = Erl
260       strES = Err.Description
270       LogError "frmCoagTotalTests", "Print_Click", intEL, strES


End Sub

Private Sub ReCalc_Click()

10        On Error GoTo ReCalc_Click_Error

20        Toolbar1.Buttons(1).Enabled = False
30        DoEvents

40        FillG

50        Exit Sub

ReCalc_Click_Error:

          Dim strES As String
          Dim intEL As Integer

60        intEL = Erl
70        strES = Err.Description
80        LogError "frmCoagTotalTests", "ReCalc_Click", intEL, strES


End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)

10        On Error GoTo Toolbar1_ButtonClick_Error

20        If Button.Tag = "Exit" Then
30            Cancel_Click
40        ElseIf Button.Tag = "Print" Then
50            Print_Click
60        Else
70            ReCalc_Click
80        End If

90        Exit Sub

Toolbar1_ButtonClick_Error:

          Dim strES As String
          Dim intEL As Integer

100       intEL = Erl
110       strES = Err.Description
120       LogError "frmCoagTotalTests", "Toolbar1_ButtonClick", intEL, strES

End Sub
