VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.ocx"
Begin VB.Form frmEndTotalTests 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire 6 - Endocrinology Total Tests"
   ClientHeight    =   6060
   ClientLeft      =   2010
   ClientTop       =   3075
   ClientWidth     =   8565
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
   Icon            =   "frmEndTotalTests.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6060
   ScaleWidth      =   8565
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
      Left            =   1035
      Picture         =   "frmEndTotalTests.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4860
      Width           =   1305
   End
   Begin VB.ListBox lstTest 
      Height          =   5325
      Left            =   8595
      TabIndex        =   15
      Top             =   180
      Width           =   870
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
      TabIndex        =   4
      Top             =   90
      Width           =   5175
      Begin VB.CommandButton brecalc 
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
         Height          =   795
         Left            =   3630
         MaskColor       =   &H8000000F&
         Picture         =   "frmEndTotalTests.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   300
         Visible         =   0   'False
         Width           =   1275
      End
      Begin Threed.SSOption obetween 
         Height          =   225
         Index           =   6
         Left            =   720
         TabIndex        =   6
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
         TabIndex        =   7
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
         TabIndex        =   8
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
         TabIndex        =   9
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
         TabIndex        =   10
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
         TabIndex        =   11
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
         TabIndex        =   12
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
         TabIndex        =   13
         Top             =   300
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   503
         _Version        =   393216
         Format          =   165937153
         CurrentDate     =   37019
      End
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   315
         Left            =   2010
         TabIndex        =   14
         Top             =   300
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         _Version        =   393216
         Format          =   165937153
         CurrentDate     =   37019
      End
   End
   Begin ComctlLib.ProgressBar pb 
      Height          =   135
      Left            =   240
      TabIndex        =   3
      Top             =   2100
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   238
      _Version        =   327682
      Appearance      =   1
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   5775
      Left            =   5430
      TabIndex        =   2
      Top             =   180
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   10186
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
      Height          =   855
      Left            =   2385
      Picture         =   "frmEndTotalTests.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4860
      Width           =   1305
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
      Left            =   3735
      Picture         =   "frmEndTotalTests.frx":0C28
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4860
      Width           =   1305
   End
   Begin VB.Label lblExcelInfo 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exporting..."
      Height          =   345
      Left            =   1035
      TabIndex        =   17
      Top             =   4455
      Visible         =   0   'False
      Width           =   1275
   End
End
Attribute VB_Name = "frmEndTotalTests"
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
60        LogError "frmEndTotalTests", "bcancel_Click", intEL, strES


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
90            g.row = P
100           g.Col = 0
110           Printer.Print g; Tab(IIf(pleft, 25, 65));
120           g.Col = 1
130           Printer.Print g; Tab(IIf(pleft, 40, 80));
140           pleft = Not pleft
150           If pleft Then Printer.Print
160           P = P + 1
170       Loop

180       Printer.EndDoc


190       Exit Sub

bprint_Click_Error:

          Dim strES As String
          Dim intEL As Integer



200       intEL = Erl
210       strES = Err.Description
220       LogError "frmEndTotalTests", "bPrint_Click", intEL, strES


End Sub

Private Sub brecalc_Click()

10        On Error GoTo brecalc_Click_Error

20        brecalc.Visible = False
30        DoEvents

40        FillG

50        Exit Sub

brecalc_Click_Error:

          Dim strES As String
          Dim intEL As Integer



60        intEL = Erl
70        strES = Err.Description
80        LogError "frmEndTotalTests", "brecalc_Click", intEL, strES


End Sub

Private Sub cmdExcel_Click()
10     ExportFlexGrid g, Me
End Sub

Private Sub dtFrom_CloseUp()

10        On Error GoTo dtFrom_CloseUp_Error

20        brecalc.Visible = True

30        Exit Sub

dtFrom_CloseUp_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmEndTotalTests", "dtFrom_CloseUp", intEL, strES


End Sub

Private Sub dtTo_CloseUp()

10        On Error GoTo dtTo_CloseUp_Error

20        brecalc.Visible = True

30        Exit Sub

dtTo_CloseUp_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmEndTotalTests", "dtTo_CloseUp", intEL, strES


End Sub

Sub FillG()

          Dim tb As New Recordset
          Dim sql As String
          Dim n As Long
          Dim FromDate As String
          Dim ToDate As String
          Dim s As String
          Dim Tot As Long

10        On Error GoTo FillG_Error

20        FromDate = Format$(dtFrom, "dd/MMM/yyyy")
30        ToDate = Format$(dtTo, "dd/MMM/yyyy")

40        With g
50            .Rows = 2
60            .AddItem ""
70            .RemoveItem 1
80        End With

90        With pb
100           .Max = lstTest.ListCount
110           .Visible = True
120       End With

130       For n = 0 To lstTest.ListCount - 1
140           With lstTest
150               .Selected(n) = True
160               .Refresh
170               pb = .ListIndex
180           End With
190           sql = "SELECT count(distinct sampleid) as tot " & _
                    "from endresults as r " & _
                    "WHERE r.runtime between '" & FromDate & " 00:00:00' and '" & ToDate & " 23:59:59' " & _
                    "and R.code = '" & eCodeForShortName(lstTest.List(n)) & "' "
200           Set tb = New Recordset
210           RecOpenClient 0, tb, sql
220           If tb!Tot <> 0 Then
230               s = lstTest.List(n) & vbTab & _
                      Format$(tb!Tot)
240               g.AddItem s
250               g.Refresh
260           End If
270       Next
280       pb.Visible = False

290       If g.Rows > 2 Then
300           g.RemoveItem 1
310       End If

320       For n = 1 To g.Rows - 1
330           Tot = Tot + Val(g.TextMatrix(n, 1))
340       Next

350       g.AddItem ""
360       g.AddItem "Total" & vbTab & Tot




370       Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer



380       intEL = Erl
390       strES = Err.Description
400       LogError "frmEndTotalTests", "FillG", intEL, strES, sql


End Sub

'---------------------------------------------------------------------------------------
' Procedure : Form_Load
' DateTime  : 28/06/2007 11:03
' Author    : Myles
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Form_Load()


10        On Error GoTo Form_Load_Error

20        dtFrom = Format(Now, "dd/mmm/yyyy")
30        dtTo = dtFrom


40        FillList
          'FillG

50        Set_Font Me



60        Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer



70        intEL = Erl
80        strES = Err.Description
90        LogError "frmEndTotalTests", "Form_Load", intEL, strES


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
80        LogError "frmEndTotalTests", "oBetween_Click", intEL, strES


End Sub




Private Sub FillList()
          Dim tb As New Recordset
          Dim sql As String

10        On Error GoTo FillList_Error

20        lstTest.Clear

30        sql = "SELECT distinct ShortName, PrintPriority from endTestDefinitions " & _
                " WHERE inuse = 1 order by PrintPriority asc"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql

60        Do While Not tb.EOF
70            lstTest.AddItem tb!ShortName & ""
80            tb.MoveNext
90        Loop

100       Exit Sub

FillList_Error:

          Dim strES As String
          Dim intEL As Integer



110       intEL = Erl
120       strES = Err.Description
130       LogError "frmEndTotalTests", "FillList", intEL, strES, sql


End Sub

