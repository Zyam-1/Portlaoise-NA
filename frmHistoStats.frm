VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frmHistoStats 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Histology - Statistics"
   ClientHeight    =   6735
   ClientLeft      =   1605
   ClientTop       =   1320
   ClientWidth     =   7380
   Icon            =   "frmHistoStats.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdExcel 
      Caption         =   "Export to Excel"
      Height          =   780
      Left            =   180
      Picture         =   "frmHistoStats.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4815
      Width           =   1230
   End
   Begin VB.CommandButton bStart 
      Caption         =   "&Start"
      Height          =   885
      Left            =   180
      Picture         =   "frmHistoStats.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3720
      Width           =   1245
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   1110
      Left            =   180
      TabIndex        =   7
      Top             =   2025
      Width           =   3255
      _Version        =   65536
      _ExtentX        =   5741
      _ExtentY        =   1958
      _StockProps     =   15
      Caption         =   "Option"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelOuter      =   1
      BevelInner      =   2
      Alignment       =   0
      Begin Threed.SSOption s 
         Height          =   195
         Index           =   1
         Left            =   1575
         TabIndex        =   9
         Top             =   360
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "Clinicians"
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
      Begin Threed.SSOption s 
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   8
         Top             =   360
         Width           =   1140
         _Version        =   65536
         _ExtentX        =   2011
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "Wards"
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
      Begin Threed.SSOption s 
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   15
         Top             =   720
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "Gp"
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
   Begin Threed.SSPanel SSPanel1 
      Height          =   975
      Left            =   180
      TabIndex        =   1
      Top             =   135
      Width           =   3285
      _Version        =   65536
      _ExtentX        =   5794
      _ExtentY        =   1720
      _StockProps     =   15
      Caption         =   "Between Dates"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   2
      BorderWidth     =   2
      BevelInner      =   1
      Alignment       =   0
      Begin MSComCtl2.DTPicker calto 
         Height          =   375
         Left            =   1755
         TabIndex        =   13
         Top             =   315
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   59441153
         CurrentDate     =   37643
      End
      Begin MSComCtl2.DTPicker calfrom 
         Height          =   375
         Left            =   180
         TabIndex        =   12
         Top             =   315
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   661
         _Version        =   393216
         Format          =   59441153
         CurrentDate     =   37643
      End
   End
   Begin VB.CommandButton bCancel 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   885
      Left            =   2025
      Picture         =   "frmHistoStats.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3735
      Width           =   1245
   End
   Begin Threed.SSPanel SSPanel4 
      Height          =   870
      Left            =   180
      TabIndex        =   2
      Top             =   1125
      Width           =   3255
      _Version        =   65536
      _ExtentX        =   5741
      _ExtentY        =   1535
      _StockProps     =   15
      Caption         =   "Hospital"
      ForeColor       =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelOuter      =   1
      BevelInner      =   2
      Alignment       =   0
      Begin Threed.SSOption h 
         Height          =   195
         Index           =   1
         Left            =   1620
         TabIndex        =   3
         Top             =   315
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "Portlaoise"
         ForeColor       =   64
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
      Begin Threed.SSOption h 
         Height          =   195
         Index           =   3
         Left            =   1620
         TabIndex        =   4
         Top             =   585
         Width           =   705
         _Version        =   65536
         _ExtentX        =   1244
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "Other"
         ForeColor       =   0
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
      Begin Threed.SSOption h 
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   5
         Top             =   585
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "Mullingar"
         ForeColor       =   0
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
      Begin Threed.SSOption h 
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   6
         Top             =   315
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "Tullamore"
         ForeColor       =   0
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
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   6450
      Left            =   3510
      TabIndex        =   14
      Top             =   135
      Width           =   3750
      _ExtentX        =   6615
      _ExtentY        =   11377
      _Version        =   393216
      FormatString    =   "Source                                    |Tests              "
   End
   Begin VB.ListBox List1 
      Height          =   5520
      Left            =   7560
      TabIndex        =   10
      Top             =   360
      Visible         =   0   'False
      Width           =   3165
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
      Left            =   210
      TabIndex        =   17
      Top             =   5640
      Visible         =   0   'False
      Width           =   1125
   End
End
Attribute VB_Name = "frmHistoStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub FillG()

          Dim sn As Recordset
          Dim sql As String
          Dim Hosp As String
          Dim n As Long
          Dim Total As Long

10        On Error GoTo FillG_Error

20        Hosp = Switch(h(0), "Tullamore", h(1), "Portlaoise", h(2), "Mullingar", h(3), "Other")

30        g.Rows = 2
40        g.AddItem ""
50        g.RemoveItem 1


60        For n = 0 To List1.ListCount - 1
70            sql = "SELECT count(demographics.sampleid) as tot " & _
                    "from demographics,historesults WHERE " & _
                    "demographics.rundate between '" & _
                    Format(calFrom, "dd/mmm/yyyy") & " 00:00:00' and '" & _
                    Format(calTo, "dd/mmm/yyyy") & " 23:59:59' " & _
                    "and demographics.hospital = '" & Hosp & "' and "

80            If s(0) Then
90                sql = sql & "demographics.ward  = '" & AddTicks(List1.List(n)) & "' "
100           ElseIf s(1) Then
110               sql = sql & "demographics.clinician = '" & AddTicks(List1.List(n)) & "' "
120           ElseIf s(2) Then
130               sql = sql & "demographics.gp = '" & AddTicks(List1.List(n)) & "' "
140           End If

150           sql = sql & "and historesults.sampleid = demographics.sampleid"
160           Set sn = New Recordset
170           RecOpenServer 0, sn, sql
180           If sn!Tot <> 0 Then
190               g.AddItem List1.List(n) & vbTab & sn!Tot
200               g.Refresh
210           End If
220       Next

230       If g.Rows > 1 Then
240           g.Col = 1
250           For n = 1 To g.Rows - 1
260               g.Row = n
270               Total = Total + Val(g)
280           Next
290           g.AddItem ""
300           g.AddItem "Total Above" & vbTab & Format(Total)
310       End If

320       sql = "SELECT count(demographics.sampleid) as tot " & _
                "from demographics, historesults WHERE " & _
                "demographics.rundate between '" & _
                Format(calFrom, "dd/mmm/yyyy") & " 00:00:00' and '" & _
                Format(calTo, "dd/mmm/yyyy") & " 23:59:59' " & _
                "and demographics.hospital = '" & Hosp & "' and  " & _
                "historesults.sampleid = demographics.sampleid"
330       Set sn = New Recordset
340       RecOpenClient 0, sn, sql

350       If sn!Tot <> 0 Then
360           g.AddItem ""
370           g.AddItem "Total Records" & vbTab & Format(sn!Tot)
380       End If

390       If g.Rows > 2 And g.TextMatrix(1, 0) = "" Then
400           g.RemoveItem 1
410       End If




420       Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer



430       intEL = Erl
440       strES = Err.Description
450       LogError "frmHistoStats", "FillG", intEL, strES, sql


End Sub

Sub FillList1()

          Dim Hosp As String
          Dim sql As String
          Dim tb As Recordset

10        On Error GoTo FillList1_Error

20        Hosp = Switch(h(0), "T", h(1), "P", h(2), "M", h(3), "O")

30        List1.Clear

40        If s(0) Then
50            sql = "SELECT * from wards WHERE hospitalcode = '" & Hosp & "'"
60        ElseIf s(1) Then
70            sql = "SELECT * from clinicians WHERE hospitalcode = '" & Hosp & "'"
80        ElseIf s(2) Then
90            sql = "SELECT * from gps WHERE hospitalcode = '" & Hosp & "'"
100       End If

110       Set tb = New Recordset
120       RecOpenServer 0, tb, sql
130       Do While Not tb.EOF
140           List1.AddItem Trim(tb!Text)
150           tb.MoveNext
160       Loop

170       Exit Sub

FillList1_Error:

          Dim strES As String
          Dim intEL As Integer



180       intEL = Erl
190       strES = Err.Description
200       LogError "frmHistoStats", "FillList1", intEL, strES, sql


End Sub

Private Sub bcancel_Click()

10        Unload Me

End Sub


Private Sub bStart_Click()

10        On Error GoTo bStart_Click_Error

20        FillG

30        Exit Sub

bStart_Click_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmHistoStats", "bStart_Click", intEL, strES


End Sub


Private Sub cmdExcel_Click()

10        On Error GoTo cmdExcel_Click_Error

20        ExportFlexGrid g, Me

30        Exit Sub

cmdExcel_Click_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmHistoStats", "cmdExcel_Click", intEL, strES


End Sub

Private Sub Form_Activate()

10        On Error GoTo Form_Activate_Error

20        g.ColWidth(0) = 2000

30        FillList1

40        Exit Sub

Form_Activate_Error:

          Dim strES As String
          Dim intEL As Integer



50        intEL = Erl
60        strES = Err.Description
70        LogError "frmHistoStats", "Form_Activate", intEL, strES


End Sub

Private Sub Form_Load()

10        On Error GoTo Form_Load_Error

20        calTo = Format(Now, "dd/mmm/yyyy")
30        calFrom = Format(Now - 7, "dd/mmm/yyyy")

40        Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer



50        intEL = Erl
60        strES = Err.Description
70        LogError "frmHistoStats", "Form_Load", intEL, strES


End Sub


Private Sub h_Click(Index As Integer, Value As Integer)

10        On Error GoTo h_Click_Error

20        FillList1

30        Exit Sub

h_Click_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmHistoStats", "h_Click", intEL, strES


End Sub


Private Sub s_Click(Index As Integer, Value As Integer)

10        On Error GoTo s_Click_Error

20        FillList1

30        Exit Sub

s_Click_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmHistoStats", "s_Click", intEL, strES


End Sub


