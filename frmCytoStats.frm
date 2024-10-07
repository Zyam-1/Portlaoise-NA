VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frmCytoStats 
   Caption         =   "NetAcquire 6 - Cytology - Totals"
   ClientHeight    =   6840
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7335
   Icon            =   "frmCytoStats.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   7335
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExcel 
      Caption         =   "Export to Excel"
      Height          =   915
      Left            =   135
      Picture         =   "frmCytoStats.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3870
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   6450
      Left            =   3465
      TabIndex        =   14
      Top             =   225
      Width           =   3750
      _ExtentX        =   6615
      _ExtentY        =   11377
      _Version        =   393216
      FormatString    =   "Source                                    |Tests              "
   End
   Begin VB.CommandButton bCancel 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   930
      Left            =   2025
      Picture         =   "frmCytoStats.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2835
      Width           =   1380
   End
   Begin VB.CommandButton bStart 
      Caption         =   "&Start"
      Height          =   930
      Left            =   135
      Picture         =   "frmCytoStats.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2835
      Width           =   1425
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   645
      Left            =   135
      TabIndex        =   2
      Top             =   2070
      Width           =   3255
      _Version        =   65536
      _ExtentX        =   5741
      _ExtentY        =   1138
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
         TabIndex        =   3
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
         TabIndex        =   4
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
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   975
      Left            =   135
      TabIndex        =   5
      Top             =   180
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
         TabIndex        =   6
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
         TabIndex        =   7
         Top             =   315
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   661
         _Version        =   393216
         Format          =   59441153
         CurrentDate     =   37643
      End
   End
   Begin Threed.SSPanel SSPanel4 
      Height          =   870
      Left            =   135
      TabIndex        =   9
      Top             =   1170
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
         TabIndex        =   10
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
         TabIndex        =   11
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
         TabIndex        =   12
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
         TabIndex        =   13
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
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   3405
      TabIndex        =   1
      Top             =   2295
      Visible         =   0   'False
      Width           =   1245
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
      Left            =   120
      TabIndex        =   16
      Top             =   4770
      Visible         =   0   'False
      Width           =   1455
   End
End
Attribute VB_Name = "frmCytoStats"
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

20        Hosp = Switch(h(0), "Tullamore", h(1), "Portlaoise", h(2), "Mullingar", h(3), "O")

30        g.Rows = 2
40        g.AddItem ""
50        g.RemoveItem 1

60        For n = 0 To List1.ListCount - 1
70            sql = "SELECT count(demographics.sampleid) as tot " & _
                    "from demographics, cytoresults WHERE " & _
                    "demographics.rundate between '" & _
                    Format(calFrom, "dd/mmm/yyyy") & " 00:00:00' and '" & _
                    Format(calTo, "dd/mmm/yyyy") & " 23:59:59' " & _
                    "and demographics.hospital = '" & Hosp & "' and " & _
                    IIf(s(0), "demographics.ward  = '", "clinician = '") & _
                    AddTicks(List1.List(n)) & "' and cytoresults.sampleid = demographics.sampleid"
80            Set sn = New Recordset
90            RecOpenClient 0, sn, sql
100           If sn!Tot <> 0 Then
110               g.AddItem List1.List(n) & vbTab & sn!Tot
120               g.Refresh
130           End If
140       Next

150       If g.Rows > 1 Then
160           g.Col = 1
170           For n = 1 To g.Rows - 1
180               g.Row = n
190               Total = Total + Val(g)
200           Next
210           g.AddItem ""
220           g.AddItem "Total Above" & vbTab & Format(Total)
230       End If

240       sql = "SELECT count(demographics.sampleid) as tot " & _
                "from demographics, cytoresults WHERE " & _
                "demographics.rundate between '" & _
                Format(calFrom, "dd/mmm/yyyy") & " 00:00:00' and '" & _
                Format(calTo, "dd/mmm/yyyy") & " 23:59:59' " & _
                "and demographics.hospital = '" & Hosp & "' and " & _
                "cytoresults.sampleid = demographics.sampleid"
250       Set sn = New Recordset
260       RecOpenClient 0, sn, sql

270       If sn!Tot <> 0 Then
280           g.AddItem ""
290           g.AddItem "Total Records" & vbTab & Format(sn!Tot)
300       End If

310       If g.Rows > 2 And g.TextMatrix(1, 0) = "" Then
320           g.RemoveItem 1
330       End If

340       Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer

350       intEL = Erl
360       strES = Err.Description
370       LogError "frmCytoStats", "FillG", intEL, strES, sql

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
60            Set tb = New Recordset
70            RecOpenServer 0, tb, sql
80            Do While Not tb.EOF
90                List1.AddItem Trim(tb!Text)
100               tb.MoveNext
110           Loop
120       Else
130           sql = "SELECT * from clinicians WHERE hospitalcode = '" & Hosp & "'"
140           Set tb = New Recordset
150           RecOpenServer 0, tb, sql
160           Do While Not tb.EOF
170               List1.AddItem Trim(tb!Text)
180               tb.MoveNext
190           Loop
200       End If

210       Exit Sub

FillList1_Error:

          Dim strES As String
          Dim intEL As Integer

220       intEL = Erl
230       strES = Err.Description
240       LogError "frmCytoStats", "FillList1", intEL, strES, sql

End Sub

Private Sub bcancel_Click()

10        Unload Me

End Sub


Private Sub bStart_Click()

10        FillG

End Sub


Private Sub cmdExcel_Click()

10        ExportFlexGrid g, Me

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
70        LogError "frmCytoStats", "Form_Activate", intEL, strES

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
70        LogError "frmCytoStats", "Form_Load", intEL, strES

End Sub


Private Sub h_Click(Index As Integer, Value As Integer)

10        FillList1

End Sub


Private Sub s_Click(Index As Integer, Value As Integer)

10        FillList1

End Sub




