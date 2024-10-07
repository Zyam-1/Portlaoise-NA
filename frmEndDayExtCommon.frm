VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.ocx"
Begin VB.Form frmEndDayExtCommon 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NetAcquire - External End Of Day Summary"
   ClientHeight    =   8805
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8805
   ScaleWidth      =   11520
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExcel 
      Caption         =   "Export to Excel"
      Height          =   960
      Left            =   5520
      Picture         =   "frmEndDayExtCommon.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   240
      Width           =   1245
   End
   Begin VB.CommandButton breport 
      Appearance      =   0  'Flat
      Caption         =   "Print &Report"
      Height          =   960
      Left            =   3690
      Picture         =   "frmEndDayExtCommon.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   1245
   End
   Begin VB.CommandButton bcancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   960
      Left            =   10140
      Picture         =   "frmEndDayExtCommon.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   1245
   End
   Begin VB.CommandButton cmdStart 
      Appearance      =   0  'Flat
      Caption         =   "&Search"
      Height          =   960
      Left            =   2280
      Picture         =   "frmEndDayExtCommon.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   1245
   End
   Begin MSComCtl2.DTPicker dtRunDate 
      Height          =   405
      Left            =   240
      TabIndex        =   3
      Top             =   555
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
      Format          =   59572227
      CurrentDate     =   36966
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   7035
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   11205
      _ExtentX        =   19764
      _ExtentY        =   12409
      _Version        =   393216
      Cols            =   7
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
      FormatString    =   $"frmEndDayExtCommon.frx":0C28
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
   Begin ComctlLib.ProgressBar pb 
      Height          =   165
      Left            =   4080
      TabIndex        =   5
      Top             =   8475
      Width           =   7365
      _ExtentX        =   12991
      _ExtentY        =   291
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   1
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
      Left            =   6900
      TabIndex        =   9
      Top             =   600
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Label lblTotalRecords 
      AutoSize        =   -1  'True
      Caption         =   "Total Records"
      Height          =   195
      Left            =   180
      TabIndex        =   7
      Top             =   8460
      Width           =   1005
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Date Requested"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   6
      Top             =   300
      Width           =   1710
   End
End
Attribute VB_Name = "frmEndDayExtCommon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub bcancel_Click()

10        Unload Me

End Sub




Private Sub breport_Click()
          Dim Y As Long
          Dim X As Long
          Dim sql As String
          Dim sn As New Recordset

10        pb.Visible = True
20        pb.Max = g.Rows - 1

30        Printer.Orientation = vbPRORLandscape
40        Printer.Font.Name = "Courier New"
50        PrintText FormatString("External End of Day report for " & Format(dtRunDate, "dd/mmm/yyyy"), 99, , AlignCenter), 10, True, , , , True
60        PrintText String(107, "-"), , , , , , True



70        For Y = 0 To g.Rows - 1
80            pb = Y

90            PrintText FormatString(g.TextMatrix(Y, 0), 10, "|"), 9, IIf(Y = 0, True, False)   'sample id
100           PrintText FormatString(g.TextMatrix(Y, 1), 20, "|"), 9, IIf(Y = 0, True, False)     'patient name
110           PrintText FormatString(g.TextMatrix(Y, 2), 35, "|"), 9, IIf(Y = 0, True, False)     'test name
120           PrintText FormatString(g.TextMatrix(Y, 3), 20, "|"), 9, IIf(Y = 0, True, False)      'gp
130           PrintText FormatString(g.TextMatrix(Y, 4), 18), 9, IIf(Y = 0, True, False), , , , True  'ward
              'PrintText FormatString(g.TextMatrix(Y, 5), 25), 9, IIf(Y = 0, True, False)     'result
              'PrintText FormatString(g.TextMatrix(Y, 6), 10), 9, IIf(Y = 0, True, False), , , , True    'return date

140           If Y = 0 Then PrintText String(107, "-"), , , , , , True
150       Next

160       Printer.EndDoc

170       pb.Visible = False
End Sub

Private Sub cmdExcel_Click()

          Dim s As String

10        On Error GoTo cmdExcel_Click_Error

20        s = "External End of Day report for " & Format(dtRunDate, "dd/mmm/yyyy") & vbCr & vbCr
30        ExportFlexGrid g, Me, s

40        Exit Sub

cmdExcel_Click_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmEndDayExtCommon", "cmdExcel_Click", intEL, strES

End Sub

Private Sub cmdStart_Click()

10        FillG

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
30        lblTotalRecords.Caption = ""

40        sql = "SELECT D.SampleID, D.PatName, D.Chart, D.GP, D.Ward, D.Clinician, " & _
                "R.Analyte, R.Result, R.SentDate, R.RetDate, R.SendTo " & _
                "FROM Demographics D " & _
                "INNER JOIN ExtResults R ON D.SampleID = R.SampleID " & _
                "WHERE DATEDIFF(hh, R.SentDate, '" & Format$(dtRunDate, "dd/mmm/yyyy 23:59:59") & "') BETWEEN 1 AND 23 " & _
                "ORDER BY D.SampleID"

50        Set tb = New Recordset
60        RecOpenClient 0, tb, sql

70        pb.Max = tb.RecordCount + 1
80        pb = 0
90        pb.Visible = True
100       g.Visible = False

110       With tb
120           Do While Not .EOF
130               pb = pb + 1
140               s = tb!SampleID & "" & vbTab & _
                      tb!PatName & "" & vbTab & _
                      tb!Analyte & "" & vbTab & _
                      tb!GP & "" & vbTab & _
                      tb!Ward & "" & vbTab & _
                      tb!Result & "" & vbTab & _
                      tb!RetDate & ""
150               g.AddItem s
160               .MoveNext
170           Loop
180       End With

190       FixG g
200       If g.Rows = 2 And g.TextMatrix(1, 0) = "" Then
210           lblTotalRecords = "Total records found: 0"
220       Else
230           lblTotalRecords = "Total records found: " & g.Rows - 1
240       End If
250       pb.Visible = False

260       Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer



270       intEL = Erl
280       strES = Err.Description
290       LogError "frmEndDayEndCommon", "FillG", intEL, strES, sql


End Sub

Private Sub Form_Load()

10        On Error GoTo Form_Load_Error

20        dtRunDate = Format(Now, "dd/mmm/yyyy")
30        lblTotalRecords.Caption = ""
40        pb.Visible = False

          'FillG

50        Set_Font Me

60        Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer



70        intEL = Erl
80        strES = Err.Description
90        LogError "frmEndDayEndCommon", "Form_Load", intEL, strES


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
90            Exit Sub
100       End If

          'If g.Row = g.RowSel Then
          '    breport.Visible = True
          'Else
          '    breport.Visible = False
          'End If




110       Exit Sub

g_Click_Error:

          Dim strES As String
          Dim intEL As Integer



120       intEL = Erl
130       strES = Err.Description
140       LogError "frmEndDayEndCommon", "g_Click", intEL, strES


End Sub


