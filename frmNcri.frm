VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmNcri 
   Caption         =   "NetAcquire - NCRI Report"
   ClientHeight    =   7620
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   14715
   Icon            =   "frmNcri.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7620
   ScaleWidth      =   14715
   Begin VB.Frame fraReport 
      Caption         =   "Report"
      Height          =   7530
      Left            =   60
      TabIndex        =   8
      Top             =   30
      Visible         =   0   'False
      Width           =   14640
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Height          =   1005
         Left            =   13590
         TabIndex        =   10
         ToolTipText     =   "Close Report"
         Top             =   315
         Width           =   870
      End
      Begin VB.TextBox txtReport 
         Height          =   7125
         Left            =   90
         MultiLine       =   -1  'True
         TabIndex        =   9
         ToolTipText     =   "Histology Re"
         Top             =   225
         Width           =   13335
      End
   End
   Begin VB.CommandButton cmdExcel 
      Caption         =   "Export to Excel"
      Height          =   870
      Left            =   12060
      Picture         =   "frmNcri.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   315
      Width           =   1230
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   870
      Left            =   13410
      Picture         =   "frmNcri.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   315
      Width           =   1050
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   690
      Left            =   4455
      Picture         =   "frmNcri.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   45
      Width           =   1050
   End
   Begin MSComCtl2.DTPicker dtTo 
      Height          =   330
      Left            =   2970
      TabIndex        =   1
      Top             =   360
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "dd/MMM/yyyy"
      Format          =   59768835
      CurrentDate     =   38979
   End
   Begin MSFlexGridLib.MSFlexGrid grdNcri 
      Height          =   5745
      Left            =   45
      TabIndex        =   0
      Top             =   1710
      Width           =   14610
      _ExtentX        =   25770
      _ExtentY        =   10134
      _Version        =   393216
      Cols            =   11
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      GridLines       =   2
      FormatString    =   $"frmNcri.frx":0C28
   End
   Begin MSComCtl2.DTPicker dtFrom 
      Height          =   330
      Left            =   990
      TabIndex        =   2
      Top             =   360
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "dd/MMM/yyyy"
      Format          =   59768835
      CurrentDate     =   38979
   End
   Begin VB.PictureBox SSPanel1 
      Height          =   750
      Index           =   1
      Left            =   135
      ScaleHeight     =   690
      ScaleWidth      =   5625
      TabIndex        =   11
      Top             =   810
      Width           =   5685
      Begin VB.OptionButton oBetween 
         Caption         =   "Last Week"
         Height          =   225
         Index           =   0
         Left            =   165
         TabIndex        =   18
         Top             =   330
         Width           =   1095
      End
      Begin VB.OptionButton oBetween 
         Caption         =   "Last Month"
         Height          =   195
         Index           =   1
         Left            =   1395
         TabIndex        =   17
         Top             =   60
         Width           =   1155
      End
      Begin VB.OptionButton oBetween 
         Caption         =   "Last Full Month"
         Height          =   195
         Index           =   2
         Left            =   1395
         TabIndex        =   16
         Top             =   330
         Width           =   1395
      End
      Begin VB.OptionButton oBetween 
         Caption         =   "Last Quarter"
         Height          =   195
         Index           =   3
         Left            =   2895
         TabIndex        =   15
         Top             =   90
         Width           =   1185
      End
      Begin VB.OptionButton oBetween 
         Caption         =   "Last Full Quarter"
         Height          =   195
         Index           =   4
         Left            =   2895
         TabIndex        =   14
         Top             =   360
         Width           =   1515
      End
      Begin VB.OptionButton oBetween 
         Caption         =   "Year To Date"
         Height          =   195
         Index           =   5
         Left            =   4275
         TabIndex        =   13
         Top             =   90
         Width           =   1305
      End
      Begin VB.OptionButton oBetween 
         Caption         =   "Today"
         Height          =   195
         Index           =   6
         Left            =   165
         TabIndex        =   12
         Top             =   90
         Value           =   -1  'True
         Width           =   795
      End
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
      Left            =   12120
      TabIndex        =   19
      Top             =   1230
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Label Label2 
      Caption         =   "To"
      Height          =   285
      Left            =   2430
      TabIndex        =   6
      Top             =   405
      Width           =   645
   End
   Begin VB.Label Label1 
      Caption         =   "From"
      Height          =   285
      Left            =   360
      TabIndex        =   5
      Top             =   405
      Width           =   645
   End
End
Attribute VB_Name = "frmNcri"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()

10        On Error GoTo cmdClose_Click_Error

20        fraReport.Visible = False
30        txtReport = ""

40        Exit Sub

cmdClose_Click_Error:

          Dim strES As String
          Dim intEL As Integer



50        intEL = Erl
60        strES = Err.Description
70        LogError "frmNcri", "cmdClose_Click", intEL, strES


End Sub

Private Sub cmdExcel_Click()

10        On Error GoTo cmdExcel_Click_Error

20        ExportFlexGrid grdNcri, Me

30        Exit Sub

cmdExcel_Click_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmNcri", "cmdExcel_Click", intEL, strES


End Sub

Private Sub cmdExit_Click()

10        Unload Me

End Sub


Private Sub cmdStart_Click()
          Dim tb As Recordset
          Dim sql As String
          Dim s As String

10        On Error GoTo cmdStart_Click_Error

20        With grdNcri
30            .Rows = 2
40            .AddItem ""
50            .RemoveItem 1
60        End With

70        sql = "select * from demographics, historesults where historesults.ncri = 1 " & _
                "and historesults.sampleid = demographics.sampleid and " & _
                "demographics.sampledate between '" & Format(dtFrom, "dd/MMM/yyyy") & " 00:00:00'" & _
                " and '" & Format(dtTo, "dd/MMM/yyyy") & " 23:59:59'"

80        Set tb = New Recordset
90        RecOpenServer 0, tb, sql
100       Do While Not tb.EOF
110           s = Trim(tb!Hyear) & "/" & Format$(Val(tb!SampleID) - (SysOptHistoOffset(0) + (Val(Swap_Year(Trim(tb!Hyear))) * 1000))) & "/H"
120           s = s & vbTab & Trim(tb!PatName) & vbTab
130           If IsDate(tb!Dob) Then s = s & Format(tb!Dob, "dd/MMM/yyyy") & vbTab Else s = s & vbTab
140           s = s & tb!sex & vbTab & tb!Addr0 & vbTab & tb!Hospital & vbTab & tb!Chart & vbTab & tb!GP & vbTab & tb!Clinician & vbTab
150           If IsDate(tb!SampleDate) Then s = s & Format(tb!SampleDate, "dd/MMM/yyyy") & vbTab Else s = s & vbTab
160           s = s & Trim(tb!historeport)
170           grdNcri.AddItem s
180           grdNcri.RowHeight(grdNcri.Rows - 1) = 4000
190           tb.MoveNext
200       Loop


210       sql = "select * from demographics, cytoresults where cytoresults.ncri = 1 " & _
                "and cytoresults.sampleid = demographics.sampleid and " & _
                "demographics.sampledate between '" & Format(dtFrom, "dd/MMM/yyyy") & " 00:00:00'" & _
                " and '" & Format(dtTo, "dd/MMM/yyyy") & " 23:59:59'"

220       Set tb = New Recordset
230       RecOpenServer 0, tb, sql
240       Do While Not tb.EOF
250           s = Trim(tb!Hyear) & "/" & Format$(Val(tb!SampleID) - (SysOptCytoOffset(0) + (Val(Swap_Year(Trim(tb!Hyear))) * 1000))) & "/C"
260           s = s & vbTab & Trim(tb!PatName) & vbTab
270           If IsDate(tb!Dob) Then s = s & Format(tb!Dob, "dd/MMM/yyyy") & vbTab Else s = s & vbTab
280           s = s & tb!sex & vbTab & tb!Addr0 & vbTab & tb!Hospital & vbTab & tb!Chart & vbTab & tb!GP & vbTab & tb!Clinician & vbTab
290           If IsDate(tb!SampleDate) Then s = s & Format(tb!SampleDate, "dd/MMM/yyyy") & vbTab Else s = s & vbTab
300           s = s & Trim(tb!cytoreport)
310           grdNcri.AddItem s
320           grdNcri.RowHeight(grdNcri.Rows - 1) = 4000
330           tb.MoveNext
340       Loop


350       If grdNcri.Rows > 2 And grdNcri.TextMatrix(1, 0) = "" Then grdNcri.RemoveItem 1


360       Exit Sub

cmdStart_Click_Error:

          Dim strES As String
          Dim intEL As Integer



370       intEL = Erl
380       strES = Err.Description
390       LogError "frmNcri", "cmdStart_Click", intEL, strES, sql


End Sub

Private Sub Form_Load()

10        On Error GoTo Form_Load_Error

20        dtTo = Format(Now, "dd/MMM/yyyy")
30        dtFrom = Format(Now - 7, "dd/MMM/yyyy")


40        Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer



50        intEL = Erl
60        strES = Err.Description
70        LogError "frmNcri", "Form_Load", intEL, strES


End Sub

Private Sub grdNcri_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        On Error GoTo grdNcri_MouseDown_Error

20        If Button = 2 Then
30            txtReport = "Lab No. : " & grdNcri.TextMatrix(grdNcri.RowSel, 0) & vbCrLf
40            txtReport = txtReport & "Name : " & grdNcri.TextMatrix(grdNcri.RowSel, 1) & "   Date of Birth : " & grdNcri.TextMatrix(grdNcri.RowSel, 2) & vbCrLf
50            txtReport = txtReport & "Addr : " & grdNcri.TextMatrix(grdNcri.RowSel, 4)
60            txtReport = txtReport & vbCrLf & vbCrLf & vbCrLf & vbCrLf & grdNcri.TextMatrix(grdNcri.RowSel, 10)
70            fraReport.Visible = True
80        End If

90        Exit Sub

grdNcri_MouseDown_Error:

          Dim strES As String
          Dim intEL As Integer



100       intEL = Erl
110       strES = Err.Description
120       LogError "frmNcri", "grdNcri_MouseDown", intEL, strES


End Sub

Private Sub grdNcri_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

      'grdNcri.ToolTipText = ""
      '
      'If grdNcri.ColSel = 3 Then
      '    grdNcri.ToolTipText = grdNcri.TextMatrix(grdNcri.RowSel, grdNcri.ColSel)
      'End If
10        On Error GoTo grdNcri_MouseMove_Error



20        Exit Sub

grdNcri_MouseMove_Error:

          Dim strES As String
          Dim intEL As Integer



30        intEL = Erl
40        strES = Err.Description
50        LogError "frmNcri", "grdNcri_MouseMove", intEL, strES

End Sub

Private Sub oBetween_Click(Index As Integer)

          Dim upto As String

10        On Error GoTo oBetween_Click_Error

20        dtFrom = BetweenDates(Index, upto)
30        dtTo = upto

40        Exit Sub

oBetween_Click_Error:

          Dim strES As String
          Dim intEL As Integer



50        intEL = Erl
60        strES = Err.Description
70        LogError "frmNcri", "oBetween_Click", intEL, strES


End Sub
