VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frmOption 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire 6.9 - System Options"
   ClientHeight    =   8655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8760
   Icon            =   "frmOption.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8655
   ScaleWidth      =   8760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Add Options"
      Height          =   1725
      Left            =   45
      TabIndex        =   3
      Top             =   6885
      Width           =   8700
      Begin VB.OptionButton optSys 
         Caption         =   "Users"
         Height          =   285
         Index           =   1
         Left            =   4410
         TabIndex        =   10
         Top             =   990
         Width           =   1680
      End
      Begin VB.OptionButton optSys 
         Caption         =   "System"
         Height          =   285
         Index           =   0
         Left            =   4410
         TabIndex        =   9
         Top             =   495
         Width           =   1680
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   825
         Left            =   6390
         Picture         =   "frmOption.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   450
         Width           =   1905
      End
      Begin VB.TextBox txtContent 
         Height          =   375
         Left            =   1575
         TabIndex        =   7
         Top             =   945
         Width           =   2490
      End
      Begin VB.TextBox txtDescription 
         Height          =   375
         Left            =   1575
         TabIndex        =   4
         Top             =   495
         Width           =   2490
      End
      Begin VB.Label Label2 
         Caption         =   "Content"
         Height          =   240
         Left            =   270
         TabIndex        =   6
         Top             =   1035
         Width           =   1185
      End
      Begin VB.Label Label1 
         Caption         =   "Description"
         Height          =   240
         Left            =   270
         TabIndex        =   5
         Top             =   540
         Width           =   1095
      End
   End
   Begin Threed.SSCommand cmdCancel 
      Height          =   870
      Left            =   4590
      TabIndex        =   2
      Top             =   5805
      Width           =   1635
      _Version        =   65536
      _ExtentX        =   2884
      _ExtentY        =   1535
      _StockProps     =   78
      Caption         =   "Exit"
      Picture         =   "frmOption.frx":0614
   End
   Begin Threed.SSCommand cmdUpdate 
      Height          =   870
      Left            =   2520
      TabIndex        =   1
      Top             =   5805
      Width           =   1770
      _Version        =   65536
      _ExtentX        =   3122
      _ExtentY        =   1535
      _StockProps     =   78
      Caption         =   "Update"
      Enabled         =   0   'False
      Picture         =   "frmOption.frx":092E
   End
   Begin MSFlexGridLib.MSFlexGrid grdOpt 
      Height          =   5505
      Left            =   360
      TabIndex        =   0
      Top             =   270
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   9710
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      FormatString    =   $"frmOption.frx":0C48
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
End
Attribute VB_Name = "frmOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub cmdadd_Click()

          Dim sql As String
          Dim tb As New Recordset
          Dim sn As New Recordset

10        On Error GoTo cmdadd_Click_Error

20        If Trim$(txtDescription) = "" Then Exit Sub

30        If Trim$(txtContent) = "" Then Exit Sub

40        If optSys(1) Then
50            sql = "SELECT name from users"
60            Set tb = New Recordset
70            RecOpenServer 0, tb, sql
80            Do While Not tb.EOF
90                sql = "SELECT * from options WHERE " & _
                        "description = '" & txtDescription & "' " & _
                        "and username = '" & AddTicks(tb!Name) & "'"
100               Set sn = New Recordset
110               RecOpenServer 0, sn, sql
120               If sn.EOF Then sn.AddNew
130               sn!Description = txtDescription
140               sn!Contents = txtContent
150               sn!Username = tb!Name
160               sn.Update
170               tb.MoveNext
180           Loop
190       Else
200           sql = "SELECT * from options WHERE " & _
                    "description = '" & txtDescription & "' "
210           Set sn = New Recordset
220           RecOpenServer 0, sn, sql
230           If sn.EOF Then sn.AddNew
240           sn!Description = txtDescription
250           sn!Contents = txtContent
260           sn.Update
270       End If

280       txtDescription = ""
290       txtContent = ""

300       Load_Options

310       Exit Sub

cmdadd_Click_Error:

          Dim strES As String
          Dim intEL As Integer

320       intEL = Erl
330       strES = Err.Description
340       LogError "frmOption", "cmdAdd_Click", intEL, strES, sql

End Sub

Private Sub cmdCancel_Click()

10        Unload Me

End Sub

Private Sub cmdUpdate_Click()

          Dim tb As New Recordset
          Dim sql As String
          Dim n As Long

10        On Error GoTo cmdUpdate_Click_Error

20        For n = 1 To grdOpt.Rows - 1
30            sql = "SELECT * from options WHERE description = '" & grdOpt.TextMatrix(n, 0) & "'"
40            Set tb = New Recordset
50            RecOpenServer 0, tb, sql
60            tb!Description = grdOpt.TextMatrix(n, 0)
70            If grdOpt.TextMatrix(n, 1) = "True" Then
80                tb!Contents = 1
90            ElseIf grdOpt.TextMatrix(n, 1) = "False" Then
100               tb!Contents = 0
110           Else
120               tb!Contents = grdOpt.TextMatrix(n, 1)
130           End If
140           tb.Update
150       Next

160       LoadOptions

170       sql = "INSERT into UPDATEs  (upd, dtime) values ('Option', '" & Format(Now, "dd/MMM/yyyy hh:mm") & "')"
180       Cnxn(0).Execute sql

190       Exit Sub

cmdUpdate_Click_Error:

          Dim strES As String
          Dim intEL As Integer

200       intEL = Erl
210       strES = Err.Description
220       LogError "frmOption", "cmdUpdate_Click", intEL, strES, sql

End Sub

Private Sub Form_Load()

10        On Error GoTo Form_Load_Error

20        Load_Options

30        Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmOption", "Form_Load", intEL, strES

End Sub

Private Sub grdOpt_Click()
          Dim s As String


10        On Error GoTo grdOpt_Click_Error

20        If grdOpt.Col = 1 Then
30            If grdOpt.TextMatrix(grdOpt.RowSel, 1) = "True" Then
40                grdOpt.TextMatrix(grdOpt.RowSel, 1) = "False"
50            ElseIf grdOpt.TextMatrix(grdOpt.RowSel, 1) = "False" Then
60                grdOpt.TextMatrix(grdOpt.RowSel, 1) = "True"
70            Else
80                s = iBOX("Change", , grdOpt.TextMatrix(grdOpt.RowSel, 1), False)
90                If s <> "" Then grdOpt.TextMatrix(grdOpt.RowSel, 1) = s
100           End If
110       End If

120       cmdUpdate.Enabled = True




130       Exit Sub

grdOpt_Click_Error:

          Dim strES As String
          Dim intEL As Integer



140       intEL = Erl
150       strES = Err.Description
160       LogError "frmOption", "grdOpt_Click", intEL, strES


End Sub

Private Sub Load_Options()

          Dim tb As New Recordset
          Dim sql As String
          Dim s As String

10        On Error GoTo Load_Options_Error

20        ClearFGrid grdOpt

30        sql = "SELECT * from options WHERE username = '' or username is null order by description"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql

60        Do While Not tb.EOF
70            s = UCase(Trim(tb!Description & "")) & vbTab
80            If Trim(tb!Contents & "") = 1 Then
90                s = s & "True" & vbTab
100           ElseIf Trim(tb!Contents & "") = 0 Then
110               s = s & "False" & vbTab
120           Else
130               s = s & Trim(tb!Contents & "") & vbTab
140           End If
150           grdOpt.AddItem s
160           tb.MoveNext
170       Loop

180       FixG grdOpt

190       Exit Sub

Load_Options_Error:

          Dim strES As String
          Dim intEL As Integer

200       intEL = Erl
210       strES = Err.Description
220       LogError "frmOption", "Load_Options", intEL, strES, sql

End Sub
