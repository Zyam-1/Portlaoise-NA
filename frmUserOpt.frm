VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmUserOpt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "User Options for "
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4590
   Icon            =   "frmUserOpt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Exit"
      Height          =   825
      Left            =   3420
      Picture         =   "frmUserOpt.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3060
      Width           =   1005
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   825
      Left            =   3420
      Picture         =   "frmUserOpt.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   270
      Width           =   1005
   End
   Begin MSFlexGridLib.MSFlexGrid grdOpt 
      Height          =   3615
      Left            =   315
      TabIndex        =   0
      Top             =   270
      Width           =   2985
      _ExtentX        =   5265
      _ExtentY        =   6376
      _Version        =   393216
      FormatString    =   "Description           | Option             "
   End
End
Attribute VB_Name = "frmUserOpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

10        Unload Me

End Sub

Private Sub grdOpt_Click()

          Dim s As String
          Dim Msg As String



10        On Error GoTo grdOpt_Click_Error

20        If grdOpt.TextMatrix(grdOpt.RowSel, 0) = "DEFAULTTAB" Then
30            Msg = "Default Tab Set to " & frmEditAll.ssTabAll.TabCaption(grdOpt.TextMatrix(grdOpt.RowSel, 0))
40        End If

50        If grdOpt.Col = 1 Then
60            If grdOpt.TextMatrix(grdOpt.RowSel, 1) = "True" Then
70                grdOpt.TextMatrix(grdOpt.RowSel, 1) = "False"
80            ElseIf grdOpt.TextMatrix(grdOpt.RowSel, 1) = "False" Then
90                grdOpt.TextMatrix(grdOpt.RowSel, 1) = "True"
100           Else
110               s = iBOX("Change", , grdOpt.TextMatrix(grdOpt.RowSel, 1), False)
120               If s <> "" Then grdOpt.TextMatrix(grdOpt.RowSel, 1) = s
130           End If
140       End If

150       cmdUpdate.Enabled = True




160       Exit Sub

grdOpt_Click_Error:

          Dim strES As String
          Dim intEL As Integer



170       intEL = Erl
180       strES = Err.Description
190       LogError "frmUserOpt", "grdOpt_Click", intEL, strES


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
140           tb!Username = Trim(Username)
150           tb.Update
160       Next

170       LoadUserOpts


180       sql = "INSERT into UPDATEs  (upd, dtime) values ('Option', '" & Format(Now, "dd/MMM/yyyy hh:mm") & "')"
190       Cnxn(0).Execute sql

200       cmdUpdate.Enabled = False



210       Exit Sub

cmdUpdate_Click_Error:

          Dim strES As String
          Dim intEL As Integer



220       intEL = Erl
230       strES = Err.Description
240       LogError "frmUserOpt", "cmdUpdate_Click", intEL, strES, sql


End Sub


Private Sub Form_Load()
          Dim tb As New Recordset
          Dim sql As String
          Dim s As String


10        On Error GoTo Form_Load_Error

20        Set_Font Me

30        Me.Caption = Me.Caption & Username

40        sql = "SELECT * from options WHERE username = '" & Username & "' order by description"
50        Set tb = New Recordset
60        RecOpenServer 0, tb, sql

70        Do While Not tb.EOF
80            s = Trim(tb!Description & "") & vbTab
90            If Trim(tb!Contents & "") = 1 Then
100               s = s & "True" & vbTab
110           ElseIf Trim(tb!Contents & "") = 0 Then
120               s = s & "False" & vbTab
130           Else
140               s = s & Trim(tb!Contents & "") & vbTab
150           End If
160           grdOpt.AddItem s
170           tb.MoveNext
180       Loop


190       If grdOpt.Rows > 2 And grdOpt.TextMatrix(1, 0) = "" Then grdOpt.RemoveItem 1



200       Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer



210       intEL = Erl
220       strES = Err.Description
230       LogError "frmUserOpt", "Form_Load", intEL, strES, sql


End Sub



