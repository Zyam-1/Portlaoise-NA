VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmCoagLimits 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Coagulation Control Limits"
   ClientHeight    =   5325
   ClientLeft      =   1980
   ClientTop       =   1935
   ClientWidth     =   6585
   Icon            =   "frmCoagLimits.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   Begin MSFlexGridLib.MSFlexGrid grdLim 
      Height          =   4185
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   3885
      _ExtentX        =   6853
      _ExtentY        =   7382
      _Version        =   393216
      Cols            =   3
   End
   Begin VB.CommandButton cmdCopyPaste 
      Caption         =   "Copy Data"
      Height          =   825
      Left            =   4500
      Picture         =   "frmCoagLimits.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   945
      Width           =   1410
   End
   Begin VB.TextBox txtNewName 
      Height          =   285
      Left            =   4320
      MaxLength       =   10
      TabIndex        =   4
      Top             =   2760
      Visible         =   0   'False
      Width           =   1785
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      Height          =   255
      Left            =   4770
      TabIndex        =   2
      Top             =   2460
      Width           =   825
   End
   Begin VB.ListBox lstName 
      Height          =   840
      Left            =   90
      TabIndex        =   1
      Top             =   60
      Width           =   3915
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   840
      Left            =   4590
      Picture         =   "frmCoagLimits.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4275
      Width           =   1365
   End
   Begin VB.Label lmsg 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Enter new control name."
      Height          =   255
      Left            =   4290
      TabIndex        =   3
      Top             =   2430
      Visible         =   0   'False
      Width           =   1815
   End
End
Attribute VB_Name = "frmCoagLimits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim block() As Single

Private Sub cmdCancel_Click()

10        Unload Me

End Sub

Private Sub cmdCopyPaste_Click()

10        On Error GoTo cmdCopyPaste_Click_Error

20        If cmdCopyPaste.Caption = "Copy Data" Then
30            copyblock
40            cmdCopyPaste.Caption = "Paste Data"
50        Else
60            pasteblock
70            cmdCopyPaste.Caption = "Copy Data"
80        End If

90        Exit Sub

cmdCopyPaste_Click_Error:

          Dim strES As String
          Dim intEL As Integer

100       intEL = Erl
110       strES = Err.Description
120       LogError "frmCoagLimits", "cmdCopyPaste_Click", intEL, strES

End Sub

Private Sub cmdNew_Click()

10        On Error GoTo cmdNew_Click_Error

20        cmdNew.Visible = False
30        grdLim.Visible = True
40        txtNewName.Visible = True

50        Exit Sub

cmdNew_Click_Error:

          Dim strES As String
          Dim intEL As Integer

60        intEL = Erl
70        strES = Err.Description
80        LogError "frmCoagLimits", "cmdNew_Click", intEL, strES


End Sub

Private Sub copyblock()

          Dim n As Long

10        On Error GoTo copyblock_Error

20        ReDim block(1 To 2, 1 To grdLim.Rows - 1)
30        For n = 1 To grdLim.Rows - 1
40            grdLim.Row = n
50            grdLim.Col = 1
60            block(1, n) = Val(grdLim)
70            grdLim.Col = 2
80            block(2, n) = Val(grdLim)
90        Next

100       Exit Sub

copyblock_Error:

          Dim strES As String
          Dim intEL As Integer

110       intEL = Erl
120       strES = Err.Description
130       LogError "frmCoagLimits", "copyblock", intEL, strES


End Sub

Private Sub drawgrid()

          Dim sn As New Recordset
          Dim sql As String
          Dim n As Long
          Dim s As String

10        On Error GoTo drawgrid_Error

20        sql = "SELECT * from CoagControls WHERE " & _
                "controlname = '" & lstName & "' " & _
                "order by parameter"
30        Set sn = New Recordset

40        RecOpenServer 0, sn, sql

50        ClearFGrid grdLim


60        grdLim.Row = 0
70        For n = 0 To 2
80            grdLim.Col = n
90            grdLim = Choose(n + 1, "Analyte", "Mean", "1 SD")
100           grdLim.ColWidth(n) = Choose(n + 1, 1500, 1000, 1000)
110       Next

120       If sn.EOF Then Exit Sub

130       Do While Not sn.EOF
140           s = CoagNameFor(sn!Parameter & "") & vbTab & _
                  sn!mean & vbTab & _
                  sn("1sd")
150           grdLim.AddItem s
160           sn.MoveNext
170       Loop

180       sn.Close

190       FixG grdLim

200       Exit Sub

drawgrid_Error:

          Dim strES As String
          Dim intEL As Integer

210       intEL = Erl
220       strES = Err.Description
230       LogError "frmCoagLimits", "drawgrid", intEL, strES, sql


End Sub

Private Sub filllname()
          Dim sn As New Recordset
          Dim sql As String

10        On Error GoTo filllname_Error

20        lstName.Clear

30        sql = "SELECT distinct controlname from coagcontrols order by controlname"
40        Set sn = New Recordset
50        RecOpenServer 0, sn, sql
60        Do While Not sn.EOF
70            lstName.AddItem Trim(sn!ControlName & "")
80            sn.MoveNext
90        Loop

100       If lstName.ListCount <> 0 Then
110           lstName.Selected(0) = True
120       End If

130       Exit Sub

filllname_Error:

          Dim strES As String
          Dim intEL As Integer

140       intEL = Erl
150       strES = Err.Description
160       LogError "frmCoagLimits", "filllname", intEL, strES


End Sub

Private Sub Form_Load()

10        On Error GoTo Form_Load_Error

20        filllname
30        drawgrid

40        Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmCoagLimits", "Form_Load", intEL, strES


End Sub

Private Sub grdLim_KeyPress(KeyAscii As Integer)

          Dim ds As Recordset
          Dim sql As String
          Dim xSave As Long

10        On Error GoTo grdLim_KeyPress_Error

20        If grdLim.Row < 1 Then Exit Sub
30        If grdLim.Col = 0 Then Exit Sub

40        If KeyAscii = 8 Then
50            grdLim = ""
60        Else
70            grdLim = grdLim & Chr(KeyAscii)
80        End If

90        xSave = grdLim.Col

100       grdLim.Col = 0
110       sql = "SELECT * from coagcontrols WHERE " & _
                "parameter = " & CoagCodeFor(grdLim) & _
                " and controlname = '" & lstName & "'"
120       Set ds = New Recordset
130       RecOpenServer 0, ds, sql
140       If ds.EOF Then ds.AddNew
150       grdLim.Col = xSave
160       Select Case grdLim.Col
          Case 1: ds!mean = Val(grdLim)
170       Case 2: ds("1sd") = Val(grdLim)
180       End Select
190       ds.Update

200       Exit Sub

grdLim_KeyPress_Error:

          Dim strES As String
          Dim intEL As Integer

210       intEL = Erl
220       strES = Err.Description
230       LogError "frmCoagLimits", "grdLim_KeyPress", intEL, strES


End Sub

Private Sub lstName_Click()

10        On Error GoTo lstName_Click_Error

20        drawgrid

30        Exit Sub

lstName_Click_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmCoagLimits", "lstName_Click", intEL, strES


End Sub

Private Sub pasteblock()

          Dim n As Long
          Dim tb As New Recordset
          Dim sql As String

10        On Error GoTo pasteblock_Error

20        For n = 1 To grdLim.Rows - 1
30            grdLim.Row = n
40            grdLim.Col = 1
50            grdLim = block(1, n)
60            grdLim.Col = 2
70            grdLim = block(2, n)
80        Next

90        For n = 1 To grdLim.Rows - 1

100           sql = "SELECT * from coagcontrols WHERE " & _
                    "controlname = '" & lstName & "' " & _
                    "and parameter = '" & CoagCodeFor(grdLim.TextMatrix(n, 0)) & "'"
110           Set tb = New Recordset
120           RecOpenServer 0, tb, sql

130           If tb.EOF Then tb.AddNew
140           tb!ControlName = lstName
150           tb!Parameter = CoagCodeFor(grdLim.TextMatrix(n, 0))
160           tb!mean = grdLim.TextMatrix(n, 1)
170           tb("1sd") = grdLim.TextMatrix(n, 2)
180           tb.Update
190       Next

200       Exit Sub

pasteblock_Error:

          Dim strES As String
          Dim intEL As Integer

210       intEL = Erl
220       strES = Err.Description
230       LogError "frmCoagLimits", "pasteblock", intEL, strES

End Sub

Private Sub txtNewName_LostFocus()

          Dim tb As New Recordset
          Dim tbref As Recordset
          Dim sql As String

10        On Error GoTo txtNewName_LostFocus_Error

20        grdLim.Visible = False
30        cmdNew.Visible = True
40        txtNewName.Visible = False

50        If Trim(txtNewName) = "" Then Exit Sub


60        sql = "SELECT * from coagcontrols WHERE controlname = '" & txtNewName & "'"
70        Set tb = New Recordset
80        RecOpenServer 0, tb, sql


90        If tb.EOF Then
100           Set tbref = New Recordset
110           sql = "SELECT distinct(code) from coagtestdefinitions WHERE inuse = '1' "
120           RecOpenServer 0, tbref, sql
130           Do While Not tbref.EOF
140               tb.AddNew
150               tb!ControlName = txtNewName
160               tb!Parameter = tbref!Code
170               tb.Update
180               tbref.MoveNext
190           Loop
200       End If

210       filllname
220       drawgrid

230       GetControlNames

240       Exit Sub

txtNewName_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

250       intEL = Erl
260       strES = Err.Description
270       LogError "frmCoagLimits", "txtNewName_LostFocus", intEL, strES, sql


End Sub
