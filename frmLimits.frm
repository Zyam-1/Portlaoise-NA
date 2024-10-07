VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmLimits 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Control Limits"
   ClientHeight    =   5325
   ClientLeft      =   1980
   ClientTop       =   1935
   ClientWidth     =   6255
   Icon            =   "frmLimits.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbQc 
      Height          =   315
      Left            =   1305
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   405
      Width           =   2670
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   4185
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   3885
      _ExtentX        =   6853
      _ExtentY        =   7382
      _Version        =   393216
      Cols            =   3
   End
   Begin VB.CommandButton bcopypaste 
      Caption         =   "Copy Data"
      Height          =   825
      Left            =   4500
      Picture         =   "frmLimits.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   135
      Width           =   1365
   End
   Begin VB.TextBox tnewname 
      Height          =   285
      Left            =   4320
      MaxLength       =   10
      TabIndex        =   3
      Top             =   2760
      Visible         =   0   'False
      Width           =   1785
   End
   Begin VB.CommandButton bnew 
      Caption         =   "&New"
      Height          =   255
      Left            =   4770
      TabIndex        =   1
      Top             =   2460
      Width           =   825
   End
   Begin VB.CommandButton bcancel 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   840
      Left            =   4455
      Picture         =   "frmLimits.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4320
      Width           =   1365
   End
   Begin VB.Label Label1 
      Caption         =   "Control Name"
      Height          =   285
      Left            =   180
      TabIndex        =   7
      Top             =   405
      Width           =   1320
   End
   Begin VB.Label lmsg 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Enter new control name."
      Height          =   255
      Left            =   4290
      TabIndex        =   2
      Top             =   2430
      Visible         =   0   'False
      Width           =   1815
   End
End
Attribute VB_Name = "frmLimits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim block() As Single

Private Sub copyblock()

          Dim n As Long

10        On Error GoTo copyblock_Error

20        ReDim block(1 To 2, 1 To g.Rows - 1)
30        For n = 1 To g.Rows - 1
40            g.Row = n
50            g.Col = 1
60            block(1, n) = Val(g)
70            g.Col = 2
80            block(2, n) = Val(g)
90        Next

100       Exit Sub

copyblock_Error:

          Dim strES As String
          Dim intEL As Integer



110       intEL = Erl
120       strES = Err.Description
130       LogError "frmLimits", "copyblock", intEL, strES


End Sub
Private Sub drawgrid()

          Dim sn As New Recordset
          Dim sql As String
          Dim n As Long
          Dim s As String
          Dim Found As String

10        On Error GoTo drawgrid_Error

20        sql = "SELECT * from controls WHERE " & _
                "controlname = '" & cmbQc & "' " & _
                "order by parameter"
30        Set sn = New Recordset

40        RecOpenServer 0, sn, sql

50        g.Rows = 2
60        g.AddItem ""
70        g.RemoveItem 1

80        g.Row = 0
90        For n = 0 To 2
100           g.Col = n
110           g = Choose(n + 1, "Analyte", "Mean", "1 SD")
120           g.ColWidth(n) = Choose(n + 1, 1500, 1000, 1000)
130       Next

140       If sn.EOF Then Exit Sub

150       Do While Not sn.EOF
160           If Found <> sn!Parameter Then
170               Found = sn!Parameter
                  '  sql = "SELECT longname from biotestdefinitions WHERE code = '" & Trim(sn!Parameter & "") & "' and inuse = '1'"
                  '  Set tb = New Recordset
                  '  RecOpenServer 0, tb, sql
                  '  If Not tb.EOF Then
                  '    s = tb!LongName
                  '  End If
180               s = LongNameforCode(sn!Parameter)
190               s = s & vbTab & sn!mean & vbTab & _
                      sn("1sd")
200               g.AddItem s
210           End If
220           sn.MoveNext
230       Loop

240       sn.Close

250       If g.Rows > 2 And g.TextMatrix(1, 0) = "" Then
260           g.RemoveItem 1
270       End If

280       Exit Sub

drawgrid_Error:

          Dim strES As String
          Dim intEL As Integer



290       intEL = Erl
300       strES = Err.Description
310       LogError "frmLimits", "drawgrid", intEL, strES, sql


End Sub

Private Sub filllname()
          Dim sn As New Recordset
          Dim sql As String

10        On Error GoTo filllname_Error

20        cmbQc.Clear

30        sql = "SELECT distinct controlname from controls order by controlname"
40        Set sn = New Recordset
50        RecOpenServer 0, sn, sql
60        Do While Not sn.EOF
70            cmbQc.AddItem Trim(sn!ControlName & "")
80            sn.MoveNext
90        Loop


100       Exit Sub

filllname_Error:

          Dim strES As String
          Dim intEL As Integer



110       intEL = Erl
120       strES = Err.Description
130       LogError "frmLimits", "filllname", intEL, strES, sql


End Sub

Private Sub pasteblock()

          Dim n As Long
          Dim tb As New Recordset
          Dim sql As String


10        On Error GoTo pasteblock_Error

20        For n = 1 To g.Rows - 1
30            g.Row = n
40            g.Col = 1
50            g = block(1, n)
60            g.Col = 2
70            g = block(2, n)
80        Next


90        For n = 1 To g.Rows - 1

100           sql = "SELECT * from controls WHERE controlname = '" & cmbQc & "' and parameter = '" & CodeForLongName(g.TextMatrix(n, 0)) & "'"
110           Set tb = New Recordset
120           RecOpenServer 0, tb, sql

130           If tb.EOF Then tb.AddNew
140           tb!ControlName = cmbQc
150           tb!Parameter = CodeForLongName(g.TextMatrix(n, 0))
160           tb!mean = g.TextMatrix(n, 1)
170           tb("1sd") = g.TextMatrix(n, 2)
180           tb.Update
190       Next

200       Exit Sub

pasteblock_Error:

          Dim strES As String
          Dim intEL As Integer



210       intEL = Erl
220       strES = Err.Description
230       LogError "frmLimits", "pasteblock", intEL, strES, sql


End Sub

Private Sub bcancel_Click()

10        Unload Me

End Sub


Private Sub bcopypaste_Click()

10        On Error GoTo bcopypaste_Click_Error

20        If bcopypaste.Caption = "Copy Data" Then
30            copyblock
40            bcopypaste.Caption = "Paste Data"
50        Else
60            pasteblock
70            bcopypaste.Caption = "Copy Data"
80        End If

90        Exit Sub

bcopypaste_Click_Error:

          Dim strES As String
          Dim intEL As Integer

100       intEL = Erl
110       strES = Err.Description
120       LogError "frmLimits", "bcopypaste_Click", intEL, strES

End Sub

Private Sub bnew_Click()

10        On Error GoTo bnew_Click_Error

20        bnew.Visible = False
30        lmsg.Visible = True
40        tnewname.Visible = True

50        Exit Sub

bnew_Click_Error:

          Dim strES As String
          Dim intEL As Integer



60        intEL = Erl
70        strES = Err.Description
80        LogError "frmLimits", "bnew_Click", intEL, strES


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
70        LogError "frmLimits", "Form_Load", intEL, strES


End Sub

Private Sub g_KeyPress(KeyAscii As Integer)

          Dim ds As Recordset
          Dim sql As String
          Dim xSave As Long

10        On Error GoTo g_KeyPress_Error

20        If g.Row < 1 Then Exit Sub
30        If g.Col = 0 Then Exit Sub

40        If KeyAscii = 8 Then
50            g = ""
60        Else
70            g = g & Chr(KeyAscii)
80        End If

90        xSave = g.Col

100       g.Col = 0
110       sql = "SELECT * from controls WHERE " & _
                "parameter = '" & CodeForLongName(g) & _
                "' and controlname = '" & cmbQc & "'"
120       Set ds = New Recordset
130       RecOpenServer 0, ds, sql
140       If ds.EOF Then ds.AddNew
150       g.Col = xSave
160       Select Case g.Col
          Case 1: ds!mean = Val(g)
170       Case 2: ds("1sd") = Val(g)
180       End Select
190       ds.Update

200       Exit Sub

g_KeyPress_Error:

          Dim strES As String
          Dim intEL As Integer



210       intEL = Erl
220       strES = Err.Description
230       LogError "frmLimits", "g_KeyPress", intEL, strES, sql


End Sub

Private Sub cmbQc_Click()

10        On Error GoTo cmbQc_Click_Error

20        drawgrid

30        Exit Sub

cmbQc_Click_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmLimits", "cmbQc_Click", intEL, strES


End Sub
Private Sub tnewname_LostFocus()

          Dim tb As New Recordset
          Dim tbref As Recordset
          Dim sql As String

10        On Error GoTo tnewname_LostFocus_Error

20        lmsg.Visible = False
30        bnew.Visible = True
40        tnewname.Visible = False

50        If Trim(tnewname) = "" Then Exit Sub

60        tnewname = UCase(tnewname)

70        sql = "SELECT * from controls WHERE controlname = '" & tnewname & "'"
80        Set tb = New Recordset
90        RecOpenServer 0, tb, sql


100       If tb.EOF Then
110           Set tbref = New Recordset
120           sql = "SELECT distinct(code) from biotestdefinitions WHERE inuse = '1' "
130           RecOpenServer 0, tbref, sql
140           Do While Not tbref.EOF
150               tb.AddNew
160               tb!ControlName = tnewname
170               tb!Parameter = tbref!Code
180               tb.Update
190               tbref.MoveNext
200           Loop
210       End If

220       filllname
230       drawgrid

240       GetControlNames

250       Exit Sub

tnewname_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer



260       intEL = Erl
270       strES = Err.Description
280       LogError "frmLimits", "tnewname_LostFocus", intEL, strES, sql


End Sub


