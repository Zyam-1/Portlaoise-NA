VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmBioPODPU 
   Caption         =   "NetAcquire - Biochemistry"
   ClientHeight    =   6465
   ClientLeft      =   1725
   ClientTop       =   1260
   ClientWidth     =   7545
   Icon            =   "frmBioPODPU.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   6465
   ScaleWidth      =   7545
   Begin VB.Timer tmrDown 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   6660
      Top             =   1950
   End
   Begin VB.Timer tmrUp 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   6660
      Top             =   960
   End
   Begin MSComctlLib.ProgressBar pbY 
      Height          =   225
      Left            =   90
      TabIndex        =   8
      Top             =   0
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   225
      Left            =   90
      TabIndex        =   7
      Top             =   210
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.ComboBox lUnits 
      Height          =   315
      Left            =   4110
      TabIndex        =   6
      Text            =   "lUnits"
      Top             =   600
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdMoveDown 
      Caption         =   "Move &Down"
      Height          =   1035
      Left            =   6030
      Picture         =   "frmBioPODPU.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1650
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdMoveUp 
      Caption         =   "Move &Up"
      Height          =   1035
      Left            =   6030
      Picture         =   "frmBioPODPU.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H0080C0FF&
      Caption         =   "&Save"
      Height          =   765
      Left            =   6255
      Picture         =   "frmBioPODPU.frx":0B8E
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5580
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   765
      Left            =   6270
      Picture         =   "frmBioPODPU.frx":0E98
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4050
      Width           =   1080
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   675
      Left            =   6270
      Picture         =   "frmBioPODPU.frx":11A2
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3150
      Width           =   1080
   End
   Begin MSFlexGridLib.MSFlexGrid grdBio 
      Height          =   6045
      Left            =   60
      TabIndex        =   2
      Top             =   420
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   10663
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      FormatString    =   "Test Name                         |^Decimal Places |<Units                 "
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
Attribute VB_Name = "frmBioPODPU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FireCounter As Long

Private Sub cmdCancel_Click()

10        Unload Me

End Sub

Private Sub cmdMoveDown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        On Error GoTo cmdMoveDown_MouseDown_Error

20        FireDown

30        tmrDown.Interval = 250
40        FireCounter = 0

50        tmrDown.Enabled = True

60        Exit Sub

cmdMoveDown_MouseDown_Error:

          Dim strES As String
          Dim intEL As Integer

70        intEL = Erl
80        strES = Err.Description
90        LogError "frmBioPODPU", "cmdMoveDown_MouseDown", intEL, strES


End Sub

Private Sub cmdMoveDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        On Error GoTo cmdMoveDown_MouseUp_Error

20        tmrDown.Enabled = False

30        Exit Sub

cmdMoveDown_MouseUp_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmBioPODPU", "cmdMoveDown_MouseUp", intEL, strES


End Sub

Private Sub cmdMoveUp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        On Error GoTo cmdMoveUp_MouseDown_Error

20        FireUp

30        tmrUp.Interval = 250
40        FireCounter = 0

50        tmrUp.Enabled = True

60        Exit Sub

cmdMoveUp_MouseDown_Error:

          Dim strES As String
          Dim intEL As Integer

70        intEL = Erl
80        strES = Err.Description
90        LogError "frmBioPODPU", "cmdMoveUp_MouseDown", intEL, strES


End Sub

Private Sub cmdMoveUp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        On Error GoTo cmdMoveUp_MouseUp_Error

20        tmrUp.Enabled = False

30        Exit Sub

cmdMoveUp_MouseUp_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmBioPODPU", "cmdMoveUp_MouseUp", intEL, strES


End Sub

Private Sub cmdPrint_Click()

          Dim Y As Long

10        On Error GoTo cmdPrint_Click_Error

20        For Y = 0 To grdBio.Rows - 1
30            Printer.Print grdBio.TextMatrix(Y, 0);
40            Printer.Print Tab(40); grdBio.TextMatrix(Y, 1);
50            Printer.Print Tab(65); grdBio.TextMatrix(Y, 2)
60        Next
70        Printer.EndDoc

80        Exit Sub

cmdPrint_Click_Error:

          Dim strES As String
          Dim intEL As Integer

90        intEL = Erl
100       strES = Err.Description
110       LogError "frmBioPODPU", "cmdPrint_Click", intEL, strES


End Sub

Private Sub cmdSave_Click()

10        SaveG

End Sub

Private Sub FillG()

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo FillG_Error

20        grdBio.Rows = 2
30        grdBio.AddItem ""
40        grdBio.RemoveItem 1

50        sql = "Select distinct ShortName, PrintPriority, " & _
                "DP, Units " & _
                "from BioTestDefinitions " & _
                "Order by PrintPriority"
60        Set tb = New Recordset
70        RecOpenServer 0, tb, sql
80        Do While Not tb.EOF
90            grdBio.AddItem tb!ShortName & vbTab & _
                             tb!DP & vbTab & _
                             tb!Units & ""
100           tb.MoveNext
110       Loop

120       If grdBio.Rows > 2 Then grdBio.RemoveItem 1

130       Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer

140       intEL = Erl
150       strES = Err.Description
160       LogError "frmBioPODPU", "FillG", intEL, strES, sql


End Sub

Private Sub FillUnits()

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo FillUnits_Error

20        sql = "Select * from Lists where " & _
                "ListType = 'UN' " & _
                "order by ListOrder"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        Do While Not tb.EOF
60            lUnits.AddItem tb!Text & ""
70            tb.MoveNext
80        Loop

90        Exit Sub

FillUnits_Error:

          Dim strES As String
          Dim intEL As Integer

100       intEL = Erl
110       strES = Err.Description
120       LogError "frmBioPODPU", "FillUnits", intEL, strES, sql


End Sub

Private Sub FireDown()

          Dim n As Long
          Dim s As String
          Dim X As Long
          Dim VisibleRows As Long

10        On Error GoTo FireDown_Error

20        If grdBio.Row = grdBio.Rows - 1 Then Exit Sub
30        n = grdBio.Row

40        VisibleRows = grdBio.Height \ grdBio.RowHeight(1) - 1

50        FireCounter = FireCounter + 1
60        If FireCounter > 5 Then
70            tmrDown.Interval = 100
80        End If

90        grdBio.Visible = False

100       s = ""
110       For X = 0 To grdBio.Cols - 1
120           s = s & grdBio.TextMatrix(n, X) & vbTab
130       Next
140       s = Left$(s, Len(s) - 1)

150       grdBio.RemoveItem n
160       If n < grdBio.Rows Then
170           grdBio.AddItem s, n + 1
180           grdBio.Row = n + 1
190       Else
200           grdBio.AddItem s
210           grdBio.Row = grdBio.Rows - 1
220       End If

230       For X = 0 To grdBio.Cols - 1
240           grdBio.Col = X
250           grdBio.CellBackColor = vbYellow
260       Next

270       If Not grdBio.RowIsVisible(grdBio.Row) Or grdBio.Row = grdBio.Rows - 1 Then
280           If grdBio.Row - VisibleRows + 1 > 0 Then
290               grdBio.TopRow = grdBio.Row - VisibleRows + 1
300           End If
310       End If

320       grdBio.Visible = True

330       cmdSave.Visible = True

340       Exit Sub

FireDown_Error:

          Dim strES As String
          Dim intEL As Integer

350       intEL = Erl
360       strES = Err.Description
370       LogError "frmBioPODPU", "FireDown", intEL, strES


End Sub

Private Sub FireUp()

          Dim n As Long
          Dim s As String
          Dim X As Long

10        On Error GoTo FireUp_Error

20        If grdBio.Row = 1 Then Exit Sub

30        FireCounter = FireCounter + 1
40        If FireCounter > 5 Then
50            tmrUp.Interval = 100
60        End If

70        n = grdBio.Row

80        grdBio.Visible = False

90        s = ""
100       For X = 0 To grdBio.Cols - 1
110           s = s & grdBio.TextMatrix(n, X) & vbTab
120       Next
130       s = Left$(s, Len(s) - 1)

140       grdBio.RemoveItem n
150       grdBio.AddItem s, n - 1

160       grdBio.Row = n - 1
170       For X = 0 To grdBio.Cols - 1
180           grdBio.Col = X
190           grdBio.CellBackColor = vbYellow
200       Next

210       If Not grdBio.RowIsVisible(grdBio.Row) Then
220           grdBio.TopRow = grdBio.Row
230       End If

240       grdBio.Visible = True

250       cmdSave.Visible = True

260       Exit Sub

FireUp_Error:

          Dim strES As String
          Dim intEL As Integer

270       intEL = Erl
280       strES = Err.Description
290       LogError "frmBioPODPU", "FireUp", intEL, strES


End Sub

Private Sub Form_Load()

10        On Error GoTo Form_Load_Error

20        grdBio.Font.Bold = True

30        FillG
40        FillUnits

50        Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

60        intEL = Erl
70        strES = Err.Description
80        LogError "frmBioPODPU", "Form_Load", intEL, strES


End Sub

Private Sub Form_Unload(Cancel As Integer)

10        On Error GoTo Form_Unload_Error

20        If cmdSave.Visible Then
30            If iMsg("Cancel without saving?", vbQuestion + vbYesNo) = vbNo Then
40                Cancel = True
50            End If
60        End If

70        Exit Sub

Form_Unload_Error:

          Dim strES As String
          Dim intEL As Integer

80        intEL = Erl
90        strES = Err.Description
100       LogError "frmBioPODPU", "Form_Unload", intEL, strES


End Sub

Private Sub grdBio_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

          Dim n As Long
          Dim yy As Long
          Dim ySave As Long

10        On Error GoTo grdBio_MouseUp_Error

20        If grdBio.MouseRow = 0 Then Exit Sub

30        If grdBio.Col = 0 Then
40            ySave = grdBio.Row
50            grdBio.Col = 0
60            For yy = 1 To grdBio.Rows - 1
70                grdBio.Row = yy
80                If grdBio.CellBackColor = vbYellow Then
90                    For n = 0 To grdBio.Cols - 1
100                       grdBio.Col = n
110                       grdBio.CellBackColor = 0
120                   Next
130                   Exit For
140               End If
150           Next
160           grdBio.Row = ySave
170           For n = 0 To grdBio.Cols - 1
180               grdBio.Col = n
190               grdBio.CellBackColor = vbYellow
200           Next
210           cmdMoveUp.Visible = True
220           cmdMoveDown.Visible = True
230       ElseIf grdBio.Col = 1 Then
240           Select Case grdBio
              Case "0": grdBio = "1"
250           Case "1": grdBio = "2"
260           Case "2": grdBio = "3"
270           Case "3": grdBio = "4"
280           Case Else: grdBio = "0"
290           End Select
300           cmdSave.Visible = True
310       Else
320           grdBio.Enabled = False
330           cmdMoveUp.Enabled = False
340           cmdMoveDown.Enabled = False
350           cmdPrint.Enabled = False
360           cmdSave.Enabled = False
370           cmdCancel.Enabled = False

380           lUnits.Top = grdBio.Top + Y - 100
390           lUnits = grdBio
400           lUnits.Visible = True
410           lUnits.SetFocus
420       End If

430       Exit Sub

grdBio_MouseUp_Error:

          Dim strES As String
          Dim intEL As Integer

440       intEL = Erl
450       strES = Err.Description
460       LogError "frmBioPODPU", "grdBio_MouseUp", intEL, strES


End Sub

Private Sub lUnits_Click()

10        On Error GoTo lUnits_Click_Error

20        grdBio = lUnits
30        cmdSave.Visible = True
40        lUnits.Visible = False

50        cmdMoveUp.Enabled = True
60        cmdMoveDown.Enabled = True
70        cmdPrint.Enabled = True
80        cmdSave.Enabled = True
90        cmdCancel.Enabled = True
100       grdBio.Enabled = True

110       Exit Sub

lUnits_Click_Error:

          Dim strES As String
          Dim intEL As Integer

120       intEL = Erl
130       strES = Err.Description
140       LogError "frmBioPODPU", "lUnits_Click", intEL, strES


End Sub

Private Sub lUnits_KeyPress(KeyAscii As Integer)

10        On Error GoTo lUnits_KeyPress_Error

20        KeyAscii = 0

30        Exit Sub

lUnits_KeyPress_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmBioPODPU", "lUnits_KeyPress", intEL, strES


End Sub

Private Sub SaveG()

          Dim sql As String
          Dim Y As Long

10        On Error GoTo SaveG_Error

20        For Y = 1 To grdBio.Rows - 1
30            sql = "Update BioTestDefinitions " & _
                    "Set DP = '" & Val(grdBio.TextMatrix(Y, 1)) & "', " & _
                    "Units = '" & grdBio.TextMatrix(Y, 2) & "', " & _
                    "PrintPriority = '" & Y & "' " & _
                    "where ShortName = '" & grdBio.TextMatrix(Y, 0) & "'"
40            Cnxn(0).Execute sql
50        Next

60        cmdSave.Visible = False

70        Exit Sub

SaveG_Error:

          Dim strES As String
          Dim intEL As Integer

80        intEL = Erl
90        strES = Err.Description
100       LogError "frmBioPODPU", "SaveG", intEL, strES, sql


End Sub

Private Sub tmrDown_Timer()

10        On Error GoTo tmrDown_Timer_Error

20        FireDown

30        Exit Sub

tmrDown_Timer_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmBioPODPU", "tmrDown_Timer", intEL, strES


End Sub

Private Sub tmrUp_Timer()

10        On Error GoTo tmrUp_Timer_Error

20        FireUp

30        Exit Sub

tmrUp_Timer_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmBioPODPU", "tmrUp_Timer", intEL, strES


End Sub
