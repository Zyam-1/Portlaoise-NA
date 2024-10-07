VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmMicroSites 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Microbiology - Sites"
   ClientHeight    =   8460
   ClientLeft      =   240
   ClientTop       =   525
   ClientWidth     =   7815
   Icon            =   "frmMicroSites.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8460
   ScaleWidth      =   7815
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrDown 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   7410
      Top             =   5730
   End
   Begin VB.Timer tmrUp 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   7410
      Top             =   4890
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   825
      Left            =   6930
      Picture         =   "frmMicroSites.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   495
      Width           =   795
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H0000FFFF&
      Caption         =   "Save"
      Height          =   825
      Left            =   6930
      Picture         =   "frmMicroSites.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1575
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.CommandButton cmdMoveDown 
      Caption         =   "Move Down"
      Enabled         =   0   'False
      Height          =   825
      Left            =   6930
      Picture         =   "frmMicroSites.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3900
      Width           =   795
   End
   Begin VB.CommandButton cmdMoveUp 
      Caption         =   "Move Up"
      Enabled         =   0   'False
      Height          =   825
      Left            =   6930
      Picture         =   "frmMicroSites.frx":0D60
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3060
      Width           =   795
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   825
      Left            =   6930
      Picture         =   "frmMicroSites.frx":11A2
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7425
      Width           =   795
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   825
      Left            =   6930
      Picture         =   "frmMicroSites.frx":14AC
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6525
      Width           =   795
   End
   Begin VB.Frame FrameAdd 
      Caption         =   "Add New Site"
      Height          =   1365
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   6705
      Begin VB.OptionButton optABs 
         Alignment       =   1  'Right Justify
         Caption         =   "4"
         Height          =   255
         Index           =   3
         Left            =   4980
         TabIndex        =   17
         Top             =   810
         Width           =   405
      End
      Begin VB.OptionButton optABs 
         Alignment       =   1  'Right Justify
         Caption         =   "3"
         Height          =   255
         Index           =   2
         Left            =   4350
         TabIndex        =   16
         Top             =   810
         Value           =   -1  'True
         Width           =   405
      End
      Begin VB.OptionButton optABs 
         Alignment       =   1  'Right Justify
         Caption         =   "2"
         Height          =   255
         Index           =   1
         Left            =   3720
         TabIndex        =   15
         Top             =   810
         Width           =   345
      End
      Begin VB.OptionButton optABs 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         Height          =   255
         Index           =   0
         Left            =   3120
         TabIndex        =   14
         Top             =   810
         Width           =   345
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   870
         Left            =   5715
         Picture         =   "frmMicroSites.frx":17B6
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtText 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1950
         MaxLength       =   50
         TabIndex        =   2
         Top             =   450
         Width           =   3495
      End
      Begin VB.TextBox txtCode 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   690
         MaxLength       =   5
         TabIndex        =   1
         Top             =   420
         Width           =   705
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Number of Antibiotics to Report"
         Height          =   195
         Left            =   720
         TabIndex        =   13
         Top             =   840
         Width           =   2235
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Text"
         Height          =   195
         Left            =   1980
         TabIndex        =   5
         Top             =   240
         Width           =   315
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Code"
         Height          =   195
         Left            =   720
         TabIndex        =   4
         Top             =   210
         Width           =   375
      End
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   6795
      Left            =   120
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1560
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   11986
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      FormatString    =   "<Code   |Text                                                                             |^AB's "
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
Attribute VB_Name = "frmMicroSites"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Activated As Boolean

Private FireCounter As Long

Private Sub FireDown()

          Dim n As Long
          Dim s As String
          Dim X As Long
          Dim VisibleRows As Long

10        On Error GoTo FireDown_Error

20        If g.Row = g.Rows - 1 Then Exit Sub
30        n = g.Row

40        FireCounter = FireCounter + 1
50        If FireCounter > 5 Then
60            tmrDown.Interval = 100
70        End If

80        VisibleRows = g.Height \ g.RowHeight(1) - 1

90        g.Visible = False

100       s = ""
110       For X = 0 To g.Cols - 1
120           s = s & g.TextMatrix(n, X) & vbTab
130       Next
140       s = Left$(s, Len(s) - 1)

150       g.RemoveItem n
160       If n < g.Rows Then
170           g.AddItem s, n + 1
180           g.Row = n + 1
190       Else
200           g.AddItem s
210           g.Row = g.Rows - 1
220       End If

230       For X = 0 To g.Cols - 1
240           g.Col = X
250           g.CellBackColor = vbYellow
260       Next

270       If Not g.RowIsVisible(g.Row) Or g.Row = g.Rows - 1 Then
280           If g.Row - VisibleRows + 1 > 0 Then
290               g.TopRow = g.Row - VisibleRows + 1
300           End If
310       End If

320       g.Visible = True

330       cmdSave.Visible = True

340       Exit Sub

FireDown_Error:

          Dim strES As String
          Dim intEL As Integer



350       intEL = Erl
360       strES = Err.Description
370       LogError "frmMicroSites", "FireDown", intEL, strES


End Sub

Private Sub FireUp()

          Dim n As Long
          Dim s As String
          Dim X As Long

10        On Error GoTo FireUp_Error

20        If g.Row = 1 Then Exit Sub

30        FireCounter = FireCounter + 1
40        If FireCounter > 5 Then
50            tmrUp.Interval = 100
60        End If

70        n = g.Row

80        g.Visible = False

90        s = ""
100       For X = 0 To g.Cols - 1
110           s = s & g.TextMatrix(n, X) & vbTab
120       Next
130       s = Left$(s, Len(s) - 1)

140       g.RemoveItem n
150       g.AddItem s, n - 1

160       g.Row = n - 1
170       For X = 0 To g.Cols - 1
180           g.Col = X
190           g.CellBackColor = vbYellow
200       Next

210       If Not g.RowIsVisible(g.Row) Then
220           g.TopRow = g.Row
230       End If

240       g.Visible = True

250       cmdSave.Visible = True

260       Exit Sub

FireUp_Error:

          Dim strES As String
          Dim intEL As Integer



270       intEL = Erl
280       strES = Err.Description
290       LogError "frmMicroSites", "FireUp", intEL, strES


End Sub



Private Sub FillG()

          Dim tb As Recordset
          Dim sql As String
          Dim s As String

10        On Error GoTo FillG_Error

20        g.Rows = 2
30        g.AddItem ""
40        g.RemoveItem 1

50        sql = "SELECT * from Lists WHERE " & _
                "ListType = 'SI' " & _
                "order by ListOrder"
60        Set tb = New Recordset
70        RecOpenServer 0, tb, sql
80        Do While Not tb.EOF
90            s = tb!Code & vbTab & tb!Text & vbTab & tb!Default & ""
100           g.AddItem s
110           tb.MoveNext
120       Loop

130       If g.Rows > 2 Then
140           g.RemoveItem 1
150       End If

160       Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer



170       intEL = Erl
180       strES = Err.Description
190       LogError "frmMicroSites", "FillG", intEL, strES, sql


End Sub


Private Sub cmdadd_Click()

          Dim n As Long
          Dim DefaultABs As String

10        On Error GoTo cmdadd_Click_Error

20        txtCode = Trim$(UCase$(txtCode))
30        txtText = Trim$(txtText)

40        If txtCode = "" Then Exit Sub
50        If txtText = "" Then Exit Sub

60        For n = 0 To 3
70            If optABs(n) Then
80                DefaultABs = Format$(n + 1)
90            End If
100       Next

110       g.AddItem txtCode & vbTab & txtText & vbTab & DefaultABs

120       txtCode = ""
130       txtText = ""
140       If SysOptDefaultABs(0) > 0 Then
150           optABs(SysOptDefaultABs(0) - 1) = True
160       End If
170       If g.Rows > 2 And g.TextMatrix(1, 0) = "" Then g.RemoveItem 1

180       cmdSave.Visible = True

190       Exit Sub

cmdadd_Click_Error:

          Dim strES As String
          Dim intEL As Integer



200       intEL = Erl
210       strES = Err.Description
220       LogError "frmMicroSites", "cmdAdd_Click", intEL, strES


End Sub


Private Sub cmdCancel_Click()

10        Unload Me

End Sub


Private Sub cmdDelete_Click()

          Dim sql As String

10        On Error GoTo cmdDelete_Click_Error

20        sql = "DELETE from Lists WHERE " & _
                "ListType = 'SI' " & _
                "and Code = '" & g.TextMatrix(g.Row, 0) & "'"
30        Cnxn(0).Execute sql

40        FillG

50        Exit Sub

cmdDelete_Click_Error:

          Dim strES As String
          Dim intEL As Integer



60        intEL = Erl
70        strES = Err.Description
80        LogError "frmMicroSites", "cmdDelete_Click", intEL, strES, sql


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
90        LogError "frmMicroSites", "cmdMoveDown_MouseDown", intEL, strES


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
60        LogError "frmMicroSites", "cmdMoveDown_MouseUp", intEL, strES


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
90        LogError "frmMicroSites", "cmdMoveUp_MouseDown", intEL, strES


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
60        LogError "frmMicroSites", "cmdMoveUp_MouseUp", intEL, strES


End Sub


Private Sub cmdPrint_Click()

10        On Error GoTo cmdPrint_Click_Error

20        Printer.Print

30        Printer.Print "List of Sites."

40        g.Col = 0
50        g.Row = 1
60        g.ColSel = g.Cols - 1
70        g.RowSel = g.Rows - 1

80        Printer.Print g.Clip

90        Printer.EndDoc


100       Exit Sub

cmdPrint_Click_Error:

          Dim strES As String
          Dim intEL As Integer



110       intEL = Erl
120       strES = Err.Description
130       LogError "frmMicroSites", "cmdPrint_Click", intEL, strES


End Sub


Private Sub cmdSave_Click()

          Dim tb As Recordset
          Dim sql As String
          Dim Y As Long

10        On Error GoTo cmdSave_Click_Error

20        For Y = 1 To g.Rows - 1
30            sql = "SELECT * from Lists WHERE " & _
                    "ListType = 'SI' " & _
                    "and Code = '" & g.TextMatrix(Y, 0) & "'"
40            Set tb = New Recordset
50            RecOpenServer 0, tb, sql
60            If tb.EOF Then
70                tb.AddNew
80            End If
90            tb!Code = g.TextMatrix(Y, 0)
100           tb!ListType = "SI"
110           tb!Text = g.TextMatrix(Y, 1)
120           tb!Default = g.TextMatrix(Y, 2)
130           tb!ListOrder = Y
140           tb!InUse = 1
150           tb.Update
160       Next

170       FillG

180       txtCode = ""
190       txtText = ""
200       If txtCode.Visible And txtCode.Enabled Then txtCode.SetFocus
210       cmdMoveUp.Enabled = False
220       cmdMoveDown.Enabled = False
230       cmdDelete.Enabled = False
240       cmdSave.Visible = False

250       Exit Sub

cmdSave_Click_Error:

          Dim strES As String
          Dim intEL As Integer



260       intEL = Erl
270       strES = Err.Description
280       LogError "frmMicroSites", "cmdsave_Click", intEL, strES, sql


End Sub


Private Sub Form_Activate()

10        On Error GoTo Form_Activate_Error

20        If Activated Then Exit Sub

30        Activated = True

40        FillG

50        Exit Sub

Form_Activate_Error:

          Dim strES As String
          Dim intEL As Integer



60        intEL = Erl
70        strES = Err.Description
80        LogError "frmMicroSites", "Form_Activate", intEL, strES


End Sub

Private Sub Form_Load()

10        On Error GoTo Form_Load_Error

20        g.Font.Bold = True

30        Activated = False

40        Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer



50        intEL = Erl
60        strES = Err.Description
70        LogError "frmMicroSites", "Form_Load", intEL, strES


End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

10        On Error GoTo Form_QueryUnload_Error

20        If cmdSave.Visible Then
30            If iMsg("Cancel without Saving?", vbQuestion + vbYesNo) = vbNo Then
40                Cancel = True
50                Exit Sub
60            End If
70        End If

80        Exit Sub

Form_QueryUnload_Error:

          Dim strES As String
          Dim intEL As Integer



90        intEL = Erl
100       strES = Err.Description
110       LogError "frmMicroSites", "Form_QueryUnload", intEL, strES


End Sub


Private Sub g_Click()

          Static SortOrder As Boolean
          Dim X As Long
          Dim Y As Long
          Dim ySave As Long

10        On Error GoTo g_Click_Error

20        If g.MouseRow = 0 Then
30            If SortOrder Then
40                g.Sort = flexSortGenericAscending
50            Else
60                g.Sort = flexSortGenericDescending
70            End If
80            SortOrder = Not SortOrder
90            Exit Sub
100       End If

110       ySave = g.Row

120       If g.Col = 2 Then
130           g.Enabled = False
140           g = iBOX("Number of Antibiotics to Report?", , SysOptDefaultABs(0))
150           g = Val(g)
160           g.Enabled = True
170           cmdSave.Visible = True
180           Exit Sub
190       End If

200       g.Visible = False
210       g.Col = 0
220       For Y = 1 To g.Rows - 1
230           g.Row = Y
240           If g.CellBackColor = vbYellow Then
250               For X = 0 To g.Cols - 1
260                   g.Col = X
270                   g.CellBackColor = 0
280               Next
290               Exit For
300           End If
310       Next
320       g.Row = ySave
330       g.Visible = True

340       For X = 0 To g.Cols - 1
350           g.Col = X
360           g.CellBackColor = vbYellow
370       Next

380       cmdMoveUp.Enabled = True
390       cmdMoveDown.Enabled = True
400       cmdDelete.Enabled = True

410       Exit Sub

g_Click_Error:

          Dim strES As String
          Dim intEL As Integer



420       intEL = Erl
430       strES = Err.Description
440       LogError "frmMicroSites", "g_Click", intEL, strES


End Sub


Private Sub optABs_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

          Dim sql As String

10        On Error GoTo optABs_MouseUp_Error

20        If SysOptDefaultABs(0) <> Index + 1 Then
30            If iMsg("Do you want to reset the Default to " & Format(Index + 1) & " ?", vbQuestion + vbYesNo) = vbYes Then
40                sql = "UPDATE Options " & _
                        "Set Contents = '" & Index + 1 & "' WHERE Description = 'DefaultABs'"
50                Cnxn(0).Execute sql
60                SysOptDefaultABs(0) = Index + 1
70            End If
80        End If

90        Exit Sub

optABs_MouseUp_Error:

          Dim strES As String
          Dim intEL As Integer



100       intEL = Erl
110       strES = Err.Description
120       LogError "frmMicroSites", "optABs_MouseUp", intEL, strES, sql


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
60        LogError "frmMicroSites", "tmrDown_Timer", intEL, strES


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
60        LogError "frmMicroSites", "tmrUp_Timer", intEL, strES


End Sub


