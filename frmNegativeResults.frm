VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmNegativeResults 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire"
   ClientHeight    =   8280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8280
   ScaleWidth      =   8205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   825
      Left            =   7140
      Picture         =   "frmNegativeResults.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5550
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.CommandButton cmdXL 
      Caption         =   "Export to Excel"
      Height          =   825
      Left            =   7140
      Picture         =   "frmNegativeResults.frx":1982
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4680
      Width           =   795
   End
   Begin VB.Timer tmrUp 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   7620
      Top             =   2370
   End
   Begin VB.Timer tmrDown 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   7620
      Top             =   3150
   End
   Begin VB.CommandButton cmdMoveDown 
      Caption         =   "Move Down"
      Height          =   1035
      Left            =   7020
      Picture         =   "frmNegativeResults.frx":1C8C
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3000
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.CommandButton cmdMoveUp 
      Caption         =   "Move Up"
      Height          =   1035
      Left            =   7020
      Picture         =   "frmNegativeResults.frx":360E
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1950
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.CommandButton bcancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   825
      Left            =   7140
      Picture         =   "frmNegativeResults.frx":4F90
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7290
      Width           =   795
   End
   Begin VB.Frame FrameAdd 
      Caption         =   "Add New "
      Height          =   1095
      Left            =   210
      TabIndex        =   7
      Top             =   690
      Width           =   7875
      Begin VB.ComboBox cmbText 
         Height          =   315
         Left            =   870
         TabIndex        =   1
         Text            =   "Combo1"
         Top             =   510
         Width           =   5895
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
         Height          =   315
         Left            =   120
         MaxLength       =   5
         TabIndex        =   0
         Top             =   510
         Width           =   705
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   825
         Left            =   6900
         Picture         =   "frmNegativeResults.frx":6912
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   210
         Width           =   795
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Code"
         Height          =   195
         Left            =   270
         TabIndex        =   9
         Top             =   300
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Text"
         Height          =   195
         Left            =   930
         TabIndex        =   8
         Top             =   300
         Width           =   315
      End
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H0000FFFF&
      Caption         =   "Save"
      Height          =   825
      Left            =   7140
      Picture         =   "frmNegativeResults.frx":8294
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6420
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.ComboBox cmbSite 
      Height          =   315
      Left            =   1050
      TabIndex        =   4
      Text            =   "cmbSite"
      Top             =   240
      Width           =   3045
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   6285
      Left            =   270
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1830
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   11086
      _Version        =   393216
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
      FormatString    =   "<Code   |Text                                                                                      "
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
   Begin VB.Label lblExcelInfo 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exporting"
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
      Left            =   7050
      TabIndex        =   14
      Top             =   4290
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Site"
      Height          =   195
      Left            =   720
      TabIndex        =   5
      Top             =   300
      Width           =   270
   End
End
Attribute VB_Name = "frmNegativeResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FireCounter As Integer

Private Sub FillG()

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo FillG_Error

20        g.Rows = 2
30        g.AddItem ""
40        g.RemoveItem 1

50        cmdDelete.Visible = False

60        sql = "SELECT Code, Name FROM Organisms WHERE " & _
                "GroupName = 'Negative Results' " & _
                "AND Site = '" & cmbSite & "' " & _
                "ORDER BY ListOrder"
70        Set tb = New Recordset
80        RecOpenServer 0, tb, sql
90        Do While Not tb.EOF
100           g.AddItem tb!Code & vbTab & tb!Name & ""
110           tb.MoveNext
120       Loop

130       If g.Rows > 2 Then
140           g.RemoveItem 1
150       End If

160       cmdSave.Visible = False

170       Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer

180       intEL = Erl
190       strES = Err.Description
200       LogError "frmNegativeResults", "FillG", intEL, strES, sql

End Sub

Private Sub FillCombos()

          Dim sql As String
          Dim tb As Recordset

10        On Error GoTo FillCombos_Error

20        cmbSite.Clear
30        cmbText.Clear

40        sql = "SELECT Text FROM Lists WHERE " & _
                "ListType = 'SI' " & _
                "ORDER BY ListOrder"
50        Set tb = New Recordset
60        RecOpenServer 0, tb, sql
70        Do While Not tb.EOF
80            cmbSite.AddItem tb!Text & ""
90            tb.MoveNext
100       Loop

110       sql = "SELECT DISTINCT Name FROM Organisms WHERE " & _
                "GroupName = 'Negative Results' " & _
                "ORDER BY Name"
120       Set tb = New Recordset
130       RecOpenServer 0, tb, sql
140       Do While Not tb.EOF
150           cmbText.AddItem tb!Name & ""
160           tb.MoveNext
170       Loop

180       Exit Sub

FillCombos_Error:

          Dim strES As String
          Dim intEL As Integer

190       intEL = Erl
200       strES = Err.Description
210       LogError "frmNegativeResults", "FillCombos", intEL, strES, sql

End Sub

Private Sub FireDown()

          Dim n As Integer
          Dim s As String
          Dim X As Integer
          Dim VisibleRows As Integer

10        On Error GoTo FireDown_Error

20        cmdDelete.Visible = False

30        If g.Row = g.Rows - 1 Then Exit Sub
40        n = g.Row

50        FireCounter = FireCounter + 1
60        If FireCounter > 5 Then
70            tmrDown.Interval = 100
80        End If

90        VisibleRows = g.Height \ g.RowHeight(1) - 1

100       g.Visible = False

110       s = ""
120       For X = 0 To g.Cols - 1
130           s = s & g.TextMatrix(n, X) & vbTab
140       Next
150       s = Left$(s, Len(s) - 1)

160       g.RemoveItem n
170       If n < g.Rows Then
180           g.AddItem s, n + 1
190           g.Row = n + 1
200       Else
210           g.AddItem s
220           g.Row = g.Rows - 1
230       End If

240       For X = 0 To g.Cols - 1
250           g.Col = X
260           g.CellBackColor = vbYellow
270       Next

280       If Not g.RowIsVisible(g.Row) Or g.Row = g.Rows - 1 Then
290           If g.Row - VisibleRows + 1 > 0 Then
300               g.TopRow = g.Row - VisibleRows + 1
310           End If
320       End If

330       g.Visible = True

340       cmdSave.Visible = True

350       Exit Sub

FireDown_Error:

          Dim strES As String
          Dim intEL As Integer

360       intEL = Erl
370       strES = Err.Description
380       LogError "frmNegativeResults", "FireDown", intEL, strES

End Sub

Private Sub FireUp()

          Dim n As Integer
          Dim s As String
          Dim X As Integer

10        On Error GoTo FireUp_Error

20        cmdDelete.Visible = False

30        If g.Row = 1 Then Exit Sub

40        FireCounter = FireCounter + 1
50        If FireCounter > 5 Then
60            tmrUp.Interval = 100
70        End If

80        n = g.Row

90        g.Visible = False

100       s = ""
110       For X = 0 To g.Cols - 1
120           s = s & g.TextMatrix(n, X) & vbTab
130       Next
140       s = Left$(s, Len(s) - 1)

150       g.RemoveItem n
160       g.AddItem s, n - 1

170       g.Row = n - 1
180       For X = 0 To g.Cols - 1
190           g.Col = X
200           g.CellBackColor = vbYellow
210       Next

220       If Not g.RowIsVisible(g.Row) Then
230           g.TopRow = g.Row
240       End If

250       g.Visible = True

260       cmdSave.Visible = True

270       Exit Sub

FireUp_Error:

          Dim strES As String
          Dim intEL As Integer

280       intEL = Erl
290       strES = Err.Description
300       LogError "frmNegativeResults", "FireUp", intEL, strES

End Sub




Private Sub cmbSite_Click()

10        FillG

End Sub

Private Sub cmdadd_Click()

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo cmdadd_Click_Error

20        cmdDelete.Visible = False

30        txtCode = Trim$(UCase$(txtCode))
40        cmbText = Trim$(cmbText)

50        If txtCode = "" Then
60            Exit Sub
70        End If

80        If cmbText = "" Then Exit Sub

90        sql = "SELECT * FROM Organisms WHERE " & _
                "GroupName = 'Negative Results' " & _
                "AND Site = '" & cmbSite & "' " & _
                "AND Code = '" & txtCode & "'"
100       Set tb = New Recordset
110       RecOpenServer 0, tb, sql
120       If tb.EOF Then
130           tb.AddNew
140       End If
150       tb!Name = cmbText
160       tb!GroupName = "Negative Results"
170       tb!ListOrder = 999
180       tb!Code = txtCode
190       tb!Site = cmbSite
200       tb.Update

210       txtCode = ""
220       cmbText = ""

230       FillG

240       txtCode.SetFocus

250       Exit Sub

cmdadd_Click_Error:

          Dim strES As String
          Dim intEL As Integer

260       intEL = Erl
270       strES = Err.Description
280       LogError "frmNegativeResults", "cmdAdd_Click", intEL, strES, sql

End Sub


Private Sub bcancel_Click()

10        Unload Me

End Sub


Private Sub cmbSite_KeyPress(KeyAscii As Integer)

10        KeyAscii = 0

End Sub


Private Sub cmdDelete_Click()

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo cmdDelete_Click_Error

20        cmdDelete.Visible = False
30        If g.CellBackColor <> vbYellow Then Exit Sub

40        sql = "SELECT * FROM Organisms WHERE " & _
                "GroupName = 'Negative Results' " & _
                "AND Site = '" & cmbSite & "' " & _
                "AND Code = '" & g.TextMatrix(g.Row, 0) & "' " & _
                "AND Name = '" & g.TextMatrix(g.Row, 1) & "'"
50        Set tb = New Recordset
60        RecOpenServer 0, tb, sql
70        If Not tb.EOF Then
80            tb.DELETE
90        End If

100       FillG
110       cmdMoveDown.Visible = False
120       cmdMoveUp.Visible = False

130       Exit Sub

cmdDelete_Click_Error:

          Dim strES As String
          Dim intEL As Integer

140       intEL = Erl
150       strES = Err.Description
160       LogError "frmNegativeResults", "cmdDelete_Click", intEL, strES, sql

End Sub

Private Sub cmdMoveDown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        FireDown

20        tmrDown.Interval = 250
30        FireCounter = 0

40        tmrDown.Enabled = True
50        cmdSave.Visible = True

End Sub


Private Sub cmdMoveDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        tmrDown.Enabled = False

End Sub


Private Sub cmdMoveUp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        FireUp

20        tmrUp.Interval = 250
30        FireCounter = 0

40        tmrUp.Enabled = True
50        cmdSave.Visible = True

End Sub


Private Sub cmdMoveUp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        tmrUp.Enabled = False

End Sub


Private Sub cmdSave_Click()

          Dim sql As String
          Dim n As Integer

10        On Error GoTo cmdSave_Click_Error

20        For n = 1 To g.Rows - 1
30            sql = "UPDATE Organisms " & _
                    "SET ListOrder = " & n & " WHERE " & _
                    "GroupName = 'Negative Results' " & _
                    "AND Site = '" & cmbSite & "' " & _
                    "AND Name = '" & g.TextMatrix(n, 1) & "'"
40            Cnxn(0).Execute sql
50        Next

60        cmdSave.Visible = False

70        Exit Sub

cmdSave_Click_Error:

          Dim strES As String
          Dim intEL As Integer

80        intEL = Erl
90        strES = Err.Description
100       LogError "frmNegativeResults", "cmdSave_Click", intEL, strES, sql

End Sub

Private Sub cmdXL_Click()

10        On Error GoTo cmdXL_Click_Error

20        ExportFlexGrid g, Me

30        Exit Sub

cmdXL_Click_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmNegativeResults", "cmdXL_Click", intEL, strES

End Sub


Private Sub Form_Activate()

10        FillG

End Sub

Private Sub Form_Load()

10        FillCombos

End Sub


Private Sub g_Click()

          Static SortOrder As Boolean
          Dim X As Integer
          Dim Y As Integer
          Dim ySave As Integer

10        ySave = g.Row

20        g.Visible = False
30        g.Col = 0
40        For Y = 1 To g.Rows - 1
50            g.Row = Y
60            If g.CellBackColor = vbYellow Then
70                For X = 0 To g.Cols - 1
80                    g.Col = X
90                    g.CellBackColor = 0
100               Next
110               Exit For
120           End If
130       Next
140       g.Row = ySave
150       g.Visible = True

160       If g.MouseRow = 0 Then
170           If SortOrder Then
180               g.Sort = flexSortGenericAscending
190           Else
200               g.Sort = flexSortGenericDescending
210           End If
220           SortOrder = Not SortOrder
230           Exit Sub
240       End If

250       For X = 0 To g.Cols - 1
260           g.Col = X
270           g.CellBackColor = vbYellow
280       Next

290       cmdMoveUp.Visible = True
300       cmdMoveDown.Visible = True
310       cmdDelete.Visible = True

End Sub


Private Sub tmrDown_Timer()

10        FireDown

End Sub


Private Sub tmrUp_Timer()

10        FireUp

End Sub


