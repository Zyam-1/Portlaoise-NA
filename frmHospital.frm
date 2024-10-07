VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmHospital 
   Caption         =   "NetAcquire - Hospitals"
   ClientHeight    =   7050
   ClientLeft      =   1515
   ClientTop       =   1095
   ClientWidth     =   7245
   Icon            =   "frmHospital.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   7245
   Begin VB.CommandButton bprint 
      Appearance      =   0  'Flat
      Caption         =   "&Print"
      Height          =   825
      Left            =   6120
      Picture         =   "frmHospital.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5040
      Width           =   795
   End
   Begin VB.CommandButton bcancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   825
      Left            =   6120
      Picture         =   "frmHospital.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5940
      Width           =   795
   End
   Begin VB.CommandButton bMoveUp 
      Caption         =   "Move Up"
      Enabled         =   0   'False
      Height          =   825
      Left            =   6060
      Picture         =   "frmHospital.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2910
      Width           =   795
   End
   Begin VB.CommandButton bMoveDown 
      Caption         =   "Move Down"
      Enabled         =   0   'False
      Height          =   825
      Left            =   6060
      Picture         =   "frmHospital.frx":0D60
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3780
      Width           =   795
   End
   Begin VB.CommandButton bSave 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   825
      Left            =   6075
      Picture         =   "frmHospital.frx":11A2
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1350
      Width           =   795
   End
   Begin VB.Frame Frame2 
      Caption         =   "Add New Hospital"
      Height          =   1065
      Left            =   90
      TabIndex        =   8
      Top             =   150
      Width           =   5925
      Begin VB.TextBox tCode 
         Height          =   285
         Left            =   810
         MaxLength       =   5
         TabIndex        =   0
         Top             =   240
         Width           =   675
      End
      Begin VB.TextBox tText 
         Height          =   285
         Left            =   810
         MaxLength       =   30
         TabIndex        =   1
         Top             =   600
         Width           =   3645
      End
      Begin VB.CommandButton bAdd 
         Caption         =   "Add"
         Height          =   855
         Left            =   4860
         Picture         =   "frmHospital.frx":14AC
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   150
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Code"
         Height          =   195
         Left            =   420
         TabIndex        =   10
         Top             =   270
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         Height          =   195
         Left            =   390
         TabIndex        =   9
         Top             =   630
         Width           =   420
      End
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   5445
      Left            =   90
      TabIndex        =   11
      Top             =   1350
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   9604
      _Version        =   393216
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
      AllowUserResizing=   1
      FormatString    =   "<Code       |<Text                                                                     "
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
Attribute VB_Name = "frmHospital"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Activated As Boolean

Private Sub FillG()

          Dim s As String
          Dim tb As New Recordset
          Dim sql As String

10        On Error GoTo FillG_Error

20        ClearFGrid g



30        sql = "SELECT * from lists WHERE listtype = 'HO'"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        Do While Not tb.EOF
70            s = Trim(tb!Code) & vbTab & Trim(tb!Text)
80            g.AddItem s
90            tb.MoveNext
100       Loop

110       FixG g

120       Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer



130       intEL = Erl
140       strES = Err.Description
150       LogError "frmHospital", "FillG", intEL, strES, sql


End Sub

Private Sub bAdd_Click()

10        On Error GoTo bAdd_Click_Error

20        tCode = Trim$(UCase$(tCode))
30        tText = Trim$(tText)

40        If tCode = "" Then
50            Exit Sub
60        End If

70        If tText = "" Then Exit Sub

80        g.AddItem tCode & vbTab & tText

90        tCode = ""
100       tText = ""

110       If g.Rows > 2 And g.TextMatrix(1, 0) = "" Then g.RemoveItem 1

120       bsave.Enabled = True

130       Exit Sub

bAdd_Click_Error:

          Dim strES As String
          Dim intEL As Integer



140       intEL = Erl
150       strES = Err.Description
160       LogError "frmHospital", "bAdd_Click", intEL, strES


End Sub


Private Sub bcancel_Click()

10        Unload Me

End Sub

Private Sub bMoveDown_Click()

          Dim n As Long
          Dim s As String
          Dim X As Long

10        On Error GoTo bMoveDown_Click_Error

20        If g.Row = g.Rows - 1 Then Exit Sub
30        n = g.Row

40        s = ""
50        For X = 0 To g.Cols - 1
60            s = s & g.TextMatrix(n, X) & vbTab
70        Next
80        s = Left$(s, Len(s) - 1)

90        g.RemoveItem n
100       If n < g.Rows Then
110           g.AddItem s, n + 1
120           g.Row = n + 1
130       Else
140           g.AddItem s
150           g.Row = g.Rows - 1
160       End If

170       For X = 0 To g.Cols - 1
180           g.Col = X
190           g.CellBackColor = vbYellow
200       Next

210       bsave.Enabled = True

220       Exit Sub

bMoveDown_Click_Error:

          Dim strES As String
          Dim intEL As Integer



230       intEL = Erl
240       strES = Err.Description
250       LogError "frmHospital", "bMoveDown_Click", intEL, strES


End Sub


Private Sub bMoveUp_Click()

          Dim n As Long
          Dim s As String
          Dim X As Long

10        On Error GoTo bMoveUp_Click_Error

20        If g.Row = 1 Then Exit Sub

30        n = g.Row

40        s = ""
50        For X = 0 To g.Cols - 1
60            s = s & g.TextMatrix(n, X) & vbTab
70        Next
80        s = Left$(s, Len(s) - 1)

90        g.RemoveItem n
100       g.AddItem s, n - 1

110       g.Row = n - 1
120       For X = 0 To g.Cols - 1
130           g.Col = X
140           g.CellBackColor = vbYellow
150       Next

160       bsave.Enabled = True

170       Exit Sub

bMoveUp_Click_Error:

          Dim strES As String
          Dim intEL As Integer



180       intEL = Erl
190       strES = Err.Description
200       LogError "frmHospital", "bMoveUp_Click", intEL, strES


End Sub


Private Sub bprint_Click()

10        On Error GoTo bprint_Click_Error

20        Printer.Print

30        Printer.Print "List of Hospitals"

40        g.Col = 0
50        g.Row = 1
60        g.ColSel = g.Cols - 1
70        g.RowSel = g.Rows - 1

80        Printer.Print g.Clip

90        Printer.EndDoc


100       Exit Sub

bprint_Click_Error:

          Dim strES As String
          Dim intEL As Integer



110       intEL = Erl
120       strES = Err.Description
130       LogError "frmHospital", "bPrint_Click", intEL, strES


End Sub


Private Sub bSave_Click()

          Dim Y As Long
          Dim sql As String
          Dim tb As New Recordset

10        On Error GoTo bSave_Click_Error

20        For Y = 1 To g.Rows - 1

30            sql = "SELECT * from lists WHERE listtype = 'HO' and code = '" & g.TextMatrix(Y, 0) & "'"
40            Set tb = New Recordset
50            RecOpenServer 0, tb, sql
60            If tb.EOF Then tb.AddNew
70            tb!Code = g.TextMatrix(Y, 0)
80            tb!ListType = "HO"
90            tb!Text = g.TextMatrix(Y, 1)
100           tb!ListOrder = Y
110           tb!InUse = 1
120           tb.Update
130       Next

140       FillG

150       tCode = ""
160       tText = ""
170       tCode.SetFocus
180       bMoveUp.Enabled = False
190       bMoveDown.Enabled = False
200       bsave.Enabled = False

210       Exit Sub

bSave_Click_Error:

          Dim strES As String
          Dim intEL As Integer



220       intEL = Erl
230       strES = Err.Description
240       LogError "frmHospital", "bSave_Click", intEL, strES, sql


End Sub


Private Sub Form_Activate()

10        On Error GoTo Form_Activate_Error

20        If Activated Then Exit Sub

30        Activated = True

40        FillG

50        Set_Font Me

60        Exit Sub

Form_Activate_Error:

          Dim strES As String
          Dim intEL As Integer



70        intEL = Erl
80        strES = Err.Description
90        LogError "frmHospital", "Form_Activate", intEL, strES


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
70        LogError "frmHospital", "Form_Load", intEL, strES


End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

10        On Error GoTo Form_QueryUnload_Error

20        If bsave.Enabled Then
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
110       LogError "frmHospital", "Form_QueryUnload", intEL, strES


End Sub


Private Sub g_Click()

          Static SortOrder As Boolean
          Dim X As Long
          Dim Y As Long
          Dim ySave As Long

10        On Error GoTo g_Click_Error

20        ySave = g.Row

30        g.Visible = False
40        g.Col = 0
50        For Y = 1 To g.Rows - 1
60            g.Row = Y
70            If g.CellBackColor = vbYellow Then
80                For X = 0 To g.Cols - 1
90                    g.Col = X
100                   g.CellBackColor = 0
110               Next
120               Exit For
130           End If
140       Next
150       g.Row = ySave
160       g.Visible = True

170       If g.MouseRow = 0 Then
180           If SortOrder Then
190               g.Sort = flexSortGenericAscending
200           Else
210               g.Sort = flexSortGenericDescending
220           End If
230           SortOrder = Not SortOrder
240           Exit Sub
250       End If

260       For X = 0 To g.Cols - 1
270           g.Col = X
280           g.CellBackColor = vbYellow
290       Next

300       bMoveUp.Enabled = True
310       bMoveDown.Enabled = True

320       Exit Sub

g_Click_Error:

          Dim strES As String
          Dim intEL As Integer



330       intEL = Erl
340       strES = Err.Description
350       LogError "frmHospital", "g_Click", intEL, strES


End Sub


