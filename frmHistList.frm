VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmHistList 
   Caption         =   "Netacquire - Histology Lists"
   ClientHeight    =   8550
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9075
   Icon            =   "frmHistList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8550
   ScaleWidth      =   9075
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton bprint 
      Appearance      =   0  'Flat
      Caption         =   "&Print"
      Height          =   825
      Left            =   6930
      Picture         =   "frmHistList.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6435
      Width           =   795
   End
   Begin VB.CommandButton bcancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   825
      Left            =   6930
      Picture         =   "frmHistList.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7425
      Width           =   795
   End
   Begin VB.Frame FrameAdd 
      Caption         =   "Add New "
      Height          =   1785
      Left            =   135
      TabIndex        =   6
      Top             =   105
      Width           =   4365
      Begin VB.TextBox tCode 
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
         Left            =   660
         MaxLength       =   5
         TabIndex        =   9
         Top             =   330
         Width           =   975
      End
      Begin VB.TextBox tText 
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
         Left            =   660
         MaxLength       =   50
         TabIndex        =   8
         Top             =   1080
         Width           =   3495
      End
      Begin VB.CommandButton bAdd 
         Caption         =   "Add"
         Height          =   825
         Left            =   3120
         Picture         =   "frmHistList.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   195
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Code"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   390
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Text"
         Height          =   195
         Left            =   300
         TabIndex        =   10
         Top             =   1140
         Width           =   315
      End
   End
   Begin VB.CommandButton bMoveUp 
      Caption         =   "Move Up"
      Enabled         =   0   'False
      Height          =   825
      Left            =   6930
      Picture         =   "frmHistList.frx":0C28
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4050
      Width           =   795
   End
   Begin VB.CommandButton bMoveDown 
      Caption         =   "Move Down"
      Enabled         =   0   'False
      Height          =   825
      Left            =   6930
      Picture         =   "frmHistList.frx":106A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4875
      Width           =   795
   End
   Begin VB.CommandButton bSave 
      BackColor       =   &H0000FFFF&
      Caption         =   "Save"
      Height          =   825
      Left            =   6930
      Picture         =   "frmHistList.frx":14AC
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2025
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Frame Frame1 
      Caption         =   "List Type"
      Height          =   1815
      Left            =   4635
      TabIndex        =   0
      Top             =   90
      Width           =   4155
      Begin VB.OptionButton o 
         Caption         =   "Histology Autotext"
         Height          =   225
         Index           =   2
         Left            =   405
         TabIndex        =   15
         Top             =   540
         Width           =   1725
      End
      Begin VB.OptionButton o 
         Caption         =   "Histology Stains"
         Height          =   225
         Index           =   0
         Left            =   390
         TabIndex        =   2
         Top             =   270
         Width           =   1455
      End
      Begin VB.OptionButton o 
         Caption         =   "Nature of Specimen"
         Height          =   225
         Index           =   1
         Left            =   2130
         TabIndex        =   1
         Top             =   270
         Width           =   1770
      End
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   6285
      Left            =   165
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1995
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
End
Attribute VB_Name = "frmHistList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Activated As Boolean

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

110       bsave.Visible = True

120       FixG g

130       Exit Sub

bAdd_Click_Error:

          Dim strES As String
          Dim intEL As Integer



140       intEL = Erl
150       strES = Err.Description
160       LogError "frmHistList", "bAdd_Click", intEL, strES


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

210       bsave.Visible = True

220       Exit Sub

bMoveDown_Click_Error:

          Dim strES As String
          Dim intEL As Integer



230       intEL = Erl
240       strES = Err.Description
250       LogError "frmHistList", "bMoveDown_Click", intEL, strES


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

160       bsave.Visible = True

170       Exit Sub

bMoveUp_Click_Error:

          Dim strES As String
          Dim intEL As Integer



180       intEL = Erl
190       strES = Err.Description
200       LogError "frmHistList", "bMoveUp_Click", intEL, strES


End Sub

Private Sub bprint_Click()

          Dim LT As String

10        On Error GoTo bprint_Click_Error

20        LT = Switch(o(0), "Histology Stains", _
                      o(1), "Nature of Specimens")

30        Printer.Print

40        Printer.Print "List of "; LT

50        g.Col = 0
60        g.Row = 1
70        g.ColSel = g.Cols - 1
80        g.RowSel = g.Rows - 1

90        Printer.Print g.Clip

100       Printer.EndDoc


110       Exit Sub

bprint_Click_Error:

          Dim strES As String
          Dim intEL As Integer



120       intEL = Erl
130       strES = Err.Description
140       LogError "frmHistList", "bPrint_Click", intEL, strES


End Sub

Private Sub bSave_Click()

          Dim LT As String
          Dim Y As Long
          Dim tb As New Recordset
          Dim sql As String

10        On Error GoTo bSave_Click_Error

20        LT = Switch(o(0), "SH", _
                      o(1), "NA", _
                      o(2), "PH")

30        For Y = 1 To g.Rows - 1
40            sql = "SELECT * from lists WHERE listtype = '" & LT & "' and code = '" & g.TextMatrix(Y, 0) & "'"
50            Set tb = New Recordset
60            RecOpenServer 0, tb, sql
70            If tb.EOF Then tb.AddNew
80            With tb
90                !Code = g.TextMatrix(Y, 0)
100               !ListType = LT
110               !Text = g.TextMatrix(Y, 1)
120               !ListOrder = Y
130               !InUse = 1
140               tb.Update
150           End With
160       Next

170       FillG

180       tCode = ""
190       tText = ""
200       tCode.SetFocus
210       bMoveUp.Enabled = False
220       bMoveDown.Enabled = False
230       bsave.Visible = False

240       Exit Sub

bSave_Click_Error:

          Dim strES As String
          Dim intEL As Integer



250       intEL = Erl
260       strES = Err.Description
270       LogError "frmHistList", "bSave_Click", intEL, strES, sql


End Sub

Private Sub FillG()

          Dim tb As New Recordset
          Dim sql As String
          Dim LT As String
          Dim s As String

10        On Error GoTo FillG_Error

20        LT = Switch(o(0), "SH", _
                      o(1), "NA", _
                      o(2), "PH")

30        g.Rows = 2
40        g.AddItem ""
50        g.RemoveItem 1

60        sql = "SELECT * from lists WHERE listtype = '" & LT & "'"
70        Set tb = New Recordset
80        RecOpenServer 0, tb, sql
90        Do While Not tb.EOF
100           s = Trim(tb!Code) & vbTab & Trim(tb!Text)
110           g.AddItem s
120           tb.MoveNext
130       Loop

140       If g.Rows > 2 Then
150           g.RemoveItem 1
160       End If

170       Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer



180       intEL = Erl
190       strES = Err.Description
200       LogError "frmHistList", "FillG", intEL, strES, sql


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
80        LogError "frmHistList", "Form_Activate", intEL, strES


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
70        LogError "frmHistList", "Form_Load", intEL, strES


End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

10        On Error GoTo Form_QueryUnload_Error

20        If bsave.Visible Then
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
110       LogError "frmHistList", "Form_QueryUnload", intEL, strES


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
350       LogError "frmHistList", "g_Click", intEL, strES


End Sub



Private Sub g_DblClick()

10        On Error GoTo g_DblClick_Error

20        If g.Col = 1 Then
30            tCode = g.TextMatrix(g.RowSel, 0)
40            tText = g.TextMatrix(g.RowSel, 1)
50            If g.Rows = 2 Then
60                g.AddItem ""
70                g.RemoveItem g.RowSel
80            Else
90                g.RemoveItem g.RowSel
100           End If
110       End If

120       Exit Sub

g_DblClick_Error:

          Dim strES As String
          Dim intEL As Integer



130       intEL = Erl
140       strES = Err.Description
150       LogError "frmHistList", "g_DblClick", intEL, strES


End Sub

Private Sub o_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

10        On Error GoTo o_MouseUp_Error

20        FillG

30        FrameAdd.Caption = "Add New " & Left$(o(Index).Caption, Len(o(Index).Caption) - 1)

40        tCode = ""
50        tText = ""
60        If tCode.Visible Then
70            tCode.SetFocus
80        End If

90        Exit Sub

o_MouseUp_Error:

          Dim strES As String
          Dim intEL As Integer



100       intEL = Erl
110       strES = Err.Description
120       LogError "frmHistList", "o_MouseUp", intEL, strES


End Sub



