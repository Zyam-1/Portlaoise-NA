VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmComments 
   Caption         =   "NetAcquire - Comments"
   ClientHeight    =   8925
   ClientLeft      =   795
   ClientTop       =   720
   ClientWidth     =   11685
   Icon            =   "frmComments.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8925
   ScaleWidth      =   11685
   Begin VB.CommandButton bprint 
      Appearance      =   0  'Flat
      Caption         =   "&Print"
      Height          =   900
      Left            =   10375
      Picture         =   "frmComments.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2625
      Width           =   1100
   End
   Begin VB.CommandButton bcancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   900
      Left            =   10375
      Picture         =   "frmComments.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7815
      Width           =   1100
   End
   Begin VB.Frame FrameAdd 
      Caption         =   "Add New Comment"
      Height          =   2175
      Left            =   180
      TabIndex        =   10
      Top             =   180
      Width           =   4530
      Begin VB.CommandButton bDelete 
         Caption         =   "Delete"
         Height          =   780
         Left            =   3240
         Picture         =   "frmComments.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   225
         Width           =   1035
      End
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
         TabIndex        =   0
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
         Height          =   960
         Left            =   660
         MaxLength       =   320
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   1080
         Width           =   3630
      End
      Begin VB.CommandButton bAdd 
         Caption         =   "&Add"
         Height          =   780
         Left            =   1710
         Picture         =   "frmComments.frx":0C28
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   225
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Code"
         Height          =   195
         Left            =   180
         TabIndex        =   12
         Top             =   375
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Text"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   1125
         Width           =   315
      End
   End
   Begin VB.CommandButton bMoveUp 
      Caption         =   "Move &Up"
      Enabled         =   0   'False
      Height          =   825
      Left            =   10110
      Picture         =   "frmComments.frx":0F32
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5070
      Width           =   795
   End
   Begin VB.CommandButton bMoveDown 
      Caption         =   "Move &Down"
      Enabled         =   0   'False
      Height          =   825
      Left            =   10110
      Picture         =   "frmComments.frx":1374
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5910
      Width           =   795
   End
   Begin VB.CommandButton bSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   900
      Left            =   10375
      Picture         =   "frmComments.frx":17B6
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3645
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Caption         =   "List Type"
      Height          =   2160
      Left            =   4800
      TabIndex        =   3
      Top             =   180
      Width           =   6675
      Begin VB.OptionButton optType 
         Caption         =   "B/C M.Scient. Comment"
         Height          =   225
         Index           =   16
         Left            =   4425
         TabIndex        =   30
         Top             =   795
         Width           =   2085
      End
      Begin VB.OptionButton optType 
         Caption         =   "B/C Consultant Comment"
         Height          =   225
         Index           =   15
         Left            =   4425
         TabIndex        =   29
         Top             =   540
         Width           =   2085
      End
      Begin VB.OptionButton optType 
         Caption         =   "Urine Comments"
         Height          =   225
         Index           =   14
         Left            =   4425
         TabIndex        =   28
         Top             =   285
         Width           =   2085
      End
      Begin VB.OptionButton optType 
         Caption         =   "Immunology Results"
         Height          =   225
         Index           =   13
         Left            =   2340
         TabIndex        =   27
         Top             =   1815
         Width           =   2115
      End
      Begin VB.OptionButton optType 
         Caption         =   "C&&S Comments"
         Height          =   225
         Index           =   11
         Left            =   2340
         TabIndex        =   26
         Top             =   1305
         Width           =   2115
      End
      Begin VB.OptionButton optType 
         Caption         =   "Cytology Comments"
         Enabled         =   0   'False
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   25
         Top             =   1065
         Width           =   2115
      End
      Begin VB.OptionButton optType 
         Caption         =   "Histology Comments"
         Enabled         =   0   'False
         Height          =   195
         Index           =   7
         Left            =   2340
         TabIndex        =   24
         Top             =   300
         Width           =   2115
      End
      Begin VB.OptionButton optType 
         Caption         =   "Immunology Comments"
         Enabled         =   0   'False
         Height          =   195
         Index           =   8
         Left            =   2340
         TabIndex        =   23
         Top             =   555
         Width           =   2115
      End
      Begin VB.OptionButton optType 
         Caption         =   "Clinical Details"
         Height          =   225
         Index           =   12
         Left            =   2340
         TabIndex        =   22
         Top             =   1560
         Width           =   2115
      End
      Begin VB.OptionButton optType 
         Caption         =   "Micro Comments"
         Height          =   225
         Index           =   9
         Left            =   2340
         TabIndex        =   21
         Top             =   795
         Width           =   2115
      End
      Begin VB.OptionButton optType 
         Caption         =   "Demographic Comments"
         Height          =   195
         Index           =   4
         Left            =   135
         TabIndex        =   19
         Top             =   1320
         Width           =   2115
      End
      Begin VB.OptionButton optType 
         Caption         =   "Semen Comments"
         Height          =   195
         Index           =   10
         Left            =   2340
         TabIndex        =   18
         Top             =   1065
         Width           =   2115
      End
      Begin VB.OptionButton optType 
         Caption         =   "Blood Gas Comments"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   17
         Top             =   555
         Width           =   2115
      End
      Begin VB.OptionButton optType 
         Caption         =   "Endocrinology Comments"
         Enabled         =   0   'False
         Height          =   195
         Index           =   5
         Left            =   135
         TabIndex        =   16
         Top             =   1575
         Width           =   2115
      End
      Begin VB.OptionButton optType 
         Caption         =   "Biochemistry Comments"
         Enabled         =   0   'False
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   6
         Top             =   300
         Width           =   2115
      End
      Begin VB.OptionButton optType 
         Caption         =   "Haematology Comments"
         Enabled         =   0   'False
         Height          =   195
         Index           =   6
         Left            =   135
         TabIndex        =   5
         Top             =   1830
         Width           =   2115
      End
      Begin VB.OptionButton optType 
         Caption         =   "Coagulation Comments"
         Enabled         =   0   'False
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   4
         Top             =   810
         Width           =   2115
      End
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   6150
      Left            =   240
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2625
      Width           =   9795
      _ExtentX        =   17277
      _ExtentY        =   10848
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
      FormatString    =   $"frmComments.frx":1AC0
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
Attribute VB_Name = "frmComments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Activated As Boolean

Private Sub FillG()

          Dim LT As String
          Dim s As String
          Dim tb As New Recordset
          Dim sql As String

          'Changed 15/Jul/2004
10        On Error GoTo FillG_Error

20        LT = Switch(optType(0), "BI", _
                      optType(1), "BG", _
                      optType(2), "CO", _
                      optType(3), "CI", _
                      optType(4), "DE", _
                      optType(5), "EN", _
                      optType(6), "HA", _
                      optType(7), "HI", _
                      optType(8), "IM", _
                      optType(9), "MG", _
                      optType(10), "SE", _
                      optType(11), "BA", _
                      optType(12), "CD", _
                      optType(13), "IR", _
                      optType(14), "UC", _
                      optType(15), "ConsComment", _
                      optType(16), "MSComment")

30        ClearFGrid g

40        If LT = "DE" Then
50            tText.MaxLength = 160
60        Else
70            tText.MaxLength = 320
80        End If

90        sql = "SELECT * from lists WHERE listtype = '" & LT & "' order by listorder"
100       Set tb = New Recordset
110       RecOpenServer 0, tb, sql
120       Do While Not tb.EOF
130           s = Trim(tb!Code) & vbTab & Trim(tb!Text)
140           g.AddItem s
150           tb.MoveNext
160       Loop

170       FixG g

180       Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer

190       intEL = Erl
200       strES = Err.Description
210       LogError "frmComments", "FillG", intEL, strES

End Sub

Private Sub bAdd_Click()

10        On Error GoTo bAdd_Click_Error

20        tCode = Trim$(UCase$(tCode))
30        tText = Trim$(tText)

40        If Trim(tCode) = "" Then
50            Exit Sub
60        End If

70        If Len(tCode) < 3 Then
80            iMsg "Please Make Code 3 to 5 Characters"
90            Exit Sub
100       End If

110       If tText = "" Then Exit Sub

120       g.AddItem tCode & vbTab & tText

130       If g.Rows > 2 And g.TextMatrix(1, 0) = "" Then g.RemoveItem 1
140       tCode = ""
150       tText = ""

160       bsave.Enabled = True

170       Exit Sub

bAdd_Click_Error:

          Dim strES As String
          Dim intEL As Integer

180       intEL = Erl
190       strES = Err.Description
200       LogError "frmComments", "bAdd_Click", intEL, strES

End Sub


Private Sub bcancel_Click()

10        Unload Me

End Sub


Private Sub bDELETE_Click()
          Dim LT As String
          Dim sql As String

10        On Error GoTo bDELETE_Click_Error

20        LT = Switch(optType(0), "BI", _
                      optType(1), "BG", _
                      optType(2), "CO", _
                      optType(3), "CI", _
                      optType(4), "DE", _
                      optType(5), "EN", _
                      optType(6), "HA", _
                      optType(7), "HI", _
                      optType(8), "IM", _
                      optType(9), "MG", _
                      optType(10), "SE", _
                      optType(11), "BA", _
                      optType(12), "CD", _
                      optType(13), "IR", _
                      optType(14), "UC")



30        sql = "DELETE from lists WHERE listtype = '" & LT & "' and code = '" & tCode & "'"
40        Cnxn(0).Execute sql

50        tCode = ""
60        tText = ""

70        Exit Sub

bDELETE_Click_Error:

          Dim strES As String
          Dim intEL As Integer

80        intEL = Erl
90        strES = Err.Description
100       LogError "frmComments", "bDELETE_Click", intEL, strES, sql


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
250       LogError "frmComments", "bMoveDown_Click", intEL, strES


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
200       LogError "frmComments", "bMoveUp_Click", intEL, strES


End Sub


Private Sub bprint_Click()

          Dim n As Long
          Dim LT As String

10        On Error GoTo bprint_Click_Error

20        If g.Rows = 2 And g.TextMatrix(1, 0) = "" Then Exit Sub

30        LT = Switch(optType(0), "Biochemistry Comments.", _
                      optType(1), "Blood Gas Comments.", _
                      optType(2), "Coagulation Comments.", _
                      optType(3), "Cytology Comments", _
                      optType(4), "Demographic Comments.", _
                      optType(5), "Endocrinology Comments.", _
                      optType(6), "Haematology Comments.", _
                      optType(7), "Histology Comments", _
                      optType(8), "Immunology Comments.", _
                      optType(9), "Micro Comments.", _
                      optType(10), "Semen Comments.", _
                      optType(11), "C&S Comments", _
                      optType(12), "Clinical Details.", _
                      optType(13), "Immunology Results", _
                      optType(14), "Urine Comments")

40        Printer.Print

50        Printer.Print Tab(10); "List of "; LT

60        For n = 1 To g.Rows - 1
70            Printer.Print Tab(10); Trim(g.TextMatrix(n, 0));
80            Printer.Print Tab(20); Trim(g.TextMatrix(n, 1))
90        Next
100       Printer.EndDoc

110       Exit Sub

bprint_Click_Error:

          Dim strES As String
          Dim intEL As Integer

120       intEL = Erl
130       strES = Err.Description
140       LogError "frmComments", "bPrint_Click", intEL, strES

End Sub


Private Sub bSave_Click()

          Dim LT As String
          Dim Y As Long
          Dim sql As String

10        On Error GoTo bSave_Click_Error

20        LT = Switch(optType(0), "BI", _
                      optType(1), "BG", _
                      optType(2), "CO", _
                      optType(3), "CI", _
                      optType(4), "DE", _
                      optType(5), "EN", _
                      optType(6), "HA", _
                      optType(7), "HI", _
                      optType(8), "IM", _
                      optType(9), "MG", _
                      optType(10), "SE", _
                      optType(11), "BA", _
                      optType(12), "CD", _
                      optType(13), "IR", _
                      optType(14), "UC", _
                      optType(15), "ConsComment", _
                      optType(16), "MSComment")

30        For Y = 1 To g.Rows - 1
40            sql = "IF EXISTS (SELECT * FROM Lists WHERE " & _
                    "           ListType = '" & LT & "' " & _
                    "           AND Code = '" & g.TextMatrix(Y, 0) & "') " & _
                    "  UPDATE Lists " & _
                    "  SET Text = '" & AddTicks(g.TextMatrix(Y, 1)) & "', " & _
                    "  ListOrder = '" & Y & "', " & _
                    "  InUse = 1 " & _
                    "  WHERE ListType = '" & LT & "' " & _
                    "  AND Code = '" & g.TextMatrix(Y, 0) & "' " & _
                    "ELSE " & _
                    "  INSERT INTO Lists " & _
                    "  (Code, ListType, Text, ListOrder, InUse) VALUES " & _
                    "  ('" & g.TextMatrix(Y, 0) & "', " & _
                    "  '" & LT & "', " & _
                    "  '" & AddTicks(g.TextMatrix(Y, 1)) & "', " & _
                    "  '" & Y & "', " & _
                    "  '1')"
50            Cnxn(0).Execute sql
60        Next

70        FillG

80        tCode = ""
90        tText = ""
100       tCode.SetFocus
110       bMoveUp.Enabled = False
120       bMoveDown.Enabled = False
130       bsave.Enabled = False

140       Exit Sub

bSave_Click_Error:

          Dim strES As String
          Dim intEL As Integer

150       intEL = Erl
160       strES = Err.Description
170       LogError "frmComments", "bSave_Click", intEL, strES, sql

End Sub


Private Sub Form_Activate()

10        On Error GoTo Form_Activate_Error

20        If Activated Then Exit Sub

30        Activated = True

40        optType(0).Enabled = SysOptDeptBio(0)
50        optType(6).Enabled = SysOptDeptHaem(0)
60        optType(2).Enabled = SysOptDeptCoag(0)
70        optType(5).Enabled = SysOptDeptEnd(0)
80        optType(1).Enabled = SysOptDeptBga(0)
90        optType(10).Enabled = SysOptDeptSemen(0)
100       optType(8).Enabled = SysOptDeptImm(0)
110       optType(13).Enabled = SysOptDeptImm(0)
120       optType(9).Enabled = SysOptDeptMicro(0)
130       optType(11).Enabled = SysOptDeptMicro(0)
140       optType(7).Enabled = SysOptDeptHisto(0)
150       optType(3).Enabled = SysOptDeptCyto(0)
160       optType(6).Enabled = True
170       optType(14).Enabled = SysOptDeptMicro(0)
180       FillG

190       Exit Sub

Form_Activate_Error:

          Dim strES As String
          Dim intEL As Integer

200       intEL = Erl
210       strES = Err.Description
220       LogError "frmComments", "Form_Activate", intEL, strES

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
70        LogError "frmComments", "Form_Load", intEL, strES

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
110       LogError "frmComments", "Form_QueryUnload", intEL, strES

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
200               bsave.Enabled = True
210           Else
220               g.Sort = flexSortGenericDescending
230               bsave.Enabled = True
240           End If
250           SortOrder = Not SortOrder
260           Exit Sub
270       End If

280       For X = 0 To g.Cols - 1
290           g.Col = X
300           g.CellBackColor = vbYellow
310       Next

320       bMoveUp.Enabled = True
330       bMoveDown.Enabled = True

340       Exit Sub

g_Click_Error:

          Dim strES As String
          Dim intEL As Integer

350       intEL = Erl
360       strES = Err.Description
370       LogError "frmComments", "g_Click", intEL, strES

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
150       LogError "frmComments", "g_DblClick", intEL, strES

End Sub

Private Sub optType_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

10        On Error GoTo optType_MouseUp_Error

20        FillG

30        FrameAdd.Caption = "Add New " & Left$(optType(Index).Caption, Len(optType(Index).Caption) - 1)

40        tCode = ""
50        tText = ""
60        If tCode.Visible Then
70            tCode.SetFocus
80        End If

90        Exit Sub

optType_MouseUp_Error:

          Dim strES As String
          Dim intEL As Integer

100       intEL = Erl
110       strES = Err.Description
120       LogError "frmComments", "optType_MouseUp", intEL, strES

End Sub

