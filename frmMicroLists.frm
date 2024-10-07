VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmMicroLists 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Micrbiology Lists"
   ClientHeight    =   8520
   ClientLeft      =   150
   ClientTop       =   345
   ClientWidth     =   8565
   Icon            =   "frmMicroLists.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   8565
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrUp 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   7020
      Top             =   5370
   End
   Begin VB.Timer tmrDown 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   7020
      Top             =   6150
   End
   Begin VB.CommandButton cmdOrganisms 
      Caption         =   "Organisms"
      Height          =   765
      Left            =   7200
      Picture         =   "frmMicroLists.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1650
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrint 
      Appearance      =   0  'Flat
      Caption         =   "&Print"
      Height          =   825
      Left            =   7245
      Picture         =   "frmMicroLists.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6435
      Width           =   1155
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   825
      Left            =   7245
      Picture         =   "frmMicroLists.frx":1F96
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7515
      Width           =   1155
   End
   Begin VB.Frame FrameAdd 
      Caption         =   "Add New Organism Group"
      Height          =   1365
      Left            =   180
      TabIndex        =   5
      Top             =   150
      Width           =   4365
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
         Left            =   660
         MaxLength       =   5
         TabIndex        =   8
         Top             =   330
         Width           =   975
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
         Left            =   660
         MaxLength       =   50
         TabIndex        =   7
         Top             =   900
         Width           =   3495
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   285
         Left            =   3480
         TabIndex        =   6
         Top             =   330
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Code"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   390
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Text"
         Height          =   195
         Left            =   300
         TabIndex        =   9
         Top             =   960
         Width           =   315
      End
   End
   Begin VB.CommandButton cmdMoveUp 
      Caption         =   "Move Up"
      Enabled         =   0   'False
      Height          =   825
      Left            =   7245
      Picture         =   "frmMicroLists.frx":3918
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3495
      Width           =   1155
   End
   Begin VB.CommandButton cmdMoveDown 
      Caption         =   "Move Down"
      Enabled         =   0   'False
      Height          =   825
      Left            =   7245
      Picture         =   "frmMicroLists.frx":529A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4335
      Width           =   1155
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H0000FFFF&
      Caption         =   "Save"
      Height          =   825
      Left            =   7245
      Picture         =   "frmMicroLists.frx":6C1C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2520
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Frame Frame1 
      Caption         =   "List Type"
      Height          =   1365
      Left            =   4710
      TabIndex        =   0
      Top             =   150
      Width           =   3705
      Begin VB.OptionButton optList 
         Alignment       =   1  'Right Justify
         Caption         =   "Qualifiers"
         Height          =   225
         Index           =   8
         Left            =   750
         TabIndex        =   21
         Top             =   960
         Width           =   975
      End
      Begin VB.OptionButton optList 
         Caption         =   "Miscellaneous"
         Height          =   225
         Index           =   7
         Left            =   1800
         TabIndex        =   20
         Top             =   960
         Width           =   1395
      End
      Begin VB.OptionButton optList 
         Caption         =   "Crystals"
         Height          =   225
         Index           =   6
         Left            =   1800
         TabIndex        =   19
         Top             =   720
         Width           =   915
      End
      Begin VB.OptionButton optList 
         Caption         =   "Casts"
         Height          =   225
         Index           =   5
         Left            =   1800
         TabIndex        =   18
         Top             =   480
         Width           =   1035
      End
      Begin VB.OptionButton optList 
         Caption         =   "Wet Preps"
         Height          =   225
         Index           =   4
         Left            =   1800
         TabIndex        =   16
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optList 
         Alignment       =   1  'Right Justify
         Caption         =   "Gram Stains"
         Height          =   225
         Index           =   3
         Left            =   510
         TabIndex        =   15
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton optList 
         Alignment       =   1  'Right Justify
         Caption         =   "Organism Groups"
         Height          =   225
         Index           =   1
         Left            =   180
         TabIndex        =   14
         Top             =   240
         Width           =   1545
      End
      Begin VB.OptionButton optList 
         Alignment       =   1  'Right Justify
         Caption         =   "Ovae"
         Height          =   225
         Index           =   2
         Left            =   960
         TabIndex        =   1
         Top             =   480
         Width           =   765
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdList 
      Height          =   6675
      Left            =   210
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1650
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   11774
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
Attribute VB_Name = "frmMicroLists"
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

20        If grdList.Row = grdList.Rows - 1 Then Exit Sub
30        n = grdList.Row

40        VisibleRows = grdList.Height \ grdList.RowHeight(1) - 1

50        FireCounter = FireCounter + 1
60        If FireCounter > 5 Then
70            tmrDown.Interval = 100
80        End If

90        grdList.Visible = False

100       s = ""
110       For X = 0 To grdList.Cols - 1
120           s = s & grdList.TextMatrix(n, X) & vbTab
130       Next
140       s = Left$(s, Len(s) - 1)

150       grdList.RemoveItem n
160       If n < grdList.Rows Then
170           grdList.AddItem s, n + 1
180           grdList.Row = n + 1
190       Else
200           grdList.AddItem s
210           grdList.Row = grdList.Rows - 1
220       End If

230       For X = 0 To grdList.Cols - 1
240           grdList.Col = X
250           grdList.CellBackColor = vbYellow
260       Next

270       If Not grdList.RowIsVisible(grdList.Row) Or grdList.Row = grdList.Rows - 1 Then
280           If grdList.Row - VisibleRows + 1 > 0 Then
290               grdList.TopRow = grdList.Row - VisibleRows + 1
300           End If
310       End If

320       grdList.Visible = True

330       cmdSave.Visible = True

340       Exit Sub

FireDown_Error:

          Dim strES As String
          Dim intEL As Integer



350       intEL = Erl
360       strES = Err.Description
370       LogError "frmMicroLists", "FireDown", intEL, strES


End Sub

Private Sub FireUp()

          Dim n As Long
          Dim s As String
          Dim X As Long

10        On Error GoTo FireUp_Error

20        If grdList.Row = 1 Then Exit Sub

30        FireCounter = FireCounter + 1
40        If FireCounter > 5 Then
50            tmrUp.Interval = 100
60        End If

70        n = grdList.Row

80        grdList.Visible = False

90        s = ""
100       For X = 0 To grdList.Cols - 1
110           s = s & grdList.TextMatrix(n, X) & vbTab
120       Next
130       s = Left$(s, Len(s) - 1)

140       grdList.RemoveItem n
150       grdList.AddItem s, n - 1

160       grdList.Row = n - 1
170       For X = 0 To grdList.Cols - 1
180           grdList.Col = X
190           grdList.CellBackColor = vbYellow
200       Next

210       If Not grdList.RowIsVisible(grdList.Row) Then
220           grdList.TopRow = grdList.Row
230       End If

240       grdList.Visible = True

250       cmdSave.Visible = True

260       Exit Sub

FireUp_Error:

          Dim strES As String
          Dim intEL As Integer



270       intEL = Erl
280       strES = Err.Description
290       LogError "frmMicroLists", "FireUp", intEL, strES


End Sub



Private Sub FillG()

          Dim tb As Recordset
          Dim sql As String
          Dim LT As String
          Dim s As String

10        On Error GoTo FillG_Error

20        LT = Switch(optList(1), "OR", _
                      optList(2), "OV", _
                      optList(3), "GS", _
                      optList(4), "WP", _
                      optList(5), "CA", _
                      optList(6), "CR", _
                      optList(7), "MI", _
                      optList(8), "MQ")

30        grdList.Rows = 2
40        grdList.AddItem ""
50        grdList.RemoveItem 1

60        sql = "SELECT * from Lists WHERE " & _
                "ListType = '" & LT & "' order by ListOrder"
70        Set tb = New Recordset
80        RecOpenServer 0, tb, sql
90        Do While Not tb.EOF
100           s = tb!Code & vbTab & initial2upper(tb!Text & "")
110           grdList.AddItem s
120           tb.MoveNext
130       Loop

140       If grdList.Rows > 2 Then
150           grdList.RemoveItem 1
160       End If

170       Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer



180       intEL = Erl
190       strES = Err.Description
200       LogError "frmMicroLists", "FillG", intEL, strES, sql


End Sub

Private Sub cmdadd_Click()

10        On Error GoTo cmdadd_Click_Error

20        txtCode = Trim$(UCase$(txtCode))
30        txtText = Trim$(txtText)

40        If txtCode = "" Then
50            Exit Sub
60        End If

70        If txtText = "" Then Exit Sub

80        grdList.AddItem txtCode & vbTab & txtText

90        txtCode = ""
100       txtText = ""

110       If grdList.Rows > 2 And grdList.TextMatrix(1, 0) = "" Then grdList.RemoveItem 1

120       cmdSave.Visible = True

130       Exit Sub

cmdadd_Click_Error:

          Dim strES As String
          Dim intEL As Integer



140       intEL = Erl
150       strES = Err.Description
160       LogError "frmMicroLists", "cmdAdd_Click", intEL, strES


End Sub


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
90        LogError "frmMicroLists", "cmdMoveDown_MouseDown", intEL, strES


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
60        LogError "frmMicroLists", "cmdMoveDown_MouseUp", intEL, strES


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
90        LogError "frmMicroLists", "cmdMoveUp_MouseDown", intEL, strES


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
60        LogError "frmMicroLists", "cmdMoveUp_MouseUp", intEL, strES


End Sub


Private Sub cmdPrint_Click()

          Dim LT As String

10        On Error GoTo cmdPrint_Click_Error

20        LT = Switch(optList(1), "Organisms.", _
                      optList(2), "Ova.", _
                      optList(3), "Gram Stains.", _
                      optList(4), "Wet Preps.", _
                      optList(5), "Casts.", _
                      optList(6), "Crystals.", _
                      optList(7), "Miscellaneous.", _
                      optList(8), "Qualifiers")

30        Printer.Print

40        Printer.Print "List of "; LT

50        grdList.Col = 0
60        grdList.Row = 1
70        grdList.ColSel = grdList.Cols - 1
80        grdList.RowSel = grdList.Rows - 1

90        Printer.Print grdList.Clip

100       Printer.EndDoc


110       Exit Sub

cmdPrint_Click_Error:

          Dim strES As String
          Dim intEL As Integer



120       intEL = Erl
130       strES = Err.Description
140       LogError "frmMicroLists", "cmdPrint_Click", intEL, strES


End Sub




Private Sub cmdSave_Click()

          Dim LT As String
          Dim Num As Long
          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo cmdSave_Click_Error

20        LT = Switch(optList(1), "OR", _
                      optList(2), "OV", _
                      optList(3), "GS", _
                      optList(4), "WP", _
                      optList(5), "CA", _
                      optList(6), "CR", _
                      optList(7), "MI", _
                      optList(8), "MQ")

30        For Num = 1 To grdList.Rows - 1
40            sql = "SELECT * from Lists WHERE " & _
                    "ListType = '" & LT & "' " & _
                    "and Code = '" & grdList.TextMatrix(Num, 0) & "'"
50            Set tb = New Recordset
60            RecOpenServer 0, tb, sql
70            If tb.EOF Then
80                tb.AddNew
90            End If
100           tb!Code = grdList.TextMatrix(Num, 0)
110           tb!ListType = LT
120           tb!Text = grdList.TextMatrix(Num, 1)
130           tb!ListOrder = Num
140           tb!InUse = 1
150           tb.Update
160       Next

170       FillG

180       txtCode = ""
190       txtText = ""
200       cmdMoveUp.Enabled = False
210       cmdMoveDown.Enabled = False
220       cmdSave.Visible = False

230       Exit Sub

cmdSave_Click_Error:

          Dim strES As String
          Dim intEL As Integer



240       intEL = Erl
250       strES = Err.Description
260       LogError "frmMicroLists", "cmdsave_Click", intEL, strES, sql


End Sub


Private Sub cmdOrganisms_Click()

10        On Error GoTo cmdOrganisms_Click_Error

20        frmOrganisms.Show 1

30        Exit Sub

cmdOrganisms_Click_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmMicroLists", "cmdOrganisms_Click", intEL, strES


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
80        LogError "frmMicroLists", "Form_Activate", intEL, strES


End Sub

Private Sub Form_Load()

10        On Error GoTo Form_Load_Error

20        grdList.Font.Bold = True

30        Activated = False

40        Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer



50        intEL = Erl
60        strES = Err.Description
70        LogError "frmMicroLists", "Form_Load", intEL, strES


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
110       LogError "frmMicroLists", "Form_QueryUnload", intEL, strES


End Sub


Private Sub grdList_Click()

          Static SortOrder As Boolean
          Dim Numx As Long
          Dim Numy As Long
          Dim NumySave As Long

10        On Error GoTo grdList_Click_Error

20        NumySave = grdList.Row

30        grdList.Visible = False
40        grdList.Col = 0
50        For Numy = 1 To grdList.Rows - 1
60            grdList.Row = Numy
70            If grdList.CellBackColor = vbYellow Then
80                For Numx = 0 To grdList.Cols - 1
90                    grdList.Col = Numx
100                   grdList.CellBackColor = 0
110               Next
120               Exit For
130           End If
140       Next
150       grdList.Row = NumySave
160       grdList.Visible = True

170       If grdList.MouseRow = 0 Then
180           If SortOrder Then
190               grdList.Sort = flexSortGenericAscending
200           Else
210               grdList.Sort = flexSortGenericDescending
220           End If
230           SortOrder = Not SortOrder
240           Exit Sub
250       End If

260       For Numx = 0 To grdList.Cols - 1
270           grdList.Col = Numx
280           grdList.CellBackColor = vbYellow
290       Next

300       cmdMoveUp.Enabled = True
310       cmdMoveDown.Enabled = True

320       Exit Sub

grdList_Click_Error:

          Dim strES As String
          Dim intEL As Integer



330       intEL = Erl
340       strES = Err.Description
350       LogError "frmMicroLists", "grdList_Click", intEL, strES


End Sub


Private Sub optList_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

10        On Error GoTo optList_MouseUp_Error

20        FillG

30        FrameAdd.Caption = "Add New " & Left$(optList(Index).Caption, Len(optList(Index).Caption) - 1)

40        txtCode = ""
50        txtText = ""
60        If txtCode.Visible Then
70            txtCode.SetFocus
80        End If

90        Exit Sub

optList_MouseUp_Error:

          Dim strES As String
          Dim intEL As Integer



100       intEL = Erl
110       strES = Err.Description
120       LogError "frmMicroLists", "optList_MouseUp", intEL, strES


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
60        LogError "frmMicroLists", "tmrDown_Timer", intEL, strES


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
60        LogError "frmMicroLists", "tmrUp_Timer", intEL, strES


End Sub


