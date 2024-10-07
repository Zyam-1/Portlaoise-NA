VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmAutoGenerateComments 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Auto-Generate Comments"
   ClientHeight    =   9000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   14520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   3888
      Left            =   2340
      TabIndex        =   7
      Top             =   3600
      Width           =   11472
      _ExtentX        =   20241
      _ExtentY        =   6853
      _Version        =   393216
      Cols            =   7
      BackColor       =   -2147483624
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      FormatString    =   $"frmAutoGenerateComments.frx":0000
   End
   Begin MSFlexGridLib.MSFlexGrid AllG 
      Height          =   2568
      Left            =   2400
      TabIndex        =   31
      Top             =   4140
      Width           =   6372
      _ExtentX        =   11245
      _ExtentY        =   4524
      _Version        =   393216
      Cols            =   8
      BackColor       =   -2147483624
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      FormatString    =   $"frmAutoGenerateComments.frx":00E8
   End
   Begin VB.CommandButton cmdAllToExcel 
      Caption         =   "Export All to Excel"
      Height          =   1100
      Left            =   4920
      Picture         =   "frmAutoGenerateComments.frx":01F5
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   7620
      Width           =   1200
   End
   Begin VB.CommandButton cmdExcel 
      Caption         =   "Export to Excel"
      Height          =   1100
      Left            =   2340
      Picture         =   "frmAutoGenerateComments.frx":0637
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   7630
      Width           =   1200
   End
   Begin VB.Frame Frame2 
      Caption         =   "Comment Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   2250
      TabIndex        =   25
      Top             =   180
      Width           =   5595
      Begin VB.OptionButton Option1 
         Caption         =   "Result Comment"
         Height          =   255
         Left            =   2820
         TabIndex        =   27
         Top             =   300
         Width           =   1755
      End
      Begin VB.OptionButton optGeneric 
         Caption         =   "Generic Comment"
         Height          =   255
         Left            =   660
         TabIndex        =   26
         Top             =   300
         Value           =   -1  'True
         Width           =   1755
      End
   End
   Begin VB.Timer tmrUp 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   13800
      Top             =   3630
   End
   Begin VB.Timer tmrDown 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   13800
      Top             =   5340
   End
   Begin VB.CommandButton cmdMoveDown 
      Height          =   615
      Left            =   13740
      Picture         =   "frmAutoGenerateComments.frx":0941
      Style           =   1  'Graphical
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   4710
      Width           =   525
   End
   Begin VB.CommandButton cmdMoveUp 
      Height          =   615
      Left            =   13740
      Picture         =   "frmAutoGenerateComments.frx":22C3
      Style           =   1  'Graphical
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   4080
      Width           =   525
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Enabled         =   0   'False
      Height          =   1100
      Left            =   9870
      Picture         =   "frmAutoGenerateComments.frx":3C45
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7630
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   1100
      Left            =   12510
      Picture         =   "frmAutoGenerateComments.frx":4B0F
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7630
      Width           =   1200
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   1100
      Left            =   11190
      Picture         =   "frmAutoGenerateComments.frx":59D9
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7630
      Width           =   1200
   End
   Begin VB.ListBox lstParameter 
      Height          =   8550
      IntegralHeight  =   0   'False
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   1965
   End
   Begin VB.Frame Frame1 
      Height          =   2385
      Left            =   2250
      TabIndex        =   5
      Top             =   990
      Width           =   11445
      Begin VB.Frame fraAlpha 
         Caption         =   "Alphanumeric Results"
         Height          =   765
         Left            =   5970
         TabIndex        =   17
         Top             =   270
         Width           =   5085
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
            Left            =   3510
            TabIndex        =   24
            Top             =   300
            Width           =   1035
         End
         Begin VB.ComboBox cmbText 
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
            Left            =   1470
            TabIndex        =   23
            Text            =   "cmbText"
            Top             =   300
            Width           =   1995
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "If this Result"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   22
            Top             =   330
            Width           =   1095
         End
      End
      Begin VB.Frame fraNumeric 
         BackColor       =   &H8000000A&
         Caption         =   "Numeric Results"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   765
         Left            =   180
         TabIndex        =   16
         Top             =   270
         Width           =   5385
         Begin VB.TextBox txtCriteria 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   0
            Left            =   3030
            TabIndex        =   20
            Top             =   300
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.TextBox txtCriteria 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   1
            Left            =   4110
            TabIndex        =   19
            Top             =   300
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.ComboBox cmbCriteria 
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
            ItemData        =   "frmAutoGenerateComments.frx":68A3
            Left            =   1470
            List            =   "frmAutoGenerateComments.frx":68A5
            TabIndex        =   18
            Text            =   "cmbCriteria"
            Top             =   300
            Width           =   1515
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "If this result is"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   210
            TabIndex        =   21
            Top             =   360
            Width           =   1215
         End
      End
      Begin MSComCtl2.DTPicker dtStart 
         Height          =   315
         Left            =   2550
         TabIndex        =   10
         Top             =   1860
         Width           =   1305
         _ExtentX        =   2275
         _ExtentY        =   556
         _Version        =   393216
         Format          =   157286401
         CurrentDate     =   40093
      End
      Begin MSComCtl2.DTPicker dtEnd 
         Height          =   315
         Left            =   4380
         TabIndex        =   9
         Top             =   1860
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Format          =   157286401
         CurrentDate     =   40093
      End
      Begin VB.TextBox txtComment 
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
         Left            =   780
         MaxLength       =   500
         TabIndex        =   1
         Top             =   1410
         Width           =   10335
      End
      Begin VB.Label lblMaxLength 
         AutoSize        =   -1  'True
         Caption         =   "(Maximum 500 characters)"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   2430
         TabIndex        =   15
         Top             =   1170
         Width           =   1860
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "and"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3960
         TabIndex        =   12
         Top             =   1920
         Width           =   330
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "This rule is active between"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   11
         Top             =   1920
         Width           =   2310
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Print this comment"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   780
         TabIndex        =   6
         Top             =   1170
         Width           =   1575
      End
   End
   Begin VB.Label lblExcelInfo 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exporting..."
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
      Left            =   3600
      TabIndex        =   29
      Top             =   7680
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label lblDiscipline 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Biochemistry"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   11865
      TabIndex        =   8
      Top             =   570
      Width           =   1830
   End
End
Attribute VB_Name = "frmAutoGenerateComments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pDiscipline As String
Private pDisc As String

Private FireCounter As Integer

Private EntryMode As String



Private Sub cmdAllToExcel_Click()
      Dim I As Integer
      Dim C As Integer
      Dim strHeading As String

10    On Error GoTo cmdAllToExcel_Click_Error
20    AllG.Rows = 2
30    AllG.FixedRows = 1
40    AllG.Rows = 1
50    For I = 0 To lstParameter.ListCount - 1
60        lstParameter.ListIndex = I
70        For C = 1 To g.Rows - 1
80            g.Row = C
90            g.Col = 3
100           If g.TextMatrix(g.Row, g.Col) <> "" Then
110               If g.CellBackColor = vbGreen Then
120                   AllG.AddItem lstParameter.List(I) & vbTab & "Result" & vbTab & g.TextMatrix(C, 0) & vbTab & g.TextMatrix(C, 1) & vbTab & g.TextMatrix(C, 2) & vbTab & g.TextMatrix(C, 3) & vbTab & g.TextMatrix(C, 4) & vbTab & g.TextMatrix(C, 5)
130               Else
140                   AllG.AddItem lstParameter.List(I) & vbTab & "Generic" & vbTab & g.TextMatrix(C, 0) & vbTab & g.TextMatrix(C, 1) & vbTab & g.TextMatrix(C, 2) & vbTab & g.TextMatrix(C, 3) & vbTab & g.TextMatrix(C, 4) & vbTab & g.TextMatrix(C, 5)
150               End If
160           End If
170       Next
          'lstParameter_Click
180   Next



190   If AllG.Rows < 2 Then
200       iMsg "Nothing to export"
210       Exit Sub
220   End If

230   strHeading = "List of Auto Generated Comment  for All Tests" & vbCr

240   ExportFlexGrid AllG, Me, strHeading


250   Exit Sub

cmdAllToExcel_Click_Error:

      Dim strES As String
      Dim intEL As Integer

260   intEL = Erl
270   strES = Err.Description
280   LogError "frmAutoGenerateComments", "cmdAllToExcel_Click", intEL, strES

End Sub

Private Sub cmdExcel_Click()

          Dim strHeading As String

10        On Error GoTo cmdExcel_Click_Error

20        If g.Rows < 2 Then
30            iMsg "Nothing to export"
40            Exit Sub
50        End If

60        strHeading = "List of Auto Generated Comment  for " & lstParameter.List(lstParameter.ListIndex) & vbCr

70        ExportFlexGrid g, Me, strHeading


80        Exit Sub

cmdExcel_Click_Error:

          Dim strES As String
          Dim intEL As Integer

90        intEL = Erl
100       strES = Err.Description
110       LogError "frmPhoneLogHistory", "cmdExcel_Click", intEL, strES

End Sub
Private Sub ClearHighlight()

          Dim SaveY As Integer
          Dim Y As Integer
          Dim X As Integer

10        SaveY = g.MouseRow

20        For Y = 1 To g.Rows - 1
30            g.Col = 1
40            g.Row = Y
50            If g.CellBackColor = vbYellow Then
60                For X = 1 To g.Cols - 1
70                    g.Col = X
80                    g.CellBackColor = 0
90                Next
100           End If
110       Next

120       g.Row = SaveY

End Sub

Private Sub FillList()

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo FillList_Error

20        lstParameter.Clear

30        sql = "SELECT DISTINCT ShortName FROM " & pDisc & "TestDefinitions " & _
                "ORDER BY ShortName"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        Do While Not tb.EOF
70            lstParameter.AddItem tb!ShortName & ""
80            tb.MoveNext
90        Loop

100       Exit Sub

FillList_Error:

          Dim strES As String
          Dim intEL As Integer

110       intEL = Erl
120       strES = Err.Description
130       LogError "frmAutoGenerateComments", "FillList", intEL, strES, sql

End Sub


Private Sub SaveListOrder()

          Dim Y As Integer
          Dim sql As String

10        On Error GoTo SaveListOrder_Error

20        For Y = 1 To g.Rows - 1
30            If g.TextMatrix(Y, 0) <> "" Then
40                sql = "UPDATE AutoComments " & _
                        "SET ListOrder = '" & Y & "' " & _
                        "WHERE Parameter = '" & lstParameter & "' " & _
                        "AND Comment = '" & AddTicks(g.TextMatrix(Y, 3)) & "'"
50                Cnxn(0).Execute sql
60            End If
70        Next

80        Exit Sub

SaveListOrder_Error:

          Dim strES As String
          Dim intEL As Integer

90        intEL = Erl
100       strES = Err.Description
110       LogError "frmAutoGenerateComments", "SaveListOrder", intEL, strES, sql


End Sub

Private Sub SetToNumeric()

10        fraNumeric.FontBold = True
20        fraNumeric.ForeColor = vbRed
30        fraAlpha.FontBold = False
40        fraAlpha.ForeColor = vbBlack
50        EntryMode = "Numeric"

End Sub

Private Sub SetToAlpha()

10        fraNumeric.FontBold = False
20        fraNumeric.ForeColor = vbBlack
30        fraAlpha.FontBold = True
40        fraAlpha.ForeColor = vbRed
50        EntryMode = "Alpha"

End Sub

Private Sub cmbCriteria_Click()

10        ClearHighlight
20        cmdDelete.Enabled = False

30        Select Case cmbCriteria
          Case "Present": txtCriteria(0).Visible = False: txtCriteria(1).Visible = False
40        Case "Equal to": txtCriteria(0).Visible = True
50        Case "Greater than": txtCriteria(0).Visible = True
60        Case "Less than": txtCriteria(0).Visible = True
70        Case "Between": txtCriteria(0).Visible = True: txtCriteria(1).Visible = True
80        Case "Not between": txtCriteria(0).Visible = True: txtCriteria(1).Visible = True
90        End Select

End Sub


Private Sub cmbCriteria_GotFocus()

10        SetToNumeric

End Sub


Private Sub cmbCriteria_KeyPress(KeyAscii As Integer)

10        KeyAscii = 0

End Sub


Private Sub cmbText_GotFocus()

10        SetToAlpha

End Sub


Private Sub cmbText_KeyPress(KeyAscii As Integer)

10        KeyAscii = 0

End Sub


Private Sub cmdCancel_Click()

10        Unload Me

End Sub

Private Sub cmdDelete_Click()

          Dim sql As String

10        On Error GoTo cmdDelete_Click_Error

20        sql = "DELETE FROM AutoComments WHERE " & _
                "Discipline = '" & pDiscipline & "' " & _
                "AND Parameter = '" & lstParameter & "' " & _
                "AND Criteria = '" & g.TextMatrix(g.Row, 0) & "' " & _
                "AND Comment = '" & g.TextMatrix(g.Row, 3) & "'"
30        Cnxn(0).Execute sql

40        FillG

50        cmdSave.Enabled = False

60        Exit Sub

cmdDelete_Click_Error:

          Dim strES As String
          Dim intEL As Integer

70        intEL = Erl
80        strES = Err.Description
90        LogError "frmAutoGenerateComments", "cmdDelete_Click", intEL, strES, sql

End Sub
Private Sub ClearEntry()

10        cmbCriteria.ListIndex = 0
20        txtCriteria(0) = ""
30        txtCriteria(1) = ""
40        txtComment = ""

End Sub

Private Sub cmdMoveDown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        FireDown

20        tmrDown.Interval = 250
30        FireCounter = 0

40        tmrDown.Enabled = True

End Sub


Private Sub cmdMoveDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        tmrDown.Enabled = False

20        SaveListOrder

End Sub


Private Sub FireDown()

          Dim n As Integer
          Dim s As String
          Dim X As Integer
          Dim VisibleRows As Integer

10        If g.Row = g.Rows - 1 Or g.Row = 0 Then Exit Sub
20        n = g.Row

30        FireCounter = FireCounter + 1
40        If FireCounter > 5 Then
50            tmrDown.Interval = 100
60        End If

70        VisibleRows = g.Height \ g.RowHeight(1) - 1

80        g.Visible = False

90        s = ""
100       For X = 0 To g.Cols - 1
110           s = s & g.TextMatrix(n, X) & vbTab
120       Next
130       s = Left$(s, Len(s) - 1)

140       g.RemoveItem n
150       If n < g.Rows Then
160           g.AddItem s, n + 1
170           g.Row = n + 1
180       Else
190           g.AddItem s
200           g.Row = g.Rows - 1
210       End If

220       For X = 1 To g.Cols - 1
230           g.Col = X
240           g.CellBackColor = vbYellow
250       Next

260       If Not g.RowIsVisible(g.Row) Or g.Row = g.Rows - 1 Then
270           If g.Row - VisibleRows + 1 > 0 Then
280               g.TopRow = g.Row - VisibleRows + 1
290           End If
300       End If

310       g.Visible = True

320       cmdSave.Visible = True

End Sub

Private Sub cmdMoveUp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        FireUp

20        tmrUp.Interval = 250
30        FireCounter = 0

40        tmrUp.Enabled = True

End Sub


Private Sub FireUp()

          Dim n As Integer
          Dim s As String
          Dim X As Integer

10        If g.Row < 2 Then Exit Sub

20        FireCounter = FireCounter + 1
30        If FireCounter > 5 Then
40            tmrUp.Interval = 100
50        End If

60        n = g.Row

70        g.Visible = False

80        s = ""
90        For X = 0 To g.Cols - 1
100           s = s & g.TextMatrix(n, X) & vbTab
110       Next
120       s = Left$(s, Len(s) - 1)

130       g.RemoveItem n
140       g.AddItem s, n - 1

150       g.Row = n - 1
160       For X = 1 To g.Cols - 1
170           g.Col = X
180           g.CellBackColor = vbYellow
190       Next

200       If Not g.RowIsVisible(g.Row) Then
210           g.TopRow = g.Row
220       End If

230       g.Visible = True

240       cmdSave.Visible = True

End Sub


Private Sub cmdMoveUp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        tmrUp.Enabled = False

20        SaveListOrder

End Sub


Private Sub cmdSave_Click()

          Dim sql As String
          Dim Value0 As String
          Dim Value1 As String
          Dim Criteria As String

10        On Error GoTo cmdSave_Click_Error

20        If Trim$(txtComment) <> "" Then
30            If EntryMode = "Numeric" Then
40                Criteria = cmbCriteria
50                If IsNumeric(txtCriteria(0)) Then
60                    If cmbCriteria = "Greater than" Then
70                        Value0 = Val(txtCriteria(0))    '- 0.00001
80                    ElseIf cmbCriteria = "Not between" Or _
                             cmbCriteria = "Between" Or _
                             cmbCriteria = "Less than" Or _
                             cmbCriteria = "Equal to" Then
90                        Value0 = Val(txtCriteria(0))    '+ 0.00001
100                   End If
110               Else
120                   Value0 = ""
130               End If
140               If IsNumeric(txtCriteria(1)) Then
150                   Value1 = Val(txtCriteria(1))
160               Else
170                   Value1 = ""
180               End If
190           ElseIf EntryMode = "Alpha" Then
200               Criteria = cmbText
210               Value0 = txtText
220               Value1 = ""
230           End If

240           sql = "INSERT INTO AutoComments " & _
                    "(Discipline, Parameter, Criteria, Value0, Value1, Comment, DateStart, DateEnd, CommentType) " & _
                    "VALUES " & _
                    "( '" & pDiscipline & "', " & _
                    "  '" & lstParameter & "', " & _
                    "  '" & Criteria & "', " & _
                    "  '" & AddTicks(Value0) & "', " & _
                    "  '" & AddTicks(Value1) & "', " & _
                    "  '" & AddTicks(txtComment) & "', " & _
                    "  '" & Format$(dtStart, "dd/MMM/yyyy") & "', " & _
                    "  '" & Format$(dtEnd, "dd/MMM/yyyy") & "', " & _
                    IIf(optGeneric.Value, 0, 1) & ")"
250           Cnxn(0).Execute sql

260           SaveListOrder

270       Else
280           txtComment.BackColor = vbRed
290       End If

300       cmdSave.Enabled = False

310       FillG

320       ClearEntry

330       Exit Sub

cmdSave_Click_Error:

          Dim strES As String
          Dim intEL As Integer

340       intEL = Erl
350       strES = Err.Description
360       LogError "frmAutoGenerateComments", "cmdSave_Click", intEL, strES, sql

End Sub


Private Sub Form_Activate()

10        lblDiscipline = pDiscipline

20        FillList

30        ClearHighlight
40        cmdDelete.Enabled = False

End Sub

Private Sub Form_Load()

10        CheckAutoCommentsInDb

20        cmbCriteria.AddItem "Present"
30        cmbCriteria.AddItem "Equal to"
40        cmbCriteria.AddItem "Less than"
50        cmbCriteria.AddItem "Greater than"
60        cmbCriteria.AddItem "Between"
70        cmbCriteria.AddItem "Not between"
80        cmbCriteria.ListIndex = 0

90        cmbText.AddItem "Starts with"
100       cmbText.AddItem "Contains Text"
110       cmbText.ListIndex = 0

120       dtStart = Format$(Now, "dd/MM/yyyy")
130       dtEnd = Format$(DateAdd("m", 6, Now), "dd/MM/yyyy")

140       EntryMode = "Numeric"

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

10        If cmdSave.Enabled Then
20            If iMsg("Cancel without Saving?", vbQuestion + vbYesNo) = vbNo Then
30                Cancel = True
40            End If
50        End If

End Sub


Private Sub g_Click()

          Dim X As Integer

10        ClearHighlight
20        cmdDelete.Enabled = False

30        If g.MouseRow > 0 And g.TextMatrix(g.MouseRow, 0) <> "" Then

40            For X = 1 To g.Cols - 1
50                g.Col = X
60                g.CellBackColor = vbYellow
70            Next

80            cmdDelete.Enabled = True
90        End If

End Sub

Private Sub FillG()

          Dim tb As Recordset
          Dim sql As String
          Dim s As String

10        On Error GoTo FillG_Error

20        g.Rows = 2
30        g.AddItem ""
40        g.RemoveItem 1

50        sql = "SELECT * FROM AutoComments WHERE " & _
                "Discipline = '" & pDiscipline & "' " & _
                "AND Parameter = '" & lstParameter & "' " & _
                "ORDER BY ListOrder"
60        Set tb = New Recordset
70        RecOpenServer 0, tb, sql
80        Do While Not tb.EOF
90            s = tb!Criteria & vbTab & _
                  tb!Value0 & vbTab & _
                  tb!Value1 & vbTab & _
                  tb!Comment & vbTab
100           If Not IsNull(tb!DateStart) Then
110               s = s & Format$(tb!DateStart, "dd/MM/yy")
120           End If
130           s = s & vbTab
140           If Not IsNull(tb!DateEnd) Then
150               s = s & Format$(tb!DateEnd, "dd/MM/yy")
160           End If



170           g.AddItem s
180           If Not IsNull(tb!CommentType) Then
190               If tb!CommentType = 1 Then
200                   g.Row = g.Rows - 1
210                   g.Col = 3
220                   g.CellBackColor = vbGreen
230               End If
240           End If
250           tb.MoveNext
260       Loop

270       If g.Rows > 2 Then
280           g.RemoveItem 1
290       End If

300       ClearHighlight
310       cmdDelete.Enabled = False

320       Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer

330       intEL = Erl
340       strES = Err.Description
350       LogError "frmAutoGenerateComments", "FillG", intEL, strES, sql

End Sub


Private Sub lstParameter_Click()

10        FillG

20        ClearEntry

End Sub

Private Sub optGeneric_Click()

10        On Error GoTo optGeneric_Click_Error

20        lblMaxLength = "(Maximum 500 characters)"
30        txtComment.MaxLength = 500

40        Exit Sub

optGeneric_Click_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmAutoGenerateComments", "optGeneric_Click", intEL, strES

End Sub

Private Sub Option1_Click()

10        On Error GoTo Option1_Click_Error

20        lblMaxLength.Caption = "(Maximum 95 characters)"
30        txtComment.MaxLength = 95
40        txtComment = Left(txtComment, 95)

50        Exit Sub

Option1_Click_Error:

          Dim strES As String
          Dim intEL As Integer

60        intEL = Erl
70        strES = Err.Description
80        LogError "frmAutoGenerateComments", "Option1_Click", intEL, strES

End Sub

Private Sub txtComment_KeyUp(KeyCode As Integer, Shift As Integer)

10        ClearHighlight
20        cmdDelete.Enabled = False
30        txtComment.BackColor = vbWhite

40        cmdSave.Enabled = Len(Trim$(txtComment)) > 0

End Sub

Private Sub txtCriteria_Change(Index As Integer)

10        ClearHighlight
20        cmdDelete.Enabled = False

End Sub



Public Property Let Discipline(ByVal sNewValue As String)

10        pDiscipline = sNewValue

20        If pDiscipline = "Biochemistry" Then
30            pDisc = "Bio"
40        ElseIf pDiscipline = "Coagulation" Then
50            pDisc = "Coag"
60        ElseIf pDiscipline = "Immunology" Then
70            pDisc = "Imm"
80        ElseIf pDiscipline = "Endocrinology" Then
90            pDisc = "End"
100       End If

End Property

Private Sub txtCriteria_GotFocus(Index As Integer)

10        SetToNumeric

End Sub

Private Sub txtText_GotFocus()
10        SetToAlpha
End Sub

