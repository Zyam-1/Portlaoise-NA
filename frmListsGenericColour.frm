VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmListsGenericColour 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - List of Generic"
   ClientHeight    =   8475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8475
   ScaleWidth      =   12885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPrint 
      Appearance      =   0  'Flat
      Caption         =   "&Print"
      Height          =   825
      Left            =   11610
      Picture         =   "frmListsGenericColour.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1560
      Width           =   795
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   825
      Left            =   11610
      Picture         =   "frmListsGenericColour.frx":1982
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7470
      Width           =   795
   End
   Begin VB.Frame FrameAdd 
      Caption         =   "Add New Generic"
      Height          =   975
      Left            =   90
      TabIndex        =   9
      Top             =   120
      Width           =   11085
      Begin VB.Frame Frame1 
         Caption         =   "Back Colour"
         Height          =   975
         Index           =   1
         Left            =   8250
         TabIndex        =   26
         Top             =   0
         Width           =   1905
         Begin VB.CommandButton cmdBackColour 
            Height          =   315
            Index           =   0
            Left            =   60
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   240
            Width           =   315
         End
         Begin VB.CommandButton cmdBackColour 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   1
            Left            =   420
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   240
            Width           =   315
         End
         Begin VB.CommandButton cmdBackColour 
            BackColor       =   &H000000FF&
            Height          =   315
            Index           =   5
            Left            =   60
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   600
            Width           =   315
         End
         Begin VB.CommandButton cmdBackColour 
            BackColor       =   &H000080FF&
            Height          =   315
            Index           =   6
            Left            =   420
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   600
            Width           =   315
         End
         Begin VB.CommandButton cmdBackColour 
            BackColor       =   &H0000FFFF&
            Height          =   315
            Index           =   2
            Left            =   780
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   240
            Width           =   315
         End
         Begin VB.CommandButton cmdBackColour 
            BackColor       =   &H0000FF00&
            Height          =   315
            Index           =   3
            Left            =   1140
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   240
            Width           =   315
         End
         Begin VB.CommandButton cmdBackColour 
            BackColor       =   &H00FFFF00&
            Height          =   315
            Index           =   7
            Left            =   780
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   600
            Width           =   315
         End
         Begin VB.CommandButton cmdBackColour 
            BackColor       =   &H00FF0000&
            Height          =   315
            Index           =   8
            Left            =   1140
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   600
            Width           =   315
         End
         Begin VB.CommandButton cmdBackColour 
            BackColor       =   &H00800080&
            Height          =   315
            Index           =   4
            Left            =   1500
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   240
            Width           =   315
         End
         Begin VB.CommandButton cmdBackColour 
            BackColor       =   &H00000001&
            Height          =   315
            Index           =   9
            Left            =   1500
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   600
            Width           =   315
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Fore Colour"
         Height          =   975
         Index           =   0
         Left            =   6330
         TabIndex        =   15
         Top             =   0
         Width           =   1905
         Begin VB.CommandButton cmdForeColour 
            BackColor       =   &H00000001&
            Height          =   315
            Index           =   9
            Left            =   1500
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   600
            Width           =   315
         End
         Begin VB.CommandButton cmdForeColour 
            BackColor       =   &H00800080&
            Height          =   315
            Index           =   4
            Left            =   1500
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   240
            Width           =   315
         End
         Begin VB.CommandButton cmdForeColour 
            BackColor       =   &H00FF0000&
            Height          =   315
            Index           =   8
            Left            =   1140
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   600
            Width           =   315
         End
         Begin VB.CommandButton cmdForeColour 
            BackColor       =   &H00FFFF00&
            Height          =   315
            Index           =   7
            Left            =   780
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   600
            Width           =   315
         End
         Begin VB.CommandButton cmdForeColour 
            BackColor       =   &H0000FF00&
            Height          =   315
            Index           =   3
            Left            =   1140
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   240
            Width           =   315
         End
         Begin VB.CommandButton cmdForeColour 
            BackColor       =   &H0000FFFF&
            Height          =   315
            Index           =   2
            Left            =   780
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   240
            Width           =   315
         End
         Begin VB.CommandButton cmdForeColour 
            BackColor       =   &H000080FF&
            Height          =   315
            Index           =   6
            Left            =   420
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   600
            Width           =   315
         End
         Begin VB.CommandButton cmdForeColour 
            BackColor       =   &H000000FF&
            Height          =   315
            Index           =   5
            Left            =   60
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   600
            Width           =   315
         End
         Begin VB.CommandButton cmdForeColour 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   1
            Left            =   420
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   240
            Width           =   315
         End
         Begin VB.CommandButton cmdForeColour 
            Height          =   315
            Index           =   0
            Left            =   60
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   240
            Width           =   315
         End
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
         Left            =   120
         MaxLength       =   15
         TabIndex        =   0
         Top             =   420
         Width           =   1545
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
         Left            =   1740
         MaxLength       =   50
         TabIndex        =   1
         Top             =   420
         Width           =   4365
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   765
         Left            =   10350
         Picture         =   "frmListsGenericColour.frx":3304
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   150
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Code"
         Height          =   195
         Left            =   150
         TabIndex        =   11
         Top             =   210
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Text"
         Height          =   195
         Left            =   1770
         TabIndex        =   10
         Top             =   210
         Width           =   315
      End
   End
   Begin VB.CommandButton cmdMoveUp 
      Caption         =   "Move Up"
      Enabled         =   0   'False
      Height          =   825
      Left            =   11610
      Picture         =   "frmListsGenericColour.frx":4C86
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   5160
      Width           =   795
   End
   Begin VB.CommandButton cmdMoveDown 
      Caption         =   "Move Down"
      Enabled         =   0   'False
      Height          =   825
      Left            =   11610
      Picture         =   "frmListsGenericColour.frx":6608
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   6030
      Width           =   795
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H0000FFFF&
      Caption         =   "Save"
      Height          =   825
      Left            =   11610
      Picture         =   "frmListsGenericColour.frx":7F8A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2730
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Enabled         =   0   'False
      Height          =   825
      Left            =   11610
      Picture         =   "frmListsGenericColour.frx":990C
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3870
      Width           =   795
   End
   Begin VB.Timer tmrDown 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   12420
      Top             =   6180
   End
   Begin VB.Timer tmrUp 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   12420
      Top             =   5400
   End
   Begin VB.CommandButton cmdXL 
      Caption         =   "Export to Excel"
      Height          =   735
      Left            =   11520
      Picture         =   "frmListsGenericColour.frx":B28E
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   240
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   7125
      Left            =   90
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1170
      Width           =   11115
      _ExtentX        =   19606
      _ExtentY        =   12568
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
      FormatString    =   $"frmListsGenericColour.frx":B598
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
      Left            =   11340
      TabIndex        =   14
      Top             =   990
      Visible         =   0   'False
      Width           =   1305
   End
End
Attribute VB_Name = "frmListsGenericColour"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Activated As Boolean

Private FireCounter As Integer

Private pListTypeNames As String
Private pListTypeName As String
Private pListType As String
'pListType = "UN", "Units"
'            "ER", "Errors"
'            "ST", "SampleTypes"
'            "MB", "Specimen Sources"



Private pForeColour As Long
Private pBackColour As Long

Private Sub FillG()

          Dim s As String
          Dim tb As Recordset
          Dim sql As String
          Dim D() As String
10        On Error GoTo FillG_Error

20        g.Visible = False
30        g.Rows = 2
40        g.AddItem ""
50        g.RemoveItem 1

60        sql = "SELECT Code, Text, [Default] FROM Lists WHERE " & _
                "ListType = '" & pListType & "' " & _
                "ORDER BY ListOrder"
70        Set tb = New Recordset
80        RecOpenServer 0, tb, sql
90        Do While Not tb.EOF
100           s = tb!Code & vbTab & tb!Text & ""
110           g.AddItem s
120           g.Row = g.Rows - 1
130           g.Col = 1
140           D = Split(tb!Default & "", "|")
150           If UBound(D) > 0 Then
160               g.CellBackColor = Val(D(0))
170               g.CellForeColor = Val(D(1))
180           End If
190           tb.MoveNext
200       Loop

210       If g.Rows > 2 Then
220           g.RemoveItem 1
230       End If
240       g.Visible = True

250       Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer

260       intEL = Erl
270       strES = Err.Description
280       LogError "frmListsGeneric", "FillG", intEL, strES, sql
290       g.Visible = True

End Sub

Private Sub FireDown()

          Dim n As Integer
          Dim s As String
          Dim VisibleRows As Integer
          Dim b As Long
          Dim f As Long

10        If g.Row = g.Rows - 1 Then Exit Sub
20        n = g.Row

30        FireCounter = FireCounter + 1
40        If FireCounter > 5 Then
50            tmrDown.Interval = 100
60        End If

70        VisibleRows = g.Height \ g.RowHeight(1) - 1

80        g.Visible = False

90        s = g.TextMatrix(n, 0) & vbTab & g.TextMatrix(n, 1)
100       g.Col = 1
110       b = g.CellBackColor
120       f = g.CellForeColor

130       g.RemoveItem n
140       If n < g.Rows Then
150           g.AddItem s, n + 1
160           g.Row = n + 1
170       Else
180           g.AddItem s
190           g.Row = g.Rows - 1
200       End If

210       g.Col = 0
220       g.CellBackColor = vbYellow

230       g.Col = 1
240       g.CellForeColor = f
250       g.CellBackColor = b

260       If Not g.RowIsVisible(g.Row) Or g.Row = g.Rows - 1 Then
270           If g.Row - VisibleRows + 1 > 0 Then
280               g.TopRow = g.Row - VisibleRows + 1
290           End If
300       End If

310       g.Visible = True

320       cmdSave.Visible = True

End Sub

Private Sub FireUp()

          Dim n As Integer
          Dim s As String
          Dim b As Long
          Dim f As Long

10        If g.Row = 1 Then Exit Sub

20        FireCounter = FireCounter + 1
30        If FireCounter > 5 Then
40            tmrUp.Interval = 100
50        End If

60        n = g.Row

70        g.Visible = False

80        s = g.TextMatrix(n, 0) & vbTab & g.TextMatrix(n, 1)
90        g.Col = 1
100       b = g.CellBackColor
110       f = g.CellForeColor

120       g.RemoveItem n
130       g.AddItem s, n - 1

140       g.Row = n - 1
150       g.Col = 0
160       g.CellBackColor = vbYellow

170       g.Col = 1
180       g.CellForeColor = f
190       g.CellBackColor = b

200       If Not g.RowIsVisible(g.Row) Then
210           g.TopRow = g.Row
220       End If

230       g.Visible = True

240       cmdSave.Visible = True

End Sub



Private Sub cmdadd_Click()

          Dim Y As Integer

10        On Error GoTo cmdadd_Click_Error

20        txtCode = Trim$(UCase$(txtCode))
30        txtText = Trim$(txtText)

40        If txtCode = "" Then
50            Exit Sub
60        End If

70        If txtText = "" Then
80            Exit Sub
90        End If

100       For Y = 1 To g.Rows - 1
110           If g.TextMatrix(Y, 0) = txtCode Then
120               If g.Rows > 2 Then
130                   g.RemoveItem Y
140               Else
150                   g.AddItem ""
160                   g.RemoveItem 1
170               End If
180               Exit For
190           End If
200       Next

210       g.AddItem txtCode & vbTab & txtText
220       g.Col = 1
230       g.Row = g.Rows - 1
240       g.CellBackColor = pBackColour
250       g.CellForeColor = pForeColour

260       If g.TextMatrix(1, 0) = "" Then
270           g.RemoveItem 1
280       End If

290       txtCode = ""
300       txtText = ""
310       pBackColour = cmdBackColour(0).BackColor
320       pForeColour = cmdForeColour(9).BackColor
330       txtText.BackColor = pBackColour
340       txtText.ForeColor = pForeColour

350       cmdSave.Visible = True

360       Exit Sub

cmdadd_Click_Error:

          Dim strES As String
          Dim intEL As Integer

370       intEL = Erl
380       strES = Err.Description
390       LogError "frmListsGenericColour", "cmdadd_Click", intEL, strES

End Sub


Private Sub cmdBackColour_Click(Index As Integer)

10        pBackColour = cmdBackColour(Index).BackColor
20        txtText.BackColor = pBackColour

End Sub

Private Sub cmdCancel_Click()

10        Unload Me

End Sub


Private Sub cmdDelete_Click()

          Dim Y As Integer
          Dim sql As String
          Dim s As String

10        On Error GoTo cmdDelete_Click_Error

20        g.Col = 0
30        For Y = 1 To g.Rows - 1
40            g.Row = Y
50            If g.CellBackColor = vbYellow Then
60                s = "Delete " & g.TextMatrix(Y, 1) & vbCrLf & _
                      "From " & pListTypeNames & " ?"
70                If iMsg(s, vbQuestion + vbYesNo) = vbYes Then
80                    sql = "Delete from Lists where " & _
                            "ListType = '" & pListType & "' " & _
                            "and Code = '" & g.TextMatrix(Y, 0) & "'"
90                    Cnxn(0).Execute sql
100               End If
110               Exit For
120           End If
130       Next

140       cmdDelete.Enabled = False
150       FillG

160       Exit Sub

cmdDelete_Click_Error:

          Dim strES As String
          Dim intEL As Integer

170       intEL = Erl
180       strES = Err.Description
190       LogError "frmListsGeneric", "cmdDelete_Click", intEL, strES, sql

End Sub


Private Sub cmdForeColour_Click(Index As Integer)

10        pForeColour = cmdForeColour(Index).BackColor
20        txtText.ForeColor = pForeColour

End Sub

Private Sub cmdMoveDown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        FireDown

20        tmrDown.Interval = 250
30        FireCounter = 0

40        tmrDown.Enabled = True

End Sub


Private Sub cmdMoveDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        tmrDown.Enabled = False

End Sub


Private Sub cmdMoveUp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        FireUp

20        tmrUp.Interval = 250
30        FireCounter = 0

40        tmrUp.Enabled = True

End Sub


Private Sub cmdMoveUp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        tmrUp.Enabled = False

End Sub


Private Sub cmdPrint_Click()

10        Printer.Print

20        Printer.Print "List of "; pListTypeNames

30        g.Col = 0
40        g.Row = 1
50        g.ColSel = g.Cols - 1
60        g.RowSel = g.Rows - 1

70        Printer.Print g.Clip

80        Printer.EndDoc


End Sub


Private Sub cmdSave_Click()

          Dim Y As Integer
          Dim tb As Recordset
          Dim sql As String
          Dim D As String

10        On Error GoTo cmdSave_Click_Error

20        g.Col = 1
30        For Y = 1 To g.Rows - 1
40            If g.TextMatrix(Y, 0) <> "" Then
50                sql = "SELECT * FROM Lists WHERE " & _
                        "ListType = '" & pListType & "' " & _
                        "AND Code = '" & g.TextMatrix(Y, 0) & "'"
60                Set tb = New Recordset
70                RecOpenServer 0, tb, sql
80                If tb.EOF Then
90                    tb.AddNew
100               End If
110               tb!Code = g.TextMatrix(Y, 0)
120               tb!ListType = pListType
130               tb!Text = g.TextMatrix(Y, 1)
140               tb!ListOrder = Y
150               g.Row = Y
160               D = Format$(g.CellBackColor) & "|" & Format$(g.CellForeColor)
170               tb!Default = D
180               tb!InUse = 1
190               tb.Update
200           End If
210       Next

220       FillG

230       txtCode = ""
240       txtText = ""
250       txtText.BackColor = cmdBackColour(1).BackColor
260       txtText.ForeColor = cmdForeColour(9).BackColor
270       txtCode.SetFocus
280       cmdMoveUp.Enabled = False
290       cmdMoveDown.Enabled = False
300       cmdSave.Visible = False
310       cmdDelete.Enabled = False

320       Exit Sub

cmdSave_Click_Error:

          Dim strES As String
          Dim intEL As Integer

330       intEL = Erl
340       strES = Err.Description
350       LogError "frmListsGeneric", "cmdSave_Click", intEL, strES, sql

End Sub


Private Sub cmdXL_Click()

10        ExportFlexGrid g, Me

End Sub


Private Sub Form_Activate()

10        If Activated Then Exit Sub

20        Activated = True

30        FillG

End Sub

Private Sub Form_Load()

10        g.Font.Bold = True

20        If pListType = "" Then
30            MsgBox "pListType not set"
40        End If
50        If pListTypeName = "" Then
60            MsgBox "pListTypeName not set"
70        End If
80        If pListTypeNames = "" Then
90            MsgBox "pListTypeNames not set"
100       End If

110       FrameAdd.Caption = "Add New " & pListTypeName
120       Me.Caption = "NetAcquire - List of " & pListTypeNames

130       pBackColour = cmdBackColour(1).BackColor
140       pForeColour = cmdForeColour(9).BackColor

150       Activated = False

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

10        If cmdSave.Visible Then
20            If iMsg("Cancel without Saving?", vbQuestion + vbYesNo) = vbNo Then
30                Cancel = True
40                Exit Sub
50            End If
60        End If

End Sub


Private Sub Form_Unload(Cancel As Integer)

10        pListType = ""
20        pListTypeName = ""
30        pListTypeNames = ""

End Sub

Private Sub g_Click()

          Static SortOrder As Boolean
          Dim Y As Integer
          Dim ySave As Integer

10        If g.MouseRow = 0 Then
20            If SortOrder Then
30                g.Sort = flexSortGenericAscending
40            Else
50                g.Sort = flexSortGenericDescending
60            End If
70            SortOrder = Not SortOrder
80            Exit Sub
90        End If

100       ySave = g.Row

          'Put the entry up ready to edit
110       txtCode = g.TextMatrix(g.Row, 0)
120       txtText = g.TextMatrix(g.Row, 1)
130       g.Col = 1
140       txtText.BackColor = g.CellBackColor
150       pBackColour = g.CellBackColor
160       txtText.ForeColor = g.CellForeColor
170       pForeColour = g.CellForeColor

          'Deselect Yellow on all rows
180       g.Visible = False
190       g.Col = 0
200       For Y = 1 To g.Rows - 1
210           g.Row = Y
220           If g.CellBackColor = vbYellow Then
230               g.CellBackColor = 0
240               Exit For
250           End If
260       Next
270       g.Row = ySave
280       g.Visible = True

290       g.CellBackColor = vbYellow

300       cmdMoveUp.Enabled = True
310       cmdMoveDown.Enabled = True
320       cmdDelete.Enabled = True

End Sub



Private Sub tmrDown_Timer()

10        FireDown

End Sub


Private Sub tmrUp_Timer()

10        FireUp

End Sub



Public Property Let ListType(ByVal strNewValue As String)

10        pListType = strNewValue

End Property
Public Property Let ListTypeName(ByVal strNewValue As String)

10        pListTypeName = strNewValue

End Property

Public Property Let ListTypeNames(ByVal strNewValue As String)

10        pListTypeNames = strNewValue

End Property

Private Sub txtText_Change()

10        txtText.ForeColor = pForeColour
20        txtText.BackColor = pBackColour

End Sub


