VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmListsGeneric 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - List of Generic"
   ClientHeight    =   8472
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   12888
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8472
   ScaleWidth      =   12888
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPrint 
      Appearance      =   0  'Flat
      Caption         =   "&Print"
      Height          =   825
      Left            =   11610
      Picture         =   "frmListsGeneric.frx":0000
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
      Picture         =   "frmListsGeneric.frx":1982
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
      Begin VB.TextBox txtCode 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
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
            Size            =   7.8
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
         Width           =   8535
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   765
         Left            =   10350
         Picture         =   "frmListsGeneric.frx":3304
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
      Picture         =   "frmListsGeneric.frx":4C86
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
      Picture         =   "frmListsGeneric.frx":6608
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
      Picture         =   "frmListsGeneric.frx":7F8A
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
      Picture         =   "frmListsGeneric.frx":990C
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
      Picture         =   "frmListsGeneric.frx":B28E
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
      _ExtentX        =   19600
      _ExtentY        =   12552
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
      FormatString    =   $"frmListsGeneric.frx":B598
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
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
         Size            =   7.8
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
Attribute VB_Name = "frmListsGeneric"
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



Private Sub FillG()

          Dim s As String
          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo FillG_Error

20        g.Visible = False
30        g.Rows = 2
40        g.AddItem ""
50        g.RemoveItem 1

60        sql = "SELECT Code, Text FROM Lists WHERE " & _
                "ListType = '" & pListType & "' " & _
                "ORDER BY ListOrder"
70        Set tb = New Recordset
80        RecOpenServer 0, tb, sql
90        Do While Not tb.EOF
100           s = tb!Code & vbTab & tb!Text & ""
110           g.AddItem s
120           tb.MoveNext
130       Loop

140       If g.Rows > 2 Then
150           g.RemoveItem 1
160       End If
170       g.Visible = True

180       Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer

190       intEL = Erl
200       strES = Err.Description
210       LogError "frmListsGeneric", "FillG", intEL, strES, sql
220       g.Visible = True

End Sub

Private Sub FireDown()

          Dim n As Integer
          Dim s As String
          Dim X As Integer
          Dim VisibleRows As Integer

10        If g.Row = g.Rows - 1 Then Exit Sub
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

220       For X = 0 To g.Cols - 1
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

Private Sub FireUp()

          Dim n As Integer
          Dim s As String
          Dim X As Integer

10        If g.Row = 1 Then Exit Sub

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
160       For X = 0 To g.Cols - 1
170           g.Col = X
180           g.CellBackColor = vbYellow
190       Next

200       If Not g.RowIsVisible(g.Row) Then
210           g.TopRow = g.Row
220       End If

230       g.Visible = True

240       cmdSave.Visible = True

End Sub



Private Sub cmdadd_Click()

10        txtCode = Trim$(UCase$(txtCode))
20        txtText = Trim$(txtText)

30        If txtCode = "" Then
40            Exit Sub
50        End If

60        If txtText = "" Then
70            Exit Sub
80        End If

90        g.AddItem txtCode & vbTab & txtText

100       txtCode = ""
110       txtText = ""

120       cmdSave.Visible = True

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
          Dim sql As String

10        On Error GoTo cmdSave_Click_Error

20        For Y = 1 To g.Rows - 1
30            If g.TextMatrix(Y, 0) <> "" Then
                  'Created on 01/02/2011 13:19:05
                  'Autogenerated by SQL Scripting

40                sql = "If Exists(Select 1 From Lists " & _
                        "Where Code = '@Code0' " & _
                        "And ListType = '@ListType2' ) " & _
                        "Begin " & _
                        "Update Lists Set " & _
                        "Code = '@Code0', " & _
                        "Text = '@Text1', " & _
                        "ListType = '@ListType2', " & _
                        "ListOrder = @ListOrder3, " & _
                        "InUse = @InUse4 " & _
                        "Where Code = '@Code0' " & _
                        "And ListType = '@ListType2' " & _
                        "End  " & _
                        "Else " & _
                        "Begin  " & _
                        "Insert Into Lists (Code, Text, ListType, ListOrder, InUse) Values " & _
                        "('@Code0', '@Text1', '@ListType2', @ListOrder3, @InUse4) " & _
                        "End"

50                sql = Replace(sql, "@Code0", g.TextMatrix(Y, 0))
60                sql = Replace(sql, "@Text1", AddTicks(g.TextMatrix(Y, 1)))
70                sql = Replace(sql, "@ListType2", pListType)
80                sql = Replace(sql, "@ListOrder3", Y)
90                sql = Replace(sql, "@InUse4", 1)

100               Cnxn(0).Execute sql

                  '        sql = "SELECT * FROM Lists WHERE " & _
                           '              "ListType = '" & pListType & "' " & _
                           '              "AND Code = '" & g.TextMatrix(Y, 0) & "'"
                  '        Set tb = New Recordset
                  '        RecOpenServer 0, tb, sql
                  '        If tb.EOF Then
                  '            tb.AddNew
                  '        End If
                  '        tb!Code = g.TextMatrix(Y, 0)
                  '        tb!ListType = pListType
                  '        tb!Text = g.TextMatrix(Y, 1)
                  '        tb!ListOrder = Y
                  '        tb!InUse = 1
                  '        tb.Update
110           End If
120       Next

130       FillG

140       txtCode = ""
150       txtText = ""
160       txtCode.SetFocus
170       cmdMoveUp.Enabled = False
180       cmdMoveDown.Enabled = False
190       cmdSave.Visible = False
200       cmdDelete.Enabled = False

210       Exit Sub

cmdSave_Click_Error:

          Dim strES As String
          Dim intEL As Integer

220       intEL = Erl
230       strES = Err.Description
240       LogError "frmListsGeneric", "cmdSave_Click", intEL, strES, sql

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

130       Activated = False

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

290       cmdMoveUp.Enabled = True
300       cmdMoveDown.Enabled = True
310       cmdDelete.Enabled = True

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

