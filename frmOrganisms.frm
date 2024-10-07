VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmOrganisms 
   Caption         =   "NetAcquire - Organisms"
   ClientHeight    =   8940
   ClientLeft      =   75
   ClientTop       =   450
   ClientWidth     =   12510
   LinkTopic       =   "Form1"
   ScaleHeight     =   8940
   ScaleWidth      =   12510
   Begin VB.CommandButton cmdXL 
      Caption         =   "Export to Excel"
      Height          =   825
      Left            =   11280
      Picture         =   "frmOrganisms.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2940
      Width           =   825
   End
   Begin VB.Timer tmrUp 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   9240
      Top             =   0
   End
   Begin VB.Timer tmrDown 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   9720
      Top             =   0
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   825
      Left            =   11280
      Picture         =   "frmOrganisms.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1770
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.CommandButton bMoveUp 
      Caption         =   "Move Up"
      Enabled         =   0   'False
      Height          =   825
      Left            =   10920
      Picture         =   "frmOrganisms.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4710
      Width           =   795
   End
   Begin VB.CommandButton bMoveDown 
      Caption         =   "Move Down"
      Enabled         =   0   'False
      Height          =   825
      Left            =   10920
      Picture         =   "frmOrganisms.frx":0B8E
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5550
      Width           =   795
   End
   Begin MSFlexGridLib.MSFlexGrid grdOrg 
      Height          =   5775
      Left            =   1080
      TabIndex        =   6
      Top             =   3030
      Width           =   9795
      _ExtentX        =   17277
      _ExtentY        =   10186
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      FormatString    =   $"frmOrganisms.frx":0FD0
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H0000FFFF&
      Caption         =   "&Save Details"
      Height          =   825
      Left            =   11280
      Picture         =   "frmOrganisms.frx":107E
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6930
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   825
      Left            =   11280
      Picture         =   "frmOrganisms.frx":16E8
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7920
      Width           =   795
   End
   Begin VB.CommandButton cmdNewGroup 
      Caption         =   "Add New Group"
      Height          =   315
      Left            =   5430
      TabIndex        =   4
      Top             =   180
      Width           =   1485
   End
   Begin VB.ComboBox cmbGroups 
      Height          =   315
      Left            =   2370
      TabIndex        =   2
      Text            =   "cmbGroups"
      Top             =   180
      Width           =   2985
   End
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   150
      TabIndex        =   12
      Top             =   660
      Width           =   10695
      Begin VB.CommandButton cmdAddSite 
         Caption         =   "Add New Site"
         Height          =   315
         Left            =   3990
         TabIndex        =   24
         Top             =   1860
         Width           =   1485
      End
      Begin VB.ComboBox cmbSite 
         Height          =   315
         Left            =   975
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   1860
         Width           =   2985
      End
      Begin VB.TextBox txtReportName 
         Height          =   285
         Left            =   975
         TabIndex        =   20
         Top             =   1440
         Width           =   1515
      End
      Begin VB.TextBox txtNewOrg 
         Height          =   285
         Left            =   975
         MaxLength       =   100
         TabIndex        =   16
         Top             =   630
         Width           =   9570
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   825
         Left            =   9750
         Picture         =   "frmOrganisms.frx":1D52
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1350
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.TextBox txtCode 
         Height          =   285
         Left            =   975
         TabIndex        =   14
         Top             =   225
         Width           =   1515
      End
      Begin VB.TextBox txtShortName 
         Height          =   285
         Left            =   975
         MaxLength       =   15
         TabIndex        =   13
         Top             =   1035
         Width           =   1515
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Site"
         Height          =   195
         Left            =   600
         TabIndex        =   23
         Top             =   1920
         Width           =   270
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Report Name"
         Height          =   375
         Left            =   345
         TabIndex        =   21
         Top             =   1395
         Width           =   570
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "New Organism"
         Height          =   390
         Left            =   225
         TabIndex        =   19
         Top             =   577
         Width           =   690
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Code"
         Height          =   195
         Left            =   540
         TabIndex        =   18
         Top             =   270
         Width           =   375
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Short Name"
         Height          =   195
         Left            =   75
         TabIndex        =   17
         Top             =   1080
         Width           =   840
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
      Left            =   11040
      TabIndex        =   11
      Top             =   3780
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Members of this Group"
      Height          =   435
      Left            =   135
      TabIndex        =   3
      Top             =   3030
      Width           =   870
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Organism Group"
      Height          =   195
      Left            =   1140
      TabIndex        =   1
      Top             =   210
      Width           =   1140
   End
End
Attribute VB_Name = "frmOrganisms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FireCounter As Integer

Private Sub FireDown()

          Dim n As Integer
          Dim s As String
          Dim X As Integer
          Dim VisibleRows As Integer

10        With grdOrg
20            If .Row = .Rows - 1 Then Exit Sub
30            n = .Row

40            FireCounter = FireCounter + 1
50            If FireCounter > 5 Then
60                tmrDown.Interval = 100
70            End If

80            VisibleRows = .Height \ .RowHeight(1) - 1

90            .Visible = False

100           s = ""
110           For X = 0 To .Cols - 1
120               s = s & .TextMatrix(n, X) & vbTab
130           Next
140           s = Left$(s, Len(s) - 1)

150           .RemoveItem n
160           If n < .Rows Then
170               .AddItem s, n + 1
180               .Row = n + 1
190           Else
200               .AddItem s
210               .Row = .Rows - 1
220           End If

230           For X = 0 To .Cols - 1
240               .Col = X
250               .CellBackColor = vbYellow
260           Next

270           If Not .RowIsVisible(.Row) Or .Row = .Rows - 1 Then
280               If .Row - VisibleRows + 1 > 0 Then
290                   .TopRow = .Row - VisibleRows + 1
300               End If
310           End If

320           .Visible = True
330       End With

340       cmdSave.Visible = True

End Sub
Private Sub FireUp()

          Dim n As Integer
          Dim s As String
          Dim X As Integer

10        With grdOrg
20            If .Row = 1 Then Exit Sub

30            FireCounter = FireCounter + 1
40            If FireCounter > 5 Then
50                tmrUp.Interval = 100
60            End If

70            n = .Row

80            .Visible = False

90            s = ""
100           For X = 0 To .Cols - 1
110               s = s & .TextMatrix(n, X) & vbTab
120           Next
130           s = Left$(s, Len(s) - 1)

140           .RemoveItem n
150           .AddItem s, n - 1

160           .Row = n - 1
170           For X = 0 To .Cols - 1
180               .Col = X
190               .CellBackColor = vbYellow
200           Next

210           If Not .RowIsVisible(.Row) Then
220               .TopRow = .Row
230           End If

240           .Visible = True

250           cmdSave.Visible = True
260       End With

End Sub





Private Sub FillGroups()

          Dim tb As Recordset
          Dim sql As String

10        cmbGroups.Clear

20        sql = "Select * from Lists where " & _
                "ListType = 'OR' " & _
                "order by ListOrder"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        Do While Not tb.EOF
60            cmbGroups.AddItem tb!Text & ""
70            tb.MoveNext
80        Loop

End Sub

Private Sub FillSites()

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo FillSites_Error


20        cmbSite.Clear
30        cmbSite.AddItem "Generic"
40        sql = "SELECT Text FROM Lists WHERE " & _
                "ListType = 'SI' " & _
                "ORDER BY ListOrder"
50        Set tb = New Recordset
60        RecOpenServer 0, tb, sql
70        Do While Not tb.EOF
80            cmbSite.AddItem tb!Text & ""
90            tb.MoveNext
100       Loop

110       cmbSite.ListIndex = 0

120       Exit Sub

FillSites_Error:

          Dim strES As String
          Dim intEL As Integer

130       intEL = Erl
140       strES = Err.Description
150       LogError "frmOrganisms", "FillSites", intEL, strES, sql

End Sub

Private Sub FillList()

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo FillList_Error

20        grdOrg.Rows = 2
30        grdOrg.AddItem ""
40        grdOrg.RemoveItem 1

50        sql = "Select * from Organisms where " & _
                "GroupName = '" & cmbGroups.Text & "' " & _
                "order by listorder"
60        Set tb = New Recordset
70        RecOpenClient 0, tb, sql
80        Do While Not tb.EOF
90            grdOrg.AddItem tb!Code & vbTab & _
                             tb!Name & vbTab & _
                             tb!ShortName & vbTab & _
                             tb!ReportName & "" & vbTab & _
                             tb!Site & ""
100           tb.MoveNext
110       Loop

120       If grdOrg.Rows > 2 Then
130           grdOrg.RemoveItem 1
140       End If

150       Exit Sub

FillList_Error:

          Dim strES As String
          Dim intEL As Integer

160       intEL = Erl
170       strES = Err.Description
180       LogError "frmOrganisms", "FillList", intEL, strES, sql

End Sub

Private Sub SaveDetails()

          Dim tb As Recordset
          Dim sql As String
          Dim n As Integer

10        On Error GoTo SaveDetails_Error

20        Cnxn(0).BeginTrans
          '  sql = "Delete from Organisms where " & _
             '        "GroupName = '" & cmbGroups.Text & "'"
          '  Cnxn(0).Execute sql

30        For n = 1 To grdOrg.Rows - 1
40            If Trim$(grdOrg.TextMatrix(n, 1)) <> "" Then
50                sql = "Select * from Organisms where " & _
                        "Name = '" & grdOrg.TextMatrix(n, 1) & "' " & _
                        "and GroupName = '" & cmbGroups.Text & "'"
60                Set tb = New Recordset
70                RecOpenClient 0, tb, sql
80                If tb.EOF Then tb.AddNew
90                tb!Code = grdOrg.TextMatrix(n, 0)
100               tb!Name = grdOrg.TextMatrix(n, 1)
110               tb!ShortName = grdOrg.TextMatrix(n, 2)
120               tb!ReportName = grdOrg.TextMatrix(n, 3)
130               tb!GroupName = cmbGroups.Text
140               tb!Site = IIf(grdOrg.TextMatrix(n, 4) = "", "Generic", grdOrg.TextMatrix(n, 4))
150               tb!ListOrder = n
160               tb.Update
170           End If
180       Next
190       Cnxn(0).CommitTrans

200       FillList
210       cmdSave.Visible = False

220       Exit Sub

SaveDetails_Error:

          Dim strES As String
          Dim intEL As Integer

230       intEL = Erl
240       strES = Err.Description
250       LogError "frmOrganisms", "SaveDetails", intEL, strES, sql

End Sub

Private Sub bMoveDown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        FireDown

20        tmrDown.Interval = 250
30        FireCounter = 0

40        tmrDown.Enabled = True

End Sub


Private Sub bMoveDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        tmrDown.Enabled = False

End Sub


Private Sub bMoveUp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        FireUp

20        tmrUp.Interval = 250
30        FireCounter = 0

40        tmrUp.Enabled = True

End Sub


Private Sub bMoveUp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        tmrUp.Enabled = False

End Sub


Private Sub cmbGroups_Click()

10        FillList

End Sub

Private Sub cmbGroups_GotFocus()

10        If cmdSave.Visible Then
20            If iMsg("Save Changes?", vbQuestion + vbYesNo) = vbYes Then
30                SaveDetails
40            End If
50            cmdSave.Visible = False
60        End If

End Sub


Private Sub cmdadd_Click()

10        grdOrg.AddItem UCase$(txtCode) & vbTab & _
                         txtNewOrg & vbTab & _
                         txtShortName & vbTab & _
                         txtReportName & vbTab & _
                         cmbSite
20        txtNewOrg = ""
30        txtCode = ""
40        txtShortName = ""
50        txtReportName = ""


60        cmdSave.Visible = True
70        txtNewOrg.SetFocus

End Sub

Private Sub cmdAddSite_Click()

10        On Error GoTo cmdAddSite_Click_Error

20        frmMicroSites.Show 1
30        FillSites

40        Exit Sub

cmdAddSite_Click_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmOrganisms", "cmdAddSite_Click", intEL, strES


End Sub

Private Sub cmdCancel_Click()

10        If cmdSave.Visible Then
20            If iMsg("Cancel without Saving?", vbQuestion + vbYesNo) = vbNo Then
30                SaveDetails
40            End If
50            cmdSave.Visible = False
60        End If

70        Unload Me

End Sub


Private Sub cmdDelete_Click()

10        If grdOrg.Rows = 2 Then
20            grdOrg.AddItem ""
30            grdOrg.RemoveItem 1
40        Else
50            grdOrg.RemoveItem grdOrg.Row
60        End If
70        cmdDelete.Visible = False
80        cmdSave.Visible = True
90        cmdSave.SetFocus

End Sub

Private Sub cmdNewGroup_Click()

10        With frmListsGeneric
20            .ListType = "OR"
30            .ListTypeName = "Organism Group"
40            .ListTypeNames = "Organism Groups"
50            .Show 1
60        End With

70        FillGroups

End Sub

Private Sub cmdSave_Click()

10        SaveDetails

End Sub

Private Sub cmdXL_Click()

10        ExportFlexGrid grdOrg, Me

End Sub

Private Sub Form_DblClick()
'
'Dim tb As Recordset
'Dim sql As String
'Dim x() As String
'Dim s As String
'
'sql = "Select * from Organisms"
'Set tb = New Recordset
'RecOpenServer 0, tb, sql
'Do While Not tb.EOF
'  x = Split(tb!Name & "", " ")
'  If UBound(x) = 1 Then
'    s = LCase$(Left$(x(0), 1)) & "." & _
     '        Left$(x(1), 16)
'    tb!ShortName = s
'    tb.Update
'  End If
'  tb.MoveNext
'Loop
'
End Sub

Private Sub Form_Load()

10        FillGroups
20        FillSites
End Sub

Private Sub grdOrg_Click()

          Static SortOrder As Boolean
          Dim X As Integer
          Dim Y As Integer
          Dim ySave As Integer
          Dim xSave As Integer

10        ySave = grdOrg.Row
20        xSave = grdOrg.Col

30        grdOrg.Visible = False
40        grdOrg.Col = 0
50        For Y = 1 To grdOrg.Rows - 1
60            grdOrg.Row = Y
70            If grdOrg.CellBackColor = vbYellow Then
80                For X = 0 To grdOrg.Cols - 1
90                    grdOrg.Col = X
100                   grdOrg.CellBackColor = 0
110               Next
120               Exit For
130           End If
140       Next
150       grdOrg.Row = ySave

160       grdOrg.Visible = True

170       If grdOrg.MouseRow = 0 Then
180           If SortOrder Then
190               grdOrg.Sort = flexSortGenericAscending
200           Else
210               grdOrg.Sort = flexSortGenericDescending
220           End If
230           SortOrder = Not SortOrder
240           Exit Sub
250       End If

260       For X = 0 To grdOrg.Cols - 1
270           grdOrg.Col = X
280           grdOrg.CellBackColor = vbYellow
290       Next

300       grdOrg.Col = xSave
310       If grdOrg.Col = 2 Then
320           grdOrg.Enabled = False
330           grdOrg = iBOX("Short Name", , grdOrg)
340           grdOrg.Enabled = True
350           cmdSave.Enabled = True
360           cmdSave.Visible = True
370       ElseIf grdOrg.Col = 3 Then
380           grdOrg.Enabled = False
390           grdOrg = iBOX("Report Name", , grdOrg)
400           grdOrg.Enabled = True
410           cmdSave.Enabled = True
420           cmdSave.Visible = True
430       End If

440       bMoveUp.Enabled = True
450       bMoveDown.Enabled = True
460       cmdDelete.Visible = True

End Sub

Private Sub tmrDown_Timer()

10        FireDown

End Sub

Private Sub tmrUp_Timer()

10        FireUp

End Sub


Private Sub txtCode_KeyUp(KeyCode As Integer, Shift As Integer)

10        cmdAdd.Visible = False

20        If Trim$(txtCode) <> "" Then
30            If Trim$(txtNewOrg) <> "" Then
40                cmdAdd.Visible = True
50            End If
60        End If

End Sub


Private Sub txtNewOrg_KeyUp(KeyCode As Integer, Shift As Integer)

10        cmdAdd.Visible = False

20        If Trim$(txtCode) <> "" Then
30            If Trim$(txtNewOrg) <> "" Then
40                cmdAdd.Visible = True
50            End If
60        End If

End Sub


