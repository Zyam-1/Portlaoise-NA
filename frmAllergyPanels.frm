VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAllergyPanels 
   Caption         =   "NetAcquire - Allergy Panels"
   ClientHeight    =   6885
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11340
   LinkTopic       =   "Form1"
   ScaleHeight     =   6885
   ScaleWidth      =   11340
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   6525
      Left            =   90
      TabIndex        =   4
      Top             =   150
      Width           =   4275
      _ExtentX        =   7541
      _ExtentY        =   11509
      _Version        =   393216
      Cols            =   3
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
      AllowUserResizing=   1
      FormatString    =   "<Code    |<Short Name    |<Long Name                            "
   End
   Begin VB.CommandButton cmdAddItem 
      Caption         =   "&Add Item"
      Enabled         =   0   'False
      Height          =   975
      Left            =   4560
      Picture         =   "frmAllergyPanels.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1650
      Width           =   1155
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   975
      Left            =   9990
      Picture         =   "frmAllergyPanels.frx":1982
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5730
      Width           =   1155
   End
   Begin VB.CommandButton cmdRemoveItem 
      Caption         =   "&Remove Item"
      Enabled         =   0   'False
      Height          =   975
      Left            =   4560
      Picture         =   "frmAllergyPanels.frx":3304
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2730
      Width           =   1155
   End
   Begin MSComctlLib.TreeView tvPanels 
      Height          =   6495
      Left            =   5910
      TabIndex        =   0
      Top             =   180
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   11456
      _Version        =   393217
      HideSelection   =   0   'False
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   6
      SingleSel       =   -1  'True
      BorderStyle     =   1
      Appearance      =   1
   End
End
Attribute VB_Name = "frmAllergyPanels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub FillList()

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo FillList_Error

20        g.Rows = 2
30        g.AddItem ""
40        g.RemoveItem 1

50        sql = "SELECT DISTINCT Code, ShortName, LongName FROM ImmTestDefinitions WHERE " & _
                "SUBSTRING(Code, 2, 1) <> 'X' " & _
                "ORDER BY ShortName"
60        Set tb = New Recordset
70        RecOpenServer 0, tb, sql
80        Do While Not tb.EOF
90            g.AddItem tb!Code & vbTab & _
                        tb!ShortName & vbTab & _
                        tb!LongName & ""
100           tb.MoveNext
110       Loop

120       If g.Rows > 2 Then
130           g.RemoveItem 1
140       End If

150       Exit Sub

FillList_Error:

          Dim strES As String
          Dim intEL As Integer

160       intEL = Erl
170       strES = Err.Description
180       LogError "frmAllergyPanels", "FillList", intEL, strES, sql

End Sub

Private Sub FillTree()

          Dim NodX As MSComctlLib.Node
          Dim Key As Long
          Dim PanelName As String
          Dim tb As New Recordset
          Dim tc As Recordset
          Dim sql As String

10        On Error GoTo FillTree_Error

20        tvPanels.Nodes.Clear

30        sql = "SELECT LongName, PrintPriority FROM ImmTestDefinitions WHERE " & _
                "SUBSTRING(Code, 2, 1) = 'X' " & _
                "AND Hospital = '" & HospName(0) & "' " & _
                "ORDER BY PrintPriority"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        Do While Not tb.EOF
70            PanelName = tb!LongName & ""

80            Set NodX = tvPanels.Nodes.Add(, , "key" & CStr(Key), PanelName)
90            Key = Key + 1

100           sql = "SELECT Content FROM IPanels WHERE " & _
                    "PanelType = 'AL' " & _
                    "AND PanelName = '" & PanelName & "' " & _
                    "AND Hospital = '" & HospName(0) & "'"
110           Set tc = New Recordset
120           RecOpenServer 0, tc, sql
130           Do While Not tc.EOF
140               Set NodX = tvPanels.Nodes.Add("key" & CStr(Key - 1), tvwChild, , tc!Content & "")
150               tc.MoveNext
160           Loop

170           tb.MoveNext
180       Loop

190       Exit Sub

FillTree_Error:

          Dim strES As String
          Dim intEL As Integer

200       intEL = Erl
210       strES = Err.Description
220       LogError "frmAllergyPanels", "FillTree", intEL, strES, sql

End Sub

Private Sub cmdAddItem_Click()

          Dim tb As Recordset
          Dim sql As String
          Dim Content As String
          Dim PanelName As String
          Dim NodX As MSComctlLib.Node

10        On Error GoTo cmdAddItem_Click_Error

20        Content = g.TextMatrix(g.Row, 2)
30        PanelName = tvPanels.SelectedItem.Text

40        sql = "SELECT * FROM IPanels WHERE " & _
                "PanelType = 'AL' " & _
                "AND Hospital = '" & HospName(0) & "' " & _
                "AND Content = '" & Content & "' " & _
                "AND PanelName = '" & PanelName & "'"
50        Set tb = New Recordset
60        RecOpenServer 0, tb, sql
70        If Not tb.EOF Then
80            Exit Sub
90        End If

100       tb.AddNew
110       tb!PanelName = PanelName
120       tb!Content = Content
130       tb!PanelType = "AL"
140       tb!Hospital = HospName(0)
150       tb!ListOrder = 999
160       tb.Update

170       FillTree

180       For Each NodX In tvPanels.Nodes
190           If NodX.Parent Is Nothing Then
200               If NodX.Text = PanelName Then
210                   If NodX.Children > 0 Then
220                       NodX.Child.EnsureVisible
230                       NodX.Selected = True
240                       Exit For
250                   End If
260               End If
270           End If
280       Next

290       cmdAddItem.Enabled = False

300       Exit Sub

cmdAddItem_Click_Error:

          Dim strES As String
          Dim intEL As Integer

310       intEL = Erl
320       strES = Err.Description
330       LogError "frmAllergyPanels", "cmdAddItem_Click", intEL, strES, sql

End Sub

Private Sub cmdCancel_Click()

10        Unload Me

End Sub


Private Sub cmdRemoveItem_Click()

          Dim sql As String
          Dim Content As String
          Dim PanelName As String
          Dim NodX As MSComctlLib.Node

10        On Error GoTo cmdRemoveItem_Click_Error

20        Content = tvPanels.SelectedItem.Text
30        PanelName = tvPanels.SelectedItem.Parent.Text

40        sql = "DELETE FROM IPanels WHERE " & _
                "PanelType = 'AL' " & _
                "AND Hospital = '" & HospName(0) & "' " & _
                "AND Content = '" & Content & "' " & _
                "AND PanelName = '" & PanelName & "'"
50        Cnxn(0).Execute sql

60        FillTree

70        For Each NodX In tvPanels.Nodes
80            If NodX.Parent Is Nothing Then
90                If NodX.Text = PanelName Then
100                   If NodX.Children > 0 Then
110                       NodX.Child.EnsureVisible
120                       NodX.Selected = True
130                       Exit For
140                   End If
150               End If
160           End If
170       Next

180       cmdRemoveItem.Enabled = False

190       Exit Sub

cmdRemoveItem_Click_Error:

          Dim strES As String
          Dim intEL As Integer

200       intEL = Erl
210       strES = Err.Description
220       LogError "frmAllergyPanels", "cmdRemoveItem_Click", intEL, strES, sql

End Sub

Private Sub Form_Load()

10        FillList
20        FillTree

End Sub


Private Sub g_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

          Static SortOrder As Boolean
          Dim SaveY As Integer
          Dim Found As Boolean

10        On Error GoTo g_MouseUp_Error

20        If g.MouseRow = 0 Then
30            If SortOrder Then
40                g.Sort = flexSortGenericAscending
50            Else
60                g.Sort = flexSortGenericDescending
70            End If
80            SortOrder = Not SortOrder
90            Exit Sub
100       End If

110       SaveY = g.Row

120       Found = False
130       g.Col = 1
140       For Y = 1 To g.Rows - 1
150           g.Row = Y
160           If g.CellBackColor = vbYellow Then
170               For x = 1 To g.Cols - 1
180                   g.Col = x
190                   g.CellBackColor = 0
200               Next
210               Found = True
220               Exit For
230           End If
240           If Found Then
250               Exit For
260           End If
270       Next

280       g.Row = SaveY
290       For x = 1 To g.Cols - 1
300           g.Col = x
310           g.CellBackColor = vbYellow
320       Next
330       If Not tvPanels.SelectedItem Is Nothing Then
340           If tvPanels.SelectedItem.Parent Is Nothing Then
350               cmdAddItem.Enabled = True
360           End If
370       End If

380       cmdRemoveItem.Enabled = False

390       Exit Sub

g_MouseUp_Error:

          Dim strES As String
          Dim intEL As Integer

400       intEL = Erl
410       strES = Err.Description
420       LogError "frmAllergyPanels", "g_MouseUp", intEL, strES

End Sub


Private Sub tvPanels_Click()

10        If tvPanels.Nodes.Count = 0 Then Exit Sub

20        If Not tvPanels.SelectedItem.Parent Is Nothing Then
30            cmdAddItem.Enabled = False
40            cmdRemoveItem.Enabled = True
50        Else
60            cmdRemoveItem.Enabled = False
70        End If

End Sub

