VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPanels 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Define Biochemistry Panels"
   ClientHeight    =   7665
   ClientLeft      =   510
   ClientTop       =   1230
   ClientWidth     =   11610
   ForeColor       =   &H80000008&
   Icon            =   "frmPanels.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7665
   ScaleWidth      =   11610
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdRemoveItem 
      Caption         =   "&Remove Item"
      Enabled         =   0   'False
      Height          =   975
      Left            =   5160
      Picture         =   "frmPanels.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2850
      Width           =   1155
   End
   Begin VB.CommandButton cmdAddItem 
      Caption         =   "&Add Item"
      Enabled         =   0   'False
      Height          =   975
      Left            =   5160
      Picture         =   "frmPanels.frx":1C8C
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1770
      Width           =   1155
   End
   Begin VB.CommandButton cmdRemovePanel 
      Caption         =   "Remove Panel"
      Enabled         =   0   'False
      Height          =   975
      Left            =   9705
      Picture         =   "frmPanels.frx":360E
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2850
      Width           =   1155
   End
   Begin VB.CommandButton cmdAddNewPanel 
      Caption         =   "Add New Panel"
      Height          =   975
      Left            =   9705
      Picture         =   "frmPanels.frx":4F90
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1770
      Width           =   1155
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   975
      Left            =   9705
      Picture         =   "frmPanels.frx":6912
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6375
      Width           =   1155
   End
   Begin MSComctlLib.TreeView Tree 
      Height          =   7080
      Left            =   6420
      TabIndex        =   2
      Top             =   270
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   12488
      _Version        =   393217
      HideSelection   =   0   'False
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   6
      SingleSel       =   -1  'True
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ComboBox cmbSampleType 
      Height          =   315
      Left            =   9270
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   270
      Width           =   2145
   End
   Begin VB.ListBox List1 
      Columns         =   3
      DragIcon        =   "frmPanels.frx":8294
      Height          =   7080
      Left            =   210
      TabIndex        =   0
      Top             =   270
      Width           =   4815
   End
End
Attribute VB_Name = "frmPanels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mDepartment As String

Private TableName As String
Private DefTableName As String

Private Sub cmdAddNewPanel_Click()

          Dim NodX As MSComctlLib.Node
          Static k As String
          Dim NewName As String

10        On Error GoTo cmdAddNewPanel_Click_Error

20        NewName = Trim$(iBOX("Enter New PanelName."))
30        If NewName = "" Then
40            Exit Sub
50        End If

60        If k = "" Then
70            k = "1"
80        Else
90            k = CStr(Val(k) + 1)
100       End If

110       Set NodX = Tree.Nodes.Add(, , "newkey" & k, NewName)
120       NodX.Selected = True
130       NodX.EnsureVisible

140       Exit Sub

cmdAddNewPanel_Click_Error:

          Dim strES As String
          Dim intEL As Integer

150       intEL = Erl
160       strES = Err.Description
170       LogError "frmPanels", "cmdAddNewPanel_Click", intEL, strES

End Sub

Private Sub FillList()

          Dim tb As New Recordset
          Dim sql As String
          Dim SampleType As String

10        On Error GoTo FillList_Error

20        List1.Clear

30        SampleType = ListCodeFor("ST", cmbSampleType)

      '  Old Code  sql = "SELECT distinct longname, PrintPriority from " & DefTableName & " WHERE " & _
      '          "SampleType = '" & SampleType & "' " & _
      '          "and Hospital = '" & HospName(0) & "' and inuse = '1' " & _
      '          "Order by PrintPriority"
                
      ' Fix issue for Panels Trevor Dunican
40       sql = "SELECT distinct longname, PrintPriority from " & DefTableName & " WHERE " & _
                "SampleType = '" & SampleType & "' " & _
                "and inuse = '1' " & _
                "Order by PrintPriority"
                
50        Set tb = New Recordset
60        RecOpenServer 0, tb, sql
70        Do While Not tb.EOF
80            List1.AddItem tb!LongName & ""
90            tb.MoveNext
100       Loop

110       Exit Sub

FillList_Error:

          Dim strES As String
          Dim intEL As Integer

120       intEL = Erl
130       strES = Err.Description
140       LogError "frmPanels", "FillList", intEL, strES, sql

End Sub

Private Sub cmbSampleType_Click()

10        On Error GoTo cmbSampleType_Click_Error

20        FillList
30        FillTree

40        Exit Sub

cmbSampleType_Click_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmPanels", "cmbSampleType_Click", intEL, strES

End Sub



Private Sub cmdAddItem_Click()

          Dim tb As Recordset
          Dim sql As String
          Dim Content As String
          Dim PanelName As String
          Dim PanelType As String
          Dim Key As String
          Dim BarCode As String
          Dim ListOrder As Integer
          Dim NodX As MSComctlLib.Node

10        On Error GoTo cmdAddItem_Click_Error

20        Content = List1.Text
30        PanelName = Tree.SelectedItem.Text
40        Key = Tree.SelectedItem.Key
50        PanelType = ListCodeFor("ST", cmbSampleType)

60        sql = "SELECT * FROM " & TableName & " WHERE " & _
                "PanelType = '" & PanelType & "' " & _
                "AND Hospital = '" & HospName(0) & "' " & _
                "AND Content = '" & Content & "' " & _
                "AND PanelName = '" & PanelName & "'"
70        Set tb = New Recordset
80        RecOpenServer 0, tb, sql
90        If Not tb.EOF Then
100           cmdAddItem.Enabled = False
110           Exit Sub
120       End If

130       sql = "SELECT BarCode, COALESCE(ListOrder, 999) ListOrder " & _
                "FROM " & TableName & " WHERE " & _
                "PanelType = '" & PanelType & "' " & _
                "AND Hospital = '" & HospName(0) & "' " & _
                "AND PanelName = '" & PanelName & "'"
140       Set tb = New Recordset
150       RecOpenServer 0, tb, sql
160       If Not tb.EOF Then
170           BarCode = tb!BarCode & ""
180           ListOrder = tb!ListOrder
190       End If

200       sql = "INSERT INTO " & TableName & " " & _
                "(PanelName, BarCode, ListOrder, Content, PanelType, Hospital) VALUES " & _
                "('" & PanelName & "', " & _
                " '" & BarCode & "', " & _
                " '" & ListOrder & "', " & _
                " '" & Content & "', " & _
                " '" & PanelType & "', " & _
                " '" & HospName(0) & "') "
210       Cnxn(0).Execute sql

220       FillTree

230       For Each NodX In Tree.Nodes
240           If NodX.Parent Is Nothing Then
250               If NodX.Text = PanelName Then
260                   If NodX.Children > 0 Then
270                       NodX.Selected = True
280                       NodX.Child.EnsureVisible
290                       Exit For
300                   End If
310               End If
320           End If
330       Next

340       cmdAddItem.Enabled = False

350       Exit Sub

cmdAddItem_Click_Error:

          Dim strES As String
          Dim intEL As Integer

360       intEL = Erl
370       strES = Err.Description
380       LogError "frmPanels", "cmdAddItem_Click", intEL, strES, sql

End Sub

Private Sub cmdCancel_Click()

10        Unload Me

End Sub

Private Sub cmdRemoveItem_Click()

          Dim PanelType As String
          Dim PanelName As String
          Dim PanelContent As String
          Dim sql As String
          Dim Key As String
          Dim NodX As MSComctlLib.Node

10        On Error GoTo cmdRemoveItem_Click_Error

20        PanelType = ListCodeFor("ST", cmbSampleType)
30        PanelName = Tree.SelectedItem.Parent.Text
40        PanelContent = Tree.SelectedItem.Text
50        Key = Tree.SelectedItem.Parent.Key

60        sql = "DELETE FROM " & TableName & " WHERE " & _
                "Hospital = '" & HospName(0) & "' " & _
                "AND PanelType = '" & PanelType & "' " & _
                "and PanelName = '" & PanelName & "' " & _
                "and Content = '" & PanelContent & "'"
70        Cnxn(0).Execute sql

80        FillTree

90        For Each NodX In Tree.Nodes
100           If NodX.Parent Is Nothing Then
110               If NodX.Text = PanelName Then
120                   If NodX.Children > 0 Then
130                       NodX.Child.EnsureVisible
140                       NodX.Selected = True
150                       Exit For
160                   End If
170               End If
180           End If
190       Next

200       cmdRemoveItem.Enabled = False

210       Exit Sub

cmdRemoveItem_Click_Error:

          Dim strES As String
          Dim intEL As Integer

220       intEL = Erl
230       strES = Err.Description
240       LogError "frmPanels", "cmdRemoveItem_Click", intEL, strES, sql

End Sub

Private Sub cmdRemovePanel_Click()

          Dim PanelType As String
          Dim sql As String

10        On Error GoTo cmdRemovePanel_Click_Error

20        PanelType = ListCodeFor("ST", cmbSampleType)

30        sql = "DELETE from " & TableName & " WHERE " & _
                "PanelType = '" & PanelType & "' " & _
                "and PanelName = '" & AddTicks(Tree.SelectedItem.Text) & "' " & _
                "and Hospital = '" & HospName(0) & "' "

40        Cnxn(0).Execute sql

50        FillTree
60        cmdRemovePanel.Enabled = False

70        Exit Sub

cmdRemovePanel_Click_Error:

          Dim strES As String
          Dim intEL As Integer

80        intEL = Erl
90        strES = Err.Description
100       LogError "frmPanels", "cmdRemovePanel_Click", intEL, strES, sql

End Sub

Private Sub Form_Load()

          Dim tb As New Recordset
          Dim sql As String

10        On Error GoTo Form_Load_Error

20        Me.Caption = "NetAcquire - Define " & mDepartment & " Panels"
30        Select Case mDepartment
          Case "Biochemistry":
40            TableName = "Panels"
50            DefTableName = "BioTestDefinitions"
60        Case "Endocrinology":
70            TableName = "EndPanels"
80            DefTableName = "EndTestDefinitions"
90        Case "Immunology":
100           TableName = "IPanels"
110           DefTableName = "ImmTestDefinitions"
120       End Select

130       sql = "SELECT * from Lists WHERE " & _
                "ListType = 'ST' " & _
                "Order by ListOrder"
140       Set tb = New Recordset
150       RecOpenServer 0, tb, sql

160       cmbSampleType.Clear

170       Do While Not tb.EOF
180           cmbSampleType.AddItem tb!Text & ""
190           tb.MoveNext
200       Loop
210       If cmbSampleType.ListCount > 0 Then
220           cmbSampleType.ListIndex = 0
230       End If

240       FillList
250       FillTree

260       Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

270       intEL = Erl
280       strES = Err.Description
290       LogError "frmPanels", "Form_Load", intEL, strES, sql

End Sub



Private Sub List1_Click()

10        If Tree.SelectedItem.Parent Is Nothing Then
20            cmdAddItem.Enabled = True
30        End If
40        cmdRemoveItem.Enabled = False

End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'
'10    On Error GoTo List1_MouseDown_Error
'
'20    List1.Drag
'
'30    Exit Sub
'
'List1_MouseDown_Error:
'
'      Dim strES As String
'      Dim intEL As Integer
'
'40    intEL = Erl
'50    strES = Err.Description
'60    LogError "frmPanels", "List1_MouseDown", intEL, strES
'
End Sub


Private Sub FillTree()

          Dim NodX As MSComctlLib.Node
          Dim Key As Long
          Dim PanelType As String
          Dim PanelName As String
          Dim tb As New Recordset
          Dim sql As String

10        On Error GoTo FillTree_Error

20        PanelType = ListCodeFor("ST", cmbSampleType)

30        Tree.Visible = False

40        Tree.Nodes.Clear
50        Key = 0

60        sql = "SELECT PanelName, Content from " & TableName & " WHERE " & _
                "PanelType = '" & PanelType & "' " & _
                "AND Hospital = '" & HospName(0) & "' " & _
                "GROUP BY PanelName, Content"
70        Set tb = New Recordset
80        RecOpenServer 0, tb, sql
90        Do While Not tb.EOF
100           If PanelName <> tb!PanelName & "" Then
110               PanelName = tb!PanelName & ""
120               Set NodX = Tree.Nodes.Add(, , "key" & CStr(Key), PanelName)
130               Set NodX = Tree.Nodes.Add("key" & CStr(Key), tvwChild, , tb!Content & "")
140               Key = Key + 1
150           Else
160               Set NodX = Tree.Nodes.Add("key" & CStr(Key - 1), tvwChild, , tb!Content & "")
170           End If
180           tb.MoveNext
190       Loop

200       If Tree.Nodes.Count > 0 Then
210           Tree.Nodes.Item(1).Selected = True
220           Tree.Nodes.Item(1).EnsureVisible
230       End If

240       Tree.Visible = True

250       Exit Sub

FillTree_Error:

          Dim strES As String
          Dim intEL As Integer

260       intEL = Erl
270       strES = Err.Description
280       LogError "frmPanels", "FillTree", intEL, strES, sql
290       Tree.Visible = True

End Sub






Public Property Let Department(ByVal strNewValue As String)

10        mDepartment = strNewValue

End Property



Private Sub Tree_NodeClick(ByVal Node As MSComctlLib.Node)

10        On Error GoTo Tree_NodeClick_Error

20        If Node.Parent Is Nothing Then
30            cmdRemovePanel.Enabled = True
40            cmdRemoveItem.Enabled = False
50        Else
60            cmdRemoveItem.Enabled = True
70            cmdRemovePanel.Enabled = False
80        End If
90        cmdAddItem.Enabled = False

100       Exit Sub

Tree_NodeClick_Error:

          Dim strES As String
          Dim intEL As Integer

110       intEL = Erl
120       strES = Err.Description
130       LogError "frmPanels", "Tree_NodeClick", intEL, strES
End Sub
