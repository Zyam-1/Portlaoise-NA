VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmSetSources 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Set Sources"
   ClientHeight    =   5445
   ClientLeft      =   540
   ClientTop       =   645
   ClientWidth     =   9270
   Icon            =   "frmSetSources.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   9270
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton bcancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   810
      Left            =   7830
      Picture         =   "frmSetSources.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4500
      Width           =   1155
   End
   Begin VB.CommandButton baddnew 
      Caption         =   "Add New Panel"
      Height          =   975
      Left            =   7830
      Picture         =   "frmSetSources.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1260
      Width           =   1155
   End
   Begin VB.CommandButton bRemovePanel 
      Caption         =   "Remove Panel"
      Enabled         =   0   'False
      Height          =   885
      Left            =   7830
      Picture         =   "frmSetSources.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2295
      Width           =   1155
   End
   Begin VB.CommandButton bRemoveItem 
      Caption         =   "Remove Item"
      Enabled         =   0   'False
      Height          =   750
      Left            =   7830
      Picture         =   "frmSetSources.frx":0C28
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3555
      Width           =   1155
   End
   Begin VB.OptionButton oSource 
      Caption         =   "GPs"
      Height          =   255
      Index           =   2
      Left            =   7950
      TabIndex        =   4
      Top             =   810
      Width           =   1035
   End
   Begin VB.OptionButton oSource 
      Caption         =   "Clinicians"
      Height          =   255
      Index           =   1
      Left            =   7950
      TabIndex        =   3
      Top             =   510
      Width           =   1035
   End
   Begin VB.OptionButton oSource 
      Caption         =   "Wards"
      Height          =   255
      Index           =   0
      Left            =   7950
      TabIndex        =   2
      Top             =   210
      Value           =   -1  'True
      Width           =   1035
   End
   Begin VB.ListBox lstSource 
      Columns         =   2
      DragIcon        =   "frmSetSources.frx":0F32
      Height          =   5130
      Left            =   90
      TabIndex        =   1
      Top             =   90
      Width           =   4815
   End
   Begin MSComctlLib.TreeView Tree 
      Height          =   5145
      Left            =   4950
      TabIndex        =   0
      Top             =   90
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   9075
      _Version        =   393217
      LineStyle       =   1
      Style           =   7
      BorderStyle     =   1
      Appearance      =   1
   End
End
Attribute VB_Name = "frmSetSources"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub FillList()

          Dim sql As String
          Dim tb As New Recordset
          Dim strHospitalCode As String

10        On Error GoTo FillList_Error

20        lstSource.Clear

30        strHospitalCode = ListCodeFor("HO", HospName(0))

40        If oSource(0) Then
50            sql = "SELECT * from wards where hospitalcode = '" & strHospitalCode & "'"
60            Set tb = New Recordset
70            RecOpenClient 0, tb, sql
80            Do While Not tb.EOF
90                lstSource.AddItem Trim(tb!Text)
100               tb.MoveNext
110           Loop
120       ElseIf oSource(1) Then
130           sql = "SELECT * from Clinicians where hospitalcode = '" & strHospitalCode & "'"
140           Set tb = New Recordset
150           RecOpenClient 0, tb, sql
160           Do While Not tb.EOF
170               lstSource.AddItem Trim(tb!Text)
180               tb.MoveNext
190           Loop
200       ElseIf oSource(2) Then
210           sql = "SELECT * from gps where hospitalcode = '" & strHospitalCode & "' ORDER BY LISTORDER"
220           Set tb = New Recordset
230           RecOpenClient 0, tb, sql
240           Do While Not tb.EOF
250               lstSource.AddItem Trim(tb!Text)
260               tb.MoveNext
270           Loop
280       End If


290       Exit Sub

FillList_Error:

          Dim strES As String
          Dim intEL As Integer



300       intEL = Erl
310       strES = Err.Description
320       LogError "frmSetSources", "FillList", intEL, strES, sql


End Sub

Private Sub FillTree()

          Dim NodX As MSComctlLib.Node
          Dim Key As Long
          Dim SourcePanelType As String
          Dim PanelName As String
          Dim Found As Boolean
          Dim tb As New Recordset
          Dim tbp As Recordset
          Dim sql As String



10        On Error GoTo FillTree_Error

20        If oSource(0) Then
30            SourcePanelType = "W"
40        ElseIf oSource(1) Then
50            SourcePanelType = "C"
60        ElseIf oSource(2) Then
70            SourcePanelType = "G"
80        End If

90        Tree.Nodes.Clear
100       sql = "SELECT * from SourcePanels WHERE " & _
                "SourcePanelType = '" & SourcePanelType & "' " & _
                "Order by ListOrder"
110       Set tb = New Recordset
120       RecOpenClient 0, tb, sql

130       Do While Not tb.EOF
140           PanelName = tb!SourcePanelName & ""
150           Found = False
160           For Each NodX In Tree.Nodes
170               If NodX.Text = PanelName Then
180                   Found = True
190                   Exit For
200               End If
210           Next
220           If Not Found Then
230               Set NodX = Tree.Nodes.Add(, , "key" & CStr(Key), PanelName)
240               Key = Key + 1
250               sql = "SELECT * from SourcePanels WHERE " & _
                        "SourcePanelType = '" & SourcePanelType & "' " & _
                        "and SourcePanelName = '" & PanelName & "' " & _
                        "Order by ListOrder"
260               Set tbp = New Recordset
270               RecOpenClient 0, tbp, sql
280               Do While Not tbp.EOF
290                   Set NodX = Tree.Nodes.Add("key" & CStr(Key - 1), tvwChild, , tbp!Content & "")
300                   tbp.MoveNext
310               Loop
320           End If
330           tb.MoveNext
340       Loop





350       Exit Sub

FillTree_Error:

          Dim strES As String
          Dim intEL As Integer



360       intEL = Erl
370       strES = Err.Description
380       LogError "frmSetSources", "FillTree", intEL, strES


End Sub

Private Sub baddnew_Click()

          Dim NodX As MSComctlLib.Node
          Static k As String

10        On Error GoTo baddnew_Click_Error

20        If k = "" Then
30            k = "1"
40        Else
50            k = CStr(Val(k) + 1)
60        End If

70        Set NodX = Tree.Nodes.Add(, , "newkey" & k, "New Panel" & k)

80        Exit Sub

baddnew_Click_Error:

          Dim strES As String
          Dim intEL As Integer



90        intEL = Erl
100       strES = Err.Description
110       LogError "frmSetSources", "baddnew_Click", intEL, strES


End Sub


Private Sub bcancel_Click()

10        Unload Me

End Sub


Private Sub bRemoveItem_Click()

          Dim SourcePanelType As String
          Dim SourcePanelName As String
          Dim Content As String
          Dim sql As String

10        On Error GoTo bRemoveItem_Click_Error

20        If oSource(0) Then
30            SourcePanelType = "W"
40        ElseIf oSource(1) Then
50            SourcePanelType = "C"
60        ElseIf oSource(2) Then
70            SourcePanelType = "G"
80        End If

90        SourcePanelName = Tree.SelectedItem.Parent.Text
100       Content = Tree.SelectedItem.Text

110       sql = "DELETE from SourcePanels WHERE " & _
                "SourcePanelType = '" & SourcePanelType & "' " & _
                "and SourcePanelName = '" & SourcePanelName & "' " & _
                "and Content = '" & Content & "'"
120       Cnxn(0).Execute sql

130       FillTree

140       Exit Sub

bRemoveItem_Click_Error:

          Dim strES As String
          Dim intEL As Integer



150       intEL = Erl
160       strES = Err.Description
170       LogError "frmSetSources", "bRemoveItem_Click", intEL, strES, sql


End Sub

Private Sub bRemovePanel_Click()

          Dim SourcePanelType As String
          Dim sql As String


10        On Error GoTo bRemovePanel_Click_Error

20        If oSource(0) Then
30            SourcePanelType = "W"
40        ElseIf oSource(1) Then
50            SourcePanelType = "C"
60        ElseIf oSource(2) Then
70            SourcePanelType = "G"
80        End If

90        sql = "DELETE from SourcePanels WHERE " & _
                "SourcePanelType = '" & SourcePanelType & "' " & _
                "and SourcePanelName = '" & Tree.SelectedItem.Text & "'"
100       Cnxn(0).Execute sql

110       FillTree




120       Exit Sub

bRemovePanel_Click_Error:

          Dim strES As String
          Dim intEL As Integer



130       intEL = Erl
140       strES = Err.Description
150       LogError "frmSetSources", "bRemovePanel_Click", intEL, strES, sql


End Sub


Private Sub Form_Load()

10        On Error GoTo Form_Load_Error

20        FillList
30        FillTree

40        Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer



50        intEL = Erl
60        strES = Err.Description
70        LogError "frmSetSources", "Form_Load", intEL, strES


End Sub



Private Sub lstSource_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

10        On Error GoTo lstSource_MouseDown_Error

20        lstSource.Drag

30        Exit Sub

lstSource_MouseDown_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmSetSources", "lstSource_MouseDown", intEL, strES


End Sub


Private Sub osource_Click(Index As Integer)

10        On Error GoTo osource_Click_Error

20        FillList
30        FillTree

40        Exit Sub

osource_Click_Error:

          Dim strES As String
          Dim intEL As Integer



50        intEL = Erl
60        strES = Err.Description
70        LogError "frmSetSources", "osource_Click", intEL, strES


End Sub


Private Sub Tree_AfterLabelEdit(Cancel As Integer, NewString As String)

          Dim sql As String
          Dim SourcePanelType As String


10        On Error GoTo Tree_AfterLabelEdit_Error

20        If oSource(0) Then
30            SourcePanelType = "W"
40        ElseIf oSource(1) Then
50            SourcePanelType = "C"
60        ElseIf oSource(2) Then
70            SourcePanelType = "G"
80        End If

90        If Trim$(NewString) = "" Then
100           Cancel = True
110           Exit Sub
120       End If

130       sql = "UPDATE SourcePanels SET " & _
                "SourcePanelName = '" & NewString & "', " & _
                "WHERE " & _
                "SourcePanelName = '" & Tree.SelectedItem.Text & "' " & _
                "and SourcePanelType = '" & SourcePanelType & "'"
140       Cnxn(0).Execute sql

150       Exit Sub



160       Exit Sub

Tree_AfterLabelEdit_Error:

          Dim strES As String
          Dim intEL As Integer

170       intEL = Erl
180       strES = Err.Description
190       LogError "frmSetSources", "Tree_AfterLabelEdit", intEL, strES, sql


End Sub

Private Sub Tree_DragDrop(Source As Control, x As Single, Y As Single)

          Dim NodX As MSComctlLib.Node
          Dim Key
          Dim SourcePanelType As String
          Dim sql As String


10        On Error GoTo Tree_DragDrop_Error

20        If oSource(0) Then
30            SourcePanelType = "W"
40        ElseIf oSource(1) Then
50            SourcePanelType = "C"
60        ElseIf oSource(2) Then
70            SourcePanelType = "G"
80        End If

90        Set NodX = Tree.HitTest(x, Y)
100       If NodX Is Nothing Then Exit Sub

110       If Tree.DropHighlight Is Nothing Then
120           Set Tree.DropHighlight = Nothing
130           Exit Sub
140       Else
150           If NodX = Tree.DropHighlight Then
160               Key = NodX.Key
170               If Key <> "" Then
180                   Set NodX = Tree.Nodes.Add(Key, tvwChild, , Source.Text)
190                   sql = "INSERT into SourcePanels " & _
                            "(SourcePanelName, SourcePanelType, Content, ListOrder) VALUES " & _
                            "('" & Tree.DropHighlight.Text & "', " & _
                            "'" & SourcePanelType & "', " & _
                            "'" & Source.Text & "', " & _
                            "'999')"
200                   Cnxn(0).Execute sql
210                   Set Tree.DropHighlight = Nothing
220                   Tree.Nodes(Key).Child.EnsureVisible
230               End If
240           End If
250       End If

260       Exit Sub

Tree_DragDrop_Error:

          Dim strES As String
          Dim intEL As Integer

270       intEL = Erl
280       strES = Err.Description
290       LogError "frmSetSources", "Tree_DragDrop", intEL, strES, sql


End Sub


Private Sub Tree_DragOver(Source As Control, x As Single, Y As Single, State As Integer)

10        On Error GoTo Tree_DragOver_Error

20        Set Tree.DropHighlight = Tree.HitTest(x, Y)

30        Exit Sub

Tree_DragOver_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmSetSources", "Tree_DragOver", intEL, strES


End Sub


Private Sub Tree_Expand(ByVal Node As MSComctlLib.Node)

10        On Error GoTo Tree_Expand_Error

20        bRemovePanel.Enabled = False
30        bRemoveItem.Enabled = False

40        Exit Sub

Tree_Expand_Error:

          Dim strES As String
          Dim intEL As Integer



50        intEL = Erl
60        strES = Err.Description
70        LogError "frmSetSources", "Tree_Expand", intEL, strES


End Sub


Private Sub Tree_NodeClick(ByVal Node As MSComctlLib.Node)

          Dim s As String

10        On Error GoTo Tree_NodeClick_Error

20        s = Node.Key
30        If s = "" Then
40            bRemovePanel.Enabled = False
50            bRemoveItem.Enabled = True
60        Else
70            bRemovePanel.Enabled = True
80            bRemoveItem.Enabled = False
90        End If


100       Exit Sub

Tree_NodeClick_Error:

          Dim strES As String
          Dim intEL As Integer



110       intEL = Erl
120       strES = Err.Description
130       LogError "frmSetSources", "Tree_NodeClick", intEL, strES


End Sub


