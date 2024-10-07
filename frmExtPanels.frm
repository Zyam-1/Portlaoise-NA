VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmExtPanels 
   Caption         =   "NetAcquire - External Panels"
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9900
   Icon            =   "frmExtPanels.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   9900
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbDepartment 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      ItemData        =   "frmExtPanels.frx":030A
      Left            =   3660
      List            =   "frmExtPanels.frx":030C
      TabIndex        =   9
      Text            =   "Combo1"
      Top             =   142
      Width           =   3555
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add to Panel"
      Height          =   945
      Left            =   3750
      Picture         =   "frmExtPanels.frx":030E
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Add to Panel"
      Top             =   1440
      Width           =   1005
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove from Panel"
      Height          =   945
      Left            =   3750
      Picture         =   "frmExtPanels.frx":1C90
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Remove from Panel"
      Top             =   2550
      Width           =   1005
   End
   Begin MSComctlLib.TreeView tvTests 
      Height          =   7695
      Left            =   180
      TabIndex        =   4
      Top             =   660
      Width           =   3435
      _ExtentX        =   6059
      _ExtentY        =   13573
      _Version        =   393217
      HideSelection   =   0   'False
      LineStyle       =   1
      Style           =   6
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   810
      Left            =   8625
      Picture         =   "frmExtPanels.frx":3612
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7560
      Width           =   1155
   End
   Begin VB.CommandButton cmdAddNew 
      Caption         =   "&Add New Panel"
      Height          =   975
      Left            =   8625
      Picture         =   "frmExtPanels.frx":4F94
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   690
      Width           =   1155
   End
   Begin VB.CommandButton cmdRemovePanel 
      Caption         =   "&Remove Panel"
      Enabled         =   0   'False
      Height          =   885
      Left            =   8625
      Picture         =   "frmExtPanels.frx":6916
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1725
      Width           =   1155
   End
   Begin MSComctlLib.TreeView tvPanels 
      Height          =   7695
      Left            =   4860
      TabIndex        =   3
      Top             =   660
      Width           =   3645
      _ExtentX        =   6429
      _ExtentY        =   13573
      _Version        =   393217
      HideSelection   =   0   'False
      LineStyle       =   1
      Style           =   7
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Department"
      Height          =   195
      Left            =   480
      TabIndex        =   8
      Top             =   255
      Width           =   825
   End
   Begin VB.Label lblDepartment 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Micro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1380
      TabIndex        =   7
      Top             =   142
      Width           =   2250
   End
End
Attribute VB_Name = "frmExtPanels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private pDepartment As String
Sub FillTvTests()

          Dim NodX As MSComctlLib.Node
          Dim n As Long
          Dim Relative As String
          Dim ThisNode As String
          Dim sql As String
          Dim tb As New Recordset

10        On Error GoTo FillTvTests_Error

20        tvTests.Visible = False
30        tvTests.Nodes.Clear

40        For n = Asc("A") To Asc("Z")
50            Set NodX = tvTests.Nodes.Add(, , chr$(n), chr$(n))
60        Next
70        For n = Asc("0") To Asc("9")
80            Set NodX = tvTests.Nodes.Add(, , "#" & chr$(n), chr$(n))
90        Next

100       If UCase(pDepartment) = "GENERAL" Then
110           sql = "SELECT AnalyteName FROM ExternalDefinitions " & _
                    "ORDER BY AnalyteName"
120       Else

130           sql = "SELECT AnalyteName FROM ExternalDefinitions " & _
                    "WHERE Department = '" & pDepartment & "' " & _
                    "ORDER BY AnalyteName"
140       End If


          'sql = "SELECT AnalyteName FROM ExternalDefinitions WHERE " & _
           '        "Department = '" & pDepartment & "' " & _
           '        "ORDER BY AnalyteName"
150       Set tb = New Recordset
160       RecOpenServer 0, tb, sql
170       Do While Not tb.EOF
180           Relative = UCase(Left(tb!AnalyteName, 1))
190           If IsNumeric(Relative) Then Relative = "#" & Relative
200           ThisNode = Trim(tb!AnalyteName)
210           Set NodX = tvTests.Nodes.Add(Relative, tvwChild, , ThisNode)
220           tb.MoveNext
230       Loop

240       tvTests.Visible = True

250       Exit Sub

FillTvTests_Error:

260       tvTests.Visible = True

          Dim strES As String
          Dim intEL As Integer

270       intEL = Erl
280       strES = Err.Description
290       LogError "frmExtPanels", "FillTvTests", intEL, strES, sql

End Sub

Private Sub cmbDepartment_Click()

10        On Error GoTo cmbDepartment_Click_Error

20        pDepartment = cmbDepartment
30        lblDepartment = cmbDepartment
40        FillTvTests
50        FilltvPanels

60        Exit Sub

cmbDepartment_Click_Error:

          Dim strES As String
          Dim intEL As Integer

70        intEL = Erl
80        strES = Err.Description
90        LogError "frmExtPanels", "cmbDepartment_Click", intEL, strES

End Sub

Private Sub cmdAddNew_Click()

          Dim NodX As MSComctlLib.Node
          Static k As String
          Dim NewPanelName As String

10        On Error GoTo cmdAddNew_Click_Error

20        NewPanelName = Trim$(iBOX("New Panel Name?"))
30        If NewPanelName = "" Then
40            iMsg "Entry Error. Cannot add.", vbExclamation
50            Exit Sub
60        End If

70        If k = "" Then
80            k = "1"
90        Else
100           k = CStr(Val(k) + 1)
110       End If

120       Set NodX = tvPanels.Nodes.Add(, , "newkey" & k, NewPanelName)

130       Exit Sub

cmdAddNew_Click_Error:

          Dim strES As String
          Dim intEL As Integer

140       intEL = Erl
150       strES = Err.Description
160       LogError "frmExtPanels", "cmdAddNew_Click", intEL, strES

End Sub

Private Sub cmdRemovePanel_Click()

          Dim sql As String

10        On Error GoTo cmdRemovePanel_Click_Error

20        sql = "DELETE FROM ExtPanels " & _
                "WHERE PanelName = '" & tvPanels.SelectedItem.Text & "' " & _
                "AND Department = '" & pDepartment & "'"
30        Cnxn(0).Execute sql

40        FilltvPanels

50        Exit Sub

cmdRemovePanel_Click_Error:

          Dim strES As String
          Dim intEL As Integer

60        intEL = Erl
70        strES = Err.Description
80        LogError "frmExtPanels", "cmdRemovePanel_Click", intEL, strES, sql

End Sub

Private Sub cmdadd_Click()

          Dim NodeParent As MSComctlLib.Node
          Dim NodeNew As MSComctlLib.Node
          Dim Key As String
          Dim sql As String
          Dim tb As Recordset
          Dim MBCode As String

10        If tvPanels.SelectedItem Is Nothing Then Exit Sub
20        If tvTests.SelectedItem Is Nothing Then Exit Sub
30        If tvTests.SelectedItem.Parent Is Nothing Then Exit Sub

40        If tvPanels.SelectedItem.Parent Is Nothing Then

50            Set NodeParent = tvPanels.SelectedItem
60            Key = NodeParent.Key
70            If Key <> "" Then
80                sql = "Select MBCode From ExternalDefinitions Where AnalyteName = '" & AddTicks(tvTests.SelectedItem.Text) & "'"
90                Set tb = New Recordset
100               RecOpenClient 0, tb, sql
110               If Not tb.EOF Then
120                   MBCode = tb!MBCode & ""
130               End If
140               Set NodeNew = tvPanels.Nodes.Add(Key, tvwChild, , tvTests.SelectedItem.Text)
150               sql = "SELECT * FROM ExtPanels WHERE 0 = 1"
160               Set tb = New Recordset
170               RecOpenServer 0, tb, sql
180               tb.AddNew
190               tb!PanelName = NodeParent.Text
200               tb!Content = MBCode
210               tb!TestName = NodeNew.Text
220               tb!Department = pDepartment
230               tb.Update
240           End If

250       End If

End Sub

Private Sub cmdCancel_Click()

10        Unload Me

End Sub

Private Sub cmdRemove_Click()

          Dim sql As String
          Dim PanelName As String

10        On Error GoTo cmdRemove_Click_Error

20        PanelName = tvPanels.SelectedItem.Parent.Text

30        sql = "DELETE FROM ExtPanels WHERE " & _
                "PanelName = '" & PanelName & "' " & _
                "AND TestName = '" & tvPanels.SelectedItem.Text & "' " & _
                "AND Department = '" & pDepartment & "'"
40        Cnxn(0).Execute sql

50        FilltvPanels

60        Exit Sub

cmdRemove_Click_Error:

          Dim strES As String
          Dim intEL As Integer

70        intEL = Erl
80        strES = Err.Description
90        LogError "frmExtPanels", "cmdRemove_Click", intEL, strES, sql

End Sub

Private Sub Form_Load()

10        PopulateDepartments

20        lblDepartment = pDepartment
30        cmbDepartment = pDepartment

40        FillTvTests
50        FilltvPanels

End Sub

Private Sub tvPanels_Click()
'
'cmdRemovePanel.Enabled = True
'
'
End Sub

Private Sub tvPanels_Collapse(ByVal Node As MSComctlLib.Node)

10        cmdRemovePanel.Enabled = False

20        If Node.Parent Is Nothing Then
30            Node.Selected = True
40        End If

End Sub


Private Sub FilltvPanels()

          Dim sql As String
          Dim tb As Recordset
          Dim NodX As MSComctlLib.Node
          Dim Key As Long
          Dim PanelName As String

10        On Error GoTo FilltvPanels_Error

20        tvPanels.Nodes.Clear

30        If UCase(pDepartment) = "GENERAL" Then
40            sql = "SELECT PanelName, TestName FROM ExtPanels " & _
                    "GROUP BY PanelName, TestName"
50        Else

60            sql = "SELECT PanelName, TestName FROM ExtPanels " & _
                    "WHERE Department = '" & pDepartment & "' " & _
                    "GROUP BY PanelName, TestName"
70        End If


80        Set tb = New Recordset
90        RecOpenServer 0, tb, sql
100       PanelName = ""
110       Key = 0
120       Do While Not tb.EOF
130           If tb!PanelName <> PanelName Then
140               PanelName = tb!PanelName
150               Key = Key + 1
160               Set NodX = tvPanels.Nodes.Add(, , "key" & CStr(Key), PanelName)
170           End If
180           Set NodX = tvPanels.Nodes.Add("key" & CStr(Key), tvwChild, , IIf(Right(tb!TestName, 2) = vbCrLf, Left(tb!TestName, Len(tb!TestName) - 2), tb!TestName) & "")
190           tb.MoveNext
200       Loop

210       Exit Sub

FilltvPanels_Error:

          Dim strES As String
          Dim intEL As Integer

220       intEL = Erl
230       strES = Err.Description
240       LogError "frmExtPanels", "FilltvPanels", intEL, strES, sql

End Sub



Private Sub tvPanels_Expand(ByVal Node As MSComctlLib.Node)

10        cmdRemovePanel.Enabled = False

20        If Node.Parent Is Nothing Then
30            Node.Selected = True
40        End If

End Sub





Private Sub tvPanels_NodeClick(ByVal Node As MSComctlLib.Node)

          Dim s As String

10        s = Node.Key
20        If s = "" Then
30            cmdRemovePanel.Enabled = False
40        Else
50            cmdRemovePanel.Enabled = True
60        End If

End Sub



Public Property Let Department(ByVal sNewValue As String)

10        pDepartment = sNewValue

End Property


Private Sub PopulateDepartments()

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo PopulateDepartments_Error


20        With cmbDepartment
30            .Clear
40            .AddItem "General"
50            sql = "SELECT DISTINCT Department FROM ExternalDefinitions " & _
                    "ORDER BY Department"
60            Set tb = New Recordset
70            RecOpenClient 0, tb, sql
80            If Not tb.EOF Then
90                While Not tb.EOF
100                   .AddItem tb!Department & ""
110                   tb.MoveNext
120               Wend
130           End If

140       End With

150       Exit Sub

PopulateDepartments_Error:

          Dim strES As String
          Dim intEL As Integer

160       intEL = Erl
170       strES = Err.Description
180       LogError "frmAddToTests", "PopulateDepartments", intEL, strES, sql

End Sub

