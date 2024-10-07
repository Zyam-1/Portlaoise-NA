VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMicroOrderExternal 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add to External Tests Requested"
   ClientHeight    =   7815
   ClientLeft      =   960
   ClientTop       =   480
   ClientWidth     =   9210
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "frmMicroOrderExternal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7815
   ScaleWidth      =   9210
   Begin Threed.SSPanel SSPanel1 
      Height          =   675
      Index           =   0
      Left            =   4350
      TabIndex        =   3
      Top             =   120
      Width           =   1965
      _Version        =   65536
      _ExtentX        =   3466
      _ExtentY        =   1191
      _StockProps     =   15
      Caption         =   "Test required    "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   2
      BorderWidth     =   2
      BevelOuter      =   0
      BevelInner      =   1
      Alignment       =   4
      Begin VB.Image Image1 
         Height          =   480
         Left            =   150
         Picture         =   "frmMicroOrderExternal.frx":030A
         Top             =   90
         Width           =   480
      End
   End
   Begin VB.ListBox lstPanels 
      Height          =   3375
      Left            =   4350
      TabIndex        =   2
      Top             =   2070
      Width           =   1965
   End
   Begin VB.ListBox lstOrder 
      Height          =   5325
      Left            =   6480
      TabIndex        =   1
      Top             =   120
      Width           =   2445
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   5940
      Picture         =   "frmMicroOrderExternal.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6090
      Width           =   1380
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   855
      Index           =   1
      Left            =   4350
      TabIndex        =   4
      Top             =   990
      Width           =   1965
      _Version        =   65536
      _ExtentX        =   3466
      _ExtentY        =   1508
      _StockProps     =   15
      Caption         =   "Panel required    "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   2
      BorderWidth     =   2
      BevelOuter      =   0
      BevelInner      =   1
      Alignment       =   6
      Begin VB.Image Image2 
         Height          =   480
         Left            =   630
         Picture         =   "frmMicroOrderExternal.frx":20CE
         Top             =   270
         Width           =   480
      End
   End
   Begin MSComctlLib.TreeView treTests 
      Height          =   7665
      Left            =   135
      TabIndex        =   5
      Top             =   90
      Width           =   4065
      _ExtentX        =   7170
      _ExtentY        =   13520
      _Version        =   393217
      LineStyle       =   1
      Style           =   6
      Appearance      =   1
   End
End
Attribute VB_Name = "frmMicroOrderExternal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

          Dim n As Long
          Dim s As String
          Dim TestName As String

10        On Error GoTo cmdCancel_Click_Error

20        With frmEditMicroExternals.grdExt

30            For n = 0 To lstOrder.ListCount - 1
40                TestName = lstOrder.List(n)
50                s = TestName & vbTab & _
                      eName2SendTo(TestName, "Micro") & vbTab & _
                      Format$(Now, "dd/MM/yyyy HH:mm")
60                .AddItem s
70            Next

80            If .Rows > 2 And .TextMatrix(1, 0) = "" Then
90                .RemoveItem 1
100           End If

110       End With

120       Unload Me

130       Exit Sub

cmdCancel_Click_Error:

          Dim strES As String
          Dim intEL As Integer

140       intEL = Erl
150       strES = Err.Description
160       LogError "frmMicroOrderExternal", "cmdcancel_Click", intEL, strES

End Sub

Sub FillPanels()
Attribute FillPanels.VB_Description = "Load Panels"

          Dim sn As New Recordset
          Dim sql As String

10        On Error GoTo FillPanels_Error

20        sql = "SELECT DISTINCT PanelName FROM ExtPanels WHERE " & _
                "Department = 'Micro'" & _
                "ORDER BY PanelName"
30        Set sn = New Recordset
40        RecOpenServer 0, sn, sql
50        lstPanels.Clear

60        Do While Not sn.EOF
70            lstPanels.AddItem sn!PanelName
80            sn.MoveNext
90        Loop

100       Exit Sub

FillPanels_Error:

          Dim strES As String
          Dim intEL As Integer

110       intEL = Erl
120       strES = Err.Description
130       LogError "frmMicroOrderExternal", "FillPanels", intEL, strES, sql

End Sub

Sub FillTV()
Attribute FillTV.VB_Description = "Fill Ndal List"

          Dim NodX As MSComctlLib.Node
          Dim Num As Long
          Dim Relative As String
          Dim ThisNode As String
          Dim tb As New Recordset
          Dim sql As String
          Dim Key As Long
          Dim T As String

10        On Error GoTo FillTV_Error

20        treTests.Nodes.Clear
30        T = "Tests"
40        Set NodX = treTests.Nodes.Add(, , "key" & CStr(Key), T)

50        For Num = Asc("A") To Asc("Z")
60            Set NodX = treTests.Nodes.Add(, , chr$(Num), chr$(Num))
70        Next
80        For Num = Asc("0") To Asc("9")
90            Set NodX = treTests.Nodes.Add(, , "#" & chr$(Num), chr$(Num))
100       Next

110       sql = "SELECT AnalyteName FROM ExternalDefinitions WHERE " & _
                "Department = 'Micro' " & _
                "ORDER BY AnalyteName"
120       Set tb = New Recordset
130       RecOpenServer 0, tb, sql
140       Do While Not tb.EOF
150           Relative = UCase(Left(Trim(tb!AnalyteName), 1))
160           If IsNumeric(Relative) Then Relative = "#" & Relative
170           ThisNode = tb!AnalyteName
180           Set NodX = treTests.Nodes.Add(Relative, tvwChild, , ThisNode)
190           tb.MoveNext
200       Loop

210       Exit Sub

FillTV_Error:

          Dim strES As String
          Dim intEL As Integer

220       intEL = Erl
230       strES = Err.Description
240       LogError "frmMicroOrderExternal", "FillTV", intEL, strES, sql

End Sub

Private Sub Form_Load()

10        On Error GoTo Form_Load_Error

20        FillTV
30        FillPanels

40        Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmMicroOrderExternal", "Form_Load", intEL, strES

End Sub

Private Sub lstOrder_Click()

          Dim Str As String

10        On Error GoTo lstOrder_Click_Error

20        Str = "Remove " & lstOrder & " from tests requested?"
30        If iMsg(Str, vbYesNo + vbQuestion) = vbYes Then
40            lstOrder.RemoveItem lstOrder.ListIndex
50        End If

60        Exit Sub

lstOrder_Click_Error:

          Dim strES As String
          Dim intEL As Integer

70        intEL = Erl
80        strES = Err.Description
90        LogError "frmMicroOrderExternal", "lstOrder_Click", intEL, strES

End Sub

Private Sub lstPanels_Click()

          Dim sn As New Recordset
          Dim sql As String

10        On Error GoTo lstPanels_Click_Error

20        sql = "SELECT Content FROM ExtPanels WHERE " & _
                "PanelName = '" & lstPanels & "' " & _
                "AND Department = 'Micro'"
30        Set sn = New Recordset
40        RecOpenServer 0, sn, sql
50        Do While Not sn.EOF
60            lstOrder.AddItem eNumber2Name(sn!Content, "Micro")
70            sn.MoveNext
80        Loop

90        Exit Sub

lstPanels_Click_Error:

          Dim strES As String
          Dim intEL As Integer

100       intEL = Erl
110       strES = Err.Description
120       LogError "frmMicroOrderExternal", "lstPanels_Click", intEL, strES, sql

End Sub


Private Sub treTests_NodeCheck(ByVal Node As MSComctlLib.Node)

          Dim sql As String
          Dim tb As New Recordset

10        On Error GoTo treTests_NodeCheck_Error

20        sql = "SELECT AnalyteName FROM ExternalDefinitions WHERE " & _
                "Department = 'Micro' " & _
                "ORDER BY AnalyteName"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        Do While Not tb.EOF
60            If Trim(Node.Text) = Trim(tb!AnalyteName) Then
70                lstOrder.AddItem Node.Text
80            End If
90            tb.MoveNext
100       Loop

110       Exit Sub

treTests_NodeCheck_Error:

          Dim strES As String
          Dim intEL As Integer

120       intEL = Erl
130       strES = Err.Description
140       LogError "frmMicroOrderExternal", "treTests_NodeCheck", intEL, strES, sql

End Sub

Private Sub treTests_NodeClick(ByVal Node As MSComctlLib.Node)

          Dim sql As String
          Dim tb As New Recordset


10        On Error GoTo treTests_NodeClick_Error

20        sql = "SELECT AnalyteName FROM ExternalDefinitions WHERE " & _
                "Department = 'Micro' " & _
                "ORDER BY AnalyteName"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        Do While Not tb.EOF
60            If Trim(Node.Text) = Trim(tb!AnalyteName) Then
70                lstOrder.AddItem Node.Text
80            End If
90            tb.MoveNext
100       Loop

110       Exit Sub

treTests_NodeClick_Error:

          Dim strES As String
          Dim intEL As Integer

120       intEL = Erl
130       strES = Err.Description
140       LogError "frmMicroOrderExternal", "treTests_NodeClick", intEL, strES

End Sub

