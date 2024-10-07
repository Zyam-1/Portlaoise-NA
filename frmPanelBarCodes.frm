VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPanelBarCodes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Panel Bar Codes"
   ClientHeight    =   7080
   ClientLeft      =   1650
   ClientTop       =   780
   ClientWidth     =   4815
   Icon            =   "frmPanelBarCodes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cSampleType 
      Height          =   315
      Left            =   1485
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   135
      Width           =   1845
   End
   Begin VB.CommandButton bcancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   795
      Left            =   1710
      Picture         =   "frmPanelBarCodes.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6165
      Width           =   1605
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   5445
      Left            =   270
      TabIndex        =   1
      Top             =   585
      Width           =   4305
      _ExtentX        =   7594
      _ExtentY        =   9604
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      FormatString    =   "<Panel Name                |<Bar Code                "
   End
End
Attribute VB_Name = "frmPanelBarCodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Tabelname As String
Public frmHeading As String
Private Sub FillG()

          Dim tb As New Recordset
          Dim sql As String
          Dim Found As Boolean
          Dim Y As Long
          Dim SampleType As String

10        On Error GoTo FillG_Error

20        SampleType = ListCodeFor("ST", cSampleType)

30        g.Visible = False
40        g.Rows = 2
50        g.AddItem ""
60        g.RemoveItem 1


70        sql = "SELECT * from " & Tabelname
80        Set tb = New Recordset
90        RecOpenServer 0, tb, sql
100       Do While Not tb.EOF
110           If Trim(tb!PanelType) = SampleType Then
120               Found = False
130               g.Col = 0
140               For Y = 1 To g.Rows - 1
150                   g.row = Y
160                   If g = Trim(tb!PanelName) Then
170                       Found = True
180                       Exit For
190                   End If
200               Next
210               If Not Found Then
220                   g.AddItem Trim(tb!PanelName) & vbTab & Trim(tb!BarCode)
230               End If
240           End If
250           tb.MoveNext
260       Loop

270       If g.Rows > 2 Then
280           g.RemoveItem 1
290       End If

300       g.Visible = True

310       Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer



320       intEL = Erl
330       strES = Err.Description
340       LogError "frmPanelBarCodes", "FillG", intEL, strES, sql


End Sub

Private Sub bcancel_Click()

10        Unload Me

End Sub


Private Sub cSampleType_Click()

10        On Error GoTo cSampleType_Click_Error

20        FillG

30        Exit Sub

cSampleType_Click_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmPanelBarCodes", "cSampleType_Click", intEL, strES


End Sub


Private Sub Form_Load()
          Dim sql As String
          Dim tb As New Recordset

10        On Error GoTo Form_Load_Error
Me.Caption = Me.Caption & " " & frmHeading
20        cSampleType.Clear

30        sql = "SELECT * from lists WHERE listtype = 'ST'"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        Do While Not tb.EOF
70            cSampleType.AddItem Trim(tb!Text)
80            tb.MoveNext
90        Loop
100       If cSampleType.ListCount > 0 Then
110           cSampleType.ListIndex = 0
120       End If

130       FillG

140       Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer



150       intEL = Erl
160       strES = Err.Description
170       LogError "frmPanelBarCodes", "Form_Load", intEL, strES, sql


End Sub


Private Sub g_Click()

          Static SortOrder As Boolean
          Dim PanelName As String
          Dim BarCode As String
          Dim PanelType As String
          Dim sql As String

10        On Error GoTo g_Click_Error

20        If g.MouseRow = 0 Then
30            If SortOrder Then
40                g.Sort = flexSortGenericAscending
50            Else
60                g.Sort = flexSortGenericDescending
70            End If
80            SortOrder = Not SortOrder
90            Exit Sub
100       End If

110       PanelType = ListCodeFor("ST", cSampleType)

120       If g.Col = 1 Then
130           g.Enabled = False
140           g = iBOX("Enter Bar Code", , g)
150           g.Enabled = True
160           PanelName = g.TextMatrix(g.row, 0)
170           BarCode = g.TextMatrix(g.row, 1)
180           sql = "UPDATE " & Tabelname & " set BarCode = '" & BarCode & "' " & _
                    "WHERE PanelName = '" & PanelName & "' " & _
                    "and PanelType = '" & PanelType & "'"
190           Cnxn(0).Execute sql
200       End If

210       Exit Sub

g_Click_Error:

          Dim strES As String
          Dim intEL As Integer



220       intEL = Erl
230       strES = Err.Description
240       LogError "frmPanelBarCodes", "g_Click", intEL, strES, sql


End Sub


