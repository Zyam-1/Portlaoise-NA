VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmExternalTests 
   Caption         =   "NetAcquire - Test List"
   ClientHeight    =   8325
   ClientLeft      =   165
   ClientTop       =   375
   ClientWidth     =   13425
   ForeColor       =   &H80000008&
   Icon            =   "frmExternalTests.frx":0000
   ScaleHeight     =   8325
   ScaleWidth      =   13425
   Begin MSComctlLib.ProgressBar pb 
      Height          =   255
      Left            =   11220
      TabIndex        =   27
      Top             =   6270
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdXL 
      Caption         =   "Export to Excel"
      Height          =   1095
      Left            =   11790
      Picture         =   "frmExternalTests.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3270
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   1095
      Left            =   11790
      Picture         =   "frmExternalTests.frx":5C1E
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5160
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   1095
      Left            =   11790
      Picture         =   "frmExternalTests.frx":6AE8
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6990
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   2985
      Left            =   240
      TabIndex        =   1
      Top             =   150
      Width           =   10815
      Begin VB.TextBox txtComment 
         Height          =   285
         Left            =   5700
         MaxLength       =   50
         TabIndex        =   29
         Top             =   2490
         Width           =   3795
      End
      Begin VB.TextBox txtUnits 
         Height          =   285
         Left            =   3360
         TabIndex        =   26
         Top             =   2490
         Width           =   1545
      End
      Begin VB.TextBox txtMBCode 
         Height          =   315
         Left            =   3390
         TabIndex        =   24
         Top             =   570
         Width           =   1545
      End
      Begin VB.Frame Frame3 
         Caption         =   "Send To Address"
         Height          =   735
         Left            =   330
         TabIndex        =   20
         Top             =   1320
         Width           =   4605
         Begin VB.CommandButton bAddToAddress 
            Caption         =   "Add to Addresses"
            Height          =   465
            Left            =   3570
            TabIndex        =   22
            Top             =   210
            Width           =   915
         End
         Begin VB.ComboBox cmbAddress 
            Height          =   315
            Left            =   120
            TabIndex        =   21
            Text            =   "cmbAddress"
            Top             =   300
            Width           =   3315
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Normal Ranges"
         Height          =   1215
         Left            =   5700
         TabIndex        =   10
         Top             =   210
         Width           =   2685
         Begin VB.TextBox txtFemaleLow 
            Height          =   285
            Left            =   1500
            TabIndex        =   18
            Top             =   750
            Width           =   705
         End
         Begin VB.TextBox txtFemaleHigh 
            Height          =   285
            Left            =   1500
            TabIndex        =   17
            Top             =   450
            Width           =   705
         End
         Begin VB.TextBox txtMaleLow 
            Height          =   285
            Left            =   720
            TabIndex        =   16
            Top             =   750
            Width           =   705
         End
         Begin VB.TextBox txtMaleHigh 
            Height          =   285
            Left            =   720
            TabIndex        =   15
            Top             =   450
            Width           =   705
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Female"
            Height          =   195
            Left            =   1560
            TabIndex        =   14
            Top             =   240
            Width           =   510
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Male"
            Height          =   195
            Left            =   810
            TabIndex        =   13
            Top             =   240
            Width           =   345
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Low"
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   360
            TabIndex        =   12
            Top             =   780
            Width           =   300
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "High"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   330
            TabIndex        =   11
            Top             =   480
            Width           =   330
         End
      End
      Begin VB.TextBox txtAnalyteName 
         Height          =   315
         Left            =   360
         TabIndex        =   9
         Top             =   570
         Width           =   2685
      End
      Begin VB.ComboBox cmbSampleType 
         Height          =   315
         Left            =   360
         TabIndex        =   3
         Text            =   "cmbSampleType"
         Top             =   2490
         Width           =   1545
      End
      Begin VB.CommandButton bAddToList 
         Caption         =   "Add To List"
         Height          =   1125
         Left            =   9600
         Picture         =   "frmExternalTests.frx":79B2
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   930
         Width           =   855
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Comment"
         Height          =   195
         Left            =   5700
         TabIndex        =   28
         Top             =   2280
         Width           =   660
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Units"
         Height          =   195
         Left            =   3360
         TabIndex        =   25
         Top             =   2280
         Width           =   360
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Medibridge Code"
         Height          =   195
         Left            =   3390
         TabIndex        =   23
         Top             =   360
         Width           =   1200
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Sample Type"
         Height          =   195
         Left            =   360
         TabIndex        =   19
         Top             =   2280
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Analyte Name"
         Height          =   195
         Left            =   420
         TabIndex        =   8
         Top             =   360
         Width           =   990
      End
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   4875
      Left            =   240
      TabIndex        =   0
      Top             =   3210
      Width           =   10845
      _ExtentX        =   19129
      _ExtentY        =   8599
      _Version        =   393216
      Cols            =   10
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      AllowUserResizing=   1
      FormatString    =   $"frmExternalTests.frx":9334
   End
   Begin VB.Label lblExcelInfo 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exporting..."
      Height          =   345
      Left            =   11625
      TabIndex        =   7
      Top             =   4470
      Visible         =   0   'False
      Width           =   1305
   End
End
Attribute VB_Name = "frmExternalTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pDepartment As String

Private Sub baddtoaddress_Click()

10    frmAddress.Show 1

20    FillAddress

End Sub

Private Sub baddtolist_Click()

      Dim s As String

10    On Error GoTo baddtolist_Click_Error

20    txtAnalyteName = Trim(txtAnalyteName)

30    If txtAnalyteName = "" Then
40      iMsg "Enter Test Name"
50      Exit Sub
60    End If

70    If cmbAddress = "" Then
80      iMsg "Enter Address"
90      Exit Sub
100   End If

110   s = txtAnalyteName & vbTab & _
          Format$(Val(txtMaleLow)) & vbTab & _
          Format$(Val(txtMaleHigh)) & vbTab & _
          Format$(Val(txtFemaleLow)) & vbTab & _
          Format$(Val(txtFemaleHigh)) & vbTab & _
          txtUnits & vbTab & _
          cmbAddress & vbTab & _
          cmbSampleType.Text & vbTab & _
          txtMBCode & vbTab & _
          txtComment
120   g.AddItem s

130   txtAnalyteName = ""
140   cmbAddress = ""
150   txtUnits = ""
160   txtMaleHigh = ""
170   txtMaleLow = ""
180   txtFemaleHigh = ""
190   txtFemaleLow = ""
200   txtMBCode = ""
210   txtComment = ""

220   cmdSave.Enabled = True

230   Exit Sub

baddtolist_Click_Error:

      Dim strES As String
      Dim intEL As Integer

240   intEL = Erl
250   strES = Err.Description
260   LogError "frmExtTests", "baddtolist_Click", intEL, strES

End Sub

Private Sub cmbSampleType_KeyPress(KeyAscii As Integer)

10    KeyAscii = 0

End Sub


Private Sub cmdCancel_Click()

10    If cmdSave.Enabled Then
20      If iMsg("Cancel without Saving?", vbQuestion + vbYesNo) = vbNo Then
30        Exit Sub
40      End If
50    End If

60    Unload Me

End Sub

Private Sub FillAddress()

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo FillAddress_Error

20    sql = "Select * from extaddress"
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql

50    cmbAddress.Clear
60    Do While Not tb.EOF
70      cmbAddress.AddItem tb!Code & ": " & tb!Addr0
80      tb.MoveNext
90    Loop

100   Exit Sub

FillAddress_Error:

      Dim strES As String
      Dim intEL As Integer

110   intEL = Erl
120   strES = Err.Description
130   LogError "frmExtTests", "FillAddress", intEL, strES, sql

End Sub

Private Sub FillG()

      Dim tb As Recordset
      Dim sql As String
      Dim s As String

10    On Error GoTo FillG_Error

20    g.Rows = 2
30    g.AddItem ""
40    g.RemoveItem 1

50    sql = "Select * from ExternalDefinitions " & _
            "WHERE Department = '" & pDepartment & "' " & _
            "Order by AnalyteName"
60    Set tb = New Recordset
70    RecOpenServer 0, tb, sql
80    Do While Not tb.EOF
90      With tb
100       s = !AnalyteName & vbTab & _
              !MaleLow & vbTab & _
              !MaleHigh & vbTab & _
              !FemaleLow & vbTab & _
              !FemaleHigh & vbTab & _
              !Units & vbTab & _
              !SendTo & vbTab & _
              !SampleType & vbTab & _
              !MBCode & vbTab & _
              !Comment & ""
110       g.AddItem s
120     End With
130     tb.MoveNext
140   Loop

150   If g.Rows > 2 Then
160     g.RemoveItem 1
170   End If

180   Exit Sub

FillG_Error:

      Dim strES As String
      Dim intEL As Integer

190   intEL = Erl
200   strES = Err.Description
210   LogError "frmExtTests", "FillG", intEL, strES, sql

End Sub

Private Sub cmdSave_Click()

      Dim tb As Recordset
      Dim sql As String
      Dim n As Integer

10    On Error GoTo cmdSave_Click_Error

20    pb.Min = 0
30    pb = 0
40    pb.Max = g.Rows - 1
50    pb.Visible = True

60    For n = 1 To g.Rows - 1
70      pb = n
80      If g.TextMatrix(n, 0) <> "" Then
90        sql = "Select * from ExternalDefinitions where " & _
                "AnalyteName  = '" & g.TextMatrix(n, 0) & "' " & _
                "AND Department = '" & pDepartment & "'"
100       Set tb = New Recordset
110       RecOpenServer 0, tb, sql
120       If tb.EOF Then
130         tb.AddNew
140       End If
150       tb!AnalyteName = g.TextMatrix(n, 0)
160       tb!MaleLow = Val(g.TextMatrix(n, 1))
170       tb!MaleHigh = Val(g.TextMatrix(n, 2))
180       tb!FemaleLow = Val(g.TextMatrix(n, 3))
190       tb!FemaleHigh = Val(g.TextMatrix(n, 4))
200       tb!Units = g.TextMatrix(n, 5)
210       tb!SendTo = g.TextMatrix(n, 6)
220       tb!SampleType = g.TextMatrix(n, 7)
230       tb!MBCode = g.TextMatrix(n, 8)
240       tb!Comment = g.TextMatrix(n, 9)
250       tb!PrintPriority = n
          tb!Department = pDepartment
260       tb.Update
270     End If
280   Next

290   cmdSave.Enabled = False

300   pb.Visible = False

310   Exit Sub

cmdSave_Click_Error:

      Dim strES As String
      Dim intEL As Integer

320   intEL = Erl
330   strES = Err.Description
340   LogError "frmExtTests", "cmdsave_Click", intEL, strES, sql

End Sub

Private Sub cmdXL_Click()

10    ExportFlexGrid g, Me

End Sub

Private Sub Form_Load()

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo Form_Load_Error

20    sql = "Select * from Lists where " & _
            "ListType = 'ST' and InUse = 1 " & _
            "order by ListOrder"
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql

50    cmbSampleType.Clear

60    Do While Not tb.EOF
70      cmbSampleType.AddItem tb!Text & ""
80      tb.MoveNext
90    Loop

100   FillAddress
110   FillG

Me.Caption = "NetAcquire - External Test List (" & pDepartment & ")"
120   Exit Sub

Form_Load_Error:

Dim strES As String
Dim intEL As Integer

130   intEL = Erl
140   strES = Err.Description
150   LogError "frmExtTests", "Form_Load", intEL, strES, sql


End Sub

Private Sub g_Click()

      Static SortOrder As Boolean
      Dim s As String
      Dim sql As String

10    On Error GoTo g_Click_Error

20    If g.MouseRow = 0 Then
30      If SortOrder Then
40        g.Sort = flexSortGenericAscending
50      Else
60        g.Sort = flexSortGenericDescending
70      End If
80      SortOrder = Not SortOrder
90      Exit Sub
100   End If

110   If g.Col = 8 Then
120     g.Enabled = False
130     s = iBOX("Medibridge Code?", , g.TextMatrix(g.Row, 8))
140     g.TextMatrix(g.Row, 8) = s
150     g.Enabled = True
160     cmdSave.Enabled = True
170     Exit Sub
180   End If

190   s = "Remove " & g.TextMatrix(g.Row, 0) & " from list?"
200   If iMsg(s, vbQuestion + vbYesNo) = vbNo Then
210     Exit Sub
220   End If

230   txtAnalyteName = g.TextMatrix(g.Row, 0)
240   txtMaleLow = g.TextMatrix(g.Row, 1)
250   txtMaleHigh = g.TextMatrix(g.Row, 2)
260   txtFemaleLow = g.TextMatrix(g.Row, 3)
270   txtFemaleHigh = g.TextMatrix(g.Row, 4)
280   txtUnits = g.TextMatrix(g.Row, 5)
290   cmbAddress = g.TextMatrix(g.Row, 6) & ":"
300   cmbSampleType = g.TextMatrix(g.Row, 7)
310   txtMBCode = g.TextMatrix(g.Row, 8)
320   txtComment = g.TextMatrix(g.Row, 9)

330   sql = "Delete from ExternalDefinitions where " & _
            "AnalyteName = '" & g.TextMatrix(g.Row, 0) & "' " & _
            "AND Department = '" & pDepartment & "'"
340   Cnxn(0).Execute sql

350   g.RemoveItem g.Row

360   Exit Sub

g_Click_Error:

      Dim strES As String
      Dim intEL As Integer

370   intEL = Erl
380   strES = Err.Description
390   LogError "frmExtTests", "g_Click", intEL, strES, sql

End Sub

Private Sub cmbaddress_KeyPress(KeyAscii As Integer)

10    KeyAscii = 0

End Sub


Public Property Let Department(ByVal sNewValue As String)

pDepartment = sNewValue

End Property
