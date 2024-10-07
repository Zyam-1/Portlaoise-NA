VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmPrintingRules 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NetAcquire --- Print Formatting"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   12135
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   1100
      Left            =   10725
      TabIndex        =   13
      Top             =   4080
      Width           =   1200
   End
   Begin VB.CommandButton cmdAddRule 
      Caption         =   "Add Rule"
      Height          =   1100
      Left            =   9315
      TabIndex        =   12
      Top             =   1073
      Width           =   1200
   End
   Begin VB.ComboBox cmbTestName 
      Height          =   315
      ItemData        =   "frmPrintingRules.frx":0000
      Left            =   615
      List            =   "frmPrintingRules.frx":0002
      TabIndex        =   11
      Text            =   "cmbType"
      Top             =   1073
      Width           =   8535
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   1100
      Left            =   10725
      TabIndex        =   10
      Top             =   5280
      Width           =   1200
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   3795
      Left            =   360
      TabIndex        =   9
      Top             =   2580
      Width           =   10155
      _ExtentX        =   17912
      _ExtentY        =   6694
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      FormatString    =   $"frmPrintingRules.frx":0004
   End
   Begin VB.CheckBox chkUnderline 
      Caption         =   "Underline"
      Height          =   255
      Left            =   4545
      TabIndex        =   8
      Top             =   1950
      Width           =   1215
   End
   Begin VB.CheckBox chkItalic 
      Caption         =   "Italic"
      Height          =   255
      Left            =   3270
      TabIndex        =   7
      Top             =   1950
      Width           =   1215
   End
   Begin VB.CheckBox chkBold 
      Caption         =   "Bold"
      Height          =   255
      Left            =   1995
      TabIndex        =   6
      Top             =   1950
      Width           =   1215
   End
   Begin VB.TextBox txtCriteria 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1560
      TabIndex        =   4
      Top             =   1500
      Width           =   4575
   End
   Begin VB.ComboBox cmbType 
      Height          =   315
      ItemData        =   "frmPrintingRules.frx":00D1
      Left            =   3360
      List            =   "frmPrintingRules.frx":00DB
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   420
      Width           =   4335
   End
   Begin VB.Shape Shape1 
      Height          =   1575
      Left            =   120
      Top             =   840
      Width           =   10695
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "then make it "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   360
      TabIndex        =   5
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "contains"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   360
      TabIndex        =   3
      Top             =   1500
      Width           =   1035
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "If"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   360
      TabIndex        =   2
      Top             =   1080
      Width           =   195
   End
   Begin VB.Label lblType 
      Caption         =   "Please select what do you want to format!"
      Height          =   255
      Left            =   3360
      TabIndex        =   1
      Top             =   180
      Width           =   3135
   End
End
Attribute VB_Name = "frmPrintingRules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmbType_Change()

10        On Error GoTo cmbType_Change_Error

20        PopulateTestNames
30        FillGrid

40        Exit Sub

cmbType_Change_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmPrintingRules", "cmbType_Change", intEL, strES

End Sub

Private Sub cmdAddRule_Click()

          Dim i As Integer
          Dim s As String
          Dim AlreadyExist As Boolean

10        On Error GoTo cmdAddRule_Click_Error

20        AlreadyExist = False

30        If cmbTestName = "" Then
40            iMsg "Please select test name"
50            cmbTestName.SetFocus
60            Exit Sub
70        End If

80        If txtCriteria = "" Then
90            iMsg "Please enter criteria"
100           txtCriteria.SetFocus
110           Exit Sub
120       End If

130       If chkBold.Value = 0 And chkItalic.Value = 0 And chkUnderline.Value = 0 Then
140           iMsg "Please select atleast one formatting option"
150           Exit Sub
160       End If

170       For i = 1 To g.Rows - 1
180           If g.TextMatrix(i, 0) = cmbTestName Then
190               AlreadyExist = True
200               Exit For

210           End If
220       Next i

230       If AlreadyExist Then
240           iMsg "Rule already exists"
250           Exit Sub
260       End If

270       s = cmbTestName & vbTab & _
              txtCriteria & vbTab & _
              IIf(chkBold.Value = 1, "True", "False") & vbTab & _
              IIf(chkItalic.Value = 1, "True", "False") & vbTab & _
              IIf(chkUnderline.Value = 1, "True", "False") & vbTab & _
              cmbType.Text

280       g.AddItem s

290       Exit Sub

cmdAddRule_Click_Error:

          Dim strES As String
          Dim intEL As Integer

300       intEL = Erl
310       strES = Err.Description
320       LogError "frmPrintingRules", "cmdAddRule_Click", intEL, strES

End Sub

Private Sub cmdExit_Click()

10        On Error GoTo cmdExit_Click_Error


20        Unload Me

30        Exit Sub

cmdExit_Click_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmPrintingRules", "cmdExit_Click", intEL, strES

End Sub

Private Sub cmdSave_Click()

          Dim sql As String
          Dim i As Integer

10        On Error GoTo cmdSave_Click_Error

20        For i = 1 To g.Rows - 1
30            sql = "IF NOT Exists (SELECT 1 FROM [PrintingRules] WHERE TestName = '" & g.TextMatrix(i, 0) & "') " & _
                    "INSERT INTO [dbo].[PrintingRules] " & _
                    "([TestName] " & _
                    ",[Criteria] " & _
                    ",[Type] " & _
                    ",[Bold] " & _
                    ",[Italic] " & _
                    ",[Underline]) " & _
                    "Values " & _
                    "('" & g.TextMatrix(i, 0) & "', " & _
                    "'" & g.TextMatrix(i, 1) & "', " & _
                    "'" & g.TextMatrix(i, 5) & "', " & _
                    "'" & g.TextMatrix(i, 2) & "', " & _
                    "'" & g.TextMatrix(i, 3) & "', " & _
                    "'" & g.TextMatrix(i, 4) & "') " & _
                    "ELSE " & _
                    "UPDATE [dbo].[PrintingRules] " & _
                    "SET [TestName] = '" & g.TextMatrix(i, 0) & "' " & _
                    ",[Criteria] = '" & g.TextMatrix(i, 1) & "' " & _
                    ",[Type] = '" & g.TextMatrix(i, 5) & "' " & _
                    ",[Bold] = '" & g.TextMatrix(i, 2) & "' " & _
                    ",[Italic] = '" & g.TextMatrix(i, 3) & "' " & _
                    ",[Underline] = '" & g.TextMatrix(i, 4) & "' " & _
                    "WHERE TestName = '" & g.TextMatrix(i, 0) & "' "

40            Cnxn(0).Execute sql
50        Next i

60        Exit Sub

cmdSave_Click_Error:

          Dim strES As String
          Dim intEL As Integer

70        intEL = Erl
80        strES = Err.Description
90        LogError "frmPrintingRules", "cmdSave_Click", intEL, strES, sql

End Sub

Private Sub Form_Load()

10        On Error GoTo Form_Load_Error

20        cmbType.ListIndex = 0
30        g.Rows = 1
40        g.ColWidth(5) = 0

50        PopulateTestNames

60        FixComboWidth cmbTestName
70        FillGrid
80        Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

90        intEL = Erl
100       strES = Err.Description
110       LogError "frmPrintingRules", "Form_Load", intEL, strES

End Sub

Private Sub FillGrid()

          Dim tb As Recordset
          Dim sql As String
          Dim s As String

10        On Error GoTo FillGrid_Error

20        sql = "SELECT * FROM PrintingRules WHERE Type = '" & cmbType & "'"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        While Not tb.EOF
60            s = tb!TestName & "" & vbTab & _
                  tb!Criteria & "" & vbTab & _
                  IIf(tb!Bold = True, "True", "False") & "" & vbTab & _
                  IIf(tb!Italic = True, "True", "False") & "" & vbTab & _
                  IIf(tb!Underline = True, "True", "False") & "" & vbTab & _
                  tb!Type & ""

70            g.AddItem s

80            tb.MoveNext
90        Wend

100       Exit Sub

FillGrid_Error:

          Dim strES As String
          Dim intEL As Integer

110       intEL = Erl
120       strES = Err.Description
130       LogError "frmPrintingRules", "FillGrid", intEL, strES, sql

End Sub

Private Sub PopulateTestNames()

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo PopulateTestNames_Error

20        cmbTestName.Clear

30        Select Case cmbType.Text
          Case "Organisms"
40            sql = "SELECT Name FROM Organisms ORDER BY Name"
50        Case "Generic Results"

60        End Select


70        Set tb = New Recordset
80        RecOpenServer 0, tb, sql
90        While Not tb.EOF
100           cmbTestName.AddItem tb!Name & ""
110           tb.MoveNext
120       Wend

130       Exit Sub

PopulateTestNames_Error:

          Dim strES As String
          Dim intEL As Integer

140       intEL = Erl
150       strES = Err.Description
160       LogError "frmPrintingRules", "PopulateTestNames", intEL, strES, sql

End Sub
