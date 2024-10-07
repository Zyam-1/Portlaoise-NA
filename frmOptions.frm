VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NetAcquire --- System Options"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   13905
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox fraProgress 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   5400
      ScaleHeight     =   885
      ScaleWidth      =   3825
      TabIndex        =   20
      Top             =   2880
      Visible         =   0   'False
      Width           =   3855
      Begin MSComctlLib.ProgressBar pbProgress 
         Height          =   255
         Left            =   0
         TabIndex        =   21
         Top             =   240
         Width           =   3840
         _ExtentX        =   6773
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         Min             =   1
      End
      Begin VB.Label lblProgress 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Fetching results ..."
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         TabIndex        =   22
         Top             =   0
         Width           =   3840
      End
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Height          =   495
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   300
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      Caption         =   "Add New Option"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   180
      TabIndex        =   5
      Top             =   7020
      Width           =   11355
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H0000FFFF&
         Caption         =   "&Save"
         Height          =   900
         Left            =   10020
         Picture         =   "frmOptions.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   420
         Width           =   1100
      End
      Begin VB.TextBox txtOptionDetail 
         Height          =   615
         Left            =   3840
         MaxLength       =   1000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Top             =   705
         Width           =   5895
      End
      Begin VB.ComboBox cmbOptionType 
         Height          =   315
         Left            =   1200
         TabIndex        =   11
         Top             =   705
         Width           =   1935
      End
      Begin VB.ComboBox cmbOptionCategory 
         Height          =   315
         Left            =   1200
         TabIndex        =   10
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox txtOptionUsername 
         BackColor       =   &H8000000B&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   9
         Top             =   1440
         Width           =   1275
      End
      Begin VB.TextBox txtOptionValue 
         Height          =   285
         Left            =   3840
         TabIndex        =   8
         Top             =   360
         Width           =   5895
      End
      Begin VB.TextBox txtOptionName 
         Height          =   285
         Left            =   1200
         TabIndex        =   7
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Details"
         Height          =   195
         Left            =   3225
         TabIndex        =   18
         Top             =   900
         Width           =   480
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Username"
         Height          =   195
         Left            =   390
         TabIndex        =   16
         Top             =   1500
         Width           =   720
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Value"
         Height          =   195
         Left            =   3300
         TabIndex        =   15
         Top             =   405
         Width           =   405
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Option Name"
         Height          =   195
         Left            =   180
         TabIndex        =   14
         Top             =   405
         Width           =   930
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Type"
         Height          =   195
         Left            =   750
         TabIndex        =   13
         Top             =   765
         Width           =   360
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Category"
         Height          =   195
         Left            =   480
         TabIndex        =   12
         Top             =   1140
         Width           =   630
      End
   End
   Begin VB.TextBox txtDescription 
      Height          =   615
      Left            =   1140
      Locked          =   -1  'True
      MaxLength       =   1000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   5760
      Width           =   12585
   End
   Begin VB.TextBox txtText 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   12840
      TabIndex        =   2
      Top             =   360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ComboBox cmbCombo 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   12720
      TabIndex        =   1
      Top             =   60
      Visible         =   0   'False
      Width           =   675
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   4725
      Left            =   180
      TabIndex        =   0
      Top             =   960
      Width           =   13545
      _ExtentX        =   23892
      _ExtentY        =   8334
      _Version        =   393216
      Cols            =   8
      FixedCols       =   3
      RowHeightMin    =   315
      BackColorFixed  =   -2147483635
      ForeColorFixed  =   -2147483634
      HighLight       =   0
      AllowUserResizing=   1
      FormatString    =   $"frmOptions.frx":066A
   End
   Begin VB.Frame fraFilter 
      Height          =   795
      Left            =   180
      TabIndex        =   23
      Top             =   60
      Width           =   8115
      Begin VB.ComboBox cmbCategories 
         Height          =   315
         Left            =   4260
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   300
         Width           =   3675
      End
      Begin VB.ComboBox cmbTypes 
         Height          =   315
         Left            =   780
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   300
         Width           =   2295
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Type"
         Height          =   195
         Left            =   240
         TabIndex        =   27
         Top             =   360
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Category"
         Height          =   195
         Left            =   3540
         TabIndex        =   26
         Top             =   360
         Width           =   630
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Description"
      Height          =   195
      Left            =   180
      TabIndex        =   4
      Top             =   5940
      Width           =   795
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mEditDescription As Boolean
Private mEditType As Boolean
Private mEditCategory As Boolean
Private mSelectedCategory As String
Private mSelectedType As String

Public Property Let EditDescription(ByVal NewValue As Boolean)
10        mEditDescription = NewValue
End Property

Public Property Get EditDescription() As Boolean
10        EditDescription = mEditDescription
End Property



Public Property Let EditType(ByVal NewValue As Boolean)
10        mEditType = NewValue
End Property

Public Property Get EditType() As Boolean
10        EditType = mEditType
End Property



Public Property Let EditCategory(ByVal NewValue As Boolean)
10        mEditCategory = NewValue
End Property

Public Property Get EditCategory() As Boolean
10        EditCategory = mEditCategory
End Property

Public Property Let SelectedCategory(ByVal NewValue As String)
10        mSelectedCategory = NewValue
End Property

Public Property Get SelectedCategory() As String
10        SelectedCategory = mSelectedCategory
End Property

Public Property Let SelectedType(ByVal NewValue As String)
10        mSelectedType = NewValue
End Property

Public Property Get SelectedType() As String
10        SelectedType = mSelectedType
End Property


Private Sub LoadAllOptions(ByVal OptionType As String, ByVal OptionCategory As String)

          Dim tb As Recordset
          Dim sql As String
          Dim n As Integer


10        On Error GoTo LoadAllOptions_Error

20        g.Clear
30        g.Rows = 2
40        g.FormatString = "     ||Option                                           |Value                                                        |Description                                                                                  |Type                 |Category              |Created By            "
50        g.ColWidth(1) = 0
60        g.Refresh



70        sql = "SELECT * FROM Options"
80        If cmbTypes <> "All Types" Or cmbCategories <> "All Categories" Then
90            sql = sql & " WHERE "
100       End If
110       If cmbTypes <> "All Types" Then
120           sql = sql & "OptType = '" & OptionType & "'"
130       End If
140       If cmbCategories <> "All Categories" Then
150           If cmbTypes <> "All Types" Then
160               sql = sql & " AND "
170           End If
180           sql = sql & "OptCategory = '" & OptionCategory & "' ORDER BY ListOrder"
190       End If

200       Set tb = New Recordset
210       RecOpenClient 0, tb, sql
220       If Not tb.EOF Then
230           g.Visible = False
240           fraProgress.Visible = True
250           pbProgress.Value = 1
260           pbProgress.Max = tb.RecordCount + 1
270           lblProgress = "Fetching results ... (0%)"
280           lblProgress.Refresh
290           While Not tb.EOF
300               g.AddItem vbTab & _
                            tb!Description & "" & vbTab & _
                            tb!OptionName & "" & vbTab & _
                            tb!Contents & "" & vbTab & _
                            tb!Details & "" & vbTab & _
                            tb!optType & "" & vbTab & _
                            tb!OptCategory & "" & vbTab & _
                            tb!Username & "" & vbTab

310               tb.MoveNext
320               pbProgress.Value = pbProgress.Value + 1
330               lblProgress = "Fetching results ... (" & Int(pbProgress.Value * 100 / pbProgress.Max) & " %)"
340               lblProgress.Refresh
350           Wend
360           fraProgress.Visible = False
370           g.Visible = True
380           If g.Rows > 2 Then g.RemoveItem 1
390       End If



400       Exit Sub

LoadAllOptions_Error:

          Dim strES As String
          Dim intEL As Integer

410       intEL = Erl
420       strES = Err.Description
430       LogError "frmOptions", "LoadAllOptions", intEL, strES, sql
440       fraProgress.Visible = False
450       g.Visible = True

End Sub

Private Sub PopulateTypes(cmb As ComboBox, Optional SelectList As Boolean = False)

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo PopulateTypes_Error

20        cmb.Clear

30        If SelectList Then
40            cmb.AddItem "All Types"
50        End If

60        sql = "SELECT DISTINCT optType From Options WHERE COALESCE(OptType, '') <> ''"
70        Set tb = New Recordset
80        RecOpenClient 0, tb, sql
90        If Not tb.EOF Then
100           While Not tb.EOF
110               cmb.AddItem tb!optType & ""
120               tb.MoveNext
130           Wend
140       End If

150       Exit Sub

PopulateTypes_Error:

          Dim strES As String
          Dim intEL As Integer

160       intEL = Erl
170       strES = Err.Description
180       LogError "frmOptions", "PopulateTypes", intEL, strES, sql

End Sub

Private Sub PopulateCategories(cmb As ComboBox, Optional SelectList As Boolean = False)

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo PopulateCategories_Error

20        cmb.Clear

30        If SelectList Then
40            cmb.AddItem "All Categories"
50        End If
60        sql = "SELECT DISTINCT optCategory From Options WHERE COALESCE(OptCategory, '') <> ''"
70        Set tb = New Recordset
80        RecOpenClient 0, tb, sql
90        If Not tb.EOF Then
100           While Not tb.EOF
110               cmb.AddItem tb!OptCategory & ""
120               tb.MoveNext
130           Wend
140       End If


150       Exit Sub

PopulateCategories_Error:

          Dim strES As String
          Dim intEL As Integer

160       intEL = Erl
170       strES = Err.Description
180       LogError "frmOptions", "PopulateCategories", intEL, strES, sql

End Sub



Private Sub cmbCategories_Click()

10        If cmbTypes = "" Then
20            iMsg "Please select option type first", vbInformation
30            cmbTypes.SetFocus
40            Exit Sub
50        End If
60        If cmbCategories = "" Then Exit Sub
70        Me.SelectedType = cmbTypes
80        Me.SelectedCategory = cmbCategories
90        LoadAllOptions cmbTypes, cmbCategories


End Sub

Private Sub LoadControls()
10        On Error GoTo LoadControls_Error

20        txtText.Visible = False
30        txtText = ""
40        cmbCombo.Text = ""
50        cmbCombo.Visible = False

60        Select Case g.Col
              '    Case 2:         'Display Name
              '        txtText.Move g.Left + g.CellLeft, _
                       '                     g.Top + g.CellTop, _
                       '                     g.CellWidth, g.CellHeight
              '        txtText.Text = g.TextMatrix(g.Row, g.Col)
              '        txtText.MaxLength = 300
              '        txtText.Visible = True
              '        txtText.SelStart = 0
              '        txtText.SelLength = Len(txtText)
              '        txtText.SetFocus
          Case 3:    'Value
70            txtText.Move g.Left + g.CellLeft, _
                           g.Top + g.CellTop, _
                           g.CellWidth, g.CellHeight
80            txtText.Text = g.TextMatrix(g.Row, g.Col)
90            txtText.MaxLength = 50
100           txtText.Visible = True
110           txtText.SelStart = 0
120           txtText.SelLength = Len(txtText)
130           txtText.SetFocus
140       Case 4:    'Details
150           If EditDescription Then
160               txtText.Move g.Left + g.CellLeft, _
                               g.Top + g.CellTop, _
                               g.CellWidth, g.CellHeight
170               txtText.Text = g.TextMatrix(g.Row, g.Col)
180               txtText.MaxLength = 1000
190               txtText.Visible = True
200               txtText.SelStart = 0
210               txtText.SelLength = Len(txtText)
220               txtText.SetFocus
230           End If
240       Case 5:    'Type
250           If EditType Then
260               PopulateTypes cmbCombo
270               cmbCombo.Move g.Left + g.CellLeft, _
                                g.Top + g.CellTop, _
                                g.CellWidth
280               cmbCombo.Text = g.TextMatrix(g.Row, g.Col)
290               cmbCombo.Visible = True
300               cmbCombo.SelStart = 0
310               cmbCombo.SelLength = Len(cmbCombo.Text)
320               cmbCombo.SetFocus
330           End If
340       Case 6:    'Category
350           If EditCategory Then
360               PopulateCategories cmbCombo
370               cmbCombo.Move g.Left + g.CellLeft, _
                                g.Top + g.CellTop, _
                                g.CellWidth
380               cmbCombo.Text = g.TextMatrix(g.Row, g.Col)
390               cmbCombo.Visible = True
400               cmbCombo.SelStart = 0
410               cmbCombo.SelLength = Len(cmbCombo.Text)
420               cmbCombo.SetFocus
430           End If
440       End Select

450       Exit Sub

LoadControls_Error:

          Dim strES As String
          Dim intEL As Integer

460       intEL = Erl
470       strES = Err.Description
480       LogError "frmOptions", "LoadControls", intEL, strES

End Sub

Private Sub cmbCategories_DropDown()
10        PopulateCategories cmbCategories, True
End Sub

Private Sub cmbCombo_Click()
10        g.TextMatrix(g.Row, g.Col) = cmbCombo.Text
20        g.CellForeColor = vbRed
30        g.TextMatrix(g.Row, 0) = "*"

End Sub

Private Sub cmbCombo_KeyUp(KeyCode As Integer, Shift As Integer)
      'If KeyCode = vbKeyUp Then
      '    cmbCombo.Text = ""
      '    GoOneRowUp
      'ElseIf KeyCode = vbKeyDown Then
      '    cmbCombo.Text = ""
      '    GoOneRowDown
      'ElseIf KeyCode = 13 Then
      '    cmbCombo.Visible = False
      '    cmdUpdate.SetFocus
      'Else
      '
      '
      'End If
10        g.TextMatrix(g.Row, g.Col) = cmbCombo.Text
20        g.CellForeColor = vbRed
30        g.TextMatrix(g.Row, 0) = "*"
End Sub


Private Sub cmbOptionCategory_DropDown()
10        PopulateCategories cmbOptionCategory
End Sub



Private Sub cmbOptionType_DropDown()
10        PopulateTypes cmbOptionType
End Sub



Private Sub cmbTypes_DropDown()
10        PopulateTypes cmbTypes, True
20        Me.SelectedType = cmbTypes
End Sub

Private Sub cmdSave_Click()

10        On Error GoTo cmdSave_Click_Error

20        SaveOptionSettingEx txtOptionName, txtOptionValue, txtOptionDetail, cmbOptionCategory, cmbOptionType

30        Exit Sub

cmdSave_Click_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmOptions", "cmdSave_Click", intEL, strES

End Sub

Private Sub cmdUpdate_Click()

          Dim i As Integer

10        On Error GoTo cmdUpdate_Click_Error

20        For i = 1 To g.Rows - 1
30            g.Row = i
40            g.Col = 0
50            If g.TextMatrix(i, 0) = "*" Then
60                SaveOptionSettingEx g.TextMatrix(i, 1), g.TextMatrix(i, 3), g.TextMatrix(i, 4), g.TextMatrix(i, 6), g.TextMatrix(i, 5), g.Row
70                g.CellForeColor = vbBlack
80            End If
90        Next
100       LoadAllOptions Me.SelectedType, Me.SelectedCategory
110       Exit Sub

cmdUpdate_Click_Error:

          Dim strES As String
          Dim intEL As Integer

120       intEL = Erl
130       strES = Err.Description
140       LogError "frmOptions", "cmdUpdate_Click", intEL, strES

End Sub

Private Sub Form_Load()
10        On Error GoTo Form_Load_Error

20        g.ColWidth(1) = 0

30        If SelectedType <> "" And SelectedCategory <> "" Then
40            Me.Caption = Me.Caption & " --> " & Me.SelectedCategory
50            LoadAllOptions SelectedType, SelectedCategory
60        End If
70        txtOptionUsername = Username
80        FixComboWidth cmbTypes
90        FixComboWidth cmbCategories


100       Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

110       intEL = Erl
120       strES = Err.Description
130       LogError "frmOptions", "Form_Load", intEL, strES

End Sub

Private Sub g_Click()

10        On Error GoTo g_Click_Error

20        LoadControls
30        txtDescription = g.TextMatrix(g.Row, 4)

40        Exit Sub

g_Click_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmOptions", "g_Click", intEL, strES

End Sub

Private Sub g_EnterCell()

10        On Error GoTo g_EnterCell_Error

20        LoadControls
30        txtDescription = g.TextMatrix(g.Row, 4)

40        Exit Sub

g_EnterCell_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmOptions", "g_EnterCell", intEL, strES

End Sub

Private Sub g_LeaveCell()
10        txtText.Visible = False
20        cmbCombo.Visible = False
End Sub

Private Sub g_Scroll()
10        LoadControls
End Sub

Private Sub txtText_KeyUp(KeyCode As Integer, Shift As Integer)
10        If KeyCode = vbKeyUp Then
20            GoOneRowUp
30        ElseIf KeyCode = vbKeyDown Then
40            GoOneRowDown
50        ElseIf KeyCode = vbKeyRight Then
              'dont do anything
60        ElseIf KeyCode = vbKeyLeft Then
              'dont do anything
70        ElseIf KeyCode = 13 Then
80            txtText.Visible = False
90            cmdUpdate.SetFocus
100       Else
110           g.TextMatrix(g.Row, g.Col) = txtText
120           g.CellForeColor = vbRed
130           g.TextMatrix(g.Row, 0) = "*"
140       End If
End Sub

Private Sub GoOneRowUp()
10        If g.Row > 1 Then
20            txtText.Visible = False
30            g.Row = g.Row - 1
40            LoadControls
50        End If
End Sub
Private Sub GoOneRowDown()
10        If g.Row < g.Rows - 1 Then
20            txtText.Visible = False
30            g.Row = g.Row + 1
40            LoadControls
50        End If
End Sub



