VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form frmManager 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NetAcquire"
   ClientHeight    =   7635
   ClientLeft      =   2355
   ClientTop       =   3150
   ClientWidth     =   8490
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H80000008&
   Icon            =   "frmManager.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7635
   ScaleWidth      =   8490
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "Show"
      Height          =   2535
      Left            =   6270
      TabIndex        =   24
      Top             =   1320
      Width           =   1845
      Begin VB.CheckBox chkHistoLookUp 
         Caption         =   "HistoLookUp"
         Height          =   195
         Left            =   240
         TabIndex        =   30
         Top             =   2070
         Width           =   1305
      End
      Begin VB.CheckBox chkLookUp 
         Caption         =   "LookUp"
         Height          =   195
         Left            =   240
         TabIndex        =   29
         Top             =   1740
         Width           =   1245
      End
      Begin VB.CheckBox chkSecretarys 
         Caption         =   "Secretarys"
         Height          =   195
         Left            =   240
         TabIndex        =   28
         Top             =   1410
         Width           =   1245
      End
      Begin VB.CheckBox chkUsers 
         Caption         =   "Users"
         Height          =   195
         Left            =   240
         TabIndex        =   27
         Top             =   1080
         Width           =   945
      End
      Begin VB.CheckBox chkManagers 
         Caption         =   "Managers"
         Height          =   195
         Left            =   240
         TabIndex        =   26
         Top             =   750
         Width           =   1305
      End
      Begin VB.CheckBox chkAdministrators 
         Caption         =   "Administrators"
         Height          =   195
         Left            =   240
         TabIndex        =   25
         Top             =   420
         Value           =   1  'Checked
         Width           =   1305
      End
   End
   Begin VB.CommandButton cmdHide 
      Caption         =   "Exit"
      Height          =   675
      Left            =   4950
      Picture         =   "frmManager.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   270
      Width           =   1035
   End
   Begin VB.CommandButton cmdOK 
      Default         =   -1  'True
      Height          =   675
      Left            =   3780
      Picture         =   "frmManager.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   270
      Width           =   855
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1110
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "gort77"
      Top             =   690
      Width           =   2235
   End
   Begin VB.ComboBox cmbUserName 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmManager.frx":14DE
      Left            =   1110
      List            =   "frmManager.frx":14E0
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   270
      Width           =   2505
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   3495
      Left            =   135
      TabIndex        =   13
      Top             =   4005
      Width           =   7995
      _ExtentX        =   14102
      _ExtentY        =   6165
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
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
      FormatString    =   "<In Use |<Operator Name                     |<Code            |<Member Of       |^Log Off Delay |<Password   |"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Add New Operator"
      Height          =   2535
      Left            =   180
      TabIndex        =   9
      Top             =   1320
      Width           =   5940
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   800
         Left            =   4515
         Picture         =   "frmManager.frx":14E2
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   1620
         Width           =   1005
      End
      Begin VB.TextBox txtAutoLogOff 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1650
         TabIndex        =   19
         Text            =   "5"
         Top             =   2100
         Width           =   495
      End
      Begin VB.ComboBox cmbMemberOf 
         Height          =   315
         Left            =   1650
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1650
         Width           =   2205
      End
      Begin VB.TextBox txtConfirm 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1650
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   1200
         Width           =   2175
      End
      Begin VB.TextBox txtName 
         DataField       =   "opname"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1650
         MaxLength       =   20
         TabIndex        =   3
         Top             =   330
         Width           =   2175
      End
      Begin VB.TextBox txtCode 
         DataField       =   "opcode"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4680
         MaxLength       =   5
         TabIndex        =   4
         Top             =   300
         Width           =   675
      End
      Begin VB.TextBox txtPass 
         DataField       =   "oppass"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1650
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   870
         Width           =   2175
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   800
         Left            =   4515
         Picture         =   "frmManager.frx":1EE4
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   750
         Width           =   1005
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   285
         Left            =   2145
         TabIndex        =   20
         Top             =   2100
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   503
         _Version        =   327681
         Value           =   5
         BuddyControl    =   "txtAutoLogOff"
         BuddyDispid     =   196622
         OrigLeft        =   2400
         OrigTop         =   2130
         OrigRight       =   2895
         OrigBottom      =   2370
         Max             =   999
         Min             =   1
         Orientation     =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Minutes"
         Height          =   195
         Left            =   2640
         TabIndex        =   21
         Top             =   2160
         Width           =   555
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Auto Log Off in"
         Height          =   195
         Left            =   480
         TabIndex        =   18
         Top             =   2130
         Width           =   1065
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Member Of"
         Height          =   195
         Left            =   750
         TabIndex        =   17
         Top             =   1710
         Width           =   780
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Confirm Password"
         Height          =   195
         Left            =   270
         TabIndex        =   14
         Top             =   1260
         Width           =   1275
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         Height          =   195
         Left            =   1110
         TabIndex        =   12
         Top             =   360
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Password"
         Height          =   195
         Left            =   840
         TabIndex        =   11
         Top             =   900
         Width           =   690
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Code"
         Height          =   195
         Left            =   4260
         TabIndex        =   10
         Top             =   330
         Width           =   375
      End
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Password"
      Height          =   195
      Left            =   300
      TabIndex        =   16
      Top             =   750
      Width           =   690
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "User Name"
      Height          =   195
      Left            =   210
      TabIndex        =   15
      Top             =   330
      Width           =   795
   End
End
Attribute VB_Name = "frmManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private temp As String

Private mLookUp As Boolean
Private mOperator As Boolean
Private mManager As Boolean
Private mAdministrator As Boolean
Private mSecretary As Boolean
Private mSysManager As Boolean
Private mHistoLookUp As Boolean
Private LoginCount As Long

Public Property Let Administrator(ByVal ShowAdministrator As Boolean)

10        On Error GoTo Administrator_Error

20        mAdministrator = ShowAdministrator

30        Exit Property

Administrator_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmManager", "Administrator", intEL, strES


End Property
Private Sub EnsureAtLeastOne()

10        If chkAdministrators.Value = 0 And _
             chkManagers.Value = 0 And _
             chkUsers.Value = 0 And _
             chkSecretarys.Value = 0 And _
             chkLookUp.Value = 0 And _
             chkHistoLookUp.Value = 0 Then

20            chkAdministrators.Value = 1

30        End If

End Sub

Public Property Let SysManager(ByVal ShowSysManager As Boolean)

10        On Error GoTo SysManager_Error

20        mSysManager = ShowSysManager

30        Exit Property

SysManager_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmManager", "SysManager", intEL, strES


End Property

Private Sub chkAdministrators_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

10        EnsureAtLeastOne

20        FillG

End Sub


Private Sub chkHistoLookUp_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

10        EnsureAtLeastOne

20        FillG

End Sub

Private Sub chkLookUp_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

10        EnsureAtLeastOne

20        FillG

End Sub


Private Sub chkManagers_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

10        EnsureAtLeastOne

20        FillG

End Sub


Private Sub chkSecretarys_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

10        EnsureAtLeastOne

20        FillG

End Sub


Private Sub chkUsers_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

10        EnsureAtLeastOne

20        FillG

End Sub


Private Sub cmbUserName_Click()

10        txtPassword = ""

End Sub

Private Sub cmbUserName_LostFocus()

          Dim sql As String
          Dim tb As New Recordset

10        On Error GoTo cmbUserName_LostFocus_Error

20        If Trim(cmbUserName) = "" Then Exit Sub

30        cmbUserName = initial2upper(cmbUserName)

40        sql = "SELECT Name FROM Users WHERE " & _
                "Name LIKE '" & AddTicks(cmbUserName) & "%' " & _
                "AND InUse = 1 " & _
                "AND MemberOf <> 'LookUp'"
50        Set tb = New Recordset
60        RecOpenServer 0, tb, sql
70        If tb.EOF Then
80            iMsg "User Name " & cmbUserName & " is incorrect !"
90            cmbUserName = ""
100           cmbUserName.SetFocus
110       Else
120           cmbUserName = tb!Name
130       End If

140       Exit Sub

cmbUserName_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

150       intEL = Erl
160       strES = Err.Description
170       LogError "frmManager", "cmbUserName_LostFocus", intEL, strES, sql

End Sub

Private Sub cmdCancel_Click()

10        On Error GoTo cmdCancel_Click_Error

          '20    FillG

20        FillCombo

30        cmbUserName.ListIndex = -1
40        txtPassword = ""
50        UserName = ""
60        Me.Width = 6165
70        Me.Height = 1635

80        Exit Sub

cmdCancel_Click_Error:

          Dim strES As String
          Dim intEL As Integer

90        intEL = Erl
100       strES = Err.Description
110       LogError "frmManager", "cmdCancel_Click", intEL, strES

End Sub

Private Sub cmdHide_Click()

10        Unload Me

End Sub

Private Sub cmdOK_Click()

          Dim tb As New Recordset
          Dim sql As String

10        On Error GoTo cmdOK_Click_Error

20        sql = "SELECT * from Users WHERE " & _
                "Name = '" & AddTicks(cmbUserName) & "' " & _
                "and Password = '" & AddTicks(txtPassword) & "' " & _
                "and InUse = 1"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        If Not tb.EOF Then
60            UserName = cmbUserName
70            UserMemberOf = tb!MemberOf & ""
80            UserCode = tb!Code & ""
90            LogOffDelayMin = Val(tb!LogOffDelay)
100           LogOffDelaySecs = Val(tb!LogOffDelay) * 60
110           If SysOptPNE(0) = True Then
120               If DateDiff("d", tb!PassDate, Now) > Val(GetOptionSetting("PWDEXPIRYDAYS", "45")) Then
130                   NewPass
140               Else
150                   UserPass = UCase(tb!PassWord & "")
160               End If
170           Else
180               UserPass = UCase(tb!PassWord & "")
190           End If
200           If UserMemberOf = "Administrators" Then
210               FillG
220               Me.Width = 8580
230               Me.Height = 8040
240               Exit Sub
250           End If
260           LoadUserOpts
270           Me.Hide
280           frmMain.Show
290       Else
300           LoginCount = LoginCount + 1
310           If LoginCount = 3 Then
320               iMsg "3 Logins tried! Program will now close!" & vbCrLf & "Contact System Administrator!"
330               End
340           End If
350           txtPassword = ""
360       End If

370       Exit Sub

cmdOK_Click_Error:

          Dim strES As String
          Dim intEL As Integer

380       intEL = Erl
390       strES = Err.Description
400       LogError "frmManager", "cmdOK_Click", intEL, strES, sql

End Sub

Private Sub cmdSave_Click()

          Dim tb As New Recordset
          Dim sql As String

10        On Error GoTo cmdSave_Click_Error

20        txtName = Trim$(txtName)
30        txtCode = UCase$(Trim$(txtCode))
40        txtPass = UCase$(Trim$(txtPass))
50        txtConfirm = UCase$(Trim$(txtConfirm))

60        If Val(txtAutoLogOff) < 1 Then
70            txtAutoLogOff = "5"
80        End If

90        If txtName = "" Or txtCode = "" Or txtPass = "" Then
100           iMsg "Must have Name, Code," & vbCrLf & "and Password.", vbCritical
110           Exit Sub
120       End If

130       If cmbMemberOf = "" Then
140           iMsg "Member Of ???", vbCritical
150           Exit Sub
160       End If

170       If txtPass <> txtConfirm Then
180           txtPass = ""
190           txtConfirm = ""
200           iMsg "Password/Confirm don't match." & vbCrLf & _
                   "Retype Password and Confirmation", vbCritical
210           Exit Sub
220       End If

230       sql = "SELECT * from Users WHERE Name = '" & AddTicks(Trim$(txtName)) & "'"
240       Set tb = New Recordset
250       RecOpenServer 0, tb, sql
260       If Not tb.EOF Then
270           iMsg "Name already used.", vbExclamation
280           txtName = ""
290           Exit Sub
300       End If

310       sql = "SELECT * from Users WHERE Code = '" & Trim$(txtCode) & "'"
320       Set tb = New Recordset
330       RecOpenServer 0, tb, sql
340       If Not tb.EOF Then
350           iMsg "Code already used.", vbExclamation
360           txtCode = ""
370           Exit Sub
380       End If

390       sql = "SELECT * from Users WHERE Password = '" & Trim$(txtPass) & "'"
400       Set tb = New Recordset
410       RecOpenServer 0, tb, sql
420       If Not tb.EOF Then
430           iMsg "Password already used." & vbCrLf & "Type another Password.", vbExclamation
440           txtPass = ""
450           txtConfirm = ""
460           Exit Sub
470       End If

480       tb.AddNew
490       tb!LogOffDelay = Val(txtAutoLogOff)
500       tb!Code = txtCode
510       tb!Name = initial2upper(txtName)
520       tb!PassWord = txtPass
530       tb!MemberOf = cmbMemberOf
540       tb!InUse = True
550       tb!ListOrder = 999
560       tb!PassDate = Format$(Now - (Val(GetOptionSetting("PWDEXPIRYDAYS", "45")) + 5), "dd/mmm/yyyy")
          'tb!PNE = SysOptExp(0)
570       tb.Update

580       FillG

590       txtCode = ""
600       txtName = ""
610       txtPass = ""
620       txtConfirm = ""
630       txtAutoLogOff = "5"

640       Exit Sub

cmdSave_Click_Error:

          Dim strES As String
          Dim intEL As Integer

650       intEL = Erl
660       strES = Err.Description
670       LogError "frmManager", "cmdsave_Click", intEL, strES, sql

End Sub

Private Sub FillG()

          Dim s As String
          Dim tb As New Recordset
          Dim sql As String

10        On Error GoTo FillG_Error

20        Screen.MousePointer = vbHourglass

30        With g
40            .Visible = False
50            .Rows = 2
60            .AddItem ""
70            .RemoveItem 1
80        End With

          '80    cmbUserName.Clear

90        sql = "SELECT " & _
                "CASE CAST(COALESCE(Inuse, '0') AS nvarchar(3)) " & _
                "     WHEN '1' THEN 'Yes' " & _
                "     ELSE 'No' " & _
                "END + CHAR(9) " & _
                "+ UPPER(Name) + CHAR(9) " & _
                "+ Code + CHAR(9) " & _
                "+ MemberOf + CHAR(9) " & _
                "+ CONVERT(nvarchar, LogOffDelay) + CHAR(9) " & _
                "+ '*****' + CHAR(9) " & _
                "+ Password Details FROM Users WHERE " & _
                "MemberOf IN ("

100       If chkLookUp.Value = 1 Then
110           sql = sql & "'LookUp', "
120       End If
130       If chkUsers.Value = 1 Then
140           sql = sql & "'Users', "
150       End If
160       If chkManagers.Value = 1 Then
170           sql = sql & "'Managers', "
180       End If
190       If chkSecretarys.Value = 1 Then
200           sql = sql & "'Secretarys', "
210       End If
220       If chkAdministrators.Value = 1 Then
230           sql = sql & "'Administrators', "
240       End If
250       If chkHistoLookUp.Value = 1 Then
260           sql = sql & "'HistoLookUp', "
270       End If
280       If Right$(sql, 2) = ", " Then
290           sql = Left$(sql, Len(sql) - 2)
300       End If
310       sql = sql & ") "
320       If SysOptAlphaOrderTechnicians(0) Then
330           sql = sql & "ORDER BY Name Asc"
340       Else
350           sql = sql & "ORDER BY ListOrder"
360       End If
370       Set tb = New Recordset
380       Set tb = Cnxn(0).Execute(sql)
          '380   RecOpenServer 0, tb, sql
390       Do While Not tb.EOF
400           s = tb!Details
410           g.AddItem s

420           tb.MoveNext
430       Loop

440       If g.Rows > 2 Then
450           g.RemoveItem 1
460       End If

470       g.Visible = True

480       Screen.MousePointer = vbNormal

490       Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer

500       intEL = Erl
510       strES = Err.Description
520       LogError "frmManager", "FillG", intEL, strES, sql
530       g.Visible = True

540       Screen.MousePointer = vbNormal

End Sub

Private Sub FillCombo()

          Dim tb As New Recordset
          Dim sql As String

10        On Error GoTo FillCombo_Error

20        Screen.MousePointer = vbHourglass

30        cmbUserName.Tag = ""
40        If cmbUserName <> "" Then cmbUserName.Tag = cmbUserName
50        cmbUserName.Clear

60        sql = "SELECT UPPER(Name) Name FROM Users " & _
                "WHERE InUse = 1 " & _
                "AND MemberOf IN ( "

70        If mLookUp Then
80            sql = sql & "'LookUp', "
90        End If
100       If mOperator Then
110           sql = sql & "'Users', "
120       End If
130       If mManager Then
140           sql = sql & "'Managers', "
150       End If
160       If mSecretary Then
170           sql = sql & "'Secretarys', "
180       End If
190       If mAdministrator Then
200           sql = sql & "'Administrators', "
210       End If
220       If mHistoLookUp Then
230           sql = sql & "'HistoLookUp', "
240       End If
250       If Right$(sql, 2) = ", " Then
260           sql = Left$(sql, Len(sql) - 2) & " "
270       End If
280       sql = sql & ")"
290       If SysOptAlphaOrderTechnicians(0) Then
300           sql = sql & "ORDER BY Name Asc"
310       Else
320           sql = sql & "ORDER BY ListOrder"
330       End If
340       Set tb = New Recordset


350       Set tb = Cnxn(0).Execute(sql)
360       Do While Not tb.EOF
370           cmbUserName.AddItem tb!Name & ""
380           tb.MoveNext
390       Loop

400       If cmbUserName.Tag <> "" Then cmbUserName = cmbUserName.Tag
410       Screen.MousePointer = vbDefault

420       Exit Sub

FillCombo_Error:

          Dim strES As String
          Dim intEL As Integer

430       intEL = Erl
440       strES = Err.Description
450       LogError "frmManager", "FillCombo", intEL, strES, sql

460       Screen.MousePointer = vbDefault

End Sub


Private Sub Form_Activate()

10        On Error GoTo Form_Activate_Error

20        FillCombo

30        If TestSys = True Then Me.Caption = Me.Caption & " - TEST SYSTEM"

40        If UserMemberOf <> "Administrators" Then
50            Me.Width = 6165
60            Me.Height = 1635
70        End If

80        txtPassword = ""

90        LoginCount = 0


100       Exit Sub

Form_Activate_Error:

          Dim strES As String
          Dim intEL As Integer

110       intEL = Erl
120       strES = Err.Description
130       LogError "frmManager", "Form_Activate", intEL, strES

End Sub

Private Sub Form_Load()

10        On Error GoTo Form_Load_Error

20        Me.Width = 5715
30        Me.Height = 1635

40        g.ColWidth(6) = 0

50        txtPassword = ""

60        With cmbMemberOf
70            .Clear
80            .AddItem "Administrators"
90            .AddItem "LookUp"
100           .AddItem "Managers"
110           .AddItem "Secretarys"
120           .AddItem "Users"
130           .AddItem "HistoLookUp"
140       End With

          '140   FillG

150       Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

160       intEL = Erl
170       strES = Err.Description
180       LogError "frmManager", "Form_Load", intEL, strES

End Sub

Private Sub g_Click()

          Static SortOrder As Boolean
          Dim LogOff As String
          Dim NewPassw As String
          Dim sql As String
          Dim tmpN As Integer

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

110       Select Case g.Col

          Case 0:    'In Use
120           If g.TextMatrix(g.Row, 1) <> "Administrator" Then
130               g.TextMatrix(g.Row, 0) = IIf(g.TextMatrix(g.Row, 0) = "No", "Yes", "No")
140               sql = "UPDATE Users " & _
                        "set InUse = " & IIf(g.TextMatrix(g.Row, 0) = "No", 0, 1) & " " & _
                        "WHERE Name = '" & AddTicks(g.TextMatrix(g.Row, 1)) & "'"
150               Cnxn(0).Execute sql
160           End If
170       Case 3:
180           If g.TextMatrix(g.Row, 3) = "Users" Then
190               g.TextMatrix(g.Row, 3) = "Managers"
200           ElseIf g.TextMatrix(g.Row, 3) = "Managers" Then
210               g.TextMatrix(g.Row, 3) = "Secretarys"
220           ElseIf g.TextMatrix(g.Row, 3) = "Secretarys" Then
230               g.TextMatrix(g.Row, 3) = "LookUp"
240           ElseIf g.TextMatrix(g.Row, 3) = "LookUp" Then
250               g.TextMatrix(g.Row, 3) = "Users"
260           End If
270           sql = "UPDATE users " & _
                    "set memberof = '" & g.TextMatrix(g.Row, 3) & "' " & _
                    "WHERE NAME = '" & AddTicks(g.TextMatrix(g.Row, 1)) & "'"
280           Cnxn(0).Execute sql

290       Case 4:    'Log Off Delay
300           g.Enabled = False
310           LogOff = g.TextMatrix(g.Row, 4)
320           LogOff = iBOX("Log Off Delay. (Minutes)", , LogOff)
330           If LogOff = "" Then
340               g.Enabled = True
350               Exit Sub
360           End If
370           g.TextMatrix(g.Row, 4) = Format$(Val(LogOff))
380           sql = "UPDATE Users " & _
                    "set LogOffDelay = " & Val(LogOff) & " " & _
                    "WHERE Name = '" & g.TextMatrix(g.Row, 1) & "'"
390           Cnxn(0).Execute sql
400           g.Enabled = True

410       Case 5:
420           g.Enabled = False
430           NewPassw = iBOX("Reset Password", , , True)
440           If NewPassw = "" Then
450               g.Enabled = True
460               Exit Sub
470           End If
480           sql = "Update Users set PassWord = '" & NewPassw & "', " & _
                    "PassDate = '" & Format(Now - (Val(GetOptionSetting("PWDEXPIRYDAYS", "45")) + 5), "dd/MMM/yyyy") & "' " & _
                    "WHERE Name = '" & g.TextMatrix(g.Row, 1) & "' AND password = '" & g.TextMatrix(g.Row, 6) & "'"
490           Cnxn(0).Execute sql
500           g.Enabled = True

510       Case Is > 7:
              '   g.Enabled = False
              '   If g.TextMatrix(g.Row, g.Col) = "" Then
              '    g.TextMatrix(g.Row, g.Col) = "?"
520       End Select

530       tmpN = g.RowSel

540       FillG

550       g.RowSel = tmpN

560       Exit Sub

g_Click_Error:

          Dim strES As String
          Dim intEL As Integer

570       intEL = Erl
580       strES = Err.Description
590       LogError "frmManager", "g_Click", intEL, strES, sql

End Sub

Private Sub g_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

          Dim n As Long
          Static PrevY As Long

10        On Error GoTo g_MouseMove_Error

20        If g.MouseRow > 0 And g.MouseCol = 5 Then

30            If SysOptPNE(0) <> True Then g.ToolTipText = g.TextMatrix(g.MouseRow, 6)
40            Exit Sub
50        ElseIf g.MouseRow > 0 And g.MouseCol = 0 Then
60            g.ToolTipText = "Click to Toggle Yes/No"
70            Exit Sub
80        ElseIf g.MouseRow > 0 And g.MouseCol = 1 And Not SysOptAlphaOrderTechnicians(0) Then
90            g.ToolTipText = "Drag to change List Order"
100       Else
110           g.ToolTipText = ""
120       End If

130       If Button = vbLeftButton And g.MouseRow > 0 And g.MouseCol = 1 Then
140           If temp = "" Then
150               PrevY = g.MouseRow
160               For n = 0 To g.Cols - 1
170                   temp = temp & g.TextMatrix(g.Row, n) & vbTab
180               Next
190               temp = Left$(temp, Len(temp) - 1)
200               Exit Sub
210           Else
220               If g.MouseRow <> PrevY Then
230                   g.RemoveItem PrevY
240                   If g.MouseRow <> PrevY Then
250                       g.AddItem temp, g.MouseRow
260                       PrevY = g.MouseRow
270                   Else
280                       g.AddItem temp
290                       PrevY = g.Rows - 1
300                   End If
310               End If
320           End If
330       End If

340       Exit Sub

g_MouseMove_Error:

          Dim strES As String
          Dim intEL As Integer

350       intEL = Erl
360       strES = Err.Description
370       LogError "frmManager", "g_MouseMove", intEL, strES

End Sub

Private Sub g_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

          Dim gy As Long
          Dim sql As String

10        On Error GoTo g_MouseUp_Error

20        If Not SysOptAlphaOrderTechnicians() Then

30            For gy = 1 To g.Rows - 1

40                sql = "UPDATE users set ListOrder = " & gy & " " & _
                        "WHERE name = '" & AddTicks(g.TextMatrix(gy, 1)) & "'"
50                Cnxn(0).Execute sql

60            Next

70        End If

80        temp = ""




90        Exit Sub

g_MouseUp_Error:

          Dim strES As String
          Dim intEL As Integer



100       intEL = Erl
110       strES = Err.Description
120       LogError "frmManager", "g_MouseUp", intEL, strES, sql


End Sub

Public Property Let LookUp(ByVal ShowLookUp As Boolean)

10        On Error GoTo LookUp_Error

20        mLookUp = ShowLookUp

30        Exit Property

LookUp_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmManager", "LookUp", intEL, strES


End Property

Public Property Let Manager(ByVal ShowManager As Boolean)

10        On Error GoTo Manager_Error

20        mManager = ShowManager

30        Exit Property

Manager_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmManager", "Manager", intEL, strES


End Property

Private Sub NewPass()

          Dim NewPass1 As String
          Dim NewPass2 As String
          Dim sql As String
          Dim tb As New Recordset
          Dim Changed As Boolean

10        On Error GoTo NewPass_Error

20        Changed = False

30        Do While Not Changed

40            Changed = True
50            NewPass1 = iBOX("Enter New PassWord", , , True)
60            NewPass2 = iBOX("Verify New PassWord", , , True)

70            Do While NewPass1 <> NewPass2 Or Len(NewPass1) < 6
80                If NewPass1 <> NewPass2 Then
90                    NewPass1 = iBOX("Passwords don't match!" & vbCrLf & "Enter New PassWord", , , True)
100               Else
110                   NewPass1 = iBOX("Password must be at least six Characters." & vbCrLf & "Enter New PassWord", , , True)
120               End If
130               NewPass2 = iBOX("Verify New PassWord", , , True)
140           Loop
              '150     temp = Cnxn(0).ConnectionString
150           sql = "SELECT * from UsersAudit WHERE " & _
                    "PassWord = '" & NewPass1 & "' " & _
                    "AND Name = '" & AddTicks(UserName) & "' AND DATEDIFF(m,DateTimeOfRecord,GETDATE()) > 12 "
160           Set tb = New Recordset
170           RecOpenServer 0, tb, sql
180           If Not tb.EOF Then
190               iMsg "Password previously used", vbExclamation
200               Changed = False
210           End If

220           sql = "SELECT * FROM Users WHERE " & _
                    "PassWord = '" & NewPass1 & "'"
230           Set tb = New Recordset
240           RecOpenServer 0, tb, sql
250           If Not tb.EOF Then
260               iMsg "Password in use or been used!", vbExclamation
270               Changed = False
280           End If
290       Loop

300       sql = "IF EXISTS (SELECT * FROM Users WHERE " & _
                "           Name = '" & AddTicks(UserName) & "') " & _
                "  UPDATE Users " & _
                "  SET Password = '" & UCase$(NewPass1) & "', " & _
                "  Passdate = getdate() " & _
                "  WHERE Name = '" & AddTicks(UserName) & "'"
310       Cnxn(0).Execute sql

320       iMsg UserName & " your Password is now Changed. Thank You!", vbInformation

330       UserPass = UCase(NewPass1)

340       Exit Sub

NewPass_Error:

          Dim strES As String
          Dim intEL As Integer

350       intEL = Erl
360       strES = Err.Description
370       LogError "frmManager", "NewPass", intEL, strES, sql

End Sub

Public Property Let Operator(ByVal ShowOperator As Boolean)

10        On Error GoTo Operator_Error

20        mOperator = ShowOperator

30        Exit Property

Operator_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmManager", "Operator", intEL, strES


End Property

Public Property Let Secretary(ByVal ShowSecretary As Boolean)

10        On Error GoTo Secretary_Error

20        mSecretary = ShowSecretary

30        Exit Property

Secretary_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmManager", "Secretary", intEL, strES


End Property



Private Sub txtPassword_LostFocus()

10        On Error GoTo txtPassword_LostFocus_Error

20        txtPassword = UCase$(txtPassword)

30        Exit Sub

txtPassword_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmManager", "txtPassword_LostFocus", intEL, strES


End Sub


Public Property Let HistoLookUp(ByVal bNewValue As Boolean)

10        mHistoLookUp = bNewValue

End Property
