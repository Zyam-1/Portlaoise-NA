VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmMicroOrderUrine 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Urine Order"
   ClientHeight    =   4170
   ClientLeft      =   5910
   ClientTop       =   5655
   ClientWidth     =   6030
   Icon            =   "frmMicroOrderUrine.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin MSFlexGridLib.MSFlexGrid grdUrineTests 
      Height          =   2295
      Left            =   3840
      TabIndex        =   15
      Top             =   660
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   4048
      _Version        =   393216
      Rows            =   1
      Cols            =   3
      FormatString    =   "|Test                 |    "
   End
   Begin VB.TextBox txtSampleID 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1560
      MaxLength       =   12
      TabIndex        =   5
      Top             =   120
      Width           =   1545
   End
   Begin VB.Frame frUrine 
      Caption         =   "Urine Requests"
      Height          =   1005
      Left            =   570
      TabIndex        =   2
      Top             =   1950
      Width           =   3105
      Begin VB.CheckBox chkUrine 
         Caption         =   "Red Sub"
         Height          =   225
         Index           =   2
         Left            =   1470
         TabIndex        =   14
         Top             =   570
         Width           =   975
      End
      Begin VB.CheckBox chkUrine 
         Alignment       =   1  'Right Justify
         Caption         =   "C && S"
         Height          =   195
         Index           =   0
         Left            =   630
         TabIndex        =   4
         Top             =   270
         Width           =   765
      End
      Begin VB.CheckBox chkUrine 
         Caption         =   "Pregnancy"
         Height          =   195
         Index           =   1
         Left            =   1470
         TabIndex        =   3
         Top             =   270
         Width           =   1155
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   735
      Left            =   2205
      Picture         =   "frmMicroOrderUrine.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3195
      Width           =   1275
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   735
      Left            =   840
      Picture         =   "frmMicroOrderUrine.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3195
      Width           =   1275
   End
   Begin VB.Frame Frame2 
      Caption         =   "Urine Sample"
      Height          =   1245
      Left            =   585
      TabIndex        =   7
      Top             =   585
      Width           =   3090
      Begin VB.OptionButton optU 
         Caption         =   "EMU"
         Height          =   195
         Index           =   5
         Left            =   1470
         TabIndex        =   13
         Top             =   900
         Width           =   675
      End
      Begin VB.OptionButton optU 
         Alignment       =   1  'Right Justify
         Caption         =   "FVU"
         Height          =   195
         Index           =   4
         Left            =   780
         TabIndex        =   12
         Top             =   900
         Width           =   645
      End
      Begin VB.OptionButton optU 
         Alignment       =   1  'Right Justify
         Caption         =   "MSU"
         Height          =   195
         Index           =   0
         Left            =   750
         TabIndex        =   11
         Top             =   360
         Value           =   -1  'True
         Width           =   675
      End
      Begin VB.OptionButton optU 
         Caption         =   "CSU"
         Height          =   195
         Index           =   1
         Left            =   1470
         TabIndex        =   10
         Top             =   360
         Width           =   645
      End
      Begin VB.OptionButton optU 
         Alignment       =   1  'Right Justify
         Caption         =   "BSU"
         Height          =   195
         Index           =   2
         Left            =   780
         TabIndex        =   9
         Top             =   630
         Width           =   645
      End
      Begin VB.OptionButton optU 
         Caption         =   "SPA"
         Height          =   195
         Index           =   3
         Left            =   1470
         TabIndex        =   8
         Top             =   630
         Width           =   645
      End
   End
   Begin VB.Image imgRedCross 
      Height          =   225
      Left            =   0
      Picture         =   "frmMicroOrderUrine.frx":091E
      Top             =   0
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image imgGreenTick 
      Height          =   225
      Left            =   0
      Picture         =   "frmMicroOrderUrine.frx":0BF4
      Top             =   210
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label Label42 
      AutoSize        =   -1  'True
      Caption         =   "Sample ID"
      Height          =   195
      Left            =   780
      TabIndex        =   6
      Top             =   150
      Width           =   735
   End
End
Attribute VB_Name = "frmMicroOrderUrine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkUrine_Click(Index As Integer)

          Dim n As Integer

10        On Error GoTo chkUrine_Click_Error

20        cmdSave.Enabled = True

30        If Index = 6 Then
40            For n = 0 To 5
50                chkUrine(n).Value = 0
60            Next
70        End If

80        Exit Sub

chkUrine_Click_Error:

          Dim strES As String
          Dim intEL As Integer

90        intEL = Erl
100       strES = Err.Description
110       LogError "frmMicroOrderUrine", "chkUrine_Click", intEL, strES

End Sub

Private Sub cmdCancel_Click()
          Dim n As Long

10        On Error GoTo cmdCancel_Click_Error

20        For n = 0 To 5
30            optU(n).Value = False
40        Next

50        If cmdSave.Enabled Then
60            If iMsg("Cancel without Saving?", vbQuestion = vbYesNo) = vbNo Then
70                Exit Sub
80            End If
90        End If

100       SaveDetails

110       Me.Hide

120       Exit Sub

cmdCancel_Click_Error:

          Dim strES As String
          Dim intEL As Integer



130       intEL = Erl
140       strES = Err.Description
150       LogError "frmMicroOrderUrine", "cmdCancel_Click", intEL, strES


End Sub


Private Sub cmdSave_Click()

10        On Error GoTo cmdSave_Click_Error

20        SaveDetails

30        Me.Hide

40        Exit Sub

cmdSave_Click_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmMicroOrderUrine", "cmdsave_Click", intEL, strES

End Sub

Private Sub InitializeGrid()
On Error GoTo InitailzeGrid_Error

With grdUrineTests
    .Rows = 2
    .Cols = 3
    .Rows = 1
    
    .ColWidth(0) = 0
    .ColWidth(1) = 1400
    .ColWidth(2) = 240
End With

Exit Sub
InitailzeGrid_Error:
   
LogError "frmMicroOrderUrine", "InitializeGrid", Erl, Err.Description


End Sub

Private Sub Form_Activate()

10        LoadDetails

End Sub

Private Sub LoadDetails()

      Dim tb         As Recordset
      Dim sql        As String
      Dim SampleIDWithOffset As Double
      Dim n          As Long

10    On Error GoTo LoadDetails_Error

20    InitializeGrid
30    SampleIDWithOffset = Val(txtSampleID) + SysOptMicroOffset(0)

40    For n = 0 To 2
50        chkUrine(n) = 0
60    Next

70    sql = "SELECT " & _
            "COALESCE(CS, 0) CS, " & _
            "COALESCE(Pregnancy, 0) Pregnancy, " & _
            "COALESCE(RedSub, 0) RedSub " & _
            "FROM UrineRequests WHERE " & _
            "SampleID = " & SampleIDWithOffset
80    Set tb = New Recordset
90    RecOpenServer 0, tb, sql
100   If Not tb.EOF Then
110       If tb!cS Then chkUrine(0) = 1
120       If tb!Pregnancy Then chkUrine(1) = 1
130       If tb!RedSub Then chkUrine(2) = 1
140   End If

150   sql = "select * from micrositedetails where sampleid = " & SampleIDWithOffset
160   Set tb = New Recordset
170   RecOpenServer 0, tb, sql
180   If Not tb.EOF Then
190       For n = 0 To 5
200           If optU(n).Caption = tb!SiteDetails Then
210               optU(n).Value = True
220           End If
230       Next
240   End If


      'Load Urine Tests

'250   sql = "SELECT distinct HostCode, LongName as Name, " & _
'            "PrintPriority, SampleType from IQ200TestDefinitions WHERE " & _
'            "KnownToAnalyser = '1' and inuse = '1'" & _
'            "order by PrintPriority"
'
'260   Set tb = New Recordset
'270   RecOpenServer 0, tb, sql
'280   Do While Not tb.EOF
'290       grdUrineTests.AddItem tb!HostCode & vbTab & tb!Name
'300       grdUrineTests.Row = grdUrineTests.Rows - 1
'310       grdUrineTests.Col = 2
'320       Set grdUrineTests.CellPicture = imgRedCross
'330       tb.MoveNext
'340   Loop
'
'      'fill known orders
'350   sql = "SELECT BR.*, BT.LongName as Name " & _
'            "from IQ200Requests as BR, IQ200TestDefinitions as BT WHERE " & _
'            "SampleID = '" & Val(SampleIDWithOffset) & "' " & _
'            "and BR.Code = BT.HostCode " & _
'            "and BR.SampleType = 'U'"
'360   Set tb = New Recordset
'370   RecOpenServer 0, tb, sql
'380   Do While Not tb.EOF
'
'390       For n = 1 To grdUrineTests.Rows - 1
'400           If tb!Name = grdUrineTests.TextMatrix(n, 1) Then
'410               grdUrineTests.Row = n
'420               grdUrineTests.Col = 2
'430               Set grdUrineTests.CellPicture = imgGreenTick
'440               Exit For
'450           End If
'460       Next
'470       tb.MoveNext
'480   Loop

490   cmdSave.Enabled = False

500   Exit Sub

LoadDetails_Error:

      Dim strES      As String
      Dim intEL      As Integer

510   intEL = Erl
520   strES = Err.Description
530   LogError "frmMicroOrderUrine", "LoadDetails", intEL, strES, sql

End Sub

Private Sub SaveDetails()

      Dim sql        As String
      Dim SampleIDWithOffset As Double
      Dim i          As Integer

10    On Error GoTo SaveDetails_Error

20    SampleIDWithOffset = Val(txtSampleID) + SysOptMicroOffset(0)

30    sql = "IF EXISTS (SELECT * FROM UrineRequests WHERE " & _
            "           SampleID = " & SampleIDWithOffset & ") " & _
            "  UPDATE UrineRequests " & _
            "  SET CS = '" & IIf(chkUrine(0), 1, 0) & "', " & _
            "  Pregnancy = '" & IIf(chkUrine(1), 1, 0) & "', " & _
            "  RedSub = '" & IIf(chkUrine(2), 1, 0) & "', " & _
            "  UserName = '" & AddTicks(UserName) & "' " & _
            "  WHERE SampleID = " & SampleIDWithOffset & " " & _
            "ELSE " & _
            "  INSERT INTO UrineRequests " & _
            "  (SampleID, CS, Pregnancy, RedSub, UserName) VALUES " & _
            "  (" & SampleIDWithOffset & ", " & _
            "  '" & IIf(chkUrine(0), 1, 0) & "', " & _
            "  '" & IIf(chkUrine(1), 1, 0) & "', " & _
            "  '" & IIf(chkUrine(2), 1, 0) & "', " & _
            "  '" & AddTicks(UserName) & "')"
40    Cnxn(0).Execute sql

      'UPDATE URINE TABLE IF ITS NEW ENTRY

      'Created on 01/02/2011 15:49:30
      'Autogenerated by SQL Scripting

50    sql = "If Exists(Select 1 From Urine " & _
            "Where SampleID = @SampleID0) " & _
            "Begin " & _
            "Update Urine Set " & _
            "SampleID = @SampleID0, " & _
            "UserName = '@UserName25' " & _
            "Where SampleID = @SampleID0 " & _
            "End  " & _
            "Else " & _
            "Begin  " & _
            "Insert Into Urine (SampleID, UserName) Values " & _
            "(@SampleID0, '@UserName25') " & _
            "End"

60    sql = Replace(sql, "@SampleID0", SampleIDWithOffset)
70    sql = Replace(sql, "@UserName25", UserName)

80    Cnxn(0).Execute sql

'90    Cnxn(0).Execute ("DELETE from IQ200Requests WHERE " & _
'                       "SampleID = '" & Val(SampleIDWithOffset) & "' " & _
'                       "and Programmed = 0")

100   For i = 1 To grdUrineTests.Rows - 1
110       grdUrineTests.Row = i
120       grdUrineTests.Col = 2
130       If grdUrineTests.CellPicture = imgGreenTick Then
140           sql = "INSERT into IQ200Requests " & _
                    "(SampleID, Code, DateTime, SampleType, Programmed, AnalyserID, Hospital) VALUES " & _
                    "('" & Val(SampleIDWithOffset) & "', " & _
                    "'" & grdUrineTests.TextMatrix(i, 0) & "', " & _
                    "'" & Format$(Now, "dd/mmm/yyyy hh:mm") & "', " & _
                    "'U', " & _
                    "'0', " & _
                    "'UWAM', " & _
                    "'')"
150           Cnxn(0).Execute sql
160       End If
          
170   Next





180   SaveInitialMicroSiteDetails "Urine", SampleIDWithOffset, SiteDetails

190   cmdSave.Enabled = False

200   Exit Sub

SaveDetails_Error:

      Dim strES      As String
      Dim intEL      As Integer

210   intEL = Erl
220   strES = Err.Description
230   LogError "frmMicroOrderUrine", "SaveDetails", intEL, strES, sql

End Sub




Private Sub grdUrineTests_Click()

10    On Error GoTo grdUrineTests_Click_Error

20    grdUrineTests.Row = grdUrineTests.MouseRow
30    grdUrineTests.Col = 2
40    If grdUrineTests.CellPicture = imgRedCross Then
50        Set grdUrineTests.CellPicture = imgGreenTick
60    Else
70        Set grdUrineTests.CellPicture = imgRedCross
80    End If
90    cmdSave.Enabled = True


100   Exit Sub
grdUrineTests_Click_Error:
         
110   LogError "frmMicroOrderUrine", "grdUrineTests_Click", Erl, Err.Description


End Sub

Private Sub optU_Click(Index As Integer)

10        On Error GoTo optU_Click_Error

20        cmdSave.Enabled = True

30        Exit Sub

optU_Click_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmMicroOrderUrine", "optU_Click", intEL, strES


End Sub

Private Sub txtSampleID_LostFocus()

10        On Error GoTo txtSampleID_LostFocus_Error

20        LoadDetails

30        Exit Sub

txtSampleID_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmMicroOrderUrine", "txtSampleID_LostFocus", intEL, strES


End Sub

Public Property Get SiteDetails() As String

          Dim n As Long

10        On Error GoTo SiteDetails_Error

20        For n = 0 To 5
30            If optU(n) Then
40                SiteDetails = optU(n).Caption
50            End If
60        Next

70        Exit Property

SiteDetails_Error:

          Dim strES As String
          Dim intEL As Integer

80        intEL = Erl
90        strES = Err.Description
100       LogError "frmMicroOrderUrine", "SiteDetails", intEL, strES

End Property
