VERSION 5.00
Begin VB.Form frmMicroOrderFaeces 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Faeces Order"
   ClientHeight    =   3570
   ClientLeft      =   6240
   ClientTop       =   5310
   ClientWidth     =   4485
   Icon            =   "frmMicroOrderFaeces.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   4485
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   705
      Left            =   2340
      Picture         =   "frmMicroOrderFaeces.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2835
      Width           =   1275
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   705
      Left            =   780
      Picture         =   "frmMicroOrderFaeces.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2835
      Width           =   1275
   End
   Begin VB.Frame frFaeces 
      Caption         =   "Faecal Requests"
      Height          =   2025
      Left            =   180
      TabIndex        =   2
      Top             =   570
      Width           =   4245
      Begin VB.CheckBox chkFaecal 
         Caption         =   "Bad Request"
         Height          =   195
         Index           =   12
         Left            =   1890
         TabIndex        =   17
         Top             =   1665
         Width           =   1245
      End
      Begin VB.CheckBox chkFaecal 
         Alignment       =   1  'Right Justify
         Caption         =   "Crypto Screen"
         Height          =   240
         Index           =   11
         Left            =   270
         TabIndex        =   16
         Top             =   1665
         Width           =   1320
      End
      Begin VB.CheckBox chkFaecal 
         Caption         =   "E/P Coli"
         Height          =   195
         Index           =   9
         Left            =   1890
         TabIndex        =   13
         Top             =   1230
         Width           =   885
      End
      Begin VB.CheckBox chkFaecal 
         Alignment       =   1  'Right Justify
         Caption         =   "C && S"
         Height          =   195
         Index           =   0
         Left            =   930
         TabIndex        =   12
         Top             =   300
         Width           =   705
      End
      Begin VB.CheckBox chkFaecal 
         Alignment       =   1  'Right Justify
         Caption         =   "O/P"
         Height          =   195
         Index           =   2
         Left            =   1020
         TabIndex        =   11
         Top             =   570
         Width           =   615
      End
      Begin VB.CheckBox chkFaecal 
         Alignment       =   1  'Right Justify
         Caption         =   "Rota/Adeno"
         Height          =   195
         Index           =   6
         Left            =   450
         TabIndex        =   10
         Top             =   810
         Width           =   1185
      End
      Begin VB.CheckBox chkFaecal 
         Alignment       =   1  'Right Justify
         Caption         =   "Coli 0157"
         Height          =   195
         Index           =   8
         Left            =   630
         TabIndex        =   9
         Top             =   1440
         Width           =   975
      End
      Begin VB.CheckBox chkFaecal 
         Caption         =   "C.Difficile"
         Height          =   195
         Index           =   1
         Left            =   1890
         TabIndex        =   8
         Top             =   300
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.CheckBox chkFaecal 
         Alignment       =   1  'Right Justify
         Caption         =   "Toxin A/B"
         Height          =   195
         Index           =   7
         Left            =   585
         TabIndex        =   7
         Top             =   1230
         Width           =   1035
      End
      Begin VB.CheckBox chkFaecal 
         Caption         =   "Check9"
         Height          =   195
         Index           =   3
         Left            =   1890
         TabIndex        =   6
         Top             =   570
         Width           =   255
      End
      Begin VB.CheckBox chkFaecal 
         Caption         =   "Check10"
         Height          =   195
         Index           =   4
         Left            =   2130
         TabIndex        =   5
         Top             =   570
         Width           =   225
      End
      Begin VB.CheckBox chkFaecal 
         Caption         =   "Occult Blood"
         Height          =   195
         Index           =   5
         Left            =   2370
         TabIndex        =   4
         Top             =   570
         Width           =   1245
      End
      Begin VB.CheckBox chkFaecal 
         Caption         =   "S/S Screen"
         Height          =   195
         Index           =   10
         Left            =   1890
         TabIndex        =   3
         Top             =   1440
         Width           =   1245
      End
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
      Left            =   1710
      MaxLength       =   12
      TabIndex        =   0
      Top             =   180
      Width           =   1545
   End
   Begin VB.Label Label42 
      AutoSize        =   -1  'True
      Caption         =   "Sample ID"
      Height          =   195
      Left            =   840
      TabIndex        =   1
      Top             =   210
      Width           =   735
   End
End
Attribute VB_Name = "frmMicroOrderFaeces"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mFaecalOrders As Long

Private Sub chkFaecal_Click(Index As Integer)
      Dim n As Integer

10    On Error GoTo chkFaecal_Click_Error

20    cmdSave.Enabled = True


30    If Index = 0 Then
40        chkFaecal(8).Value = chkFaecal(Index).Value
50        chkFaecal(10).Value = chkFaecal(Index).Value
60    ElseIf Index = 12 Then
70      For n = 0 To 11
80        chkFaecal(n).Value = 0
90      Next
100   End If

110   Exit Sub

chkFaecal_Click_Error:

      Dim strES As String
      Dim intEL As Integer

120   Screen.MousePointer = 0

130   intEL = Erl
140   strES = Err.Description
150   LogError "frmMicroOrderFaeces", "chkFaecal_Click", intEL, strES


End Sub

Private Sub cmdcancel_Click()

10    On Error GoTo cmdCancel_Click_Error

20    If cmdSave.Enabled Then
30      If iMsg("Cancel without Saving?", vbQuestion = vbYesNo) = vbNo Then
40        Exit Sub
50      End If
60    End If

70    Unload Me

80    Exit Sub

cmdCancel_Click_Error:

      Dim strES As String
      Dim intEL As Integer

90    Screen.MousePointer = 0

100   intEL = Erl
110   strES = Err.Description
120   LogError "frmMicroOrderFaeces", "cmdCancel_Click", intEL, strES


End Sub


Private Sub cmdsave_Click()

10    On Error GoTo cmdsave_Click_Error

20    SaveDetails

30    Unload Me

40    Exit Sub

cmdsave_Click_Error:

      Dim strES As String
      Dim intEL As Integer

50    Screen.MousePointer = 0

60    intEL = Erl
70    strES = Err.Description
80    LogError "frmMicroOrderFaeces", "cmdsave_Click", intEL, strES


End Sub


Private Sub Form_Activate()

10    On Error GoTo Form_Activate_Error

20    LoadDetails

30    Exit Sub

Form_Activate_Error:

      Dim strES As String
      Dim intEL As Integer

40    Screen.MousePointer = 0

50    intEL = Erl
60    strES = Err.Description
70    LogError "frmMicroOrderFaeces", "Form_Activate", intEL, strES


End Sub

Private Sub Form_Load()

10    On Error GoTo Form_Load_Error

20    If SysOptShortFaeces(0) Then
30      frFaeces.Height = 1185
40    Else
50      frFaeces.Height = 2025
60    End If

70    Exit Sub

Form_Load_Error:

      Dim strES As String
      Dim intEL As Integer

80    Screen.MousePointer = 0

90    intEL = Erl
100   strES = Err.Description
110   LogError "frmMicroOrderFaeces", "Form_Load", intEL, strES


End Sub


Private Sub LoadDetails()

      Dim tb As Recordset
      Dim sql As String
      Dim SampleIDWithOffset As Long
      Dim n As Long
      Dim lngDaysOld As Long

10    On Error GoTo LoadDetails_Error

20    SampleIDWithOffset = Val(txtSampleID) + SysOptMicroOffset(0)

30    For n = 0 To 12
40      chkFaecal(n) = 0
50    Next

60    sql = "SELECT Faecal " & _
            "from MicroRequests WHERE " & _
            "SampleID = '" & SampleIDWithOffset & "' "
70    Set tb = New Recordset
80    RecOpenServer 0, tb, sql
90    If Not tb.EOF Then

100     For n = 0 To 12
110       If tb!Faecal And 2 ^ n Then
120         chkFaecal(n) = 1
130          If n = 0 Then
140           chkFaecal(8).Value = 0
150           chkFaecal(10).Value = 0
160          End If
170       End If
180     Next
  
  
190   cmdSave.Enabled = False
200   Else
210     If frmEditMicrobiology.txtDoB <> "" And frmEditMicrobiology.dtSampleDate <> "" Then
220       lngDaysOld = DateDiff("d", frmEditMicrobiology.txtDoB, frmEditMicrobiology.dtSampleDate)
    
'220       If Left(CalcAge(frmEditMicrobiology.txtDoB, frmEditMicrobiology.dtSampleDate), Len(CalcAge(frmEditMicrobiology.txtDoB)) - 2) < 5 Then
230       If lngDaysOld < (5 * 365) Then
240         chkFaecal(0).Value = 1
250         chkFaecal(6).Value = 1
260         chkFaecal(9).Value = 1
270         chkFaecal(10).Value = 1
280         chkFaecal(11).Value = 1
290       End If
300     End If
310   End If

320   Exit Sub

LoadDetails_Error:

      Dim strES As String
      Dim intEL As Integer

330   Screen.MousePointer = 0

340   intEL = Erl
350   strES = Err.Description
360   LogError "frmMicroOrderFaeces", "LoadDetails", intEL, strES, sql


End Sub

Private Sub SaveDetails()

      Dim lngF As Long
      Dim n As Long
      Dim tb As Recordset
      Dim sql As String
      Dim SampleIDWithOffset As Long

10    On Error GoTo SaveDetails_Error

20    lngF = 0
30    For n = 0 To 12
40      If chkFaecal(n) Then
50        Debug.Print 2 ^ n
60        lngF = lngF + 2 ^ n
70      End If
80    Next

90    If lngF = 0 Then
100     iMsg "Nothing to Save!", vbExclamation
110     cmdSave.Enabled = False
120     Exit Sub
130   End If

140   SampleIDWithOffset = Val(txtSampleID) + SysOptMicroOffset(0)

150   sql = "SELECT * from MicroRequests WHERE " & _
            "SampleID = '" & SampleIDWithOffset & "'"
160   Set tb = New Recordset
170   RecOpenServer 0, tb, sql
180   If tb.EOF Then
190     tb.AddNew
200   End If
210   tb!SampleID = SampleIDWithOffset
220   tb!RequestDate = Format(Now, "dd/mmm/yyyy hh:mm")
230   tb!Faecal = lngF
240   tb!Urine = 0
250   tb!Valid = 0
260   tb.Update
  

270   sql = "SELECT * from faeces WHERE sampleid = '" & SampleIDWithOffset & "'"
280   Set tb = New Recordset
290   RecOpenServer 0, tb, sql
300   If tb.EOF Then
310     tb.AddNew
320     tb!SampleID = SampleIDWithOffset
330     tb.Update
340   End If

350   mFaecalOrders = lngF

360   cmdSave.Enabled = False

370   Exit Sub

SaveDetails_Error:

      Dim strES As String
      Dim intEL As Integer

380   Screen.MousePointer = 0

390   intEL = Erl
400   strES = Err.Description
410   LogError "frmMicroOrderFaeces", "SaveDetails", intEL, strES, sql


End Sub




Private Sub txtSampleID_LostFocus()

10    On Error GoTo txtSampleID_LostFocus_Error

20    LoadDetails

30    Exit Sub

txtSampleID_LostFocus_Error:

      Dim strES As String
      Dim intEL As Integer

40    Screen.MousePointer = 0

50    intEL = Erl
60    strES = Err.Description
70    LogError "frmMicroOrderFaeces", "txtSampleID_LostFocus", intEL, strES


End Sub


Public Property Get FaecalOrders() As Long

10    FaecalOrders = mFaecalOrders

End Property
