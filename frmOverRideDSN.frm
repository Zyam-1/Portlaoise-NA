VERSION 5.00
Begin VB.Form frmOverRideDSN 
   Caption         =   "NetAcquire - DSN OverRide"
   ClientHeight    =   3150
   ClientLeft      =   7110
   ClientTop       =   2730
   ClientWidth     =   5220
   Icon            =   "frmOverRideDSN.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3150
   ScaleWidth      =   5220
   Begin VB.CommandButton cmdContinue 
      Caption         =   "Continue"
      Height          =   885
      Left            =   1755
      Picture         =   "frmOverRideDSN.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2115
      Width           =   1560
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select DSN"
      Height          =   1455
      Left            =   300
      TabIndex        =   0
      Top             =   450
      Width           =   4545
      Begin VB.OptionButton optDSN 
         Caption         =   "Not Set"
         Enabled         =   0   'False
         Height          =   225
         Index           =   2
         Left            =   210
         TabIndex        =   3
         Top             =   990
         Width           =   3465
      End
      Begin VB.OptionButton optDSN 
         Caption         =   "Not Set"
         Enabled         =   0   'False
         Height          =   225
         Index           =   1
         Left            =   210
         TabIndex        =   2
         Top             =   660
         Width           =   3465
      End
      Begin VB.OptionButton optDSN 
         Caption         =   "Not Set"
         Enabled         =   0   'False
         Height          =   225
         Index           =   0
         Left            =   210
         TabIndex        =   1
         Top             =   330
         Value           =   -1  'True
         Width           =   3465
      End
   End
End
Attribute VB_Name = "frmOverRideDSN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ptb As Recordset

Private pDSN As String

Private Sub cmdContinue_Click()

10    On Error GoTo cmdContinue_Click_Error

20    If optDSN(0) Then
30      pDSN = "DSN"
40    ElseIf optDSN(1) Then
50      pDSN = "Live69DSN"
60    ElseIf optDSN(2) Then
70      pDSN = "Test69DSN"
80      TestSys = True
90    End If

100   Unload Me

110   Exit Sub

cmdContinue_Click_Error:

      Dim strES As String
      Dim intEL As Integer

120   Screen.MousePointer = 0

130   intEL = Erl
140   strES = Err.Description
150   LogError "frmOverRideDSN", "cmdContinue_Click", intEL, strES


End Sub

Private Sub Form_Activate()

      Dim n As Long

10    On Error GoTo Form_Activate_Error

20    n = 0

30    If Trim$(ptb!DSN & "") <> "" Then
40      optDSN(0).Caption = Trim$(ptb!DSN)
50      optDSN(0).Enabled = True
60      n = 1
70    End If
80    If Trim$(ptb!Live69DSN & "") <> "" Then
90      optDSN(1).Caption = Trim$(ptb!Live69DSN)
100     optDSN(1).Enabled = True
110     n = n + 1
120   End If
130   If Trim$(ptb!Test69DSN & "") <> "" Then
140     optDSN(2).Caption = Trim$(ptb!Test69DSN)
150     optDSN(2).Enabled = True
160     n = n + 1
170   End If

180   If n = 1 Then
190     pDSN = "DSN"
200     Unload Me
210   End If

220   Exit Sub

Form_Activate_Error:

      Dim strES As String
      Dim intEL As Integer

230   Screen.MousePointer = 0

240   intEL = Erl
250   strES = Err.Description
260   LogError "frmOverRideDSN", "Form_Activate", intEL, strES


End Sub

Public Property Let Rec(ByVal R As Recordset)

10    On Error GoTo Rec_Error

20    Set ptb = R

30    Exit Property

Rec_Error:

      Dim strES As String
      Dim intEL As Integer

40    Screen.MousePointer = 0

50    intEL = Erl
60    strES = Err.Description
70    LogError "frmOverRideDSN", "Rec", intEL, strES


End Property

Public Property Get DSN() As String

10    On Error GoTo DSN_Error

20    DSN = pDSN

30    Exit Property

DSN_Error:

      Dim strES As String
      Dim intEL As Integer

40    Screen.MousePointer = 0

50    intEL = Erl
60    strES = Err.Description
70    LogError "frmOverRideDSN", "DSN", intEL, strES


End Property

