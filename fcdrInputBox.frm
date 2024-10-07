VERSION 5.00
Begin VB.Form fcdrInputBox 
   BackColor       =   &H00808000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   2505
   ClientLeft      =   1845
   ClientTop       =   2190
   ClientWidth     =   5700
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   5700
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtInput 
      Height          =   255
      Left            =   660
      TabIndex        =   3
      Top             =   1950
      Width           =   3375
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "O. K."
      Default         =   -1  'True
      Height          =   525
      Left            =   4200
      TabIndex        =   1
      Top             =   180
      Width           =   1245
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   525
      Left            =   4200
      TabIndex        =   0
      Top             =   1350
      Width           =   1245
   End
   Begin VB.Label lblPrompt 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   1695
      Left            =   660
      TabIndex        =   2
      Top             =   180
      Width           =   3375
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "fcdrInputBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ReturnValue As String
Private mPass As Boolean

Private Sub cmdCancel_Click()

10        On Error GoTo cmdCancel_Click_Error

20        ReturnValue = ""
30        Unload Me

40        Exit Sub

cmdCancel_Click_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "fcdrInputBox", "cmdCancel_Click", intEL, strES

End Sub

Private Sub cmdOK_Click()

10        On Error GoTo cmdOK_Click_Error

20        ReturnValue = txtInput
30        Unload Me

40        Exit Sub

cmdOK_Click_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "fcdrInputBox", "cmdOK_Click", intEL, strES

End Sub

Private Sub Form_Activate()

10        On Error GoTo Form_Activate_Error

20        On Error GoTo Form_Activate_Error

30        If mPass Then
40            txtInput.PasswordChar = "*"
50        Else
60            txtInput.PasswordChar = ""
70        End If

80        txtInput.SelStart = 0
90        txtInput.SelLength = Len(txtInput)
100       txtInput.SetFocus

110       Exit Sub

Form_Activate_Error:

          Dim strES As String
          Dim intEL As Integer

120       intEL = Erl
130       strES = Err.Description
140       LogError "fcdrInputBox", "Form_Activate", intEL, strES




End Sub

Private Sub Form_Deactivate()

10        mPass = False

End Sub

Public Property Let PassWord(ByVal blnNewValue As Boolean)

10        On Error GoTo PassWord_Error

20        mPass = blnNewValue

30        Exit Property

PassWord_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "fcdrInputBox", "PassWord", intEL, strES

End Property

Public Property Get RetVal() As String

10        On Error GoTo Retval_Error

20        RetVal = ReturnValue

30        Exit Property

Retval_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "fcdrInputBox", "Retval", intEL, strES

End Property
