VERSION 5.00
Begin VB.Form fcdrDBox 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire"
   ClientHeight    =   2700
   ClientLeft      =   3600
   ClientTop       =   2535
   ClientWidth     =   5910
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox optOptions 
      Height          =   315
      Left            =   480
      TabIndex        =   3
      Text            =   "optOptions"
      Top             =   2040
      Width           =   3405
   End
   Begin VB.CommandButton bCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   705
      Left            =   4320
      Picture         =   "fcdrDBox.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1200
      Width           =   1245
   End
   Begin VB.CommandButton bOK 
      Caption         =   "O. K."
      Default         =   -1  'True
      Height          =   645
      Left            =   4320
      Picture         =   "fcdrDBox.frx":066A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   210
      Width           =   1245
   End
   Begin VB.Label lPrompt 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   1695
      Left            =   480
      TabIndex        =   2
      Top             =   210
      Width           =   3375
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "fcdrDBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private RetVal As String

Private pListOrCombo
Private Sub bcancel_Click()

10        On Error GoTo bCancel_Click_Error

20        RetVal = ""
30        Me.Hide

40        Exit Sub

bCancel_Click_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "fcdrDBox", "bcancel_Click", intEL, strES


End Sub

Public Property Get ReturnValue() As String

10        On Error GoTo ReturnValue_Error

20        ReturnValue = RetVal

30        Exit Property

ReturnValue_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "fcdrDBox", "ReturnValue", intEL, strES


End Property

Private Sub bOK_Click()

10        On Error GoTo bOK_Click_Error

20        RetVal = Trim$(optOptions)
30        Me.Hide

40        Exit Sub

bOK_Click_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "fcdrDBox", "bOK_Click", intEL, strES


End Sub


Function dBOX(ByVal Prompt As String)

10        On Error GoTo dBOX_Error

20        lPrompt = Prompt

30        Exit Function

dBOX_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "fcdrDBox", "dBOX", intEL, strES

End Function




Public Property Let Options(ByRef varOptions As Variant)

          Dim n As Integer

10        On Error GoTo Options_Error

20        optOptions.Clear
30        For n = 0 To UBound(varOptions)
40            optOptions.AddItem varOptions(n)
50        Next

60        Exit Property

Options_Error:

          Dim strES As String
          Dim intEL As Integer

70        intEL = Erl
80        strES = Err.Description
90        LogError "fcdrDBox", "Options", intEL, strES


End Property

Public Property Let Prompt(ByVal strPrompt As String)

10        lPrompt = strPrompt

End Property

Public Property Let ListOrCombo(ByVal strListOrCombo As String)

10        pListOrCombo = strListOrCombo

End Property

Private Sub optOptions_KeyPress(KeyAscii As Integer)

10        On Error GoTo optOptions_KeyPress_Error

20        If pListOrCombo = "List" Then
30            KeyAscii = 0
40        End If

50        Exit Sub

optOptions_KeyPress_Error:

          Dim strES As String
          Dim intEL As Integer

60        intEL = Erl
70        strES = Err.Description
80        LogError "fcdrDBox", "optOptions_KeyPress", intEL, strES


End Sub


