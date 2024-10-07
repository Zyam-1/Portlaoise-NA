VERSION 5.00
Begin VB.Form frmMessages 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Comments"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8640
   ClipControls    =   0   'False
   Icon            =   "frmMessages.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   8640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstComm 
      Height          =   4155
      Left            =   90
      TabIndex        =   0
      Top             =   240
      Width           =   8445
   End
End
Attribute VB_Name = "frmMessages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public f As Form
Public T As TextBox


Private Sub lstComm_Click()

10        On Error GoTo lstComm_Click_Error

20        T = T & lstComm

30        Unload Me

40        Exit Sub

lstComm_Click_Error:

          Dim strES As String
          Dim intEL As Integer



50        intEL = Erl
60        strES = Err.Description
70        LogError "frmMessages", "lstComm_Click", intEL, strES


End Sub
