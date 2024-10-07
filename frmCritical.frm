VERSION 5.00
Begin VB.Form frmCritical 
   Caption         =   "                                                        Critical Problem"
   ClientHeight    =   4380
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6375
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4380
   ScaleWidth      =   6375
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRestart 
      Caption         =   "Restart"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   2430
      Picture         =   "frmCritical.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3330
      Width           =   1860
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   $"frmCritical.frx":0442
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   450
      TabIndex        =   2
      Top             =   2250
      Width           =   5100
   End
   Begin VB.Label Label1 
      Caption         =   "NetAcquire Has Encountered a Critical Problem"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   135
      TabIndex        =   1
      Top             =   1575
      Width           =   6135
   End
   Begin VB.Image Image1 
      Height          =   1425
      Left            =   2295
      Picture         =   "frmCritical.frx":04E9
      Stretch         =   -1  'True
      Top             =   90
      Width           =   2100
   End
End
Attribute VB_Name = "frmCritical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Const EWX_SHUTDOWN = 1
Const EWX_REBOOT = 2
'Const EWX_LOGOFF = 0
'Const EWX_FORCE = 4



Private Declare Function ExitWindowsEx Lib "user32" _
    (ByVal uFlags As Long, _
    ByVal dwReserved As Long) _
    As Long
  
Private Sub cmdRestart_Click()
          'Restart Windows (works on Windows 95/NT)
10        ExitWindowsEx EWX_REBOOT, 0
End Sub



