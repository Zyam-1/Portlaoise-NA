VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmcdrInputTime 
   Caption         =   "Form2"
   ClientHeight    =   2505
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5310
   LinkTopic       =   "Form2"
   ScaleHeight     =   2505
   ScaleWidth      =   5310
   StartUpPosition =   3  'Windows Default
   Begin MSMask.MaskEdBox txtIP 
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   1980
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "hh:mm"
      Mask            =   "##:##"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   555
      Left            =   3810
      Picture         =   "frmcdrInputTime.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1350
      Width           =   1245
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "O. K."
      Default         =   -1  'True
      Height          =   645
      Left            =   3810
      Picture         =   "frmcdrInputTime.frx":066A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   210
      Width           =   1245
   End
   Begin VB.Label lblPrompt 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   1695
      Left            =   270
      TabIndex        =   3
      Top             =   210
      Width           =   3375
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmcdrInputTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ReturnValue As String

Public Property Get RetVal() As String

10        RetVal = ReturnValue

End Property

Private Sub cmdCancel_Click()

10        ReturnValue = ""
20        Unload Me

End Sub


Private Sub cmdOK_Click()

10        ReturnValue = txtIP
20        Unload Me

End Sub


Private Sub Form_Activate()

10        txtIP.Text = "__:__"

End Sub

