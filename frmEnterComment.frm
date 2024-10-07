VERSION 5.00
Begin VB.Form frmEnterComment 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire"
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6480
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   6480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   750
      Left            =   4920
      Picture         =   "frmEnterComment.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1320
      Width           =   1245
   End
   Begin VB.TextBox txtComment 
      Height          =   1575
      Left            =   270
      MaxLength       =   400
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   570
      Width           =   4485
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Comment"
      Height          =   195
      Left            =   270
      TabIndex        =   1
      Top             =   330
      Width           =   660
   End
End
Attribute VB_Name = "frmEnterComment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSave_Click()

10        Me.Hide

End Sub


