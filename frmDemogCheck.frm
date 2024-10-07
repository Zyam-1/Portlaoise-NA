VERSION 5.00
Begin VB.Form frmDemogCheck 
   Caption         =   "NetAcquire 6 - Demographic Conflict"
   ClientHeight    =   3585
   ClientLeft      =   1245
   ClientTop       =   1500
   ClientWidth     =   6585
   Icon            =   "frmDemogCheck.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   6585
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton bSelect 
      Caption         =   "Select"
      Height          =   915
      Index           =   2
      Left            =   4140
      Picture         =   "frmDemogCheck.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2400
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.Frame Frame1 
      Caption         =   "A && E"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Index           =   2
      Left            =   360
      TabIndex        =   9
      Top             =   1860
      Width           =   3705
      Begin VB.Label lDoB 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   2
         Left            =   780
         TabIndex        =   15
         Top             =   990
         Width           =   1635
      End
      Begin VB.Label lAddress 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   2
         Left            =   780
         TabIndex        =   14
         Top             =   630
         Width           =   2775
      End
      Begin VB.Label lName 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   2
         Left            =   780
         TabIndex        =   13
         Top             =   270
         Width           =   2775
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "D.o.B."
         Height          =   195
         Index           =   0
         Left            =   270
         TabIndex        =   12
         Top             =   1020
         Width           =   450
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Address"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   630
         Width           =   570
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         Height          =   195
         Index           =   0
         Left            =   270
         TabIndex        =   10
         Top             =   300
         Width           =   420
      End
   End
   Begin VB.CommandButton bSelect 
      Caption         =   "Select"
      Height          =   915
      Index           =   1
      Left            =   4140
      Picture         =   "frmDemogCheck.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   780
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton bNone 
      Caption         =   "Select None"
      Height          =   915
      Left            =   5340
      Picture         =   "frmDemogCheck.frx":0B8E
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Chart"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Index           =   1
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   3705
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         Height          =   195
         Index           =   1
         Left            =   270
         TabIndex        =   6
         Top             =   300
         Width           =   420
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Address"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   630
         Width           =   570
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "D.o.B."
         Height          =   195
         Index           =   1
         Left            =   270
         TabIndex        =   4
         Top             =   1020
         Width           =   450
      End
      Begin VB.Label lName 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   1
         Left            =   780
         TabIndex        =   3
         Top             =   270
         Width           =   2775
      End
      Begin VB.Label lAddress 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   1
         Left            =   780
         TabIndex        =   2
         Top             =   630
         Width           =   2775
      End
      Begin VB.Label lDoB 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   1
         Left            =   780
         TabIndex        =   1
         Top             =   990
         Width           =   1635
      End
   End
End
Attribute VB_Name = "frmDemogCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pIDType As String

Private Sub bNone_Click()

10        On Error GoTo bNone_Click_Error

20        pIDType = ""
30        Me.Hide

40        Exit Sub

bNone_Click_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmDemogCheck", "bNone_Click", intEL, strES

End Sub

Private Sub bSELECT_Click(Index As Integer)

10        On Error GoTo bSELECT_Click_Error

20        pIDType = Choose(Index, "ChART", "AandE")

30        Me.Hide

40        Exit Sub

bSELECT_Click_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmDemogCheck", "bSELECT_Click", intEL, strES

End Sub

Private Sub Form_Unload(Cancel As Integer)

10        On Error GoTo Form_Unload_Error

20        pIDType = ""

30        Exit Sub

Form_Unload_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmDemogCheck", "Form_Unload", intEL, strES

End Sub

Public Property Get IDType() As String

10        IDType = pIDType

End Property

