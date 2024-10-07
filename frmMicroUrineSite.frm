VERSION 5.00
Begin VB.Form frmMicroUrineSite 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire Urines"
   ClientHeight    =   2550
   ClientLeft      =   7185
   ClientTop       =   4155
   ClientWidth     =   2145
   Icon            =   "frmMicroUrineSite.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   2145
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   825
      Left            =   1170
      Picture         =   "frmMicroUrineSite.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1500
      Width           =   765
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   825
      Left            =   180
      Picture         =   "frmMicroUrineSite.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1500
      Width           =   765
   End
   Begin VB.Frame Frame2 
      Caption         =   "Urine Sample"
      Height          =   1095
      Left            =   150
      TabIndex        =   0
      Top             =   180
      Width           =   1785
      Begin VB.OptionButton optU 
         Caption         =   "SPA"
         Height          =   195
         Index           =   3
         Left            =   900
         TabIndex        =   6
         Top             =   630
         Width           =   645
      End
      Begin VB.OptionButton optU 
         Alignment       =   1  'Right Justify
         Caption         =   "BSU"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   5
         Top             =   630
         Width           =   645
      End
      Begin VB.OptionButton optU 
         Caption         =   "CSU"
         Height          =   195
         Index           =   1
         Left            =   930
         TabIndex        =   4
         Top             =   360
         Width           =   645
      End
      Begin VB.OptionButton optU 
         Alignment       =   1  'Right Justify
         Caption         =   "MSU"
         Height          =   195
         Index           =   0
         Left            =   210
         TabIndex        =   3
         Top             =   360
         Value           =   -1  'True
         Width           =   675
      End
   End
End
Attribute VB_Name = "frmMicroUrineSite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

          Dim n As Long

10        On Error GoTo cmdCancel_Click_Error

20        For n = 0 To 3
30            optU(n).Value = False
40        Next
50        Me.Hide

60        Exit Sub

cmdCancel_Click_Error:

          Dim strES As String
          Dim intEL As Integer



70        intEL = Erl
80        strES = Err.Description
90        LogError "frmMicroUrineSite", "cmdCancel_Click", intEL, strES


End Sub


Private Sub cmdSave_Click()


10        On Error GoTo cmdSave_Click_Error

20        Me.Hide

30        Exit Sub

cmdSave_Click_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmMicroUrineSite", "cmdsave_Click", intEL, strES


End Sub


Public Property Get Details() As String

          Dim n As Long

10        On Error GoTo Details_Error

20        For n = 0 To 3
30            If optU(n) = True Then
40                Details = optU(n).Caption
50            End If
60        Next

70        Exit Property

Details_Error:

          Dim strES As String
          Dim intEL As Integer



80        intEL = Erl
90        strES = Err.Description
100       LogError "frmMicroUrineSite", "Details", intEL, strES


End Property


