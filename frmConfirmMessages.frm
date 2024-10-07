VERSION 5.00
Begin VB.Form frmConfirmMessages 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   5460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   1065
      Left            =   3900
      Picture         =   "frmConfirmMessages.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1650
      Width           =   1035
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   1065
      Left            =   2490
      Picture         =   "frmConfirmMessages.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1650
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CheckBox chkClinician 
      Alignment       =   1  'Right Justify
      Caption         =   "Clinician"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   930
      TabIndex        =   5
      Top             =   2460
      Width           =   1065
   End
   Begin VB.CheckBox chkWard 
      Alignment       =   1  'Right Justify
      Caption         =   "Ward"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1200
      TabIndex        =   4
      Top             =   2160
      Width           =   795
   End
   Begin VB.CheckBox chkDoB 
      Alignment       =   1  'Right Justify
      Caption         =   "Date of Birth"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   570
      TabIndex        =   3
      Top             =   1560
      Width           =   1425
   End
   Begin VB.CheckBox chkChart 
      Alignment       =   1  'Right Justify
      Caption         =   "Chart"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1200
      TabIndex        =   2
      Top             =   1860
      Width           =   795
   End
   Begin VB.CheckBox chkName 
      Alignment       =   1  'Right Justify
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1170
      TabIndex        =   1
      Top             =   1260
      Value           =   1  'Checked
      Width           =   825
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ask for confirmation when the following details of Patients Demographics are changed"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   750
      TabIndex        =   0
      Top             =   450
      Width           =   4185
   End
End
Attribute VB_Name = "frmConfirmMessages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkChart_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        cmdSave.Visible = True

End Sub

Private Sub chkClinician_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        cmdSave.Visible = True

End Sub

Private Sub chkDoB_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        cmdSave.Visible = True

End Sub

Private Sub chkName_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        cmdSave.Visible = True

End Sub


Private Sub chkWard_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        cmdSave.Visible = True

End Sub

Private Sub cmdCancel_Click()

10        If cmdSave.Visible Then
20            If iMsg("Cancel without Saving?", vbQuestion + vbYesNo) = vbNo Then
30                Exit Sub
40            End If
50        End If

60        Unload Me

End Sub


Private Sub cmdSave_Click()

10        SaveOptionSetting "MicroConfirmChangeName", chkName.Value = 1
20        SaveOptionSetting "MicroConfirmChangeDoB", chkDoB.Value = 1
30        SaveOptionSetting "MicroConfirmChangeChart", chkChart.Value = 1
40        SaveOptionSetting "MicroConfirmChangeWard", chkWard.Value = 1
50        SaveOptionSetting "MicroConfirmChangeClinician", chkClinician.Value = 1

60        Unload Me

End Sub


Private Sub Form_Load()

10        chkName.Value = IIf(GetOptionSetting("MicroConfirmChangeName", "True") = "True", 1, 0)
20        chkDoB.Value = IIf(GetOptionSetting("MicroConfirmChangeDoB", "True") = "True", 1, 0)
30        chkChart.Value = IIf(GetOptionSetting("MicroConfirmChangeChart", "True") = "True", 1, 0)
40        chkWard.Value = IIf(GetOptionSetting("MicroConfirmChangeWard", "True") = "True", 1, 0)
50        chkClinician.Value = IIf(GetOptionSetting("MicroConfirmChangeClinician", "True") = "True", 1, 0)

End Sub

