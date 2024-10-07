VERSION 5.00
Begin VB.Form frmBioChart 
   Caption         =   "NetAcquire - Control Chart for"
   ClientHeight    =   1455
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5085
   Icon            =   "frmBioChart.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   5085
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Exit"
      Height          =   600
      Left            =   3465
      Picture         =   "frmBioChart.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   810
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   645
      Left            =   3480
      Picture         =   "frmBioChart.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   45
      Width           =   1215
   End
   Begin VB.TextBox txtName 
      Height          =   345
      Left            =   720
      MaxLength       =   45
      TabIndex        =   4
      Top             =   360
      Width           =   2415
   End
   Begin VB.TextBox txtChart 
      Height          =   315
      Left            =   1200
      TabIndex        =   2
      Top             =   390
      Width           =   1245
   End
   Begin VB.Label Label3 
      Caption         =   "Name"
      Height          =   225
      Left            =   240
      TabIndex        =   3
      Top             =   900
      Width           =   495
   End
   Begin VB.Label lblType 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "QCB"
      Height          =   285
      Left            =   780
      TabIndex        =   1
      Top             =   390
      Width           =   405
   End
   Begin VB.Label Label1 
      Caption         =   "Chart"
      Height          =   285
      Left            =   270
      TabIndex        =   0
      Top             =   390
      Width           =   525
   End
End
Attribute VB_Name = "frmBioChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdCancel_Click()

10        Unload Me

End Sub

Private Sub cmdSave_Click()
          Dim sql As String
          Dim tb As New Recordset

10        On Error GoTo cmdSave_Click_Error

20        If txtName = "" Then
30            iMsg "Chart must have name"

40            Exit Sub
50        End If

60        If txtChart = "" Then
70            iMsg "You must have Chart"
80            Exit Sub
90        End If

100       sql = "SELECT * from patientifs WHERE " & _
              "Chart = '" & lblType & txtChart & "'"
110       Set tb = New Recordset
120       RecOpenServer 0, tb, sql

130       If Not tb.EOF Then
140           iMsg "Chart already Exists!"
150           Exit Sub
160       Else
170           tb.AddNew
180           tb!Chart = lblType & txtChart
190           tb!PatName = txtName
200           tb!Dob = "01/Jan/1900"
210           tb!Address0 = "Laboratory @ " & initial2upper(HospName(0))
220           tb!sex = "M"
230           tb!Ward = "Laboratory"
240           tb.Update
250       End If

260       txtChart = ""
270       txtName = ""

280       Exit Sub

cmdSave_Click_Error:

          Dim strES As String
          Dim intEL As Integer

290       intEL = Erl
300       strES = Err.Description
310       LogError "frmBioChart", "cmdsave_Click", intEL, strES, sql


End Sub


