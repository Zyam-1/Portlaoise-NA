VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmDelta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Delta Checking"
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10560
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   10560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2505
      Left            =   270
      TabIndex        =   4
      Top             =   1560
      Width           =   8715
      _ExtentX        =   15372
      _ExtentY        =   4419
      _Version        =   393216
      Cols            =   4
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      FormatString    =   "<Parameter              |^Delta Check |^Change %  |<Interval          "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   1065
      Left            =   2850
      Picture         =   "frmDelta.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5250
      Width           =   1125
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   1065
      Left            =   4590
      Picture         =   "frmDelta.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5250
      Width           =   1125
   End
   Begin VB.Frame fraSampleType 
      Caption         =   "Sample Type"
      Height          =   855
      Left            =   4140
      TabIndex        =   0
      Top             =   270
      Width           =   2895
      Begin VB.ComboBox cmbSampleType 
         Height          =   315
         Left            =   525
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   315
         Width           =   1860
      End
   End
End
Attribute VB_Name = "frmDelta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pDiscipline As String
Private pSampleType As String

Private Sub cmbSample_Change()

End Sub

Private Sub cmbSample_Click()

End Sub


Private Sub cmdCancel_Click()

10        Unload Me

End Sub


Private Sub Form_Activate()

10        If pSampleType = "" Then
20            FillSampleType
30        Else
40            cmbSampleType.Clear
50            cmbSampleType.AddItem pSampleType
60            cmbSampleType.ListIndex = 0
70            fraSampleType.Enabled = False
80        End If

End Sub
Private Sub FillSampleType()

          Dim sql As String
          Dim tb As New Recordset

10        sql = "SELECT Text FROM Lists " & _
                "WHERE ListType = 'ST' " & _
                "ORDER BY ListOrder"
20        Set tb = New Recordset
30        RecOpenServer 0, tb, sql
40        Do While Not tb.EOF
50            cmbSampleType.AddItem Trim(tb!Text & "")
60            tb.MoveNext
70        Loop

80        cmbSampleType.ListIndex = 0

End Sub

Public Property Let Discipline(ByVal sNewValue As String)

10        pDiscipline = sNewValue

End Property
Public Property Let SampleType(ByVal sNewValue As String)

10        pSampleType = sNewValue

End Property

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

10        If cmdSave.Enabled Then
20            If iMsg("Cancel without Saving?", vbQuestion + vbYesNo) = vbNo Then
30                Cancel = True
40            End If
50        End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

10        pSampleType = ""
20        pDiscipline = ""

End Sub


