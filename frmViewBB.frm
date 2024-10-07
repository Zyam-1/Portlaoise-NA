VERSION 5.00
Begin VB.Form frmViewBB 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Blood Bank View"
   ClientHeight    =   4080
   ClientLeft      =   1710
   ClientTop       =   1395
   ClientWidth     =   6210
   Icon            =   "frmViewBB.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   6210
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton bCancel 
      Caption         =   "Exit"
      Height          =   840
      Left            =   4140
      Picture         =   "frmViewBB.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1035
      Width           =   1290
   End
   Begin VB.Label lGroup 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1860
      TabIndex        =   18
      Top             =   900
      Width           =   1770
   End
   Begin VB.Label lAnti3Reported 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1860
      TabIndex        =   17
      Top             =   1500
      Width           =   1770
   End
   Begin VB.Label lAIDr 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1860
      TabIndex        =   16
      Top             =   1890
      Width           =   1770
   End
   Begin VB.Label lProcedure 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1860
      TabIndex        =   15
      Top             =   2280
      Width           =   3855
   End
   Begin VB.Label lConditions 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1860
      TabIndex        =   14
      Top             =   2700
      Width           =   3855
   End
   Begin VB.Label lComment 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1860
      TabIndex        =   13
      Top             =   3120
      Width           =   3855
   End
   Begin VB.Label lSampleComment 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1860
      TabIndex        =   12
      Top             =   3540
      Width           =   3855
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Antibody ID"
      Height          =   195
      Left            =   900
      TabIndex        =   11
      Top             =   1920
      Width           =   825
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Antibodies Reported"
      Height          =   195
      Left            =   285
      TabIndex        =   10
      Top             =   1530
      Width           =   1440
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Conditions"
      Height          =   195
      Left            =   990
      TabIndex        =   9
      Top             =   2730
      Width           =   735
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Procedure"
      Height          =   195
      Left            =   990
      TabIndex        =   8
      Top             =   2310
      Width           =   735
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Sample Comment"
      Height          =   195
      Left            =   510
      TabIndex        =   7
      Top             =   3570
      Width           =   1230
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Comment"
      Height          =   195
      Left            =   1065
      TabIndex        =   6
      Top             =   3150
      Width           =   660
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Group"
      Height          =   195
      Left            =   1290
      TabIndex        =   5
      Top             =   930
      Width           =   435
   End
   Begin VB.Label lName 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1860
      TabIndex        =   4
      Top             =   510
      Width           =   3855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Name"
      Height          =   195
      Left            =   1410
      TabIndex        =   3
      Top             =   540
      Width           =   420
   End
   Begin VB.Label lChart 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1830
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Chart"
      Height          =   195
      Left            =   1440
      TabIndex        =   0
      Top             =   150
      Width           =   375
   End
End
Attribute VB_Name = "frmViewBB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub bcancel_Click()

10        Unload Me

End Sub


Private Sub Form_Activate()

          Dim sql As String
          Dim tb As New Recordset

10        On Error GoTo Form_Activate_Error

20        sql = "SELECT * from PatientDetails WHERE " & _
                "PatNum = '" & lChart & "'"
30        Set tb = New Recordset
40        RecOpenClientBB tb, sql
50        If tb.EOF Then
60            lName = "No Record in Blood Bank"
70        Else
80            lName = tb!Name & ""
90            lProcedure = tb!Procedure & ""
100           lConditions = tb!conditions & ""
110           lGroup = tb!fgroup & ""
120           lAnti3Reported = tb!anti3reported & ""
130           lcomment = tb!Comment & ""
140           lAIDr = tb!aidr & ""
              'lSampleComment = tb!samplecomment & ""
150       End If

160       Exit Sub

Form_Activate_Error:

          Dim strES As String
          Dim intEL As Integer



170       intEL = Erl
180       strES = Err.Description
190       LogError "frmViewBB", "Form_Activate", intEL, strES, sql


End Sub

