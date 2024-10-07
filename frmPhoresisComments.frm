VERSION 5.00
Begin VB.Form frmPhoresisComments 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NetAcquire - Phoresis Comments"
   ClientHeight    =   6465
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8025
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   8025
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtImmunologyD 
      Height          =   1155
      Left            =   180
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   5130
      Width           =   6615
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   825
      Left            =   6960
      Picture         =   "frmPhoresisComments.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Exit Screen"
      Top             =   5460
      Width           =   885
   End
   Begin VB.TextBox txtImmunologyC 
      Height          =   1155
      Left            =   180
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   3510
      Width           =   6615
   End
   Begin VB.TextBox txtImmunologyB 
      Height          =   1155
      Left            =   180
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1935
      Width           =   6615
   End
   Begin VB.TextBox txtImmunologyA 
      Height          =   1155
      Left            =   180
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   360
      Width           =   6615
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Serum Electrophoresis Comments"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   180
      TabIndex        =   8
      Top             =   4860
      Width           =   3045
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Electrophoresis Comments"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   180
      TabIndex        =   6
      Top             =   3240
      Width           =   2415
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Serum Immunofixation Comments"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   180
      TabIndex        =   5
      Top             =   1680
      Width           =   2955
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Serum Immunotyping Comments"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   180
      TabIndex        =   4
      Top             =   120
      Width           =   2880
   End
End
Attribute VB_Name = "frmPhoresisComments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public SampleID As String

Private Sub cmdCancel_Click()

10        SampleID = ""
20        Unload Me
End Sub

Private Sub Form_Load()

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo Form_Load_Error

20        sql = "Select Code, Result From ImmResults Where SampleID = '" & SampleID & "' And " & _
                "Code In ('FIX','ECOM','IT','ELEC')"
30        Set tb = New Recordset
40        RecOpenClient 0, tb, sql

50        If Not tb.EOF Then
60            While Not tb.EOF
70                If tb!Code = "IT" Then
                      'serum immunotyping
80                    txtImmunologyA = tb!Result & ""
90                ElseIf tb!Code = "FIX" Then
                      'serum immunofixation
100                   txtImmunologyB = tb!Result & ""
110               ElseIf tb!Code = "ECOM" Then
                      'electrophoresis (extended)
120                   txtImmunologyC = tb!Result & ""
130               ElseIf tb!Code = "ELEC" Then
                      'serum electrophoresis
140                   txtImmunologyD = tb!Result & ""
150               End If
160               tb.MoveNext
170           Wend
180       End If

190       Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

200       intEL = Erl
210       strES = Err.Description
220       LogError "frmPhoresisComments", "Form_Load", intEL, strES, sql

End Sub
