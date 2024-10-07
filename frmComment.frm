VERSION 5.00
Begin VB.Form frmComment 
   Caption         =   "NetAcquire - Comments"
   ClientHeight    =   3435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmComment.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Exit"
      Height          =   825
      Left            =   3105
      Picture         =   "frmComment.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2565
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   825
      Left            =   405
      Picture         =   "frmComment.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2565
      Width           =   1275
   End
   Begin VB.TextBox txtComment 
      Height          =   2130
      Left            =   135
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   315
      Width           =   4425
   End
End
Attribute VB_Name = "frmComment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mDiscipline As String

Private Sub cmdCancel_Click()

10        Unload Me

End Sub

Private Sub cmdSave_Click()

10        Me.Hide

End Sub

Public Property Let Discipline(ByVal Discipline As String)

10        On Error GoTo Discipline_Error

20        mDiscipline = Discipline

30        Exit Property

Discipline_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmComment", "Discipline", intEL, strES


End Property



Private Sub txtComment_KeyDown(KeyCode As Integer, Shift As Integer)

          Dim n As Long
          Dim s As String
          Dim sql As String
          Dim tb As Recordset

10        On Error GoTo txtComment_KeyDown_Error

20        If KeyCode = vbKeyF2 Then

30            If Len(Trim$(txtComment)) < 2 Then Exit Sub

40            n = txtComment.SelStart
50            If n < 3 Then Exit Sub

60            s = UCase(Mid(txtComment, n - 2, 3))

70            If mDiscipline = "IMM" Then
80                If ListText("IM", s) <> "" Then
90                    s = ListText("IM", s)
100               End If
110           End If

120           txtComment = Left(txtComment, n - 3)
130           txtComment = txtComment & s
140           txtComment.SelStart = Len(txtComment)

150       ElseIf KeyCode = vbKeyF3 Then

160           If mDiscipline = "IMM" Then
170               sql = "SELECT * from lists WHERE listtype = 'IM'"
180               Set tb = New Recordset
190               RecOpenServer 0, tb, sql
200               Do While Not tb.EOF
210                   s = Trim(tb!Text)
220                   frmMessages.lstComm.AddItem s
230                   tb.MoveNext
240               Loop

250               Set frmMessages.f = Me
260               Set frmMessages.T = txtComment
270               frmMessages.Show 1

280           End If
290       End If

300       Exit Sub

txtComment_KeyDown_Error:

          Dim strES As String
          Dim intEL As Integer

310       intEL = Erl
320       strES = Err.Description
330       LogError "frmComment", "txtComment_KeyDown", intEL, strES

End Sub


Public Property Get Comment() As String

10        Comment = txtComment.Text

End Property

Public Property Let Comment(ByVal strNewValue As String)

10        txtComment.Text = strNewValue

End Property
