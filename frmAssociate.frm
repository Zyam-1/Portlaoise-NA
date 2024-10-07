VERSION 5.00
Begin VB.Form frmAssociate 
   Caption         =   "NetAcquire - Associated Sample ID's"
   ClientHeight    =   5190
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3810
   LinkTopic       =   "Form1"
   ScaleHeight     =   5190
   ScaleWidth      =   3810
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&save"
      Height          =   525
      Left            =   2310
      Picture         =   "frmAssociate.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4230
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   795
      Left            =   240
      TabIndex        =   4
      Top             =   1530
      Width           =   3285
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   525
         Left            =   2070
         Picture         =   "frmAssociate.frx":0A02
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   180
         Width           =   975
      End
      Begin VB.TextBox txtAdd 
         Height          =   285
         Left            =   210
         TabIndex        =   5
         Top             =   300
         Width           =   1545
      End
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove"
      Height          =   525
      Left            =   2310
      Picture         =   "frmAssociate.frx":1404
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2520
      Width           =   975
   End
   Begin VB.ListBox lstAss 
      Height          =   2205
      IntegralHeight  =   0   'False
      Left            =   450
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   2520
      Width           =   1545
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Name"
      Height          =   195
      Left            =   240
      TabIndex        =   11
      Top             =   120
      Width           =   420
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Chart"
      Height          =   195
      Left            =   240
      TabIndex        =   10
      Top             =   720
      Width           =   375
   End
   Begin VB.Label lblName 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   210
      TabIndex        =   9
      Top             =   330
      Width           =   3315
   End
   Begin VB.Label lblChart 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   210
      TabIndex        =   8
      Top             =   930
      Width           =   1515
   End
   Begin VB.Label lblSID 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   2010
      TabIndex        =   1
      Top             =   930
      Width           =   1515
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Sample ID"
      Height          =   195
      Left            =   2040
      TabIndex        =   0
      Top             =   720
      Width           =   735
   End
End
Attribute VB_Name = "frmAssociate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub GetAss()

          Dim sql As String
          Dim tb As Recordset

10        lstAss.Clear

20        sql = "SELECT AssID FROM AssociatedIDs WHERE " & _
                "SampleID = '" & Val(lblSID) + SysOptMicroOffset(0) & "' " & _
                "OR AssID = '" & Val(lblSID) + SysOptMicroOffset(0) & "'"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        Do While Not tb.EOF
60            lstAss.AddItem tb!AssID - SysOptMicroOffset(0)
70            tb.MoveNext
80        Loop

90        If lstAss.ListCount = 0 Then    'AssID unknown - look in previous for possibility
100           sql = "SELECT SampleID FROM Demographics WHERE " & _
                    "SampleID = '" & Val(lblSID) + SysOptMicroOffset(0) - 1 & "' " & _
                    "AND Chart = '" & lblChart & "' " & _
                    "AND PatName = '" & AddTicks(lblName) & "'"
110           Set tb = New Recordset
120           RecOpenServer 0, tb, sql
130           If tb.EOF Then    'not in previous
140               txtAdd = ""
150           Else
160               txtAdd = CStr(Val(tb!SampleID) - SysOptMicroOffset(0))
170           End If
180       Else    'AssID already known
190           GetAssRecurse
200       End If

End Sub

Private Sub GetAssRecurse()

          Dim sql As String
          Dim tb As Recordset
          Dim n As Integer
          Dim Y As Integer
          Dim FoundAss As Boolean
          Dim FoundSID As Boolean

10        For n = lstAss.ListCount - 1 To 0 Step -1

20            sql = "SELECT SampleID, AssID FROM AssociatedIDs WHERE " & _
                    "SampleID = '" & Val(lstAss.List(n)) + SysOptMicroOffset(0) & "' " & _
                    "OR AssID  = '" & Val(lstAss.List(n)) + SysOptMicroOffset(0) & "'"
30            Set tb = New Recordset
40            RecOpenServer 0, tb, sql
50            Do While Not tb.EOF
60                FoundAss = False
70                FoundSID = False
80                For Y = lstAss.ListCount - 1 To 0 Step -1
90                    If lstAss.List(Y) = tb!SampleID - SysOptMicroOffset(0) Then FoundSID = True
100                   If lstAss.List(Y) = tb!AssID - SysOptMicroOffset(0) Then FoundAss = True
110                   If FoundAss And FoundSID Then Exit For
120               Next
130               If Not FoundAss Then lstAss.AddItem tb!AssID - SysOptMicroOffset(0)
140               If Not FoundSID Then lstAss.AddItem tb!SampleID - SysOptMicroOffset(0)
150               If Not FoundAss Or Not FoundSID Then GetAssRecurse
160               tb.MoveNext
170           Loop

180       Next

End Sub

Private Sub cmdadd_Click()

          Dim tb As Recordset
          Dim sql As String

10        If Trim$(txtAdd) = "" Then Exit Sub

20        sql = "SELECT SampleID, AssID FROM AssociatedIDs WHERE " & _
                "( SampleID = '" & Val(lblSID) + SysOptMicroOffset(0) & "' " & _
                "  AND AssID = '" & Val(txtAdd) + SysOptMicroOffset(0) & "' ) " & _
                "OR " & _
                "( SampleID = '" & Val(txtAdd) + SysOptMicroOffset(0) & "' " & _
                "  AND AssID = '" & Val(lblSID) + SysOptMicroOffset(0) & "' ) "
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        If tb.EOF Then
60            tb.AddNew
70            tb!SampleID = Val(lblSID) + SysOptMicroOffset(0)
80            tb!AssID = Val(txtAdd) + SysOptMicroOffset(0)
90            tb.Update
100       End If

110       lstAss.AddItem txtAdd
120       txtAdd = ""
130       GetAssRecurse

End Sub

Private Sub cmdCancel_Click()

10        Unload Me

End Sub

Private Sub cmdRemove_Click()

          Dim n As Integer
          Dim sql As String

10        If lstAss.SelCount = 0 Then Exit Sub

20        For n = lstAss.ListCount - 1 To 0 Step -1
30            If lstAss.Selected(n) Then

40                sql = "DELETE FROM AssociatedIDs WHERE " & _
                        "( SampleID = '" & Val(lstAss.List(n)) + SysOptMicroOffset(0) & "' " & _
                        "  AND AssID = '" & Val(lblSID) + SysOptMicroOffset(0) & "' ) " & _
                        "OR " & _
                        "( AssID = '" & Val(lstAss.List(n)) + SysOptMicroOffset(0) & "' " & _
                        "  AND SampleID = '" & Val(lblSID) + SysOptMicroOffset(0) & "' ) "
50                Cnxn(0).Execute sql

60                lstAss.RemoveItem n

70            End If
80        Next

End Sub


Private Sub Form_Activate()

10        GetAss

End Sub

Private Sub txtAdd_KeyPress(KeyAscii As Integer)

10        KeyAscii = VI(KeyAscii, Numeric_Only)

End Sub


