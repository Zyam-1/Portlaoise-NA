VERSION 5.00
Begin VB.Form frmFastings 
   Caption         =   "NetAcquire - Biochemistry Fasting Ranges"
   ClientHeight    =   3945
   ClientLeft      =   705
   ClientTop       =   1440
   ClientWidth     =   7020
   Icon            =   "frmFastings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   7020
   Begin VB.CommandButton bSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   795
      Left            =   5220
      Picture         =   "frmFastings.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   360
      Width           =   1515
   End
   Begin VB.CommandButton bCancel 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   795
      Left            =   5220
      Picture         =   "frmFastings.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2850
      Width           =   1530
   End
   Begin VB.Frame Frame3 
      Caption         =   "Glucose"
      Height          =   1035
      Left            =   270
      TabIndex        =   13
      Top             =   180
      Width           =   4605
      Begin VB.TextBox tText 
         Height          =   285
         Index           =   0
         Left            =   2580
         TabIndex        =   2
         Top             =   510
         Width           =   1815
      End
      Begin VB.TextBox tHigh 
         Height          =   285
         Index           =   0
         Left            =   1380
         TabIndex        =   1
         Top             =   510
         Width           =   975
      End
      Begin VB.TextBox tLow 
         Height          =   285
         Index           =   0
         Left            =   150
         TabIndex        =   0
         Top             =   510
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Printout Text"
         Height          =   195
         Index           =   0
         Left            =   3030
         TabIndex        =   20
         Top             =   300
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "High"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   0
         Left            =   1680
         TabIndex        =   15
         Top             =   270
         Width           =   330
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Low"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   0
         Left            =   450
         TabIndex        =   14
         Top             =   270
         Width           =   300
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Triglyceride"
      Height          =   1035
      Left            =   270
      TabIndex        =   12
      Top             =   2670
      Width           =   4605
      Begin VB.TextBox tText 
         Height          =   285
         Index           =   2
         Left            =   2580
         TabIndex        =   8
         Top             =   570
         Width           =   1815
      End
      Begin VB.TextBox tHigh 
         Height          =   285
         Index           =   2
         Left            =   1380
         TabIndex        =   7
         Top             =   570
         Width           =   975
      End
      Begin VB.TextBox tLow 
         Height          =   285
         Index           =   2
         Left            =   150
         TabIndex        =   6
         Top             =   570
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Printout Text"
         Height          =   195
         Index           =   2
         Left            =   3030
         TabIndex        =   22
         Top             =   360
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "High"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   2
         Left            =   1680
         TabIndex        =   19
         Top             =   330
         Width           =   330
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Low"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   2
         Left            =   450
         TabIndex        =   18
         Top             =   330
         Width           =   300
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cholesterol"
      Height          =   1035
      Left            =   270
      TabIndex        =   11
      Top             =   1380
      Width           =   4605
      Begin VB.TextBox tText 
         Height          =   285
         Index           =   1
         Left            =   2580
         TabIndex        =   5
         Top             =   510
         Width           =   1815
      End
      Begin VB.TextBox tHigh 
         Height          =   285
         Index           =   1
         Left            =   1380
         TabIndex        =   4
         Top             =   510
         Width           =   975
      End
      Begin VB.TextBox tLow 
         Height          =   285
         Index           =   1
         Left            =   150
         TabIndex        =   3
         Top             =   540
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Printout Text"
         Height          =   195
         Index           =   1
         Left            =   3030
         TabIndex        =   21
         Top             =   300
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "High"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   1
         Left            =   1680
         TabIndex        =   17
         Top             =   270
         Width           =   330
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Low"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   1
         Left            =   450
         TabIndex        =   16
         Top             =   270
         Width           =   300
      End
   End
End
Attribute VB_Name = "frmFastings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub bcancel_Click()

10        Unload Me

End Sub


Private Sub bSave_Click()

          Dim n As Long
          Dim tb As New Recordset
          Dim sql As String

10        For n = 0 To 2
20            sql = "SELECT * from fastings WHERE testname = '" & Choose(n + 1, "GLU", "CHO", "TRI") & "'"
30            Set tb = New Recordset
40            RecOpenServer 0, tb, sql
50            If tb.EOF Then tb.AddNew
60            tb!FastingLow = Format$(Val(tLow(n)))
70            tb!FastingHigh = Format$(Val(tHigh(n)))
80            tb!FastingText = tText(n)
90            tb!TestName = Choose(n + 1, "GLU", "CHO", "TRI")
100           If n + 1 = 1 Then
110               tb!Code = SysOptBioCodeForGlucose(0)
120           ElseIf n + 1 = 2 Then
130               tb!Code = SysOptBioCodeForChol(0)
140           ElseIf n + 1 = 3 Then
150               tb!Code = SysOptBioCodeForTrig(0)
160           End If
170           tb.Update
180       Next

190       bsave.Enabled = False

End Sub

Private Sub Form_Load()

          Dim tb As New Recordset
          Dim sql As String

10        sql = "SELECT * from fastings"
20        Set tb = New Recordset
30        RecOpenServer 0, tb, sql

40        Do While Not tb.EOF
50            Select Case Trim(tb!TestName)
              Case "GLU"
60                tLow(0) = tb!FastingLow
70                tHigh(0) = tb!FastingHigh
80                tText(0) = tb!FastingText
90            Case "CHO"
100               tLow(1) = tb!FastingLow
110               tHigh(1) = tb!FastingHigh
120               tText(1) = tb!FastingText
130           Case "TRI"
140               tLow(2) = tb!FastingLow
150               tHigh(2) = tb!FastingHigh
160               tText(2) = tb!FastingText
170           End Select
180           tb.MoveNext
190       Loop

End Sub


Private Sub tHigh_Change(Index As Integer)

10        tText(Index) = "( " & tLow(Index) & " - " & tHigh(Index) & " )"

End Sub

Private Sub tHigh_KeyPress(Index As Integer, KeyAscii As Integer)

10        bsave.Enabled = True

End Sub


Private Sub tLow_Change(Index As Integer)

10        tText(Index) = "( " & tLow(Index) & " - " & tHigh(Index) & " )"

End Sub

Private Sub tLow_KeyPress(Index As Integer, KeyAscii As Integer)

10        bsave.Enabled = True

End Sub


Private Sub tText_KeyPress(Index As Integer, KeyAscii As Integer)

10        bsave.Enabled = True

End Sub


