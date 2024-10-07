VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmOutstandingMicro 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Outstanding Microbiology"
   ClientHeight    =   4515
   ClientLeft      =   105
   ClientTop       =   1080
   ClientWidth     =   10095
   Icon            =   "frmOutstandingMicro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   10095
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   825
      Left            =   9090
      Picture         =   "frmOutstandingMicro.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Exit Screen"
      Top             =   3600
      Width           =   885
   End
   Begin MSFlexGridLib.MSFlexGrid grd 
      Height          =   4365
      Left            =   60
      TabIndex        =   0
      Top             =   90
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   7699
      _Version        =   393216
      Cols            =   3
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      AllowUserResizing=   1
      FormatString    =   $"frmOutstandingMicro.frx":0614
   End
End
Attribute VB_Name = "frmOutstandingMicro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

10        Unload Me

End Sub

Private Sub Form_Activate()

10        On Error GoTo Form_Activate_Error

20        LoadDetails

30        Exit Sub

Form_Activate_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmOutstandingMicro", "Form_Activate", intEL, strES


End Sub

Private Sub LoadDetails()

          Dim tb As Recordset
          Dim sql As String
          Dim s As String

10        On Error GoTo LoadDetails_Error

20        grd.Rows = 2
30        grd.AddItem ""
40        grd.RemoveItem 1

50        sql = "SELECT R.*, D.PatName " & _
                "FROM FaecalRequests R LEFT JOIN Demographics D " & _
                "ON R.SampleID = D.SampleID " & _
                "ORDER BY R.SampleID ASC"
60        Set tb = New Recordset
70        RecOpenServer 0, tb, sql
80        Do While Not tb.EOF

90            s = tb!SampleID - SysOptMicroOffset(0) & vbTab & _
                  tb!PatName & vbTab & _
                  "Faeces - "

100           If tb!cS Then s = s & "C & S "
110           If tb!ToxinAB Then s = s & "C. Difficile "
120           If tb!OP Then s = s & "O/P "
130           If tb!OB0 Or tb!OB1 Or tb!OB2 Then s = s & "Occult Blood "
140           If tb!Rota Then s = s & "Rota "
150           If tb!Adeno Then s = s & "Adeno "
160           If tb!HPylori Then s = s & "H.Pylori "
170           If tb!Coli0157 Then s = s & "Coli 0157 "
180           If tb!ssScreen Then s = s & "S/S Screen "
190           If tb!GDH Then s = s & "GDH "
200           If tb!PCR Then s = s & "PCR "

210           grd.AddItem s
220           tb.MoveNext

230       Loop

240       sql = "SELECT R.*, D.PatName " & _
                "FROM UrineRequests R LEFT JOIN Demographics D " & _
                "ON R.SampleID = D.SampleID " & _
                "ORDER BY R.SampleID ASC"
250       Set tb = New Recordset
260       RecOpenServer 0, tb, sql
270       Do While Not tb.EOF

280           s = tb!SampleID - SysOptMicroOffset(0) & vbTab & _
                  tb!PatName & vbTab & _
                  "Urine - "

290           If tb!cS Then s = s & "C & S "
300           If tb!Pregnancy Then s = s & "Pregnancy "
310           If tb!RedSub Then s = s & "Red Sub"

320           grd.AddItem s
330           tb.MoveNext

340       Loop

350       If grd.Rows > 2 Then
360           grd.RemoveItem 1
370       End If

380       Exit Sub

LoadDetails_Error:

          Dim strES As String
          Dim intEL As Integer

390       intEL = Erl
400       strES = Err.Description
410       LogError "frmOutstandingMicro", "LoadDetails", intEL, strES, sql

End Sub


