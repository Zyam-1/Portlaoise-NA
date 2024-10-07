VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmManageSensitivities 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Sensitivities"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13050
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   13050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid grdCurrent 
      Height          =   1305
      Left            =   900
      TabIndex        =   6
      Top             =   510
      Width           =   9825
      _ExtentX        =   17330
      _ExtentY        =   2302
      _Version        =   393216
      Cols            =   13
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      AllowUserResizing=   1
      FormatString    =   $"frmManageSensitivities.frx":0000
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   525
      Left            =   11340
      Picture         =   "frmManageSensitivities.frx":0089
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5100
      Width           =   1245
   End
   Begin MSFlexGridLib.MSFlexGrid grdArc 
      Height          =   2625
      Left            =   900
      TabIndex        =   7
      Top             =   1860
      Width           =   11985
      _ExtentX        =   21140
      _ExtentY        =   4630
      _Version        =   393216
      Cols            =   15
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      AllowUserResizing=   1
      FormatString    =   $"frmManageSensitivities.frx":06F3
   End
   Begin MSFlexGridLib.MSFlexGrid grdRepeat 
      Height          =   1305
      Left            =   900
      TabIndex        =   8
      Top             =   4560
      Width           =   9825
      _ExtentX        =   17330
      _ExtentY        =   2302
      _Version        =   393216
      Cols            =   13
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      AllowUserResizing=   1
      FormatString    =   $"frmManageSensitivities.frx":079A
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Sample ID"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   150
      Width           =   735
   End
   Begin VB.Label lblSampleID 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   900
      TabIndex        =   4
      Top             =   120
      Width           =   1245
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Current"
      Height          =   195
      Left            =   300
      TabIndex        =   3
      Top             =   540
      Width           =   510
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Repeats"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   4620
      Width           =   600
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Archive"
      Height          =   195
      Left            =   300
      TabIndex        =   1
      Top             =   1890
      Width           =   540
   End
End
Attribute VB_Name = "frmManageSensitivities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

10        Unload Me

End Sub


Private Sub Form_Activate()

10        FillCurrent
20        FillRepeat
30        FillArchive

End Sub

Private Sub FillCurrent()

          Dim sql As String
          Dim tb As Recordset
          Dim s As String

10        With grdCurrent
20            .Rows = 2
30            .AddItem ""
40            .RemoveItem 1

50            sql = "SELECT * FROM Sensitivities WHERE " & _
                    "SampleID = '" & Val(lblSampleID) + SysOptMicroOffset(0) & "'"
60            Set tb = New Recordset
70            RecOpenServer 0, tb, sql
80            Do While Not tb.EOF
90                s = tb!IsolateNumber & vbTab & _
                      tb!AntibioticCode & vbTab & _
                      tb!Result & vbTab & _
                      tb!Report & vbTab & _
                      tb!CPOFlag & vbTab & _
                      tb!Rundate & vbTab & _
                      tb!RunDateTime & vbTab & _
                      tb!RSI & vbTab & _
                      tb!Username & vbTab & _
                      tb!Forced & vbTab & _
                      tb!Secondary & vbTab & _
                      tb!Valid & vbTab & _
                      tb!AuthoriserCode & ""
100               .AddItem s
110               tb.MoveNext
120           Loop

130           If .Rows > 2 Then
140               .RemoveItem 1
150           End If
160       End With

End Sub

Private Sub FillRepeat()

          Dim sql As String
          Dim tb As Recordset
          Dim s As String

10        With grdRepeat
20            .Rows = 2
30            .AddItem ""
40            .RemoveItem 1

50            sql = "SELECT * FROM SensitivitiesRepeats WHERE " & _
                    "SampleID = '" & Val(lblSampleID) + SysOptMicroOffset(0) & "'"
60            Set tb = New Recordset
70            RecOpenServer 0, tb, sql
80            Do While Not tb.EOF
90                s = tb!IsolateNumber & vbTab & _
                      tb!AntibioticCode & vbTab & _
                      tb!Result & vbTab & _
                      tb!Report & vbTab & _
                      tb!CPOFlag & vbTab & _
                      tb!Rundate & vbTab & _
                      tb!RunDateTime & vbTab & _
                      tb!RSI & vbTab & _
                      tb!Username & vbTab & _
                      tb!Forced & vbTab & _
                      tb!Secondary & vbTab & _
                      tb!Valid & vbTab & _
                      tb!AuthoriserCode & ""
100               .AddItem s
110               tb.MoveNext
120           Loop

130           If .Rows > 2 Then
140               .RemoveItem 1
150           End If
160       End With

End Sub


Private Sub FillArchive()

          Dim sql As String
          Dim tb As Recordset
          Dim s As String

10        With grdArc
20            .Rows = 2
30            .AddItem ""
40            .RemoveItem 1

50            sql = "SELECT * FROM SensitivitiesAudit WHERE " & _
                    "SampleID = '" & Val(lblSampleID) + SysOptMicroOffset(0) & "'"
60            Set tb = New Recordset
70            RecOpenServer 0, tb, sql
80            Do While Not tb.EOF
90                s = tb!IsolateNumber & vbTab & _
                      tb!AntibioticCode & vbTab & _
                      tb!Result & vbTab & _
                      tb!Report & vbTab & _
                      tb!CPOFlag & vbTab & _
                      tb!Rundate & vbTab & _
                      tb!RunDateTime & vbTab & _
                      tb!RSI & vbTab & _
                      tb!Username & vbTab & _
                      tb!Forced & vbTab & _
                      tb!Secondary & vbTab & _
                      tb!Valid & vbTab & _
                      tb!AuthoriserCode & vbTab & _
                      tb!ArchivedBy & vbTab & _
                      tb!ArchiveDateTime
100               .AddItem s
110               tb.MoveNext
120           Loop

130           If .Rows > 2 Then
140               .RemoveItem 1
150           End If
160       End With

End Sub



