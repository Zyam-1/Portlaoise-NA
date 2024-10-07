VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmManageIsolates 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Isolates"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9645
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   9645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   525
      Left            =   7860
      Picture         =   "frmManageIsolates.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5430
      Width           =   1245
   End
   Begin MSFlexGridLib.MSFlexGrid grdCurrent 
      Height          =   1215
      Left            =   1140
      TabIndex        =   2
      Top             =   690
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   2143
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      AllowUserResizing=   1
      FormatString    =   "^Isolate # |<Organism Group           |<Organism Name             |<Qualifier              "
   End
   Begin MSFlexGridLib.MSFlexGrid grdRepeat 
      Height          =   2175
      Left            =   1140
      TabIndex        =   6
      Top             =   4230
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   3836
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      AllowUserResizing=   1
      FormatString    =   "^Isolate # |<Organism Group           |<Organism Name             |<Qualifier              "
   End
   Begin MSFlexGridLib.MSFlexGrid grdArc 
      Height          =   2175
      Left            =   1140
      TabIndex        =   7
      Top             =   1980
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   3836
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      AllowUserResizing=   1
      FormatString    =   "^Isolate # |<Organism Group           |<Organism Name             |<Qualifier              |<Archived By |<Archived Time "
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Archive"
      Height          =   195
      Left            =   540
      TabIndex        =   5
      Top             =   2070
      Width           =   540
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Repeats"
      Height          =   195
      Left            =   480
      TabIndex        =   4
      Top             =   4290
      Width           =   600
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Current"
      Height          =   195
      Left            =   540
      TabIndex        =   3
      Top             =   720
      Width           =   510
   End
   Begin VB.Label lblSampleID 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1140
      TabIndex        =   1
      Top             =   300
      Width           =   1245
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Sample ID"
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   330
      Width           =   735
   End
End
Attribute VB_Name = "frmManageIsolates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub FillCurrent()

          Dim sql As String
          Dim tb As Recordset
          Dim s As String

10        With grdCurrent
20            .Rows = 2
30            .AddItem ""
40            .RemoveItem 1

50            sql = "SELECT * FROM Isolates WHERE " & _
                    "SampleID = '" & Val(lblSampleID) + SysOptMicroOffset(0) & "'"
60            Set tb = New Recordset
70            RecOpenServer 0, tb, sql
80            Do While Not tb.EOF
90                s = tb!IsolateNumber & vbTab & _
                      tb!OrganismGroup & vbTab & _
                      tb!OrganismName & vbTab & _
                      tb!Qualifier & ""
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

10        On Error GoTo FillArchive_Error

20        With grdArc
30            .Rows = 2
40            .AddItem ""
50            .RemoveItem 1

60            sql = "SELECT * FROM IsolatesAudit WHERE " & _
                    "SampleID = '" & Val(lblSampleID) + SysOptMicroOffset(0) & "'"
70            Set tb = New Recordset
80            RecOpenServer 0, tb, sql
90            Do While Not tb.EOF
100               s = tb!IsolateNumber & vbTab & _
                      tb!OrganismGroup & vbTab & _
                      tb!OrganismName & vbTab & _
                      tb!Qualifier & vbTab & _
                      tb!ArchivedBy & vbTab & _
                      tb!ArchiveDateTime
110               .AddItem s
120               tb.MoveNext
130           Loop

140           If .Rows > 2 Then
150               .RemoveItem 1
160           End If

170       End With

180       Exit Sub

FillArchive_Error:

          Dim strES As String
          Dim intEL As Integer

190       intEL = Erl
200       strES = Err.Description
210       LogError "frmManageIsolates", "FillArchive", intEL, strES, sql

End Sub
Private Sub FillRepeat()

          Dim sql As String
          Dim tb As Recordset
          Dim s As String

10        With grdRepeat
20            .Rows = 2
30            .AddItem ""
40            .RemoveItem 1

50            sql = "SELECT * FROM IsolatesRepeats WHERE " & _
                    "SampleID = '" & Val(lblSampleID) + SysOptMicroOffset(0) & "'"
60            Set tb = New Recordset
70            RecOpenServer 0, tb, sql
80            Do While Not tb.EOF
90                s = tb!IsolateNumber & vbTab & _
                      tb!OrganismGroup & vbTab & _
                      tb!OrganismName & vbTab & _
                      tb!Qualifier & ""
100               .AddItem s
110               tb.MoveNext
120           Loop

130           If .Rows > 2 Then
140               .RemoveItem 1
150           End If
160       End With

End Sub
Private Sub cmdCancel_Click()

10        Unload Me

End Sub


Private Sub Form_Activate()

10        FillCurrent
20        FillRepeat
30        FillArchive

End Sub

