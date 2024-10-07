VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmStatSources 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Statistics"
   ClientHeight    =   8235
   ClientLeft      =   615
   ClientTop       =   690
   ClientWidth     =   7995
   Icon            =   "frmStatSources.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8235
   ScaleWidth      =   7995
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ProgressBar pb 
      Height          =   195
      Left            =   240
      TabIndex        =   9
      Top             =   1020
      Visible         =   0   'False
      Width           =   7485
      _ExtentX        =   13203
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.OptionButton oSource 
      Caption         =   "Wards"
      Height          =   255
      Index           =   0
      Left            =   4065
      TabIndex        =   8
      Top             =   150
      Value           =   -1  'True
      Width           =   1035
   End
   Begin VB.OptionButton oSource 
      Caption         =   "Clinicians"
      Height          =   255
      Index           =   1
      Left            =   4065
      TabIndex        =   7
      Top             =   450
      Width           =   1035
   End
   Begin VB.OptionButton oSource 
      Caption         =   "GPs"
      Height          =   255
      Index           =   2
      Left            =   4065
      TabIndex        =   6
      Top             =   750
      Width           =   1035
   End
   Begin VB.CommandButton bCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   825
      Left            =   6525
      Picture         =   "frmStatSources.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   180
      Width           =   1245
   End
   Begin VB.CommandButton bCalc 
      Caption         =   "Start"
      Height          =   825
      Left            =   5220
      Picture         =   "frmStatSources.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   180
      Width           =   1245
   End
   Begin VB.Frame Frame1 
      Caption         =   "Between Dates"
      Height          =   885
      Left            =   240
      TabIndex        =   1
      Top             =   90
      Width           =   3645
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   315
         Left            =   2160
         TabIndex        =   2
         Top             =   330
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   556
         _Version        =   393216
         Format          =   16580609
         CurrentDate     =   37606
      End
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   315
         Left            =   630
         TabIndex        =   3
         Top             =   330
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         _Version        =   393216
         Format          =   16580609
         CurrentDate     =   37606
      End
      Begin VB.Label Label2 
         Caption         =   "From"
         Height          =   240
         Left            =   180
         TabIndex        =   11
         Top             =   360
         Width           =   600
      End
      Begin VB.Label Label1 
         Caption         =   "To"
         Height          =   240
         Left            =   1935
         TabIndex        =   10
         Top             =   360
         Width           =   375
      End
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   6645
      Left            =   240
      TabIndex        =   0
      Top             =   1230
      Width           =   7485
      _ExtentX        =   13203
      _ExtentY        =   11721
      _Version        =   393216
      Cols            =   5
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
      FormatString    =   "<Source                 |<Total Samples |<Coag Samples |<Bio Samples |<Haem Samples    "
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
End
Attribute VB_Name = "frmStatSources"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub FillG()

          Dim sql As String
          Dim tb As New Recordset
          Dim tb1 As Recordset
          Dim tbd As Recordset
          Dim SourcePanelType As String
          Dim s As String
          Dim Total As Long
          Dim TotCoag As Long
          Dim TotHaem As Long
          Dim TotBio As Long

10        On Error GoTo FillG_Error

20        If oSource(0) Then
30            SourcePanelType = "W"
40        ElseIf oSource(1) Then
50            SourcePanelType = "C"
60        ElseIf oSource(2) Then
70            SourcePanelType = "G"
80        End If

90        g.Rows = 2
100       g.AddItem ""
110       g.RemoveItem 1
120       g.Visible = False

130       sql = "SELECT distinct SourcePanelName from SourcePanels WHERE " & _
                "SourcePanelType = '" & SourcePanelType & "'"
140       Set tb = New Recordset
150       RecOpenClient 0, tb, sql

160       If tb.EOF Then
170           g.Visible = True
180           Exit Sub
190       End If
200       pb.Visible = True
210       pb.Max = tb.RecordCount
220       pb = 0

230       Do While Not tb.EOF
240           pb = pb + 1
250           Total = 0
260           TotCoag = 0
270           TotHaem = 0
280           TotBio = 0
290           s = tb!SourcePanelName & vbTab
300           sql = "SELECT Content from SourcePanels WHERE " & _
                    "SourcePanelName = '" & tb!SourcePanelName & "' " & _
                    "and SourcePanelType = '" & SourcePanelType & "'"
310           Set tb1 = New Recordset
320           RecOpenClient 0, tb1, sql
330           Do While Not tb1.EOF
                  'Totals
340               sql = "SELECT Count (*) as Tot from Demographics WHERE "
350               Select Case SourcePanelType
                  Case "W": sql = sql & "Ward"
360               Case "C": sql = sql & "Clinician"
370               Case "G": sql = sql & "GP"
380               End Select
390               sql = sql & " = '" & tb1!Content & "' " & _
                        "and RunDate between '" & _
                        Format$(dtFrom, "dd/mmm/yyyy") & _
                        "' and '" & Format$(dtTo, "dd/mmm/yyyy") & "'"
400               Set tbd = New Recordset
410               RecOpenClient 0, tbd, sql
420               Total = Total + tbd!Tot

                  'Coag
430               sql = "SELECT distinct Demographics.SampleID from Demographics, CoagResults WHERE "
440               Select Case SourcePanelType
                  Case "W": sql = sql & "Ward"
450               Case "C": sql = sql & "Clinician"
460               Case "G": sql = sql & "GP"
470               End Select
480               sql = sql & " = '" & tb1!Content & "' " & _
                        "and Demographics.SampleID = CoagResults.SampleID " & _
                        "and Demographics.RunDate between '" & _
                        Format$(dtFrom, "dd/mmm/yyyy") & _
                        "' and '" & Format$(dtTo, "dd/mmm/yyyy") & "'"
490               Set tbd = New Recordset
500               RecOpenClient 0, tbd, sql
510               If Not tbd.EOF Then
520                   TotCoag = TotCoag + tbd.RecordCount
530               End If

                  'Bio
540               sql = "SELECT distinct Demographics.SampleID from Demographics, BioResults WHERE "
550               Select Case SourcePanelType
                  Case "W": sql = sql & "Ward"
560               Case "C": sql = sql & "Clinician"
570               Case "G": sql = sql & "GP"
580               End Select
590               sql = sql & " = '" & tb1!Content & "' " & _
                        "and Demographics.SampleID = BioResults.SampleID " & _
                        "and Demographics.RunDate between '" & _
                        Format$(dtFrom, "dd/mmm/yyyy") & _
                        "' and '" & Format$(dtTo, "dd/mmm/yyyy") & "'"
600               Set tbd = New Recordset
610               RecOpenClient 0, tbd, sql
620               If Not tbd.EOF Then
630                   TotBio = TotBio + tbd.RecordCount
640               End If

                  'Haem
650               sql = "SELECT distinct Demographics.SampleID from Demographics, HaemResults WHERE "
660               Select Case SourcePanelType
                  Case "W": sql = sql & "Ward"
670               Case "C": sql = sql & "Clinician"
680               Case "G": sql = sql & "GP"
690               End Select
700               sql = sql & " = '" & tb1!Content & "' " & _
                        "and Demographics.SampleID = HaemResults.SampleID " & _
                        "and Demographics.RunDate between '" & _
                        Format$(dtFrom, "dd/mmm/yyyy") & _
                        "' and '" & Format$(dtTo, "dd/mmm/yyyy") & "'"
710               Set tbd = New Recordset
720               RecOpenClient 0, tbd, sql
730               If Not tbd.EOF Then
740                   TotHaem = TotHaem + tbd.RecordCount
750               End If

760               tb1.MoveNext
770           Loop
780           s = s & Format$(Total) & vbTab & _
                  Format$(TotCoag) & vbTab & _
                  Format$(TotBio) & vbTab & _
                  Format$(TotHaem)

790           g.AddItem s
800           tb.MoveNext
810       Loop

820       pb.Visible = False
830       g.Visible = True
840       If g.Rows > 2 Then
850           g.RemoveItem 1
860       End If


870       Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer



880       intEL = Erl
890       strES = Err.Description
900       LogError "frmStatSources", "FillG", intEL, strES, sql


End Sub

Private Sub bCalc_Click()

10        On Error GoTo bCalc_Click_Error

20        FillG

30        Exit Sub

bCalc_Click_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmStatSources", "bCalc_Click", intEL, strES


End Sub

Private Sub bcancel_Click()

10        Unload Me

End Sub


Private Sub Form_Load()

10        On Error GoTo Form_Load_Error

20        dtTo = Format$(Now, "dd/mm/yyyy")
30        dtFrom = Format$(Now - 365, "dd/mm/yyyy")

40        Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer



50        intEL = Erl
60        strES = Err.Description
70        LogError "frmStatSources", "Form_Load", intEL, strES


End Sub


