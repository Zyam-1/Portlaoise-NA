VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmGPStats 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire-GP Stats"
   ClientHeight    =   6870
   ClientLeft      =   60
   ClientTop       =   540
   ClientWidth     =   9885
   Icon            =   "frmGPStats.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   9885
   ShowInTaskbar   =   0   'False
   Begin MSComCtl2.DTPicker calToDate 
      Height          =   330
      Left            =   2745
      TabIndex        =   11
      Top             =   135
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "dd/MMM/yyyy"
      Format          =   59375619
      CurrentDate     =   38985
   End
   Begin MSComCtl2.DTPicker calFromDate 
      Height          =   330
      Left            =   855
      TabIndex        =   10
      Top             =   135
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "dd/MMM/yyyy"
      Format          =   59375619
      CurrentDate     =   38985
   End
   Begin VB.Frame Frame3 
      Caption         =   "Laboratory Totals"
      Height          =   1425
      Left            =   4500
      TabIndex        =   3
      Top             =   90
      Width           =   1845
      Begin VB.TextBox tpers 
         Height          =   285
         Left            =   840
         TabIndex        =   6
         Top             =   990
         Width           =   885
      End
      Begin VB.TextBox btests 
         Height          =   285
         Left            =   840
         TabIndex        =   5
         Text            =   "0"
         Top             =   600
         Width           =   885
      End
      Begin VB.TextBox bsamples 
         Height          =   285
         Left            =   840
         TabIndex        =   4
         Text            =   "0"
         Top             =   240
         Width           =   885
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tests"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   9
         Top             =   630
         Width           =   390
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "Tests per Sample"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   930
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Samples"
         Height          =   195
         Left            =   150
         TabIndex        =   7
         Top             =   270
         Width           =   600
      End
   End
   Begin VB.CommandButton bPrint 
      Caption         =   "&Print"
      Height          =   840
      Left            =   6480
      Picture         =   "frmGPStats.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   90
      Width           =   1245
   End
   Begin VB.CommandButton bcancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   885
      Left            =   8280
      Picture         =   "frmGPStats.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   90
      Width           =   1245
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   5235
      Left            =   135
      TabIndex        =   1
      Top             =   1530
      Width           =   9645
      _ExtentX        =   17013
      _ExtentY        =   9234
      _Version        =   393216
      Cols            =   8
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      FormatString    =   $"frmGPStats.frx":091E
   End
   Begin VB.Label Label4 
      Caption         =   "From"
      Height          =   240
      Left            =   225
      TabIndex        =   13
      Top             =   180
      Width           =   690
   End
   Begin VB.Label Label2 
      Caption         =   "to"
      Height          =   330
      Left            =   2475
      TabIndex        =   12
      Top             =   180
      Width           =   330
   End
End
Attribute VB_Name = "frmGPStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Calculate()

          Dim Bsn As Recordset
          Dim sn As New Recordset
          Dim sql As String
          Dim n As Long
          Dim SampleCounter As Long
          Dim TestCounter As Long
          Dim tpers As Single
          Dim TotTests As Long
          Dim TotSamples As Long
          Dim TW As Single
          Dim SW As Single


10        On Error GoTo Calculate_Error

20        g.Rows = 2
30        g.AddItem ""
40        g.RemoveItem 1

50        sql = "SELECT distinct gp from demographics WHERE " & _
                "rundate between '" & _
                Format(calFromDate, "dd/mmm/yyyy") & "' and " & _
                Format(calToDate, "dd/mmm/yyyy") & "'"
60        RecOpenClient 0, sn, sql
70        Do While Not sn.EOF
80            g.AddItem sn!GP & ""
90            sn.MoveNext
100       Loop

110       If g.Rows > 2 Then
120           g.RemoveItem 1
130       End If

140       For n = g.Rows - 1 To 1 Step -1
150           g.Col = 0
160           g.Row = n
170           If Trim(g) <> "" Then
180               sql = "SELECT * from demographics WHERE " & _
                        "gp = '" & g & "' and " & _
                        "rundate between '" & _
                        Format(calFromDate, "dd/mmm/yyyy") & "' and '" & _
                        Format(calToDate, "dd/mmm/yyyy") & "'"
190               RecOpenServer 0, sn, sql
200               SampleCounter = 0
210               TestCounter = 0
220               Do While Not sn.EOF
230                   sql = "SELECT * from bioresults WHERE " & _
                            "sampleid = '" & sn!SampleID & "'"
240                   RecOpenServer 0, Bsn, sql
250                   If Not Bsn.EOF Then
260                       SampleCounter = SampleCounter + 1
270                       Bsn.MoveLast
280                       TestCounter = TestCounter + Bsn.RecordCount
290                   End If
300                   sn.MoveNext
310               Loop
320           End If
330           If SampleCounter <> 0 Then
340               g.Col = 1
350               g = Format(TestCounter)
360               g.Col = 2
370               g = Format(SampleCounter)
380           Else
390               If g.Rows = 2 Then
400                   g.AddItem ""
410                   g.RemoveItem 1
420               Else
430                   g.RemoveItem g.Row
440               End If
450           End If
460       Next

470       g.Visible = False

480       TotTests = 0
490       TotSamples = 0
500       For n = 1 To g.Rows - 1
510           g.Row = n
520           g.Col = 1
530           TotTests = TotTests + Val(g)
540           g.Col = 2
550           TotSamples = TotSamples + Val(g)
560       Next

570       If TotSamples * TotTests <> 0 Then
580           g.AddItem "Totals" & vbTab & Format(TotTests) & vbTab & Format(TotSamples)
590           For n = 1 To g.Rows - 1
600               g.Row = n
610               g.Col = 1
620               TestCounter = Val(g)
630               If SampleCounter <> 0 Then
640                   TW = Val(g) / TotTests
650                   g.Col = 2
660                   SampleCounter = Val(g)
670                   If SampleCounter <> 0 Then
680                       SW = Val(g) / TotSamples
690                       tpers = TestCounter / SampleCounter
700                       g.Col = 3
710                       g = Format(tpers, "0.0")
720                       g.Col = 4
730                       g = Format(TW, "0.0%")
740                       g.Col = 5
750                       g = Format(SW, "0.0%")
760                       If Val(btests) <> 0 Then
770                           g.Col = 6
780                           g = Format(TestCounter / Val(btests), "0.0%")
790                       End If
800                       If Val(bsamples) <> 0 Then
810                           g.Col = 7
820                           g = Format(SampleCounter / Val(bsamples), "0.0%")
830                       End If
840                   End If
850               End If
860           Next
870       End If

880       If g.Rows > 2 Then
890           g.RemoveItem 1
900       End If

910       g.Visible = True



920       Exit Sub

Calculate_Error:

          Dim strES As String
          Dim intEL As Integer

930       intEL = Erl
940       strES = Err.Description
950       LogError "frmGPStats", "calculate", intEL, strES, sql


End Sub

Private Sub bcancel_Click()

10        Unload Me

End Sub


Private Sub bprint_Click()

          Dim Y As Long

10        On Error GoTo bprint_Click_Error

20        Printer.Print "General Hospital Portlaoise"
30        Printer.Print "GP Workload ";
40        Printer.Print Format(calFromDate, "dd/mmmm/yyyy"); " to ";
50        Printer.Print Format(calToDate, "dd/mmmm/yyyy")
60        Printer.Print

70        For Y = 0 To g.Rows - 1
80            g.Row = Y
90            g.Col = 0
100           Printer.Print Left(g, 20);
110           g.Col = 1
120           Printer.Print Tab(20); g;
130           g.Col = 2
140           Printer.Print Tab(26); g;
150           g.Col = 3
160           Printer.Print Tab(34); g;
170           g.Col = 4
180           Printer.Print Tab(39); g;
190           g.Col = 5
200           Printer.Print Tab(53); g;
210           g.Col = 6
220           Printer.Print Tab(69); g;
230           g.Col = 7
240           Printer.Print Tab(80); g
250       Next
260       Printer.EndDoc

270       Exit Sub

bprint_Click_Error:

          Dim strES As String
          Dim intEL As Integer



280       intEL = Erl
290       strES = Err.Description
300       LogError "frmGPStats", "bPrint_Click", intEL, strES


End Sub




Private Sub Form_Activate()

10        On Error GoTo Form_Activate_Error

20        Calculate

30        Exit Sub

Form_Activate_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmGPStats", "Form_Activate", intEL, strES


End Sub

Private Sub g_Click()

          Static SortOrder As Boolean

10        On Error GoTo g_Click_Error

20        If g.MouseRow = 0 Then
30            If SortOrder Then
40                g.Sort = flexSortGenericAscending
50            Else
60                g.Sort = flexSortGenericDescending
70            End If
80            SortOrder = Not SortOrder
90        End If

100       Exit Sub

g_Click_Error:

          Dim strES As String
          Dim intEL As Integer



110       intEL = Erl
120       strES = Err.Description
130       LogError "frmGPStats", "g_Click", intEL, strES


End Sub



