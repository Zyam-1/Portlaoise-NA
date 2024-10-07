VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPrintOptions 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Print Results"
   ClientHeight    =   4770
   ClientLeft      =   1635
   ClientTop       =   1935
   ClientWidth     =   6075
   ForeColor       =   &H80000008&
   Icon            =   "frmPrintOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4770
   ScaleWidth      =   6075
   Begin VB.Frame fraNumSample2Print 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   90
      TabIndex        =   23
      Top             =   1380
      Width           =   2880
      Begin VB.Label lblNum2Print 
         Height          =   225
         Left            =   105
         TabIndex        =   24
         Top             =   180
         Width           =   2655
      End
   End
   Begin VB.CommandButton cmdGet 
      Caption         =   "Get Numbers"
      Height          =   735
      Left            =   675
      Picture         =   "frmPrintOptions.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   540
      Width           =   1320
   End
   Begin VB.OptionButton optDept 
      Caption         =   "Faeces"
      Height          =   195
      Index           =   6
      Left            =   4230
      TabIndex        =   19
      Top             =   630
      Width           =   870
   End
   Begin VB.OptionButton optDept 
      Caption         =   "Urine"
      Height          =   195
      Index           =   5
      Left            =   4230
      TabIndex        =   18
      Top             =   360
      Width           =   735
   End
   Begin VB.OptionButton optDept 
      Caption         =   "Immunology"
      Height          =   195
      Index           =   4
      Left            =   4230
      TabIndex        =   17
      Top             =   90
      Width           =   1140
   End
   Begin VB.OptionButton optDept 
      Alignment       =   1  'Right Justify
      Caption         =   "Endocrinology"
      Height          =   195
      Index           =   3
      Left            =   2835
      TabIndex        =   16
      Top             =   900
      Width           =   1320
   End
   Begin VB.OptionButton optDept 
      Alignment       =   1  'Right Justify
      Caption         =   "Coagulation"
      Height          =   195
      Index           =   2
      Left            =   3000
      TabIndex        =   15
      Top             =   630
      Width           =   1155
   End
   Begin MSComCtl2.DTPicker dtFrom 
      Height          =   345
      Left            =   0
      TabIndex        =   14
      Top             =   90
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   609
      _Version        =   393216
      Format          =   95682561
      CurrentDate     =   37112
   End
   Begin VB.OptionButton optDept 
      Alignment       =   1  'Right Justify
      Caption         =   "Biochemistry"
      Height          =   195
      Index           =   1
      Left            =   2970
      TabIndex        =   13
      Top             =   90
      Value           =   -1  'True
      Width           =   1185
   End
   Begin VB.OptionButton optDept 
      Alignment       =   1  'Right Justify
      Caption         =   "Haematology"
      Height          =   195
      Index           =   0
      Left            =   2925
      TabIndex        =   12
      Top             =   360
      Width           =   1230
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   705
      Left            =   1350
      Picture         =   "frmPrintOptions.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3855
      Width           =   975
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1485
      Left            =   3015
      TabIndex        =   7
      Top             =   2010
      Width           =   2985
      Begin VB.OptionButton v 
         Caption         =   "All (Valid, Not Valid, Printed and Not Printed)"
         Height          =   330
         Index           =   2
         Left            =   270
         TabIndex        =   11
         Top             =   1020
         Width           =   2580
      End
      Begin VB.OptionButton v 
         Caption         =   "Valid (Printed or Not Printed)"
         Height          =   195
         Index           =   1
         Left            =   270
         TabIndex        =   10
         Top             =   690
         Width           =   2310
      End
      Begin VB.OptionButton v 
         Caption         =   "Valid, not Printed"
         Height          =   195
         Index           =   0
         Left            =   270
         TabIndex        =   9
         Top             =   360
         Value           =   -1  'True
         Width           =   1545
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Between Numbers"
      Height          =   1485
      Left            =   90
      TabIndex        =   2
      Top             =   2010
      Width           =   2880
      Begin ComCtl2.UpDown UpDown2 
         Height          =   285
         Left            =   1710
         TabIndex        =   6
         Top             =   900
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   503
         _Version        =   327681
         Value           =   1
         BuddyControl    =   "tto"
         BuddyDispid     =   196621
         OrigLeft        =   1650
         OrigTop         =   1140
         OrigRight       =   1890
         OrigBottom      =   1335
         Max             =   99999
         Min             =   1
         Orientation     =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   285
         Left            =   1710
         TabIndex        =   5
         Top             =   450
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   503
         _Version        =   327681
         Value           =   1
         BuddyControl    =   "tfrom"
         BuddyDispid     =   196620
         OrigLeft        =   1620
         OrigTop         =   540
         OrigRight       =   1860
         OrigBottom      =   945
         Max             =   99999
         Min             =   1
         Orientation     =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.TextBox tfrom 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   270
         MaxLength       =   12
         TabIndex        =   4
         Top             =   420
         Width           =   1395
      End
      Begin VB.TextBox tto 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   270
         MaxLength       =   12
         TabIndex        =   3
         Top             =   870
         Width           =   1395
      End
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "&Stop Print"
      Height          =   705
      Left            =   2700
      Picture         =   "frmPrintOptions.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3855
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   705
      Left            =   3960
      Picture         =   "frmPrintOptions.frx":0D60
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3855
      Width           =   975
   End
   Begin MSComCtl2.DTPicker dtTo 
      Height          =   345
      Left            =   1530
      TabIndex        =   20
      Top             =   90
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   609
      _Version        =   393216
      Format          =   131727361
      CurrentDate     =   37112
   End
   Begin VB.Label Label1 
      Caption         =   "to"
      Height          =   195
      Left            =   1305
      TabIndex        =   21
      Top             =   135
      Width           =   285
   End
End
Attribute VB_Name = "frmPrintOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim tonumber As Long
Dim fromnumber As Long

Dim printing As Boolean


Private Sub cmdCancel_Click()

10        On Error GoTo cmdCancel_Click_Error

20        If printing Then Exit Sub

30        Unload Me

40        Exit Sub

cmdCancel_Click_Error:

          Dim strES As String
          Dim intEL As Integer



50        intEL = Erl
60        strES = Err.Description
70        LogError "frmPrintOptions", "cmdCancel_Click", intEL, strES

End Sub

Private Sub cmdGet_Click()

10        FillToAndFrom

End Sub

Private Sub cmdPrint_Click()

          Dim sql As String
          Dim tb As New Recordset
          Dim temp As Long
          Dim T As Single
          Dim PrintIt As Long
          Dim Total As Long
          Dim Dept As String
          Dim FromDate As String
          Dim ToDate As String

10        On Error GoTo cmdPrint_Click_Error

20        If printing Then Exit Sub

30        FromDate = Format$(dtFrom, "dd/mmm/yyyy")
40        ToDate = Format$(dtTo, "dd/mmm/yyyy")

50        printing = True

60        If Val(tfrom) > Val(tto) Then
70            temp = tto
80            tto = tfrom
90            tfrom = tto
100       End If

110       If optDept(5).Value = True Then    'Urine

120           sql = "SELECT DISTINCT U.SampleID FROM Demographics AS D, Urine AS U WHERE " & _
                    "D.RunDate BETWEEN '" & FromDate & "' AND '" & ToDate & "' " & _
                    "AND D.SampleID = U.SampleID " & _
                    "AND U.SampleID IN " & _
                    "( " & _
                    "  SELECT DISTINCT SampleID FROM PrintValidLog WHERE " & _
                    "  (Printed = 0 OR Printed IS NULL) AND Valid = 1 " & _
                    "  AND Department = 'D' ) " & _
                    "ORDER BY U.SampleID"
130           Set tb = New Recordset
140           RecOpenClient 0, tb, sql
150           Total = tb.recordCount
160           If Total > SysOptMaxSampleUrineBatchPrinting(0) Then
170               iMsg "Too many to print (" & Format$(Total) & ") reports." & vbCrLf & "Maximum " & SysOptMaxSampleUrineBatchPrinting(0)
180               printing = False
190               Exit Sub
200           End If

210           Do While Not tb.EOF
220               PrintResultUrnWin (Val(tb!SampleID) - SysOptMicroOffset(0))
230               tb.MoveNext
240           Loop
250           tb.Close
260       Else
270           Total = Abs(Val(tfrom) - Val(tto)) + 1

280           If Total > 50 Then


290               If optDept(6) Then
300                   sql = "SELECT count (distinct SampleID) as Tot from " & _
                            "Faeces WHERE " & _
                            "SampleID between '" & Val(tfrom) + SysOptMicroOffset(0) & "' and '" & Val(tto) + SysOptMicroOffset(0) & "'"
310               Else
320                   If optDept(0).Value = True Then
330                       Dept = "Haem"
340                   ElseIf optDept(1).Value = True Then
350                       Dept = "Bio"
360                   ElseIf optDept(2).Value = True Then
370                       Dept = "Coag"
380                   ElseIf optDept(3).Value = True Then
390                       Dept = "End"
400                   ElseIf optDept(4).Value = True Then
410                       Dept = "Imm"
420                   End If
430                   sql = "SELECT count (distinct SampleID) as Tot from " & _
                            Dept & "Results WHERE " & _
                            "SampleID between '" & Val(tfrom) & "' and '" & Val(tto) & "'"
440                   If v(0) Then
450                       sql = sql & " and valid = 1 and (printed = 0 or printed is null) "
460                   ElseIf v(1) Then
470                       sql = sql & " and valid = 1"
480                   End If
490               End If
500               Set tb = New Recordset
510               RecOpenServer 0, tb, sql
520               Total = tb!Tot
530               If Total > 50 Then
540                   iMsg "Too many to print (" & Format$(Total) & ") reports." & vbCrLf & "Maximum 50"
550                   printing = False
560                   Exit Sub
570               End If

580           End If

590           If Total > 20 Then
600               If iMsg("You requested to print " & Format$(Total) & " reports." & vbCrLf & _
                          "Are you sure?", vbYesNo + vbQuestion) = vbNo Then
610                   printing = False
620                   Exit Sub
630               End If
640           End If

650           For temp = Val(tfrom) To Val(tto)
660               PrintIt = True
670               If optDept(0) Then
680                   sql = "SELECT * from HaemResults WHERE " & _
                            "SampleID = '" & Format$(temp) & "'"
690               ElseIf optDept(1) Then
700                   sql = "SELECT * from BioResults WHERE " & _
                            "SampleID = '" & Format$(temp) & "'"
710               ElseIf optDept(2) Then
720                   sql = "SELECT * from CoagResults WHERE " & _
                            "SampleID = '" & Format$(temp) & "'"
730               ElseIf optDept(3) Then
740                   sql = "SELECT * from endResults WHERE " & _
                            "SampleID = '" & Format$(temp) & "'"
750               ElseIf optDept(4) Then
760                   sql = "SELECT * from immResults WHERE " & _
                            "SampleID = " & Format$(temp) & ""
770               ElseIf optDept(6) Then
780                   sql = "SELECT SampleID FROM PrintValidLog WHERE " & _
                            "(Printed = 0 OR Printed IS NULL) AND Valid = 1 " & _
                            "AND (   Department = 'A' " & _
                            "     OR Department = 'D' " & _
                            "     OR Department = 'F' " & _
                            "     OR Department = 'G' " & _
                            "     OR Department = 'O' " & _
                            "     OR Department = 'Y' ) " & _
                            "AND SampleID = " & Format$(temp) + SysOptMicroOffset(0) & ""
790               End If

800               Set tb = New Recordset
810               RecOpenServer 0, tb, sql
820               If tb.EOF Then
830                   PrintIt = False
840               Else
850                   If v(0) Then
860                       If optDept(6) = True Then
870                           PrintIt = True
880                       Else
890                           If tb!Valid = 0 Or IsNull(tb!Valid) Then PrintIt = False
900                           If tb!Printed = 1 Then PrintIt = False
910                       End If
920                   ElseIf v(1) Then
930                       If tb!Valid = 0 Then PrintIt = False
940                   End If
950               End If
960               If PrintIt Then
970                   cmdPrint.Visible = False
980                   cmdStop.Visible = True
990                   cmdCancel.Visible = False
1000                  T = Timer
1010                  If optDept(0) Then
1020                      PrintResultHaemWin Format$(temp)
1030                  ElseIf optDept(1) Then
1040                      sql = "UPDATE BioResults " & _
                                "Set Valid = 1, Printed = 0 " & _
                                "WHERE SampleID = '" & Format$(temp) & "'"
1050                      Cnxn(0).Execute sql
1060                      PrintResultBioWin Format$(temp)
1070                  ElseIf optDept(2) Then
1080                      sql = "UPDATE CoagResults " & _
                                "Set Valid = 1, Printed = 0 " & _
                                "WHERE SampleID = '" & Format$(temp) & "'"
1090                      Cnxn(0).Execute sql
1100                      PrintResultCWin Format$(temp)
1110                  ElseIf optDept(3) Then
1120                      sql = "UPDATE EndResults " & _
                                "Set Valid = 1, Printed = 0 " & _
                                "WHERE SampleID = '" & Format$(temp) & "'"
1130                      Cnxn(0).Execute sql
1140                      PrintResultEndWin Format$(temp)
1150                  ElseIf optDept(4) Then
1160                      sql = "UPDATE immResults " & _
                                "Set Valid = 1, Printed = 0 " & _
                                "WHERE SampleID = '" & Format$(temp) & "'"
1170                      Cnxn(0).Execute sql
1180                      PrintResultImmWin Format$(temp)
1190                  ElseIf optDept(6) Then
1200                      PrintResultUrnWin Format$(temp)
1210                  End If
                      '    Do While Timer - t < 2
                      '      DoEvents
                      '    Loop
1220              End If
1230          Next
1240      End If




1250      printing = False
1260      cmdPrint.Visible = True
1270      cmdStop.Visible = False
1280      cmdCancel.Visible = True

1290      Exit Sub

cmdPrint_Click_Error:

          Dim strES As String
          Dim intEL As Integer



1300      intEL = Erl
1310      strES = Err.Description
1320      LogError "frmPrintOptions", "cmdPrint_Click", intEL, strES, sql


End Sub

Private Sub cmdStop_Click()

10        On Error GoTo cmdStop_Click_Error


20        Unload Me

30        Exit Sub

cmdStop_Click_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmPrintOptions", "cmdStop_Click", intEL, strES

End Sub

Private Sub FillToAndFrom()

          Dim tb As New Recordset
          Dim sql As String
          Dim Index As Integer
          Dim FromDate As String
          Dim ToDate As String

10        On Error GoTo FillToAndFrom_Error

20        For Index = 0 To 6
30            If optDept(Index) Then
40                Exit For
50            End If
60        Next

70        FromDate = Format$(dtFrom, "dd/mmm/yyyy")
80        ToDate = Format$(dtTo, "dd/mmm/yyyy")

90        Select Case Index
          Case 0:
100           sql = "SELECT SampleID from HaemResults WHERE " & _
                    "rundatetime between '" & FromDate & " 00:00:00'  and '" & ToDate & " 23:59:59' " & _
                    " and sampleid < 9000000 "
110       Case 1:
120           sql = "SELECT distinct sampleid from BioResults WHERE " & _
                    "rundate between '" & FromDate & "' and  '" & ToDate & "' "
130       Case 2:
140           sql = "SELECT distinct sampleid from CoagResults WHERE " & _
                    "rundate between '" & FromDate & "' and  '" & ToDate & "' and sampleid > '1000' "
150       Case 3:
160           sql = "SELECT distinct sampleid from endResults WHERE " & _
                    "rundate = '" & FromDate & "' and  '" & ToDate & "') "
170       Case 4:
180           sql = "SELECT distinct sampleid from immResults WHERE " & _
                    "rundate between '" & FromDate & "' and  '" & ToDate & "' "
190       Case 5:    'Urine
200           sql = "SELECT DISTINCT U.SampleID FROM Demographics AS D, Urine AS U WHERE " & _
                    "D.RunDate BETWEEN '" & FromDate & "' AND '" & ToDate & "' " & _
                    "AND D.SampleID = U.SampleID " & _
                    "AND U.SampleID IN " & _
                    "( " & _
                    "  SELECT DISTINCT SampleID FROM PrintValidLog WHERE " & _
                    "  (Printed = 0 OR Printed IS NULL) AND Valid = 1 " & _
                    "  AND Department = 'D' ) " & _
                    "ORDER BY U.SampleID"
              'Urine sample must have validated C&S before printing hence 'D'

210       Case 6:
              'A Rota/Adeno
              'B Biochemistry
              'C Coagulation
              'D C and S
              'E Endocrinology
              'F FOB
              'G C.diff
              'H Haematology
              'I Immunology
              'M Micro
              'O Ova/Parasites
              'R Red Sub
              'S ESR
              'U Urine
              'V RSV
              'X External
              'Y H.Pylori
220           sql = "SELECT DISTINCT F.SampleID FROM Demographics AS D, Faeces AS F WHERE " & _
                    "D.RunDate BETWEEN '" & FromDate & "' AND '" & ToDate & "' " & _
                    "AND D.SampleID = F.SampleID " & _
                    "AND F.SampleID IN " & _
                    "( SELECT DISTINCT SampleID FROM PrintValidLog WHERE " & _
                    "  (Printed = 0 OR Printed IS NULL) AND Valid = 1 " & _
                    "  AND (   Department = 'A' " & _
                    "       OR Department = 'D' " & _
                    "       OR Department = 'F' " & _
                    "       OR Department = 'G' " & _
                    "       OR Department = 'O' " & _
                    "       OR Department = 'Y' ) " & _
                    ") ORDER BY F.SampleID"
230       End Select

240       If Index < 5 Then
250           If v(0) Then
260               sql = sql & " and ((Valid = 1 and printed is null) or (Valid = 1 and printed = '')  or (Valid = 1 and printed = 0)) "
270           ElseIf v(1) Then
280               sql = sql & " and Valid = 1 "
290           End If
300           sql = sql & " order by sampleid"
310       End If

320       Set tb = New Recordset
330       RecOpenClient 0, tb, sql

340       If tb.EOF Then
350           fromnumber = 0
360           tfrom = ""
370           tonumber = 0
380           tto = ""
390           Exit Sub
400       Else
410           If Index < 5 Then
420               tfrom = tb!SampleID
430               fromnumber = Val(tfrom)
440               tb.MoveLast
450               tto = tb!SampleID
460               tonumber = Val(tto)
470           Else
480               tfrom = Format$(Val(tb!SampleID) - SysOptMicroOffset(0))
490               fromnumber = Val(tfrom)
500               tb.MoveLast
510               tto = Format$(Val(tb!SampleID) - SysOptMicroOffset(0))
520               tonumber = Val(tto)
530               lblNum2Print = "Number of samples to print: " & tb.recordCount
540               fraNumSample2Print.Visible = True
550           End If
560       End If

570       Exit Sub

FillToAndFrom_Error:

          Dim strES As String
          Dim intEL As Integer

580       intEL = Erl
590       strES = Err.Description
600       LogError "frmPrintOptions", "FillToAndFrom", intEL, strES, sql

End Sub

Private Sub Form_Load()

10        On Error GoTo Form_Load_Error

20        Me.Refresh

30        dtTo = Format$(Now, "dd/mm/yyyy")
40        dtFrom = Format$(Now, "dd/mm/yyyy")

50        UpDown1.Max = 9999999
60        UpDown2.Max = 9999999

70        optDept(3).Visible = SysOptDeptEnd(0)
80        optDept(4).Visible = SysOptDeptImm(0)

90        FillToAndFrom

100       printing = False

110       Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer



120       intEL = Erl
130       strES = Err.Description
140       LogError "frmPrintOptions", "Form_Load", intEL, strES


End Sub




Private Sub optDept_Click(Index As Integer)

          Dim n As Long

10        On Error GoTo optDept_Click_Error

20        v(0).Visible = True
30        v(0).Value = True

40        For n = 0 To 6
50            If n <> Index Then
60                optDept(n).ForeColor = vbBlack
70            End If
80        Next

90        optDept(Index).ForeColor = vbRed

100       If Index = 5 Or Index = 6 Then
110           v(1).Enabled = False
120           v(2).Enabled = False
130       Else
140           v(1).Enabled = True
150           v(2).Enabled = True
160       End If
170       tfrom = ""
180       tto = ""
190       lblNum2Print = ""
200       fraNumSample2Print.Visible = False

210       Exit Sub

optDept_Click_Error:

          Dim strES As String
          Dim intEL As Integer



220       intEL = Erl
230       strES = Err.Description
240       LogError "frmPrintOptions", "optDept_Click", intEL, strES

End Sub


Private Sub tfrom_Change()
10        lblNum2Print = ""
20        fraNumSample2Print.Visible = False
End Sub


Private Sub tto_Change()
10        lblNum2Print = ""
20        fraNumSample2Print.Visible = False
End Sub
