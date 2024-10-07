VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmNotValidatedPrinted 
   Caption         =   "Form1"
   ClientHeight    =   8655
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12360
   LinkTopic       =   "Form1"
   ScaleHeight     =   8655
   ScaleWidth      =   12360
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPrintSamples 
      Appearance      =   0  'Flat
      Caption         =   "&Print Samples"
      Height          =   1035
      Left            =   10920
      Picture         =   "frmNotValidatedPrinted.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   6000
      Width           =   1245
   End
   Begin VB.CommandButton bcancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   975
      Left            =   10920
      Picture         =   "frmNotValidatedPrinted.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   7260
      Width           =   1245
   End
   Begin VB.CommandButton bprint 
      Appearance      =   0  'Flat
      Caption         =   "&Print List"
      Height          =   975
      Left            =   10920
      Picture         =   "frmNotValidatedPrinted.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   2175
      Width           =   1245
   End
   Begin VB.Frame Frame1 
      Height          =   1635
      Index           =   1
      Left            =   240
      TabIndex        =   14
      Top             =   240
      Width           =   4455
      Begin VB.OptionButton optNotValidatedPrinted 
         Caption         =   "Neither validated nor printed"
         Height          =   315
         Left            =   300
         TabIndex        =   17
         Top             =   1290
         Width           =   2475
      End
      Begin VB.OptionButton optNotPrinted 
         Caption         =   "Not Printed"
         Height          =   195
         Left            =   300
         TabIndex        =   16
         Top             =   1035
         Width           =   2475
      End
      Begin VB.OptionButton optNotValidated 
         Caption         =   "Not validated"
         Height          =   255
         Left            =   300
         TabIndex        =   15
         Top             =   720
         Value           =   -1  'True
         Width           =   2895
      End
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   315
         Left            =   765
         TabIndex        =   18
         Top             =   300
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         Format          =   175636481
         CurrentDate     =   37096
      End
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   315
         Left            =   2760
         TabIndex        =   19
         Top             =   300
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         Format          =   175636481
         CurrentDate     =   37096
      End
      Begin VB.Label lblTo 
         Caption         =   "To"
         Height          =   195
         Left            =   2430
         TabIndex        =   21
         Top             =   360
         Width           =   435
      End
      Begin VB.Label lblFrom 
         Caption         =   "From"
         Height          =   195
         Left            =   300
         TabIndex        =   20
         Top             =   360
         Width           =   435
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Discipline"
      Height          =   1635
      Index           =   0
      Left            =   4920
      TabIndex        =   1
      Top             =   240
      Width           =   3930
      Begin VB.OptionButton optBio 
         Caption         =   "Biochemistry"
         Height          =   255
         Left            =   585
         TabIndex        =   12
         Top             =   300
         Value           =   -1  'True
         Width           =   1485
      End
      Begin VB.OptionButton optHaem 
         Caption         =   "Haematology"
         Height          =   255
         Left            =   585
         TabIndex        =   11
         Top             =   615
         Width           =   1485
      End
      Begin VB.OptionButton optCoag 
         Caption         =   "Coagulation"
         Height          =   255
         Left            =   585
         TabIndex        =   10
         Top             =   930
         Width           =   1485
      End
      Begin VB.OptionButton optExt 
         Caption         =   "External"
         Height          =   255
         Left            =   2100
         TabIndex        =   9
         Top             =   615
         Width           =   1485
      End
      Begin VB.OptionButton optEnd 
         Caption         =   "Endocrinology"
         Height          =   255
         Left            =   585
         TabIndex        =   8
         Top             =   1245
         Width           =   1575
      End
      Begin VB.OptionButton optImm 
         Caption         =   "Immunology"
         Height          =   255
         Left            =   2100
         TabIndex        =   7
         Top             =   300
         Width           =   1575
      End
      Begin VB.OptionButton optBG 
         Caption         =   "Blood Gas"
         Height          =   255
         Left            =   2100
         TabIndex        =   6
         Top             =   930
         Width           =   1575
      End
      Begin VB.OptionButton optHisto 
         Caption         =   "Histology"
         Height          =   255
         Left            =   4080
         TabIndex        =   5
         Top             =   615
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.OptionButton optCyto 
         Caption         =   "Cytology"
         Height          =   255
         Left            =   4080
         TabIndex        =   4
         Top             =   930
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.OptionButton optMicro 
         Caption         =   "Microbiology"
         Height          =   255
         Left            =   4080
         TabIndex        =   3
         Top             =   300
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.OptionButton optSemen 
         Caption         =   "Semen Analysis"
         Height          =   255
         Left            =   4080
         TabIndex        =   2
         Top             =   1245
         Visible         =   0   'False
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdRefresh 
      Appearance      =   0  'Flat
      Caption         =   "Start"
      Height          =   975
      Left            =   9240
      Picture         =   "frmNotValidatedPrinted.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   660
      Width           =   1245
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   6060
      Left            =   240
      TabIndex        =   22
      Top             =   2175
      Width           =   10545
      _ExtentX        =   18600
      _ExtentY        =   10689
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      GridLines       =   3
      GridLinesFixed  =   3
      AllowUserResizing=   1
      FormatString    =   $"frmNotValidatedPrinted.frx":0C28
   End
   Begin VB.Label lblClick 
      Caption         =   "Please click on the sample id to view details"
      Height          =   255
      Left            =   240
      TabIndex        =   25
      Top             =   1920
      Width           =   8595
   End
   Begin VB.Label lblTotal 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   240
      TabIndex        =   13
      Top             =   8340
      Width           =   480
   End
End
Attribute VB_Name = "frmNotValidatedPrinted"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub SetFormCaption()

10        On Error GoTo SetFormCaption_Error

20        If optBio.Value = True Then
30            Me.Caption = "NetAcquire - " & optBio.Caption & " Unvalidated / Not Printed Samples"
40        ElseIf optHaem.Value = True Then
50            Me.Caption = "NetAcquire - " & optHaem.Caption & " Unvalidated / Not Printed Samples"
60        ElseIf optCoag.Value = True Then
70            Me.Caption = "NetAcquire - " & optCoag.Caption & " Unvalidated / Not Printed Samples"
80        ElseIf optEnd.Value = True Then
90            Me.Caption = "NetAcquire - " & optEnd.Caption & " Unvalidated / Not Printed Samples"
100       ElseIf optImm.Value = True Then
110           Me.Caption = "NetAcquire - " & optImm.Caption & " Unvalidated / Not Printed Samples"
120       ElseIf optExt.Value = True Then
130           Me.Caption = "NetAcquire - " & optExt.Caption & " Unvalidated / Not Printed Samples"
140       ElseIf optBG.Value = True Then
150           Me.Caption = "NetAcquire - " & optBG.Caption & " Unvalidated / Not Printed Samples"
160       ElseIf optHisto.Value = True Then
170           Me.Caption = "NetAcquire - " & optHisto.Caption & " Unvalidated / Not Printed Samples"
180       ElseIf optCyto.Value = True Then
190           Me.Caption = "NetAcquire - " & optCyto.Caption & " Unvalidated / Not Printed Samples"
200       ElseIf optMicro.Value = True Then
210           Me.Caption = "NetAcquire - " & optMicro.Caption & " Unvalidated / Not Printed Samples"
220       ElseIf optSemen.Value = True Then
230           Me.Caption = "NetAcquire - " & optSemen.Caption & " Unvalidated / Not Printed Samples"
240       End If
250       Exit Sub

SetFormCaption_Error:

          Dim strES As String
          Dim intEL As Integer

260       intEL = Erl
270       strES = Err.Description
280       LogError "frmDaily", "SetFormCaption", intEL, strES

End Sub


Private Sub FillG()

      Dim tb As New Recordset
      Dim s As String
      Dim sql As String
      Dim Asql As String
      Dim Bsql As String
      Dim OldSampleID As String
      Dim NewSampleID As String
      Dim Disc As String
      Dim TestColumn As String
      Dim ResultColumn As String
      Dim TableName As String
      Dim DateColumn As String
      Dim Selection As String


10    On Error GoTo FillG_Error

20    ClearFGrid g

30    DateColumn = "RunDate"
40    If optBio Then
50        Disc = "Bio"
60        TestColumn = "ShortName"
70        TableName = "BioResults"
80        DateColumn = "RunTime"
90    ElseIf optHaem Then
100       Disc = "Haem"
110       TestColumn = "AnalyteName"
120       ResultColumn = "RBC"
130       TableName = "HaemResults"
140   ElseIf optCoag Then
150       Disc = "Coag"
160       TestColumn = "TestName"
170       TableName = "CoagResults"
180   ElseIf optExt Then
190       Disc = "Ext"
200       TestColumn = "Analyte"
210       TableName = "ExtResults"
220       DateColumn = "RetDate"
230   ElseIf optEnd Then
240       Disc = "End"
250       TestColumn = "ShortName"
260       TableName = "EndResults"
270   ElseIf optImm Then
280       Disc = "Imm"
290       TestColumn = "ShortName"
300       TableName = "ImmResults"
310   ElseIf optBG Then
320       Disc = "Bga"
330       TestColumn = "ShortName"
340       TableName = "BgaResults"
350   ElseIf optCyto Then
360       Disc = "Cyto"
370       TestColumn = ""
380   ElseIf optHisto Then
390       Disc = "Histo"
400       TestColumn = ""
410   ElseIf optMicro Then
420       TableName = "PrintValidLog"
430   ElseIf optSemen Then
440       Disc = "Semen"
450       TestColumn = ""
460       ResultColumn = "Motility"
470       TableName = "SemenResults"
480       DateColumn = "DateTimeOfRecord"
490   Else
500       g.Visible = True
510       iMsg "No Discipline Choosen!"
520       Exit Sub
530   End If

540   If optNotValidated Then
550       Selection = "AND R.Valid = 0"
560   ElseIf optNotPrinted Then
570       Selection = "AND R.Printed = 0"
580   ElseIf optNotValidatedPrinted Then
590       Selection = "AND R.Valid = 0 AND R.Printed = 0"
600   End If

610   If optCoag Or optEnd Or optImm Or optBG Or optExt Or optSemen Then
620       sql = "SELECT DISTINCT R.SampleID, D.PatName,D.Chart,D.GP,D.Ward,D.Clinician FROM " & TableName & " R " & _
                "INNER JOIN Demographics D ON R.SampleId = D.SampleID " & _
                "INNER JOIN " & Disc & "TestDefinitions T ON R.Code = T.Code " & _
                "WHERE R." & DateColumn & " BETWEEN '" & Format$(dtFrom, "dd/mmm/yyyy") & "' AND '" & Format$(dtTo, "dd/mmm/yyyy") & "' " & _
                "AND T.Printable = '1' " & _
                Selection
630   ElseIf optHaem Then
640       sql = "SELECT DISTINCT R.SampleID, D.PatName,D.Chart,D.GP,D.Ward,D.Clinician FROM " & TableName & " R " & _
                "INNER JOIN Demographics D ON R.SampleId = D.SampleID " & _
                "WHERE R." & DateColumn & " BETWEEN '" & Format$(dtFrom, "dd/mmm/yyyy") & "' AND '" & Format$(dtTo, "dd/mmm/yyyy") & "' " & _
                Selection

650   ElseIf optBio Then
660       sql = "SELECT DISTINCT R.SampleID, D.PatName,D.Chart,D.GP,D.Ward,D.Clinician FROM " & TableName & " R " & _
                "INNER JOIN Demographics D ON R.SampleId = D.SampleID " & _
                "INNER JOIN " & Disc & "TestDefinitions T ON R.Code = T.Code " & _
                "WHERE R." & DateColumn & " BETWEEN '" & Format$(dtFrom, "dd/mmm/yyyy 00:00:00") & "' AND '" & Format$(dtTo, "dd/mmm/yyyy 23:59:00") & "' " & _
                "AND T.Printable = '1' " & _
                Selection
670   End If
680   Set tb = New Recordset
690   RecOpenServer 0, tb, sql
700   While Not tb.EOF
710       If optHisto Then
720           s = Val(tb!Hyear) & "/" & Val(tb!SampleID) - (SysOptHistoOffset(0) + (Val(Swap_Year(Trim(tb!Hyear))) * 1000)) & "H"
730       ElseIf optCyto Then
740           s = Val(tb!Hyear) & "/" & Val(tb!SampleID) - (SysOptCytoOffset(0) + (Val(Swap_Year(Trim(tb!Hyear))) * 1000)) & "C"
750       ElseIf optMicro Then
760           s = Val(tb!SampleID) - SysOptMicroOffset(0)
770       ElseIf optSemen Then
780           s = Val(tb!SampleID) - SysOptSemenOffset(0)
790       Else
800           s = tb!SampleID
810       End If

820       s = s & vbTab & _
              tb!PatName & vbTab & _
              tb!Chart & vbTab & _
              tb!GP & vbTab & _
              tb!Ward & vbTab & _
              tb!Clinician & ""

830       g.AddItem s
840       tb.MoveNext

850   Wend


860   FixG g

870   If g.Rows > 2 Then
880       lblTotal = "Total samples : " & g.Rows - 1
890   Else
900       lblTotal = ""
910   End If

920   Exit Sub

FillG_Error:

      Dim strES As String
      Dim intEL As Integer

930   intEL = Erl
940   strES = Err.Description
950   LogError "frmNotValidatedPrinted", "FillG", intEL, strES, sql

End Sub


Private Sub bcancel_Click()

10        Unload Me

End Sub

Private Sub bprint_Click()

10        On Error GoTo bprint_Click_Error

          Dim Y As Long
          Dim X As Long
          Dim sql As String
          Dim sn As New Recordset




20        Printer.Orientation = vbPRORLandscape
30        Printer.Font.Name = "Courier New"
40        PrintText FormatString("Unvalidated / Not Printed Samples List", 99, , AlignCenter), 10, True, , , , True
50        PrintText FormatString("From " & Format(dtFrom, "dd/mmm/yyyy") & " to " & Format(dtTo, "dd/mmm/yyyy"), 99, , AlignCenter), 10, True, , , , True
60        PrintText String(107, "-"), , , , , , True



70        For Y = 0 To g.Rows - 1


80            PrintText FormatString(g.TextMatrix(Y, 0), 10, "|"), 9, IIf(Y = 0, True, False)   'sample id
90            PrintText FormatString(g.TextMatrix(Y, 1), 30, "|"), 9, IIf(Y = 0, True, False)     'patient name
100           PrintText FormatString(g.TextMatrix(Y, 2), 10, "|"), 9, IIf(Y = 0, True, False)     'test name
110           PrintText FormatString(g.TextMatrix(Y, 3), 20, "|"), 9, IIf(Y = 0, True, False)      'gp
120           PrintText FormatString(g.TextMatrix(Y, 4), 20, "|"), 9, IIf(Y = 0, True, False)  'ward
130           PrintText FormatString(g.TextMatrix(Y, 5), 20), 9, IIf(Y = 0, True, False), , , , True     'result
              'PrintText FormatString(g.TextMatrix(Y, 6), 10), 9, IIf(Y = 0, True, False), , , , True    'return date

140           If Y = 0 Then PrintText String(107, "-"), , , , , , True
150       Next

160       Printer.EndDoc

170       Exit Sub

bprint_Click_Error:

          Dim strES As String
          Dim intEL As Integer

180       intEL = Erl
190       strES = Err.Description
200       LogError "frmNotValidatedPrinted", "bprint_Click", intEL, strES

End Sub

Private Sub cmdPrintSamples_Click()
10    On Error GoTo cmdPrintSamples_Click_Error
      Dim DeptSel    As String
      Dim i          As Integer
      Dim NewDepartment As String
      Dim Department As String
      Dim sql As String

20    If g.Rows = 1 Then Exit Sub
30    If iMsg("Are you sure you want to send " & g.Rows - 1 & " samples for printing?", vbYesNo, "Batch Printing") = vbNo Then
40         Exit Sub
50    End If

60    With g
70        For i = 1 To .Rows - 1
80            If .TextMatrix(i, 0) <> "" Then
90                If optBio.Value = True Then
100                   Department = "B"
110               ElseIf optHaem.Value = True Then
120                   Department = "H"
130               ElseIf optCoag.Value = True Then
140                   Department = "C"
150               ElseIf optEnd.Value = True Then
160                   Department = "E"
170               ElseIf optImm.Value = True Then
180                   Department = "I"
190               ElseIf optBG.Value = True Then
200                   Department = "Q"
210               ElseIf optExt.Value = True Then
220                   Department = "X"
230               End If
                  
                  'If Department = "I" And IsAllergy() Then Department = "W"
240               If SysOptRealImm(0) And Department = "I" Then
250                   NewDepartment = "J"
260               Else
270                   NewDepartment = Department
280               End If
290               LogTimeOfPrinting .TextMatrix(i, 0), Department
300               sql = "IF EXISTS (SELECT * FROM PrintPending WHERE " & _
                        "           Department = '" & Department & "' " & _
                        "           AND SampleID = '" & .TextMatrix(i, 0) & "' " & _
                        "           AND COALESCE(FaxNumber, '') = '') " & _
                        "    UPDATE PrintPending " & _
                        "    SET Department = '" & NewDepartment & "', " & _
                        "    Initiator = '" & UserName & "', " & _
                        "    Ward = '" & AddTicks(.TextMatrix(i, 4)) & "', " & _
                        "    Clinician = '" & AddTicks(.TextMatrix(i, 5)) & "', " & _
                        "    GP = '" & AddTicks(.TextMatrix(i, 3)) & "', " & _
                        "    UsePrinter = '', " & _
                        "    pTime = getdate() " & _
                        "    WHERE Department = '" & Department & "' " & _
                        "    AND SampleID = '" & .TextMatrix(i, 0) & "' " & _
                        "    AND COALESCE(FaxNumber, '') = '' " & _
                        "ELSE " & _
                        "    INSERT INTO PrintPending " & _
                          "    (SampleID, Department, Initiator, Ward, Clinician, GP, UsePrinter, pTime) "
310               sql = sql & _
                        "    VALUES ( " & _
                        "    '" & .TextMatrix(i, 0) & "', " & _
                        "    '" & NewDepartment & "', " & _
                        "    '" & UserName & "', " & _
                        "    '" & AddTicks(.TextMatrix(i, 4)) & "', " & _
                        "    '" & AddTicks(.TextMatrix(i, 5)) & "', " & _
                        "    '" & AddTicks(.TextMatrix(i, 3)) & "', " & _
                        "    '', " & _
                        "    getdate() )"
320               Cnxn(0).Execute sql

330           End If
340       Next i
350       MsgBox .Rows - 1 & " Samples sent for printing"
360   End With

370   FillG

380   Exit Sub

cmdPrintSamples_Click_Error:

      Dim strES      As String
      Dim intEL      As Integer

390   intEL = Erl
400   strES = Err.Description
410   LogError "frmNotValidatedPrinted", "cmdPrintSamples_Click", intEL, strES
End Sub

Private Sub cmdRefresh_Click()
10        FillG

End Sub

Private Sub Form_Load()

10        On Error GoTo Form_Load_Error

20        dtFrom = Format$(Now, "dd/mm/yyyy")
30        dtTo = Format$(Now, "dd/mm/yyyy")
          'optEnd.Enabled = SysOptDeptEnd(0)
          'optImm.Enabled = SysOptDeptImm(0)
          'optBG.Enabled = SysOptDeptBga(0)
          'optMicro.Enabled = False 'SysOptDeptMicro(0)
          'optSemen.Enabled = SysOptDeptSemen(0)
          'optHisto.Enabled = SysOptDeptHisto(0)
          'optCyto.Enabled = SysOptDeptCyto(0)

40        SetFormCaption

50        Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

60        intEL = Erl
70        strES = Err.Description
80        LogError "frmNotValidatedPrinted", "Form_Load", intEL, strES

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
90        Else
100           If g.MouseCol = 0 Then
110               If UserName = "" Then
120                   iMsg "Please logon to system to view sample"
130               Else

140                   With frmEditAll
150                       .txtSampleID = g.TextMatrix(g.MouseRow, g.MouseCol)
160                       Unload Me
170                       .txtSampleID_LostFocus
180                       .Show 1
190                   End With
200               End If
210           End If
220       End If

230       Exit Sub

g_Click_Error:

          Dim strES As String
          Dim intEL As Integer

240       intEL = Erl
250       strES = Err.Description
260       LogError "frmDaily", "g_Click", intEL, strES

End Sub

Private Sub optBG_Click()
10        On Error GoTo optBG_Click_Error

20        SetFormCaption
30        FillG


40        Exit Sub

optBG_Click_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmDaily", "optBG_Click", intEL, strES

End Sub

Private Sub optBio_Click()
10        On Error GoTo optBio_Click_Error

20        SetFormCaption
30        FillG
40        'g.ColWidth(6) = 2500

50        Exit Sub

optBio_Click_Error:

          Dim strES As String
          Dim intEL As Integer

60        intEL = Erl
70        strES = Err.Description
80        LogError "frmDaily", "optBio_Click", intEL, strES

End Sub

Private Sub optCoag_Click()
10        On Error GoTo optCoag_Click_Error

20        SetFormCaption
30        FillG


40        Exit Sub

optCoag_Click_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmDaily", "optCoag_Click", intEL, strES

End Sub

Private Sub optCyto_Click()
10        On Error GoTo optCyto_Click_Error

20        SetFormCaption
30        FillG
40        g.ColWidth(6) = 2500

50        Exit Sub

optCyto_Click_Error:

          Dim strES As String
          Dim intEL As Integer

60        intEL = Erl
70        strES = Err.Description
80        LogError "frmDaily", "optCyto_Click", intEL, strES

End Sub

Private Sub optEnd_Click()
10        On Error GoTo optEnd_Click_Error

20        SetFormCaption
30        FillG


40        Exit Sub

optEnd_Click_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmDaily", "optEnd_Click", intEL, strES

End Sub

Private Sub optExt_Click()

10        On Error GoTo optExt_Click_Error

20        SetFormCaption
30        FillG


40        Exit Sub

optExt_Click_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmDaily", "optExt_Click", intEL, strES

End Sub

Private Sub optHaem_Click()
10        On Error GoTo optHaem_Click_Error

20        SetFormCaption
30        FillG


40        Exit Sub

optHaem_Click_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmDaily", "optHaem_Click", intEL, strES

End Sub

Private Sub optHisto_Click()
10        On Error GoTo optHisto_Click_Error

20        SetFormCaption
30        FillG
40        g.ColWidth(6) = 2500

50        Exit Sub

optHisto_Click_Error:

          Dim strES As String
          Dim intEL As Integer

60        intEL = Erl
70        strES = Err.Description
80        LogError "frmDaily", "optHisto_Click", intEL, strES

End Sub

Private Sub optImm_Click()
10        On Error GoTo optImm_Click_Error

20        SetFormCaption
30        FillG


40        Exit Sub

optImm_Click_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmDaily", "optImm_Click", intEL, strES

End Sub

Private Sub optMicro_Click()
10        On Error GoTo optMicro_Click_Error

20        SetFormCaption
30        FillG
40        g.ColWidth(6) = 2500

50        Exit Sub

optMicro_Click_Error:

          Dim strES As String
          Dim intEL As Integer

60        intEL = Erl
70        strES = Err.Description
80        LogError "frmDaily", "optMicro_Click", intEL, strES

End Sub

Private Sub optNotPrinted_Click()

10        On Error GoTo optNotPrinted_Click_Error

20        If optExt Then
30            iMsg "Only unvalidated samples can be searched for externals"
40            optNotValidated.Value = True
50        End If

60        Exit Sub

optNotPrinted_Click_Error:

          Dim strES As String
          Dim intEL As Integer

70        intEL = Erl
80        strES = Err.Description
90        LogError "frmNotValidatedPrinted", "optNotPrinted_Click", intEL, strES

End Sub

Private Sub optNotValidatedPrinted_Click()

10        On Error GoTo optNotValidatedPrinted_Click_Error
20        If optExt Then
30            iMsg "Only unvalidated samples can be searched for externals"
40            optNotValidated.Value = True
50        End If


60        Exit Sub

optNotValidatedPrinted_Click_Error:

          Dim strES As String
          Dim intEL As Integer

70        intEL = Erl
80        strES = Err.Description
90        LogError "frmNotValidatedPrinted", "optNotValidatedPrinted_Click", intEL, strES

End Sub

Private Sub optSemen_Click()
10        On Error GoTo optSemen_Click_Error

20        SetFormCaption
30        FillG


40        Exit Sub

optSemen_Click_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmDaily", "optSemen_Click", intEL, strES

End Sub
