VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmRFT 
   Caption         =   "Netacquire - Report Viewer"
   ClientHeight    =   10695
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13725
   Icon            =   "frmRFTs.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   18.865
   ScaleMode       =   7  'Centimeter
   ScaleWidth      =   24.209
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdUpdateNotes 
      DisabledPicture =   "frmRFTs.frx":030A
      Enabled         =   0   'False
      Height          =   465
      Left            =   13020
      Picture         =   "frmRFTs.frx":09F4
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Update report notes"
      Top             =   1350
      Width           =   555
   End
   Begin VB.TextBox txtNotes 
      Height          =   525
      Left            =   1200
      MaxLength       =   1000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   1320
      Width           =   11775
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "Move Report"
      Height          =   945
      Left            =   9480
      TabIndex        =   10
      Top             =   210
      Width           =   735
   End
   Begin MSFlexGridLib.MSFlexGrid grdPTimes 
      Height          =   1125
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   5985
      _ExtentX        =   10557
      _ExtentY        =   1984
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      FormatString    =   "<Printed Time                 |<Printed By                      |<Printer                                   "
   End
   Begin VB.Frame fraPage 
      Caption         =   "Viewing Page"
      Height          =   1125
      Left            =   6210
      TabIndex        =   3
      Top             =   90
      Width           =   1635
      Begin ComCtl2.UpDown udPage 
         Height          =   285
         Left            =   510
         TabIndex        =   7
         Top             =   750
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   503
         _Version        =   327681
         Value           =   1
         BuddyControl    =   "lblCurrentPage"
         BuddyDispid     =   196615
         OrigLeft        =   25
         OrigTop         =   8
         OrigRight       =   25
         OrigBottom      =   9
         Max             =   99
         Min             =   1
         Orientation     =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label lblTotalPages 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         Height          =   285
         Left            =   930
         TabIndex        =   6
         Top             =   390
         Width           =   465
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "of"
         Height          =   195
         Left            =   720
         TabIndex        =   5
         Top             =   420
         Width           =   135
      End
      Begin VB.Label lblCurrentPage 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         Height          =   285
         Left            =   210
         TabIndex        =   4
         Top             =   390
         Width           =   465
      End
   End
   Begin RichTextLib.RichTextBox rtb 
      Height          =   8715
      Left            =   120
      TabIndex        =   2
      Top             =   1890
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   15372
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmRFTs.frx":10DE
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "  Re-Print this page"
      Height          =   945
      Left            =   10410
      Picture         =   "frmRFTs.frx":1160
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Re Print already Printed Results"
      Top             =   210
      Width           =   1545
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   945
      Left            =   12030
      Picture         =   "frmRFTs.frx":146A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   210
      Width           =   1545
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Report Notes"
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   1455
      Width           =   945
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sample ID 999999999 No Reports"
      Height          =   1125
      Left            =   7920
      TabIndex        =   8
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmRFT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mSampleID As Double
Private mDept As String


Private Sub HighlightRow()

          Dim x As Integer
          Dim Y As Integer
          Dim ySave As Integer

10        With grdPTimes
20            ySave = .Row

30            .Col = 0
40            For Y = 1 To .Rows - 1
50                .Row = Y
60                If .CellBackColor = vbYellow Then
70                    For x = 0 To .Cols - 1
80                        .Col = x
90                        .CellBackColor = 0
100                   Next
110                   Exit For
120               End If
130           Next

140           .Row = ySave
150           For x = 0 To .Cols - 1
160               .Col = x
170               .CellBackColor = vbYellow
180           Next

190       End With

End Sub

Private Function PagesPerReport(ByVal pTime As String) As Integer

          Dim sql As String
          Dim tb As Recordset

10        On Error GoTo PagesPerReport_Error

20        sql = "SELECT COUNT(*) Tot FROM Reports WHERE " & _
                "SampleID = " & mSampleID & " " & _
                "AND PrintTime = '" & Format(pTime, "dd/MMM/yyyy HH:mm:ss") & "'"
          '            "AND Dept = '" & mDept & "' "
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        PagesPerReport = tb!Tot

60        Exit Function

PagesPerReport_Error:

          Dim strES As String
          Dim intEL As Integer

70        intEL = Erl
80        strES = Err.Description
90        LogError "frmRFT", "PagesPerReport", intEL, strES, sql

End Function

Private Sub cmdExit_Click()

10        Unload Me

End Sub

Private Sub cmdMove_Click()

          Dim sql As String

10        On Error GoTo cmdMove_Click_Error

20        If grdPTimes.TextMatrix(grdPTimes.Row, 0) = "" Then
30            iMsg "No print time available. Report cannot be moved"
40            Exit Sub
50        End If

60        If iMsg("Are you sure you want to move this report to Immunology department", vbYesNo) = vbYes Then
70            sql = "Update Reports Set Dept = 'I' " & _
                    "Where PrintTime = '" & Format(grdPTimes.TextMatrix(grdPTimes.Row, 0), "yyyy-MM-dd hh:mm:ss") & "' " & _
                    "And Dept = 'E' " & _
                    "And SampleID = " & mSampleID
80            Cnxn(0).Execute sql
90            iMsg "Report successfully moved"
100           If grdPTimes.Rows = 2 Then
110               grdPTimes.AddItem ""
120           End If
130           grdPTimes.RemoveItem grdPTimes.Row
140           HighlightRow
150           FillReport
160       End If

170       Exit Sub

cmdMove_Click_Error:

          Dim strES As String
          Dim intEL As Integer

180       intEL = Erl
190       strES = Err.Description
200       LogError "frmRFT", "cmdMove_Click", intEL, strES, sql

End Sub

Private Sub cmdPrint_Click()

          Dim OriginalPrinter As String
          Dim FullDept As String
          Dim sql As String

10        On Error GoTo cmdPrint_Click_Error

20        Select Case mDept
          Case "B", "R", "T", "S", "G", "Q": FullDept = "CHBIO"
30        Case "C", "M": FullDept = "CHCOAG"
40        Case "X": FullDept = "CHEXT"
50        Case "H", "K": FullDept = "CHHAEM"
60        Case "I", "J": FullDept = "CHIMM"
70        Case "E": FullDept = "CHEND"
80        Case "Y": FullDept = "CHHIST"
90        Case "F", "N", "U", "Z": FullDept = "CHMICRO"
100       Case Else: Exit Sub
110       End Select

120       OriginalPrinter = Printer.DeviceName
130       rtb.SelStart = 0
140       rtb.SelLength = 10000000#
150       rtb.SelPrint Printer.hDC

          'xFound = False
          'sql = "SELECT * FROM Printers WHERE MappedTo = '" & FullDept & "'"
          'Set tb = New Recordset
          'RecOpenClient 0, tb, sql
          'If Not tb.EOF Then
          '  TargetPrinter = UCase$(tb!PrinterName & "")
          '  For Each Px In Printers
          '    If UCase(Px.DeviceName) = TargetPrinter Then
          '      Set Printer = Px
          '      xFound = True
          '      Exit For
          '    End If
          '  Next
          '  If xFound Then
          '
          '    rtb.SelStart = 0
          '    rtb.SelLength = 10000000#
          '    rtb.SelPrint Printer.hDC
          '
          '  End If
          '
          'End If

160       Exit Sub

cmdPrint_Click_Error:

          Dim strES As String
          Dim intEL As Integer

170       intEL = Erl
180       strES = Err.Description
190       LogError "frmRFT", "cmdPrint_Click", intEL, strES, sql

End Sub


Private Sub Form_Load_old()

          Dim tb As Recordset
          Dim sql As String
          Dim TotalPages As Integer

10        On Error GoTo Form_Load_old_Error

20        sql = "SELECT COUNT(*) AS Tot FROM Reports WHERE " & _
                "SampleID = " & mSampleID & " and Dept = '" & mDept & "'"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        TotalPages = tb!Tot
60        If TotalPages > 0 Then
70            fraPage.Visible = True
80            lblInfo.Visible = False
90            lblTotalPages = TotalPages
100           lblCurrentPage = "1"
110           udPage.Max = TotalPages
120       Else
130           fraPage.Visible = False
140           lblInfo.Visible = True
150           lblInfo = "Sample ID" & vbCrLf & _
                        mSampleID & vbCrLf & _
                        "No Reports"
160       End If

170       FillReport

180       Exit Sub

Form_Load_old_Error:

          Dim strES As String
          Dim intEL As Integer

190       intEL = Erl
200       strES = Err.Description
210       LogError "frmRFT", "Form_Load_old", intEL, strES, sql

End Sub

Private Sub cmdUpdateNotes_Click()

          Dim sql As String

10        On Error GoTo cmdUpdateNotes_Click_Error

20        If grdPTimes.TextMatrix(grdPTimes.Row, 0) = "" Then Exit Sub

30        sql = "UPDATE Reports Set Notes = '%notes' WHERE SampleID = %sampleid AND PrintTime = '%printtime'"
40        sql = Replace(sql, "%notes", txtNotes)
50        sql = Replace(sql, "%sampleid", mSampleID)
60        sql = Replace(sql, "%printtime", Format(grdPTimes.TextMatrix(grdPTimes.Row, 0), "yyyy-MM-dd hh:mm:ss"))

70        Cnxn(0).Execute sql
80        cmdUpdateNotes.Enabled = False

90        Exit Sub

cmdUpdateNotes_Click_Error:

          Dim strES As String
          Dim intEL As Integer

100       intEL = Erl
110       strES = Err.Description
120       LogError "frmRFT", "cmdUpdateNotes_Click", intEL, strES, sql

End Sub

Private Sub Form_Load()

          Dim tb As Recordset
          Dim tb1 As Recordset
          Dim sql As String
          Dim TotalReports As Integer
          Dim Y As Integer
          Dim Target As String

10        On Error GoTo Form_Load_Error

          '***************************************
          'Author: Babar Shahzad
          'BLR:immunology allergy reports were saved in endrocrinology department
          'all those reports need to be moved to immunology department
          'this can only be done step by step. (so move button is there)
          'this button will only be visible for Endrocrinology department
20        If (mDept <> "E") Or (UCase$(UserName) <> "NUALA HENNESSY") Then cmdMove.Visible = False
          '***************************************


30        With grdPTimes
40            .Rows = 2
50            .AddItem ""
60            .RemoveItem 1

70            sql = "SELECT PrintTime FROM Reports WHERE " & _
                    "SampleID = " & mSampleID & " " & _
                    "ORDER BY PrintTime DESC"
              '              "AND Dept = '" & mDept & "' "
80            Set tb = New Recordset
90            RecOpenServer 0, tb, sql
100           If Not tb.EOF Then
110               Do While Not tb.EOF
120                   .AddItem Format(tb!printTime, "dd/MM/yy HH:mm:ss")
130                   tb.MoveNext
140               Loop
150               If .Rows > 2 Then
160                   .RemoveItem 1
170               End If
180               If .Rows > 2 Then
190                   Target = .TextMatrix(1, 0)
200                   For Y = 2 To .Rows - 1
210                       If DateDiff("s", .TextMatrix(Y, 0), Target) < 10 And .TextMatrix(Y, 0) <> Target Then
220                           sql = "UPDATE Reports SET PrintTime = '" & Format(Target, "dd/MMM/yyyy HH:mm:ss") & "' " & _
                                    "WHERE SampleID = " & mSampleID & " " & _
                                    "AND PrintTime = '" & Format(.TextMatrix(Y, 0), "dd/MMM/yyyy HH:mm:ss") & "' " & _
                                    "AND Dept = '" & mDept & "'"
230                           Set tb1 = New Recordset
240                           RecOpenServer 0, tb1, sql
250                       Else
260                           Target = .TextMatrix(Y, 0)
270                       End If
280                   Next
290               End If
300           End If

310           .Rows = 2
320           .AddItem ""
330           .RemoveItem 1

340           sql = "SELECT DISTINCT(PrintTime), Initiator, Printer FROM Reports WHERE " & _
                    "SampleID = " & mSampleID & " " & _
                    "ORDER BY PrintTime DESC"
              '              "AND Dept = '" & mDept & "' "
350           Set tb = New Recordset
360           RecOpenServer 0, tb, sql
370           Do While Not tb.EOF
380               .AddItem Format(tb!printTime, "dd/MM/yy HH:mm:ss") & vbTab & _
                           tb!Initiator & vbTab & _
                           tb!Printer & ""
390               tb.MoveNext
400           Loop
410           If .Rows > 2 Then
420               .RemoveItem 1
430           End If
440       End With

450       TotalReports = grdPTimes.Rows - 1
460       If TotalReports > 0 Then
470           grdPTimes.Row = 1
480           HighlightRow
490           fraPage.Visible = True
500           lblInfo.Visible = False
510           lblTotalPages = PagesPerReport(grdPTimes.TextMatrix(1, 0))
520           lblCurrentPage = "1"
530           udPage.Max = lblTotalPages
540           FillReport
550       Else
560           fraPage.Visible = False
570           lblInfo.Visible = True
580           lblInfo = "Sample ID" & vbCrLf & _
                        mSampleID & vbCrLf & _
                        "No Reports"
590       End If

600       Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer



610       intEL = Erl
620       strES = Err.Description
630       LogError "frmRFT", "Form_Load", intEL, strES, sql

End Sub

Private Sub FillReport()

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo FillReport_Error

20        rtb = ""
30        rtb.SelText = ""

40        If Val(lblCurrentPage) = 0 Then Exit Sub

50        sql = "SELECT Report, Notes FROM Reports WHERE " & _
                "SampleID = " & mSampleID & " " & _
                "AND PrintTime = '" & Format(grdPTimes.TextMatrix(grdPTimes.Row, 0), "dd/MMM/yyyy HH:mm:ss") & "' " & _
                "AND PageNumber = '" & Val(lblCurrentPage) - 1 & "'"
          '"AND Dept = '" & mDept & "' " '
60        Set tb = New Recordset
70        RecOpenServer 0, tb, sql
80        If Not tb.EOF Then
90            If Trim(tb!Report & "") <> "" Then
100               rtb.SelText = Trim(tb!Report)
110           End If
120           If Trim(tb!Notes & "") <> "" Then
130               txtNotes = Trim(tb!Notes)
140           Else
150               txtNotes = ""
160           End If
170       End If

180       Exit Sub

FillReport_Error:

          Dim strES As String
          Dim intEL As Integer

190       intEL = Erl
200       strES = Err.Description
210       LogError "frmRFT", "FillReport", intEL, strES, sql

End Sub

Public Property Let SampleID(ByVal SID As Double)

10        On Error GoTo SampleID_Error

20        mSampleID = SID

30        Exit Property

SampleID_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmRFT", "SampleID", intEL, strES

End Property

Public Property Let Dept(ByVal Dep As String)

10        On Error GoTo Dept_Error

20        mDept = Dep

30        Exit Property

Dept_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmRFT", "Dept", intEL, strES


End Property

Private Sub grdPTimes_Click()

10        HighlightRow
20        lblTotalPages = PagesPerReport(grdPTimes.TextMatrix(grdPTimes.Row, 0))
30        lblCurrentPage = "1"
40        udPage.Max = lblTotalPages
50        cmdUpdateNotes.Enabled = False
60        FillReport

End Sub

Private Sub txtNotes_KeyPress(KeyAscii As Integer)
10        If grdPTimes.TextMatrix(grdPTimes.Row, 0) = "" Then
20            KeyAscii = 0
30        Else
40            cmdUpdateNotes.Enabled = True
50        End If
End Sub

Private Sub UdPage_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

10        FillReport

End Sub


