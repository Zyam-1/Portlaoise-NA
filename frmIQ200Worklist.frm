VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frmIQ200Worklist 
   Caption         =   "NetAcquire - Urine Worklist"
   ClientHeight    =   7170
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11970
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7170
   ScaleWidth      =   11970
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExcel 
      Caption         =   "Export to Excel"
      Height          =   1155
      Left            =   10740
      Picture         =   "frmIQ200Worklist.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3570
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   1170
      Left            =   10740
      Picture         =   "frmIQ200Worklist.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   12
      Tag             =   "cancel"
      ToolTipText     =   "Exit"
      Top             =   5820
      Width           =   1035
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   1170
      Left            =   10740
      Picture         =   "frmIQ200Worklist.frx":11D4
      Style           =   1  'Graphical
      TabIndex        =   11
      Tag             =   "cancel"
      ToolTipText     =   "Print"
      Top             =   1860
      Width           =   1035
   End
   Begin Threed.SSPanel panSampleDates 
      Height          =   1425
      Left            =   3390
      TabIndex        =   0
      Top             =   210
      Width           =   7095
      _Version        =   65536
      _ExtentX        =   12515
      _ExtentY        =   2514
      _StockProps     =   15
      Caption         =   "Between Sample Dates"
      ForeColor       =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelInner      =   1
      Alignment       =   0
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Calculate"
         Default         =   -1  'True
         Height          =   1170
         Left            =   5640
         Picture         =   "frmIQ200Worklist.frx":209E
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Search"
         Top             =   120
         Width           =   1035
      End
      Begin MSComCtl2.DTPicker calfrom 
         Height          =   345
         Left            =   210
         TabIndex        =   1
         Top             =   600
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   194183169
         CurrentDate     =   37753
      End
      Begin MSComCtl2.DTPicker calTo 
         Height          =   345
         Left            =   2760
         TabIndex        =   2
         Top             =   600
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   194183169
         CurrentDate     =   37753
      End
      Begin MSMask.MaskEdBox tFromTime 
         Height          =   300
         Left            =   1770
         TabIndex        =   3
         Tag             =   "SampleTime"
         Top             =   645
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox tToTime 
         Height          =   300
         Left            =   4320
         TabIndex        =   4
         Tag             =   "SampleTime"
         Top             =   645
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   195
         TabIndex        =   6
         Top             =   330
         Width           =   945
      End
      Begin VB.Label Label1 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   2745
         TabIndex        =   5
         Top             =   330
         Width           =   945
      End
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   1425
      Left            =   240
      TabIndex        =   7
      Top             =   210
      Width           =   3045
      _Version        =   65536
      _ExtentX        =   5371
      _ExtentY        =   2514
      _StockProps     =   15
      Caption         =   "Display"
      ForeColor       =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelInner      =   1
      Alignment       =   0
      Begin VB.OptionButton optDisplay 
         Caption         =   "All work-list samples"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   120
         TabIndex        =   10
         Top             =   1020
         Width           =   2670
      End
      Begin VB.OptionButton optDisplay 
         Caption         =   "Previous work-list samples"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   690
         Width           =   2880
      End
      Begin VB.OptionButton optDisplay 
         Caption         =   "New work-list samples"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   375
         Value           =   -1  'True
         Width           =   2715
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdReport 
      Height          =   5145
      Left            =   240
      TabIndex        =   13
      Top             =   1830
      Width           =   10320
      _ExtentX        =   18203
      _ExtentY        =   9075
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
      BackColor       =   -2147483624
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483634
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      AllowUserResizing=   1
      FormatString    =   $"frmIQ200Worklist.frx":2F68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblExcelInfo 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exporting..."
      Height          =   345
      Left            =   10620
      TabIndex        =   16
      Top             =   4710
      Visible         =   0   'False
      Width           =   1305
   End
End
Attribute VB_Name = "frmIQ200Worklist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

10        Unload Me

End Sub


Private Sub cmdExcel_Click()
          Dim FromTime As String, ToTime As String
          Dim strHeading As String

10        FromTime = Format(calFrom, "dd/mmm/yyyy") & " " & tFromTime
20        ToTime = Format(calTo, "dd/mmm/yyyy") & " " & tToTime


30        strHeading = "Regional Hospital " & StrConv(HospName(0), vbProperCase) & " : Microbiology " & vbCr & _
                       "IQ200 WorkList Report" & vbCr & FromTime & " To " & ToTime & vbCr

40        ExportFlexGrid grdReport, Me, strHeading

End Sub

Private Sub cmdSearch_Click()

      Dim tb As Recordset
      Dim sn As Recordset
      Dim sql As String
      Dim FromTime As String, ToTime As String
      Dim s As String
      Dim SampleID As String
      Dim Rundate As String
      Dim Result() As String
      Dim tempResult As String
      Dim i As Integer
      Dim AddToCultureList As Boolean
      Dim AddSampleIDOnly As Boolean

      Dim X As Integer

10    On Error GoTo cmdSearch_Click_Error

20    AddToCultureList = False
30    AddSampleIDOnly = False
      '30    AddToWBC = False
      '40    AddToPC = False
      '50    AddtoBact = False
40    ClearGrid

50    ReDim Result(0 To 5)

60    If tFromTime = "__:__" Then
70        tFromTime = "00:00"
80    End If

90    If tToTime = "__:__" Then
100       tToTime = "00:00"
110   End If

120   FromTime = Format(calFrom, "dd/mmm/yyyy") & " " & tFromTime
130   ToTime = Format(calTo, "dd/mmm/yyyy") & " " & tToTime

140   If DateDiff("d", FromTime, ToTime) < 0 Then
150       iMsg "From Date must be before To Date!", vbExclamation
160   End If

170   sql = "SELECT DISTINCT I.SampleId FROM IQ200 I " & _
            "INNER JOIN Demographics D ON I.SampleId = D.SampleId " & _
            "WHERE I.DateTimeOfRecord BETWEEN '" & FromTime & "' AND '" & ToTime & "' " & _
            "AND D.Valid = 1"

180   If optDisplay(0) Then
190       sql = sql & " AND WorkListPrinted = 0"
200   ElseIf optDisplay(1) Then
210       sql = sql & " AND WorkListPrinted = 1"
220   End If
230   Set tb = New Recordset
240   RecOpenServer 0, tb, sql

250   Do While Not tb.EOF
260       For i = 0 To 5
270           Result(i) = ""
280       Next
290       sql = "SELECT I.SampleId, [ShortName],[LongName],[Result], I.DateTimeOfRecord, " & _
                "(DATEDIFF(dd, D.DOB, I.DateTimeOfRecord)  / 365) Age, " & _
                "COALESCE(D.Pregnant, 0) Preg, D.Ward FROM IQ200 I " & _
                "INNER JOIN Demographics D ON I.SampleId = D.SampleId " & _
                "WHERE I.SampleId = '" & tb!SampleID & "'"
300       Set sn = New Recordset
310       RecOpenServer 0, sn, sql

320       SampleID = Val(tb!SampleID & "") - SysOptMicroOffset(0)

330       If Not sn.EOF Then
340           If (sn!Ward & "" = "OHIU" Or sn!Ward & "" = "ROHDU" Or sn!Ward & "" = "Oncology OPD" _
                  Or sn!Ward & "" = "Ante-natal Clinic" Or sn!Ward & "" = "Haematology OPD") Then
350               AddToCultureList = True
360           End If
370           Result(0) = sn!Ward & ""

380           If Val(sn!Age & "") < 16 Then
390               AddToCultureList = True
400           End If
410           Result(1) = sn!Age & ""

420           If sn!Preg = 0 Then
430               Result(5) = "No"
440           Else
450               Result(5) = "Yes"
460               AddToCultureList = True
470           End If
480           Rundate = Format(sn!DateTimeOfRecord & "", "dd/mm/yyyy")
490           Do While Not sn.EOF
500               X = InStr(sn!Result & "", " ") - 1
510               If X > 0 Then
520                   tempResult = Left(sn!Result & "", X)
530               Else
540                   tempResult = sn!Result & ""
550               End If

560               If UCase$(sn!ShortName & "") = "WBC" Then
570                   If IsNumeric(tempResult) Then
580                       If Val(tempResult) > 25 Then
590                           AddToCultureList = True
600                       Else
610                           AddSampleIDOnly = True
620                       End If
630                   End If
640                   Result(2) = sn!Result & ""

650               ElseIf sn!ShortName & "" = "BACT" Then
660                   If IsNumeric(tempResult) Then
670                       If Val(tempResult) > 2 Then
680                           AddToCultureList = True
690                       Else
700                           AddSampleIDOnly = True
710                       End If
720                   End If
730                   Result(3) = sn!Result & ""

740               ElseIf sn!ShortName & "" = "PC" Then
750                   If IsNumeric(tempResult) Then
760                       If Val(tempResult) > 8500 Then
770                           AddToCultureList = True
780                       Else
790                           AddSampleIDOnly = True
800                       End If
810                   End If
820                   Result(4) = sn!Result & ""

830               End If

840               sn.MoveNext
850           Loop
860           For i = 2 To 4
870               If Result(i) = "" Then
880                   Result(i) = "0"
890               End If
900           Next
910       End If

920       If AddToCultureList = True Then
930           s = SampleID & vbTab & Rundate & vbTab & Result(2) & vbTab & Result(3) & vbTab & _
                  Result(4) & vbTab & Result(1) & vbTab & Result(5) & vbTab & Result(0)
940           grdReport.AddItem s
950       ElseIf AddSampleIDOnly = True Then
960           s = SampleID & vbTab & Rundate & vbTab & "" & vbTab & "" & vbTab & _
                  "" & vbTab & Result(1) & vbTab & Result(5) & vbTab & Result(0)
970           grdReport.AddItem s
980       End If
990       AddToCultureList = False
1000      AddSampleIDOnly = False

1010      tb.MoveNext
1020  Loop

1030  If grdReport.Rows > 2 And grdReport.TextMatrix(1, 0) = "" Then
1040      grdReport.RemoveItem 1
1050  End If

1060  Exit Sub

cmdSearch_Click_Error:

      Dim strES As String
      Dim intEL As Integer

1070  intEL = Erl
1080  strES = Err.Description
1090  LogError "frmIQ200Worklist", "cmdSearch_Click", intEL, strES, sql

End Sub

Private Sub Form_Load()

10        calTo = Format(Now, "dd/mmm/yyyy")
20        calFrom = Format(Now, "dd/mmm/yyyy")
30        tFromTime = "00:00"
40        tToTime = "23:59"

End Sub

Private Sub ClearGrid()

10        grdReport.AddItem "", 1
20        grdReport.Rows = 2

End Sub


Private Sub cmdPrint_Click()
          Dim n As Integer
          Dim X As Integer
          Dim sql As String
          Dim FromTime As String
          Dim ToTime As String

          'Heading

10        On Error GoTo cmdPrint_Click_Error

20        FromTime = Format(calFrom, "dd/mmm/yyyy") & " " & tFromTime
30        ToTime = Format(calTo, "dd/mmm/yyyy") & " " & tToTime

40        Printer.Font.Name = "Courier New"
50        Printer.Font.Size = 14
60        Printer.Font.Bold = True

70        Printer.ForeColor = vbRed
80        Printer.Print FormatString("Regional Hospital " & StrConv(HospName(0), vbProperCase) & " : Microbiology ", 70, , AlignCenter)

90        Printer.Print FormatString("IQ200 WorkList Report", 70, , AlignCenter)
100       Printer.Print FormatString(FromTime & " To " & ToTime, 70, , AlignCenter)
110       Printer.CurrentY = 1000

120       Printer.Font.Size = 4
130       Printer.Print Space(3);
140       Printer.Print String(240, "-")
150       Printer.CurrentY = Printer.CurrentY + 150


160       Printer.ForeColor = vbBlack

170       Printer.Font.Name = "Courier New"
180       Printer.Font.Size = 10
190       Printer.Font.Bold = False
          'End of heading


200       Printer.Print

210       For n = 0 To grdReport.Rows - 1
220           grdReport.row = n
230           If n = 0 Then
240               Printer.Font.Bold = True
250           Else
260               Printer.Font.Bold = False
270           End If
280           Printer.Print FormatString(" ", 2, , AlignLeft);
290           For X = 0 To 7
300               grdReport.Col = X

310               If X = 0 Or X = 4 Then
320                   Printer.Print FormatString(grdReport, 11, , AlignLeft);
330               ElseIf X = 1 Then
340                   Printer.Print FormatString(grdReport, 15, , AlignLeft);
350               ElseIf X = 5 Then
360                   Printer.Print FormatString(grdReport, 5, , AlignLeft);
370               ElseIf X = 7 Then
380                   Printer.Print FormatString(grdReport, 25, , AlignLeft);
390               Else
400                   Printer.Print FormatString(grdReport, 10, , AlignLeft);
410               End If
420           Next
430           Printer.Print
440           If n > 0 Then
450               If grdReport.TextMatrix(n, 0) <> "" Then
460                   sql = "UPDATE IQ200 set WorkListPrinted = '1' where SampleId = '" & grdReport.TextMatrix(n, 0) + SysOptMicroOffset(0) & "' "
470                   Cnxn(0).Execute sql
480               End If
490           End If
500       Next

510       Printer.Print

520       Printer.EndDoc

530       Exit Sub

cmdPrint_Click_Error:

          Dim strES As String
          Dim intEL As Integer

540       intEL = Erl
550       strES = Err.Description
560       LogError "frmIQ200Worklist", "cmdPrint_Click", intEL, strES

End Sub
