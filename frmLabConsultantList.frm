VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmLabConsultantList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NetAcquire --- Microbiology Samples for Consultant Review"
   ClientHeight    =   8670
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8670
   ScaleWidth      =   15960
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   900
      Left            =   14805
      Picture         =   "frmLabConsultantList.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      Tag             =   "bprint"
      ToolTipText     =   "Print Result"
      Top             =   1560
      Width           =   1000
   End
   Begin VB.CommandButton cmdApplyFilter 
      Caption         =   "Search"
      Height          =   1000
      Left            =   11580
      TabIndex        =   14
      Top             =   300
      Width           =   1100
   End
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   315
      Left            =   9780
      TabIndex        =   10
      Top             =   360
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
      _Version        =   393216
      Format          =   244514817
      CurrentDate     =   41850
   End
   Begin VB.Frame fraStatus 
      Caption         =   "Filter By Status"
      Height          =   1215
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   8835
      Begin VB.OptionButton optStatus 
         Caption         =   "Released by Lab"
         Height          =   255
         Index           =   4
         Left            =   3600
         TabIndex        =   17
         Top             =   720
         Width           =   2355
      End
      Begin VB.OptionButton optStatus 
         Caption         =   "All"
         Height          =   195
         Index           =   3
         Left            =   7260
         TabIndex        =   9
         Top             =   300
         Width           =   735
      End
      Begin VB.OptionButton optStatus 
         Caption         =   "In Lab for Review"
         Height          =   195
         Index           =   2
         Left            =   3600
         TabIndex        =   8
         Top             =   300
         Width           =   1815
      End
      Begin VB.OptionButton optStatus 
         Caption         =   "Released by consultant"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   2355
      End
      Begin VB.OptionButton optStatus 
         Caption         =   "With Consutant"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   300
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Enabled         =   0   'False
      Height          =   900
      Left            =   14805
      Picture         =   "frmLabConsultantList.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2640
      Width           =   1000
   End
   Begin VB.CommandButton bCancel 
      Appearance      =   0  'Flat
      Caption         =   "Exit"
      Height          =   900
      Left            =   14805
      Picture         =   "frmLabConsultantList.frx":1C8C
      Style           =   1  'Graphical
      TabIndex        =   3
      Tag             =   "101"
      ToolTipText     =   "Exit Screen"
      Top             =   6675
      Width           =   1000
   End
   Begin VB.CommandButton cmdRefresh 
      Appearance      =   0  'Flat
      Caption         =   "Refresh"
      Height          =   900
      Index           =   0
      Left            =   14805
      Picture         =   "frmLabConsultantList.frx":1F96
      Style           =   1  'Graphical
      TabIndex        =   2
      Tag             =   "101"
      ToolTipText     =   "Exit Screen"
      Top             =   3720
      Width           =   1000
   End
   Begin MSFlexGridLib.MSFlexGrid grdSID 
      Height          =   6675
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   14550
      _ExtentX        =   25665
      _ExtentY        =   11774
      _Version        =   393216
      Cols            =   11
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      FocusRect       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      SelectionMode   =   1
      FormatString    =   $"frmLabConsultantList.frx":2860
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   135
      Left            =   120
      TabIndex        =   1
      Top             =   8460
      Width           =   14550
      _ExtentX        =   25665
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   1
      Max             =   600
      Scrolling       =   1
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   315
      Left            =   9780
      TabIndex        =   11
      Top             =   900
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
      _Version        =   393216
      Format          =   244580353
      CurrentDate     =   41850
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Click on sample ID to remove from list"
      Height          =   195
      Left            =   240
      TabIndex        =   16
      Top             =   1440
      Width           =   2670
   End
   Begin VB.Label lblTo 
      AutoSize        =   -1  'True
      Caption         =   "To"
      Height          =   195
      Left            =   9360
      TabIndex        =   13
      Top             =   960
      Width           =   195
   End
   Begin VB.Label lblFrom 
      AutoSize        =   -1  'True
      Caption         =   "From"
      Height          =   195
      Left            =   9240
      TabIndex        =   12
      Top             =   420
      Width           =   345
   End
   Begin VB.Image imgGreenTick 
      Height          =   225
      Left            =   0
      Picture         =   "frmLabConsultantList.frx":292C
      Top             =   210
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image imgRedCross 
      Height          =   225
      Left            =   0
      Picture         =   "frmLabConsultantList.frx":2C02
      Top             =   0
      Visible         =   0   'False
      Width           =   210
   End
End
Attribute VB_Name = "frmLabConsultantList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private CurrentGridRow As Integer
Private PreviousGridRow As Integer

Private Sub FillGrid()

      Dim sql As String
      Dim tb As Recordset
      Dim s As String
      Dim PatInfo As String
      Dim strType As String
      Dim ShowFaecesWardEnq As Boolean
      Dim ShowUrineWardEnq As Boolean
      Dim ShowBloodCultureWardEnq As Boolean
      Dim ShowSwabWardEnq As Boolean

10    On Error GoTo FillGrid_Error

      'grdSID
20    With grdSID
30        .Rows = 2
40        .Cols = 11
50        .AddItem ""
60        .RemoveItem 1


          '.FormatString = "<SampleID   |<Run Date     |<Sample Date         |<Pat Name                               |<DOB             |<Age   |<Sex  |<Address                                                      |<   |<   "
70        .ColWidth(0) = 1240
80        .ColWidth(1) = 1700
90        .ColWidth(2) = 1000
100       .ColWidth(3) = 1600
110       .ColWidth(4) = 2500
120       .ColWidth(5) = 1000
130       .ColWidth(6) = 600
140       .ColWidth(7) = 500
150       .ColWidth(8) = 1800
          '.ColWidth(9) = 0
160       .ColWidth(9) = 500
170       .ColWidth(10) = 800




180       sql = "SELECT     D.SampleID,D.SampleDate,D.Rundate ,D.PatName,D.DoB,D.Age,D.Sex,D.Addr0, C.Status, COALESCE(c.ack, 0) ack, COALESCE(c.ConAck, 0) ConAck, S.Site, S.SiteDetails "
190       sql = sql & "FROM  ConsultantList as C INNER JOIN Demographics as D ON C.SampleID = D.SampleID "
200       sql = sql & "INNER JOIN MicroSiteDetails as S ON D.SampleID = S.SampleID "
210       sql = sql & "WHERE DateTimeOfRecord Between '" & Format(dtpFrom, "dd/MMM/yyyy hh:mm:ss") & "' AND '" & Format(dtpTo, "dd/MMM/yyyy hh:mm:ss") & "' "
220       If optStatus(0).Value = True Then
230           sql = sql & "AND C.Status = 0 "
240       ElseIf optStatus(1).Value = True Then
250           sql = sql & "AND C.Status = 1 "
260       ElseIf optStatus(2).Value = True Then
270           sql = sql & "AND C.Status = 2 "
280       ElseIf optStatus(4).Value = True Then
290           sql = sql & "AND C.Status = 3 "
300       End If
          
310       sql = sql & "Order by C.DateTimeOfRecord "

320       Set tb = New Recordset
330       RecOpenServer 0, tb, sql

340       Do While Not tb.EOF

350           PatInfo = tb!PatName & vbTab & tb!Dob & vbTab & tb!Age & vbTab & tb!sex

360           If Val(tb!SampleID & "") > SysOptMicroOffset(0) Then
370               s = Format$(Val(tb!SampleID) - SysOptMicroOffset(0)) & vbTab & _
                      tb!Site & " " & tb!SiteDetails & vbTab & _
                      tb!Rundate & vbTab
380           Else
390               s = Val(tb!SampleID) & vbTab & vbTab & _
                      tb!Rundate & vbTab
400           End If

410           If IsDate(tb!SampleDate) Then
420               If Format(tb!SampleDate, "hh:mm") <> "00:00" Then
430                   s = s & Format(tb!SampleDate, "dd/MM/yyyy hh:mm")
440               Else
450                   s = s & Format(tb!SampleDate, "dd/MM/yyyy")
460               End If
470           Else
480               s = s & "Not Specified"
490           End If
500           s = s & vbTab & PatInfo
510           Select Case tb!Status
              Case 0:
520               s = s & vbTab & "With Consultant"
530           Case 1, 3:
540               s = s & vbTab & "Released to Ward"
550           Case 2:
560               s = s & vbTab & "In Lab for Review"
570           End Select
             '----------- by farhan--------------
580               If tb!Ack = 0 Then
590                 s = s & vbTab & "No"
600               ElseIf tb!Ack = 1 Then
610                 s = s & vbTab & "Yes"
620               End If
                  
630               If tb!ConAck = 0 Then
640                 s = s & vbTab & "No"
650               ElseIf tb!ConAck = 1 Then
660                 s = s & vbTab & "Yes"
670               End If
              '=================================
              

              'strType = LoadOutstandingMicro(tb!SampleID, 0)
              '        s = s & strType


680           .AddItem s
690           .row = .Rows - 1
      '580           .Col = 9
      '590           Set .CellPicture = imgRedCross.Picture

700           tb.MoveNext
710       Loop

720       If .Rows > 2 Then
730           .RemoveItem 1
740       End If

750   End With

      'Call ShowSignals(grdSID, ConIndex)
760   Exit Sub

FillGrid_Error:

      Dim strES As String
      Dim intEL As Integer

770   intEL = Erl
780   strES = Err.Description
790   LogError "frmViewResults", "FillGrid", intEL, strES, sql


End Sub
Private Function LoadOutstandingMicro(ByVal SampleIDWithOffset As String, ConIndex As String) As String

          Dim tb As Recordset
          Dim sql As String
          Dim s As String
          Dim n As Integer

10        On Error GoTo LoadOutstandingMicro_Error

20        sql = "Select * from MicroSiteDetails where " & _
                "SampleID = '" & SampleIDWithOffset & "' "
30        Set tb = New Recordset
40        RecOpenServer Val(ConIndex), tb, sql

50        If Not tb.EOF Then

60            If UCase(Trim$(tb!Site & "")) = "FAECES" Then
70                LoadOutstandingMicro = UCase(Trim$(tb!Site & ""))
80                Exit Function
90            End If

100           s = tb!Site & " " & tb!SiteDetails & " "
110           If tb!Site & "" = "Urine" Or tb!Site & "" = "Faeces" Then
120               sql = "Select * from MicroRequests where " & _
                        "SampleID = '" & SampleIDWithOffset & "'"
130               Set tb = New Recordset
140               RecOpenServer Val(ConIndex), tb, sql

150               If Not tb.EOF Then

160                   For n = 0 To 2
170                       If tb!Faecal And 2 ^ n Then
180                           s = s & Choose(n + 1, "C & S ", "C. Difficile ", "O/P ")
190                       End If
200                   Next

210                   For n = 3 To 5
220                       If tb!Faecal And 2 ^ n Then
230                           s = s & "Occult Blood "
240                           Exit For
250                       End If
260                   Next

270                   If tb!Faecal And 2 ^ 6 Then
280                       s = s & "Rota/Adeno "
290                   End If

300                   For n = 7 To 10
310                       If tb!Faecal And 2 ^ n Then
320                           s = s & Choose(n + 1, "Toxin A ", "Coli 0157 ", _
                                             "E/P Coli ", "S/S Screen ")
330                       End If
340                   Next

350                   For n = 0 To 5
360                       If tb!Urine And 2 ^ n Then
370                           s = s & Choose(n + 1, "C & S", "Pregnancy ", "Fat Globules ", _
                                             "Bence Jones ", "SG ", "HCG ")
380                       End If
390                   Next
400               End If
410           End If
420       End If
430       LoadOutstandingMicro = Trim(s)

440       Exit Function

LoadOutstandingMicro_Error:

          Dim strES As String
          Dim intEL As Integer

450       intEL = Erl
460       strES = Err.Description
470       LogError "frmViewResults", "LoadOutstandingMicro", intEL, strES, sql


End Function

Private Sub bcancel_Click()

10        On Error GoTo bCancel_Click_Error

20        Unload Me

30        Exit Sub

bCancel_Click_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmLabConsultantList", "bCancel_Click", intEL, strES

End Sub

Private Sub cmdApplyFilter_Click()
10    FillGrid
End Sub

Private Sub cmdPrint_Click()
      Dim Y As Integer
      Dim Px As Printer
      Dim OriginalPrinter As String
      Dim i As Integer

10    On Error GoTo cmdPrint_Click_Error

20    Screen.MousePointer = vbHourglass

30    OriginalPrinter = Printer.DeviceName



40    Printer.FontName = "Courier New"
50    Printer.Orientation = vbPRORLandscape


      '****Report heading
60    Printer.FontSize = 10
70    Printer.Font.Bold = True
80    Printer.Print
90    Printer.Print FormatString("Micro Worklist", 108, , AlignCenter)

      '****Report body heading

100   Printer.Font.Size = 9
110   For i = 1 To 135
120       Printer.Print "-";
130   Next i
140   Printer.Print

150   For Y = 0 To grdSID.Rows - 1
160       PrintText FormatString("", 0, "|"), 9, IIf(Y = 0, True, False)
170       PrintText FormatString(grdSID.TextMatrix(Y, 0), 12, "|", AlignCenter), 9, IIf(Y = 0, True, False)
180       PrintText FormatString(grdSID.TextMatrix(Y, 1), 20, "|", AlignLeft), 9, IIf(Y = 0, True, False)
190       PrintText FormatString(grdSID.TextMatrix(Y, 2), 11, "|", AlignLeft), 9, IIf(Y = 0, True, False)
200       PrintText FormatString(grdSID.TextMatrix(Y, 3), 16, "|", AlignLeft), 9, IIf(Y = 0, True, False)
210       PrintText FormatString(grdSID.TextMatrix(Y, 4), 26, "|", AlignLeft), 9, IIf(Y = 0, True, False)
220       PrintText FormatString(grdSID.TextMatrix(Y, 5), 11, "|", AlignLeft), 9, IIf(Y = 0, True, False)
230       PrintText FormatString(grdSID.TextMatrix(Y, 6), 6, "|", AlignLeft), 9, IIf(Y = 0, True, False)
240       PrintText FormatString(grdSID.TextMatrix(Y, 7), 5, "|", AlignLeft), 9, IIf(Y = 0, True, False)
250       PrintText FormatString(grdSID.TextMatrix(Y, 8), 18, "|", AlignLeft), 9, IIf(Y = 0, True, False), , , , True
260   Next

270   Printer.EndDoc

280   Screen.MousePointer = vbDefault


290   Exit Sub

cmdPrint_Click_Error:

       Dim strES As String
       Dim intEL As Integer

300    intEL = Erl
310    strES = Err.Description
320    LogError "frmLabConsultantList", "cmdPrint_Click", intEL, strES
          

End Sub

Private Sub cmdRefresh_Click(Index As Integer)

10        On Error GoTo cmdRefresh_Click_Error

20        FillGrid

30        Exit Sub

cmdRefresh_Click_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmLabConsultantList", "cmdRefresh_Click", intEL, strES

End Sub

Private Sub cmdRemove_Click()

      Dim n As Integer
      Dim sql As String

10    On Error GoTo cmdRemove_Click_Error

20    If iMsg("Are you sure you want to remove selected report from consultant list", vbYesNo) = vbYes Then
30        With grdSID
40            sql = "Delete from ConsultantList " & _
                    "where SampleID = '" & SysOptMicroOffset(0) + .TextMatrix(.row, 0) & "'"
50            Cnxn(0).Execute sql
60            RemoveReport 0, SysOptMicroOffset(0) + .TextMatrix(.row, 0), "N", 0
70        End With
80        cmdRemove.Enabled = False
90        FillGrid
100   End If

110   Exit Sub

cmdRemove_Click_Error:

      Dim strES As String
      Dim intEL As Integer

120   intEL = Erl
130   strES = Err.Description
140   LogError "frmLabConsultantList", "cmdRemove_Click", intEL, strES

End Sub

Private Sub grdSID_Click()
      Dim sql As String

10    On Error GoTo grdSID_Click_Error
20    cmdRemove.Enabled = False
30    With grdSID

40        If InStr(UCase(Trim(.TextMatrix(.row, 8))), UCase("In Lab For Review")) Then
50            If .ColSel = 0 Then
60                If MsgBox("Are You Sure you want to delete sampleID " & .TextMatrix(.row, 0), vbYesNo) = vbYes Then

70                    sql = "Delete from ConsultantList " & _
                            "where SampleID = '" & SysOptMicroOffset(0) + .TextMatrix(.row, 0) & "'"
80                    Cnxn(0).Execute sql
90                    RemoveReport 0, SysOptMicroOffset(0) + .TextMatrix(.row, 0), "N", 0
100               End If
110               FillGrid
120           ElseIf .ColSel = 9 Then
130               If .TextMatrix(.row, 9) = "No" Then
140                   .TextMatrix(.row, 9) = "Yes"
150               Else
160                   .TextMatrix(.row, 9) = "No"
170               End If
180               sql = "Update ConsultantList set Ack =" & IIf(.TextMatrix(.row, 9) = "No", 0, 1) & _
                        "where SampleID = '" & SysOptMicroOffset(0) + .TextMatrix(.row, 0) & "'"
190               Cnxn(0).Execute sql
200           End If
210       ElseIf InStr(UCase(Trim(.TextMatrix(.row, 8))), UCase("With Consultant")) Then
220           cmdRemove.Enabled = True

230       End If
240   End With

250   Exit Sub


grdSID_Click_Error:

      Dim strES As String
      Dim intEL As Integer

260   intEL = Erl
270   strES = Err.Description
280   LogError "frmLabConsultantList", "grdSID_Click", intEL, strES, sql
End Sub


Private Sub Form_Load()
10        dtpFrom.Value = Now - 10
20        dtpTo.Value = Now
30        FillGrid
End Sub


