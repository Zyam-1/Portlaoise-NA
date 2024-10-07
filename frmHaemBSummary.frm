VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmHaemBSummary 
   Caption         =   "NetAcquire 6 - Haematology Blood Film Daily Summary"
   ClientHeight    =   8220
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15240
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "frmHaemBSummary.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8220
   ScaleWidth      =   15240
   Begin VB.CommandButton cmdPrintUn 
      Caption         =   "&Print Unvalidated"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7515
      Picture         =   "frmHaemBSummary.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   90
      Width           =   1140
   End
   Begin VB.CommandButton cmdUnValidate 
      Caption         =   "Unvalidate Selected Rows"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8775
      Picture         =   "frmHaemBSummary.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   90
      Width           =   1230
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1290
      Left            =   4185
      TabIndex        =   12
      Top             =   -45
      Width           =   2895
      Begin VB.OptionButton v 
         Caption         =   "Not Valid"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   17
         Top             =   180
         Width           =   1545
      End
      Begin VB.OptionButton v 
         Caption         =   "Valid, not Printed"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   15
         Top             =   405
         Width           =   1545
      End
      Begin VB.OptionButton v 
         Caption         =   "Valid (Printed or Not Printed)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   14
         Top             =   630
         Width           =   2310
      End
      Begin VB.OptionButton v 
         Caption         =   "All (Valid, Not Valid, Printed and Not Printed)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   180
         TabIndex        =   13
         Top             =   855
         Value           =   -1  'True
         Width           =   2580
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Sort By"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   150
      TabIndex        =   7
      Top             =   90
      Width           =   2115
      Begin VB.OptionButton oSort 
         Caption         =   "Chart"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   1110
         TabIndex        =   11
         Top             =   480
         Width           =   855
      End
      Begin VB.OptionButton oSort 
         Caption         =   "Clinician"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   1110
         TabIndex        =   10
         Top             =   240
         Width           =   885
      End
      Begin VB.OptionButton oSort 
         Alignment       =   1  'Right Justify
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   735
      End
      Begin VB.OptionButton oSort 
         Alignment       =   1  'Right Justify
         Caption         =   "Run #"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   210
         TabIndex        =   8
         Top             =   240
         Value           =   -1  'True
         Width           =   765
      End
   End
   Begin MSComCtl2.DTPicker dtRunDate 
      Height          =   315
      Left            =   2580
      TabIndex        =   5
      Top             =   450
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
      _Version        =   393216
      Format          =   59637761
      CurrentDate     =   36965
   End
   Begin VB.CommandButton bValidate 
      Caption         =   "Validate Selected Rows"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   10125
      Picture         =   "frmHaemBSummary.frx":0A56
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   90
      Width           =   1185
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   6840
      Left            =   -30
      TabIndex        =   3
      Top             =   1245
      Width           =   15105
      _ExtentX        =   26644
      _ExtentY        =   12065
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      BackColorSel    =   65535
      ForeColorSel    =   12583104
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      GridLines       =   3
      GridLinesFixed  =   3
      AllowUserResizing=   1
      FormatString    =   $"frmHaemBSummary.frx":0D60
   End
   Begin VB.CommandButton bview 
      Caption         =   "&View Results"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   11430
      Picture         =   "frmHaemBSummary.frx":0E4F
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   90
      Width           =   1140
   End
   Begin VB.CommandButton bPrint 
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   12690
      Picture         =   "frmHaemBSummary.frx":1159
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   90
      Width           =   1140
   End
   Begin VB.CommandButton bexit 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   13920
      Picture         =   "frmHaemBSummary.frx":1463
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   90
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Run Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2610
      TabIndex        =   6
      Top             =   240
      Width           =   690
   End
End
Attribute VB_Name = "frmHaemBSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit  '© Custom Software 2001

Private Sub bexit_Click()

10        Unload Me

End Sub

Private Sub bprint_Click()

          Dim n As Long
          Dim sortby As Long
          Dim OrigOrientation As Long

10        On Error GoTo bprint_Click_Error

20        If g.Rows = 2 And g.TextMatrix(1, 0) = "" Then Exit Sub

30        For n = 0 To 3
40            If oSort(n) Then
50                sortby = n
60                Exit For
70            End If
80        Next

90        OrigOrientation = Printer.Orientation
100       Printer.Orientation = vbPRORPortrait

110       Printer.FontName = "Courier"
120       Printer.FontSize = 12
130       Printer.Print HospName(0)
140       Printer.Print "Haematology - Summary "; dtRunDate
150       Printer.Print "Sorted by ";
160       Printer.Print Choose(sortby + 1, "Lab Run Number", "Patient Name", _
                               "Clinician", "Chart")
170       Printer.Print

180       Printer.Font.Size = 10
190       GoSub PrintHead
200       For n = 1 To g.Rows - 1
210           If Printer.CurrentY > Printer.Height - 800 Then
220               Printer.NewPage
230               GoSub PrintHead
240           End If
250           Printer.Print g.TextMatrix(n, 0);                    'lab
260           Printer.Print Tab(10); g.TextMatrix(n, 1);            'Chart
270           Printer.Print Tab(20); g.TextMatrix(n, 2);           'dob
280           Printer.Print Tab(34); Left(g.TextMatrix(n, 3), 19);    'name
290           Printer.Print Tab(55); Left(g.TextMatrix(n, 4) & Space(19), 19);    'clinician

300           Printer.Print "  ";
310           Printer.Print Tab(85); Left(g.TextMatrix(n, 5) & Space(6), 6);
320           Printer.Print Tab(92); Left(g.TextMatrix(n, 6) & Space(6), 6);
330           Printer.Print Tab(99); Left(g.TextMatrix(n, 7) & Space(6), 6);
340           Printer.Print Tab(106); Left(g.TextMatrix(n, 8) & Space(6), 6);
350           Printer.Print Tab(113); Left(g.TextMatrix(n, 9) & Space(6), 6);
360           Printer.Print Tab(120); Left(g.TextMatrix(n, 10) & Space(6), 6);
370           Printer.Print Tab(127); Left(g.TextMatrix(n, 11) & Space(6), 6);
380           Printer.Print Tab(130); Left(g.TextMatrix(n, 12) & Space(6), 6);
390           Printer.Print Tab(140); Left(g.TextMatrix(n, 13) & Space(6), 6);
400           Printer.Font.Size = 10
410           Printer.Print
420       Next

430       Printer.EndDoc

440       Printer.Orientation = OrigOrientation



450       Exit Sub

PrintHead:
460       Printer.Print g.TextMatrix(n, 0);                    'lab
470       Printer.Print Tab(10); g.TextMatrix(n, 1);            'Chart
480       Printer.Print Tab(20); g.TextMatrix(n, 2);           'dob
490       Printer.Print Tab(34); Left(g.TextMatrix(n, 3), 19);    'name
500       Printer.Print Tab(55); Left(g.TextMatrix(n, 4) & Space(19), 19);    'clinician

510       Printer.Print "  ";
520       Printer.Print Tab(85); Left(g.TextMatrix(n, 5) & Space(6), 6);
530       Printer.Print Tab(92); Left(g.TextMatrix(n, 6) & Space(6), 6);
540       Printer.Print Tab(99); Left(g.TextMatrix(n, 7) & Space(6), 6);
550       Printer.Print Tab(106); Left(g.TextMatrix(n, 8) & Space(6), 6);
560       Printer.Print Tab(113); Left(g.TextMatrix(n, 9) & Space(6), 6);
570       Printer.Print Tab(120); Left(g.TextMatrix(n, 10) & Space(6), 6);
580       Printer.Print Tab(127); Left(g.TextMatrix(n, 11) & Space(6), 6);
590       Printer.Print Tab(130); Left(g.TextMatrix(n, 12) & Space(6), 6);
600       Printer.Print Tab(140); Left(g.TextMatrix(n, 13) & Space(6), 6);
610       Printer.Font.Size = 10
620       Printer.Print
630       Return


640       Exit Sub

bprint_Click_Error:

          Dim strES As String
          Dim intEL As Integer



650       intEL = Erl
660       strES = Err.Description
670       LogError "frmHaemBSummary", "bPrint_Click", intEL, strES


End Sub

Private Sub FillG()

          Dim n As Long
          Dim sn As New Recordset
          Dim tb As New Recordset
          Dim tbd As Recordset
          Dim sql As String
          Dim s As String

10        On Error GoTo FillG_Error

20        g.Rows = 2
30        g.AddItem ""
40        g.RemoveItem 1

50        sql = "SELECT * from haemresults WHERE " & _
                "(rundatetime between '" & Format(dtRunDate, "dd/mmm/yyyy 00:00:00") & "' " & _
                "AND '" & Format(dtRunDate, "dd/mmm/yyyy 23:59:59") & "') " & _
                "AND cfilm = 1 "

60        If v(0) Then
70            sql = sql & " and (valid = 1 and printed <> 1)"
80        ElseIf v(1) Then
90            sql = sql & " and valid = 1"
100       ElseIf v(3) Then
110           sql = sql & " and (valid = 0 or valid is null)"
120       ElseIf v(2) Then
130           sql = sql & " and (valid = 0 or valid is null or valid = 1)"
140       End If

150       Set sn = New Recordset
160       RecOpenServer 0, sn, sql

170       g.Visible = False

180       Do While Not sn.EOF
190           If IsNumeric(sn!SampleID) Then
200               sql = "SELECT * FROM Demographics WHERE " & _
                        "SampleID = '" & sn!SampleID & "'"

210               Set tb = New Recordset
220               RecOpenServer 0, tb, sql
230               If Not tb.EOF Then
240                   s = sn!SampleID & vbTab
250                   s = s & tb!Chart & vbTab
260                   s = s & Format(tb!Dob, "dd/MM/yy") & vbTab & _
                          tb!PatName & vbTab & _
                          tb!Clinician & vbTab & tb!GP & "" & vbTab & tb!Ward & ""
270               Else
280                   sql = "SELECT * FROM Demographics WHERE " & _
                            "SampleID = '" & sn!SampleID & "'"

290                   Set tbd = New Recordset
300                   RecOpenServer 0, tbd, sql
310                   If Not tbd.EOF Then
320                       s = s & Format(tbd!Dob, "dd/MM/yy") & vbTab & _
                              tbd!PatName & vbTab & _
                              tbd!Clinician & vbTab & tb!GP & "" & vbTab & tb!Ward & ""
330                   Else
340                       s = s & vbTab & vbTab & vbTab & vbTab
350                   End If
360               End If


370               sql = "SELECT Comment FROM Observations WHERE " & _
                        "SampleID = '" & sn!SampleID & "' " & _
                        "AND Discipline = 'Haematology'"
380               Set tbd = New Recordset
390               RecOpenServer 0, tbd, sql
400               If Not tbd.EOF Then
410                   s = s & vbTab & tbd!Comment & ""
420               End If
430               g.AddItem s
440           End If
450           sn.MoveNext
460       Loop

470       For n = 0 To 3
480           If oSort(n) Then
490               g.Col = Choose(n + 1, 0, 3, 4, 1)
500               Exit For
510           End If
520       Next

530       If g.Rows > 2 Then g.RemoveItem 1

540       g.Sort = flexSortGenericAscending

550       g.Visible = True






560       Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer



570       intEL = Erl
580       strES = Err.Description
590       LogError "frmHaemBSummary", "FillG", intEL, strES, sql


End Sub

Private Sub cmdPrintUn_Click()

          Dim n As Integer
          Dim sql As String
          Dim tb As Recordset
          Dim StartRow As Long
          Dim StopRow As Long

10        On Error GoTo cmdPrintUn_Click_Error

20        If g.Rows = 2 And g.TextMatrix(1, 0) = "" Then Exit Sub

30        If g.Row > g.RowSel Then
40            StartRow = g.RowSel
50            StopRow = g.Row
60        Else
70            StartRow = g.Row
80            StopRow = g.RowSel
90        End If

100       For n = StartRow To StopRow
110           sql = "SELECT * FROM PrintPending WHERE " & _
                    "Department = 'K' " & _
                    "AND SampleID = '" & g.TextMatrix(n, 0) & "'"
120           Set tb = New Recordset
130           RecOpenServer 0, tb, sql
140           tb.AddNew
150           tb!SampleID = g.TextMatrix(n, 0)
160           tb!Department = "K"
170           tb!Initiator = UserCode
180           tb!Ward = g.TextMatrix(n, 6)
190           tb!GP = g.TextMatrix(n, 5)
200           tb!Clinician = g.TextMatrix(n, 4)
210           tb.Update
220       Next

230       Exit Sub

cmdPrintUn_Click_Error:

          Dim strES As String
          Dim intEL As Integer



240       intEL = Erl
250       strES = Err.Description
260       LogError "frmHaemBSummary", "cmdPrintUn_Click", intEL, strES, sql

End Sub

Private Sub cmdUnValidate_Click()

          Dim n As Long
          Dim StartRow As Long
          Dim StopRow As Long
          Dim tb As New Recordset
          Dim sql As String


10        On Error GoTo cmdUnValidate_Click_Error

20        If g.Rows = 2 And g.TextMatrix(1, 0) = "" Then Exit Sub

30        If g.Row > g.RowSel Then
40            StartRow = g.RowSel
50            StopRow = g.Row
60        Else
70            StartRow = g.Row
80            StopRow = g.RowSel
90        End If


100       If UCase(iBOX("Password!", , , True)) <> UCase(UserPass) Then
110           iMsg "Invalid password!"
120           Exit Sub
130       End If


140       If StartRow = StopRow Then
150           g.Col = 0
160           If iMsg("unValidate Lab # " & g & "?", vbQuestion + vbYesNo) = vbNo Then
170               Exit Sub
180           End If
190       Else
200           If iMsg("unValidate all SELECTED rows?", vbQuestion + vbYesNo) = vbNo Then
210               Exit Sub
220           End If
230       End If


240       g.Col = 0
250       For n = StartRow To StopRow
260           g.Row = n
270           sql = "Select * from Haemresults where sampleid = " & Trim(g) & ""
280           Set tb = New Recordset
290           RecOpenServer 0, tb, sql
300           If Not tb.EOF Then
310               ArchiveHaem Trim$(g)
320               tb!Valid = 0
330               tb!HealthLink = 0
340               tb!Operator = UserCode
350               tb.Update
360           End If
370       Next

380       FillG

390       Exit Sub

cmdUnValidate_Click_Error:

          Dim strES As String
          Dim intEL As Integer



400       intEL = Erl
410       strES = Err.Description
420       LogError "frmHaemBSummary", "cmdUnValidate_Click", intEL, strES, sql


End Sub
Private Sub bValidate_Click()

          Dim n As Long
          Dim StartRow As Long
          Dim StopRow As Long
          Dim tb As New Recordset
          Dim sql As String

10        On Error GoTo bValidate_Click_Error

20        If g.Rows = 2 And g.TextMatrix(1, 0) = "" Then Exit Sub


30        If g.Row > g.RowSel Then
40            StartRow = g.RowSel
50            StopRow = g.Row
60        Else
70            StartRow = g.Row
80            StopRow = g.RowSel
90        End If



100       If StartRow = StopRow Then
110           g.Col = 0
120           If iMsg("Validate Lab # " & g & "?", vbQuestion + vbYesNo) = vbNo Then
130               Exit Sub
140           End If
150       Else
160           If iMsg("Validate all SELECTED rows?", vbQuestion + vbYesNo) = vbNo Then
170               Exit Sub
180           End If
190       End If

200       If iMsg("Validate all demographics for SELECTED rows?", vbQuestion + vbYesNo) = vbNo Then
210           Exit Sub
220       End If

230       g.Col = 0
240       For n = StartRow To StopRow
250           g.Row = n
260           sql = "Update demographics set valid = 1, operator  = '" & UserCode & "' where sampleid = " & Trim(g) & ""
270           Cnxn(0).Execute sql
280           sql = "Select * from Haemresults where sampleid = " & Trim(g) & ""
290           Set tb = New Recordset
300           RecOpenServer 0, tb, sql
310           If Not tb.EOF Then
320               ArchiveHaem Trim$(g)
330               tb!Valid = 1
340               tb!Operator = UserCode
350               tb.Update
360           End If
370       Next

380       FillG

390       Exit Sub

bValidate_Click_Error:

          Dim strES As String
          Dim intEL As Integer



400       intEL = Erl
410       strES = Err.Description
420       LogError "frmHaemBSummary", "bValidate_Click", intEL, strES, sql


End Sub

Private Sub bview_Click()


10        On Error GoTo bview_Click_Error

20        If g.TextMatrix(g.RowSel, 0) = "" Then Exit Sub


30        With frmEditAll
40            .txtSampleID = g.TextMatrix(g.RowSel, 0)
50            .txtSampleID_LostFocus
60            .Show 1
70        End With

80        Exit Sub

bview_Click_Error:

          Dim strES As String
          Dim intEL As Integer



90        intEL = Erl
100       strES = Err.Description
110       LogError "frmHaemBSummary", "bview_Click", intEL, strES


End Sub

Private Sub dtRunDate_CloseUp()

10        On Error GoTo dtRunDate_CloseUp_Error

20        FillG

30        Exit Sub

dtRunDate_CloseUp_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmHaemBSummary", "dtRunDate_CloseUp", intEL, strES


End Sub


Private Sub Form_Load()

10        dtRunDate = Format(Now, "dd/mmm/yyyy")

20        oSort(3).Caption = "Chart"
30        g.TextMatrix(0, 1) = "Chart"

40        If TestSys = True Then Me.Caption = Me.Caption & " - TEST SYSTEM"

50        Set_Font Me

End Sub

Private Sub g_Click()
          Static SortOrder As Boolean
          Dim ySave As Long



10        On Error GoTo g_Click_Error

20        ySave = g.Row

30        If g.MouseRow = 0 Then
40            If SortOrder Then
50                g.Sort = flexSortGenericAscending
60            Else
70                g.Sort = flexSortGenericDescending
80            End If
90            SortOrder = Not SortOrder
100       End If


110       Exit Sub

g_Click_Error:

          Dim strES As String
          Dim intEL As Integer



120       intEL = Erl
130       strES = Err.Description
140       LogError "frmHaemBSummary", "g_Click", intEL, strES


End Sub

Private Sub g_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        On Error GoTo g_MouseMove_Error

20        Y = g.MouseCol
30        X = g.MouseRow
40        g.ToolTipText = "Haematology Results"

50        If g.MouseCol = 7 Then
60            If Trim(g.TextMatrix(X, Y)) <> "" Then g.ToolTipText = g.TextMatrix(X, Y)
70        End If

80        Exit Sub

g_MouseMove_Error:

          Dim strES As String
          Dim intEL As Integer



90        intEL = Erl
100       strES = Err.Description
110       LogError "frmHaemBSummary", "g_MouseMove", intEL, strES


End Sub

Private Sub osort_Click(Index As Integer)

          Dim n As Long

10        On Error GoTo osort_Click_Error

20        For n = 0 To 3
30            If oSort(n) Then
40                g.Col = Choose(n + 1, 0, 3, 4, 1)
50                Exit For
60            End If
70        Next

80        g.Sort = flexSortGenericAscending

90        Exit Sub

osort_Click_Error:

          Dim strES As String
          Dim intEL As Integer



100       intEL = Erl
110       strES = Err.Description
120       LogError "frmHaemBSummary", "osort_Click", intEL, strES


End Sub

Private Sub v_Click(Index As Integer)

10        On Error GoTo v_Click_Error

20        If Index = 3 Then
30            cmdPrintUn.Visible = True
40        Else
50            cmdPrintUn.Visible = False
60        End If

70        FillG

80        Exit Sub

v_Click_Error:

          Dim strES As String
          Dim intEL As Integer



90        intEL = Erl
100       strES = Err.Description
110       LogError "frmHaemBSummary", "v_Click", intEL, strES


End Sub
