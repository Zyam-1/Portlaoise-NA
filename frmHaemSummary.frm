VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmHaemSummary 
   Caption         =   "NetAcquire 6 - Haematology Daily Summary"
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
   Icon            =   "frmHaemSummary.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8220
   ScaleWidth      =   15240
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
      Left            =   8880
      Picture         =   "frmHaemSummary.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   90
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   6840
      Left            =   45
      TabIndex        =   3
      Top             =   1215
      Width           =   15105
      _ExtentX        =   26644
      _ExtentY        =   12065
      _Version        =   393216
      Cols            =   15
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
      FormatString    =   $"frmHaemSummary.frx":0614
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
      Left            =   10440
      Picture         =   "frmHaemSummary.frx":06FA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   90
      Width           =   1455
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
      Left            =   12000
      Picture         =   "frmHaemSummary.frx":0A04
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   90
      Width           =   1455
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
      Left            =   13560
      Picture         =   "frmHaemSummary.frx":0D0E
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   90
      Width           =   1455
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
Attribute VB_Name = "frmHaemSummary"
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

20        For n = 0 To 3
30            If oSort(n) Then
40                sortby = n
50                Exit For
60            End If
70        Next

80        OrigOrientation = Printer.Orientation
90        Printer.Orientation = vbPRORPortrait

100       Printer.FontName = "Courier"
110       Printer.FontSize = 12
120       Printer.Print HospName(0)
130       Printer.Print "Haematology - Summary "; dtRunDate
140       Printer.Print "Sorted by ";
150       Printer.Print Choose(sortby + 1, "Lab Run Number", "Patient Name", _
                               "Clinician", "Chart")
160       Printer.Print

170       Printer.Font.Size = 10
180       GoSub PrintHead
190       For n = 1 To g.Rows - 1
200           If Printer.CurrentY > Printer.Height - 800 Then
210               Printer.NewPage
220               GoSub PrintHead
230           End If
240           Printer.Print g.TextMatrix(n, 0);                    'lab
250           Printer.Print Tab(10); g.TextMatrix(n, 1);            'Chart
260           Printer.Print Tab(20); g.TextMatrix(n, 2);           'dob
270           Printer.Print Tab(34); Left(g.TextMatrix(n, 3), 19);    'name
280           Printer.Print Tab(55); Left(g.TextMatrix(n, 4) & Space(19), 19);    'clinician

290           Printer.Print "  ";
300           Printer.Print Tab(85); Left(g.TextMatrix(n, 5) & Space(6), 6);
310           Printer.Print Tab(92); Left(g.TextMatrix(n, 6) & Space(6), 6);
320           Printer.Print Tab(99); Left(g.TextMatrix(n, 7) & Space(6), 6);
330           Printer.Print Tab(106); Left(g.TextMatrix(n, 8) & Space(6), 6);
340           Printer.Print Tab(113); Left(g.TextMatrix(n, 9) & Space(6), 6);
350           Printer.Print Tab(120); Left(g.TextMatrix(n, 10) & Space(6), 6);
360           Printer.Print Tab(127); Left(g.TextMatrix(n, 11) & Space(6), 6);
370           Printer.Print Tab(130); Left(g.TextMatrix(n, 12) & Space(6), 6);
380           Printer.Print Tab(140); Left(g.TextMatrix(n, 13) & Space(6), 6);
390           Printer.Font.Size = 10
400           Printer.Print
410       Next

420       Printer.EndDoc

430       Printer.Orientation = OrigOrientation



440       Exit Sub

PrintHead:
450       Printer.Print g.TextMatrix(n, 0);                    'lab
460       Printer.Print Tab(10); g.TextMatrix(n, 1);            'nopas
470       Printer.Print Tab(20); g.TextMatrix(n, 2);           'dob
480       Printer.Print Tab(34); Left(g.TextMatrix(n, 3), 19);    'name
490       Printer.Print Tab(55); Left(g.TextMatrix(n, 4) & Space(19), 19);    'clinician

500       Printer.Print "  ";
510       Printer.Print Tab(85); Left(g.TextMatrix(n, 5) & Space(6), 6);
520       Printer.Print Tab(92); Left(g.TextMatrix(n, 6) & Space(6), 6);
530       Printer.Print Tab(99); Left(g.TextMatrix(n, 7) & Space(6), 6);
540       Printer.Print Tab(106); Left(g.TextMatrix(n, 8) & Space(6), 6);
550       Printer.Print Tab(113); Left(g.TextMatrix(n, 9) & Space(6), 6);
560       Printer.Print Tab(120); Left(g.TextMatrix(n, 10) & Space(6), 6);
570       Printer.Print Tab(127); Left(g.TextMatrix(n, 11) & Space(6), 6);
580       Printer.Print Tab(130); Left(g.TextMatrix(n, 12) & Space(6), 6);
590       Printer.Print Tab(140); Left(g.TextMatrix(n, 13) & Space(6), 6);
600       Printer.Font.Size = 10
610       Printer.Print
620       Return


630       Exit Sub

bprint_Click_Error:

          Dim strES As String
          Dim intEL As Integer



640       intEL = Erl
650       strES = Err.Description
660       LogError "frmHaemSummary", "bPrint_Click", intEL, strES


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
                "rundatetime between '" & Format(dtRunDate, "dd/mmm/yyyy 00:00:00") & "' and '" & Format(dtRunDate, "dd/mmm/yyyy 23:59:59") & "'"
60        Set sn = New Recordset
70        RecOpenServer 0, sn, sql

80        g.Visible = False

90        Do While Not sn.EOF
100           If IsNumeric(sn!SampleID) Then
110               sql = "SELECT * FROM Demographics WHERE " & _
                        "SampleID = '" & sn!SampleID & "'"

120               Set tb = New Recordset
130               RecOpenServer 0, tb, sql
140               If Not tb.EOF Then
150                   s = sn!SampleID & vbTab
160                   s = s & tb!Chart & vbTab
170                   s = s & Format(tb!Dob, "dd/MM/yy") & vbTab & _
                          tb!PatName & vbTab & _
                          tb!Clinician & vbTab & tb!GP & "" & vbTab
180               Else
190                   sql = "SELECT * FROM Demographics WHERE " & _
                            "SampleID = '" & sn!SampleID & "'"

200                   Set tbd = New Recordset
210                   RecOpenServer 0, tbd, sql
220                   If Not tbd.EOF Then
230                       s = s & Format(tbd!Dob, "dd/MM/yy") & vbTab & _
                              tbd!PatName & vbTab & _
                              tbd!Clinician & vbTab & tb!GP & "" & vbTab
240                   Else
250                       s = s & vbTab & vbTab & vbTab & vbTab
260                   End If
270               End If
280               s = s & Format(sn!wbc, "#0.0") & vbTab & _
                      Format(sn!rbc, "0.00") & vbTab & _
                      Format(sn!Hgb, "#0.0") & vbTab & _
                      Format(sn!MCV, "###") & vbTab & _
                      Format(sn!Plt, "###") & vbTab

290               s = s & Trim(sn!esr) & ""
300               s = s & vbTab
310               If Not IsNull(sn!cFilm) Then
320                   If sn!cFilm Then s = s & "Yes"
330               End If
340               s = s & vbTab & sn!Monospot & "" & vbTab & Trim(sn!reta & "")
350               g.AddItem s
360           End If
370           sn.MoveNext
380       Loop

390       For n = 0 To 3
400           If oSort(n) Then
410               g.Col = Choose(n + 1, 0, 3, 4, 1)
420               Exit For
430           End If
440       Next

450       If g.Rows > 2 Then g.RemoveItem 1

460       g.Sort = flexSortGenericAscending

470       g.Visible = True





480       Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer



490       intEL = Erl
500       strES = Err.Description
510       LogError "frmHaemSummary", "FillG", intEL, strES, sql


End Sub

Private Sub bValidate_Click()

          Dim n As Long
          Dim StartRow As Long
          Dim StopRow As Long
          Dim sql As String


10        On Error GoTo bValidate_Click_Error

20        If g.Row > g.RowSel Then
30            StartRow = g.RowSel
40            StopRow = g.Row
50        Else
60            StartRow = g.Row
70            StopRow = g.RowSel
80        End If

90        If StartRow = StopRow Then
100           g.Col = 0
110           If iMsg("Validate Lab # " & g & "?", vbQuestion + vbYesNo) = vbNo Then
120               Exit Sub
130           End If
140       Else
150           If iMsg("Validate all SELECTED rows?", vbQuestion + vbYesNo) = vbNo Then
160               Exit Sub
170           End If
180       End If

190       If iMsg("Validate all demographics for SELECTED rows?", vbQuestion + vbYesNo) = vbNo Then
200           Exit Sub
210       End If

220       g.Col = 0
230       For n = StartRow To StopRow
240           g.Row = n
250           If g <> "Lab #" Then
260               sql = "update demographics set  valid = 1, username = '" & Username & "' where sampleid = '" & Trim(g) & "' and valid <> 1"
270               Cnxn(0).Execute sql
280               sql = "update haemresults set  valid = 1, operator = '" & UserCode & "' where sampleid = '" & Trim(g) & "' and valid <> 1"
290               Cnxn(0).Execute sql
300           End If
310       Next

320       Exit Sub

bValidate_Click_Error:

          Dim strES As String
          Dim intEL As Integer



330       intEL = Erl
340       strES = Err.Description
350       LogError "frmHaemSummary", "bValidate_Click", intEL, strES


End Sub

Private Sub bview_Click()


10        On Error GoTo bview_Click_Error

20        If g.TextMatrix(g.RowSel, 0) = "" Then Exit Sub

30        g.Col = 0
40        If g = "Lab #" Then Exit Sub
50        g.Col = 1

60        With frmFullHaem
70            g.Col = 3
80            .lblName = g
90            g.Col = 2
100           .lblDoB = g
110           .Show 1
120       End With

130       Exit Sub

bview_Click_Error:

          Dim strES As String
          Dim intEL As Integer



140       intEL = Erl
150       strES = Err.Description
160       LogError "frmHaemSummary", "bview_Click", intEL, strES


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
60        LogError "frmHaemSummary", "dtRunDate_CloseUp", intEL, strES


End Sub


Private Sub Form_Load()

10        On Error GoTo Form_Load_Error

20        dtRunDate = Format(Now, "dd/mmm/yyyy")

30        oSort(3).Caption = "Chart"
40        g.TextMatrix(0, 1) = "Chart"

50        If TestSys = True Then Me.Caption = Me.Caption & " - TEST SYSTEM"

60        Set_Font Me

70        Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

80        intEL = Erl
90        strES = Err.Description
100       LogError "frmHaemSummary", "Form_Load", intEL, strES

End Sub

Private Sub osort_Click(Index As Integer)

          Dim n As Long

10        For n = 0 To 3
20            If oSort(n) Then
30                g.Col = Choose(n + 1, 0, 3, 4, 1)
40                Exit For
50            End If
60        Next

70        g.Sort = flexSortGenericAscending

End Sub

