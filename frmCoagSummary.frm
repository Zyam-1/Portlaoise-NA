VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.ocx"
Begin VB.Form frmCoagSummary 
   Caption         =   "NetAcquire 6 - Coagulation Daily Summary"
   ClientHeight    =   8985
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   12195
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
   Icon            =   "frmCoagSummary.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8985
   ScaleWidth      =   12195
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   900
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   12195
      _ExtentX        =   21511
      _ExtentY        =   1588
      ButtonWidth     =   1905
      ButtonHeight    =   1429
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   11
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Validate"
            Key             =   ""
            Object.ToolTipText     =   "Vaildate Selected Rows"
            Object.Tag             =   "Val"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Start"
            Key             =   ""
            Object.ToolTipText     =   "Start Search"
            Object.Tag             =   "Search"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "View Results"
            Key             =   ""
            Object.ToolTipText     =   "View Results"
            Object.Tag             =   "View"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Print"
            Key             =   ""
            Object.ToolTipText     =   "Print"
            Object.Tag             =   "Print"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Export"
            Key             =   ""
            Object.ToolTipText     =   "Export to excel"
            Object.Tag             =   "Export"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Exit"
            Key             =   ""
            Object.ToolTipText     =   "Exit"
            Object.Tag             =   "Exit"
            ImageIndex      =   4
         EndProperty
      EndProperty
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
      Height          =   525
      Left            =   60
      TabIndex        =   3
      Top             =   1035
      Width           =   6420
      Begin VB.ComboBox cmbWard 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4410
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   180
         Width           =   1860
      End
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
         Left            =   2910
         TabIndex        =   7
         Top             =   240
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
         TabIndex        =   6
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
         Left            =   2100
         TabIndex        =   5
         Top             =   240
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
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   765
      End
      Begin VB.Label Wa 
         Caption         =   "Ward"
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
         Left            =   3960
         TabIndex        =   10
         Top             =   225
         Width           =   735
      End
   End
   Begin MSComCtl2.DTPicker dtRunDate 
      Height          =   315
      Left            =   7350
      TabIndex        =   1
      Top             =   1140
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd/MMM/yyyy"
      Format          =   77660163
      CurrentDate     =   36965
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   7065
      Left            =   90
      TabIndex        =   0
      Top             =   1665
      Width           =   12045
      _ExtentX        =   21246
      _ExtentY        =   12462
      _Version        =   393216
      Cols            =   12
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
      FormatString    =   $"frmCoagSummary.frx":030A
   End
   Begin VB.Label lblExcelInfo 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exporting..."
      Height          =   345
      Left            =   10740
      TabIndex        =   11
      Top             =   1320
      Visible         =   0   'False
      Width           =   1305
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   6870
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   8
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCoagSummary.frx":03E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCoagSummary.frx":0702
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCoagSummary.frx":0A1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCoagSummary.frx":0D36
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCoagSummary.frx":1050
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCoagSummary.frx":136A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCoagSummary.frx":1684
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCoagSummary.frx":199E
            Key             =   ""
         EndProperty
      EndProperty
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
      Left            =   6600
      TabIndex        =   2
      Top             =   1200
      Width           =   690
   End
End
Attribute VB_Name = "frmCoagSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit  '© Custom Software 2001

Private Sub dtRunDate_CloseUp()

10        On Error GoTo dtRunDate_CloseUp_Error

20        FillG

30        Exit Sub

dtRunDate_CloseUp_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmCoagSummary", "dtRunDate_CloseUp", intEL, strES


End Sub

Private Sub Exit_Click()

10        Unload Me

End Sub

Private Sub FillG()

          Dim Ward As String
          Dim n As Long
          Dim sn As New Recordset
          Dim snr As Recordset
          Dim sql As String
          Dim s As String
          Dim A As String
          Dim b As String
          Dim c As String
          Dim D As String
          Dim e As String

10        On Error GoTo FillG_Error

20        ClearFGrid g
30        g.Visible = False

40        sql = "SELECT distinct(demographics.sampleid), demographics.* from demographics, coagresults WHERE " & _
                "demographics.rundate = '" & Format(dtRunDate, "dd/MMM/yyyy") & "' and coagresults.sampleid = demographics.sampleid"
50        Set sn = New Recordset
60        RecOpenServer 0, sn, sql

70        g.Visible = False

80        Do While Not sn.EOF
90            Ward = ""
100           s = sn!SampleID & vbTab & vbTab
110           If sn!Dob & "" <> "" Then
120               s = s & Format(sn!Dob, "dd/MM/yyyy") & vbTab & _
                      sn!PatName & vbTab & _
                      sn!Clinician & vbTab & sn!Ward & vbTab & sn!GP & "" & vbTab
130           Else
140               s = s & vbTab & sn!PatName & vbTab & sn!Clinician & vbTab & sn!Ward & vbTab & sn!GP & "" & vbTab
150           End If
160           Ward = UCase(sn!Ward) & ""
170           A = ""
180           b = ""
190           c = ""
200           D = ""
210           e = ""
220           sql = "SELECT coagresults.*, coagtestdefinitions.testname from coagresults, coagtestdefinitions WHERE " & _
                    "coagresults.sampleid = " & sn!SampleID & " and coagtestdefinitions.code = coagresults.code"
230           If SysOptExp(0) Then
240               sql = sql & " and coagtestdefinitions.units = coagresults.units"
250           End If

260           Set snr = New Recordset
270           RecOpenServer 0, snr, sql
280           Do While Not snr.EOF
290               Select Case UCase(Trim(snr!TestName))
                  Case "PT"
300                   A = snr!Result
310               Case "INR"
320                   b = snr!Result
330               Case "APTT"
340                   c = snr!Result
350               Case "D-DIMER"
360                   D = snr!Result
370               Case "DDIMERS"
380                   D = snr!Result
390               Case "FIB"
400                   e = snr!Result
410               Case "FIBRINOGEN"
420                   e = snr!Result
430               Case "FIB-CLAUSS"
440                   e = snr!Result
450               End Select
460               snr.MoveNext
470           Loop
480           If UCase(cmbWard) = UCase(Ward) Or cmbWard = "All" Then
490               s = s & A & vbTab & b & vbTab & c & vbTab & D & vbTab & e
500               g.AddItem s
510           End If
520           sn.MoveNext
530       Loop

540       For n = 0 To 4
550           If oSort(n) Then
560               g.Col = Choose(n + 1, 0, 3, 4, 5, 1)
570               Exit For
580           End If
590       Next

600       g.Sort = flexSortGenericAscending

610       FixG g

620       Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer

630       intEL = Erl
640       strES = Err.Description
650       LogError "frmCoagSummary", "FillG", intEL, strES

End Sub

Private Sub Form_Load()

10        On Error GoTo Form_Load_Error

20        dtRunDate = Format(Now, "dd/MMM/yyyy")

30        g.TextMatrix(0, 1) = "Chart"
40        oSort(3).Caption = "Chart"

50        FillWard

60        Set_Font Me

70        Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

80        intEL = Erl
90        strES = Err.Description
100       LogError "frmCoagSummary", "Form_Load", intEL, strES

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
120       LogError "frmCoagSummary", "osort_Click", intEL, strES


End Sub

Private Sub Print_Click()

          Dim n As Long
          Dim sortby As Long
          Dim OrigOrientation As Long

10        On Error GoTo Print_Click_Error

20        For n = 0 To 3
30            If oSort(n) Then
40                sortby = n
50                Exit For
60            End If
70        Next

80        OrigOrientation = Printer.Orientation
90        Printer.Orientation = vbPRORLandscape

100       Printer.FontName = "Courier New"
110       Printer.FontSize = 12
120       Printer.Print HospName(0)
130       Printer.Print "Coagulation - Summary "; dtRunDate
140       Printer.Print "Sorted by ";
150       Printer.Print Choose(sortby + 1, "Lab Run Number", "Patient Name", _
                               "Clinician", "Chart")
160       Printer.Print

170       Printer.Font.Size = 8
180       GoSub PrintHead
190       For n = 1 To g.Rows - 1
200           If Printer.CurrentY > Printer.Height - 800 Then
210               Printer.NewPage
220               GoSub PrintHead
230           End If
240           Printer.Print Trim(g.TextMatrix(n, 0));                       'lab
250           Printer.Print Tab(10); Trim(g.TextMatrix(n, 1));            'Chart
260           Printer.Print Tab(20); Trim(g.TextMatrix(n, 2));           'dob
270           Printer.Print Tab(32); Trim(Left(g.TextMatrix(n, 3), 29));    'name
280           Printer.Print Tab(60); Trim(Left(g.TextMatrix(n, 4), 24));    'clinician
290           Printer.Print Tab(95); (g.TextMatrix(n, 5));  'ward
300           Printer.Print Tab(115); g.TextMatrix(n, 6);
310           Printer.Print Tab(125); g.TextMatrix(n, 7);
320           Printer.Print Tab(135); g.TextMatrix(n, 8);
330           Printer.Print Tab(145); g.TextMatrix(n, 9);
340           Printer.Print Tab(155); g.TextMatrix(n, 10);
350           Printer.Font.Size = 8
360           Printer.Print
370       Next

380       Printer.EndDoc

390       Printer.Orientation = OrigOrientation

400       Exit Sub

PrintHead:
410       Printer.Print g.TextMatrix(0, 0);                    'lab
420       Printer.Print Tab(10); g.TextMatrix(0, 1);            'Chart
430       Printer.Print Tab(20); g.TextMatrix(0, 2);           'dob
440       Printer.Print Tab(32); g.TextMatrix(0, 3);    'name
450       Printer.Print Tab(60); g.TextMatrix(0, 4);  'clinician
460       Printer.Print Tab(95); g.TextMatrix(0, 5);  'clinician
470       Printer.Print Tab(115); g.TextMatrix(0, 6);
480       Printer.Print Tab(125); g.TextMatrix(0, 7);
490       Printer.Print Tab(135); g.TextMatrix(0, 8);
500       Printer.Print Tab(145); g.TextMatrix(0, 9);
510       Printer.Print Tab(155); g.TextMatrix(0, 10);
520       Printer.Font.Size = 8
530       Printer.Print
540       Return

550       Exit Sub

Print_Click_Error:

          Dim strES As String
          Dim intEL As Integer

560       intEL = Erl
570       strES = Err.Description
580       LogError "frmCoagSummary", "Print_Click", intEL, strES

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)

10        On Error GoTo Toolbar1_ButtonClick_Error

20        If Button.Tag = "Exit" Then
30            Exit_Click
40        ElseIf Button.Tag = "Print" Then
50            Print_Click
60        ElseIf Button.Tag = "View" Then
70            view_Click
80        ElseIf Button.Tag = "Search" Then
90            FillG
100       ElseIf Button.Tag = "Export" Then
110           ExportToExcel
120       Else
130           Validate_Click
140       End If

150       Exit Sub

Toolbar1_ButtonClick_Error:

          Dim strES As String
          Dim intEL As Integer

160       intEL = Erl
170       strES = Err.Description
180       LogError "frmCoagSummary", "Toolbar1_ButtonClick", intEL, strES


End Sub

Private Sub ExportToExcel()

          Dim strHeading As String

10        On Error GoTo ExportToExcel_Error

20        strHeading = "Coagulation Daily Summary" & vbCr & vbCr
30        ExportFlexGrid g, Me, strHeading


40        Exit Sub

ExportToExcel_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmCoagSummary", "ExportToExcel", intEL, strES

End Sub

Private Sub Validate_Click()

          Dim n As Long
          Dim StartRow As Long
          Dim StopRow As Long
          Dim sql As String



10        On Error GoTo Validate_Click_Error

20        If g.row > g.RowSel Then
30            StartRow = g.RowSel
40            StopRow = g.row
50        Else
60            StartRow = g.row
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
240           g.row = n
250           If g <> "Lab #" Then
260               sql = "update demographics set  valid = 1, username = '" & UserName & "' where sampleid = '" & Trim(g) & "' and valid <> 1"
270               Cnxn(0).Execute sql
280               sql = "UPDATE coagresults set valid = 1, Username = '" & AddTicks(UserCode) & "' WHERE sampleid = '" & Trim(g) & "'"
290               Cnxn(0).Execute sql
300           End If
310       Next




320       Exit Sub

Validate_Click_Error:

          Dim strES As String
          Dim intEL As Integer

330       intEL = Erl
340       strES = Err.Description
350       LogError "frmCoagSummary", "Validate_Click", intEL, strES, sql


End Sub

Private Sub view_Click()

10        On Error GoTo view_Click_Error

20        If g.TextMatrix(g.RowSel, 0) = "" Then Exit Sub

30        g.Col = 0
40        If g = "Lab #" Then Exit Sub
50        g.Col = 1

60        With frmViewResults
70            .lblChart = g
80            g.Col = 3
90            .lblName = g
100           g.Col = 0
110           .lblSampleID = g
              '  .trundatetime = Format(dtRunDate, "dd/mmm/yyyy")
120           .Show 1
130       End With

140       Exit Sub

view_Click_Error:

          Dim strES As String
          Dim intEL As Integer

150       intEL = Erl
160       strES = Err.Description
170       LogError "frmCoagSummary", "view_Click", intEL, strES

End Sub


Private Sub FillWard()
          Dim sql As String
          Dim tb As Recordset
          Dim strHospitalCode As String

10        On Error GoTo FillWard_Error

20        strHospitalCode = ListCodeFor("HO", HospName(0))


30        sql = "SELECT * from wards WHERE hospitalcode = '" & strHospitalCode & "' order by text"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        Do While Not tb.EOF
70            cmbWard.AddItem Trim(tb!Text)
80            tb.MoveNext
90        Loop

100       cmbWard.AddItem "All", 0
110       cmbWard.ListIndex = 0

120       Exit Sub

FillWard_Error:

          Dim strES As String
          Dim intEL As Integer

130       intEL = Erl
140       strES = Err.Description
150       LogError "frmCoagSummary", "FillWard", intEL, strES, sql

End Sub
