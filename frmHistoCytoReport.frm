VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.ocx"
Begin VB.Form frmHistoCytoReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NetAcquire --- Histology Yearly Report"
   ClientHeight    =   8325
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9210
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   9210
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox fraProgress 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   2220
      ScaleHeight     =   525
      ScaleWidth      =   3825
      TabIndex        =   12
      Top             =   4560
      Visible         =   0   'False
      Width           =   3855
      Begin MSComctlLib.ProgressBar pbProgress 
         Height          =   255
         Left            =   0
         TabIndex        =   13
         Top             =   240
         Width           =   3840
         _ExtentX        =   6773
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         Min             =   1
      End
      Begin VB.Label lblProgress 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Fetching results ..."
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   3840
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   1875
      Left            =   300
      TabIndex        =   0
      Top             =   300
      Width           =   8595
      _Version        =   65536
      _ExtentX        =   15161
      _ExtentY        =   3307
      _StockProps     =   15
      BackColor       =   14215660
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelInner      =   1
      Begin VB.CommandButton cmdExcel 
         Caption         =   "Export to Excel"
         Height          =   870
         Left            =   5550
         Picture         =   "frmHistoCytoReport.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   600
         Width           =   1185
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
         Height          =   870
         Left            =   2640
         Picture         =   "frmHistoCytoReport.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   600
         Width           =   1185
      End
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "Exit"
         Height          =   870
         Left            =   7155
         Picture         =   "frmHistoCytoReport.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   600
         Width           =   1185
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print"
         Height          =   870
         Left            =   4245
         Picture         =   "frmHistoCytoReport.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   600
         Width           =   1185
      End
      Begin VB.ComboBox cmbYear 
         Height          =   315
         Left            =   780
         TabIndex        =   3
         Top             =   300
         Width           =   1335
      End
      Begin VB.OptionButton optHistoSamples 
         Caption         =   "Histology Samples"
         Height          =   195
         Left            =   780
         TabIndex        =   2
         Top             =   900
         Value           =   -1  'True
         Width           =   1755
      End
      Begin VB.OptionButton optCytoSamples 
         Caption         =   "Cytology Samples"
         Height          =   195
         Left            =   780
         TabIndex        =   1
         Top             =   1215
         Width           =   1755
      End
      Begin VB.Image imgYearInfo 
         Height          =   360
         Left            =   2100
         Picture         =   "frmHistoCytoReport.frx":0C28
         ToolTipText     =   "Shows only avaiable years for histology and cytology"
         Top             =   277
         Width           =   360
      End
      Begin VB.Label lblExcelInfo 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Exporting..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5550
         TabIndex        =   10
         Top             =   180
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Year"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   330
      End
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   5295
      Left            =   300
      TabIndex        =   5
      Top             =   2280
      Width           =   8595
      _ExtentX        =   15161
      _ExtentY        =   9340
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      GridLines       =   3
      GridLinesFixed  =   3
      SelectionMode   =   1
      FormatString    =   $"frmHistoCytoReport.frx":1312
   End
   Begin ComctlLib.ProgressBar pb 
      Height          =   285
      Left            =   300
      TabIndex        =   11
      Top             =   7740
      Width           =   8595
      _ExtentX        =   15161
      _ExtentY        =   503
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   1
   End
End
Attribute VB_Name = "frmHistoCytoReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmbYear_KeyPress(KeyAscii As Integer)
10        KeyAscii = 0
End Sub

Private Sub cmdExcel_Click()

          Dim strHeading As String

10        On Error GoTo cmdExcel_Click_Error

20        If g.Rows = 2 Then
30            iMsg "Nothing to export", vbInformation
40            Exit Sub
50        End If

60        cmdExcel.Enabled = False

70        strHeading = IIf(optHistoSamples, "Histology", "Cytology") & " Yearly Report" & vbCr
80        strHeading = strHeading & "Year Ending 31/12/"" & cmbYear & vbCr & vbCr"
90        ExportFlexGrid g, Me, strHeading

100       cmdExcel.Enabled = True

110       Exit Sub

cmdExcel_Click_Error:

          Dim strES As String
          Dim intEL As Integer

120       cmdExcel.Enabled = True
130       intEL = Erl
140       strES = Err.Description
150       LogError "frmHistoCytoReport", "cmdExcel_Click", intEL, strES

End Sub

Private Sub cmdExit_Click()
10        Unload Me
End Sub

Private Sub cmdPrint_Click()

          Dim i As Integer
          Dim J As Integer
          Dim CurrentPage As Integer
          Dim TotalPages As Integer
          Const LinesPerPage As Integer = 75
          Dim StartIndex As Integer
          Dim EndIndex As Integer

10        On Error GoTo cmdPrint_Click_Error

20        If g.Rows = 2 Then
30            iMsg "Nothing to print", vbInformation
40            Exit Sub
50        End If

60        cmdPrint.Enabled = False

70        CurrentPage = 1
80        TotalPages = (g.Rows \ LinesPerPage) + 1

          'Printer.Orientation = vbPRORPortrait
90        Printer.FontName = "Courier New"


100       StartIndex = 0
110       EndIndex = 0
120       For i = 1 To TotalPages
130           StartIndex = EndIndex + 1
140           If (g.Rows Mod LinesPerPage) < LinesPerPage And i = TotalPages Then
150               EndIndex = EndIndex + ((g.Rows - 1) Mod LinesPerPage)
160           Else
170               EndIndex = EndIndex + LinesPerPage
180           End If
190           PrintHeaderPortrait CurrentPage, TotalPages
              'PRINT HEADINGS
200           Printer.FontBold = True
210           Printer.FontSize = 9
220           Printer.Print FormatString("", 0, "|");
230           Printer.Print FormatString(g.TextMatrix(0, 1), 15, "|", AlignLeft);   'Sample ID
240           Printer.Print FormatString(g.TextMatrix(0, 2), 40, "|", AlignLeft);   'Patient Name
250           Printer.Print FormatString(g.TextMatrix(0, 3), 15, "|", AlignLeft);   'DoB
260           Printer.Print FormatString(g.TextMatrix(0, 4), 16, "|", AlignLeft);   'Chart
270           Printer.Print FormatString(g.TextMatrix(0, 5), 16, "|", AlignLeft)    'Sample Date
              'PRINT LINE
280           Printer.FontBold = False
290           Printer.FontSize = 2
300           Printer.Print String(500, "-")
310           For J = StartIndex To EndIndex
                  'PRINT DATA
320               Printer.FontBold = False
330               Printer.FontSize = 9
340               Printer.Print FormatString("", 0, "|");
350               Printer.Print FormatString(g.TextMatrix(J, 1), 15, "|", AlignLeft);
360               Printer.Print FormatString(g.TextMatrix(J, 2), 40, "|", AlignLeft);
370               Printer.Print FormatString(g.TextMatrix(J, 3), 15, "|", AlignLeft);
380               Printer.Print FormatString(g.TextMatrix(J, 4), 16, "|", AlignLeft);
390               Printer.Print FormatString(g.TextMatrix(J, 5), 16, "|", AlignLeft)
400           Next J
410           PrintFooterPortrait
420           Printer.NewPage
430           Printer.FontBold = False
440           Printer.FontSize = 9
450           CurrentPage = CurrentPage + 1
460       Next i

470       Printer.EndDoc

480       cmdPrint.Enabled = True

490       Exit Sub

cmdPrint_Click_Error:

          Dim strES As String
          Dim intEL As Integer

500       cmdPrint.Enabled = False
510       intEL = Erl
520       strES = Err.Description
530       LogError "frmHistoCytoReport", "cmdPrint_Click", intEL, strES

End Sub

Private Sub PrintFooterPortrait()

10        On Error GoTo PrintFooterPortrait_Error

          'PRINT LINE
20        Printer.FontBold = False
30        Printer.FontSize = 2
40        Printer.Print String(500, "-")
          'PRINT PRINTED BY
50        Printer.FontSize = 7
60        Printer.FontBold = False
70        Printer.Print FormatString("Printed by " & Username, 140, , AlignCenter)
          'PRINT LINE
80        Printer.FontBold = False
90        Printer.FontSize = 2
100       Printer.Print String(500, "-")

110       Exit Sub

PrintFooterPortrait_Error:

          Dim strES As String
          Dim intEL As Integer

120       intEL = Erl
130       strES = Err.Description
140       LogError "frmHistoCytoReport", "PrintFooterPortrait", intEL, strES

End Sub
Private Sub PrintHeaderPortrait(CurrentPage As Integer, TotalPages As Integer)

10        On Error GoTo PrintHeaderPortrait_Error

          'PRINT HEADING
20        Printer.FontSize = 10
30        Printer.FontBold = True
40        Printer.Print
50        Printer.Print FormatString(IIf(optHistoSamples, "Histology", "Cytology") & " Yearly Report", 99, , AlignCenter)
60        Printer.Print FormatString("Year Ending 31/12/" & Trim(cmbYear), 99, , AlignCenter)
70        Printer.Print
          'PRINT LINE
80        Printer.FontBold = False
90        Printer.FontSize = 2
100       Printer.Print String(500, "-")
          'PRINT PRINTED TIME AND PAGE M OF N
110       Printer.FontSize = 7
120       Printer.FontBold = False
130       Printer.Print FormatString("Printed on " & Format$(Now, "dd/mm/yy") & " at  " & Format$(Now, "hh:mm"), 85, , AlignRight) & _
                        FormatString("Page " & CurrentPage & " Of " & TotalPages, 55, , AlignRight)
          'PRINT LINE
140       Printer.FontBold = False
150       Printer.FontSize = 2
160       Printer.Print String(500, "-")

170       Exit Sub

PrintHeaderPortrait_Error:

          Dim strES As String
          Dim intEL As Integer

180       intEL = Erl
190       strES = Err.Description
200       LogError "frmHistoCytoReport", "PrintHeaderPortrait", intEL, strES

End Sub
Private Sub cmdSearch_Click()

          Dim tb As Recordset
          Dim sql As String
          Dim s As String

10        On Error GoTo cmdSearch_Click_Error

20        If cmbYear = "" Then
30            iMsg "Please select year first", vbInformation
40            Exit Sub
50        End If

60        cmdSearch.Enabled = False
70        sql = "SELECT D.SampleID - @offset - (Reverse(DatePart(""yyyy"",RunDate)) * 1000) As SampleIDWithoutOffset, " & _
                "DatePart(""yyyy"",D.RunDate) As SYear, " & _
                "D.SampleID , D.PatName, D.Dob, D.Chart, D.SampleDate " & _
                "FROM Demographics D INNER JOIN @departmentResults R " & _
                "ON D.SampleID = R.SampleID " & _
                "WHERE DatePart(""yyyy"",D.RunDate) = @hyear " & _
                "ORDER BY D.RunDate DESC"

80        sql = Replace(sql, "@offset", IIf(optHistoSamples.Value = True, SysOptHistoOffset(0), SysOptCytoOffset(0)))
90        sql = Replace(sql, "@hyear", cmbYear)
100       sql = Replace(sql, "@department", IIf(optHistoSamples, "Histo", "Cyto"))


110       Set tb = New Recordset
120       RecOpenClient 0, tb, sql
130       If Not tb.EOF Then
140           pbProgress.Max = tb.RecordCount + 1
150           g.Visible = False
160           fraProgress.Visible = True
170           While Not tb.EOF
180               pbProgress.Value = pbProgress.Value + 1
190               lblProgress = "Fetch results ... (" & Int(pbProgress.Value * 100 / pbProgress.Max) & " %)"
200               lblProgress.Refresh
210               s = vbTab & tb!SampleIDWithoutOffset & "" & vbTab & _
                      tb!PatName & "" & vbTab & _
                      tb!Dob & "" & vbTab & _
                      tb!Chart & "" & vbTab & _
                      tb!SampleDate
220               g.AddItem s
230               tb.MoveNext
240           Wend
250           fraProgress.Visible = False
260           g.Visible = True
270           If g.Rows > 2 Then g.RemoveItem 1
280           pbProgress.Value = 1
290       End If

300       cmdSearch.Enabled = True

310       Exit Sub

cmdSearch_Click_Error:

320       g.Visible = True
330       fraProgress.Visible = False
340       cmdSearch.Enabled = True
          Dim strES As String
          Dim intEL As Integer

350       intEL = Erl
360       strES = Err.Description
370       LogError "frmHistoCytoReport", "cmdSearch_Click", intEL, strES, sql

End Sub

Private Sub Form_Load()
10        PopulateYear
20        g.ColWidth(0) = 0
End Sub


Private Sub PopulateYear()

          Dim tb As Recordset
          Dim sql As String


10        On Error GoTo PopulateYear_Error

20        sql = "SELECT DISTINCT HYear FROM @departmentResults ORDER BY HYear DESC"
30        sql = Replace(sql, "@department", IIf(optHistoSamples = True, "Histo", "Cyto"))

40        Set tb = New Recordset
50        RecOpenClient 0, tb, sql
60        cmbYear.Clear
70        If Not tb.EOF Then

80            While Not tb.EOF

90                cmbYear.AddItem tb!Hyear & ""
100               tb.MoveNext
110           Wend
120       End If

130       Exit Sub

PopulateYear_Error:

          Dim strES As String
          Dim intEL As Integer

140       intEL = Erl
150       strES = Err.Description
160       LogError "frmHistoCytoReport", "PopulateYear", intEL, strES, sql

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
90            Exit Sub
100       End If


110       Exit Sub

g_Click_Error:

          Dim strES As String
          Dim intEL As Integer

120       intEL = Erl
130       strES = Err.Description
140       LogError "frmHistoCytoReport", "g_Click", intEL, strES


End Sub

Private Sub optCytoSamples_Click()
10        PopulateYear
End Sub

Private Sub optHistoSamples_Click()
10        PopulateYear
End Sub
