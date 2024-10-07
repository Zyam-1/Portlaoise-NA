VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMicroStatsGeneral 
   Caption         =   "NetAcquire - Microbiology Statistics"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11475
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   11475
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton bcancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   1000
      Left            =   10230
      Picture         =   "frmMicroStatsGeneral.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   660
      Width           =   1100
   End
   Begin VB.CommandButton bprint 
      Appearance      =   0  'Flat
      Caption         =   "&Print"
      Height          =   1000
      Left            =   8970
      Picture         =   "frmMicroStatsGeneral.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   660
      Width           =   1100
   End
   Begin VB.PictureBox SSPanel1 
      Height          =   1395
      Index           =   1
      Left            =   180
      ScaleHeight     =   1335
      ScaleWidth      =   6225
      TabIndex        =   15
      Top             =   240
      Width           =   6285
      Begin VB.OptionButton oBetween 
         Caption         =   "Custom"
         Height          =   195
         Index           =   7
         Left            =   4770
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   1155
      End
      Begin VB.OptionButton oBetween 
         Caption         =   "Last Week"
         Height          =   225
         Index           =   0
         Left            =   180
         TabIndex        =   4
         Top             =   810
         Width           =   1095
      End
      Begin VB.OptionButton oBetween 
         Caption         =   "Last Month"
         Height          =   195
         Index           =   1
         Left            =   1455
         TabIndex        =   5
         Top             =   540
         Width           =   1155
      End
      Begin VB.OptionButton oBetween 
         Caption         =   "Last Full Month"
         Height          =   195
         Index           =   2
         Left            =   1455
         TabIndex        =   6
         Top             =   810
         Width           =   1395
      End
      Begin VB.OptionButton oBetween 
         Caption         =   "Last Quarter"
         Height          =   195
         Index           =   3
         Left            =   3105
         TabIndex        =   7
         Top             =   570
         Width           =   1185
      End
      Begin VB.OptionButton oBetween 
         Caption         =   "Last Full Quarter"
         Height          =   195
         Index           =   4
         Left            =   3105
         TabIndex        =   8
         Top             =   840
         Width           =   1515
      End
      Begin VB.OptionButton oBetween 
         Caption         =   "Year To Date"
         Height          =   195
         Index           =   5
         Left            =   4770
         TabIndex        =   9
         Top             =   840
         Width           =   1305
      End
      Begin VB.OptionButton oBetween 
         Caption         =   "Today"
         Height          =   195
         Index           =   6
         Left            =   180
         TabIndex        =   3
         Top             =   570
         Width           =   795
      End
      Begin MSComCtl2.DTPicker calFrom 
         Height          =   375
         Left            =   645
         TabIndex        =   0
         Top             =   150
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   661
         _Version        =   393216
         Format          =   193265665
         CurrentDate     =   38037
      End
      Begin MSComCtl2.DTPicker calTo 
         Height          =   375
         Left            =   2670
         TabIndex        =   1
         Top             =   150
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   661
         _Version        =   393216
         Format          =   193265665
         CurrentDate     =   38037
      End
      Begin VB.Label Label1 
         Caption         =   "From"
         Height          =   330
         Left            =   195
         TabIndex        =   17
         Top             =   210
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "To"
         Height          =   285
         Left            =   2400
         TabIndex        =   16
         Top             =   210
         Width           =   285
      End
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Start"
      Height          =   1000
      Left            =   6660
      Picture         =   "frmMicroStatsGeneral.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   660
      Width           =   1100
   End
   Begin VB.CommandButton cmdExcel 
      Caption         =   "Export to Excel"
      Height          =   1000
      Left            =   7815
      Picture         =   "frmMicroStatsGeneral.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   660
      Width           =   1100
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   6075
      Left            =   180
      TabIndex        =   14
      Top             =   1740
      Width           =   11145
      _ExtentX        =   19659
      _ExtentY        =   10716
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      BackColor       =   -2147483624
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      GridLines       =   2
      SelectionMode   =   1
      FormatString    =   $"frmMicroStatsGeneral.frx":0C28
   End
   Begin MSComctlLib.ProgressBar PBar 
      Height          =   195
      Left            =   180
      TabIndex        =   19
      Top             =   7860
      Width           =   11145
      _ExtentX        =   19659
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label lblExcelInfo 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exporting..."
      Height          =   345
      Left            =   7800
      TabIndex        =   18
      Top             =   300
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "frmMicroStatsGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub bcancel_Click()

10        On Error GoTo bCancel_Click_Error
20        Unload Me


30        Exit Sub

bCancel_Click_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmMicroStatsGeneral", "bcancel_Click", intEL, strES, sql

End Sub

Private Sub bprint_Click()

          Dim i As Integer

10        On Error GoTo bprint_Click_Error

20        With Printer
30            .Orientation = vbPRORPortrait
40            .FontName = "Courier New"

50        End With

60        With g

              'Title
70            PrintText "", , , , , , True
80            PrintText FormatString("Microbiology Statistics", 99, , AlignCenter), 10, True, , , , True
90            PrintText FormatString("From " & Format(calFrom.Value, "dd/MM/yyyy") & " to " & Format(calTo.Value, "dd/MM/yyyy"), 99, , AlignCenter), 10, True, , , , True
100           PrintText "", , , , , , True
110           PrintText String(248, "-"), 4, , , , , True
120           PrintText "", , , , , , True
              'Heading

130           PrintText FormatString("", 4, , AlignLeft), , True
140           PrintText FormatString(.TextMatrix(0, 0), 50, , AlignLeft), , True
150           PrintText FormatString(.TextMatrix(0, 1), 10, , AlignLeft), , True
160           PrintText FormatString(.TextMatrix(0, 2), 10, , AlignLeft), , True
170           PrintText FormatString(.TextMatrix(0, 3), 10, , AlignLeft), , True
180           PrintText FormatString(.TextMatrix(0, 4), 20, , AlignRight), , True, , , , True
190           PrintText String(248, "-"), 4, , , , , True
              'Data
200           For i = 1 To g.Rows - 1
210               PrintText FormatString("", 4, , AlignLeft), , True
220               PrintText FormatString(.TextMatrix(i, 0), 50, , AlignLeft), , IIf(i = g.Rows - 1, True, False)
230               PrintText FormatString(.TextMatrix(i, 1), 10, , AlignLeft), , IIf(i = g.Rows - 1, True, False)
240               PrintText FormatString(.TextMatrix(i, 2), 10, , AlignLeft), , IIf(i = g.Rows - 1, True, False)
250               PrintText FormatString(.TextMatrix(i, 3), 10, , AlignLeft), , IIf(i = g.Rows - 1, True, False)
260               PrintText FormatString(.TextMatrix(i, 4), 20, , AlignRight), , True, , , , True
270               PrintText String(248, "-"), 4, , , , , True
280           Next i

290           Printer.EndDoc
300       End With

310       Exit Sub

bprint_Click_Error:

          Dim strES As String
          Dim intEL As Integer

320       intEL = Erl
330       strES = Err.Description
340       LogError "frmMicroStatsGeneral", "bprint_Click", intEL, strES

End Sub

Private Sub cmdExcel_Click()
          Dim strTitle As String
10        On Error GoTo cmdExcel_Click_Error

20        If g.Rows < 2 Then
30            iMsg "Nothing to export", vbInformation
40            Exit Sub
50        End If

60        strTitle = "Microbiology Statistics" & vbCr
70        strTitle = "From " & calFrom.Value & " to " & calTo.Value & vbCr

80        ExportFlexGrid g, Me, strTitle

90        Exit Sub

cmdExcel_Click_Error:

          Dim strES As String
          Dim intEL As Integer

100       intEL = Erl
110       strES = Err.Description
120       LogError "frmMicroStatsGeneral", "cmdExcel_Click", intEL, strES, sql

End Sub

Private Sub calFrom_Click()
10        On Error GoTo calFrom_Click_Error

20        oBetween(7).Value = True

30        Exit Sub

calFrom_Click_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmMicroStatsGeneral", "calFrom_Click", intEL, strES


End Sub

Private Sub calFrom_DropDown()

10        On Error GoTo calFrom_DropDown_Error
20        oBetween(7).Value = True


30        Exit Sub

calFrom_DropDown_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmMicroStatsGeneral", "calFrom_DropDown", intEL, strES

End Sub

Private Sub calTo_Click()

10        On Error GoTo calTo_Click_Error
20        oBetween(7).Value = True


30        Exit Sub

calTo_Click_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmMicroStatsGeneral", "calTo_Click", intEL, strES

End Sub

Private Sub calTo_DropDown()

10        On Error GoTo calTo_DropDown_Error
20        oBetween(7).Value = True


30        Exit Sub

calTo_DropDown_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmMicroStatsGeneral", "calTo_DropDown", intEL, strES

End Sub

Private Sub cmdSearch_Click()
10        On Error GoTo cmdSearch_Click_Error

20        FillGrid

30        Exit Sub

cmdSearch_Click_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmMicroStatsGeneral", "cmdSearch_Click", intEL, strES


End Sub

Private Sub FillGrid()

10        On Error GoTo FillGrid_Error

          Dim tb As Recordset
          Dim sql As String
          Dim MicroOffset As String
          Dim OldSite As String
          Dim NewSite As String
          Dim i As Integer
          Dim GpTotal As Long
          Dim ClinTotal As Long
          Dim WardTotal As Long

20        With g
30            .Clear
40            .Rows = 1
50            .FormatString = "<Micro Site                                                                                                 |<GP                      |<Wards                   |<Totals                       "
60        End With

70        MicroOffset = GetOptionSetting("MICROOFFSET", "0")
80        If MicroOffset = "0" Or MicroOffset = "" Then Exit Sub

90        If DateDiff("d", calFrom.Value, calTo.Value) < 0 Then
100           iMsg "Invalid date selection, Please amend", vbInformation
110           Exit Sub
120       End If

130       sql = "SELECT COUNT(M.SampleID) AS Cnt, M.Site, MAX(D.GP) GP,  '' Ward from MicroSiteDetails M " & _
                "INNER JOIN " & _
                "(SELECT SampleID, GP FROM Demographics WHERE SampleID > " & MicroOffset & " AND " & _
                "RunDate between '" & Format(calFrom.Value, "dd/MMM/yyyy") & "' and ' " & Format(calTo.Value, "dd/MMM/yyyy") & "' AND Ward = 'gp') D " & _
                "ON M.SampleID = D.SampleID " & _
                "GROUP BY M.Site " & _
                "UNION " & _
                "SELECT COUNT(M.SampleID) AS Cnt, M.Site, '' GP, '' Ward from MicroSiteDetails M " & _
                "INNER JOIN " & _
                "(SELECT SampleID, Clinician FROM Demographics WHERE SampleID >  " & MicroOffset & " AND  " & _
                "RunDate between '" & Format(calFrom.Value, "dd/MMM/yyyy") & "' and ' " & Format(calTo.Value, "dd/MMM/yyyy") & "' AND coalesce(Clinician,'') <> '') D " & _
                "ON M.SampleID = D.SampleID " & _
                "GROUP BY M.Site " & _
                "UNION " & _
                "SELECT COUNT(M.SampleID) AS Cnt, M.Site, '' GP,  MAX(Ward) Ward from MicroSiteDetails M " & _
                "INNER JOIN " & _
                "(SELECT SampleID, Ward FROM Demographics WHERE SampleID >  " & MicroOffset & "  " & _
                "AND RunDate between '" & Format(calFrom.Value, "dd/MMM/yyyy") & "' and ' " & Format(calTo.Value, "dd/MMM/yyyy") & "' AND coalesce(Ward,'') <> '' AND coalesce(Ward,'') <> 'GP') D " & _
                "ON M.SampleID = D.SampleID " & _
                "GROUP BY M.Site " & _
                "ORDER By M.Site"

140       Set tb = New Recordset
150       RecOpenServer 0, tb, sql
160       OldSite = "*"
170       NewSite = tb!Site & ""
180       If Not tb.EOF Then
190           While Not tb.EOF
200               With g
210                   If NewSite <> OldSite Then
220                       .AddItem tb!Site & ""

230                   End If
240                   If tb!GP & "" <> "" Then
250                       .row = .Rows - 1
260                       .Col = 1
270                       .TextMatrix(.row, .Col) = tb!Cnt
280                       GpTotal = GpTotal + tb!Cnt
290                   ElseIf tb!Ward & "" <> "" Then
300                       .row = .Rows - 1
310                       .Col = 2
320                       .TextMatrix(.row, .Col) = tb!Cnt
330                       WardTotal = WardTotal + tb!Cnt
'340                   ElseIf tb!Clinician & "" <> "" Then
'350                       .row = .Rows - 1
'360                       .Col = 3
'370                       .TextMatrix(.row, .Col) = tb!Cnt
'380                       ClinTotal = ClinTotal + tb!Cnt
390                   End If
400                   tb.MoveNext
410                   If Not tb.EOF Then
420                       OldSite = NewSite
430                       NewSite = tb!Site & ""
440                       .row = .Rows - 1
450                       .Col = 3
460                       .CellFontBold = True
470                       .CellForeColor = &HC0&
480                       .CellBackColor = &H80000013
490                       .TextMatrix(.row, .Col) = Val(.TextMatrix(.row, 1)) + Val(.TextMatrix(.row, 2)) '+ Val(.TextMatrix(.row, 3))
500                   End If
510               End With
520           Wend

530           g.AddItem "Total Samples" & vbTab & _
                        GpTotal & vbTab & _
                        WardTotal & vbTab & _
                        GpTotal + WardTotal '+ ClinTotal
540           For i = 0 To g.Cols - 1
550               g.row = g.Rows - 1
560               g.Col = i
570               g.CellFontBold = True
580               g.CellBackColor = &H80000013
590               g.CellForeColor = &HC0&
600           Next



610       End If


620       Exit Sub

FillGrid_Error:

          Dim strES As String
          Dim intEL As Integer

630       intEL = Erl
640       strES = Err.Description
650       LogError "frmMicroStatsGeneral", "FillGrid", intEL, strES, sql

End Sub

'Private Sub FillGrid()
'
'    On Error GoTo FillGrid_Error
'
'    Dim tb As Recordset
'    Dim sql As String
'    Dim MicroOffset As String
'    Dim OldSite As String
'    Dim NewSite As String
'    Dim i As Integer
'    Dim GpTotal As Long
'    Dim ClinTotal As Long
'    Dim WardTotal As Long
'
'    With g
'        .Clear
'        .Rows = 1
'        .FormatString = "<Micro Site                                                                                                 |<GP                      |<Wards                   |<Clin                    |<Totals                       "
'    End With
'
'    MicroOffset = GetOptionSetting("MICROOFFSET", "0")
'    If MicroOffset = "0" Or MicroOffset = "" Then Exit Sub
'
'    If DateDiff("d", calFrom.Value, calTo.Value) < 0 Then
'        iMsg "Invalid date selection, Please amend", vbInformation
'        Exit Sub
'    End If
'
'
'
'    sql = "SELECT COUNT(M.SampleID) AS Cnt, M.Site, MAX(D.GP) GP, '' Clinician, '' Ward from MicroSiteDetails M " & _
'          "INNER JOIN " & _
'          "(SELECT SampleID, GP FROM Demographics WHERE SampleID > " & MicroOffset & " AND " & _
'          "RunDate between '" & Format(calFrom.Value, "dd/MMM/yyyy") & "' and ' " & Format(calTo.Value, "dd/MMM/yyyy") & "' AND Ward = 'gp') D " & _
'          "ON M.SampleID = D.SampleID " & _
'          "GROUP BY M.Site " & _
'          "UNION " & _
'          "SELECT COUNT(M.SampleID) AS Cnt, M.Site, '' GP, Max(D.Clinician) Clinician, '' Ward from MicroSiteDetails M " & _
'          "INNER JOIN " & _
'          "(SELECT SampleID, Clinician FROM Demographics WHERE SampleID >  " & MicroOffset & " AND  " & _
'          "RunDate between '" & Format(calFrom.Value, "dd/MMM/yyyy") & "' and ' " & Format(calTo.Value, "dd/MMM/yyyy") & "' AND coalesce(Clinician,'') <> '') D " & _
'          "ON M.SampleID = D.SampleID " & _
'          "GROUP BY M.Site " & _
'          "UNION " & _
'          "SELECT COUNT(M.SampleID) AS Cnt, M.Site, '' GP, '' Clinician, MAX(Ward) Ward from MicroSiteDetails M " & _
'          "INNER JOIN " & _
'          "(SELECT SampleID, Ward FROM Demographics WHERE SampleID >  " & MicroOffset & "  " & _
'          "AND RunDate between '" & Format(calFrom.Value, "dd/MMM/yyyy") & "' and ' " & Format(calTo.Value, "dd/MMM/yyyy") & "' AND coalesce(Ward,'') <> '' AND coalesce(Ward,'') <> 'GP') D " & _
'          "ON M.SampleID = D.SampleID " & _
'          "GROUP BY M.Site " & _
'          "ORDER By M.Site"
'
'    Set tb = New Recordset
'    RecOpenServer 0, tb, sql
'    OldSite = "*"
'    NewSite = tb!Site & ""
'    If Not tb.EOF Then
'        While Not tb.EOF
'            With g
'                If NewSite <> OldSite Then
'                    .AddItem tb!Site & ""
'
'                End If
'                If tb!GP & "" <> "" Then
'                    .row = .Rows - 1
'                    .Col = 1
'                    .TextMatrix(.row, .Col) = tb!Cnt
'                    GpTotal = GpTotal + tb!Cnt
'                ElseIf tb!Ward & "" <> "" Then
'                    .row = .Rows - 1
'                    .Col = 2
'                    .TextMatrix(.row, .Col) = tb!Cnt
'                    WardTotal = WardTotal + tb!Cnt
'                ElseIf tb!Clinician & "" <> "" Then
'                    .row = .Rows - 1
'                    .Col = 3
'                    .TextMatrix(.row, .Col) = tb!Cnt
'                    ClinTotal = ClinTotal + tb!Cnt
'                End If
'                tb.MoveNext
'                If Not tb.EOF Then
'                    OldSite = NewSite
'                    NewSite = tb!Site & ""
'                    .row = .Rows - 1
'                    .Col = 4
'                    .CellFontBold = True
'                    .CellForeColor = &HC0&
'                    .CellBackColor = &H80000013
'                    .TextMatrix(.row, .Col) = Val(.TextMatrix(.row, 1)) + Val(.TextMatrix(.row, 2)) + Val(.TextMatrix(.row, 3))
'                End If
'            End With
'        Wend
'
'        g.AddItem "Total Samples" & vbTab & _
'                  GpTotal & vbTab & _
'                  WardTotal & vbTab & _
'                  ClinTotal & vbTab & _
'                  GpTotal + WardTotal + ClinTotal
'        For i = 0 To g.Cols - 1
'            g.row = g.Rows - 1
'            g.Col = i
'            g.CellFontBold = True
'            g.CellBackColor = &H80000013
'            g.CellForeColor = &HC0&
'        Next
'
'
'
'    End If
'
'
'    Exit Sub
'
'FillGrid_Error:
'
'    Dim strES As String
'    Dim intEL As Integer
'
'    intEL = Erl
'    strES = Err.Description
'    LogError "frmMicroStatsGeneral", "FillGrid", intEL, strES, sql
'
'End Sub



Private Sub Form_Load()

10        On Error GoTo Form_Load_Error

20        calFrom.Value = Now - 1
30        calTo.Value = Now


40        Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmMicroStatsGeneral", "Form_Load", intEL, strES

End Sub

Private Sub oBetween_Click(Index As Integer)

          Dim upto As String

10        On Error GoTo oBetween_Click_Error

20        calFrom = BetweenDates(Index, upto)
30        calTo = upto

40        bReCalc.Visible = True

50        Exit Sub

oBetween_Click_Error:

          Dim strES As String
          Dim intEL As Integer



60        intEL = Erl
70        strES = Err.Description
80        LogError "frmTotals", "oBetween_Click", intEL, strES


End Sub

