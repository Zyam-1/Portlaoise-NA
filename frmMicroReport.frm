VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMicroReport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Micro Reports"
   ClientHeight    =   10170
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   12165
   Icon            =   "frmMicroReport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10170
   ScaleWidth      =   12165
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdPrint 
      Appearance      =   0  'Flat
      Caption         =   "&Print"
      Height          =   750
      Left            =   9720
      Picture         =   "frmMicroReport.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   9270
      Width           =   1065
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   7065
      Left            =   12180
      TabIndex        =   16
      Top             =   1980
      Width           =   8355
      _ExtentX        =   14737
      _ExtentY        =   12462
      _Version        =   393216
      Rows            =   7
      Cols            =   7
      FixedRows       =   6
      FocusRect       =   0
      HighLight       =   0
      AllowUserResizing=   3
   End
   Begin VB.Frame Frame3 
      Caption         =   "Report"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5145
      Left            =   225
      TabIndex        =   12
      Top             =   3945
      Width           =   11775
      Begin RichTextLib.RichTextBox txtReport 
         Height          =   4695
         Left            =   90
         TabIndex        =   14
         Top             =   315
         Width           =   11475
         _ExtentX        =   20241
         _ExtentY        =   8281
         _Version        =   393217
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"frmMicroReport.frx":0614
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   11715
      Top             =   180
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Exit"
      Height          =   750
      Left            =   10935
      Picture         =   "frmMicroReport.frx":0694
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   9270
      Width           =   1065
   End
   Begin VB.Frame Frame1 
      Height          =   1395
      Left            =   240
      TabIndex        =   0
      Top             =   345
      Width           =   11835
      Begin VB.Label lblClDetails 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   675
         TabIndex        =   13
         Top             =   1020
         Width           =   10605
      End
      Begin VB.Label lblName 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   675
         TabIndex        =   9
         Top             =   225
         Width           =   3540
      End
      Begin VB.Label lblChart 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   8220
         TabIndex        =   8
         Top             =   225
         Width           =   1200
      End
      Begin VB.Label lblDoB 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   5970
         TabIndex        =   7
         Top             =   225
         Width           =   1230
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   6
         Top             =   255
         Width           =   420
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Chart"
         Height          =   195
         Index           =   1
         Left            =   7740
         TabIndex        =   5
         Top             =   255
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Date of Birth"
         Height          =   195
         Index           =   1
         Left            =   4965
         TabIndex        =   4
         Top             =   255
         Width           =   885
      End
      Begin VB.Label lblSex 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   10230
         TabIndex        =   3
         Top             =   225
         Width           =   1020
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Sex"
         Height          =   195
         Index           =   1
         Left            =   9900
         TabIndex        =   2
         Top             =   255
         Width           =   270
      End
      Begin VB.Label lblDemogComment 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   675
         TabIndex        =   1
         Top             =   660
         Width           =   10605
      End
   End
   Begin MSComctlLib.ProgressBar PBar 
      Height          =   195
      Left            =   240
      TabIndex        =   11
      Top             =   60
      Width           =   11685
      _ExtentX        =   20611
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSFlexGridLib.MSFlexGrid grdSID 
      Height          =   1935
      Left            =   225
      TabIndex        =   15
      Top             =   1860
      Width           =   11835
      _ExtentX        =   20876
      _ExtentY        =   3413
      _Version        =   393216
      Cols            =   6
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      FormatString    =   $"frmMicroReport.frx":099E
   End
End
Attribute VB_Name = "frmMicroReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Activated As Boolean

Private Type ABResult
    Antibiotic As String
    ReportName As String
    Result(1 To 8) As String
    Report(1 To 8) As Boolean
    Group(1 To 8) As String
    Qualifier(1 To 8) As String
End Type


Private pPatName As String
Private pPatChart As String
Private pPatDoB As String
Private pPatSex As String
Private pPatWard As String
Private pPatClinician As String
Private pPatGP As String

Private SortOrder As Boolean

Private Type LineText
    LineType As String
    LineText As String
End Type
Private udtPL() As LineText
Private udtPR() As LineText
Private ABExists As Boolean

Private Sub FillComments(ByVal SampleIDWithOffset As Double)

          Dim s As String
          Dim OBS As Observations
          Dim OB As Observation

10        On Error GoTo FillComments_Error

20        Set OBS = New Observations
30        Set OBS = OBS.Load(SampleIDWithOffset, _
                             "Demographic", "MicroCS", "MicroIdent", _
                             "MicroGeneral", "MicroConsultant")

40        If Not OBS Is Nothing Then
50            For Each OB In OBS
60                Select Case UCase$(OB.Discipline)
                  Case "DEMOGRAPHIC": s = s & OB.Comment & vbCrLf
70                Case "MICROCS": s = s & OB.Comment & vbCrLf
80                Case "MICROIDENT": s = s & OB.Comment & vbCrLf
90                Case "MICROGENERAL": s = s & OB.Comment & vbCrLf
100               Case "MICROCOMSULTANT": s = s & OB.Comment & vbCrLf
110               End Select

120               txtReport.SelColor = vbBlack
130               txtReport.SelUnderline = False
140               txtReport.SelText = vbCrLf & s

150           Next
160       End If

170       Exit Sub

FillComments_Error:

          Dim strES As String
          Dim intEL As Integer

180       intEL = Erl
190       strES = Err.Description
200       LogError "frmMicroReport", "FillComments", intEL, strES

End Sub

Private Sub FillFluidReport(ByVal Site As String)

          Dim sql As String
          Dim tb As Recordset
          Dim SampleIDWithOffset As Double

10        On Error GoTo FillFluidReport_Error

20        SampleIDWithOffset = grdSID.TextMatrix(grdSID.row, 0) + SysOptMicroOffset(0)

30        sql = "SELECT * FROM GenericResults " & _
                "WHERE SampleID = " & SampleIDWithOffset & " " & _
                "AND (TestName LIKE  'CSF%' " & _
                "     OR TestName LIKE  'Fluid%' )"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        If tb.EOF Then Exit Sub

70        With txtReport
80            If ValidStatus4MicroDept(SampleIDWithOffset, "C") = False Then
90                .SelBold = True
100               .SelColor = vbRed
110               .SelText = "Not Validated" & vbCrLf
120           End If
130           .SelColor = vbBlack
140           .SelBold = False

150           .SelUnderline = True
160           .SelBold = True
170           .SelText = Site & " Report:" & vbCrLf
180           .SelBold = False
190           .SelUnderline = False

200           ShowReportFor "FluidAppearance0", "Cell Count"
210           ShowReportFor "FluidAppearance1", "Appearance"
220           ShowReportFor "FluidGram", "Gram"
230           ShowReportFor "FluidGram(2)", "Gram(2)"
240           ShowReportFor "FluidLeishmans", "Leishmans"
250           ShowReportFor "FluidZN", "ZN"
260           ShowReportFor "FluidWetPrep", "Wet Prep"
270           ShowReportFor "FluidCrystals", "Crystals"

280           .SelText = vbCrLf

290           ShowReportFor "FluidGlucose", "Glucose", "mmol/L"
300           ShowReportFor "FluidProtein", "Protein", "g/L"
310           ShowReportFor "FluidAlbumin", "Albumin", "g/L"
320           ShowReportFor "FluidGlobulin", "Globulin", "g/L"
330           ShowReportFor "FluidLDH", "LDH", "IU/L"
340           ShowReportFor "FluidAmylase", "Amylase", "IU/L"
350           ShowReportForCSF "CSFGlucose", "mmol/L"
360           ShowReportForCSF "CSFProtein", "g/L"

370           .SelText = vbCrLf

380           sql = "SELECT * FROM GenericResults " & _
                    "WHERE SampleID = " & SampleIDWithOffset & " " & _
                    "AND TestName LIKE  'CSFH%'"
390           Set tb = New Recordset
400           RecOpenServer 0, tb, sql
410           If tb.EOF Then Exit Sub

420           .SelBold = False
430           .SelText = "            "
440           .SelUnderline = True
450           .SelText = "Specimen         1         2         3             " & _
                         vbCrLf
460           .SelUnderline = False
470           .SelText = "                 RCC   "
480           ShowReportForHaem 0
490           .SelText = "/cmm" & vbCrLf

500           .SelText = "                 WCC   "
510           ShowReportForHaem 3
520           .SelText = "/cmm" & vbCrLf

530           .SelText = "         Polymorphic   "
540           ShowReportForHaem 6
550           .SelText = "%" & vbCrLf

560           .SelText = "       Mononucleated   "
570           ShowReportForHaem 9
580           .SelText = "%" & vbCrLf

590       End With

600       Exit Sub

FillFluidReport_Error:

          Dim strES As String
          Dim intEL As Integer

610       intEL = Erl
620       strES = Err.Description
630       LogError "frmMicroReport", "FillFluidReport", intEL, strES, sql

End Sub
Private Sub FillRSVReport(ByVal SampleIDWithOffset As Double)

          Dim sql As String
          Dim tb As Recordset

10        On Error GoTo FillRSVReport_Error

20        sql = "SELECT * FROM GenericResults " & _
                "WHERE SampleID = " & SampleIDWithOffset & " " & _
                "AND TestName = 'RSV'"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        If tb.EOF Then Exit Sub

60        With txtReport
70            If ValidStatus4MicroDept(SampleIDWithOffset, "V") = False Then
80                .SelBold = True
90                .SelColor = vbRed
100               .SelText = "Not Validated" & vbCrLf
110           End If
120           .SelColor = vbBlack
130           .SelBold = False
140           .SelText = "RSV Report: " & tb!Result
150           .SelBold = False
160           .SelUnderline = False
170           .SelText = vbCrLf
180           .SelText = vbCrLf
190           .SelColor = vbBlack

200       End With

210       Exit Sub

FillRSVReport_Error:

          Dim strES As String
          Dim intEL As Integer

220       intEL = Erl
230       strES = Err.Description
240       LogError "frmMicroReport", "FillRSVReport", intEL, strES, sql

End Sub


Private Sub FillFungalKOHReport(ByVal SampleIDWithOffset As Double)

          Dim sql As String
          Dim tb As Recordset

10        On Error GoTo FillFungalKOHReport_Error

20        sql = "SELECT * FROM GenericResults " & _
                "WHERE SampleID = " & SampleIDWithOffset & " " & _
                "AND TestName = 'PneumococcalAT'"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        If Not tb.EOF Then
60            With txtReport
70                If ValidStatus4MicroDept(SampleIDWithOffset, "C") = False Then
80                    .SelBold = True
90                    .SelColor = vbRed
100                   .SelText = "Pneumococcal A/T Not Validated" & vbCrLf
110                   .SelColor = vbBlack
120                   .SelBold = False
130               End If
140               .SelColor = vbBlack
150               .SelBold = False
160               .SelText = "Pneumococcal A/T Report: " & tb!Result & "" & vbCrLf
170               .SelBold = False
180               .SelUnderline = False
190               .SelColor = vbBlack
200               .SelText = vbCrLf
210               .SelText = vbCrLf
220           End With
230       End If

240       sql = "SELECT * FROM GenericResults " & _
                "WHERE SampleID = " & SampleIDWithOffset & " " & _
                "AND TestName = 'LegionellaAT'"
250       Set tb = New Recordset
260       RecOpenServer 0, tb, sql
270       If Not tb.EOF Then
280           With txtReport
290               If ValidStatus4MicroDept(SampleIDWithOffset, "C") = False Then
300                   .SelBold = True
310                   .SelColor = vbRed
320                   .SelText = "Legionella A/T Not Validated" & vbCrLf
330                   .SelColor = vbBlack
340                   .SelBold = False
350               End If
360               .SelColor = vbBlack
370               .SelBold = False
380               .SelText = "Legionella A/T Report: " & tb!Result & "" & vbCrLf
390               .SelBold = False
400               .SelUnderline = False
410               .SelColor = vbBlack
420               .SelText = vbCrLf
430               .SelText = vbCrLf
440           End With
450       End If

460       sql = "SELECT * FROM GenericResults " & _
                "WHERE SampleID = " & SampleIDWithOffset & " " & _
                "AND TestName = 'FungalElements'"
470       Set tb = New Recordset
480       RecOpenServer 0, tb, sql
490       If Not tb.EOF Then
500           With txtReport
510               If ValidStatus4MicroDept(SampleIDWithOffset, "C") = False Then
520                   .SelBold = True
530                   .SelColor = vbRed
540                   .SelText = "Fungal Elements Not Validated" & vbCrLf
550                   .SelColor = vbBlack
560                   .SelBold = False
570               End If
580               .SelColor = vbBlack
590               .SelBold = False
600               .SelText = "Fungal Elements Report: " & tb!Result & "" & vbCrLf
610               .SelBold = False
620               .SelUnderline = False
630               .SelColor = vbBlack
640               .SelText = vbCrLf
650               .SelText = vbCrLf
660           End With
670       End If

680       Exit Sub

FillFungalKOHReport_Error:

          Dim strES As String
          Dim intEL As Integer

690       intEL = Erl
700       strES = Err.Description
710       LogError "frmMicroReport", "FillFungalKOHReport", intEL, strES, sql


End Sub



Private Sub FillGrid()

          Dim sql As String
          Dim tb As Recordset
          Dim s As String


10        On Error GoTo FillGrid_Error



20        With grdSID
30            .Rows = 2
40            .AddItem ""
50            .RemoveItem 1
60        End With

70        sql = "SELECT * from Demographics where "
80        sql = sql & "PatName = '" & AddTicks(pPatName) & "' and "
90        If Trim$(pPatChart) <> "" Then
100           sql = sql & "Chart = '" & pPatChart & "' and "
110       Else
120           sql = sql & "(Chart = '' OR Chart IS NULL) and "
130       End If
140       If Trim$(pPatSex) <> "" Then
150           sql = sql & "Sex = '" & pPatSex & "' and "
160       Else
170           sql = sql & "(Sex = '' OR Sex IS NULL) and "
180       End If
190       If IsDate(pPatDoB) Then
200           sql = sql & "DoB = '" & Format(pPatDoB, "yyyy/mm/dd") & "' "
210       Else
220           sql = sql & "(COALESCE(DoB, '')  = '') "
230       End If
240       sql = sql & "ORDER BY RunDate desc"

250       Set tb = New Recordset
260       RecOpenClient 0, tb, sql


270       Do While Not tb.EOF


280           If Trim(tb!Hyear & "") = "" Then
290               If Val(tb!SampleID) > SysOptMicroOffset(0) Then

300                   s = Format$(Val(tb!SampleID) - SysOptMicroOffset( _
                                  0)) & vbTab & tb!Rundate & vbTab
310                   If IsDate(tb!SampleDate) Then
320                       If Format(tb!SampleDate, "hh:mm") <> "00:00" Then
330                           s = s & Format(tb!SampleDate, "dd/MM/yyyy hh:mm")
340                       Else
350                           s = s & Format(tb!SampleDate, "dd/MM/yyyy")
360                       End If
370                   Else
380                       s = s & "Not Specified"
390                   End If

400                   s = s & vbTab & LoadOutstandingMicro(tb!SampleID) & vbTab

410                   If IsDate(tb!SampleDate) Then
420                       s = s & CalcAge(pPatDoB, tb!SampleDate)
430                   End If
440                   s = s & vbTab
450                   s = s & tb!Addr0 & " " & tb!Addr1 & ""
460                   grdSID.AddItem s


470               End If
480           End If
490           tb.MoveNext
500       Loop

510       If grdSID.Rows > 2 Then
520           grdSID.RemoveItem 1
530       End If

540       Exit Sub

FillGrid_Error:

          Dim strES As String
          Dim intEL As Integer



550       intEL = Erl
560       strES = Err.Description
570       LogError "frmMicroReport", "FillGrid", intEL, strES, sql


End Sub
Private Function LoadOutstandingMicro(ByVal SampleIDWithOffset As Double) As String

          Dim tb As Recordset
          Dim sql As String
          Dim s As String

10        On Error GoTo LoadOutstandingMicro_Error

20        sql = "SELECT * from MicroSiteDetails " & _
                "WHERE SampleID = " & SampleIDWithOffset
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql

50        If Not tb.EOF Then
60            s = tb!Site & " " & tb!SiteDetails & " "
70            If tb!Site & "" = "Urine" Then
80                sql = "SELECT * FROM UrineRequests WHERE SampleID = " & SampleIDWithOffset
90                Set tb = New Recordset
100               RecOpenServer 0, tb, sql
110               If Not tb.EOF Then
120                   If tb!cS Then s = s & "C & S "
130                   If tb!Pregnancy Then s = s & "Pregnancy "
140                   If tb!RedSub Then s = s & "Red Sub"
150               End If
160           ElseIf tb!Site & "" = "Faeces" Then
170               sql = "SELECT * FROM FaecalRequests WHERE SampleID = " & SampleIDWithOffset
180               Set tb = New Recordset
190               RecOpenServer 0, tb, sql

200               If Not tb.EOF Then
210                   If tb!cS Then s = s & "C & S "
220                   If tb!ToxinAB Then s = s & "C. Difficile "
230                   If tb!OP Then s = s & "O/P "
240                   If tb!OB0 Or tb!OB1 Or tb!OB2 Then s = s & "Occult Blood "
250                   If tb!Rota Then s = s & "Rota "
260                   If tb!Adeno Then s = s & "Adeno "
270                   If tb!HPylori Then s = s & "H.Pylori "
280                   If tb!Coli0157 Then s = s & "Coli 0157 "
290                   If tb!ssScreen Then s = s & "S/S Screen "
300                   If tb!GDH Then s = s & "GDH "
310                   If tb!PCR Then s = s & "PCR "
320               End If
330           End If
340       End If
350       LoadOutstandingMicro = s

360       Exit Function

LoadOutstandingMicro_Error:

          Dim strES As String
          Dim intEL As Integer

370       intEL = Erl
380       strES = Err.Description
390       LogError "frmMicroReport", "LoadOutstandingMicro", intEL, strES, sql

End Function

Private Sub ShowReportFor(ByVal Parameter As String, _
                          ByVal DisplayName As String, Optional ByVal Units As String)

          Dim sql As String
          Dim tb As Recordset

10        On Error GoTo ShowReportFor_Error

20        sql = "SELECT * FROM GenericResults WHERE " & "SampleID = '" & _
                grdSID.TextMatrix(grdSID.row, _
                                  0) + SysOptMicroOffset(0) & "' " & "AND TestName = '" & Parameter & "'"

30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        If tb.EOF Then Exit Sub

60        With txtReport
70            .SelColor = vbBlack
80            .SelText = "              "
90            .SelBold = False
100           .SelText = Left$(DisplayName & Space$(15), 15)
110           .SelBold = True
120           .SelText = tb!Result & " " & Units & vbCrLf
130           .SelColor = vbBlack
140           .SelBold = False
150       End With

160       Exit Sub

ShowReportFor_Error:

          Dim strES As String
          Dim intEL As Integer

170       intEL = Erl
180       strES = Err.Description
190       LogError "frmMicroReport", "ShowReportFor", intEL, strES, sql

End Sub

Private Sub ShowReportForCSF(ByVal Parameter As String, ByVal Units As String)

          Dim sql As String
          Dim tb As Recordset

10        On Error GoTo ShowReportForCSF_Error

20        sql = "SELECT * FROM GenericResults WHERE " & "SampleID = '" & _
                grdSID.TextMatrix(grdSID.row, _
                                  0) + SysOptMicroOffset(0) & "' " & "AND TestName = '" & Parameter & "'"

30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        If tb.EOF Then Exit Sub

60        With txtReport
70            .SelText = "              "
80            .SelBold = False
90            .SelText = Left$(Mid$(Parameter, 4) & Space$(15), 15)
100           .SelBold = True
110           .SelText = tb!Result & " " & Units & vbCrLf
120           .SelBold = False
130       End With

140       Exit Sub

ShowReportForCSF_Error:

          Dim strES As String
          Dim intEL As Integer

150       intEL = Erl
160       strES = Err.Description
170       LogError "frmMicroReport", "ShowReportForCSF", intEL, strES, sql

End Sub

Private Sub ShowReportForHaem(ByVal pNumber As Integer)

          Dim sql As String
          Dim tb As Recordset
          Dim s As String

10        On Error GoTo ShowReportForHaem_Error

20        s = "      "
30        sql = "SELECT Result FROM GenericResults WHERE " & "SampleID = '" & _
                grdSID.TextMatrix(grdSID.row, _
                                  0) + SysOptMicroOffset(0) & "' " & "AND TestName = 'CSFHAEM" & Format$( _
                                  pNumber) & "'"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        If tb.EOF Then
70            s = s & "          "
80        Else
90            s = s & Left$(tb!Result & Space$(10), 10)
100       End If

110       sql = "SELECT Result FROM GenericResults WHERE " & "SampleID = '" & _
                grdSID.TextMatrix(grdSID.row, _
                                  0) + SysOptMicroOffset(0) & "' " & "AND TestName = 'CSFHAEM" & Format$( _
                                  pNumber + 1) & "'"
120       Set tb = New Recordset
130       RecOpenServer 0, tb, sql
140       If tb.EOF Then
150           s = s & "          "
160       Else
170           s = s & Left$(tb!Result & Space$(10), 10)
180       End If

190       sql = "SELECT Result FROM GenericResults WHERE " & "SampleID = '" & _
                grdSID.TextMatrix(grdSID.row, _
                                  0) + SysOptMicroOffset(0) & "' " & "AND TestName = 'CSFHAEM" & Format$( _
                                  pNumber + 2) & "'"
200       Set tb = New Recordset
210       RecOpenServer 0, tb, sql
220       If tb.EOF Then
230           s = s & "          "
240       Else
250           s = s & Left$(tb!Result & Space$(10), 10)
260       End If

270       txtReport.SelBold = True
280       txtReport.SelText = s
290       txtReport.SelBold = False

300       Exit Sub

ShowReportForHaem_Error:

          Dim strES As String
          Dim intEL As Integer

310       intEL = Erl
320       strES = Err.Description
330       LogError "frmMicroReport", "ShowReportForHaem", intEL, strES, sql

End Sub


Private Sub cmdCancel_Click()

10        Unload Me

End Sub

Private Sub cmdPrint_Click()
10        PrintThis
End Sub

Private Sub Form_Activate()

10        On Error GoTo Form_Activate_Error


20        PBar.Max = LogOffDelaySecs
30        PBar = 0

40        Timer1.Enabled = True

50        If Not Activated Then
60            Activated = True
70            FillGrid
80        End If

90        Exit Sub

Form_Activate_Error:

          Dim strES As String
          Dim intEL As Integer



100       intEL = Erl
110       strES = Err.Description
120       LogError "frmMicroReport", "Form_Activate", intEL, strES


End Sub

Private Sub Form_Deactivate()

10        On Error GoTo Form_Deactivate_Error

20        Timer1.Enabled = False

30        Exit Sub

Form_Deactivate_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmMicroReport", "Form_Deactivate", intEL, strES


End Sub


Private Sub Form_Load()

10        On Error GoTo Form_Load_Error

20        Activated = False

30        PBar.Max = LogOffDelaySecs
40        PBar = 0

50        Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer



60        intEL = Erl
70        strES = Err.Description
80        LogError "frmMicroReport", "Form_Load", intEL, strES


End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, _
                           Y As Single)

10        On Error GoTo Form_MouseMove_Error

20        PBar = 0

30        Exit Sub

Form_MouseMove_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmMicroReport", "Form_MouseMove", intEL, strES


End Sub


Private Sub Form_Unload(Cancel As Integer)

10        On Error GoTo Form_Unload_Error

20        Activated = False

30        Exit Sub

Form_Unload_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmMicroReport", "Form_Unload", intEL, strES


End Sub



Private Sub grdSID_Click()

          Dim X As Long
          Dim Y As Long

10        On Error GoTo grdSID_Click_Error

20        txtReport = ""

30        If grdSID.MouseRow = 0 Then
40            If grdSID.MouseCol = 1 Or grdSID.MouseCol = 2 Then
50                grdSID.Col = grdSID.MouseCol
60                grdSID.Sort = 9
70            Else
80                If SortOrder Then
90                    grdSID.Sort = flexSortGenericAscending
100               Else
110                   grdSID.Sort = flexSortGenericDescending
120               End If
130           End If
140           SortOrder = Not SortOrder
150           Exit Sub
160       End If

170       For Y = 1 To grdSID.Rows - 1
180           grdSID.row = Y
190           For X = 1 To grdSID.Cols - 1
200               grdSID.Col = X
210               grdSID.CellBackColor = 0
220           Next
230       Next

240       grdSID.row = grdSID.MouseRow
250       For X = 1 To grdSID.Cols - 1
260           grdSID.Col = X
270           grdSID.CellBackColor = vbYellow
280       Next

290       FillResultMicro Val(grdSID.TextMatrix(grdSID.row, _
                                                0)) + SysOptMicroOffset(0)

300       Exit Sub

grdSID_Click_Error:

          Dim strES As String
          Dim intEL As Integer

310       intEL = Erl
320       strES = Err.Description
330       LogError "frmMicroReport", "grdSID_Click", intEL, strES

End Sub

Private Sub grdSID_Compare(ByVal Row1 As Long, ByVal Row2 As Long, _
                           Cmp As Integer)

          Dim d1 As String
          Dim d2 As String

10        If Not IsDate(grdSID.TextMatrix(Row1, grdSID.Col)) Then
20            Cmp = 0
30            Exit Sub
40        End If

50        If Not IsDate(grdSID.TextMatrix(Row2, grdSID.Col)) Then
60            Cmp = 0
70            Exit Sub
80        End If

90        d1 = Format(grdSID.TextMatrix(Row1, grdSID.Col), "dd/mmm/yyyy hh:mm:ss")
100       d2 = Format(grdSID.TextMatrix(Row2, grdSID.Col), "dd/mmm/yyyy hh:mm:ss")

110       If SortOrder Then
120           Cmp = Sgn(DateDiff("s", d1, d2))
130       Else
140           Cmp = Sgn(DateDiff("s", d2, d1))
150       End If

End Sub


Private Sub FillResultMicro(ByVal SampleIDWithOffset As Double)

          Dim sql As String
          Dim tb As Recordset
          Dim s As String
          Dim pDefault As Long

10        On Error GoTo FillResultMicro_Error

20        lblClDetails = ""

30        sql = "SELECT D.Valid, D.ClDetails, M.Site, M.PCA0, M.PCA1, M.PCA2, M.PCA3, L.[Default] " & _
                "FROM Demographics D, MicroSiteDetails M, Lists L " & _
                "WHERE D.SampleID = " & SampleIDWithOffset & " " & _
                "AND D.SampleID = M.SampleID " & _
                "AND L.ListType = 'SI' " & _
                "AND L.[Text] LIKE M.Site "
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        pDefault = 3

70        If tb.EOF Then Exit Sub
80        lblClDetails = tb!ClDetails & ""
90        s = Trim$(tb!PCA0 & " " & tb!PCA1 & " " & tb!PCA2 & " " & tb!PCA3 & "")
100       pDefault = Val(tb!Default)

110       If IsNull(tb!Valid) Or tb!Valid = 0 Then
120           txtReport.SelUnderline = False
130           txtReport.SelBold = True
140           txtReport.SelColor = vbRed
150           txtReport.SelText = Space$(20)
160           txtReport.SelUnderline = True
170           txtReport.SelText = "DEMOGRAPHICS NOT VALIDATED" & vbCrLf
180           txtReport.SelBold = False
190           txtReport.SelColor = vbBlack
200       End If

210       txtReport.SelUnderline = False

220       FillAssIDs

230       If Trim$(s) <> "" Then
240           txtReport.SelText = "Current Antibiotics:" & s & vbCrLf
250       End If

260       txtReport.SelUnderline = False
270       txtReport.SelBold = False

280       If Trim$(lblClDetails) <> "" Then
290           txtReport.SelText = "Clinical Details:"
300           txtReport.SelBold = True
310           txtReport.SelText = lblClDetails & vbCrLf
320       End If

330       txtReport.SelUnderline = False
340       txtReport.SelBold = False

350       FillRSVReport SampleIDWithOffset

360       If InStr(grdSID.TextMatrix(grdSID.row, 3), "Faeces") > 0 Then
370           FillFaeces SampleIDWithOffset
380       Else
390           FillMicroscopy SampleIDWithOffset

400       End If

          'FillComments SampleIDWithOffset


410       FillFungalKOHReport SampleIDWithOffset

          'If IsFluid(tb!Site & "") Then
420           FillFluidReport tb!Site & ""
          'End If

430       ReDim udtPL(0 To 0)
440       ReDim udtPR(0 To 0)
450       FillG SampleIDWithOffset

          'GetPrintLineInterimHeading
460       If UCase$(tb!Site) = "BLOOD CULTURE" Then
470           PrepareSensitivitiesBloodCulture SampleIDWithOffset, pDefault
480           PrintResultLines
490           PrintTextRTB txtReport, "COMMENTS" & vbCrLf, 10, True
500           GetPrintLineComments SampleIDWithOffset, "COMMENTS", "Demographic"
510           GetPrintLineComments SampleIDWithOffset, "COMMENTS", "MicroCS"
520           GetPrintLineComments SampleIDWithOffset, "COMMENTS", "MicroConsultant"
530       Else
540           PrepareSensitivitiesOther
550           PrintResultLines
560           PrintTextRTB txtReport, "COMMENTS" & vbCrLf, 10, True
570           GetPrintLineComments SampleIDWithOffset, "COMMENTS", "Demographic"
580           GetPrintLineComments SampleIDWithOffset, "COMMENTS", "CSFFluid"
590           GetPrintLineComments SampleIDWithOffset, "COMMENTS", "MicroGeneral"
600           GetPrintLineComments SampleIDWithOffset, "COMMENTS", "Semen"
610           GetPrintLineComments SampleIDWithOffset, "COMMENTS", "MicroCS"          'Medical scientist comments
620           GetPrintLineComments SampleIDWithOffset, "COMMENTS", "MicroConsultant"
630       End If



640       Exit Sub

FillResultMicro_Error:

          Dim strES As String
          Dim intEL As Integer

650       intEL = Erl
660       strES = Err.Description
670       LogError "frmMicroReport", "FillResultMicro", intEL, strES, sql

End Sub

Private Sub FillAssIDs()

          Dim sql As String
          Dim tb As Recordset
          Dim SampleID As String
          Dim Added As Boolean
          Dim SIDs As New Collection
          Dim SID As Variant
          Dim Found As Boolean

10        On Error GoTo FillAssIDs_Error

20        SampleID = Val(grdSID.TextMatrix(grdSID.row, 0)) + SysOptMicroOffset(0)

30        Added = True
40        Do While Added
50            Added = False
60            sql = "SELECT AssID FROM AssociatedIDs WHERE " & "SampleID = '" & _
                    SampleID & "' "
70            For Each SID In SIDs
80                sql = sql & "OR SampleID = '" & SID & "' "
90            Next
100           Set tb = New Recordset
110           RecOpenServer 0, tb, sql
120           Do While Not tb.EOF
130               Found = False
140               For Each SID In SIDs
150                   If Format$(SID) = Format$(tb!AssID) Then
160                       Found = True
170                       Exit For
180                   End If
190               Next
200               If Not Found Then
210                   SIDs.Add tb!AssID & ""
220                   Added = True
230               End If
240               tb.MoveNext
250           Loop

260           sql = "SELECT SampleID FROM AssociatedIDs WHERE " & "AssID = '" & _
                    SampleID & "' "
270           For Each SID In SIDs
280               sql = sql & "OR AssID = '" & SID & "' "
290           Next
300           Set tb = New Recordset
310           RecOpenServer 0, tb, sql
320           Do While Not tb.EOF
330               Found = False
340               For Each SID In SIDs
350                   If Format$(SID) = Format$(tb!SampleID) Then
360                       Found = True
370                       Exit For
380                   End If
390               Next
400               If Not Found Then
410                   SIDs.Add tb!SampleID & ""
420                   Added = True
430               End If
440               tb.MoveNext
450           Loop
460       Loop

470       If SIDs.Count > 0 Then
480           txtReport.SelText = "Associated Samples: "
490           For Each SID In SIDs
500               txtReport.SelText = Format$(Val(SID) - SysOptMicroOffset(0)) & " "
510           Next
520           txtReport.SelText = vbCrLf
530       End If

540       Exit Sub

FillAssIDs_Error:

          Dim strES As String
          Dim intEL As Integer

550       intEL = Erl
560       strES = Err.Description
570       LogError "frmMicroReport", "FillAssIDs", intEL, strES, sql

End Sub


Private Function FillIQ200(ByVal SampleIDWithOffset As Double) As Boolean

          Dim tb As Recordset
          Dim sql As String
          Dim ShortName As String
          Dim TestNameLength As Integer
          Dim ResultLength As Integer

          Dim WBCFound As Boolean
          Dim RBCFound As Boolean
          Dim CastFound As Boolean
          Dim CrystalFound As Boolean
          Dim EpithelialFound As Boolean
          Dim PrintThis As Boolean
          Dim tempResult As String
          Dim X As Integer

10        On Error GoTo FillIQ200_Error

20        FillIQ200 = False

30        TestNameLength = 26
40        ResultLength = 16

50        sql = "SELECT ShortName,LongName,Result FROM IQ200 WHERE " & _
                "SampleID = " & SampleIDWithOffset & " " & _
                "AND Result <> '[none]'"

60        PrintThis = True

70        Set tb = New Recordset
80        RecOpenServer 0, tb, sql
90        If Not tb.EOF Then
100           PrintTextRTB txtReport, "MICROSCOPY" & vbCrLf, , True
110           Do While Not tb.EOF


120               X = InStr(tb!Result & "", " ") - 1
130               If X > 0 Then
140                   tempResult = Left(tb!Result & "", X)
150               Else
160                   tempResult = tb!Result & ""
170               End If

180               If tb!ShortName & "" = "WBC" Then
190                   WBCFound = True
200                   If IsNumeric(tempResult) Then
210                       If Val(tempResult) > 25 Then PrintThis = False
220                   End If
230               ElseIf tb!ShortName & "" = "RBC" Then
240                   RBCFound = True
250               ElseIf InStr(tb!LongName & "", "Cast") > 0 Then
260                   CastFound = True
270               ElseIf InStr(tb!LongName & "", "Crystal") > 0 Then
280                   CrystalFound = True
290               ElseIf InStr(tb!LongName & "", "Epithelial") > 0 Then
300                   EpithelialFound = True
310               ElseIf tb!ShortName & "" = "BACT" Then
320                   If IsNumeric(tempResult) Then
330                       If Val(tempResult) > 2 Then PrintThis = False
340                   End If
350               ElseIf tb!ShortName & "" = "PC" Then
360                   If IsNumeric(tempResult) Then
370                       If Val(tempResult) > 8500 Then PrintThis = False
380                   End If
390               End If


400               If (InStr(tb!LongName & "", "Cast") > 0) Or _
                     (InStr(tb!LongName & "", "Crystal") > 0) Or _
                     (InStr(tb!LongName & "", "Epithelial") > 0) Then
410                   If tb!Result > 0 Then
420                       PrintTextRTB txtReport, FormatString(" ", 4, , AlignLeft) & _
                                                  FormatString(tb!LongName & "", TestNameLength, , AlignLeft) & _
                                                  FormatString(tb!Result & "", ResultLength, , AlignLeft) & vbCrLf, 10
430                   End If
440               Else
450                   If tb!ShortName & "" <> "BACT" And tb!ShortName & "" <> "PC" Then
460                       PrintTextRTB txtReport, FormatString(" ", 4, , AlignLeft) & _
                                                  FormatString(tb!LongName & "", TestNameLength, , AlignLeft) & _
                                                  FormatString(tb!Result & "", ResultLength, , AlignLeft) & vbCrLf, 10
470                   End If
480               End If
490               tb.MoveNext
500           Loop


510           If Not WBCFound Then
520               PrintTextRTB txtReport, FormatString(" ", 4, , AlignLeft) & _
                                          FormatString("WBC", TestNameLength, , AlignLeft) & _
                                          FormatString("0 /uL", ResultLength, , AlignLeft) & vbCrLf, 10


530           End If

540           If Not RBCFound Then
550               PrintTextRTB txtReport, FormatString(" ", 4, , AlignLeft) & _
                                          FormatString("RBC", TestNameLength, , AlignLeft) & _
                                          FormatString("0 /uL", ResultLength, , AlignLeft) & vbCrLf, 10

560           End If

570           If Not EpithelialFound Then
580               PrintTextRTB txtReport, FormatString(" ", 4, , AlignLeft) & _
                                          FormatString("Epithelial Cells", TestNameLength, , AlignLeft) & _
                                          FormatString("0 /uL", ResultLength, , AlignLeft) & vbCrLf, 10


590           End If
600           FillIQ200 = True

610       End If




620       Exit Function

FillIQ200_Error:

          Dim strES As String
          Dim intEL As Integer

630       intEL = Erl
640       strES = Err.Description
650       LogError "modNewMicro", "FillIQ200", intEL, strES, sql


End Function


Private Sub FillMicroscopy(ByVal SampleIDWithOffset As Double)

          Dim tb As Recordset
          Dim sql As String
          Dim ColonyCount As String

10        On Error GoTo FillMicroscopy_Error

20        ReDim Comments(1 To 4) As String
          Dim n As Integer

          '******If IQ200 results are avaiable then do not print manual microscopy
30        If FillIQ200(SampleIDWithOffset) = True Then Exit Sub


40        sql = "SELECT * from Urine WHERE SampleID = " & SampleIDWithOffset
50        Set tb = New Recordset
60        RecOpenServer 0, tb, sql
70        If tb.EOF Then Exit Sub

80        If Trim$( _
             tb!Bacteria & tb!WCC & tb!RCC & tb!Crystals & tb!Casts & tb!Misc0 & tb!Misc1 & _
             tb!Misc2 & tb!Pregnancy & "") <> "" Then

90            With txtReport
100               If ValidStatus4MicroDept(SampleIDWithOffset, "U") = False Then
110                   .SelBold = True
120                   .SelColor = vbRed
130                   .SelText = "Not Validated" & vbCrLf
140               End If

150               .SelColor = vbBlack
160               .SelBold = False

                  '160       .SelText = "Bacteria:" & Trim$(tb!Bacteria & "") & vbCrLf

170               .SelBold = True
180               .SelText = "MICROSCOPY"
190               .SelUnderline = False
200               .SelBold = False
210               .SelText = Left(" " & Space(28), 28)
220               If SysOptDipStick(0) Then
230                   .SelUnderline = True
240                   .SelText = "Biochemistry" & vbCrLf
250                   .SelUnderline = False
260               Else
270                   .SelText = vbCrLf
280               End If

290               .SelBold = False

300               .SelText = Left("     WCC: " & LeftOfBar(tb!WCC & "") & Space(35), 35)
310               If SysOptDipStick(0) Then
320                   .SelText = "          Ph:" & Trim$(tb!pH & " ") & vbCrLf
330               Else
340                   .SelText = vbCrLf
350               End If

360               .SelText = Left("     RCC: " & LeftOfBar(tb!RCC & "") & Space(35), 35)
370               If SysOptDipStick(0) Then
380                   .SelText = "     Protein: " & Trim$(tb!Protein & "") & vbCrLf
390               Else
400                   .SelText = vbCrLf
410               End If

420               .SelText = Left("   Casts: " & Trim$(tb!Casts & "") & Space(35), 35)
430               If SysOptDipStick(0) Then
440                   .SelText = "     Glucose: " & Trim$(tb!Glucose & "") & vbCrLf
450               Else
460                   .SelText = vbCrLf
470               End If

480               .SelText = Left("Crystals: " & Trim$(tb!Crystals & "") & Space(35), _
                                  35)
490               If SysOptDipStick(0) Then
500                   .SelText = "     Ketones: " & Trim$(tb!Ketones & " ") & vbCrLf
510               Else
520                   .SelText = vbCrLf
530               End If

540               If SysOptDipStick(0) Then
550                   .SelText = Left(" " & Space(35), _
                                      35) & "Urobilinogen: " & Trim$(tb!Urobilinogen & " ") & vbCrLf
560                   .SelText = Left(" " & Space(35), _
                                      35) & "   Bilirubin: " & Trim$(tb!Bilirubin & " ") & vbCrLf
570                   .SelText = Left(" " & Space(35), _
                                      35) & "    Blood Hb: " & Trim$(tb!BloodHb & " ") & vbCrLf
580               End If

590               If Trim(tb!Misc0 & "") <> "" Or Trim(tb!Misc1 & "") <> "" Or Trim( _
                     tb!Misc2 & "") <> "" Then
600                   .SelText = "    Misc: " & Trim$(tb!Misc0 & " ") & vbCrLf
610                   If Trim(tb!Misc1 & "") <> "" Then
620                       .SelText = "          " & Trim(tb!Misc1 & " ") & vbCrLf
630                   End If
640                   If Trim(tb!Misc2 & "") <> "" Then
650                       .SelText = "          " & Trim(tb!Misc2 & "") & vbCrLf
660                   End If
670               End If
680           End With
690       End If

700       With txtReport
710           ColonyCount = Trim$(getColonyCount(SampleIDWithOffset))
              '  If ColonyCount <> "" Then
              '    .SelUnderline = True
              '    .SelText = Left(" " & Space(60), 60) & vbCrLf
              '    .SelText = "Colony Count :"
              '    .SelUnderline = False
              '    .SelBold = False
              '    .SelText = " " & getColonyCount(SampleIDWithOffset) & vbCrLf
              '    .SelUnderline = True
              '    .SelText = Left(" " & Space(60), 60) & vbCrLf
              '    .SelUnderline = False
              '  End If

720           If Trim(tb!HCGLevel & "") <> "" Then
730               .SelText = "         "
740               .SelBold = True
750               .SelText = "hCG:"
760               .SelBold = False
770               .SelText = " "
780               .SelText = Left(Trim(tb!HCGLevel) & " mIU/mL" & Space(20), 20)
790           End If

800           If Trim(tb!BenceJones & "") <> "" Then
810               .SelText = " "
820               .SelBold = True
830               .SelText = "Bence Jones Protein:"
840               .SelBold = False
850               .SelText = " "
860               .SelText = Trim(tb!BenceJones) & vbCrLf
870           End If

880           If Trim(tb!FatGlobules & "") <> "" Then
890               .SelText = " "
900               .SelBold = True
910               .SelText = "Fat Globules:"
920               .SelBold = False
930               .SelText = " "
940               .SelText = Left(Trim(tb!FatGlobules) & Space(20), 20)
950           End If

960           If Trim(tb!SG & "") <> "" Then
970               .SelText = "                 "
980               .SelBold = True
990               .SelText = "SG:"
1000              .SelBold = False
1010              .SelText = " "
1020              .SelText = Trim(tb!SG)
1030          End If

1040          If Trim(tb!Pregnancy & "") <> "" Then
1050              .SelText = "      Pregnancy "
1060              .SelBold = True
1070              If tb!Pregnancy = "P" Then
1080                  .SelText = "Positive"
1090              ElseIf tb!Pregnancy = "N" Then
1100                  .SelText = "Negative"
1110              ElseIf tb!Pregnancy = "E" Then
1120                  .SelText = "Equivocal"
1130              ElseIf tb!Pregnancy = "U" Then
1140                  .SelText = "Specimen Unsuitable"
1150              Else
1160                  .SelText = tb!Pregnancy & ""
1170              End If
1180              .SelBold = False
1190          End If

1200          .SelText = vbCrLf

1210          .SelBold = False

1220      End With

1230      Exit Sub

FillMicroscopy_Error:

          Dim strES As String
          Dim intEL As Integer

1240      intEL = Erl
1250      strES = Err.Description
1260      LogError "frmMicroReport", "FillMicroscopy", intEL, strES, sql

End Sub
Private Function getColonyCount(SampleIDWithOffset As Double) As String

          Dim tb As Recordset
          Dim sql As String

          'Get Colony count for sampleid and Isolate 1 only
          'PC
10        On Error GoTo getColonyCount_Error

20        getColonyCount = ""

30        sql = "SELECT Qualifier from Isolates " & _
                "WHERE SampleID = " & SampleIDWithOffset & " AND IsolateNumber = 1"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        If Not tb.EOF Then
70            getColonyCount = tb!Qualifier & ""
80        End If

90        Exit Function

getColonyCount_Error:

          Dim strES As String
          Dim intEL As Integer

100       intEL = Erl
110       strES = Err.Description
120       LogError "frmMicroReport", "getColonyCount", intEL, strES, sql

End Function


Private Sub FillFaeces(ByVal SampleIDWithOffset As Double)

      Dim sql As String
      Dim tbRequests As Recordset
      Dim tbFaeces As Recordset
      Dim tb As Recordset
      Dim n As Integer
      Dim ColName As String
      Dim ShortName As String
      Dim Title As String
      Dim TestNameLength As Integer
      Dim ResultLength As Integer





10    On Error GoTo FillFaeces_Error

20    TestNameLength = 22
30    ResultLength = 24




40    sql = "SELECT " & "COALESCE(CS, 0) CS, " & _
            "COALESCE(Cryptosporidium, 0) Crypto, " & "COALESCE(Rota, 0) Rota, " & _
            "COALESCE(Adeno, 0) Adeno, " & "COALESCE(OB0, 0) OB0, " & _
            "COALESCE(OB1, 0) OB1, " & "COALESCE(OB2, 0) OB2, " & "COALESCE(OP, 0) OP, " & _
            "COALESCE(ToxinAB, 0) ToxinAB, " & "COALESCE(HPylori, 0) HPylori, " & _
            "COALESCE(RedSub, 0) RedSub, " & "COALESCE(CDiff, 0) CDiff, " & _
            "COALESCE(GDH, 0) GDH, " & "COALESCE(PCR, 0) PCR , COALESCE(GL, 0) GL " & _
            "FROM FaecalRequests " & _
            "WHERE SampleID = " & SampleIDWithOffset

50    Set tbRequests = New Recordset
60    RecOpenServer 0, tbRequests, sql

70    If Not tbRequests.EOF Then
80        With txtReport

90            sql = "SELECT " & _
                    "COALESCE(LTRIM(RTRIM(LEFT(F.OB0 + '|', CHARINDEX('|', F.OB0) - 1))), '') B0, " & _
                    "COALESCE(LTRIM(RTRIM(LEFT(F.OB1 + '|', CHARINDEX('|', F.OB1) - 1))), '') B1, " & _
                    "COALESCE(LTRIM(RTRIM(LEFT(F.OB2 + '|', CHARINDEX('|', F.OB2) - 1))), '') B2, " & _
                    "COALESCE(LTRIM(RTRIM(LEFT(F.Rota + '|', CHARINDEX('|', F.Rota) - 1))), '') R, " & _
                    "COALESCE(LTRIM(RTRIM(LEFT(F.Adeno + '|', CHARINDEX('|', F.Adeno) - 1))), '') A, " & _
                    "COALESCE(LTRIM(RTRIM(LEFT(ToxinAB + '|', CHARINDEX('|', ToxinAB) - 1))), '') ToxAB, " & _
                    "COALESCE(LTRIM(RTRIM(LEFT(CDiffCulture + '|', CHARINDEX('|', CDiffCulture) - 1))), '') ToxC, " & _
                    "COALESCE(GDHDetail, '') GDHDetail, " & _
                    "COALESCE(PCRDetail, '') PCRDetail, " & _
                    "COALESCE(LTRIM(RTRIM(LEFT(F.Cryptosporidium + '|', CHARINDEX('|', F.Cryptosporidium) - 1))), '') Crypto, " & _
                    "COALESCE(LTRIM(RTRIM(LEFT(F.GiardiaLambila + '|', CHARINDEX('|', F.GiardiaLambila) - 1))), '') GL, " & _
                    "OP0, OP1, OP2, " & _
                    "COALESCE(LTRIM(RTRIM(LEFT(F.HPylori + '|', CHARINDEX('|', F.HPylori) - 1))), '') H " & _
                    "FROM Faeces F WHERE " & _
                    "F.SampleID = '" & SampleIDWithOffset & "' "

100           Set tbFaeces = New Recordset
110           RecOpenServer 0, tbFaeces, sql

120           If Not tbFaeces.EOF Then

130               .SelColor = vbBlack
140               .SelUnderline = False
150               .SelBold = True

160               .SelText = "FAECES" & vbCrLf

170               .SelColor = vbBlack
180               .SelUnderline = False
190               .SelBold = False

                  '110               NumberOfTitles = NumberOfTitles + 1
200               For n = 1 To 15
210                   ColName = Choose(n, "B0", "B1", "B2", "R", "A", "ToxAB", "ToxC", "GDHDetail", "PCRDetail", "Crypto", "GL", "OP0", "OP1", "OP2", "H")
220                   ShortName = Choose(n, "Occult Blood (1):", "Occult Blood (2):", "Occult Blood (3):", _
                                         "Rota Virus:", "Adeno Virus:", _
                                         "C.difficile Toxin A/B:", "C.difficile Culture:", "GDH:", "PCR:", _
                                         "Cryptosporidium:", "Giardia Lambila", _
                                         "Ova and Parasites (1):", "                  (2):", "                  (3):", _
                                         "H.pylori Antigen Test:")
230                   If Trim$(tbFaeces(ColName) & "") <> "" Then
240                       If ColName = "H" Then
                              'AddNotAccreditedTest "HP", True

250                       End If
260                       .SelText = FormatString(ShortName, TestNameLength) & tbFaeces(ColName) & vbCrLf   '"Reducing Substances   : Report Not Ready" & vbCrLf


270                   End If
280               Next

290           End If


              '280           If tbRequests!RedSub Then
              '290               sql = "Select * from GenericResults where " & _
               '                        "SampleID = " & SampleIDWithOffset & " " & _
               '                        "and TestName = 'RedSub'"
              '300               Set tb = New Recordset
              '310               RecOpenServer 0, tb, sql
              '320               If tb.EOF Then
              '330                   .SelText = "Reducing Substances   : Report Not Ready" & vbCrLf
              '340               Else
              '350                   .SelText = "Reducing Substances   : " & tb!Result & vbCrLf
              '360               End If
              '370           End If

300           .SelColor = vbBlack
310           .SelUnderline = False
320           .SelBold = False

330       End With
340   End If



350   Exit Sub

FillFaeces_Error:

      Dim strES As String
      Dim intEL As Integer

360   intEL = Erl
370   strES = Err.Description
380   LogError "frmMicroReport", "FillFaeces", intEL, strES, sql

End Sub

Private Sub PrepareSensitivitiesBloodCulture(ByVal SampleIDWithOffset As Double, _
                                             ByVal MaxSensitives As String)

          Dim ABCount As Integer
          Dim lpc As Integer
          Dim ResultsPerPage As Integer

10        On Error GoTo PrepareSensitivitiesBloodCulture_Error

20        ResultsPerPage = Val(GetOptionSetting("ResultsPerPage", "25"))

30        ABCount = 0
40        If g.Rows > 7 Then
50            ABCount = g.Rows - 6
60        ElseIf g.TextMatrix(6, 0) <> "" Then
70            ABCount = 1
80        End If

90        If ValidStatus4MicroDept(SampleIDWithOffset, "B") = False Then
100           txtReport.SelBold = True
110           txtReport.SelColor = vbRed
120           txtReport.SelText = "Not Validated" & vbCrLf
130       End If

          'PRINT LINE FOR HEADING (PRINT ONLY WHEN ATLEAST ONE BOTTLE IS +VE
140       If Not BloodCultureBottleIsPositive(SampleIDWithOffset, "Aerobic") And Not BloodCultureBottleIsPositive(SampleIDWithOffset, "Anaerobic") And _
              Not BloodCultureBottleIsPositive(SampleIDWithOffset, "Fan") Then
              'print negative culture
150           GetPrintLineBloodCultureOrganismResult SampleIDWithOffset
160           If ColHasValue(1) Or ColHasValue(2) Or ColHasValue(3) Or ColHasValue(4) Or ColHasValue(5) Or ColHasValue(6) Then
170               GetPrintLineBloodCultureOrganisms SampleIDWithOffset
180           End If
190       Else
200           GetPrintLineBloodCultureOrganismResult SampleIDWithOffset
              'Get Print Bottle Line1
210           If GramIdentificationExists(SampleIDWithOffset) Then
220               lpc = UBound(udtPL) + 1
230               ReDim Preserve udtPL(0 To lpc)
240               udtPL(lpc).LineType = "BOLD10"
250               udtPL(lpc).LineText = "GRAM STAIN"
260           End If
270           GetPrintLineBloodCultureBottle SampleIDWithOffset, 1
280           GetPrintLineBloodCultureBottle SampleIDWithOffset, 2
290           GetPrintLineBloodCultureBottle SampleIDWithOffset, 3
300           GetPrintLineBloodCultureBottle SampleIDWithOffset, 4
310           GetPrintLineBloodCultureBottle SampleIDWithOffset, 5
320           GetPrintLineBloodCultureBottle SampleIDWithOffset, 6

330           If ColHasValue(1) Or ColHasValue(2) Or ColHasValue(3) Or ColHasValue(4) Or ColHasValue(5) Or ColHasValue(6) Then
340               GetPrintLineBloodCultureOrganisms SampleIDWithOffset
350           End If

360           If ABCount > 0 Then
370               GetPrintLineBloodCultureSensitivities
380           End If

390           lpc = UBound(udtPL)

400       End If

410       Exit Sub

PrepareSensitivitiesBloodCulture_Error:

          Dim strES As String
          Dim intEL As Integer

420       intEL = Erl
430       strES = Err.Description
440       LogError "frmMicroReport", "PrepareSensitivitiesBloodCulture", intEL, strES

End Sub







Private Function ColHasValue(Col As Integer) As Boolean

          Dim i As Integer

10        On Error GoTo ColHasValue_Error


20        ColHasValue = False
30        With g

40            For i = 0 To g.Rows - 1
50                If .TextMatrix(i, Col) <> "" Then
60                    ColHasValue = True
70                    Exit For
80                End If
90            Next i

100       End With

110       Exit Function

ColHasValue_Error:

          Dim strES As String
          Dim intEL As Integer

120       intEL = Erl
130       strES = Err.Description
140       LogError "frmMicroReport", "ColHasValue", intEL, strES

End Function

Private Function IsolateHasAntibiotics(Col As Integer) As Boolean

          Dim i As Integer

10        On Error GoTo IsolateHasAntibiotics_Error

20        IsolateHasAntibiotics = False

30        If g.Rows < 6 Then Exit Function

40        With g

50            For i = 6 To g.Rows - 1
60                If .TextMatrix(i, Col) <> "" Then
70                    IsolateHasAntibiotics = True
80                    Exit For
90                End If
100           Next i

110       End With


120       Exit Function

IsolateHasAntibiotics_Error:

          Dim strES As String
          Dim intEL As Integer

130       intEL = Erl
140       strES = Err.Description
150       LogError "frmMicroReport", "IsolateHasAntibiotics", intEL, strES

End Function


Public Sub FillG(ByVal SampleIDWithOffset As Double)

          Dim tb As Recordset
          Dim sql As String
          Dim R As Integer
          Dim Y As Integer
          Dim X As Integer
          Dim Found As Boolean
          Dim IsolateCount As Integer
          Dim MicroSite As String
          Dim Organisms As String

10        On Error GoTo FillG_Error

20        MicroSite = GetMicroSite(SampleIDWithOffset)

30        With g
40            .Clear
50            .TextMatrix(0, 0) = "Org Group"
60            .TextMatrix(1, 0) = "Org Name"
70            .TextMatrix(2, 0) = "Qualifier"
80            .TextMatrix(3, 0) = "Short Name"
90            .TextMatrix(4, 0) = "Report Name"
100           .TextMatrix(5, 0) = "Isolate #"

110           sql = "SET NOCOUNT ON " & _
                    "DECLARE @Tab table " & _
                    "( OrganismGroup nvarchar(100), OrganismName nvarchar(100), Qualifier nvarchar(50), " & _
                    "  IsolateNumber nvarchar(50), ShortName nvarchar(50), ReportName nvarchar(100), RowIndex int identity) " & _
                    "INSERT INTO @tab (OrganismGroup, OrganismName, Qualifier, IsolateNumber, ShortName, ReportName) " & _
                    "SELECT DISTINCT I.OrganismGroup, I.OrganismName, I.Qualifier, I.IsolateNumber, O.ShortName, O.ReportName " & _
                    "FROM Isolates I LEFT JOIN Organisms O ON O.Name = I.OrganismName " & _
                    "WHERE I.SampleID = '" & SampleIDWithOffset & "' " & _
                    "ORDER BY IsolateNumber " & _
                    "SELECT * FROM @Tab"
120           Set tb = New Recordset
130           RecOpenClient 0, tb, sql
140           If tb.EOF Then
150               .Clear
160           End If

170           IsolateCount = tb.recordCount
180           Do While Not tb.EOF
190               R = tb!IsolateNumber

200               .TextMatrix(0, R) = tb!OrganismGroup & ""
210               .TextMatrix(1, R) = tb!OrganismName & ""
220               .TextMatrix(2, R) = tb!Qualifier & ""
230               .TextMatrix(3, R) = tb!ShortName & ""
240               .TextMatrix(4, R) = tb!ReportName & ""
250               .TextMatrix(5, R) = tb!IsolateNumber

260               Organisms = Organisms & "'" & tb!OrganismGroup & "',"
270               tb.MoveNext
280           Loop
290           If Len(Organisms) > 1 Then
300               Organisms = Left(Organisms, Len(Organisms) - 1)

310               .Rows = 7
320               .AddItem ""
330               .RemoveItem 6

340               sql = "SELECT DISTINCT LTRIM(RTRIM(S.Antibiotic)) Antibiotic, B.ListOrder, LTRIM(RTRIM(COALESCE(A.ReportName, ''))) ReportName, " & _
                        "CASE S.RSI WHEN 'R' THEN 'Resistant' " & _
                        "         WHEN 'S' THEN 'Sensitive' " & _
                        "         WHEN 'I' THEN 'Intermediate' " & _
                        "         ELSE '' END RSI, S.IsolateNumber " & _
                        "FROM Sensitivities S INNER JOIN Antibiotics A ON S.AntibioticCode = A.Code " & _
                        "Inner Join (Select AntibioticName, Listorder from ABDefinitions Where Site = '" & MicroSite & "' " & _
                        "And OrganismGroup = '" & .TextMatrix(0, 1) & "') B on S.Antibiotic = B.AntibioticName " & _
                        "Where S.SampleID = '" & SampleIDWithOffset & "' AND S.Report = 1 AND COALESCE(S.Antibiotic,'') <> '' "
350               sql = sql & " UNION "
360               sql = sql & _
                        "SELECT DISTINCT LTRIM(RTRIM(S.Antibiotic)) Antibiotic, B.ListOrder, LTRIM(RTRIM(COALESCE(A.ReportName, ''))) ReportName, " & _
                        "CASE S.RSI WHEN 'R' THEN 'Resistant' " & _
                        "         WHEN 'S' THEN 'Sensitive' " & _
                        "         WHEN 'I' THEN 'Intermediate' " & _
                        "         ELSE '' END RSI, S.IsolateNumber " & _
                        "FROM Sensitivities S INNER JOIN Antibiotics A ON S.AntibioticCode = A.Code " & _
                        "Inner Join (Select AntibioticName, Listorder from ABDefinitions Where Site = '" & MicroSite & "' " & _
                        "And OrganismGroup = '" & .TextMatrix(0, 2) & "') B on S.Antibiotic = B.AntibioticName " & _
                        "Where S.SampleID = '" & SampleIDWithOffset & "' AND S.Report = 1 AND COALESCE(S.Antibiotic,'') <> '' "
370               sql = sql & " UNION "
380               sql = sql & _
                        "SELECT DISTINCT LTRIM(RTRIM(S.Antibiotic)) Antibiotic, B.ListOrder, LTRIM(RTRIM(COALESCE(A.ReportName, ''))) ReportName, " & _
                        "CASE S.RSI WHEN 'R' THEN 'Resistant' " & _
                        "         WHEN 'S' THEN 'Sensitive' " & _
                        "         WHEN 'I' THEN 'Intermediate' " & _
                        "         ELSE '' END RSI, S.IsolateNumber " & _
                        "FROM Sensitivities S INNER JOIN Antibiotics A ON S.AntibioticCode = A.Code " & _
                        "Inner Join (Select AntibioticName, Listorder from ABDefinitions Where Site = '" & MicroSite & "' " & _
                        "And OrganismGroup = '" & .TextMatrix(0, 3) & "') B on S.Antibiotic = B.AntibioticName " & _
                        "Where S.SampleID = '" & SampleIDWithOffset & "' AND S.Report = 1 AND COALESCE(S.Antibiotic,'') <> '' "
390               sql = sql & " UNION "
400               sql = sql & _
                        "SELECT DISTINCT LTRIM(RTRIM(S.Antibiotic)) Antibiotic, B.ListOrder, LTRIM(RTRIM(COALESCE(A.ReportName, ''))) ReportName, " & _
                        "CASE S.RSI WHEN 'R' THEN 'Resistant' " & _
                        "         WHEN 'S' THEN 'Sensitive' " & _
                        "         WHEN 'I' THEN 'Intermediate' " & _
                        "         ELSE '' END RSI, S.IsolateNumber " & _
                        "FROM Sensitivities S INNER JOIN Antibiotics A ON S.AntibioticCode = A.Code " & _
                        "Inner Join (Select AntibioticName, Listorder from ABDefinitions Where Site = '" & MicroSite & "' " & _
                        "And OrganismGroup = '" & .TextMatrix(0, 4) & "') B on S.Antibiotic = B.AntibioticName " & _
                        "Where S.SampleID = '" & SampleIDWithOffset & "' AND S.Report = 1 AND COALESCE(S.Antibiotic,'') <> '' "
410               sql = sql & " UNION "
420               sql = sql & _
                        "SELECT DISTINCT LTRIM(RTRIM(S.Antibiotic)) Antibiotic, B.ListOrder, LTRIM(RTRIM(COALESCE(A.ReportName, ''))) ReportName, " & _
                        "CASE S.RSI WHEN 'R' THEN 'Resistant' " & _
                        "         WHEN 'S' THEN 'Sensitive' " & _
                        "         WHEN 'I' THEN 'Intermediate' " & _
                        "         ELSE '' END RSI, S.IsolateNumber " & _
                        "FROM Sensitivities S INNER JOIN Antibiotics A ON S.AntibioticCode = A.Code " & _
                        "Inner Join (Select AntibioticName, Listorder from ABDefinitions Where Site = '" & MicroSite & "' " & _
                        "And OrganismGroup = '" & .TextMatrix(0, 5) & "') B on S.Antibiotic = B.AntibioticName " & _
                        "Where S.SampleID = '" & SampleIDWithOffset & "' AND S.Report = 1 AND COALESCE(S.Antibiotic,'') <> '' "
430               sql = sql & " UNION "
440               sql = sql & _
                        "SELECT DISTINCT LTRIM(RTRIM(S.Antibiotic)) Antibiotic, B.ListOrder, LTRIM(RTRIM(COALESCE(A.ReportName, ''))) ReportName, " & _
                        "CASE S.RSI WHEN 'R' THEN 'Resistant' " & _
                        "         WHEN 'S' THEN 'Sensitive' " & _
                        "         WHEN 'I' THEN 'Intermediate' " & _
                        "         ELSE '' END RSI, S.IsolateNumber " & _
                        "FROM Sensitivities S INNER JOIN Antibiotics A ON S.AntibioticCode = A.Code " & _
                        "Inner Join (Select AntibioticName, Listorder from ABDefinitions Where Site = '" & MicroSite & "' " & _
                        "And OrganismGroup = '" & .TextMatrix(0, 6) & "') B on S.Antibiotic = B.AntibioticName " & _
                        "Where S.SampleID = '" & SampleIDWithOffset & "' AND S.Report = 1 AND COALESCE(S.Antibiotic,'') <> '' "
450               sql = sql & "Order By B.ListOrder"

460               Set tb = New Recordset
470               RecOpenServer 0, tb, sql

480               Do While Not tb.EOF
490                   ABExists = True
500                   Found = False
510                   For X = 7 To .Rows - 1
520                       If .TextMatrix(X, 0) = tb!Antibiotic Or .TextMatrix(X, 0) = tb!ReportName Then
                              'antibiotic already added
530                           .row = X
540                           For Y = 1 To .Cols - 1
550                               If .TextMatrix(5, Y) = tb!IsolateNumber Then
560                                   .TextMatrix(.row, tb!IsolateNumber) = tb!RSI
570                                   Found = True
580                                   Exit For
590                               End If
600                           Next
610                       End If
620                   Next X
630                   If Not Found Then
640                       .AddItem IIf(tb!ReportName <> "", tb!ReportName & "", tb!Antibiotic & "")
650                       .row = g.Rows - 1
660                       For Y = 1 To .Cols - 1
670                           If .TextMatrix(5, Y) = tb!IsolateNumber Then
680                               .TextMatrix(.row, tb!IsolateNumber) = tb!RSI
690                               Exit For
700                           End If
710                       Next
720                   End If

730                   tb.MoveNext
740               Loop

                  'FILL IN FORECED ONES ********************************

750               sql = "SELECT LTRIM(RTRIM(A.AntibioticName)) AS Antibiotic, LTRIM(RTRIM(COALESCE(A.ReportName,''))) ReportName, " & _
                        "CASE S.RSI WHEN 'R' THEN 'Resistant' " & _
                        "         WHEN 'S' THEN 'Sensitive' " & _
                        "         WHEN 'I' THEN 'Intermediate' " & _
                        "         ELSE '' END RSI, S.IsolateNumber, " & _
                        "S.Report, S.RSI, S.CPOFlag, S.Result, S.RunDateTime, S.UserName " & _
                        "FROM Sensitivities S, Antibiotics A " & _
                        "Where S.SampleID = '" & SampleIDWithOffset & "' AND S.Report = 1 AND COALESCE(Antibiotic,'') <> '' " & _
                        "AND S.AntibioticCode = A.Code " & _
                        "AND S.Forced = 1"
760               Set tb = New Recordset
770               RecOpenServer 0, tb, sql
780               Do While Not tb.EOF
790                   Found = False
800                   For X = 7 To .Rows - 1
810                       If .TextMatrix(X, 0) = tb!Antibiotic Or .TextMatrix(X, 0) = tb!ReportName Then
                              'antibiotic already added
820                           .row = X
830                           For Y = 1 To .Cols - 1
840                               If .TextMatrix(5, Y) = tb!IsolateNumber Then
850                                   .TextMatrix(.row, tb!IsolateNumber) = tb!RSI
860                                   Found = True
870                                   Exit For
880                               End If
890                           Next
900                       End If
910                   Next X
920                   If Not Found Then
930                       .AddItem IIf(tb!ReportName <> "", tb!ReportName & "", tb!Antibiotic & "")
940                       .row = g.Rows - 1
950                       For Y = 1 To .Cols - 1
960                           If .TextMatrix(5, Y) = tb!IsolateNumber Then
970                               .TextMatrix(.row, tb!IsolateNumber) = tb!RSI
980                               Exit For
990                           End If
1000                      Next
1010                  End If

1020                  tb.MoveNext
1030              Loop



1040              If g.Rows > 7 Then
1050                  g.RemoveItem 6
1060              End If
1070          End If
1080      End With

1090      Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer

1100      intEL = Erl
1110      strES = Err.Description
1120      LogError "frmMicroReport", "FillG", intEL, strES, sql

End Sub

Private Sub PrepareSensitivitiesOther()

          Dim ABCount As Integer
          Dim lpc As Integer
          Dim ResultsPerPage As Integer
          Dim StartIndex As Integer
          Dim EndIndex As Integer
          Dim OldPrintLines As Integer

10        On Error GoTo PrepareSensitivitiesOther_Error

20        ResultsPerPage = Val(GetOptionSetting("ResultsPerPage", "25"))

30        ABCount = 0
40        If g.Rows > 7 Then
50            ABCount = g.Rows - 6
60        ElseIf g.TextMatrix(6, 0) <> "" Then
70            ABCount = 1
80        End If

90        StartIndex = UBound(udtPL) + 1
100       OldPrintLines = StartIndex - 1

110       StartIndex = UBound(udtPL) + 1

120       If ColHasValue(1) Or ColHasValue(2) Or ColHasValue(3) Or ColHasValue(4) Then

130           GetPrintLineOrganismsOther
140       End If

150       EndIndex = UBound(udtPL)

160       If ABCount > 0 Then
170           GetPrintLineBloodCultureSensitivities
180       End If

190       lpc = UBound(udtPL)

200       Exit Sub

PrepareSensitivitiesOther_Error:

          Dim strES As String
          Dim intEL As Integer

210       intEL = Erl
220       strES = Err.Description
230       LogError "modNewMicro", "PrepareSensitivitiesOther", intEL, strES

End Sub


Private Sub PrepareSensitivities(ByVal SampleIDWithOffset As Double, _
                                 ByVal MaxSensitives As String)

          Dim tb As Recordset
          Dim tbr As Recordset
          Dim sql As String
          Dim intSensCounter(1 To 8) As Long
          Dim strOrgGroup As String
          Dim RSI As String
          Dim Site As String
          Dim sqlBase As String
10        ReDim ABResults(0 To 0) As ABResult
          Dim n As Long
          Dim X As Long
          Dim Y As Long
          Dim Found As Boolean
          Dim s As String
          Dim OrgNames As String
          Dim OrgGroups As String

20        On Error GoTo PrepareSensitivities_Error

30        sql = "SELECT Site from MicroSiteDetails " & _
                "WHERE SampleID = " & SampleIDWithOffset
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        If tb.EOF Then
70            Site = "Generic"
80        ElseIf tb!Site & "" = "" Then
90            Site = "Generic"
100       Else
110           Site = tb!Site
120       End If

130       For X = 1 To 8
140           intSensCounter(X) = Val(MaxSensitives)
150       Next

160       sql = "SELECT * from isolates " & _
                "WHERE SampleID = " & SampleIDWithOffset
170       Set tb = New Recordset
180       RecOpenServer 0, tb, sql
190       If tb.EOF Then Exit Sub

200       strOrgGroup = ""
210       ABResults(0).Antibiotic = Left$("Sensitivities:" & Space$(20), 20)
220       Do While Not tb.EOF
230           ABResults(0).Result(tb!IsolateNumber) = tb!OrganismName & ""
240           ABResults(0).Group(tb!IsolateNumber) = tb!OrganismGroup & ""
250           ABResults(0).Qualifier(tb!IsolateNumber) = tb!Qualifier & ""
260           strOrgGroup = strOrgGroup & " OrganismGroup = '" & tb!OrganismGroup & _
                            "' or "
270           tb.MoveNext
280       Loop
290       If strOrgGroup <> "" Then
300           strOrgGroup = Left(strOrgGroup, Len(strOrgGroup) - 3)
310       End If

320       sqlBase = "SELECT Distinct AntibioticName, ListOrder from ABDefinitions " & _
                    "WHERE AntibioticName IN " & _
                    "  (SELECT AntibioticName from Antibiotics " & _
                    "  WHERE Code IN " & _
                    "     (SELECT DISTINCT AntibioticCode " & _
                    "     FROM Sensitivities " & _
                    "     WHERE SampleID = " & SampleIDWithOffset & " ) ) " & _
                    "AND (" & strOrgGroup & ") "
330       sql = sqlBase & "AND Site = '" & Site & "' " & "ORDER BY ListOrder"

340       sql = "SELECT DISTINCT AntibioticCode FROM Sensitivities WHERE " & _
                "SampleID = " & SampleIDWithOffset
350       Set tb = New Recordset
360       RecOpenServer 0, tb, sql
370       If tb.EOF Then
380           sql = sqlBase & "AND Site = 'Generic' " & "ORDER BY ListOrder"
390           Set tb = New Recordset
400           RecOpenServer 0, tb, sql
410       End If

420       Do While Not tb.EOF
              '  For n = 0 To UBound(ABResults)
              '    Found = False
              '    If ABResults(n).Antibiotic = AntibioticNameFor(tb!Antibioticcode) Then
              '      Found = True
              '      Exit For
              '    End If
              '  Next
              '  If Not Found Then
430           ReDim Preserve ABResults(UBound(ABResults) + 1)
440           ABResults(UBound(ABResults)).Antibiotic = AntibioticNameFor( _
                                                        tb!AntibioticCode)
450           ABResults(UBound(ABResults)).ReportName = AntibioticRerportNameFor( _
                                                        tb!AntibioticCode)
              '  End If
460           tb.MoveNext
470       Loop

480       For Y = 0 To UBound(ABResults)
490           For X = 1 To 4
500               If ABResults(Y).Result(X) = "" Then
510                   sql = "SELECT S.RSI, S.Report, A.Code " & _
                            "from Sensitivities as S, Antibiotics as A " & _
                            "WHERE SampleID = " & SampleIDWithOffset & " " & _
                            "AND S.IsolateNumber = '" & X & "' " & _
                            "AND A.Code = S.AntibioticCode " & _
                            "AND A.AntibioticName = '" & ABResults(Y).Antibiotic & "'"
520                   Set tbr = New Recordset
530                   RecOpenServer 0, tbr, sql
540                   If Not tbr.EOF Then
550                       ABResults(Y).Report(X) = IIf(Not IsNull(tbr!Report), tbr!Report, False)
560                       If tbr!RSI & "" = "R" Then
570                           ABResults(Y).Result(X) = "R"
580                       ElseIf tbr!RSI & "" = "S" Then
590                           If intSensCounter(X) > 0 Then
600                               ABResults(Y).Result(X) = "S"
610                               intSensCounter(X) = intSensCounter(X) - 1
620                           End If
630                           If tbr!Report = True Then
640                               ABResults(Y).Result(X) = "S"
650                           End If
660                       ElseIf tbr!RSI & "" = "I" Then
670                           ABResults(Y).Result(X) = "I"
680                       End If
690                   Else
700                       ABResults(Y).Result(X) = ""
710                   End If
720               End If
730           Next
740       Next

          'Is C&S valid
750       If ValidStatus4MicroDept(SampleIDWithOffset, "D") = False Then
760           txtReport.SelBold = True
770           txtReport.SelColor = vbRed
780           txtReport.SelText = "Not Validated"
790       End If

800       OrgNames = ""
810       OrgGroups = ""
820       s = ""
830       For X = 1 To 4
840           If ABResults(0).Result(X) <> "" Then
850               s = ABResults(0).Result(X)
860               If ABResults(0).Qualifier(X) <> "" Then
870                   s = ABResults(0).Qualifier(X) & " " & ABResults(0).Result(X)
880               End If

890               If Len(OrgNames & s) < 80 Then
900                   OrgNames = OrgNames & s & ". "
910               Else
920                   OrgNames = OrgNames & vbCrLf & s
930               End If
940           End If
950           If ABResults(0).Group(X) <> "" Then
960               s = ABResults(0).Group(X)
970               If Len(OrgGroups & s) < 80 Then
980                   OrgGroups = OrgGroups & s & ". "
990               Else
1000                  OrgGroups = OrgGroups & ". " & vbCrLf & s
1010              End If
1020          End If
1030      Next

1040      If OrgNames <> "" Or OrgGroups <> "" Then
1050          txtReport.SelText = vbCrLf
1060          txtReport.SelUnderline = True
1070          txtReport.SelColor = vbBlack
1080          txtReport.SelText = "Culture:"
1090          txtReport.SelUnderline = False
1100          txtReport.SelText = " " & vbCrLf
1110          txtReport.SelUnderline = False
1120      End If
          '*********************no need to print orggroups
          'If OrgGroups <> "" Then
          '    txtReport.SelBold = False
          '    txtReport.SelColor = vbBlack
          '    txtReport.SelText = Trim(OrgGroups) & vbCrLf
          'End If
1130      If OrgNames <> "" Then
1140          txtReport.SelBold = False
1150          txtReport.SelColor = vbBlack
1160          txtReport.SelText = Trim(OrgNames) & vbCrLf
1170      End If


1180      FillSensitivities ABResults

1190      Exit Sub

PrepareSensitivities_Error:

          Dim strES As String
          Dim intEL As Integer

1200      intEL = Erl
1210      strES = Err.Description
1220      LogError "frmMicroReport", "PrepareSensitivities", intEL, strES, sql

End Sub

Private Sub FillSensitivities(ByRef ABResults() As ABResult)

          Dim ResultPresent As Boolean
          Dim Y As Long
          Dim X As Long
          Dim s As String

10        On Error GoTo FillSensitivities_Error

20        If UBound(ABResults) = 0 Then Exit Sub

30        For X = 1 To 8
40            If Trim$(ABResults(0).Result(X)) <> "" Then
50                txtReport.SelText = vbCrLf
60                txtReport.SelColor = vbBlack

70                s = "Sensitivities for Organism "
80                s = s & Trim(ABResults(0).Result(X))
90                txtReport.SelBold = True
100               txtReport.SelText = s & vbCrLf & vbCrLf
110               txtReport.SelColor = vbBlack
120               txtReport.SelBold = False
130               For Y = 1 To UBound(ABResults)
140                   If ABResults(Y).ReportName <> "" Then
150                       s = Left$(ABResults(Y).ReportName & Space$(20), 20)
160                   Else
170                       s = Left$(ABResults(Y).Antibiotic & Space$(20), 20)
180                   End If

190                   ResultPresent = True
200                   If ABResults(Y).Result(X) = "R" Then
210                       s = s & "Resistant"
220                   ElseIf ABResults(Y).Result(X) = "S" Then
230                       s = s & "Sensitive"
240                   ElseIf ABResults(Y).Result(X) = "I" Then
250                       s = s & "Intermediate"
260                   Else
270                       ResultPresent = False
280                   End If
290                   If ResultPresent And ABResults(Y).Report(X) Then
300                       txtReport.SelUnderline = False
310                       txtReport.SelText = s & vbCrLf
320                       txtReport.SelColor = vbBlack
330                   End If
340               Next
350           End If
360       Next
370       txtReport.SelColor = vbBlack

380       Exit Sub

FillSensitivities_Error:

          Dim strES As String
          Dim intEL As Integer

390       intEL = Erl
400       strES = Err.Description
410       LogError "frmMicroReport", "FillSensitivities", intEL, strES

End Sub


Private Sub Timer1_Timer()

      'tmrRefresh.Interval set to 1000
10        On Error GoTo Timer1_Timer_Error

20        PBar = PBar + 1

30        If PBar = PBar.Max Then
40            Unload Me
50        End If

60        Exit Sub

Timer1_Timer_Error:

          Dim strES As String
          Dim intEL As Integer



70        intEL = Erl
80        strES = Err.Description
90        LogError "frmMicroReport", "Timer1_Timer", intEL, strES


End Sub




Public Property Let PatName(ByVal NewValue As String)

10        pPatName = NewValue
20        lblName = NewValue

End Property
Public Property Let PatChart(ByVal NewValue As String)

10        pPatChart = NewValue
20        lblChart = NewValue

End Property

Public Property Let PatDoB(ByVal NewValue As String)

10        pPatDoB = NewValue
20        lblDoB = NewValue

End Property


Public Property Let PatSex(ByVal NewValue As String)

10        pPatSex = NewValue
20        lblSex = NewValue

End Property

Public Property Let PatWard(ByVal NewValue As String)
10        pPatWard = NewValue
End Property

Public Property Get PatWard() As String
10        PatWard = pPatWard
End Property


Public Property Let PatClinician(ByVal NewValue As String)
10        pPatClinician = NewValue
End Property

Public Property Get PatClinician() As String
10        PatClinician = pPatClinician
End Property

Public Property Let PatGP(ByVal NewValue As String)
10        pPatGP = NewValue
End Property

Public Property Get PatGP() As String
10        PatGP = pPatGP
End Property


Public Function GetGramIdentification(ByVal SampleIDWithOffset As Double, ByVal Isolate As Byte) As String

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo GetGramIdentification_Error

20        sql = "Select Gram From UrineIdent Where SampleID = %sampleid And Isolate = %isolate"
30        sql = Replace(sql, "%sampleid", SampleIDWithOffset)
40        sql = Replace(sql, "%isolate", Isolate)

50        Set tb = New Recordset
60        RecOpenClient 0, tb, sql

70        If Not tb.EOF Then
80            GetGramIdentification = tb!Gram & ""
90        Else
100           GetGramIdentification = ""
110       End If

120       Exit Function

GetGramIdentification_Error:

          Dim strES As String
          Dim intEL As Integer

130       intEL = Erl
140       strES = Err.Description
150       LogError "modNewMicro", "GetGramIdentification", intEL, strES, sql

End Function

Public Function GetBloodCultureBottleInterval(ByVal SampleIDWithOffset As String, ByVal BottleType As String) As String

          Dim tb As Recordset
          Dim sql As String
          Dim TypeOfTest As String

10        On Error GoTo GetBloodCultureBottleInterval_Error

20        Select Case BottleType
          Case "Aerobic": TypeOfTest = GetOptionSetting("BcAerobicBottle", "BSA")
30        Case "Anaerobic": TypeOfTest = GetOptionSetting("BcAnarobicBottle", "BSN")
40        Case "Fan": TypeOfTest = GetOptionSetting("BcFanBottle", "BFA")
50        Case Else: TypeOfTest = ""
60        End Select

70        If TypeOfTest = "" Then
80            GetBloodCultureBottleInterval = ""
90            Exit Function
100       End If

110       sql = "Select TTD From BloodCultureResults Where SampleID = %sampleid And TypeOfTest = '%typeoftest'"
120       sql = Replace(sql, "%sampleid", SampleIDWithOffset)
130       sql = Replace(sql, "%typeoftest", TypeOfTest)

140       Set tb = New Recordset
150       RecOpenClient 0, tb, sql

160       If tb.EOF Then
170           GetBloodCultureBottleInterval = ""
180       Else
190           GetBloodCultureBottleInterval = tb!TTD & ""
200       End If

210       Exit Function

GetBloodCultureBottleInterval_Error:

          Dim strES As String
          Dim intEL As Integer

220       intEL = Erl
230       strES = Err.Description
240       LogError "modNewMicro", "GetBloodCultureBottleInterval", intEL, strES, sql

End Function

Public Function BloodCultureBottleIsPositive(ByVal SampleIDWithOffset As Double, ByVal BottleType As String) As Boolean

          Dim tb As Recordset
          Dim sql As String
          Dim TypeOfTest As String

10        On Error GoTo BloodCultureBottleIsPositive_Error

20        Select Case BottleType
          Case "Aerobic": TypeOfTest = GetOptionSetting("BcAerobicBottle", "BSA")
30        Case "Anaerobic": TypeOfTest = GetOptionSetting("BcAnarobicBottle", "BSN")
40        Case "Fan": TypeOfTest = GetOptionSetting("BcFanBottle", "BFA")
50        Case Else: TypeOfTest = ""
60        End Select

70        If TypeOfTest = "" Then
80            BloodCultureBottleIsPositive = False
90            Exit Function
100       End If

110       sql = "Select Result From BloodCultureResults Where SampleID = %sampleid And TypeOfTest = '%typeoftest'"
120       sql = Replace(sql, "%sampleid", SampleIDWithOffset)
130       sql = Replace(sql, "%typeoftest", TypeOfTest)

140       Set tb = New Recordset
150       RecOpenClient 0, tb, sql

160       If tb.EOF Then
170           BloodCultureBottleIsPositive = False
180       Else
190           BloodCultureBottleIsPositive = (tb!Result & "" = "+")
200       End If


210       Exit Function

BloodCultureBottleIsPositive_Error:

          Dim strES As String
          Dim intEL As Integer

220       intEL = Erl
230       strES = Err.Description
240       LogError "modNewMicro", "BloodCultureBottleIsPositive", intEL, strES, sql

End Function

Public Function BloodCultureBottleExists(ByVal SampleIDWithOffset As Double, ByVal BottleType As String) As Boolean

      Dim tb As Recordset
      Dim sql As String
      Dim TypeOfTest As String

10    On Error GoTo BloodCultureBottleExists_Error

20    Select Case BottleType
          Case "Aerobic": TypeOfTest = GetOptionSetting("BcAerobicBottle", "BSA")
30        Case "Anaerobic": TypeOfTest = GetOptionSetting("BcAnarobicBottle", "BSN")
40        Case "Fan": TypeOfTest = GetOptionSetting("BcFanBottle", "BFA")
50        Case Else: TypeOfTest = ""
60    End Select
70    sql = "SELECT Count(*) AS Cnt FROM BloodCultureResults WHERE SampleID = " & SampleIDWithOffset & " AND TypeOfTest = '" & TypeOfTest & "'"
80    Set tb = New Recordset
90    RecOpenServer 0, tb, sql
100   BloodCultureBottleExists = (tb!Cnt > 0)

110   Exit Function

BloodCultureBottleExists_Error:

      Dim strES As String
      Dim intEL As Integer

120   intEL = Erl
130   strES = Err.Description
140   LogError "frmMicroReport", "BloodCultureBottleExists", intEL, strES, sql

End Function

Private Function GramIdentificationExists(ByVal SampleIDWithOffset As Double) As Boolean

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo GramIdentificationExists_Error

20        sql = "Select Count(*) as Cnt From UrineIdent Where SampleID = %sampleid"
30        sql = Replace(sql, "%sampleid", SampleIDWithOffset)

40        Set tb = New Recordset
50        RecOpenClient 0, tb, sql

60        GramIdentificationExists = (tb!Cnt > 0)
70        Exit Function

GramIdentificationExists_Error:

          Dim strES As String
          Dim intEL As Integer

80        intEL = Erl
90        strES = Err.Description
100       LogError "modNewMicro", "GramIdentificationExists", intEL, strES, sql

End Function

Private Sub GetPrintLineBloodCultureBottle(ByVal SampleIDWithOffset As Double, ByVal BottleLine As Integer)

      Dim BottleName As String
      Dim GramStain As String
      Dim Interval As String
      Dim lpc As Integer

10    On Error GoTo GetPrintLineBloodCultureBottle_Error

20    Select Case BottleLine
          Case 1:
30            If BloodCultureBottleExists(SampleIDWithOffset, "Aerobic") Then
40                BottleName = FormatString("Bottle A", 10, , AlignLeft)
50                Interval = GetBloodCultureBottleInterval(SampleIDWithOffset, "Aerobic")
60                If Interval <> "" Then
70                    If Interval < 12 Then
80                        Interval = Interval & " hr(s)"
90                    ElseIf Interval >= 12 And Interval < 24 Then
100                       Interval = "<1 day"
110                   ElseIf Interval >= 24 Then
120                       Interval = (Interval \ 24) & " day(s)"
130                   End If
140               End If
150               Interval = FormatString(Interval, 8, , AlignRight)
160           End If
170       Case 2:
180           BottleName = FormatString("", 10, , AlignCenter)
190           Interval = FormatString("", 8, , AlignCenter)
200       Case 3:
210           If BloodCultureBottleExists(SampleIDWithOffset, "Anaerobic") Then
220               BottleName = FormatString("Bottle B", 10, , AlignLeft)
230               Interval = GetBloodCultureBottleInterval(SampleIDWithOffset, "Anaerobic")
240               If Interval <> "" Then
250                   If Interval < 12 Then
260                       Interval = Interval & " hr(s)"
270                   ElseIf Interval >= 12 And Interval < 24 Then
280                       Interval = "<1 day"
290                   ElseIf Interval >= 24 Then
300                       Interval = (Interval \ 24) & " day(s)"
310                   End If
320               End If
330               Interval = FormatString(Interval, 8, , AlignRight)
340           End If
350       Case 4:
360           BottleName = FormatString("", 10, , AlignCenter)
370           Interval = FormatString("", 8, , AlignCenter)
380       Case 5:
390           If BloodCultureBottleExists(SampleIDWithOffset, "Fan") Then
400               BottleName = FormatString("Bottle C", 10, , AlignLeft)
410               Interval = GetBloodCultureBottleInterval(SampleIDWithOffset, "Fan")
420               If Interval <> "" Then
430                   If Interval < 12 Then
440                       Interval = Interval & " hr(s)"
450                   ElseIf Interval >= 12 And Interval < 24 Then
460                       Interval = "<1 day"
470                   ElseIf Interval >= 24 Then
480                       Interval = (Interval \ 24) & " day(s)"
490                   End If
500               End If
510               Interval = FormatString(Interval, 8, , AlignRight)
520           End If
530       Case 6:
540           BottleName = FormatString("", 10, , AlignCenter)
550           Interval = FormatString("", 8, , AlignCenter)
560   End Select

570   GramStain = FormatString(GetGramIdentification(SampleIDWithOffset, BottleLine), 36, , AlignLeft)

580   If LTrim(RTrim((GramStain))) <> "" Then
590       lpc = UBound(udtPL) + 1
600       ReDim Preserve udtPL(0 To lpc)
610       udtPL(lpc).LineType = "NORMAL10"
620       udtPL(lpc).LineText = BottleName & " " & BottleLine & ". " & GramStain
630   End If


640   Exit Sub

GetPrintLineBloodCultureBottle_Error:

      Dim strES As String
      Dim intEL As Integer

650   intEL = Erl
660   strES = Err.Description
670   LogError "modNewMicro", "GetPrintLineBloodCultureBottle", intEL, strES

End Sub

Private Sub GetPrintLineBloodCultureOrganisms(ByVal SampleIDWithOffset As Double)

      Dim lpc As Integer
      Dim i As Integer
      Dim Duplicates As String
      Dim AlreadyPrinted As String
      Dim s As String
      Dim Start As Integer
      Dim Interval As String
      Dim OrgNo As String

10    On Error GoTo GetPrintLineBloodCultureOrganisms_Error

20    Duplicates = ""
30    AlreadyPrinted = ""
40    Start = 1

      'Culture Heading
50    lpc = UBound(udtPL) + 1
60    ReDim Preserve udtPL(0 To lpc)
70    udtPL(lpc).LineType = "BOLD10"
80    udtPL(lpc).LineText = FormatString("CULTURE", 46, , AlignLeft)


90    For i = 1 To 6

100       If ((i = 1 Or i = 2) And Not BloodCultureBottleIsPositive(SampleIDWithOffset, "Aerobic")) _
             Or ((i = 3 Or i = 4) And Not BloodCultureBottleIsPositive(SampleIDWithOffset, "Anaerobic")) _
             Or ((i = 5 Or i = 6) And Not BloodCultureBottleIsPositive(SampleIDWithOffset, "Fan")) Then
110           If i = 1 Or i = 2 Then

120               Interval = GetBloodCultureBottleInterval(SampleIDWithOffset, "Aerobic")
130               OrgNo = i & "."
140           ElseIf i = 3 Or i = 4 Then
150               Interval = GetBloodCultureBottleInterval(SampleIDWithOffset, "Anaerobic")
160               OrgNo = i & "."
170           ElseIf i = 5 Or i = 6 Then
180               Interval = GetBloodCultureBottleInterval(SampleIDWithOffset, "Fan")
190               OrgNo = i & "."
200           End If
210           If Interval <> "" Then
220               If Interval < 12 Then
230                   Interval = Interval & " hour(s)"
240               ElseIf Interval >= 12 And Interval < 24 Then
250                   Interval = "0.5 day"
260               ElseIf Interval >= 24 Then
270                   Interval = (Interval \ 24) & " day(s)"
280               End If
290           End If
300           s = FormatString("No growth at " & Interval, 30, , AlignLeft)
310           Select Case i
                  Case 1, 2:
320                   If BloodCultureBottleExists(SampleIDWithOffset, "Aerobic") Then
330                       s = s & FormatString("Bottle A", 12, , AlignRight)
340                   Else
350                       s = ""
360                   End If
370               Case 3, 4:
380                   If BloodCultureBottleExists(SampleIDWithOffset, "Anaerobic") Then
390                       s = s & FormatString("Bottle B", 12, , AlignRight)
400                   Else
410                       s = ""
420                   End If
430               Case 5, 6:
440                   If BloodCultureBottleExists(SampleIDWithOffset, "Fan") Then
450                       s = s & FormatString("Bottle C", 12, , AlignRight)
460                   Else
470                       s = ""
480                   End If
490           End Select
500           If Trim(s) <> "" Then
510           lpc = UBound(udtPL) + 1
520           ReDim Preserve udtPL(0 To lpc)
530           udtPL(lpc).LineType = "NORMAL10"
540           udtPL(lpc).LineText = FormatString(OrgNo, 2, , AlignLeft) & s
550           End If
560           OrgNo = ""
570           i = i + 1
580       Else

590           s = FormatString(g.TextMatrix(1, i), 30, , AlignLeft)
600           OrgNo = i & "."
610           If Trim(s) <> "" Then
620               Select Case i
                      Case 1, 2:
630                       If BloodCultureBottleExists(SampleIDWithOffset, "Aerobic") Then
640                           s = FormatString(OrgNo, 2, , AlignLeft) & s & FormatString("Bottle A", 12, , AlignRight)
650                       Else
660                           s = ""
670                       End If
680                   Case 3, 4:
690                       If BloodCultureBottleExists(SampleIDWithOffset, "Anaerobic") Then
700                           s = FormatString(OrgNo, 2, , AlignLeft) & s & FormatString("Bottle B", 12, , AlignRight)
710                       Else
720                           s = ""
730                       End If
740                   Case 5, 6:
750                       If BloodCultureBottleExists(SampleIDWithOffset, "Fan") Then
760                           s = FormatString(OrgNo, 2, , AlignLeft) & s & FormatString("Bottle C", 12, , AlignRight)
770                       Else
780                           s = ""
790                       End If
800               End Select
810           End If
820           If Trim(s) <> "" Then
830               lpc = UBound(udtPL) + 1
840               ReDim Preserve udtPL(0 To lpc)
850               udtPL(lpc).LineType = "NORMAL10"
860               udtPL(lpc).LineText = s
870           End If
880           OrgNo = ""
890       End If
900   Next i

910   Exit Sub

GetPrintLineBloodCultureOrganisms_Error:

      Dim strES As String
      Dim intEL As Integer

920   intEL = Erl
930   strES = Err.Description
940   LogError "modNewMicro", "GetPrintLineBloodCultureOrganisms", intEL, strES


End Sub

Private Sub GetPrintLineBloodCultureSensitivities()

          Dim lpc As Integer
          Dim s As String
          Dim i As Integer
          Dim SIndex As Integer

10        On Error GoTo GetPrintLineBloodCultureSensitivities_Error

20        SIndex = 6


30        lpc = UBound(udtPR) + 1
40        ReDim Preserve udtPR(0 To lpc)
50        udtPR(lpc).LineType = "BOLD10"
60        udtPR(lpc).LineText = "SUSCEPTIBILITIES" & FormatString(" ", 5)

70        If IsolateHasAntibiotics(1) Then
80            udtPR(lpc).LineText = udtPR(lpc).LineText & FormatString("1", 3, , AlignCenter)
90        Else
100           udtPR(lpc).LineText = udtPR(lpc).LineText & FormatString(" ", 3, , AlignCenter)
110       End If
120       If IsolateHasAntibiotics(2) Then
130           udtPR(lpc).LineText = udtPR(lpc).LineText & FormatString("2", 2, , AlignRight)
140       Else
150           udtPR(lpc).LineText = udtPR(lpc).LineText & FormatString(" ", 2, , AlignRight)
160       End If
170       If IsolateHasAntibiotics(3) Then
180           udtPR(lpc).LineText = udtPR(lpc).LineText & FormatString("3", 2, , AlignRight)
190       Else
200           udtPR(lpc).LineText = udtPR(lpc).LineText & FormatString(" ", 2, , AlignRight)
210       End If
220       If IsolateHasAntibiotics(4) Then
230           udtPR(lpc).LineText = udtPR(lpc).LineText & FormatString("4", 2, , AlignRight)
240       Else
250           udtPR(lpc).LineText = udtPR(lpc).LineText & FormatString(" ", 2, , AlignRight)
260       End If
270       If IsolateHasAntibiotics(5) Then
280           udtPR(lpc).LineText = udtPR(lpc).LineText & FormatString("5", 2, , AlignRight)
290       Else
300           udtPR(lpc).LineText = udtPR(lpc).LineText & FormatString(" ", 2, , AlignRight)
310       End If
320       If IsolateHasAntibiotics(6) Then
330           udtPR(lpc).LineText = udtPR(lpc).LineText & FormatString("6", 2, , AlignRight)
340       Else
350           udtPR(lpc).LineText = udtPR(lpc).LineText & FormatString(" ", 2, , AlignRight)
360       End If




370       For i = SIndex To g.Rows - 1
380           lpc = UBound(udtPR) + 1
390           ReDim Preserve udtPR(0 To lpc)
400           s = FormatString(g.TextMatrix(i, 0), 21, , AlignLeft) & _
                  FormatString(Left$(g.TextMatrix(i, 1), 1), 3, , AlignCenter) & _
                  FormatString(Left$(g.TextMatrix(i, 2), 1), 2, , AlignRight) & _
                  FormatString(Left$(g.TextMatrix(i, 3), 1), 2, , AlignRight) & _
                  FormatString(Left$(g.TextMatrix(i, 4), 1), 2, , AlignRight) & _
                  FormatString(Left$(g.TextMatrix(i, 5), 1), 2, , AlignRight) & _
                  FormatString(Left$(g.TextMatrix(i, 6), 1), 2, , AlignRight)

410           udtPR(lpc).LineType = "NORMAL10"
420           udtPR(lpc).LineText = s
430       Next i




440       Exit Sub

GetPrintLineBloodCultureSensitivities_Error:

          Dim strES As String
          Dim intEL As Integer

450       intEL = Erl
460       strES = Err.Description
470       LogError "modNewMicro", "GetPrintLineBloodCultureSensitivities", intEL, strES

End Sub

Private Sub GetPrintLineBloodCultureOrganismResult(ByVal SampleIDWithOffset As String)

          Dim lpc As Integer
          Dim i As Integer
          Dim s As String
          Dim Start As Integer
          Dim Interval As String
          Dim BottleAPrinted As Boolean
          Dim BottleBPrinted As Boolean
          Dim BottleCPrinted As Boolean

10        On Error GoTo GetPrintLineBloodCultureOrganismResult_Error


20        Start = 1
30        BottleAPrinted = False
40        BottleBPrinted = False
50        BottleCPrinted = False

          'Culture Heading
60        lpc = UBound(udtPL) + 1
70        ReDim Preserve udtPL(0 To lpc)
80        udtPL(lpc).LineType = "BOLD10"
90        udtPL(lpc).LineText = FormatString("RESULT", 46, , AlignLeft)


100       For i = 1 To 6

110           If (i = 1 Or i = 2) And Not BottleAPrinted And BloodCultureBottleExists(SampleIDWithOffset, "Aerobic") Then
120               If BloodCultureBottleIsPositive(SampleIDWithOffset, "Aerobic") Then
130                   s = "Bottle A Flagged Positive @ "
140               Else
150                   s = "Bottle A Flagged Negative @ "
160               End If
170               Interval = GetBloodCultureBottleInterval(SampleIDWithOffset, "Aerobic")
180               BottleAPrinted = True
190           ElseIf (i = 3 Or i = 4) And Not BottleBPrinted And BloodCultureBottleExists(SampleIDWithOffset, "Anaerobic") Then
200               If BloodCultureBottleIsPositive(SampleIDWithOffset, "Anaerobic") Then
210                   s = "Bottle B Flagged Positive @ "
220               Else
230                   s = "Bottle B Flagged Negative @ "
240               End If
250               Interval = GetBloodCultureBottleInterval(SampleIDWithOffset, "Anaerobic")
260               BottleBPrinted = True
270           ElseIf (i = 5 Or i = 6) And Not BottleCPrinted And BloodCultureBottleExists(SampleIDWithOffset, "Fan") Then
280               If BloodCultureBottleIsPositive(SampleIDWithOffset, "Fan") Then
290                   s = "Bottle C Flagged Positive @ "
300               Else
310                   s = "Bottle C Flagged Negative @ "
320               End If
330               Interval = GetBloodCultureBottleInterval(SampleIDWithOffset, "Fan")
340               BottleCPrinted = True
350           End If
360           If Interval <> "" Then
370               If Interval < 12 Then
380                   Interval = Interval & " hr(s)"
390               ElseIf Interval >= 12 And Interval < 24 Then
400                   Interval = "<1 day"
410               ElseIf Interval >= 24 Then
420                   Interval = (Interval \ 24) & " day(s)"
430               End If
440           End If
450           If s <> "" Then
460               lpc = UBound(udtPL) + 1
470               ReDim Preserve udtPL(0 To lpc)
480               udtPL(lpc).LineType = "NORMAL10"
490               udtPL(lpc).LineText = FormatString(s & Interval, 46, , AlignLeft)
500           End If
510           Interval = ""
520           s = ""

530       Next i

540       Exit Sub

GetPrintLineBloodCultureOrganismResult_Error:

          Dim strES As String
          Dim intEL As Integer

550       intEL = Erl
560       strES = Err.Description
570       LogError "frmMicroReport", "GetPrintLineBloodCultureOrganismResult", intEL, strES


End Sub


Private Sub GetPrintLineInterimHeading()

          Dim lpc As Integer
          'PRINT LINE FOR REPORT TYPE

10        On Error GoTo GetPrintLineInterimHeading_Error

20        lpc = UBound(udtPL) + 1
30        ReDim Preserve udtPL(0 To lpc)
40        udtPL(lpc).LineType = "BOLD10"
50        udtPL(lpc).LineText = Space(33) & FormatString(UCase$("Final Report"), 12, , AlignCenter) & Space(37)

60        lpc = UBound(udtPR) + 1
70        ReDim Preserve udtPR(0 To lpc)
80        udtPR(lpc).LineType = "BOLD10"
90        udtPR(lpc).LineText = ""

100       Exit Sub

GetPrintLineInterimHeading_Error:

          Dim strES As String
          Dim intEL As Integer

110       intEL = Erl
120       strES = Err.Description
130       LogError "modNewMicro", "GetPrintLineInterimHeading", intEL, strES


End Sub

Private Sub GetPrintLineOrganismsOther()

          Dim lpc As Integer
          Dim i As Integer
          Dim s As String
          Dim Start As Integer
          Dim OrgNo As String
          Dim Multiline As String
          Dim LastWordIndex As Integer


10        On Error GoTo GetPrintLineOrganismsOther_Error

20        Start = 1
30        Multiline = ""

          'Culture Heading
40        lpc = UBound(udtPL) + 1
50        ReDim Preserve udtPL(0 To lpc)
60        udtPL(lpc).LineType = "BOLD10"
70        udtPL(lpc).LineText = FormatString("CULTURE", 46, , AlignLeft)

80        For i = 1 To 4

90            If g.TextMatrix(2, i) <> "" Then
100               s = g.TextMatrix(2, i) & " "
110           End If

120           If g.TextMatrix(4, i) <> "" Then
130               s = i & ". " & s & g.TextMatrix(4, i)
140           Else
150               If g.TextMatrix(1, i) <> "" Then
160                   If g.TextMatrix(0, i) = "Microscopy Negative" Then
170                       s = s & g.TextMatrix(1, i)
180                   Else
190                       s = i & ". " & s & g.TextMatrix(1, i)
200                   End If
210               End If
220           End If
230           If Len(s) > 46 Then
240               LastWordIndex = InStrRev(Left(s, 46), " ")
250               Multiline = FormatString(Mid(s, LastWordIndex + 1, Len(s)), 46, , AlignLeft)
260               s = Left(s, LastWordIndex - 1)
270           End If
280           If Trim(s) <> "" Then
290               lpc = UBound(udtPL) + 1
300               ReDim Preserve udtPL(0 To lpc)
310               udtPL(lpc).LineType = "NORMAL10"
320               udtPL(lpc).LineText = FormatString(s, 46, , AlignLeft)
330               If Multiline <> "" Then
340                   lpc = UBound(udtPL) + 1
350                   ReDim Preserve udtPL(0 To lpc)
360                   udtPL(lpc).LineType = "NORMAL10"
370                   udtPL(lpc).LineText = FormatString("  " & Multiline, 46, , AlignLeft)
380               End If
390           End If
400           OrgNo = ""
410           Multiline = ""
420           s = ""
430       Next i

440       Exit Sub

GetPrintLineOrganismsOther_Error:

          Dim strES As String
          Dim intEL As Integer

450       intEL = Erl
460       strES = Err.Description
470       LogError "modNewMicro", "GetPrintLineOrganismsOther", intEL, strES


End Sub



Private Sub PrintResultLines()

          Dim i As Integer
          Dim lpc As Integer

10        On Error GoTo PrintResultLines_Error

          'Make both arrays same index.
20        If UBound(udtPL) > UBound(udtPR) Then
30            For i = UBound(udtPR) + 1 To UBound(udtPL)
40                ReDim Preserve udtPR(0 To i)
50                udtPR(i).LineText = ""
60            Next
70        ElseIf UBound(udtPR) > UBound(udtPL) Then
80            For i = UBound(udtPL) + 1 To UBound(udtPR)
90                ReDim Preserve udtPL(0 To i)
100               udtPL(i).LineText = FormatString("", 46)
110           Next
120       End If




130       For lpc = 0 To UBound(udtPL)

140           If udtPL(lpc).LineType = "BOLD10" Then
150               PrintTextRTB txtReport, FormatString(udtPL(lpc).LineText, 46) & "    ", 10, True
160           ElseIf udtPL(lpc).LineType = "TITLE10" Then
170               PrintTextRTB txtReport, FormatString(udtPL(lpc).LineText, 46) & "    ", 10, True, , True
180           ElseIf udtPL(lpc).LineType = "NORMAL10" Then
190               PrintTextRTB txtReport, FormatString(udtPL(lpc).LineText, 46) & "    ", 10, False
200           Else
210               PrintTextRTB txtReport, FormatString(udtPL(lpc).LineText, 46) & "    ", 10, False
220           End If

230           If udtPR(lpc).LineType = "BOLD10" Then
240               PrintTextRTB txtReport, FormatString(udtPR(lpc).LineText, 34) & vbCrLf, 10, True
250           ElseIf udtPL(lpc).LineType = "TITLE10" Then
260               PrintTextRTB txtReport, FormatString(udtPL(lpc).LineText, 46) & "    ", 10, True, , True
270           ElseIf udtPR(lpc).LineType = "NORMAL10" Then
280               PrintTextRTB txtReport, FormatString(udtPR(lpc).LineText, 34) & vbCrLf, 10, False
290           Else
300               PrintTextRTB txtReport, FormatString(udtPR(lpc).LineText, 34) & vbCrLf, 10, False
310           End If

320       Next lpc

330       Exit Sub

PrintResultLines_Error:

          Dim strES As String
          Dim intEL As Integer

340       intEL = Erl
350       strES = Err.Description
360       LogError "frmMicroReport", "PrintResultLines", intEL, strES

End Sub

Private Sub GetPrintLineComments(SampleIDWithOffset As Double, ByVal CommentTitle As String, _
                                 ByVal FieldName As String)

10        On Error GoTo GetPrintLineComments_Error

20        ReDim Comments(1 To 8) As String
          Dim sql As String
          Dim tb As Recordset
          Dim n As Integer
          Dim lpc As Integer
          Dim OB As Observation
          Dim OBS As New Observations

30        Set OBS = OBS.Load(SampleIDWithOffset, FieldName)

40        If Not OBS Is Nothing Then
50            For Each OB In OBS

60                FillCommentLines OB.Comment, 8, Comments(), 80
70                For n = 1 To 8
80                    If Trim(Comments(n) & "") <> "" Then
90                        PrintTextRTB txtReport, FormatString(Comments(n), 80, , AlignLeft) & vbCrLf, 10
100                   End If
110               Next
120           Next
130       End If

140       Exit Sub

GetPrintLineComments_Error:

          Dim strES As String
          Dim intEL As Integer

150       intEL = Erl
160       strES = Err.Description
170       LogError "modNewMicro", "GetPrintLineComments", intEL, strES, sql

End Sub

Public Sub FillCommentLines(ByVal FullComment As String, _
                            ByVal NumberOfLines As Integer, _
                            ByRef Comments() As String, _
                            Optional ByVal MaxLen As Integer = 80)

          Dim n As Integer
          Dim CurrentLine As Integer
          Dim X As Integer
          Dim ThisLine As String
          Dim SpaceFound As Boolean

10        On Error GoTo FillCommentLines_Error

20        For n = 1 To UBound(Comments)
30            Comments(n) = ""
40        Next

50        CurrentLine = 0
60        FullComment = Trim(FullComment)
70        n = Len(FullComment)

80        For X = n - 1 To 1 Step -1
90            If Mid(FullComment, X, 1) = vbCr Or Mid(FullComment, X, 1) = vbLf Or Mid(FullComment, X, 1) = vbTab Then
100               Mid(FullComment, X, 1) = " "
110           End If
120       Next

130       For X = n - 3 To 1 Step -1
140           If Mid(FullComment, X, 2) = "  " Then
150               FullComment = Left(FullComment, X) & Mid(FullComment, X + 2)
160           End If
170       Next
180       n = Len(FullComment)

190       Do While n > MaxLen
200           SpaceFound = False
210           For X = MaxLen To 1 Step -1
220               If Mid(FullComment, X, 1) = " " Then
230                   ThisLine = Left(FullComment, X - 1)
240                   FullComment = Mid(FullComment, X + 1)

250                   CurrentLine = CurrentLine + 1
260                   If CurrentLine <= NumberOfLines Then
270                       Comments(CurrentLine) = ThisLine
280                   End If
290                   SpaceFound = True
300                   Exit For
310               End If
320           Next
330           If Not SpaceFound Then
340               ThisLine = Left(FullComment, MaxLen)
350               FullComment = Mid(FullComment, MaxLen + 1)

360               CurrentLine = CurrentLine + 1
370               If CurrentLine <= NumberOfLines Then
380                   Comments(CurrentLine) = ThisLine
390               End If
400           End If
410           n = Len(FullComment)
420       Loop

430       CurrentLine = CurrentLine + 1
440       If CurrentLine <= NumberOfLines Then
450           Comments(CurrentLine) = FullComment
460       End If

470       Exit Sub

FillCommentLines_Error:

          Dim strES As String
          Dim intEL As Integer

480       intEL = Erl
490       strES = Err.Description
500       LogError "Other", "FillCommentLines", intEL, strES

End Sub

Private Sub PrintThis()

          Dim tb As Recordset
          Dim sql As String
          Dim SampleIDWithOffset As Double

10        On Error GoTo PrintThis_Error

20        PBar = 0

30        If grdSID.TextMatrix(grdSID.row, 0) = "" Then Exit Sub

40        SampleIDWithOffset = grdSID.TextMatrix(grdSID.row, 0) + SysOptMicroOffset(0)

50        sql = "Select * from PrintPending where " & _
                "Department = 'N' " & _
                "and SampleID = " & SampleIDWithOffset
60        Set tb = New Recordset
70        RecOpenClient 0, tb, sql
80        If tb.EOF Then
90            tb.AddNew
100       End If
110       tb!SampleID = SampleIDWithOffset
120       tb!Ward = PatWard
130       tb!Clinician = PatClinician
140       tb!GP = PatGP
150       tb!Department = "N"
160       tb!Initiator = UserName
170       tb!UsePrinter = ""   'pPrintToPrinter
180       tb!NoOfCopies = 1
190       tb!FinalInterim = "F"
200       tb!pTime = Now
210       tb.Update

220       Exit Sub

PrintThis_Error:

          Dim strES As String
          Dim intEL As Integer

230       intEL = Erl
240       strES = Err.Description
250       LogError "frmMicroReport", "PrintThis", intEL, strES, sql

End Sub



