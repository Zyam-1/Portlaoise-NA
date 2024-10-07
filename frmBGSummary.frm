VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmBGSummary 
   Caption         =   "NetAcquire  - Blood Gas Daily Summary"
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   12090
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
   Icon            =   "frmBGSummary.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8220
   ScaleWidth      =   12090
   StartUpPosition =   2  'CenterScreen
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
      Format          =   130547713
      CurrentDate     =   36965
   End
   Begin VB.CommandButton cmdValidate 
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
      Height          =   885
      Left            =   4740
      Picture         =   "frmBGSummary.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   45
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid grdBga 
      Height          =   7065
      Left            =   -30
      TabIndex        =   3
      Top             =   1035
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
      FormatString    =   $"frmBGSummary.frx":0614
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
      Height          =   885
      Left            =   6300
      Picture         =   "frmBGSummary.frx":06AF
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   45
      Width           =   1455
   End
   Begin VB.CommandButton cmdPrint 
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
      Height          =   885
      Left            =   7860
      Picture         =   "frmBGSummary.frx":09B9
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   45
      Width           =   1455
   End
   Begin VB.CommandButton cmdExit 
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
      Height          =   885
      Left            =   9420
      Picture         =   "frmBGSummary.frx":0CC3
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   45
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
Attribute VB_Name = "frmBGSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit  '© Custom Software 2001


Private Sub cmdExit_Click()

10        Unload Me

End Sub

Private Sub cmdPrint_Click()

          Dim n As Long
          Dim sortby As Long
          Dim OrigOrientation As Long

10        On Error GoTo cmdPrint_Click_Error

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
120       Printer.Print "St Johns Hospital Limerick"
130       Printer.Print "Blood Gas - Summary "; dtRunDate
140       Printer.Print "Sorted by ";
150       Printer.Print Choose(sortby + 1, "Lab Run Number", "Patient Name", "Clinician")
160       Printer.Print

170       Printer.Font.Size = 10
180       GoSub PrintHead
190       For n = 1 To grdBga.Rows - 1
200           If Printer.CurrentY > Printer.Height - 800 Then
210               Printer.NewPage
220               GoSub PrintHead
230           End If
240           Printer.Print grdBga.TextMatrix(n, 0);                    'lab
250           Printer.Print Tab(7); grdBga.TextMatrix(n, 1);            'nopas
260           Printer.Print Tab(15); grdBga.TextMatrix(n, 2);           'dob
270           Printer.Print Tab(25); Left(grdBga.TextMatrix(n, 3), 19);    'name
280           Printer.Print Tab(45); Left(grdBga.TextMatrix(n, 4) & Space(19), 19);    'clinician

290           Printer.Font.Size = 8
300           Printer.Print "  ";
310           Printer.Print Left(grdBga.TextMatrix(n, 5) & Space(6), 6);
320           Printer.Print Left(grdBga.TextMatrix(n, 6) & Space(6), 6);
330           Printer.Print Left(grdBga.TextMatrix(n, 7) & Space(6), 6);
340           Printer.Print Left(grdBga.TextMatrix(n, 8) & Space(6), 6);
350           Printer.Print Left(grdBga.TextMatrix(n, 9) & Space(6), 6);
360           Printer.Print Left(grdBga.TextMatrix(n, 10) & Space(6), 6);
370           Printer.Print Left(grdBga.TextMatrix(n, 11) & Space(6), 6);
380           Printer.Font.Size = 10
390           Printer.Print
400       Next

410       Printer.EndDoc

420       Printer.Orientation = OrigOrientation

430       Exit Sub

PrintHead:
440       Printer.Print grdBga.TextMatrix(0, 0);                    'lab
450       Printer.Print Tab(7); grdBga.TextMatrix(0, 1);            'nopas
460       Printer.Print Tab(15); grdBga.TextMatrix(0, 2);           'dob
470       Printer.Print Tab(25); Left(grdBga.TextMatrix(0, 3), 19);    'name
480       Printer.Print Tab(45); Left(grdBga.TextMatrix(0, 4) & Space(19), 19);    'clinician

490       Printer.Font.Size = 8
500       Printer.Print "  ";
510       Printer.Print Left(grdBga.TextMatrix(0, 5) & Space(6), 6);
520       Printer.Print Left(grdBga.TextMatrix(0, 6) & Space(6), 6);
530       Printer.Print Left(grdBga.TextMatrix(0, 7) & Space(6), 6);
540       Printer.Print Left(grdBga.TextMatrix(0, 8) & Space(6), 6);
550       Printer.Font.Size = 10
560       Printer.Print
570       Return

580       Exit Sub

cmdPrint_Click_Error:

          Dim strES As String
          Dim intEL As Integer

590       intEL = Erl
600       strES = Err.Description
610       LogError "frmBGSummary", "cmdPrint_Click", intEL, strES

End Sub

Private Sub cmdvalidate_Click()

          Dim n As Long
          Dim StartRow As Long
          Dim StopRow As Long
          Dim tb As New Recordset
          Dim sql As String

10        On Error GoTo cmdvalidate_Click_Error

20        If grdBga.Row > grdBga.RowSel Then
30            StartRow = grdBga.RowSel
40            StopRow = grdBga.Row
50        Else
60            StartRow = grdBga.Row
70            StopRow = grdBga.RowSel
80        End If

90        If StartRow = StopRow Then
100           grdBga.Col = 0
110           If iMsg("Validate Lab # " & grdBga & "?", vbQuestion + vbYesNo) = vbNo Then
120               Exit Sub
130           End If
140       Else
150           If iMsg("Validate all SELECTed rows?", vbQuestion + vbYesNo) = vbNo Then
160               Exit Sub
170           End If
180       End If

190       grdBga.Col = 0
200       For n = StartRow To StopRow
210           grdBga.Row = n
220           If grdBga <> "Lab #" Then
230               sql = "SELECT * from BGAresults WHERE rundate = '" & Format(dtRunDate, "dd/mmm/yyyy") & "' and " & _
                        "sampleid = '" & grdBga & "'"
240               Set tb = New Recordset
250               RecOpenServer 0, tb, sql
260               If Not tb.EOF Then
270                   tb!Valid = True
280                   tb.Update
290               End If
300           End If
310       Next

320       Exit Sub

cmdvalidate_Click_Error:

          Dim strES As String
          Dim intEL As Integer

330       intEL = Erl
340       strES = Err.Description
350       LogError "frmBGSummary", "cmdvalidate_Click", intEL, strES, sql

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
60        LogError "frmBGSummary", "dtRunDate_CloseUp", intEL, strES


End Sub

Private Sub FillG()

          Dim n As Long
          Dim sn As New Recordset
          Dim snr As Recordset
          Dim tb As New Recordset
          Dim sql As String
          Dim s As String

10        On Error GoTo FillG_Error

20        ClearFGrid grdBga

30        sql = "SELECT distinct sampleid from bgaresults WHERE " & _
                "rundate = '" & Format((dtRunDate), "yyyy/MM/dd") & "'"
40        Set sn = New Recordset
50        RecOpenServer 0, sn, sql

60        grdBga.Visible = False

70        Do While Not sn.EOF
80            sql = "SELECT * FROM Demographics WHERE " & _
                    "SampleID = '" & Trim$(sn!SampleID) & "' " & _
                    "AND RunDate = '" & Format((dtRunDate), "yyyy/MM/dd") & "'"
90            Set tb = New Recordset
100           RecOpenServer 0, tb, sql
110           sql = "SELECT * from bgaresults WHERE " & _
                    "rundate = '" & Format((dtRunDate), "yyyy/MM/dd") & "' and sampleid = '" & Trim(sn!SampleID) & "'"
120           Set snr = New Recordset
130           RecOpenServer 0, snr, sql
140           If Not snr.EOF Then
150               s = snr!SampleID & vbTab
160               If Not tb.EOF Then
170                   s = s & Format(tb!Dob, "dd/MM/yy") & vbTab & _
                          tb!PatName & vbTab & _
                          tb!Clinician & vbTab
180               Else
190                   s = s & vbTab & vbTab & vbTab
200               End If
210               sql = "SELECT * from bgaresults WHERE " & _
                        "rundate = '" & Format((dtRunDate), "yyyy/MM/dd") & "' and sampleid = '" & Trim(sn!SampleID) & "'"
220               Set snr = New Recordset
230               RecOpenServer 0, snr, sql
240               Do While Not snr.EOF
250                   s = s & snr!pH & "" & vbTab & snr!PCO2 & "" & vbTab
260                   s = s & snr!PO2 & "" & vbTab & snr!HCO3 & "" & vbTab
270                   s = s & snr!BE & "" & vbTab & snr!O2SAT & "" & vbTab
280                   s = s & snr!O2SAT & "" & vbTab & snr!TotCO2 & ""
290                   snr.MoveNext
300               Loop
310               grdBga.AddItem s
320           End If
330           sn.MoveNext
340       Loop

350       For n = 0 To 2
360           If oSort(n) Then
370               grdBga.Col = Choose(n + 1, 0, 3, 4)
380               Exit For
390           End If
400       Next

410       grdBga.Sort = flexSortGenericAscending

420       FixG grdBga

430       Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer

440       intEL = Erl
450       strES = Err.Description
460       LogError "frmBGSummary", "FillG", intEL, strES, sql

End Sub

Private Sub Form_Load()

10        On Error GoTo Form_Load_Error

20        dtRunDate = Format(Now, "dd/mmm/yyyy")

30        Set_Font Me

40        Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmBGSummary", "Form_Load", intEL, strES


End Sub

Private Sub osort_Click(Index As Integer)

          Dim n As Long

10        On Error GoTo osort_Click_Error

20        For n = 0 To 2
30            If oSort(n) Then
40                grdBga.Col = Choose(n + 1, 0, 3, 4)
50                Exit For
60            End If
70        Next

80        grdBga.Sort = flexSortGenericAscending

90        Exit Sub

osort_Click_Error:

          Dim strES As String
          Dim intEL As Integer

100       intEL = Erl
110       strES = Err.Description
120       LogError "frmBGSummary", "osort_Click", intEL, strES

End Sub
