VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmGPClinWard 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - GP/Clinician/Ward Totals"
   ClientHeight    =   7695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12600
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   12600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExcel 
      Caption         =   "Export to Excel"
      Height          =   1095
      Index           =   1
      Left            =   8280
      Picture         =   "frmGPClinWard.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1440
      Width           =   765
   End
   Begin MSFlexGridLib.MSFlexGrid gOPD 
      Height          =   6405
      Left            =   9060
      TabIndex        =   9
      Top             =   1110
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   11298
      _Version        =   393216
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      ScrollTrack     =   -1  'True
      FormatString    =   "<Analyte                         |<Total                 "
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   3870
      Picture         =   "frmGPClinWard.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   150
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Caption         =   "Between Dates"
      Height          =   885
      Left            =   150
      TabIndex        =   4
      Top             =   60
      Width           =   3645
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   330
         Left            =   1830
         TabIndex        =   5
         Top             =   330
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MMM/yyyy"
         Format          =   59768835
         CurrentDate     =   38985
      End
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   330
         Left            =   150
         TabIndex        =   6
         Top             =   330
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MMM/yyyy"
         Format          =   59768835
         CurrentDate     =   38985
      End
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   6405
      Left            =   150
      TabIndex        =   3
      Top             =   1110
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   11298
      _Version        =   393216
      Cols            =   5
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      ScrollTrack     =   -1  'True
      FormatString    =   "<Analyte                         |<GP             |<Clinician    |<Ward         |<Total         "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   1095
      Left            =   7605
      Picture         =   "frmGPClinWard.frx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6420
      Width           =   975
   End
   Begin VB.CommandButton cmdExcel 
      Caption         =   "Export to Excel"
      Height          =   1095
      Index           =   0
      Left            =   7140
      Picture         =   "frmGPClinWard.frx":1A9E
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1440
      Width           =   765
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Clinicians - OPD only"
      ForeColor       =   &H80000018&
      Height          =   285
      Left            =   9090
      TabIndex        =   10
      Top             =   840
      Width           =   3270
   End
   Begin VB.Label lblCalculating 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Calculating..."
      Height          =   285
      Left            =   4950
      TabIndex        =   8
      Top             =   420
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label lblExcelInfo 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exporting..."
      Height          =   345
      Left            =   7380
      TabIndex        =   2
      Top             =   2580
      Visible         =   0   'False
      Width           =   1305
   End
End
Attribute VB_Name = "frmGPClinWard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub FillG()

          Dim tb As Recordset
          Dim tbA As Recordset
          Dim sql As String
          Dim FromDate As String
          Dim ToDate As String
          Dim s As String
          Dim n As Integer

10        On Error GoTo FillG_Error

20        FromDate = Format$(dtFrom, "dd/MMM/yyyy")
30        ToDate = Format$(dtTo, "dd/MMM/yyyy") & " 23:59"

40        g.Rows = 2
50        g.AddItem ""
60        g.RemoveItem 1

70        n = 0

          'sql = "SELECT DISTINCT(T.Code), T.Shortname FROM BioResults R JOIN BioTestDefinitions T " & _
           '      "ON R.Code = T.Code " & _
           '      "JOIN Demographics D " & _
           '      "ON R.SampleID = D.SampleID " & _
           '      "WHERE D.RunDate BETWEEN '" & FromDate & "' AND '" & ToDate & "'"
          'Set tb = New Recordset
          'RecOpenServer 0, tb, sql
          'Do While Not tb.EOF
          '
          '  sql = "SELECT " & _
             '        "TotGP = (SELECT Count(GP) FROM BioResults R JOIN Demographics D " & _
             '        "     ON R.SampleID = D.SampleID " & _
             '        "     WHERE Code = '" & tb!Code & "' " & _
             '        "     AND D.RunDate BETWEEN '" & FromDate & "' AND '" & ToDate & "' " & _
             '        "     AND COALESCE(GP, '') <> ''), " & _
             '        "TotClin = (SELECT Count(Clinician) FROM BioResults R JOIN Demographics D " & _
             '        "     ON R.SampleID = D.SampleID " & _
             '        "     WHERE Code = '" & tb!Code & "' " & _
             '        "     AND D.RunDate BETWEEN '" & FromDate & "' AND '" & ToDate & "' " & _
             '        "     AND COALESCE(Clinician, '') <> ''), " & _
             '        "TotWard = (SELECT Count(Ward) FROM BioResults R JOIN Demographics D " & _
             '        "     ON R.SampleID = D.SampleID " & _
             '        "     WHERE Code = '" & tb!Code & "' " & _
             '        "     AND D.RunDate BETWEEN '" & FromDate & "' AND '" & ToDate & "' " & _
             '        "     AND Ward <> 'GP')"
          ''              , " & _
           ''              "TotAll = (SELECT Count(*) FROM BioResults R JOIN Demographics D " & _
           ''              "     ON R.SampleID = D.SampleID " & _
           ''              "     WHERE Code = '" & tb!Code & "' " & _
           ''              "     AND D.RunDate BETWEEN '" & FromDate & "' AND '" & ToDate & "')"
          '  Set tbA = New Recordset
          '  RecOpenServer 0, tbA, sql
          '  If Not tbA.EOF Then


80        sql = "SELECT DISTINCT(T.Code), T.Shortname, " & _
                "TotGP = (SELECT Count(GP) FROM BioResults R JOIN Demographics D ON R.SampleID = D.SampleID " & _
                "WHERE Code = T.Code AND D.RunDate BETWEEN '" & FromDate & "' AND '" & ToDate & "' " & _
                "AND COALESCE(GP, '') <> ''), " & _
                "TotClin = (SELECT Count(Clinician) FROM BioResults R JOIN Demographics D ON R.SampleID = D.SampleID " & _
                "WHERE Code =  T.Code AND D.RunDate BETWEEN '" & FromDate & "' AND '" & ToDate & "' " & _
                "AND COALESCE(Clinician, '') <> ''), " & _
                "TotWard = (SELECT Count(Ward) FROM BioResults R JOIN Demographics D ON R.SampleID = D.SampleID " & _
                "WHERE Code =  T.Code AND D.RunDate BETWEEN '" & FromDate & "' AND '" & ToDate & "' " & _
                "AND Ward <> 'GP'), " & _
                "TotOPD = (SELECT Count(*) FROM BioResults R JOIN Demographics D ON R.SampleID = D.SampleID " & _
                "WHERE Code = T.Code AND D.RunDate BETWEEN '01/Apr/2011' AND '15/Apr/2011 23:59' " & _
                "AND Ward LIKE '%OPD%') " & _
                "FROM BioResults R JOIN BioTestDefinitions T ON R.Code = T.Code " & _
                "JOIN Demographics D ON R.SampleID = D.SampleID " & _
                "WHERE D.RunDate BETWEEN '" & FromDate & "' AND '" & ToDate & "' "
90        Set tb = New Recordset
100       RecOpenClient 0, tb, sql
110       g.Visible = False
120       gOPD.Visible = False

130       While Not tb.EOF
140           s = tb!ShortName & "" & vbTab & _
                  tb!TotGP & vbTab & _
                  tb!TotClin & vbTab & _
                  tb!TotWard & vbTab & _
                  "Tot" & ""
150           g.AddItem s
160           gOPD.AddItem tb!ShortName & "" & vbTab & _
                           tb!TotOPD

170           tb.MoveNext
180       Wend
190       g.Visible = True
200       gOPD.Visible = True
          '  n = n + 1
          '  If n > 5 Then
          '    n = 0
          '    g.Refresh
          '  End If
          '  tb.MoveNext
          'Loop

210       If g.Rows > 2 Then
220           CalculateTotal
230           g.RemoveItem 1
240       End If
250       If gOPD.Rows > 2 Then gOPD.RemoveItem 1

260       Exit Sub

FillG_Error:

270       g.Visible = True
280       gOPD.Visible = True

          Dim strES As String
          Dim intEL As Integer

290       intEL = Erl
300       strES = Err.Description
310       LogError "frmGPClinWard", "FillG", intEL, strES, sql


End Sub

Private Sub FillGOPD()

          Dim tb As Recordset
          Dim tbA As Recordset
          Dim sql As String
          Dim FromDate As String
          Dim ToDate As String
          Dim s As String
          Dim n As Integer

10        FromDate = Format$(dtFrom, "dd/MMM/yyyy")
20        ToDate = Format$(dtTo, "dd/MMM/yyyy") & " 23:59"

30        gOPD.Rows = 2
40        gOPD.AddItem ""
50        gOPD.RemoveItem 1

60        n = 0

70        sql = "SELECT DISTINCT(T.Code), T.Shortname FROM BioResults R JOIN BioTestDefinitions T " & _
                "ON R.Code = T.Code " & _
                "JOIN Demographics D " & _
                "ON R.SampleID = D.SampleID " & _
                "WHERE D.RunDate BETWEEN '" & FromDate & "' AND '" & ToDate & "'"
80        Set tb = New Recordset
90        RecOpenServer 0, tb, sql
100       Do While Not tb.EOF

110           sql = "SELECT Count(*) Tot FROM BioResults R JOIN Demographics D " & _
                    "ON R.SampleID = D.SampleID " & _
                    "WHERE Code = '" & tb!Code & "' " & _
                    "AND D.RunDate BETWEEN '" & FromDate & "' AND '" & ToDate & "' " & _
                    "AND Ward LIKE '%OPD%'"
120           Set tbA = New Recordset
130           RecOpenServer 0, tbA, sql
140           If Not tbA.EOF Then
150               s = tb!ShortName & vbTab & _
                      tbA!Tot & ""
160               gOPD.AddItem s
170           End If
180           n = n + 1
190           If n > 5 Then
200               n = 0
210               gOPD.Refresh
220           End If
230           tb.MoveNext
240       Loop

250       If gOPD.Rows > 2 Then
260           gOPD.RemoveItem 1
270       End If

End Sub

Private Sub cmdCancel_Click()

10        Unload Me

End Sub


Private Sub cmdExcel_Click(Index As Integer)

10        If Index = 0 Then
20            ExportFlexGrid g, Me
30        Else
40            ExportFlexGrid gOPD, Me
50        End If

End Sub

Private Sub cmdStart_Click()

10        lblCalculating.Visible = True
20        lblCalculating.Refresh
30        FillG
          'FillGOPD
40        lblCalculating.Visible = False

End Sub

Private Sub Form_Load()

10        dtFrom = Format$(Now - 30, "dd/MMM/yyyy")
20        dtTo = Format$(Now, "dd/MMM/yyyy")

End Sub


Private Sub CalculateTotal()

          Dim i As Integer

10        On Error GoTo CalculateTotal_Error

20        For i = 1 To g.Rows - 1
30            g.TextMatrix(i, 4) = Val(g.TextMatrix(i, 1)) + Val(g.TextMatrix(i, 2)) + Val(g.TextMatrix(i, 3))
40        Next i

50        Exit Sub

CalculateTotal_Error:

          Dim strES As String
          Dim intEL As Integer

60        intEL = Erl
70        strES = Err.Description
80        LogError "frmGPClinWard", "CalculateTotal", intEL, strES

End Sub
