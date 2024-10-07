VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmExtStats 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - External Statistics"
   ClientHeight    =   9240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14790
   Icon            =   "frmExtStats.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9240
   ScaleWidth      =   14790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExcelTot 
      Caption         =   "Export to Excel"
      Height          =   1110
      Left            =   8220
      Picture         =   "frmExtStats.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5160
      Width           =   735
   End
   Begin MSFlexGridLib.MSFlexGrid grdTot 
      Height          =   7890
      Left            =   8970
      TabIndex        =   13
      Top             =   1155
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   13917
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      ScrollTrack     =   -1  'True
      FormatString    =   "<Analyte                                     |<Total       "
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
   Begin VB.CommandButton cmdExcelRes 
      Caption         =   "Export to Excel"
      Height          =   1110
      Left            =   7080
      Picture         =   "frmExtStats.frx":5F28
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5160
      Width           =   735
   End
   Begin VB.Frame Frame2 
      Caption         =   "Source"
      Height          =   915
      Left            =   4080
      TabIndex        =   6
      Top             =   90
      Width           =   1245
      Begin VB.OptionButton optGP 
         Caption         =   "GP"
         Height          =   195
         Left            =   180
         TabIndex        =   9
         Top             =   660
         Width           =   945
      End
      Begin VB.OptionButton optWard 
         Caption         =   "Ward"
         Height          =   195
         Left            =   180
         TabIndex        =   8
         Top             =   450
         Width           =   975
      End
      Begin VB.OptionButton optClinician 
         Caption         =   "Clinician"
         Height          =   195
         Left            =   180
         TabIndex        =   7
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Between Dates"
      Height          =   915
      Left            =   240
      TabIndex        =   3
      Top             =   90
      Width           =   3195
      Begin MSComCtl2.DTPicker calTo 
         Height          =   375
         Left            =   1620
         TabIndex        =   4
         Top             =   360
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   661
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
         Format          =   59179009
         CurrentDate     =   37951
      End
      Begin MSComCtl2.DTPicker calFrom 
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   661
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
         Format          =   59179009
         CurrentDate     =   37951
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Exit"
      Height          =   1110
      Left            =   13440
      Picture         =   "frmExtStats.frx":BB46
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7920
      Width           =   1035
   End
   Begin MSFlexGridLib.MSFlexGrid grdRes 
      Height          =   7890
      Left            =   225
      TabIndex        =   1
      Top             =   1155
      Width           =   6810
      _ExtentX        =   12012
      _ExtentY        =   13917
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      AllowUserResizing=   1
      FormatString    =   "<Source                                  |<Test                                                               |<Total      "
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
      Left            =   5970
      Picture         =   "frmExtStats.frx":D4CC
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   180
      Width           =   1035
   End
   Begin VB.Label lblTotals 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total from all Clinicians"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   285
      Left            =   8970
      TabIndex        =   15
      Top             =   870
      Width           =   4335
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
      Left            =   7050
      TabIndex        =   12
      Top             =   6330
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Label lblCalculating 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Calculating..."
      Height          =   285
      Left            =   7020
      TabIndex        =   11
      Top             =   450
      Visible         =   0   'False
      Width           =   1035
   End
End
Attribute VB_Name = "frmExtStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub FillTotals()

          Dim sql As String
          Dim tb As Recordset
          Dim FromDate As String
          Dim ToDate As String
          Dim s As String
          Dim Source As String
          Dim NoGP As String

10        On Error GoTo FillTotals_Error

20        Screen.MousePointer = vbHourglass

30        FromDate = Format(calFrom, "dd/MMM/yyyy")
40        ToDate = Format(calTo, "dd/MMM/yyyy")

50        lblCalculating.Visible = True
60        lblCalculating.Refresh

70        If optClinician Then
80            Source = "Clinician"
90            NoGP = ""
100       ElseIf optWard Then
110           Source = "Ward"
120           NoGP = "AND Ward <> 'GP' "
130       Else
140           Source = "GP"
150           NoGP = ""
160       End If

170       sql = "SELECT CAST(T.AnalyteName AS nvarchar(50)) + CHAR(9) + CAST(COUNT(Analyte) AS nvarchar(50)) s " & _
                "FROM ExtResults R LEFT JOIN ExternalDefinitions T " & _
                "ON T.AnalyteName = R.Analyte " & _
                "LEFT JOIN Demographics D " & _
                "ON R.SampleID = D.SampleID " & _
                "WHERE SentDate BETWEEN '" & FromDate & "' AND '" & ToDate & "' " & _
                "AND COALESCE(D." & Source & ", '') <> '' " & _
                NoGP & _
                "GROUP BY T.AnalyteName " & _
                "ORDER BY T.AnalyteName"

180       Set tb = New Recordset
190       RecOpenServer 0, tb, sql
200       Do While Not tb.EOF
210           grdTot.AddItem tb!s & ""
220           tb.MoveNext
230       Loop

240       If grdTot.Rows > 2 Then
250           grdTot.RemoveItem 1
260       End If
270       grdTot.Visible = True
280       lblCalculating.Visible = False
290       Screen.MousePointer = vbNormal

300       Exit Sub

FillTotals_Error:

          Dim strES As String
          Dim intEL As Integer

310       intEL = Erl
320       strES = Err.Description
330       LogError "frmExtStats", "FillTotals", intEL, strES
340       grdTot.Visible = True
350       lblCalculating.Visible = False
360       Screen.MousePointer = vbNormal

End Sub

Private Sub cmdCancel_Click()

10        Unload Me

End Sub

Private Sub cmdExcelTot_Click()

          Dim s As String

10        s = "External Statistics" & vbCr & _
              "Total Requests from all "
20        If optClinician Then
30            s = s & "Clinicians"
40        ElseIf optWard Then
50            s = s & "Wards"
60        Else
70            s = s & "GP's"
80        End If
90        s = s & vbCr & _
              "Between " & calFrom & " and " & calTo & vbCr

100       ExportFlexGrid grdTot, Me, s

End Sub

Private Sub cmdPrint_Click()
          Dim Num As Long

10        On Error GoTo cmdPrint_Click_Error

20        For Num = 0 To grdRes.Rows - 1
30            Printer.Print grdRes.TextMatrix(Num, 0);
40            Printer.Print Tab(35); grdRes.TextMatrix(Num, 1);
50            Printer.Print Tab(80); grdRes.TextMatrix(Num, 2)
60        Next

70        Printer.EndDoc

80        Exit Sub

cmdPrint_Click_Error:

          Dim strES As String
          Dim intEL As Integer

90        intEL = Erl
100       strES = Err.Description
110       LogError "frmExtStats", "cmdPrint_Click", intEL, strES

End Sub

Private Sub cmdExcelRes_Click()

          Dim s As String

10        s = "External Statistics" & vbCr & _
              "List of Requests from "
20        If optClinician Then
30            s = s & "Clinicians"
40        ElseIf optWard Then
50            s = s & "Wards"
60        Else
70            s = s & "GP's"
80        End If
90        s = s & vbCr & _
              "Between " & calFrom & " and " & calTo & vbCr


100       ExportFlexGrid grdRes, Me, s

End Sub

Private Sub cmdStart_Click()

          Dim sql As String
          Dim tb As Recordset
          Dim FromDate As String
          Dim ToDate As String
          Dim s As String
          Dim Source As String

10        On Error GoTo cmdStart_Click_Error

20        Screen.MousePointer = vbHourglass

30        FromDate = Format(calFrom, "dd/MMM/yyyy")
40        ToDate = Format(calTo, "dd/MMM/yyyy")

50        lblCalculating.Visible = True
60        lblCalculating.Refresh

70        With grdRes
80            .Rows = 2
90            .AddItem ""
100           .RemoveItem 1
110           .Refresh
120           .Visible = False
130       End With

140       With grdTot
150           .Rows = 2
160           .AddItem ""
170           .RemoveItem 1
180           .Refresh
190           .Visible = False
200       End With

210       If optClinician Then
220           Source = "Clinician"
230           lblTotals = "Total from all Clinicians"
240       ElseIf optWard Then
250           Source = "Ward"
260           lblTotals = "Total from all Wards"
270       Else
280           Source = "GP"
290           lblTotals = "Total from all GPs"
300       End If

310       sql = "SELECT D." & Source & " + CHAR(9) + CAST(T.AnalyteName COLLATE DATABASE_DEFAULT AS nvarchar(50)) + " & _
                "CHAR(9) + CAST(COUNT(Analyte COLLATE DATABASE_DEFAULT) AS nvarchar(50)) s " & _
                "FROM ExtResults R LEFT JOIN ExternalDefinitions T " & _
                "ON T.AnalyteName = R.Analyte " & _
                "LEFT JOIN Demographics D " & _
                "ON R.SampleID = D.SampleID " & _
                "WHERE SentDate BETWEEN '" & FromDate & "' AND '" & ToDate & "' " & _
                "AND COALESCE(D." & Source & ", '') <> '' " & _
                "GROUP BY T.AnalyteName, D." & Source & " " & " ORDER BY D." & Source

320       Set tb = New Recordset
330       RecOpenServer 0, tb, sql
340       Do While Not tb.EOF
350           grdRes.AddItem tb!s & ""
360           tb.MoveNext
370       Loop

380       If grdRes.Rows > 2 And grdRes.TextMatrix(1, 0) = "" Then
390           grdRes.RemoveItem 1
400       End If
410       grdRes.Visible = True
420       lblCalculating.Visible = False
430       Screen.MousePointer = vbNormal

440       FillTotals

450       Exit Sub

cmdStart_Click_Error:

          Dim strES As String
          Dim intEL As Integer

460       intEL = Erl
470       strES = Err.Description
480       LogError "frmExtStats", "cmdStart_Click", intEL, strES
490       grdRes.Visible = True
500       lblCalculating.Visible = False
510       Screen.MousePointer = vbNormal

End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()

10        On Error GoTo Form_Load_Error

20        calFrom = Format(Now - 7, "dd/MMM/yyyy")
30        calTo = Format(Now, "dd/MMM/yyyy")

40        Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer



50        intEL = Erl
60        strES = Err.Description
70        LogError "frmExtStats", "Form_Load", intEL, strES


End Sub

Private Sub grdRes_Click()

          Static SortOrder As Boolean

10        If grdRes.MouseRow = 0 Then
20            If SortOrder Then
30                grdRes.Sort = flexSortGenericAscending
40            Else
50                grdRes.Sort = flexSortGenericDescending
60            End If
70            SortOrder = Not SortOrder
80        End If

End Sub

