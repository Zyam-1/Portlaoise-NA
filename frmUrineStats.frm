VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frmUrineStats 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Urine Statistics"
   ClientHeight    =   8940
   ClientLeft      =   675
   ClientTop       =   1485
   ClientWidth     =   8040
   Icon            =   "frmUrineStats.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8940
   ScaleWidth      =   8040
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove Blanks"
      Height          =   735
      Left            =   6420
      Picture         =   "frmUrineStats.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   4860
      Width           =   1275
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   255
      Left            =   90
      TabIndex        =   23
      Top             =   1560
      Width           =   5985
      _ExtentX        =   10557
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdPrint 
      Appearance      =   0  'Flat
      Caption         =   "&Print"
      Height          =   735
      Left            =   6420
      Picture         =   "frmUrineStats.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   7080
      Width           =   1290
   End
   Begin VB.CommandButton cmdXL 
      Caption         =   "&Export to Excel"
      Height          =   735
      Left            =   6420
      Picture         =   "frmUrineStats.frx":0A56
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   5880
      Width           =   1275
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   735
      Left            =   6420
      Picture         =   "frmUrineStats.frx":0D60
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   8070
      Width           =   1290
   End
   Begin VB.Frame Frame1 
      Height          =   480
      Left            =   90
      TabIndex        =   12
      Top             =   1695
      Width           =   6000
      Begin Threed.SSOption o 
         Height          =   225
         Index           =   0
         Left            =   3870
         TabIndex        =   15
         Top             =   180
         Width           =   915
         _Version        =   65536
         _ExtentX        =   1614
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Clinicians"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   -1  'True
      End
      Begin Threed.SSOption o 
         Height          =   225
         Index           =   1
         Left            =   1890
         TabIndex        =   14
         Top             =   180
         Width           =   795
         _Version        =   65536
         _ExtentX        =   1402
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Wards"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption o 
         Height          =   225
         Index           =   2
         Left            =   150
         TabIndex        =   13
         Top             =   180
         Width           =   825
         _Version        =   65536
         _ExtentX        =   1455
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "G. P. s"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   6645
      Left            =   75
      TabIndex        =   11
      Top             =   2175
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   11721
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      FormatString    =   "<Source                                                          |<C & S       |<Preg        |<Red Sub  "
   End
   Begin VB.Frame Frame2 
      Caption         =   "Between Dates"
      Height          =   1530
      Left            =   90
      TabIndex        =   0
      Top             =   45
      Width           =   6000
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   330
         Left            =   1170
         TabIndex        =   17
         Top             =   360
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   582
         _Version        =   393216
         Format          =   59506689
         CurrentDate     =   37643
      End
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   330
         Left            =   2880
         TabIndex        =   16
         Top             =   360
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   582
         _Version        =   393216
         Format          =   59506689
         CurrentDate     =   37643
      End
      Begin VB.CommandButton bReCalc 
         Caption         =   "&Start"
         Height          =   795
         Left            =   4500
         MaskColor       =   &H8000000F&
         Picture         =   "frmUrineStats.frx":106A
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   180
         Width           =   1155
      End
      Begin Threed.SSOption obetween 
         Height          =   210
         Index           =   6
         Left            =   465
         TabIndex        =   2
         Top             =   945
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   370
         _StockProps     =   78
         Caption         =   "Today"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Value           =   -1  'True
      End
      Begin Threed.SSOption obetween 
         Height          =   210
         Index           =   3
         Left            =   3030
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1185
         Width           =   1155
         _Version        =   65536
         _ExtentX        =   2037
         _ExtentY        =   370
         _StockProps     =   78
         Caption         =   "Last Quarter"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin Threed.SSOption obetween 
         Height          =   210
         Index           =   4
         Left            =   2730
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   945
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   370
         _StockProps     =   78
         Caption         =   "Last Full Quarter"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin Threed.SSOption obetween 
         Height          =   210
         Index           =   2
         Left            =   1245
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1185
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   370
         _StockProps     =   78
         Caption         =   "Last Full Month"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption obetween 
         Height          =   210
         Index           =   5
         Left            =   4245
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1185
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   370
         _StockProps     =   78
         Caption         =   "Year to Date"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption obetween 
         Height          =   210
         Index           =   1
         Left            =   135
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1185
         Width           =   1065
         _Version        =   65536
         _ExtentX        =   1879
         _ExtentY        =   370
         _StockProps     =   78
         Caption         =   "Last Month"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin Threed.SSOption obetween 
         Height          =   210
         Index           =   0
         Left            =   1245
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   945
         Width           =   1065
         _Version        =   65536
         _ExtentX        =   1879
         _ExtentY        =   370
         _StockProps     =   78
         Caption         =   "Last Week"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label3 
         Caption         =   "From"
         Height          =   240
         Left            =   765
         TabIndex        =   22
         Top             =   405
         Width           =   555
      End
      Begin VB.Label Label2 
         Caption         =   "To"
         Height          =   240
         Left            =   2520
         TabIndex        =   21
         Top             =   405
         Width           =   420
      End
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
      Left            =   6420
      TabIndex        =   25
      Top             =   6600
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label lUrineSamples 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "123"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6180
      TabIndex        =   10
      Top             =   840
      Width           =   1725
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total Number of Urine Samples between the specified dates"
      Height          =   705
      Left            =   6180
      TabIndex        =   9
      Top             =   120
      Width           =   1725
   End
End
Attribute VB_Name = "frmUrineStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub FillG()

          Dim tb As Recordset
          Dim sql As String
          Dim Source As String
          Dim n As Integer
          Dim s As String
          Dim X As Integer
          Dim Y As Integer
          Dim Tot(1 To 7) As Integer

10        With g
20            .Visible = False
30            .Rows = 2
40            .AddItem ""
50            .RemoveItem 1
60        End With

70        If o(0) Then
80            Source = "Clinician"
90        ElseIf o(1) Then
100           Source = "Ward"
110       Else
120           Source = "GP"
130       End If

140       sql = "SELECT DISTINCT " & Source & " AS Source FROM Demographics WHERE " & _
                "RunDate BETWEEN '" & _
                Format(dtFrom, "dd/mmm/yyyy") & " 00:00:00' AND '" & _
                Format(dtTo, "dd/mmm/yyyy") & " 23:59:59' " & _
                "AND " & Source & " IS NOT NULL " & _
                "AND " & Source & " <> ''"
150       Set tb = New Recordset
160       RecOpenServer 0, tb, sql
170       Do While Not tb.EOF
180           g.AddItem tb!Source & ""
190           tb.MoveNext
200       Loop

210       If g.Rows > 2 Then
220           g.RemoveItem 1
230       End If

240       pb.Max = g.Rows
250       pb.Visible = True

260       For n = 1 To g.Rows - 1
270           pb = n
280           sql = "SELECT COUNT(DISTINCT(S.SampleID)) AS CS " & _
                    "FROM Sensitivities AS S, MicroSiteDetails M, Demographics AS D WHERE " & _
                    "S.Rundate BETWEEN '" & _
                    Format(dtFrom, "dd/MMM/yyyy") & " 00:00:00' AND '" & _
                    Format(dtTo, "dd/MMM/yyyy") & " 23:59:59' " & _
                    "AND D." & Source & " = '" & AddTicks(g.TextMatrix(n, 0)) & "' " & _
                    "AND D.SampleID = S.SampleID " & _
                    "AND M.Site LIKE 'Urine' " & _
                    "AND M.SampleID = S.SampleID " & _
                    "AND D.SampleID = S.SampleID"
290           Set tb = New Recordset
300           RecOpenServer 0, tb, sql
310           g.TextMatrix(n, 1) = tb!cS
320       Next

330       For n = 1 To g.Rows - 1
340           pb = n
350           sql = "SELECT COUNT(Pregnancy) AS Pregnancy " & _
                    "FROM Urine U, Demographics D " & _
                    "WHERE D.RunDate BETWEEN '" & _
                    Format(dtFrom, "dd/mmm/yyyy") & " 00:00:00' AND '" & _
                    Format(dtTo, "dd/mmm/yyyy") & " 23:59:59' " & _
                    "AND D." & Source & " = '" & AddTicks(g.TextMatrix(n, 0)) & "' " & _
                    "AND D.SampleID = U.SampleID " & _
                    "AND U.Pregnancy IS NOT NULL " & _
                    "AND U.Pregnancy <> ''"
360           Set tb = New Recordset
370           RecOpenServer 0, tb, sql
380           g.TextMatrix(n, 2) = tb!Pregnancy
390       Next

400       For n = 1 To g.Rows - 1
410           pb = n
420           sql = "SELECT COUNT(*) AS RedSub " & _
                    "FROM GenericResults G, Demographics D " & _
                    "WHERE D.RunDate BETWEEN '" & _
                    Format(dtFrom, "dd/mmm/yyyy") & " 00:00:00' and '" & _
                    Format(dtTo, "dd/mmm/yyyy") & " 23:59:59' " & _
                    "AND D." & Source & " = '" & AddTicks(g.TextMatrix(n, 0)) & "' " & _
                    "AND D.SampleID = G.SampleID " & _
                    "AND G.TestName = 'RedSub'"
430           Set tb = New Recordset
440           RecOpenServer 0, tb, sql
450           g.TextMatrix(n, 3) = tb!RedSub
460       Next

470       For X = 1 To 3
480           Tot(X) = 0
490       Next
500       For Y = 1 To g.Rows - 1
510           g.Row = Y
520           For X = 1 To 3
530               g.Col = X
540               Tot(X) = Tot(X) + Val(g)
550           Next
560       Next

570       s = "Total of Below"
580       For X = 1 To 3
590           s = s & vbTab & Format(Tot(X))
600       Next
610       g.AddItem s, 1

620       g.Visible = True
630       pb.Visible = False

640       Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer

650       intEL = Erl
660       strES = Err.Description
670       LogError "frmUrineStats", "FillG", intEL, strES, sql

680       g.Visible = True

End Sub

Private Sub ReCalc()

          Dim sn As Recordset
          Dim sql As String

10        On Error GoTo ReCalc_Error

20        sql = "SELECT COUNT(*) AS Tot FROM MicroSiteDetails M, Demographics D WHERE " & _
                "D.RunDate BETWEEN '" & _
                Format(dtFrom, "dd/mmm/yyyy") & " 00:00:00' AND '" & _
                Format(dtTo, "dd/mmm/yyyy") & " 23:59:59' " & _
                "AND D.SampleID = M.SampleID " & _
                "AND M.Site Like 'Urine'"

30        Set sn = New Recordset
40        RecOpenServer 0, sn, sql

50        lUrineSamples = sn!Tot

60        FillG

70        Exit Sub

ReCalc_Error:

          Dim strES As String
          Dim intEL As Integer

80        intEL = Erl
90        strES = Err.Description
100       LogError "frmUrineStats", "ReCalc", intEL, strES, sql

End Sub



Private Sub brecalc_Click()

10        ReCalc

End Sub

Private Sub cmdCancel_Click()

10        Unload Me

End Sub

Private Sub cmdPrint_Click()

          Dim Y As Integer
          Dim X As Integer

10        On Error GoTo cmdPrint_Click_Error

20        cmdPrint.Caption = "Printing"

30        With Printer
40            .Font.Name = "Courier New"
50            .Font.Bold = True
60            Printer.Print "Urine Statistics."
70            Printer.Print Format(dtFrom, "dd/mmmm/yyyy"); " to "; Format(dtTo, "dd/mmmm/yyyy")
80            Printer.Print
90            .Font.Bold = False
100       End With

110       For Y = 0 To g.Rows - 1
120           g.Row = Y
130           g.Col = 0
140           Printer.Print Left(g & Space(29), 29);
150           For X = 1 To g.Cols - 1
160               g.Col = X
170               Printer.Print Tab(Choose(X, 30, 38, 46, 54, 62, 70, 78)); g;
180           Next
190           Printer.Print
200       Next

210       Printer.EndDoc

220       cmdPrint.Caption = "&Print"

230       Exit Sub

cmdPrint_Click_Error:

          Dim strES As String
          Dim intEL As Integer

240       intEL = Erl
250       strES = Err.Description
260       LogError "frmUrineStats", "cmdPrint_Click", intEL, strES

End Sub

Private Sub cmdRemove_Click()

          Dim n As Integer

10        For n = g.Rows - 1 To 1 Step -1
20            If g.TextMatrix(n, 1) = "0" And _
                 g.TextMatrix(n, 2) = "0" And _
                 g.TextMatrix(n, 3) = "0" Then

30                g.RemoveItem n
40            End If
50        Next

End Sub

Private Sub cmdXL_Click()

10        On Error GoTo cmdXL_Click_Error

20        ExportFlexGrid g, Me

30        Exit Sub

cmdXL_Click_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmUrineStats", "cmdXL_Click", intEL, strES


End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()

10        On Error GoTo Form_Load_Error

20        dtFrom = Format(Now - 7, "dd/mmm/yyyy")
30        dtTo = Format(Now, "dd/mmm/yyyy")

40        ReCalc

50        Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

60        intEL = Erl
70        strES = Err.Description
80        LogError "frmUrineStats", "Form_Load", intEL, strES

End Sub

Private Sub o_Click(Index As Integer, Value As Integer)

10        ReCalc

End Sub

Private Sub oBetween_Click(Index As Integer, Value As Integer)

          Dim upto As String

10        dtFrom = BetweenDates(Index, upto)
20        dtTo = upto

30        ReCalc

End Sub


