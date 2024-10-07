VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmTotHaem 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Totals for Haematology"
   ClientHeight    =   6735
   ClientLeft      =   330
   ClientTop       =   660
   ClientWidth     =   9405
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
   Icon            =   "frmTotHaem.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6735
   ScaleWidth      =   9405
   Begin VB.PictureBox fraProgress 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   3000
      ScaleHeight     =   525
      ScaleWidth      =   3825
      TabIndex        =   44
      Top             =   3480
      Visible         =   0   'False
      Width           =   3855
      Begin MSComctlLib.ProgressBar pbProgress 
         Height          =   255
         Left            =   0
         TabIndex        =   45
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
         TabIndex        =   46
         Top             =   0
         Width           =   3840
      End
   End
   Begin VB.OptionButton o 
      Caption         =   "Wards"
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
      Left            =   1260
      TabIndex        =   31
      Top             =   1620
      Width           =   855
   End
   Begin VB.CommandButton cmdXL 
      Caption         =   "Export to Excel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   8145
      Picture         =   "frmTotHaem.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   4215
      Width           =   1230
   End
   Begin VB.OptionButton o 
      Caption         =   "Clinicians"
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
      Left            =   240
      TabIndex        =   32
      Top             =   1620
      Value           =   -1  'True
      Width           =   1065
   End
   Begin VB.OptionButton o 
      Caption         =   "G.P.s"
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
      Left            =   2115
      TabIndex        =   30
      Top             =   1620
      Width           =   825
   End
   Begin VB.CommandButton cmdGraph 
      Appearance      =   0  'Flat
      Caption         =   "&Graph"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   6795
      Picture         =   "frmTotHaem.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   4215
      Width           =   1290
   End
   Begin VB.ListBox lstSource 
      BackColor       =   &H00C0C0C0&
      Height          =   4335
      IntegralHeight  =   0   'False
      Left            =   240
      TabIndex        =   18
      Top             =   1950
      Width           =   2685
   End
   Begin VB.CommandButton cmdPrint 
      Appearance      =   0  'Flat
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
      Height          =   840
      Left            =   6795
      Picture         =   "frmTotHaem.frx":0A56
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5475
      Width           =   1290
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   8145
      Picture         =   "frmTotHaem.frx":0D60
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5475
      Width           =   1245
   End
   Begin VB.PictureBox SSPanel1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1425
      Index           =   1
      Left            =   210
      ScaleHeight     =   1365
      ScaleWidth      =   6465
      TabIndex        =   16
      Top             =   135
      Width           =   6525
      Begin VB.ComboBox cmbHosp 
         Height          =   315
         Left            =   3465
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   990
         Width           =   2715
      End
      Begin MSComCtl2.DTPicker calTo 
         Height          =   315
         Left            =   2310
         TabIndex        =   29
         Top             =   120
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   556
         _Version        =   393216
         Format          =   59244545
         CurrentDate     =   38041
      End
      Begin MSComCtl2.DTPicker calFrom 
         Height          =   315
         Left            =   495
         TabIndex        =   28
         Top             =   90
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         Format          =   59244545
         CurrentDate     =   38041
      End
      Begin VB.OptionButton oBetween 
         Caption         =   "Today"
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
         Index           =   6
         Left            =   150
         TabIndex        =   25
         Top             =   540
         Width           =   795
      End
      Begin VB.OptionButton oBetween 
         Caption         =   "Year To Date"
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
         Index           =   5
         Left            =   3195
         TabIndex        =   24
         Top             =   540
         Width           =   1305
      End
      Begin VB.OptionButton oBetween 
         Caption         =   "Last Full Quarter"
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
         Index           =   4
         Left            =   1530
         TabIndex        =   23
         Top             =   1080
         Width           =   1515
      End
      Begin VB.OptionButton oBetween 
         Caption         =   "Last Quarter"
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
         Left            =   1530
         TabIndex        =   22
         Top             =   810
         Width           =   1185
      End
      Begin VB.OptionButton oBetween 
         Caption         =   "Last Full Month"
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
         Left            =   1530
         TabIndex        =   21
         Top             =   540
         Width           =   1395
      End
      Begin VB.OptionButton oBetween 
         Caption         =   "Last Month"
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
         Left            =   135
         TabIndex        =   20
         Top             =   1080
         Width           =   1155
      End
      Begin VB.OptionButton oBetween 
         Caption         =   "Last Week"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   150
         TabIndex        =   19
         Top             =   780
         Width           =   1095
      End
      Begin VB.CommandButton cmdRecalc 
         Caption         =   "&Start"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   4725
         MaskColor       =   &H8000000F&
         Picture         =   "frmTotHaem.frx":106A
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   120
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.Label Label11 
         Caption         =   "To"
         Height          =   195
         Left            =   1980
         TabIndex        =   43
         Top             =   180
         Width           =   420
      End
      Begin VB.Label Label9 
         Caption         =   "From"
         Height          =   195
         Left            =   45
         TabIndex        =   42
         Top             =   135
         Width           =   420
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdHaem 
      Height          =   4665
      Left            =   3150
      TabIndex        =   27
      Top             =   1650
      Width           =   3585
      _ExtentX        =   6324
      _ExtentY        =   8229
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      GridLines       =   2
      FormatString    =   "<Source                            |<Tests      "
   End
   Begin VB.Label lblExcelInfo 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exporting..."
      Height          =   345
      Left            =   8145
      TabIndex        =   40
      Top             =   5025
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Label lblFilm 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   8190
      TabIndex        =   38
      Top             =   2970
      Width           =   1005
   End
   Begin VB.Label Label10 
      Caption         =   "Film"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7020
      TabIndex        =   37
      Top             =   3030
      Width           =   1095
   End
   Begin VB.Label lblRa 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   8190
      TabIndex        =   36
      Top             =   2610
      Width           =   1005
   End
   Begin VB.Label Label7 
      Caption         =   "Ra"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7020
      TabIndex        =   35
      Top             =   2670
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "Asot"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7020
      TabIndex        =   34
      Top             =   2310
      Width           =   1095
   End
   Begin VB.Label lblAsot 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   8190
      TabIndex        =   33
      Top             =   2250
      Width           =   1005
   End
   Begin VB.Label lblTot 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   270
      Left            =   8190
      TabIndex        =   15
      Top             =   3600
      Width           =   1005
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7020
      TabIndex        =   14
      Top             =   3615
      Width           =   555
   End
   Begin VB.Label lblSickle 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   8190
      TabIndex        =   13
      Top             =   1890
      Width           =   1005
   End
   Begin VB.Label lblMalaria 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   8190
      TabIndex        =   12
      Top             =   1530
      Width           =   1005
   End
   Begin VB.Label lblRetics 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   8190
      TabIndex        =   11
      Top             =   1170
      Width           =   1005
   End
   Begin VB.Label lblMonospot 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   8190
      TabIndex        =   10
      Top             =   810
      Width           =   1005
   End
   Begin VB.Label lblEsr 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   8190
      TabIndex        =   9
      Top             =   450
      Width           =   1005
   End
   Begin VB.Label lblFbc 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   8190
      TabIndex        =   8
      Top             =   90
      Width           =   1005
   End
   Begin VB.Label Label6 
      Caption         =   "Monospot"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7020
      TabIndex        =   7
      Top             =   870
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "ESR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7020
      TabIndex        =   6
      Top             =   510
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "FBC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7020
      TabIndex        =   5
      Top             =   150
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Retics"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7020
      TabIndex        =   4
      Top             =   1230
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Malaria"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7020
      TabIndex        =   3
      Top             =   1590
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Sickle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7020
      TabIndex        =   2
      Top             =   1950
      Width           =   1095
   End
End
Attribute VB_Name = "frmTotHaem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit




Private Sub cmbHosp_Click()

          Dim n As Long

10        On Error GoTo cmbHosp_Click_Error

20        For n = 0 To 2
30            If o(n).Value = True Then
40                Filllstsource
50            End If
60        Next

70        cmdRecalc.Visible = True

80        Exit Sub

cmbHosp_Click_Error:

          Dim strES As String
          Dim intEL As Integer



90        intEL = Erl
100       strES = Err.Description
110       LogError "frmTotHaem", "cmbHosp_Click", intEL, strES


End Sub

Private Sub cmdCancel_Click()

10        Unload Me

End Sub

Private Sub cmdPrint_Click()

          Dim n As Long


10        On Error GoTo cmdPrint_Click_Error

20        Printer.Print "Haematology  Tests from " & calFrom & " to " & calTo
30        Printer.Print

40        For n = 0 To 2
50            If o(n).Value = True Then
60                Printer.Print o(n).Caption
70            End If
80        Next

90        For n = 0 To grdHaem.Rows - 1
100           Printer.Print grdHaem.TextMatrix(n, 0);
110           Printer.Print Tab(40); grdHaem.TextMatrix(n, 1)
120       Next


130       Printer.Print Tab(10); "        FBC : " & lblFbc
140       Printer.Print Tab(10); "        ESR : " & lblEsr
150       Printer.Print Tab(10); "  MONOSPOTS : " & lblMonospot
160       Printer.Print Tab(10); "     RETICS : " & lblRetics

170       Printer.Print "----------End of REport-----------"

180       Printer.EndDoc

190       Exit Sub

cmdPrint_Click_Error:

          Dim strES As String
          Dim intEL As Integer



200       intEL = Erl
210       strES = Err.Description
220       LogError "frmTotHaem", "cmdPrint_Click", intEL, strES


End Sub

Private Sub cmdRecalc_Click()

10        On Error GoTo cmdRecalc_Click_Error

20        If DateDiff("d", calFrom, calTo) < 0 Then
30            iMsg "Wrong date selection. From date cannot be greater than To date", vbInformation
40            Exit Sub
50        End If
60        FilllTotal

70        Exit Sub

cmdRecalc_Click_Error:

          Dim strES As String
          Dim intEL As Integer



80        intEL = Erl
90        strES = Err.Description
100       LogError "frmTotHaem", "cmdRecalc_Click", intEL, strES


End Sub

Private Sub Filllstsource()
          Dim Found As Boolean
          Dim n As Long
          Dim tb As New Recordset
          Dim sql As String

10        On Error GoTo Filllstsource_Error

20        lstSource.Clear

30        If o(0) Then
40            sql = "SELECT * from Clinicians WHERE hospitalcode = '" & ListCodeFor("HO", cmbHosp) & "'"
50            Set tb = New Recordset
60            RecOpenServer 0, tb, sql
70            Do While Not tb.EOF
80                For n = 1 To lstSource.ListCount
90                    If lstSource.List(n) = Trim(tb!Text) Then
100                       Found = True
110                       Exit For
120                   End If
130               Next
140               If Found = False Then lstSource.AddItem Trim(tb!Text)
150               Found = False
160               tb.MoveNext
170           Loop

180       ElseIf o(1) Then
190           sql = "SELECT * from wards WHERE hospitalcode = '" & ListCodeFor("HO", cmbHosp) & "'"
200           Set tb = New Recordset
210           RecOpenServer 0, tb, sql
220           Do While Not tb.EOF
230               For n = 1 To lstSource.ListCount
240                   If lstSource.List(n) = Trim(tb!Text) Then
250                       Found = True
260                       Exit For
270                   End If
280               Next
290               If Found = False Then lstSource.AddItem Trim(tb!Text)
300               Found = False
310               tb.MoveNext
320           Loop
330       Else
340           sql = "SELECT * from gps WHERE hospitalcode = '" & ListCodeFor("HO", cmbHosp) & "' ORDER BY LISTORDER"
350           Set tb = New Recordset
360           RecOpenServer 0, tb, sql
370           Do While Not tb.EOF
380               For n = 1 To lstSource.ListCount
390                   If lstSource.List(n) = Trim(tb!Text) Then
400                       Found = True
410                       Exit For
420                   End If
430               Next
440               If Found = False Then lstSource.AddItem Trim(tb!Text)
450               Found = False
460               tb.MoveNext
470           Loop
480       End If

490       Exit Sub

Filllstsource_Error:

          Dim strES As String
          Dim intEL As Integer



500       intEL = Erl
510       strES = Err.Description
520       LogError "frmTotHaem", "Filllstsource", intEL, strES, sql


End Sub

Private Sub FilllTotal()

          Dim tb As New Recordset
          Dim sql As String
          Dim n As Long
          Dim Count As Long
          Dim s As String
          Dim Tot As Long
          Dim FromDate As String
          Dim ToDate As String
          Dim Source As String
          Dim SourceTable As String
          Dim MonthIndex As Integer
          Dim StartDate As Date
          Dim EndDate As Date
          Dim SrcUpdated As Boolean
          Dim DiffMonths As Integer

10        On Error GoTo FilllTotal_Error

20        FromDate = Format$(calFrom, "dd/mmm/yyyy") & " 00:00:00"
30        ToDate = Format$(calTo, "dd/mmm/yyyy") & " 23:59:59"

40        lblFbc = "0"
50        lblEsr = "0"
60        lblMonospot = "0"
70        lblRetics = "0"
80        lblMalaria = "0"
90        lblSickle = "0"
100       lblTot = "0"
110       lblAsot = "0"


120       StartDate = calFrom
130       pbProgress.Value = 1
140       DiffMonths = DateDiff("m", calFrom, calTo)
150       If DiffMonths = 0 Then
160           EndDate = calTo
170           pbProgress.Max = 4
180           DiffMonths = DiffMonths + 1
190       Else
200           EndDate = DateAdd("m", 1, calFrom)
210           EndDate = "01/" & Month(EndDate) & "/" & Year(EndDate)
220           EndDate = DateAdd("d", -1, EndDate)
230           pbProgress.Max = DiffMonths + 3
240       End If

250       fraProgress.Visible = True
260       pbProgress.Value = pbProgress.Value + 1
270       lblProgress = "Fetching results ... (" & Int(pbProgress.Value * 100 / pbProgress.Max) & " %)"
280       lblProgress.Refresh

290       sql = "IF OBJECT_ID('tempdb..##temp') IS NOT NULL DROP TABLE ##temp " & _
                "SELECT WBC, cESR, cMonospot, RetP, cMalaria,cSickledex,cAsot,cRA,cFilm into ##temp " & _
                "From HaemResults " & _
                "WHERE (RunDateTime BETWEEN '" & FromDate & "' AND '" & ToDate & "') "
300       Cnxn(0).Execute sql

310       sql = "SELECT " & _
                "TotWBC = (SELECT COUNT(WBC) FROM ##temp WHERE WBC <> ''), " & _
                "TotESR = (SELECT COUNT(cESR) FROM ##temp WHERE cESR = 1), " & _
                "TotMonoSpot = (SELECT COUNT(cMonospot) FROM ##temp WHERE cMonoSpot = 1), " & _
                "TotRetP = (SELECT COUNT(RetP) FROM ##temp WHERE RetP <> ''), " & _
                "TotMalaria = (SELECT COUNT(cMalaria) FROM ##temp WHERE cMalaria = 1), " & _
                "TotSickledex = (SELECT COUNT(cSickledex) FROM ##temp WHERE cSickledex = 1), " & _
                "TotASOT = (SELECT COUNT(cAsot) FROM ##temp WHERE cASOT = 1), " & _
                "TotRA = (SELECT COUNT(cRA) FROM ##temp WHERE cRA = 1), " & _
                "TotFilm = (SELECT COUNT(cFilm) FROM ##temp WHERE cFilm = 1)"




          'sql = "SELECT " & _
           '      "TotWBC = (SELECT COUNT(*) FROM HaemResults WHERE WBC <> '' " & _
           '      "          AND RunDateTime BETWEEN '" & FromDate & "' AND '" & ToDate & "'), " & _
           '      "TotESR = (SELECT COUNT(*) FROM HaemResults WHERE cESR = 1 " & _
           '      "          AND RunDateTime BETWEEN '" & FromDate & "' AND '" & ToDate & "'), " & _
           '      "TotMonoSpot = (SELECT COUNT(*) FROM HaemResults WHERE cMonoSpot = 1 " & _
           '      "          AND RunDateTime BETWEEN '" & FromDate & "' AND '" & ToDate & "'), " & _
           '      "TotRetP = (SELECT COUNT(*) FROM HaemResults WHERE RetP <> '' " & _
           '      "          AND RunDateTime BETWEEN '" & FromDate & "' AND '" & ToDate & "'), " & _
           '      "TotMalaria = (SELECT COUNT(*) FROM HaemResults WHERE cMalaria = 1 " & _
           '      "          AND RunDateTime BETWEEN '" & FromDate & "' AND '" & ToDate & "'), " & _
           '      "TotSickledex = (SELECT COUNT(*) FROM HaemResults WHERE cSickledex = 1 " & _
           '      "          AND RunDateTime BETWEEN '" & FromDate & "' AND '" & ToDate & "'), " & _
           '      "TotASOT = (SELECT COUNT(*) FROM HaemResults WHERE cASOT = 1 " & _
           '      "          AND RunDateTime BETWEEN '" & FromDate & "' AND '" & ToDate & "'), " & _
           '      "TotRA = (SELECT COUNT(*) FROM HaemResults WHERE cRA = 1 " & _
           '      "          AND RunDateTime BETWEEN '" & FromDate & "' AND '" & ToDate & "'), " & _
           '      "TotFilm = (SELECT COUNT(*) FROM HaemResults WHERE cFilm = 1 " & _
           '      "          AND RunDateTime BETWEEN '" & FromDate & "' AND '" & ToDate & "')"

320       Set tb = New Recordset

330       RecOpenServer 0, tb, sql
340       lblFbc = tb!TotWBC
350       lblEsr = tb!TotESR
360       lblMonospot = tb!TotMonoSpot
370       lblRetics = tb!TotRetP
380       lblMalaria = tb!TotMalaria
390       lblSickle = tb!TotSickledex
400       lblAsot = tb!TotASOT
410       lblRa = tb!TotRA
420       lblFilm = tb!TotFilm

          'lblTot = Val(lblFbc) + Val(lblEsr) + Val(lblMonospot) + Val(lblRetics) + Val(lblMalaria) + Val(lblSickle) + Val(lblAsot) + Val(lblRa) + Val(lblFilm)
          'lblTot.Refresh



430       grdHaem.Rows = 2
440       grdHaem.AddItem ""
450       grdHaem.RemoveItem 1

460       If o(0) Then
470           Source = "Clinician"
480           SourceTable = "Clinicians"
490       ElseIf o(1) Then
500           Source = "Ward"
510           SourceTable = "Wards"
520       Else
530           Source = "GP"
540           SourceTable = "GPs"
550       End If


560       For MonthIndex = 1 To DiffMonths
570           pbProgress.Value = pbProgress.Value + 1
580           lblProgress = "Fetching results ... (" & Int(pbProgress.Value * 100 / pbProgress.Max) & " %)"
590           lblProgress.Refresh

600           sql = "SELECT D." & Source & " AS Src, COUNT(DISTINCT(D.SampleID)) Tot " & _
                    "FROM Demographics D JOIN HaemResults R " & _
                    "ON D.SampleID = R.SampleID " & _
                    "WHERE " & Source & " IN (SELECT Text FROM " & SourceTable & " " & _
                    "WHERE HospitalCode = '" & ListCodeFor("HO", cmbHosp) & "') " & _
                    "AND D.Hospital = '" & cmbHosp & "' " & _
                    "AND D.RunDate BETWEEN '" & _
                    Format$(StartDate, "dd/mmm/yyyy 00:00:00") & "' AND '" & _
                    Format$(EndDate, "dd/mmm/yyyy 23:59:59") & "' " & _
                    "GROUP BY D." & Source

610           Set tb = New Recordset
620           RecOpenClient 0, tb, sql
630           If Not tb.EOF Then
640               While Not tb.EOF
650                   Tot = Tot + tb!Tot
660                   SrcUpdated = False
670                   For n = 1 To grdHaem.Rows - 1
680                       If grdHaem.TextMatrix(n, 0) = tb!Src Then
690                           grdHaem.TextMatrix(n, 1) = Val(grdHaem.TextMatrix(n, 1)) + tb!Tot
700                           SrcUpdated = True
710                       End If
720                   Next n
730                   If Not SrcUpdated Then
740                       grdHaem.AddItem tb!Src & vbTab & tb!Tot
750                   End If

760                   tb.MoveNext
770               Wend
780           End If
790           StartDate = DateAdd("d", 1, EndDate)
800           EndDate = DateAdd("m", 1, StartDate)
810           EndDate = "01/" & Month(EndDate) & "/" & Year(EndDate)
820           EndDate = DateAdd("d", -1, EndDate)
830           If MonthIndex = DiffMonths And DateDiff("d", ToDate, EndDate) > 0 Then
840               EndDate = ToDate
850           End If
860       Next MonthIndex




          'For n = 0 To lstSource.ListCount - 1
          '    lstSource.Selected(n) = True
          '    sql = "SELECT COUNT(DISTINCT(D.SampleID)) Tot " & _
               '          "FROM Demographics D JOIN HaemResults R " & _
               '          "ON D.SampleID = R.SampleID " & _
               '          "WHERE " & Source & " = '" & AddTicks(lstSource.List(n)) & "' " & _
               '          "AND D.Hospital = '" & cmbHosp & "' " & _
               '          "AND D.RunDate BETWEEN '" & _
               '          Format$(calFrom, "dd/mmm/yyyy") & "' AND '" & _
               '          Format$(calTo, "dd/mmm/yyyy") & "'"
          '    Set tb = New Recordset
          '    RecOpenServer 0, tb, sql
          '    Count = tb!Tot
          '    If Count <> 0 Then
          '        s = initial2upper(lstSource.List(n)) & vbTab & Format$(Count)
          '        Tot = Tot + Count
          '        grdHaem.AddItem s
          '    End If
          'Next

870       grdHaem.AddItem "Total" & vbTab & Tot

880       If grdHaem.Rows > 2 And grdHaem.TextMatrix(1, 0) = "" And grdHaem.TextMatrix(1, 1) = "" Then grdHaem.RemoveItem 1

890       sql = "SELECT Count(D.SampleID) As TotCnt FROM Demographics D " & _
                "INNER JOIN HaemResults R ON D.SampleID = R.SampleID " & _
                "WHERE D.Hospital = '" & cmbHosp & "' " & _
                "AND D.RunDate BETWEEN '" & Format$(calFrom, "dd/mmm/yyyy") & "' AND '" & _
                Format$(calTo, "dd/mmm/yyyy") & "' "

900       Set tb = New Recordset
910       RecOpenClient 0, tb, sql
920       pbProgress.Value = pbProgress.Value + 1
930       lblProgress = "Fetching results ... (" & Int(pbProgress.Value * 100 / pbProgress.Max) & " %)"
940       lblProgress.Refresh
950       lblTot = tb!TotCnt

960       fraProgress.Visible = False

970       Exit Sub

FilllTotal_Error:

          Dim strES As String
          Dim intEL As Integer

980       intEL = Erl
990       strES = Err.Description
1000      LogError "frmTotHaem", "FilllTotal", intEL, strES, sql
1010      fraProgress.Visible = False

End Sub




Private Sub calFrom_CloseUp()

10        On Error GoTo calFrom_CloseUp_Error

20        cmdRecalc.Visible = True

30        Exit Sub

calFrom_CloseUp_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmTotHaem", "calFrom_CloseUp", intEL, strES


End Sub



Private Sub calTo_CloseUp()

10        On Error GoTo calTo_CloseUp_Error

20        cmdRecalc.Visible = True

30        Exit Sub

calTo_CloseUp_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmTotHaem", "calTo_CloseUp", intEL, strES


End Sub

Private Sub cmdGraph_Click()

10        On Error GoTo cmdGraph_Click_Error

20        If grdHaem.TextMatrix(1, 1) = "" Then Exit Sub

30        With frmGraph
40            .DrawGraph Me, grdHaem
50            .Show 1
60        End With

70        Exit Sub

cmdGraph_Click_Error:

          Dim strES As String
          Dim intEL As Integer



80        intEL = Erl
90        strES = Err.Description
100       LogError "frmTotHaem", "cmdGraph_Click", intEL, strES


End Sub
Private Sub FillHosp()
          Dim sql As String
          Dim tb As Recordset

10        On Error GoTo FillHosp_Error

20        sql = "SELECT * from lists WHERE listtype = 'HO'"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        Do While Not tb.EOF
60            If UCase(tb!Text) = HospName(0) Then
70                cmbHosp.AddItem tb!Text, 0
80            Else
90                cmbHosp.AddItem tb!Text
100           End If
110           tb.MoveNext
120       Loop

130       cmbHosp.ListIndex = 0

140       Exit Sub

FillHosp_Error:

          Dim strES As String
          Dim intEL As Integer



150       intEL = Erl
160       strES = Err.Description
170       LogError "frmTotHaem", "FillHosp", intEL, strES, sql


End Sub

Private Sub Form_Load()

10        On Error GoTo Form_Load_Error

20        calFrom = Format$(Now, "dd/mmm/yyyy")
30        calTo = calFrom

40        ResetControls
50        FillHosp
60        Filllstsource
          'FilllTotal
70        Set_Font Me

80        cmdRecalc.Visible = True

90        Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer



100       intEL = Erl
110       strES = Err.Description
120       LogError "frmTotHaem", "Form_Load", intEL, strES


End Sub
Private Sub cmdXL_Click()

          Dim strHeading As String

10        On Error GoTo cmdXL_Click_Error

20        strHeading = "Totals for Haematology" & vbCr
30        strHeading = strHeading & "From " & calFrom & " To " & calTo
40        strHeading = strHeading & vbCr
50        ExportFlexGrid grdHaem, Me, strHeading

60        Exit Sub

cmdXL_Click_Error:

          Dim strES As String
          Dim intEL As Integer



70        intEL = Erl
80        strES = Err.Description
90        LogError "frmTotHaem", "cmdXL_Click", intEL, strES


End Sub

Private Sub oBetween_Click(Index As Integer)

          Dim upto As String

10        On Error GoTo oBetween_Click_Error

20        calFrom = Format$(BetweenDates(Index, upto), "dd/mmm/yyyy")
30        calTo = Format$(upto, "dd/mmm/yyyy")

          'FilllTotal

40        Exit Sub

oBetween_Click_Error:

          Dim strES As String
          Dim intEL As Integer



50        intEL = Erl
60        strES = Err.Description
70        LogError "frmTotHaem", "oBetween_Click", intEL, strES


End Sub


Private Sub o_Click(Index As Integer)

10        On Error GoTo o_Click_Error

20        ResetControls
30        Filllstsource
          'FilllTotal

40        Exit Sub

o_Click_Error:

          Dim strES As String
          Dim intEL As Integer



50        intEL = Erl
60        strES = Err.Description
70        LogError "frmTotHaem", "o_Click", intEL, strES


End Sub

Private Sub ResetControls()
10        With grdHaem
20            .Clear
30            .Rows = 2
40            .FixedRows = 1
50            .Refresh
60        End With
70        lblAsot = ""
80        lblEsr = ""
90        lblFbc = ""
100       lblFilm = ""
110       lblMalaria = ""
120       lblMonospot = ""
130       lblRa = ""
140       lblRetics = ""
150       lblSickle = ""
160       lblTot = ""
170       If o(0).Value = True Then
180           grdHaem.FormatString = "<" & Left(o(0).Caption & Space(34), 34) & "|<Tests      "
190       ElseIf o(1).Value = True Then
200           grdHaem.FormatString = "<" & Left(o(1).Caption & Space(34), 34) & "|<Tests      "
210       ElseIf o(2).Value = True Then
220           grdHaem.FormatString = "<" & Left(o(2).Caption & Space(34), 34) & "|<Tests      "
230       End If

End Sub
