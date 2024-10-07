VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form f 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Biochemistry Usage"
   ClientHeight    =   5130
   ClientLeft      =   315
   ClientTop       =   1110
   ClientWidth     =   8625
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
   Icon            =   "frmStats.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5130
   ScaleWidth      =   8625
   ShowInTaskbar   =   0   'False
   Begin ComctlLib.ProgressBar bar 
      Height          =   225
      Left            =   225
      TabIndex        =   3
      Top             =   990
      Visible         =   0   'False
      Width           =   6120
      _ExtentX        =   10795
      _ExtentY        =   397
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin MSComCtl2.DTPicker calToDate 
      Height          =   375
      Left            =   2745
      TabIndex        =   44
      Top             =   135
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   661
      _Version        =   393216
      Format          =   175046657
      CurrentDate     =   38503
   End
   Begin VB.CommandButton bGP 
      Caption         =   "GP's"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   930
      Left            =   7020
      Picture         =   "frmStats.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   1980
      Width           =   1020
   End
   Begin VB.Frame Frame3 
      Caption         =   "Totals"
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
      Left            =   6660
      TabIndex        =   24
      Top             =   45
      Width           =   1845
      Begin VB.TextBox bsamples 
         Height          =   285
         Left            =   840
         TabIndex        =   27
         Text            =   "0"
         Top             =   240
         Width           =   885
      End
      Begin VB.TextBox btests 
         Height          =   285
         Left            =   840
         TabIndex        =   26
         Text            =   "0"
         Top             =   600
         Width           =   885
      End
      Begin VB.TextBox tpers 
         Height          =   285
         Left            =   840
         TabIndex        =   25
         Top             =   990
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Samples"
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
         Left            =   150
         TabIndex        =   30
         Top             =   270
         Width           =   600
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "Tests per Sample"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   29
         Top             =   930
         Width           =   675
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tests"
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
         Left            =   360
         TabIndex        =   28
         Top             =   630
         Width           =   390
      End
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   225
      TabIndex        =   16
      Top             =   1125
      Width           =   6090
      Begin VB.OptionButton o 
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
         Left            =   4815
         TabIndex        =   23
         Top             =   270
         Width           =   1215
      End
      Begin VB.OptionButton o 
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
         Left            =   3105
         TabIndex        =   22
         Top             =   495
         Width           =   1485
      End
      Begin VB.OptionButton o 
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
         Left            =   3105
         TabIndex        =   21
         Top             =   270
         Width           =   1410
      End
      Begin VB.OptionButton o 
         Caption         =   "Year to Date"
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
         Left            =   1530
         TabIndex        =   20
         Top             =   495
         Width           =   1215
      End
      Begin VB.OptionButton o 
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
         Left            =   1530
         TabIndex        =   19
         Top             =   270
         Width           =   1155
      End
      Begin VB.OptionButton o 
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
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   18
         Top             =   510
         Value           =   -1  'True
         Width           =   1125
      End
      Begin VB.OptionButton o 
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
         TabIndex        =   17
         Top             =   270
         Width           =   825
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3165
      Left            =   45
      TabIndex        =   4
      Top             =   1845
      Width           =   6645
      Begin MSFlexGridLib.MSFlexGrid g 
         Height          =   2745
         Left            =   1110
         TabIndex        =   5
         Top             =   270
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   4842
         _Version        =   393216
         Rows            =   11
         Cols            =   7
         FixedCols       =   0
         BackColor       =   -2147483624
         ForeColor       =   -2147483635
         BackColorFixed  =   -2147483647
         ForeColorFixed  =   -2147483624
         BackColorBkg    =   -2147483647
         AllowBigSelection=   0   'False
         GridLines       =   3
         GridLinesFixed  =   3
         ScrollBars      =   0
         FormatString    =   "<Samples  |<%     ||<Tests        |<%     ||<T/S     "
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Misc."
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
         Left            =   5235
         TabIndex        =   40
         Top             =   2760
         Width           =   375
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Medical"
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
         Left            =   5235
         TabIndex        =   39
         Top             =   570
         Width           =   555
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Surgical"
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
         Left            =   5235
         TabIndex        =   38
         Top             =   810
         Width           =   570
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Paeds"
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
         Left            =   5235
         TabIndex        =   37
         Top             =   1050
         Width           =   450
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "G.P."
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
         Left            =   5235
         TabIndex        =   36
         Top             =   1305
         Width           =   315
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Ext. Hosp"
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
         Left            =   5235
         TabIndex        =   35
         Top             =   1545
         Width           =   690
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Maternity"
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
         Left            =   5235
         TabIndex        =   34
         Top             =   1785
         Width           =   645
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Out Patients"
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
         Left            =   5235
         TabIndex        =   33
         Top             =   2025
         Width           =   870
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Rooms"
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
         Left            =   5235
         TabIndex        =   32
         Top             =   2265
         Width           =   495
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Casualty"
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
         Left            =   5235
         TabIndex        =   31
         Top             =   2520
         Width           =   600
      End
      Begin VB.Label l 
         AutoSize        =   -1  'True
         Caption         =   "Casualty"
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
         Index           =   9
         Left            =   450
         TabIndex        =   15
         Top             =   2520
         Width           =   600
      End
      Begin VB.Label l 
         AutoSize        =   -1  'True
         Caption         =   "Rooms"
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
         Index           =   8
         Left            =   555
         TabIndex        =   14
         Top             =   2265
         Width           =   495
      End
      Begin VB.Label l 
         AutoSize        =   -1  'True
         Caption         =   "Out Patients"
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
         Index           =   7
         Left            =   180
         TabIndex        =   13
         Top             =   2025
         Width           =   870
      End
      Begin VB.Label l 
         AutoSize        =   -1  'True
         Caption         =   "Maternity"
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
         Left            =   405
         TabIndex        =   12
         Top             =   1785
         Width           =   645
      End
      Begin VB.Label l 
         AutoSize        =   -1  'True
         Caption         =   "Ext. Hosp"
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
         Left            =   360
         TabIndex        =   11
         Top             =   1545
         Width           =   690
      End
      Begin VB.Label l 
         AutoSize        =   -1  'True
         Caption         =   "G.P."
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
         Left            =   750
         TabIndex        =   10
         Top             =   1305
         Width           =   315
      End
      Begin VB.Label l 
         AutoSize        =   -1  'True
         Caption         =   "Paeds"
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
         Left            =   600
         TabIndex        =   9
         Top             =   1050
         Width           =   450
      End
      Begin VB.Label l 
         AutoSize        =   -1  'True
         Caption         =   "Surgical"
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
         Left            =   480
         TabIndex        =   8
         Top             =   810
         Width           =   570
      End
      Begin VB.Label l 
         AutoSize        =   -1  'True
         Caption         =   "Medical"
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
         Left            =   495
         TabIndex        =   7
         Top             =   570
         Width           =   555
      End
      Begin VB.Label l 
         AutoSize        =   -1  'True
         Caption         =   "Misc."
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
         Index           =   10
         Left            =   675
         TabIndex        =   6
         Top             =   2760
         Width           =   375
      End
   End
   Begin VB.CommandButton bcancel 
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
      Height          =   900
      Left            =   7020
      Picture         =   "frmStats.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3915
      Width           =   1020
   End
   Begin VB.CommandButton Bstart 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   4680
      Picture         =   "frmStats.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   45
      Width           =   1470
   End
   Begin VB.CommandButton bprint 
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
      Height          =   855
      Left            =   7020
      Picture         =   "frmStats.frx":0C28
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2970
      Width           =   1020
   End
   Begin MSComCtl2.DTPicker calFromDate 
      Height          =   375
      Left            =   630
      TabIndex        =   45
      Top             =   135
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   661
      _Version        =   393216
      Format          =   175112193
      CurrentDate     =   38503
   End
   Begin VB.Label Label14 
      Caption         =   "To"
      Height          =   240
      Left            =   2430
      TabIndex        =   47
      Top             =   180
      Width           =   330
   End
   Begin VB.Label Label13 
      Caption         =   "From"
      Height          =   240
      Left            =   45
      TabIndex        =   46
      Top             =   180
      Width           =   465
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "l(0) - do not remove"
      Height          =   195
      Left            =   11670
      TabIndex        =   43
      Top             =   1020
      Width           =   1680
   End
   Begin VB.Label l 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   0
      Left            =   12210
      TabIndex        =   42
      Top             =   1230
      Width           =   645
   End
End
Attribute VB_Name = "f"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim Cntr As Counterb
Dim Cntrs As New Counterbs



Private Sub bcancel_Click()

10        Unload Me

End Sub

Private Sub bGP_Click()

10        On Error GoTo bGP_Click_Error

20        With frmGPStats
30            .bsamples = bsamples
40            .btests = btests
50            .tpers = tpers
60            .calFromDate = calFromDate
70            .calToDate = calToDate
80            .Show 1
90        End With

100       Exit Sub

bGP_Click_Error:

          Dim strES As String
          Dim intEL As Integer



110       intEL = Erl
120       strES = Err.Description
130       LogError "frmStats", "bGP_Click", intEL, strES


End Sub

Private Sub bprint_Click()

          Dim n As Long

10        On Error GoTo bprint_Click_Error

20        Printer.Print "Analysis of Tests and Samples"
30        Printer.Print
40        Printer.Print "Between dates "; calFromDate; " and "; calToDate
50        Printer.Print

60        Printer.Print Tab(25); "Biochemistry Tests "; btests
70        Printer.Print Tab(25); "           Samples "; bsamples
80        Printer.Print Tab(25); "  Tests per Sample "; tpers
90        Printer.Print

100       For n = 0 To 10
110           g.Row = n
120           Printer.Print l(n);
130           g.Col = 0
140           Printer.Print Tab(16); g;
150           g.Col = 1
160           Printer.Print Tab(25); g;
170           g.Col = 3
180           Printer.Print Tab(34); g;
190           g.Col = 4
200           Printer.Print Tab(42); g;
210           g.Col = 6
220           Printer.Print Tab(51); g
230       Next

240       Printer.EndDoc

250       Exit Sub

bprint_Click_Error:

          Dim strES As String
          Dim intEL As Integer



260       intEL = Erl
270       strES = Err.Description
280       LogError "frmStats", "bPrint_Click", intEL, strES


End Sub

Private Sub bStart_Click()

          Dim tb As New Recordset
          Dim sn As New Recordset
          Dim sql As String

10        On Error GoTo bStart_Click_Error

20        g.Rows = 2

30        bar.Visible = True
40        bar.Max = 100

50        sql = "SELECT demographics.rundate, demographics.sampleid, " & _
                "demographics.clinician, demographics.ward, demographics.gp from demographics, bioresults WHERE (" & _
                "demographics.rundate between '" & _
                Format(calFromDate, "dd/mmm/yyyy") & "' and '" & _
                Format(calToDate, "dd/mmm/yyyy") & "') and bioresults.sampleid = demographics.sampleid"
60        Set sn = New Recordset
70        RecOpenServer 0, sn, sql

80        If Not sn.EOF Then
90            Do While Not sn.EOF
100               sql = "SELECT * from bioresults WHERE " & _
                        "sampleid = '" & sn!SampleID & "'"
110               Set tb = New Recordset
120               RecOpenServer 0, tb, sql
130               If Not tb.EOF Then
140                   Cntrs.Add sn!Clinician & "", sn!Ward & "", "S"
150                   Do While Not tb.EOF
160                       Cntrs.Add sn!Clinician & "", sn!Ward & "", "T"
170                       tb.MoveNext
180                   Loop
190               End If
200               sn.MoveNext
210           Loop
220       End If

230       RemoveUnwanted

240       bar.Visible = False
250       bGP.Enabled = True

260       Exit Sub

bStart_Click_Error:

          Dim strES As String
          Dim intEL As Integer

270       intEL = Erl
280       strES = Err.Description
290       LogError "frmStats", "bStart_Click", intEL, strES, sql

End Sub
Private Sub RemoveUnwanted()

          Dim Link As Long
          Dim n As Long
          Dim PerCent As Single
          Dim Tot As Long
          Dim tests As Single

10        On Error GoTo RemoveUnwanted_Error

20        g.Rows = 2
30        g.AddItem ""
40        g.RemoveItem 1
50        g.Rows = 11

60        For Each Cntr In Cntrs

70            g.Row = Link
80            g.Col = 0
90            g = Format(Val(g) + Cntr.SampleCounter)

100           g.Col = 3
110           g = Format(Val(g) + Cntr.TestCounter)



120           g.Row = 10
130           g.Col = 0
140           g = Format(Val(g) + Cntr.SampleCounter)

150           g.Col = 3
160           g = Format(Val(g) + Cntr.TestCounter)

170       Next

180       g.Col = 0
190       Tot = 0
200       For n = 1 To 10
210           g.Row = n
220           Tot = Tot + Val(g)
230       Next
240       bsamples = Tot

250       g.Col = 3
260       Tot = 0
270       For n = 1 To 10
280           g.Row = n
290           Tot = Tot + Val(g)
300       Next
310       btests = Tot

320       If Val(btests) * Val(bsamples) <> 0 Then
330           tpers = Format(Val(btests) / Val(bsamples), "0.00")
340       Else
350           tpers = ""
360       End If

370       For n = 1 To 10
380           g.Row = n
390           g.Col = 0
400           tests = Val(g)
410           If tests > 0 Then
420               g.Col = 3
430               tests = Val(g) / tests
440               g.Col = 6
450               g = Format(tests, "0.00")
460           End If
470       Next

480       For n = 1 To 10
490           g.Row = n

500           If Val(bsamples) > 0 Then
510               g.Col = 0
520               PerCent = Format((Val(g) / Val(bsamples)) * 100, "#0.0")
530               g.Col = 1
540               g = PerCent
550           Else
560               g.Col = 1
570               g = ""
580           End If

590           If Val(btests) > 0 Then
600               g.Col = 3
610               PerCent = Format((Val(g) / Val(btests)) * 100, "#0.0")
620               g.Col = 4
630               g = PerCent
640           Else
650               g.Col = 4
660               g = ""
670           End If

680       Next

690       For n = 1 To 10
700           g.Row = n
710           g.Col = 2
720           g.CellBackColor = &H80000001
730           g.Col = 5
740           g.CellBackColor = &H80000001
750       Next

760       Exit Sub

RemoveUnwanted_Error:

          Dim strES As String
          Dim intEL As Integer



770       intEL = Erl
780       strES = Err.Description
790       LogError "frmStats", "RemoveUnwanted", intEL, strES


End Sub


Private Sub Form_Load()

          Dim n As Long

10        On Error GoTo Form_Load_Error

20        calFromDate = Format(Now - 7, "dd/mmm/yyyy")
30        calToDate = Format(Now, "dd/mmm/yyyy")

40        For n = 1 To 10
50            g.Row = n
60            g.Col = 2
70            g.CellBackColor = &H80000001
80            g.Col = 5
90            g.CellBackColor = &H80000001
100       Next

110       Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer



120       intEL = Erl
130       strES = Err.Description
140       LogError "frmStats", "Form_Load", intEL, strES


End Sub



Private Sub o_Click(Index As Integer)

          Dim upto As String

10        On Error GoTo o_Click_Error

20        calFromDate = BetweenDates(Index, upto)
30        calToDate = upto

40        Exit Sub

o_Click_Error:

          Dim strES As String
          Dim intEL As Integer



50        intEL = Erl
60        strES = Err.Description
70        LogError "frmStats", "o_Click", intEL, strES


End Sub


