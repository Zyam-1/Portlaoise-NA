VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFullBio 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   10485
   ClientLeft      =   165
   ClientTop       =   465
   ClientWidth     =   12030
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
   Icon            =   "frmFullBio.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   10485
   ScaleWidth      =   12030
   Begin VB.ComboBox cmbResultCount 
      Height          =   315
      ItemData        =   "frmFullBio.frx":030A
      Left            =   690
      List            =   "frmFullBio.frx":0320
      TabIndex        =   48
      Text            =   "All"
      Top             =   900
      Width           =   1215
   End
   Begin VB.CheckBox chkChartNumber 
      Caption         =   "Ignore chart number"
      Height          =   255
      Left            =   210
      TabIndex        =   45
      Top             =   600
      Width           =   2535
   End
   Begin VB.Frame Frame3 
      Height          =   855
      Left            =   2850
      TabIndex        =   39
      Top             =   360
      Width           =   6375
      Begin VB.CommandButton cmdRefresh 
         Height          =   615
         Left            =   5460
         Picture         =   "frmFullBio.frx":0352
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Click to refresh biochemistry history"
         Top             =   135
         Width           =   735
      End
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   330
         Left            =   1440
         TabIndex        =   40
         Top             =   315
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
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
         CustomFormat    =   "dd/MMM/yyyy"
         Format          =   287309827
         CurrentDate     =   38629
      End
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   330
         Left            =   3510
         TabIndex        =   41
         Top             =   315
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
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
         CustomFormat    =   "dd/MMM/yyyy"
         Format          =   287309827
         CurrentDate     =   38629
      End
      Begin VB.Label Label6 
         Caption         =   "to"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3105
         TabIndex        =   43
         Top             =   315
         Width           =   330
      End
      Begin VB.Label lblBetween 
         Caption         =   "Between"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   180
         TabIndex        =   42
         Top             =   315
         Width           =   1185
      End
   End
   Begin VB.CommandButton cmdExcel 
      Caption         =   "Export"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   9240
      Picture         =   "frmFullBio.frx":0C1C
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   9540
      Width           =   1245
   End
   Begin VB.Frame fraHL 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4470
      TabIndex        =   31
      Top             =   7170
      Visible         =   0   'False
      Width           =   3315
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Abnormal Results are shown "
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
         Left            =   90
         TabIndex        =   34
         Top             =   30
         Width           =   2055
      End
      Begin VB.Label Label10 
         BackColor       =   &H000000FF&
         Caption         =   "HIGH"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   2310
         TabIndex        =   33
         Top             =   30
         Width           =   465
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "LOW"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   2820
         TabIndex        =   32
         Top             =   30
         Width           =   405
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Print Cumulative Report"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   7560
      TabIndex        =   22
      Top             =   7800
      Width           =   4305
      Begin VB.CommandButton cmdPrint 
         Height          =   435
         Left            =   3570
         Picture         =   "frmFullBio.frx":683A
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Print Cumulative Report"
         Top             =   690
         Width           =   555
      End
      Begin VB.ComboBox cmbPrinter 
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
         Left            =   990
         TabIndex        =   23
         Text            =   "cmbPrinter"
         Top             =   330
         Width           =   3195
      End
      Begin MSComCtl2.UpDown udPrevious 
         Height          =   285
         Left            =   2415
         TabIndex        =   25
         Top             =   780
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   503
         _Version        =   393216
         Value           =   6
         BuddyControl    =   "lblPrevious"
         BuddyDispid     =   196626
         OrigLeft        =   2460
         OrigTop         =   630
         OrigRight       =   3120
         OrigBottom      =   945
         Max             =   99
         Min             =   1
         Orientation     =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label lblPrevious 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1890
         TabIndex        =   28
         Top             =   780
         Width           =   435
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Current and Previous"
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
         Left            =   210
         TabIndex        =   27
         Top             =   840
         Width           =   1515
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Use Printer"
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
         Left            =   120
         TabIndex        =   26
         Top             =   420
         Width           =   780
      End
   End
   Begin VB.CommandButton bcancel 
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
      Height          =   750
      Left            =   10620
      Picture         =   "frmFullBio.frx":6EA4
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   9540
      Width           =   1245
   End
   Begin VB.Frame Frame1 
      Caption         =   "Plot between"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2325
      Left            =   225
      TabIndex        =   8
      Top             =   7875
      Width           =   1785
      Begin VB.CommandButton cmdGo 
         BackColor       =   &H80000016&
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
         Height          =   675
         Left            =   360
         Picture         =   "frmFullBio.frx":71AE
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1560
         Width           =   1005
      End
      Begin VB.ComboBox cmbPlotFrom 
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
         Left            =   180
         TabIndex        =   10
         Text            =   "cmbPlotFrom"
         Top             =   540
         Width           =   1455
      End
      Begin VB.ComboBox cmbPlotTo 
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
         Left            =   150
         TabIndex        =   9
         Text            =   "cmbPlotTo"
         Top             =   1140
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   210
         TabIndex        =   17
         Top             =   270
         Width           =   495
      End
      Begin VB.Label T 
         Caption         =   "to"
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
         Left            =   210
         TabIndex        =   16
         Top             =   870
         Width           =   285
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   11370
      Top             =   1260
   End
   Begin VB.PictureBox pb 
      AutoRedraw      =   -1  'True
      Height          =   2325
      Left            =   2025
      ScaleHeight     =   2265
      ScaleWidth      =   4515
      TabIndex        =   6
      Top             =   7875
      Width           =   4575
   End
   Begin MSComctlLib.ProgressBar PBar 
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   1290
      Width           =   11625
      _ExtentX        =   20505
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSFlexGridLib.MSFlexGrid gDem 
      Height          =   2745
      Left            =   13050
      TabIndex        =   20
      Top             =   3900
      Visible         =   0   'False
      Width           =   4035
      _ExtentX        =   7117
      _ExtentY        =   4842
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      FormatString    =   "<SampleID    |<SampleDate   |<SampleTime |<Cnxn "
   End
   Begin Threed.SSPanel pnlFetching 
      Height          =   255
      Left            =   2880
      TabIndex        =   21
      Top             =   7170
      Visible         =   0   'False
      Width           =   6585
      _Version        =   65536
      _ExtentX        =   11615
      _ExtentY        =   450
      _StockProps     =   15
      ForeColor       =   -2147483630
      BackColor       =   65535
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FloodType       =   1
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   5595
      Left            =   210
      TabIndex        =   36
      Top             =   1530
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   9869
      _Version        =   393216
      Rows            =   5
      Cols            =   4
      FixedRows       =   4
      FixedCols       =   3
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      FormatString    =   "<Code|<Parameter |<Ref Ranges "
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "Results"
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
      Left            =   1950
      TabIndex        =   49
      Top             =   960
      Width           =   525
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "Show "
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
      Left            =   210
      TabIndex        =   47
      Top             =   960
      Width           =   450
   End
   Begin VB.Label Lbl1 
      Caption         =   "Amended  Results are underline"
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
      Left            =   9225
      TabIndex        =   46
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label lblExcelInfo 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exporting..."
      Height          =   285
      Left            =   9240
      TabIndex        =   38
      Top             =   9240
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "A/E"
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
      Left            =   10170
      TabIndex        =   35
      Top             =   60
      Width           =   285
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Sex"
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
      Left            =   8820
      TabIndex        =   30
      Top             =   60
      Width           =   270
   End
   Begin VB.Label lblSex 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   9120
      TabIndex        =   29
      Top             =   30
      Width           =   705
   End
   Begin VB.Label lblTest 
      Alignment       =   2  'Center
      Height          =   225
      Left            =   3180
      TabIndex        =   19
      Top             =   7635
      Width           =   2535
   End
   Begin VB.Label lblAandE 
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Left            =   10560
      TabIndex        =   18
      Top             =   37
      Width           =   1185
   End
   Begin VB.Label lblMaxVal 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6645
      TabIndex        =   14
      Top             =   7875
      Width           =   480
   End
   Begin VB.Label lblMeanVal 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6645
      TabIndex        =   13
      Top             =   8865
      Width           =   480
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6645
      TabIndex        =   12
      Top             =   9945
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
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
      Left            =   180
      TabIndex        =   5
      Top             =   60
      Width           =   420
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "DoB"
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
      Left            =   6750
      TabIndex        =   4
      Top             =   60
      Width           =   315
   End
   Begin VB.Label lblName 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   630
      TabIndex        =   3
      Top             =   30
      Width           =   3765
   End
   Begin VB.Label lblDoB 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   7080
      TabIndex        =   2
      Top             =   30
      Width           =   1545
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
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
      Left            =   4590
      TabIndex        =   1
      Top             =   60
      Width           =   375
   End
   Begin VB.Label lblChart 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   5010
      TabIndex        =   0
      Top             =   30
      Width           =   1545
   End
End
Attribute VB_Name = "frmFullBio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type ChartPosition
    xPos As Long
    yPos As Long
    Value As Single
    Date As String
End Type

Private ChartPositions() As ChartPosition

Private NumberOfDays As Long

Dim MaxPrevious As Integer



Private Sub FillResultGrid()

          Dim sql As String
          Dim intCols As Integer

10        On Error GoTo FillResultGrid_Error

20        intCols = gDem.Rows + 1
30        MaxPrevious = intCols - 3
40        udPrevious.Max = MaxPrevious
50        If MaxPrevious < 7 Then
60            lblPrevious = Format$(MaxPrevious)
70            udPrevious.Value = MaxPrevious
80        Else
90            lblPrevious = "6"
100           udPrevious.Value = 6
110       End If

120       g.Cols = intCols + 1

130       Exit Sub

FillResultGrid_Error:

          Dim strES As String
          Dim intEL As Integer

140       intEL = Erl
150       strES = Err.Description
160       LogError "frmFullBio", "FillResultGrid", intEL, strES, sql

End Sub

Private Sub FillCodeGrid()

          Dim tb As Recordset
          Dim sql As String
          Dim s As String

10        On Error GoTo FillCodeGrid_Error

20        With g
30            .Rows = 5
40            .AddItem ""
50            .RemoveItem 4
60        End With

          'Get Codes and ShortNames
70        sql = "SELECT D.PrintPriority, D.ShortName, D.Code " & _
                "FROM BioResults R, BioTestDefinitions D, Demographics G WHERE " & _
                "R.Code = D.Code " & _
                "AND R.SampleID = G.SampleID " & _
                "AND PatName = '" & AddTicks(lblName) & "' "

80        If lblChart <> "" And chkChartNumber.Value = 0 Then
90            sql = sql & "AND Chart = '" & lblChart & "' "
100       End If
110       If IsDate(lblDoB) Then
120           sql = sql & "and DoB = '" & Format(lblDoB, "dd/mmm/yyyy") & "' "
130       Else
140           sql = sql & "and (DoB is null or DoB = '') "
150       End If
160       sql = sql & "GROUP BY D.PrintPriority, D.ShortName, D.Code"

170       Set tb = New Recordset
180       RecOpenClient 0, tb, sql
190       Do While Not tb.EOF
200           s = tb!Code & vbTab & _
                  tb!ShortName & ""
210           g.AddItem s
220           tb.MoveNext
230       Loop

240       If g.Rows > 5 Then
250           g.RemoveItem 4
260       End If

270       Exit Sub

FillCodeGrid_Error:

          Dim strES As String
          Dim intEL As Integer

280       intEL = Erl
290       strES = Err.Description
300       LogError "frmFullBio", "FillCodeGrid", intEL, strES, sql

End Sub

Private Sub TransferDems()

          Dim gDemY As Integer
          Dim gResX As Integer

10        On Error GoTo TransferDems_Error

20        For gDemY = 1 To gDem.Rows - 1
30            gResX = gDemY + 2
              '  g.TextMatrix(0, gResX) = gDem.TextMatrix(gDemY, 3) 'Cnxn
40            g.TextMatrix(1, gResX) = gDem.TextMatrix(gDemY, 0)    'SampleID
50            g.TextMatrix(2, gResX) = gDem.TextMatrix(gDemY, 1)    'SampleDate
60            g.TextMatrix(3, gResX) = gDem.TextMatrix(gDemY, 2)    'SampleTime
70        Next

80        Exit Sub

TransferDems_Error:

          Dim strES As String
          Dim intEL As Integer

90        intEL = Erl
100       strES = Err.Description
110       LogError "frmFullBio", "TransferDems", intEL, strES

End Sub

Private Sub DrawChart()

          Dim n As Integer
          Dim Counter As Integer
          Dim DaysInterval As Long
          Dim x As Integer
          Dim Y As Integer
          Dim PixelsPerDay As Single
          Dim PixelsPerPointY As Single
          Dim FirstDayFilled As Boolean
          Dim MaxVal As Single
          Dim cVal As Single
          Dim StartGridX As Integer
          Dim StopGridX As Integer

10        On Error GoTo DrawChart_Error

20        MaxVal = 0
30        lblMaxVal = ""
40        lblMeanVal = ""
50        lblTest = ""

60        pb.Cls
70        pb.Picture = LoadPicture("")

80        NumberOfDays = DateDiff("d", Format(cmbPlotFrom, "dd/mmm/yyyy"), Format(cmbPlotTo, "dd/mmm/yyyy"))
90        If NumberOfDays < 1 Then Exit Sub
100       ReDim ChartPositions(0 To NumberOfDays)

110       For n = 1 To NumberOfDays
120           ChartPositions(n).xPos = 0
130           ChartPositions(n).yPos = 0
140           ChartPositions(n).Value = 0
150           ChartPositions(n).Date = ""
160       Next

170       For n = 2 To g.Cols - 1
180           If Format$(g.TextMatrix(2, n), "dd/mmm/yyyy") = cmbPlotTo Then StartGridX = n
190           If Format$(g.TextMatrix(2, n), "dd/mmm/yyyy") = cmbPlotFrom Then StopGridX = n
200       Next

210       FirstDayFilled = False
220       Counter = 0
230       For x = StartGridX To StopGridX
240           If g.TextMatrix(g.Row, x) <> "" Then
250               If Not FirstDayFilled Then
260                   FirstDayFilled = True
270                   MaxVal = Val(g.TextMatrix(g.Row, x))
280                   ChartPositions(NumberOfDays).Date = Format(g.TextMatrix(2, x), "dd/mmm/yyyy")
290                   ChartPositions(NumberOfDays).Value = Val(g.TextMatrix(g.Row, x))
300               Else
310                   DaysInterval = Abs(DateDiff("D", cmbPlotTo, Format$(g.TextMatrix(2, x), "dd/mmm/yyyy")))
320                   ChartPositions(NumberOfDays - DaysInterval).Date = g.TextMatrix(2, x)
330                   cVal = Val(g.TextMatrix(g.Row, x))
340                   ChartPositions(NumberOfDays - DaysInterval).Value = cVal
350                   If cVal > MaxVal Then MaxVal = cVal
360               End If
370           End If
380       Next

390       PixelsPerDay = (pb.Width - 1060) / NumberOfDays
400       MaxVal = MaxVal * 1.1
410       If MaxVal = 0 Then Exit Sub
420       PixelsPerPointY = pb.Height / MaxVal

430       x = 580 + (NumberOfDays * PixelsPerDay)
440       Y = pb.Height - (ChartPositions(NumberOfDays).Value * PixelsPerPointY)
450       ChartPositions(NumberOfDays).yPos = Y
460       ChartPositions(NumberOfDays).xPos = x

470       pb.ForeColor = vbBlue
480       pb.Circle (x, Y), 30
490       pb.Line (x - 15, Y - 15)-(x + 15, Y + 15), vbBlue, BF
500       pb.PSet (x, Y)

510       For n = NumberOfDays - 1 To 0 Step -1
520           If ChartPositions(n).Value <> 0 And ChartPositions(n).Date <> "" Then
530               DaysInterval = Abs(DateDiff("d", cmbPlotFrom, Format(ChartPositions(n).Date, "dd/mmm/yyyy")))
540               x = 580 + (DaysInterval * PixelsPerDay)
550               ChartPositions(n).xPos = x
560               Y = pb.Height - (ChartPositions(n).Value * PixelsPerPointY)
570               ChartPositions(n).yPos = Y
580               pb.Line -(x, Y)
590               pb.Line (x - 15, Y - 15)-(x + 15, Y + 15), vbBlue, BF
600               pb.Circle (x, Y), 30
610               pb.PSet (x, Y)
620           End If
630       Next

640       pb.Line (0, pb.Height / 2)-(pb.Width, pb.Height / 2), vbBlack, BF

650       lblMaxVal = Format$(Int(MaxVal + 0.5), "###0")
660       lblMeanVal = Format$(Val(lblMaxVal) / 2, "###0.0")

670       lblTest = g.TextMatrix(g.Row, 1)

680       Exit Sub

DrawChart_Error:

          Dim strES As String
          Dim intEL As Integer

690       intEL = Erl
700       strES = Err.Description
710       LogError "frmFullBio", "DrawChart", intEL, strES

End Sub

Private Sub FillCombos()

          Dim x As Integer
          Dim Px As Printer

10        On Error GoTo FillCombos_Error

20        cmbPlotFrom.Clear
30        cmbPlotTo.Clear

40        For x = 3 To g.Cols - 1
50            cmbPlotFrom.AddItem Format$(g.TextMatrix(2, x), "dd/mmm/yyyy")
60            cmbPlotTo.AddItem Format$(g.TextMatrix(2, x), "dd/mmm/yyyy")
70        Next

80        cmbPlotTo = Format$(g.TextMatrix(2, 3), "dd/mmm/yyyy")
90        If Not IsDate(cmbPlotTo) Then Exit Sub

100       For x = g.Cols - 1 To 3 Step -1
110           If DateDiff("d", Format$(g.TextMatrix(2, x), "dd/mmm/yyyy"), cmbPlotTo) < 365 Then
120               cmbPlotFrom = Format$(g.TextMatrix(2, x), "dd/mmm/yyyy")
130               Exit For
140           End If
150       Next

160       cmbPrinter.Clear
170       For Each Px In Printers
180           cmbPrinter.AddItem Px.DeviceName
190       Next
200       For x = 0 To cmbPrinter.ListCount - 1
210           If Printer.DeviceName = cmbPrinter.List(x) Then
220               cmbPrinter.ListIndex = x
230               Exit For
240           End If
250       Next

260       Exit Sub

FillCombos_Error:

          Dim strES As String
          Dim intEL As Integer

270       intEL = Erl
280       strES = Err.Description
290       LogError "frmFullBio", "FillCombos", intEL, strES

End Sub

Private Sub bcancel_Click()

10        Unload Me

End Sub


Private Sub chkChartNumber_Click()
          If cmbResultCount.Text <> "" Then
10            If FillgDem(Trim(cmbResultCount.Text)) Then
20                FillResultGrid
30                FillCodeGrid
40                TransferDems
50                FillAllResults
60            End If
          End If
End Sub

Private Sub cmbPlotFrom_KeyPress(KeyAscii As Integer)

10        KeyAscii = 0

End Sub


Private Sub cmbPlotTo_KeyPress(KeyAscii As Integer)

10        KeyAscii = 0

End Sub


Private Sub cmbResultCount_Change()
    On Error GoTo cmbResultCount_Change_Error
    
    pnlFetching.Visible = False
    g.Visible = False
     
    If cmbResultCount.Text <> "" Then
        If FillgDem(Trim(cmbResultCount.Text)) Then
            FillResultGrid
            FillCodeGrid
            TransferDems
            FillAllResults
        End If
    End If
    
    g.Visible = True
    pnlFetching.Visible = True
    
cmbResultCount_Change_Error:

    Dim strES As String
    Dim intEL As Integer
    
    intEL = Erl
    strES = Err.Description
    LogError "frmFullBio", "cmbResultCount_Change", intEL, strES
End Sub

Private Sub cmbResultCount_Click()
    On Error GoTo cmbResultCount_Click_Error
    
    pnlFetching.Visible = False
    g.Visible = False
     
    If cmbResultCount.Text <> "" Then
        If FillgDem(Trim(cmbResultCount.Text)) Then
            FillResultGrid
            FillCodeGrid
            TransferDems
            FillAllResults
        End If
    End If
    
    g.Visible = True
    pnlFetching.Visible = True
    
cmbResultCount_Click_Error:

    Dim strES As String
    Dim intEL As Integer
    
    intEL = Erl
    strES = Err.Description
    LogError "frmFullBio", "cmbResultCount_Click", intEL, strES
End Sub

Private Sub cmdExcel_Click()
          Dim strHeading As String
10        On Error GoTo cmdExcel_Click_Error

20        strHeading = "Biochemistry History" & vbCr
30        strHeading = strHeading & "Patient Name: " & lblName & vbCr
40        strHeading = strHeading & vbCr

50        ExportFlexGrid g, Me, strHeading

60        Exit Sub

cmdExcel_Click_Error:

          Dim strES As String
          Dim intEL As Integer

70        intEL = Erl
80        strES = Err.Description
90        LogError "frmFullBio", "cmdExcel_Click", intEL, strES
End Sub

Private Sub cmdGo_Click()

10        DrawChart

End Sub

Private Sub cmdPrint_Click()

          Dim Px As Printer
          Dim x As Integer
          Dim Y As Integer
          Dim z As Integer
          Dim PageCounter As Integer
          Dim CurrentPage As Integer
          Dim TabPos As Integer
          Dim ArrayStart As Integer
          Dim ArrayStop As Integer
          Dim MaxCols As Integer
          Dim TotalLines As Integer
          Dim LinesToPrint As Integer
          Dim Start As Integer
          Dim Last As Integer

10        On Error GoTo cmdPrint_Click_Error

20        For Each Px In Printers
30            If Px.DeviceName = cmbPrinter.Text Then
40                Set Printer = Px
50                Exit For
60            End If
70        Next

80        Printer.Font = "Courier New"
90        Printer.Orientation = vbPRORLandscape


100       TotalLines = 27
110       LinesToPrint = g.Rows - 4
120       Start = 4
130       PageCounter = ((g.Rows - 4) \ TotalLines)
140       If (g.Rows - 4) Mod TotalLines > 0 Then PageCounter = PageCounter + 1
150       CurrentPage = 1

160       If LinesToPrint > TotalLines Then
170           Last = TotalLines
180           LinesToPrint = LinesToPrint - TotalLines
190       Else
200           Last = LinesToPrint + 3
210           LinesToPrint = 0
220       End If

230       Do While CurrentPage <= PageCounter
240           Printer.Print
250           PrintText FormatString("Cumulative Biochemistry Report", 70, , AlignCenter), 14, True, , , , True
260           PrintText FormatString("Page " & Format$(CurrentPage) & " of " & PageCounter, 100, , AlignCenter), 10, , , , , True

270           PrintText "  Patient Name: " & lblName, 14, True, , , , True
280           PrintText " Date of Birth: " & Format$(lblDoB, "dd/mm/yyyy"), 14, True, , , , True
290           PrintText "         Chart: " & lblChart, 14, True, , , , True

300           Printer.Print

310           With g
320               If .Cols > 9 Then
330                   MaxCols = 8
340               Else
350                   MaxCols = .Cols - 2
360               End If

                  'Add seperator
370               PrintText String(217, "-") & vbCrLf, 4, True
                  'Print SampleID row
380               PrintText FormatString(.TextMatrix(1, 1), 16, "|"), 10, True
390               For z = 2 To MaxCols
400                   PrintText FormatString(.TextMatrix(1, z), 9, "|", AlignCenter), 10
410               Next z
420               PrintText vbCrLf
                  'Print Sample Date Row
430               PrintText FormatString(.TextMatrix(2, 1), 16, "|"), 10, True
440               For z = 2 To MaxCols
450                   PrintText FormatString(Format(.TextMatrix(2, z), "dd/MM/yy"), 9, "|", AlignCenter), 10
460               Next z
470               PrintText vbCrLf
                  'Print Sample Time Row is exists

480               PrintText FormatString("SAMPLE TIME", 16, "|"), 10, True
490               For z = 2 To MaxCols
500                   If Format(.TextMatrix(2, z), "hh:mm") = "00:00" Then
510                       PrintText FormatString("", 9, "|", AlignCenter), 10
520                   Else
530                       PrintText FormatString(Format(.TextMatrix(2, z), "hh:mm"), 9, "|", AlignCenter), 10
540                   End If
550               Next z
560               PrintText vbCrLf

                  'Print Run Date Row
570               PrintText FormatString(.TextMatrix(3, 1), 16, "|"), 10, True
580               For z = 2 To MaxCols
590                   PrintText FormatString(Format(.TextMatrix(3, z), "dd/MM/yy"), 9, "|", AlignCenter), 10
600               Next z
610               PrintText vbCrLf
                  'Print Run Time Row is exists

620               PrintText FormatString("RUN TIME", 16, "|"), 10, True
630               PrintText FormatString("Ref Range", 9, "|"), 10
640               For z = 3 To MaxCols
650                   If Format(.TextMatrix(3, z), "hh:mm") = "00:00" Then
660                       PrintText FormatString("", 9, "|", AlignCenter), 10
670                   Else
680                       PrintText FormatString(Format(.TextMatrix(3, z), "hh:mm"), 9, "|", AlignCenter), 10
690                   End If
700               Next z
710               PrintText vbCrLf
                  'Add seperator
720               PrintText String(217, "-") & vbCrLf, 4, True
                  'Print results
730               For Y = Start To Last
740                   PrintText FormatString(.TextMatrix(Y, 1), 16, "|"), 10
750                   For z = 2 To MaxCols
760                       PrintText FormatString(.TextMatrix(Y, z), 9, "|", IIf(z = 2, AlignLeft, AlignCenter)), 10
770                   Next z
780                   PrintText vbCrLf
790               Next Y

800           End With


              'End of Page Line
810           PrintText String(217, "-"), 4, True
820           If CurrentPage < PageCounter Then Printer.NewPage
830           CurrentPage = CurrentPage + 1
840           Start = Last + 1
850           If LinesToPrint > TotalLines Then
860               Last = Last + TotalLines
870               LinesToPrint = LinesToPrint - TotalLines
880           Else
890               Last = Last + LinesToPrint + 3
900               LinesToPrint = 0
910           End If

920       Loop

930       Printer.EndDoc



          '
          '        For y = 1 To g.Rows - 1
          '
          '            PrintText FormatString(g.TextMatrix(y, 1), 11, "|", AlignLeft), 10, IIf(y = 1 Or y = 2 Or y = 3, True, False)
          '
          '            ArrayStart = ((CurrentPage - 1) * 7) + 3
          '            ArrayStop = ArrayStart + 6
          '
          '            'MaxCol = ((CurrentPage - 1) * 7) + 7
          '            If ArrayStop > Val(lblPrevious) + 3 Then
          '                ArrayStop = Val(lblPrevious) + 3
          '            End If
          '
          '            For x = ArrayStart To ArrayStop
          '
          '                PrintText FormatString(g.TextMatrix(y, x), 14, "|", AlignCenter), 10
          '            Next
          '
          '            Printer.Print
          '
          '        Next



940       Exit Sub

cmdPrint_Click_Error:

          Dim strES As String
          Dim intEL As Integer

950       intEL = Erl
960       strES = Err.Description
970       LogError "frmFullBio", "cmdPrint_Click", intEL, strES

End Sub

Private Sub cmdRefresh_Click()

10    On Error GoTo cmdRefresh_Click_Error

20    If dtFrom.Value > dtTo.Value Then
30        iMsg "From cannot be greater than To date", vbInformation
40        Exit Sub
50    End If

      'Initiate
60    g.Rows = 4
70    g.Row = 0
80    g.Cols = 3
90    g.Col = 0


100   pnlFetching.Visible = False
110   g.Visible = False

120   If cmbResultCount.Text <> "" Then
          If FillgDem(Trim(cmbResultCount.Text)) Then
130           FillResultGrid
140           FillCodeGrid
150           TransferDems
160           FillAllResults
170       End If
      End If

180   g.Visible = True
190   pnlFetching.Visible = True

200   FillCombos

210   Exit Sub

cmdRefresh_Click_Error:

      Dim strES As String
      Dim intEL As Integer

220   intEL = Erl
230   strES = Err.Description
240   LogError "frmFullBio", "cmdRefresh_Click", intEL, strES


End Sub

Private Sub Form_Activate()

10        On Error GoTo Form_Activate_Error

20        If GetOptionSetting("EnableiPMSChart", "0") = 0 Then
30            chkChartNumber.Value = 0
40            chkChartNumber.Enabled = Not (lblChart = "")
50        Else
60            chkChartNumber.Value = 1

70        End If

80        pnlFetching.Visible = False
90        g.Visible = False
            
          If cmbResultCount.Text <> "" Then
100           If FillgDem(Trim(cmbResultCount.Text)) Then
110               FillResultGrid
120               FillCodeGrid
130               TransferDems
140               FillAllResults
150           End If
          End If
          
160       g.Visible = True
170       pnlFetching.Visible = True

180       FillCombos


190       PBar.Max = LogOffDelaySecs
200       PBar = 0

210       Timer1.Enabled = True

220       Exit Sub

Form_Activate_Error:

          Dim strES As String
          Dim intEL As Integer

230       intEL = Erl
240       strES = Err.Description
250       LogError "frmFullBio", "Form_Activate", intEL, strES

End Sub
Private Sub FillAllResults()

      Dim tb As Recordset
      Dim sql As String
      Dim x As Integer
      Dim Y As Integer
      Dim SampleID As Long
      Dim strSex As String
      Dim PerCent As Integer
      Dim DaysOld As Long

10    On Error GoTo FillAllResults_Error

20    If Left$(lblSex, 1) = "M" Then
30        strSex = "CT.MaleLow as Low, CT.MaleHigh as High "
40    ElseIf Left$(lblSex, 1) = "F" Then
50        strSex = "CT.FemaleLow as Low, CT.FemaleHigh as High "
60    Else
70        strSex = "CT.FemaleLow as Low, CT.MaleHigh as High "
80    End If

90    PerCent = 0

100   For x = 3 To g.Cols - 1
110       g.ColWidth(x) = 1500
120       If x = 11 Then
130           g.Visible = True
140           pnlFetching.Visible = True
150       End If

160       PerCent = (x / g.Cols) * 100
170       pnlFetching.FloodPercent = PerCent

          '180     g.TextMatrix(0, x) = ""

180       SampleID = Val(g.TextMatrix(1, x))

190       DaysOld = 365 * 20
200       If IsDate(lblDoB) Then
210           DaysOld = DateDiff("d", lblDoB, g.TextMatrix(2, x))
220       End If

230       sql = "SELECT " & _
                "CASE WHEN ISNUMERIC(R.Result) = 1 AND R.Result <> '+' AND R.Result <> '.' THEN " & _
                "  STR(CONVERT(FLOAT,R.Result) + 0.00000001, 6, CT.DP)  " & _
                "ELSE " & _
                "  R.Result " & _
                "END Result, R.Valid, CT.ShortName, R.Code, " & _
                "CT.PlausibleLow, CT.PlausibleHigh, " & _
                strSex & _
                "FROM BioResults R, BioTestDefinitions CT " & _
                "WHERE SampleID = '" & SampleID & "' " & _
                "AND R.Code = CT.Code " & _
                "AND AgeFromDays <= '" & DaysOld & "' " & _
                "AND AgeToDays >= '" & DaysOld & "'"

240       Set tb = New Recordset
250       RecOpenClient 0, tb, sql
260       Do While Not tb.EOF
              'Zyam
              If InStr(1, ">", tb!Result & "") Then
                tb!Result = Right(tb!Result, Len(tb!Result) - 1)
              End If
              'Zyam
270           For Y = 4 To g.Rows - 1
280               If tb!Code = g.TextMatrix(Y, 0) Then
290                   If g.TextMatrix(Y, 2) = "" Then
300                       g.TextMatrix(Y, 2) = tb!Low & "-" & tb!High
310                   End If
320                   If tb!Valid Or UCase(UserMemberOf) = "MANAGERS" Or UCase(UserMemberOf) = "USERS" Then     'QMS Ref #818192
                          '----------------------------
                          If IsResultAmended("Bio", SampleID, tb!Code, tb!Result) = True Then
                              g.Col = x
                              g.Row = Y
                              g.CellFontUnderline = True
                          End If
                          '============================
330                       g.TextMatrix(Y, x) = tb!Result
                          '400   g.CellFontItalic = False
340                       If IsNumeric(tb!Result) Then
350                           If Val(tb!Result) > tb!PlausibleHigh Then
360                               g.Col = x
370                               g.Row = Y
380                               g.CellBackColor = vbBlue
390                               g.CellForeColor = vbYellow
400                               g.TextMatrix(Y, x) = "*****"
410                           ElseIf Val(tb!Result) < tb!PlausibleLow Then
420                               g.Col = x
430                               g.Row = Y
440                               g.CellBackColor = vbBlack
450                               g.CellForeColor = vbYellow
460                               g.TextMatrix(Y, x) = "*****"
470                           ElseIf Val(tb!Result) > tb!High Then
480                               fraHL.Visible = True
490                               g.Col = x
500                               g.Row = Y
510                               g.CellBackColor = vbRed
520                               g.CellForeColor = vbYellow
530                           ElseIf Val(tb!Result) < tb!Low Then
540                               fraHL.Visible = True
550                               g.Col = x
560                               g.Row = Y
570                               g.CellBackColor = vbBlue
580                               g.CellForeColor = vbYellow
590                           End If
600                       End If
610                       If tb!Valid = 0 Then g.TextMatrix(0, x) = "NV"
620                   Else
630                       g.TextMatrix(Y, x) = "NV"

640                   End If
650                   Exit For
660               End If
670           Next
680           tb.MoveNext
690       Loop
700   Next

710   Exit Sub

FillAllResults_Error:

      Dim strES As String
      Dim intEL As Integer

720   intEL = Erl
730   strES = Err.Description
740   LogError "frmFullBio", "FillAllResults", intEL, strES, sql

End Sub


Private Function FillgDem(p_RCount As String) As Boolean

          Dim n As Integer
          Dim tb As Recordset
          Dim sql As String
          Dim s As String
          Dim xdate As String

10        On Error GoTo FillgDem_Error

20        With gDem
30            .Visible = False
40            .Rows = 2
50            .AddItem ""
60            .RemoveItem 1
70        End With
80        p_RCount = UCase(p_RCount)
90        If p_RCount = "First 5" Then
100           sql = "SELECT DISTINCT top 5 D.SampleID, D.SampleDate, D.RunDate, " & _
                  "(SELECT Top 1 Q.RunTime FROM BioResults Q WHERE Q.SampleID = D.SampleID ORDER BY Q.RunTime) RunDateTime " & _
                  "FROM Demographics D INNER JOIN BioResults R " & _
                  "ON D.SampleID = R.SampleID WHERE "
110       ElseIf p_RCount = "First 10" Then
120           sql = "SELECT DISTINCT top 10 D.SampleID, D.SampleDate, D.RunDate, " & _
                  "(SELECT Top 1 Q.RunTime FROM BioResults Q WHERE Q.SampleID = D.SampleID ORDER BY Q.RunTime) RunDateTime " & _
                  "FROM Demographics D INNER JOIN BioResults R " & _
                  "ON D.SampleID = R.SampleID WHERE "
130       ElseIf p_RCount = "First 20" Then
140           sql = "SELECT DISTINCT top 20 D.SampleID, D.SampleDate, D.RunDate, " & _
                  "(SELECT Top 1 Q.RunTime FROM BioResults Q WHERE Q.SampleID = D.SampleID ORDER BY Q.RunTime) RunDateTime " & _
                  "FROM Demographics D INNER JOIN BioResults R " & _
                  "ON D.SampleID = R.SampleID WHERE "
150       ElseIf p_RCount = "First 50" Then
160           sql = "SELECT DISTINCT top 50 D.SampleID, D.SampleDate, D.RunDate, " & _
                  "(SELECT Top 1 Q.RunTime FROM BioResults Q WHERE Q.SampleID = D.SampleID ORDER BY Q.RunTime) RunDateTime " & _
                  "FROM Demographics D INNER JOIN BioResults R " & _
                  "ON D.SampleID = R.SampleID WHERE "
170       ElseIf p_RCount = "ALL" Then
180           sql = "SELECT DISTINCT D.SampleID, D.SampleDate, D.RunDate, " & _
                  "(SELECT Top 1 Q.RunTime FROM BioResults Q WHERE Q.SampleID = D.SampleID ORDER BY Q.RunTime) RunDateTime " & _
                  "FROM Demographics D INNER JOIN BioResults R " & _
                  "ON D.SampleID = R.SampleID WHERE "
190       End If
200       If Trim(lblChart) <> "" And chkChartNumber.Value = 0 Then
210           sql = sql & "(D.Chart = '" & lblChart & "') AND"
220       End If
230       sql = sql & " PatName = '" & AddTicks(lblName) & "' " & _
              "AND D.SampleDate BETWEEN '" & Format(dtFrom, "dd/MMM/yyyy") & "' AND '" & Format(dtTo.Value + 1, "dd/MMM/yyyy") & "' "

          ' "AND D.SampleDate Between getdate()-90 And getdate()+1 "              'QMS Ref #818126
240       If IsDate(lblDoB) Then
250           sql = sql & "AND DoB = '" & Format(lblDoB, "dd/mmm/yyyy") & "' "
260       Else
270           sql = sql & "AND (DoB IS NULL OR DoB = '') "
280       End If
          '+++ Junaid 10-08-2023
          '180       sql = sql & "AND R.SampleID = D.SampleID ORDER BY RunDateTime DESC"
290       sql = sql & "AND R.SampleID = D.SampleID ORDER BY SampleDate DESC"
          '--- Junaid
300       Set tb = New Recordset
310       RecOpenClient n, tb, sql
320       Do While Not tb.EOF
330           s = tb!SampleID & vbTab
              'Sample DateTime
340           If Not IsNull(tb!SampleDate) Then
350               xdate = Format(tb!SampleDate, "dd/mm/yy")
360               If Format(tb!SampleDate, "hh:mm") <> "00:00" Then
370                   xdate = xdate & " " & Format(tb!SampleDate, "hh:mm")
380               End If
390           Else
400               xdate = ""
410           End If
420           s = s & xdate & vbTab
              'Run DateTime
430           If tb!RunDateTime <> "" Then
440               xdate = Format(tb!RunDateTime, "dd/mm/yy")
450               If Format(tb!RunDateTime, "hh:mm") <> "00:00" Then
460                   xdate = xdate & " " & Format(tb!RunDateTime, "hh:mm")
470               End If
480           Else
490               xdate = ""
500           End If
510           s = s & xdate & vbTab
520           s = s & vbTab & Format$(n)

              '    If Format(tb!SampleDate, "HH:mm") <> "00:00" Then
              '        s = s & Format$(tb!SampleDate, "HH:mm")
              '    End If

530           gDem.AddItem s
540           tb.MoveNext
550       Loop

560       With gDem
570           If .Rows > 2 Then
580               .RemoveItem 1
590               .Visible = True
600               FillgDem = True
610           Else
620               FillgDem = False
630           End If
640       End With

650       Exit Function

FillgDem_Error:

          Dim strES As String
          Dim intEL As Integer

660       intEL = Erl
670       strES = Err.Description
680       LogError "frmFullBio", "FillgDem", intEL, strES, sql

End Function

Private Sub Form_Deactivate()

10        On Error GoTo Form_Deactivate_Error

20        Timer1.Enabled = False

30        Exit Sub

Form_Deactivate_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmFullBio", "Form_Deactivate", intEL, strES

End Sub

Private Sub Form_Load()

10        On Error GoTo Form_Load_Error

20        g.ColWidth(0) = 0
30        g.ColWidth(1) = 1600
40        g.TextMatrix(1, 1) = "SAMPLE ID"
50        g.TextMatrix(2, 1) = "SAMPLE DATE"
60        g.TextMatrix(3, 1) = "RUN DATE"
70        dtFrom.Value = Now - 90
80        dtTo.Value = Now

90        Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

100       intEL = Erl
110       strES = Err.Description
120       LogError "frmFullBio", "Form_Load", intEL, strES

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

10        On Error GoTo Form_MouseMove_Error

20        PBar = 0

30        Exit Sub

Form_MouseMove_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmFullBio", "Form_MouseMove", intEL, strES


End Sub

Private Sub g_Click()

10        On Error GoTo g_Click_Error

20        DrawChart

30        Exit Sub

g_Click_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmFullBio", "g_Click", intEL, strES

End Sub


Private Sub g_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

10        On Error GoTo g_MouseDown_Error

20        If Button = vbRightButton Then
30            g.ColWidth(g.Col) = 1500
40        End If

50        Exit Sub

g_MouseDown_Error:

          Dim strES As String
          Dim intEL As Integer

60        intEL = Erl
70        strES = Err.Description
80        LogError "frmFullBio", "g_MouseDown", intEL, strES

End Sub

Private Sub g_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

10        On Error GoTo g_MouseMove_Error

20        PBar = 0

30        Exit Sub

g_MouseMove_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmFullBio", "g_MouseMove", intEL, strES

End Sub


Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

10        On Error GoTo Label1_MouseMove_Error

20        PBar = 0

30        Exit Sub

Label1_MouseMove_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmFullBio", "Label1_MouseMove", intEL, strES


End Sub


Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

10        On Error GoTo Label2_MouseMove_Error

20        PBar = 0

30        Exit Sub

Label2_MouseMove_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmFullBio", "Label2_MouseMove", intEL, strES


End Sub


Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

10        On Error GoTo Label3_MouseMove_Error

20        PBar = 0

30        Exit Sub

Label3_MouseMove_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmFullBio", "Label3_MouseMove", intEL, strES


End Sub


Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

10        On Error GoTo Label4_MouseMove_Error

20        PBar = 0

30        Exit Sub

Label4_MouseMove_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmFullBio", "Label4_MouseMove", intEL, strES


End Sub


Private Sub lblChart_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

10        On Error GoTo lblChart_MouseMove_Error

20        PBar = 0

30        Exit Sub

lblChart_MouseMove_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmFullBio", "lblChart_MouseMove", intEL, strES


End Sub


Private Sub lblDoB_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

10        On Error GoTo lblDoB_MouseMove_Error

20        PBar = 0

30        Exit Sub

lblDoB_MouseMove_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmFullBio", "lblDoB_MouseMove", intEL, strES


End Sub


Private Sub lblName_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

10        On Error GoTo lblName_MouseMove_Error

20        PBar = 0

30        Exit Sub

lblName_MouseMove_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmFullBio", "lblName_MouseMove", intEL, strES


End Sub


Private Sub pb_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

          Dim i As Long
          Dim CurrentDistance As Long
          Dim BestDistance As Long
          Dim BestIndex As Integer

10        On Error GoTo pb_MouseMove_Error

20        PBar = 0

30        If NumberOfDays = 0 Then Exit Sub

40        BestIndex = -1
50        BestDistance = 99999
60        For i = 0 To NumberOfDays
70            CurrentDistance = ((x - ChartPositions(i).xPos) ^ 2 + (Y - ChartPositions(i).yPos) ^ 2) ^ (1 / 2)
80            If i = 0 Or CurrentDistance < BestDistance Then
90                BestDistance = CurrentDistance
100               BestIndex = i
110           End If
120       Next

130       If BestIndex <> -1 Then
140           pb.ToolTipText = Format$(ChartPositions(BestIndex).Date, "dd/mmm/yyyy") & " " & ChartPositions(BestIndex).Value
150       End If

160       Exit Sub

pb_MouseMove_Error:

          Dim strES As String
          Dim intEL As Integer

170       intEL = Erl
180       strES = Err.Description
190       LogError "frmFullBio", "pb_MouseMove", intEL, strES

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
90        LogError "frmFullBio", "Timer1_Timer", intEL, strES

End Sub

