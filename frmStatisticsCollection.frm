VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmStatisticsCollection 
   Caption         =   "Satistics Collection"
   ClientHeight    =   12975
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19545
   LinkTopic       =   "Form1"
   ScaleHeight     =   12975
   ScaleWidth      =   19545
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.ListBox lstHours 
      Height          =   450
      ItemData        =   "frmStatisticsCollection.frx":0000
      Left            =   7440
      List            =   "frmStatisticsCollection.frx":000D
      TabIndex        =   55
      Top             =   12000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CheckBox chkClinician 
      Caption         =   "Clinician"
      Height          =   195
      Left            =   12420
      TabIndex        =   49
      Top             =   1020
      Width           =   975
   End
   Begin VB.CheckBox chkWard 
      Caption         =   "Ward"
      Height          =   195
      Left            =   16560
      TabIndex        =   48
      Top             =   1020
      Width           =   735
   End
   Begin VB.CheckBox chkGP 
      Caption         =   "GP"
      Height          =   195
      Left            =   8640
      TabIndex        =   47
      Top             =   1020
      Width           =   615
   End
   Begin VB.Frame fraMainTab1 
      BackColor       =   &H00404000&
      BorderStyle     =   0  'None
      Height          =   12420
      Left            =   75
      TabIndex        =   1
      Top             =   420
      Width           =   19290
      Begin VB.Frame Frame1 
         Caption         =   "Tests Selected"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   2820
         Left            =   19200
         TabIndex        =   50
         Top             =   1920
         Visible         =   0   'False
         Width           =   915
         Begin VB.ListBox lstTestSelected 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   2220
            ItemData        =   "frmStatisticsCollection.frx":001B
            Left            =   240
            List            =   "frmStatisticsCollection.frx":001D
            Style           =   1  'Checkbox
            TabIndex        =   51
            Top             =   330
            Width           =   5025
         End
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         Caption         =   "Total Results"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   900
         Left            =   240
         TabIndex        =   36
         Top             =   11280
         Width           =   7050
         Begin VB.ComboBox cmbHours 
            Height          =   315
            ItemData        =   "frmStatisticsCollection.frx":001F
            Left            =   4200
            List            =   "frmStatisticsCollection.frx":0032
            TabIndex        =   54
            Text            =   "01"
            Top             =   465
            Width           =   630
         End
         Begin VB.TextBox txtOver1hr 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   5280
            TabIndex        =   52
            Top             =   345
            Width           =   1455
         End
         Begin VB.CommandButton Command 
            Height          =   360
            Index           =   2
            Left            =   6690
            TabIndex        =   38
            Top             =   240
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.TextBox txtResultTotal 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   1620
            TabIndex        =   37
            Top             =   345
            Width           =   1455
         End
         Begin VB.Label Label3 
            Caption         =   "hr"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   4850
            TabIndex        =   56
            Top             =   480
            Width           =   375
         End
         Begin VB.Label Label2 
            Caption         =   "TAT Over"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   3350
            TabIndex        =   53
            Top             =   465
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Result Total"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   540
            TabIndex        =   39
            Top             =   465
            Width           =   1095
         End
      End
      Begin VB.CommandButton cmdExit 
         Height          =   900
         Left            =   13185
         Picture         =   "frmStatisticsCollection.frx":004A
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   11280
         Width           =   1210
      End
      Begin VB.CommandButton cmdExcel 
         Caption         =   "Export to Excel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   9960
         Picture         =   "frmStatisticsCollection.frx":0F14
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   11280
         Width           =   1210
      End
      Begin VB.Frame fraResultType 
         Appearance      =   0  'Flat
         Caption         =   "Refine Filters"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   1620
         Left            =   6840
         TabIndex        =   20
         Top             =   180
         Width           =   12210
         Begin VB.ComboBox cmbClinician 
            Height          =   315
            ItemData        =   "frmStatisticsCollection.frx":121E
            Left            =   4080
            List            =   "frmStatisticsCollection.frx":1220
            TabIndex        =   42
            Top             =   720
            Width           =   3975
         End
         Begin VB.ComboBox cmbWard 
            Height          =   315
            Left            =   8040
            TabIndex        =   41
            Top             =   720
            Width           =   3975
         End
         Begin VB.ComboBox cmbGp 
            Height          =   315
            Left            =   120
            TabIndex        =   40
            Top             =   720
            Width           =   3975
         End
         Begin VB.CommandButton cmdCheck 
            Caption         =   "Check"
            Height          =   360
            Index           =   1
            Left            =   11160
            TabIndex        =   21
            Top             =   1200
            Visible         =   0   'False
            Width           =   885
         End
      End
      Begin VB.Frame fraDates 
         Caption         =   "Between Dates"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   1620
         Left            =   255
         TabIndex        =   9
         Top             =   180
         Width           =   6450
         Begin VB.CommandButton cmdRecalc 
            Caption         =   "Start"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   825
            Left            =   5160
            MaskColor       =   &H8000000F&
            Picture         =   "frmStatisticsCollection.frx":1222
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   300
            Width           =   1080
         End
         Begin VB.OptionButton optBetween 
            Caption         =   "Last 30 Days"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   1545
            TabIndex        =   16
            Top             =   750
            Width           =   1455
         End
         Begin VB.OptionButton optBetween 
            Caption         =   "Last Full Month"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   1545
            TabIndex        =   15
            Top             =   1125
            Width           =   1635
         End
         Begin VB.OptionButton optBetween 
            Caption         =   "Year To Date"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   5
            Left            =   3405
            TabIndex        =   14
            Top             =   390
            Width           =   1455
         End
         Begin VB.OptionButton optBetween 
            Caption         =   "Today"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   6
            Left            =   180
            TabIndex        =   13
            Top             =   750
            Width           =   1095
         End
         Begin VB.OptionButton optBetween 
            Caption         =   "Last Week"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   180
            TabIndex        =   12
            Top             =   1125
            Width           =   1275
         End
         Begin VB.OptionButton optBetween 
            Caption         =   "Last Quarter"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   3
            Left            =   3405
            TabIndex        =   11
            Top             =   750
            Width           =   1395
         End
         Begin VB.OptionButton optBetween 
            Caption         =   "Last Full Quarter"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   4
            Left            =   3405
            TabIndex        =   10
            Top             =   1125
            Width           =   1755
         End
         Begin MSComCtl2.DTPicker dtFrom 
            Height          =   315
            Left            =   180
            TabIndex        =   18
            Top             =   360
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   126025729
            CurrentDate     =   40969
         End
         Begin MSComCtl2.DTPicker dtTo 
            Height          =   315
            Left            =   1830
            TabIndex        =   19
            Top             =   360
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   126025729
            CurrentDate     =   40976
         End
      End
      Begin VB.TextBox txtTest 
         Height          =   285
         Left            =   4680
         TabIndex        =   8
         Top             =   11880
         Width           =   2535
      End
      Begin VB.Timer tmrUp 
         Enabled         =   0   'False
         Interval        =   200
         Left            =   9225
         Top             =   11625
      End
      Begin VB.Timer tmrDown 
         Enabled         =   0   'False
         Interval        =   200
         Left            =   8790
         Top             =   11625
      End
      Begin VB.Frame fraOrganismGroup 
         Caption         =   "Tests"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   2820
         Left            =   2670
         TabIndex        =   6
         Top             =   1890
         Width           =   16380
         Begin VB.ListBox lstDefinitions 
            BackColor       =   &H00C0FFFF&
            Columns         =   4
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   2220
            ItemData        =   "frmStatisticsCollection.frx":152C
            Left            =   210
            List            =   "frmStatisticsCollection.frx":152E
            Style           =   1  'Checkbox
            TabIndex        =   7
            Top             =   345
            Width           =   15930
         End
      End
      Begin VB.Frame fraSites 
         Caption         =   "Disapline"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   2820
         Left            =   255
         TabIndex        =   4
         Top             =   1890
         Width           =   2350
         Begin VB.ListBox lstSites 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   2220
            ItemData        =   "frmStatisticsCollection.frx":1530
            Left            =   240
            List            =   "frmStatisticsCollection.frx":1543
            Style           =   1  'Checkbox
            TabIndex        =   5
            Top             =   330
            Width           =   1905
         End
      End
      Begin VB.CommandButton Command 
         Height          =   150
         Index           =   0
         Left            =   4200
         TabIndex        =   3
         Top             =   11040
         Visible         =   0   'False
         Width           =   7140
      End
      Begin VB.TextBox txtExporting 
         BackColor       =   &H00000080&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   900
         Left            =   11587
         MultiLine       =   -1  'True
         TabIndex        =   2
         Text            =   "frmStatisticsCollection.frx":158A
         Top             =   11280
         Visible         =   0   'False
         Width           =   1210
      End
      Begin TabDlg.SSTab tabResults 
         Height          =   6240
         Left            =   240
         TabIndex        =   24
         Top             =   4800
         Width           =   18885
         _ExtentX        =   33311
         _ExtentY        =   11007
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   8
         TabHeight       =   520
         WordWrap        =   0   'False
         BackColor       =   4210688
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Statistics"
         TabPicture(0)   =   "frmStatisticsCollection.frx":15AA
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "G(7)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "fraCSProgressBar2"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Statistics2"
         TabPicture(1)   =   "frmStatisticsCollection.frx":15C6
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fraCSProgressBar"
         Tab(1).Control(1)=   "G(1)"
         Tab(1).ControlCount=   2
         Begin VB.Frame fraCSProgressBar2 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000C&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   900
            Left            =   7560
            TabIndex        =   44
            Top             =   2400
            Visible         =   0   'False
            Width           =   3840
            Begin MSComctlLib.ProgressBar CSProgressBar2 
               Height          =   375
               Left            =   -15
               TabIndex        =   45
               Top             =   360
               Width           =   3855
               _ExtentX        =   6800
               _ExtentY        =   661
               _Version        =   393216
               Appearance      =   1
            End
            Begin VB.Label lblCSProgressBar2 
               Alignment       =   2  'Center
               BackColor       =   &H80000009&
               Caption         =   "Fetching Results........."
               Height          =   255
               Left            =   -15
               TabIndex        =   46
               Top             =   135
               Width           =   3855
            End
         End
         Begin VB.Frame fraCSProgressBar 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000C&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   900
            Left            =   -69870
            TabIndex        =   25
            Top             =   2490
            Visible         =   0   'False
            Width           =   3840
            Begin MSComctlLib.ProgressBar CSProgressBar 
               Height          =   375
               Left            =   -15
               TabIndex        =   26
               Top             =   375
               Width           =   3855
               _ExtentX        =   6800
               _ExtentY        =   661
               _Version        =   393216
               Appearance      =   1
            End
            Begin VB.Label lblCSProgressBar 
               Alignment       =   2  'Center
               BackColor       =   &H80000009&
               Caption         =   "Fetching Results........."
               Height          =   255
               Left            =   -15
               TabIndex        =   27
               Top             =   135
               Width           =   3855
            End
         End
         Begin MSFlexGridLib.MSFlexGrid G 
            Height          =   4215
            Index           =   0
            Left            =   -74730
            TabIndex        =   28
            Top             =   930
            Width           =   13515
            _ExtentX        =   23839
            _ExtentY        =   7435
            _Version        =   393216
            Cols            =   1
            FixedCols       =   0
            BackColor       =   12648447
            BackColorFixed  =   8421376
            ForeColorFixed  =   -2147483624
            BackColorSel    =   4210688
            AllowBigSelection=   0   'False
            ScrollTrack     =   -1  'True
            GridLines       =   3
            GridLinesFixed  =   3
            SelectionMode   =   1
            FormatString    =   "                                      "
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSFlexGridLib.MSFlexGrid G 
            Height          =   4875
            Index           =   1
            Left            =   -74790
            TabIndex        =   29
            Top             =   405
            Width           =   13785
            _ExtentX        =   24315
            _ExtentY        =   8599
            _Version        =   393216
            Cols            =   10
            FixedCols       =   0
            BackColor       =   12648447
            BackColorFixed  =   8421376
            ForeColorFixed  =   -2147483624
            BackColorSel    =   4210688
            AllowBigSelection=   0   'False
            ScrollTrack     =   -1  'True
            GridLines       =   3
            GridLinesFixed  =   3
            SelectionMode   =   1
            FormatString    =   "                                      "
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSFlexGridLib.MSFlexGrid G 
            Height          =   4245
            Index           =   2
            Left            =   -74685
            TabIndex        =   30
            Top             =   1380
            Width           =   13515
            _ExtentX        =   23839
            _ExtentY        =   7488
            _Version        =   393216
            Rows            =   3
            Cols            =   10
            FixedRows       =   2
            FixedCols       =   0
            BackColor       =   12648447
            BackColorFixed  =   8421376
            ForeColorFixed  =   -2147483624
            BackColorSel    =   4210688
            AllowBigSelection=   0   'False
            ScrollTrack     =   -1  'True
            GridLines       =   3
            GridLinesFixed  =   3
            SelectionMode   =   1
            AllowUserResizing=   3
            FormatString    =   "                                      "
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSFlexGridLib.MSFlexGrid G 
            Height          =   4725
            Index           =   3
            Left            =   -74850
            TabIndex        =   31
            Top             =   705
            Width           =   13785
            _ExtentX        =   24315
            _ExtentY        =   8334
            _Version        =   393216
            Rows            =   3
            Cols            =   10
            FixedRows       =   2
            FixedCols       =   0
            BackColor       =   12648447
            BackColorFixed  =   8421376
            ForeColorFixed  =   -2147483624
            BackColorSel    =   4210688
            AllowBigSelection=   0   'False
            ScrollTrack     =   -1  'True
            GridLines       =   3
            GridLinesFixed  =   3
            SelectionMode   =   1
            AllowUserResizing=   3
            FormatString    =   "                                      "
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSFlexGridLib.MSFlexGrid G 
            Height          =   4920
            Index           =   4
            Left            =   -74835
            TabIndex        =   32
            Top             =   750
            Width           =   13860
            _ExtentX        =   24448
            _ExtentY        =   8678
            _Version        =   393216
            Rows            =   3
            Cols            =   198
            FixedCols       =   0
            BackColor       =   12648447
            BackColorFixed  =   8421376
            ForeColorFixed  =   -2147483624
            BackColorSel    =   4210688
            AllowBigSelection=   0   'False
            ScrollTrack     =   -1  'True
            GridLines       =   3
            GridLinesFixed  =   3
            SelectionMode   =   1
            AllowUserResizing=   3
            FormatString    =   "                                      "
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSFlexGridLib.MSFlexGrid G 
            Height          =   4725
            Index           =   5
            Left            =   -74850
            TabIndex        =   33
            Top             =   750
            Width           =   13785
            _ExtentX        =   24315
            _ExtentY        =   8334
            _Version        =   393216
            Rows            =   3
            Cols            =   10
            FixedCols       =   0
            BackColor       =   12648447
            BackColorFixed  =   8421376
            ForeColorFixed  =   -2147483624
            BackColorSel    =   4210688
            AllowBigSelection=   0   'False
            ScrollTrack     =   -1  'True
            GridLines       =   3
            GridLinesFixed  =   3
            SelectionMode   =   1
            AllowUserResizing=   3
            FormatString    =   "                                      "
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSFlexGridLib.MSFlexGrid G 
            Height          =   4875
            Index           =   6
            Left            =   -74790
            TabIndex        =   34
            Top             =   405
            Width           =   13785
            _ExtentX        =   24315
            _ExtentY        =   8599
            _Version        =   393216
            Rows            =   3
            Cols            =   10
            FixedRows       =   2
            FixedCols       =   0
            BackColor       =   12648447
            BackColorFixed  =   8421376
            ForeColorFixed  =   -2147483624
            BackColorSel    =   4210688
            AllowBigSelection=   0   'False
            ScrollTrack     =   -1  'True
            GridLines       =   3
            GridLinesFixed  =   3
            SelectionMode   =   1
            AllowUserResizing=   3
            FormatString    =   "                                      "
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSFlexGridLib.MSFlexGrid G 
            Height          =   5595
            Index           =   7
            Left            =   165
            TabIndex        =   43
            Top             =   480
            Width           =   18585
            _ExtentX        =   32782
            _ExtentY        =   9869
            _Version        =   393216
            Cols            =   10
            FixedCols       =   0
            BackColor       =   12648447
            BackColorFixed  =   8421376
            ForeColorFixed  =   -2147483624
            BackColorSel    =   4210688
            AllowBigSelection=   0   'False
            ScrollTrack     =   -1  'True
            GridLines       =   3
            GridLinesFixed  =   3
            SelectionMode   =   1
            FormatString    =   "                                      "
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
   End
   Begin VB.CommandButton cmdFilterOrganisms 
      Caption         =   "Filter Organisms"
      Enabled         =   0   'False
      Height          =   495
      Left            =   5805
      TabIndex        =   0
      Top             =   1680
      Width           =   1215
   End
   Begin TabDlg.SSTab tabHospitals 
      Height          =   12930
      Left            =   0
      TabIndex        =   35
      Top             =   0
      Width           =   19440
      _ExtentX        =   34290
      _ExtentY        =   22807
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Mullingar"
      TabPicture(0)   =   "frmStatisticsCollection.frx":15E2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
   End
End
Attribute VB_Name = "frmStatisticsCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Disapline As String
Dim LongTestName As String
Dim TestCode As String
Dim testCount As Integer
Dim tests() As Variant
Dim TestNames() As Variant
Dim ResultDisapline As String
Dim filterType As String
Dim Hours As Integer
Private Sub FillG()
      Dim tb As Recordset
      Dim sql As String
      Dim s As String
      Dim i As Integer
      Dim calFrom As String
      Dim calTo As String
      Dim NumberOfColums As Integer
      Dim NumberOfRows As Integer
      Dim J As Integer
      Dim k As Integer
      Dim temp As String
      Dim f As Integer
      Dim T As Integer
      Dim isList As Boolean

10    isList = False

20    On Error GoTo FillG_Error

30    For T = 0 To lstDefinitions.ListCount - 1

40        If lstDefinitions.Selected(T) = True Then
50            isList = True
60        End If
70    Next T

80    calFrom = Format(dtFrom, "dd/MMM/yyyy 00:00:00")
90    calTo = Format(dtTo, "dd/MMM/yyyy 23:59:59")
100   fraCSProgressBar.Visible = True

110   If isList Then


120       If ResultDisapline = "HaemResults" Then
130           sql = "SELECT * from (select  TOP (100) PERCENT d.SampleID As Lab_Number, d.DoB,d.Sex,d.gp,d.ward,d.Clinician, d.RunDate,d.RecDate StartTime, "
140           For i = 0 To UBound(tests) - 1
150               temp = Left(TestNames(i), 1)
160               If IsNumeric(temp) Then
170                   TestNames(i) = "[" & TestNames(i) & "]"
180               End If
190               sql = sql & "(SELECT RunDateTime  From " & ResultDisapline & "  WHERE  (SampleID = d.SampleID))  AS FinishTime, "
200               sql = sql & "(SELECT  " & tests(i) & " From " & ResultDisapline & "  WHERE (SampleID = d.SampleID))  AS " & TestNames(i) & ""
210               If i <> UBound(tests) - 1 Then
220                   sql = sql & ","
230               End If
240           Next i
250       Else
260           sql = "SELECT * from (select  TOP (100) PERCENT d.SampleID As Lab_Number, d.DoB,d.Sex,d.gp,d.ward,d.Clinician, d.RunDate,d.RecDate StartTime, "
270           For i = 0 To UBound(tests) - 1
280               temp = Left(TestNames(i), 1)
290               If IsNumeric(temp) Then
300                   TestNames(i) = "[" & TestNames(i) & "]"
310               End If
320               sql = sql & "(SELECT RunTime  From " & ResultDisapline & "  WHERE (code = '" & tests(i) & "') AND (SampleID = d.SampleID))  AS FinishTime, "
330               sql = sql & "(SELECT Result  From " & ResultDisapline & "  WHERE (code = '" & tests(i) & "') AND (SampleID = d.SampleID)and isnumeric (result)<>0)  AS " & TestNames(i) & ""
340               If i <> UBound(tests) - 1 Then
350                   sql = sql & ","
360               End If
370           Next i
380       End If

390       sql = sql & " FROM Demographics D "
400       sql = sql & "where "
410       If cmbGp <> "" Then
420           sql = sql & " d.gp= '" & cmbGp.Text & "' And "
430       End If
440       sql = sql & "D.DateTimeDemographics Between '" & calFrom & "' AND '" & calTo & "' ) m Where "

450       For i = 0 To UBound(tests) - 1
460           sql = sql & "" & TestNames(i) & " is Not NULL "
470           If i <> UBound(tests) - 1 Then
480               sql = sql & " or "
490           End If
500       Next i

510       sql = sql & " order by Lab_Number desc "

520       Set tb = New Recordset
530       RecOpenServer 0, tb, sql

540       NumberOfColums = tb.Fields.Count
550       G(1).Cols = NumberOfColums
560       NumberOfRows = 1

570       Do While Not tb.EOF
580           If NumberOfRows = 1 Then
590               For J = 10 To NumberOfColums - 1
600                   G(1).ColWidth(J) = 400
610                   G(1).TextMatrix(0, J) = tb(J).Name
620                   If G(1).ColWidth(J) < frmStatisticsCollection.TextWidth(G(1).TextMatrix(0, J)) + 100 Then
630                       G(1).ColWidth(J) = frmStatisticsCollection.TextWidth(G(1).TextMatrix(0, J)) + 100
640                   End If
650               Next J
660           End If

670           For J = 0 To NumberOfColums - 1
680               If tb(J) <> "NULL" Then
690                   G(1).TextMatrix(NumberOfRows, J) = tb(J)
700                   If G(1).ColWidth(J) < frmStatisticsCollection.TextWidth(G(1).TextMatrix(NumberOfRows, J)) + 100 Then
710                       G(1).ColWidth(J) = frmStatisticsCollection.TextWidth(G(1).TextMatrix(NumberOfRows, J)) + 100
720                   End If
730               End If
740           Next J

750           NumberOfRows = NumberOfRows + 1
760           G(1).Rows = NumberOfRows + 1
770           tb.MoveNext
780           CSProgressBar.Value = CSProgressBar.Value + 1
790           lblCSProgressBar = "Fetching results ... (" & Int(CSProgressBar.Value * 100 / CSProgressBar.Max) & " %)"
800           If CSProgressBar.Value = 100 Then
810               CSProgressBar.Value = 0
820           End If
830           lblCSProgressBar.Refresh
840       Loop

850       fraCSProgressBar.Visible = False
860       txtResultTotal.Text = NumberOfRows - 1
870       testCount = 0

880   End If

890   Exit Sub

FillG_Error:

      Dim strES As String
      Dim intEL As Integer

900   intEL = Erl
910   strES = Err.Description
920   LogError "frmStatisticsCollection", "FillG", intEL, strES, sql

End Sub
Private Sub FillG2()
Dim tb As Recordset
Dim sql As String
Dim s As String
Dim i As Integer
Dim calFrom As String
Dim calTo As String
Dim NumberOfColums As Integer
Dim NumberOfRows As Integer
Dim J As Integer
Dim k As Integer
Dim temp As String
Dim T As Integer
Dim isList As Boolean
Dim TAT As Date
Dim TATOver As Integer

On Error GoTo FillG2_Error

isList = False
TATOver = 0
txtOver1hr.Text = ""
For T = 0 To lstDefinitions.ListCount - 1

    If lstDefinitions.Selected(T) = True Then
        isList = True
    End If
Next T

If isList Then
    calFrom = Format(dtFrom, "dd/MMM/yyyy 00:00:00")
    calTo = Format(dtTo, "dd/MMM/yyyy 23:59:59")
    fraCSProgressBar2.Visible = True

    If ResultDisapline = "HaemResults" Then
        sql = "SELECT * from (select  TOP (100) PERCENT d.SampleID As Lab_Number, d.DoB,d.Sex,d.gp,d.ward,d.Clinician, d.RunDate, d.RecDate StartTime, "
        sql = sql & "(SELECT top 1 RunDateTime  From " & ResultDisapline & "  WHERE  (SampleID = d.SampleID))  AS FinishTime, d.SampleDate, "
        For i = 0 To UBound(tests) - 1
            temp = Left(TestNames(i), 1)
            If IsNumeric(temp) Then
                TestNames(i) = "[" & TestNames(i) & "]"
            End If
            If tests(i) = "RDW" Then
                tests(i) = "RDWCV"
            End If
            sql = sql & "(SELECT  " & tests(i) & " From " & ResultDisapline & "  WHERE (SampleID = d.SampleID))  AS " & TestNames(i) & ""
            If i <> UBound(tests) - 1 Then
                sql = sql & ","
            End If
        Next i
    Else
        sql = "SELECT * from (select  TOP (100) PERCENT d.SampleID As Lab_Number, d.DoB,d.Sex,d.gp,d.ward,d.Clinician, d.RunDate, d.RecDate StratTime,  "
        sql = sql & "(SELECT top 1 RunTime  From " & ResultDisapline & "  WHERE  (SampleID = d.SampleID))  AS FinishTime, d.SampleDate, "
        For i = 0 To UBound(tests) - 1
            sql = sql & "(SELECT Result  From " & ResultDisapline & "  WHERE (code = '" & tests(i) & "') AND (SampleID = d.SampleID)and isnumeric (result)<>0)  AS " & TestNames(i) & ""
            If i <> UBound(tests) - 1 Then
                sql = sql & ","
            End If
        Next i
        CSProgressBar2.Value = CSProgressBar2.Value + 1
        lblCSProgressBar2 = "Fetching results ... (" & Int(CSProgressBar2.Value * 100 / CSProgressBar2.Max) & " %)"
        If CSProgressBar2.Value = 100 Then
            CSProgressBar2.Value = 0
        End If
        lblCSProgressBar2.Refresh
    End If

    sql = sql & " FROM Demographics D "
    sql = sql & "where "

    If chkGP Then
        If cmbGp.Text <> "" Then
            sql = sql & " d.Gp = '" & cmbGp.Text & "' And "
        Else
             sql = sql & " d.Gp<>'' And "
        End If
    End If

    If chkClinician Then
        If cmbClinician.Text <> "" Then
            sql = sql & " d.Clinician = '" & cmbClinician.Text & "' And "
        Else
             sql = sql & " d.Clinician<>'' And "
        End If
    End If

    If chkWard Then
         If cmbWard.Text <> "" Then
            sql = sql & " d.Ward = '" & cmbWard.Text & "' And "
        Else
             sql = sql & " d.Ward<>'' And "
        End If
       
        
    End If

    sql = sql & "D.DateTimeDemographics Between '" & calFrom & "' AND '" & calTo & "' ) m Where "

    For i = 0 To UBound(tests) - 1
        sql = sql & "" & TestNames(i) & " is Not NULL "
        If i <> UBound(tests) - 1 Then
            sql = sql & " or "
        End If
    Next i

    sql = sql & " order by Lab_Number desc "
    Set tb = New Recordset
    RecOpenServer 0, tb, sql

    NumberOfColums = tb.Fields.Count
    G(7).Cols = NumberOfColums
    NumberOfRows = 1

    Do While Not tb.EOF
        If NumberOfRows = 1 Then
            For J = 10 To NumberOfColums - 1
                G(7).ColWidth(J) = 400
                G(7).TextMatrix(0, J) = tb(J).Name
                If G(7).ColWidth(J) < frmStatisticsCollection.TextWidth(G(7).TextMatrix(0, J)) + 100 Then
                    G(7).ColWidth(J) = frmStatisticsCollection.TextWidth(G(7).TextMatrix(0, J)) + 100
                End If
            Next J
        End If

        For J = 0 To NumberOfColums - 1
            If tb(J) <> "NULL" Then
                If J = 9 Then
                   G(7).TextMatrix(NumberOfRows, J) = DateDiff("n", tb(7), tb(8))
                Else
                    G(7).TextMatrix(NumberOfRows, J) = tb(J)
                    If G(7).ColWidth(J) < frmStatisticsCollection.TextWidth(G(7).TextMatrix(NumberOfRows, J)) + 100 Then
                        G(7).ColWidth(J) = frmStatisticsCollection.TextWidth(G(7).TextMatrix(NumberOfRows, J)) + 100
                    End If
                End If
            End If
        Next J

        If G(7).TextMatrix(NumberOfRows, 9) > (Hours * 60) Then
            G(7).row = NumberOfRows
            G(7).Col = 9
            G(7).CellBackColor = vbRed
            G(7).CellForeColor = vbWhite
            TATOver = TATOver + 1
        Else
            G(7).row = NumberOfRows
            G(7).Col = 9
            G(7).CellBackColor = &HC0FFFF
            G(7).CellForeColor = &H80000008
        End If
'             FormatMin G(7).TextMatrix(NumberOfRows, 9)
        'G(7).TextMatrix(NumberOfRows, 9) = Formatin
        
        NumberOfRows = NumberOfRows + 1
        G(7).Rows = NumberOfRows + 1
        tb.MoveNext
        CSProgressBar2.Value = CSProgressBar2.Value + 1
        lblCSProgressBar2 = "Fetching results ... (" & Int(CSProgressBar2.Value * 100 / CSProgressBar2.Max) & " %)"
        If CSProgressBar2.Value = 100 Then
            CSProgressBar2.Value = 0
        End If
        lblCSProgressBar2.Refresh
    Loop

    fraCSProgressBar2.Visible = False
    txtResultTotal.Text = NumberOfRows - 1
    txtOver1hr.Text = TATOver
    testCount = 0
End If

Exit Sub

FillG2_Error:

Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "frmStatisticsCollection", "FillG2", intEL, strES, sql

End Sub
'Private Function FormatMin(ByVal iMin As Integer) As String
'    FormatMin = Format$(iMin \ 60, "00") & ":" & _
'                Format$(iMin Mod 60, "00")
'End Function



Private Sub cmdExcel_Click()
      Dim strHeading As String

10    On Error GoTo cmdExcel_Click_Error


20    If tabResults.Tab = 0 Then
30        If G(7).Rows < 2 Then
40            iMsg "No Demographics to export"
50            Exit Sub
60        End If
70        strHeading = "Counts from Statistics Collection Searches " & vbCr
80        ExpotToExcell 6, strHeading
90    End If

100   Exit Sub

cmdExcel_Click_Error:

      Dim strES As String
      Dim intEL As Integer

110   intEL = Erl
120   strES = Err.Description
130   LogError "frmMicroSurveillanceSearches", "cmdExcel_Click", intEL, strES

End Sub
Public Sub ExportFlexGrid(ByVal objGrid As MSFlexGrid, _
                          ByVal CallingForm As Form, _
                          Optional ByVal HeadingMatrix As String = "")
                          
      Dim objXL As Object
      Dim objWB As Object
      Dim objWS As Object
      Dim R As Long
      Dim c As Long
      Dim T As Single
      Dim i As Boolean
      Dim filterRows As Integer
      Dim AmountOfCols As Integer
      Dim x As Integer
      Dim intLineCount As Integer
      Dim intColCount As Integer

10    On Error GoTo ExportFlexGrid_Error
20    AmountOfCols = objGrid.Cols
30    For x = 0 To objGrid.Cols - 1
40        If objGrid.ColWidth(x) = 0 Then
50            AmountOfCols = AmountOfCols - 1
60        End If
70    Next x
80    txtExporting.Visible = True
90    Set objXL = CreateObject("Excel.Application")
100   Set objWB = objXL.Workbooks.Add
110   Set objWS = objWB.Worksheets(1)
120   intLineCount = 0
130   If HeadingMatrix <> "" Then
140       With objWS
              Dim strTokens() As String
150           strTokens = Split(HeadingMatrix, vbCr)
160           intLineCount = UBound(strTokens)
170           For R = LBound(strTokens) To UBound(strTokens) - 1
180               .range(.Cells(R + 1, 1), .Cells(R + 1, AmountOfCols)).MergeCells = True
190               .range(.Cells(R + 1, 1), .Cells(R + 1, AmountOfCols)).HorizontalAlignment = 3
200               .range(.Cells(R + 1, 1), .Cells(R + 1, AmountOfCols)).Font.Bold = True
210               .range(.Cells(R + 1, 1), .Cells(R + 1, AmountOfCols)).Font.Size = 16
220               .range(.Cells(R + 1, 1), .Cells(R + 1, AmountOfCols)).Borders.Weight = 4
230               objWS.Cells(R + 1, 1) = "'" & strTokens(R)
240           Next
250       End With
260   End If
270   With objWS
280       filterRows = UBound(strTokens)
290       For R = 0 To objGrid.Rows - 1
300           If objGrid.RowHeight(R) > 0 Then
310               filterRows = filterRows + 1
320               intColCount = 1
330               For c = 0 To objGrid.Cols - 1
340                   If objGrid.ColWidth(c) > 0 Then
350                       If R = 0 And c = 0 Then
360                           .range(.Cells(R + 1 + intLineCount, 1), .Cells(R + 1 + intLineCount, AmountOfCols)).Font.Bold = True
370                           .range(.Cells(R + 1 + intLineCount, 1), .Cells(R + 1 + intLineCount, AmountOfCols)).WrapText = True
380                           .range(.Cells(R + 1, 1), .Cells(R + 1, AmountOfCols)).Borders.Weight = 2
390                       End If
400                       objGrid.row = R
410                       objGrid.Col = c
420                       .Cells(filterRows, intColCount) = "'" & objGrid.TextMatrix(R, c)
430                       .Cells(filterRows, intColCount).Borders.Weight = 2
440                       intColCount = intColCount + 1
450                   End If
460               Next
470           End If
480       Next
490       objXL.ActiveSheet.PageSetup.LeftMargin = 0.25
500       objXL.ActiveSheet.PageSetup.RightMargin = 0.25
510       objXL.ActiveSheet.PageSetup.Orientation = 2
520       objXL.ActiveSheet.PageSetup.Zoom = 60  ' Reduce to 60% when printing
530       .Cells.Columns.AutoFit
540       .Cells.Columns(14).ColumnWidth = 50
550       .Cells.Columns(15).ColumnWidth = 50
560   End With
570   objXL.Visible = True
580   Set objWS = Nothing
590   Set objWB = Nothing
600   Set objXL = Nothing
610   txtExporting.Visible = False
620   Exit Sub

ExportFlexGrid_Error:

      Dim strES As String
      Dim lngErr As Long

630   txtExporting.Visible = False
640   With CallingForm.lblExcelInfo
650       .Caption = "Error " & Format(lngErr)
660       .Refresh
670       T = Timer
680       Do While Timer - T < 1: Loop
690       .Visible = False
700   End With

End Sub

Public Function ExpotToExcell(ByVal TabNumber As Integer, strHeading As String)

10    On Error GoTo ExpotToExcell_Error

20    Select Case TabNumber
      Case 0:    'General
30        ExportFlexGrid Me.G, Me, strHeading
40    Case 1:    'Demographics
50        ExportFlexGrid Me.G(1), Me, strHeading
60    Case 2:    'Results
          'ExportFlexGrid Me.g3, Me, strHeading
70    Case 5:    'All Results
          'ExportFlexGrid Me.g5, Me, strHeading
80    Case 6:    'Other Results
          ExportFlexGrid Me.G(7), Me, strHeading
90    Case 7:    'Other Results
          'ExportFlexGrid Me.GeneralG, Me, strHeading
100   End Select

110   Exit Function

ExpotToExcell_Error:

      Dim strES As String
      Dim intEL As Integer

120   intEL = Erl
130   strES = Err.Description
140   LogError "frmMicroSurveillanceSearches", "ExpotToExcell", intEL, strES
End Function

Private Sub cmdExit_Click()
10        Unload Me
End Sub

Private Sub cmdRecalc_Click()

10    On Error GoTo cmdRecalc_Click_Error

20    If lstDefinitions.ListCount <> 0 Then
30        Erase TestNames
40        Erase tests
50        G(1).Clear
60        G(7).Clear
70        CreateTestArray
80        FillGridHeadings
90        FillGridHeadings2
100       Hours = cmbHours.Text
110       FillG2
120   End If
130   Exit Sub

cmdRecalc_Click_Error:

      Dim strES As String
      Dim intEL As Integer

140   intEL = Erl
150   strES = Err.Description
160   LogError "frmStatisticsCollection", "cmdRecalc_Click", intEL, strES

End Sub
Private Sub FillGridHeadings()

10    On Error GoTo FillGridHeadings_Error

20    With G(1)

30        .ColWidth(0) = 950: .TextMatrix(0, 0) = "Sample ID"
40        .TextMatrix(0, 1) = "DOB"
50        .ColWidth(2) = 400: .TextMatrix(0, 2) = "Sex"
60        .TextMatrix(0, 3) = "GP"
70        .TextMatrix(0, 4) = "Ward"
80        .TextMatrix(0, 5) = "Clinician"
90        .TextMatrix(0, 6) = "Run Date"
100       .TextMatrix(0, 7) = "Time In"
110       .TextMatrix(0, 8) = "Time Out"
120       .TextMatrix(0, 9) = "TAT"
130       .Rows = 2
140   End With

150   Exit Sub

FillGridHeadings_Error:

      Dim strES As String
      Dim intEL As Integer

160   intEL = Erl
170   strES = Err.Description
180   LogError "frmStatisticsCollection", "FillGridHeadings", intEL, strES

End Sub
Private Sub FillGridHeadings2()


10    On Error GoTo FillGridHeadings2_Error

20    With G(7)
30        .ColWidth(0) = 950: .TextMatrix(0, 0) = "Sample ID"
40        .TextMatrix(0, 1) = "DOB"
50        .ColWidth(2) = 400: .TextMatrix(0, 2) = "Sex"
60        .TextMatrix(0, 3) = "GP"
70        .TextMatrix(0, 4) = "Ward"
80        .TextMatrix(0, 5) = "Clinician"
90        .TextMatrix(0, 6) = "Run Date"
100       .TextMatrix(0, 7) = "Time In"
110       .TextMatrix(0, 8) = "Time Out"
120       .TextMatrix(0, 9) = "TAT"
130       .Rows = 2
140   End With

150   Exit Sub

FillGridHeadings2_Error:

      Dim strES As String
      Dim intEL As Integer

160   intEL = Erl
170   strES = Err.Description
180   LogError "frmStatisticsCollection", "FillGridHeadings2", intEL, strES

End Sub

Private Sub Form_Load()


10    Me.Top = (Screen.Height - Me.Height) / 2
20    Me.Left = (Screen.Width - Me.Width) / 2

30    FillGPsClinWard Me, HospName(0)

40    dtTo = Format(Now, "dd/MM/yyyy")
50    dtFrom = Format(Now, "dd/MM/yyyy")

60    'cmbClinician.Visible = False
70    'cmbWard.Visible = False
80    tabResults.TabVisible(1) = False

90    Exit Sub


End Sub

Private Sub CreateTestArray()

      Dim TestName As String
      Dim i, k, T As Integer
      Dim Msg As String
      Dim chr As String

10    On Error GoTo CreateTestArray_Error

20    i = 0

30    For T = 0 To lstDefinitions.ListCount - 1

40        If lstDefinitions.Selected(T) = True Then
50            k = 1
60            TestName = lstDefinitions.List(T)
70            FindTestCode TestName
80            ReDim Preserve tests(testCount + 1)
90            ReDim Preserve TestNames(testCount + 1)
100           tests(testCount) = TestCode
110           Do While k < Len(TestName) + 1
120               If (Asc(Mid$(TestName, k, 1)) < 48) Or ((Asc(Mid$(TestName, k, 1)) > 57) And (Asc(Mid$(TestName, k, 1)) < 65)) Or ((Asc(Mid$(TestName, k, 1)) > 90) And (Asc(Mid$(TestName, k, 1)) < 97)) Or (Asc(Mid$(TestName, k, 1)) > 122) Then
130                   TestName = Replace(TestName, Mid$(TestName, k, 1), "")
140                   k = k - 1
150               End If
160               k = k + 1
170           Loop

180           TestNames(testCount) = UCase$(TestName)
190           If testCount > 0 Then
200               For i = 0 To testCount - 1
210                   If TestNames(i) = UCase$(TestName) Then
220                       TestNames(testCount) = TestName + "1"
230                       Exit For
240                   End If
250               Next i
260           End If

270           testCount = testCount + 1
280       Else
290       End If
300   Next T

310   Exit Sub

CreateTestArray_Error:

      Dim strES As String
      Dim intEL As Integer

320   intEL = Erl
330   strES = Err.Description
340   LogError "frmStatisticsCollection", "CreateTestArray", intEL, strES

End Sub

Private Sub lstSites_Click()

      Dim DisaplineNumber As Integer
      Dim i As Integer

10    On Error GoTo lstSites_Click_Error

20    With lstSites

30        DisaplineNumber = .ListIndex
40        Disapline = .Text

50        For i = 0 To .ListCount - 1
              ' If an item is checked, add it to the string
60            If i <> DisaplineNumber Then
70                .Selected(i) = False
80            End If
90        Next

100       Select Case DisaplineNumber
          Case 0
110           Disapline = "BioTestDefinitions"
120           LongTestName = "LongName"
130           ResultDisapline = "BioResults"
140       Case 1
150           Disapline = "CoagTestDefinitions"
160           LongTestName = "TestName"
170           ResultDisapline = "CoagResults"
180       Case 2
190           Disapline = "EndTestDefinitions"
200           LongTestName = "LongName"
210           ResultDisapline = "EndResults"
220       Case 3
230           Disapline = "HaemTestDefinitions"
240           LongTestName = "AnalyteName"
250           ResultDisapline = "HaemResults"
260       Case 4
270           Disapline = "ImmTestDefinitions"
280           LongTestName = "LongName"
290           ResultDisapline = "ImmResults"
300       End Select

310       FillTests Disapline

320   End With

330   Exit Sub

lstSites_Click_Error:

      Dim strES As String
      Dim intEL As Integer

340   intEL = Erl
350   strES = Err.Description
360   LogError "frmStatisticsCollection", "lstSites_Click", intEL, strES
End Sub

Private Sub FillTests(ByVal Disapline As String)

          Dim tb As Recordset
          Dim sql As String
          Dim s As String

10        On Error GoTo FillTests_Error
          
20        lstDefinitions.Clear
          
30        sql = "SELECT Distinct Code,longname " & _
                "FROM " & Disapline & " " & _
                "Where inuse = 1 " & _
                "ORDER BY longname"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql

60        Do While Not tb.EOF
70            With tb
80                lstDefinitions.AddItem !LongName
90            End With
100           tb.MoveNext
110       Loop

120       Exit Sub

FillTests_Error:

          Dim strES As String
          Dim intEL As Integer

130       intEL = Erl
140       strES = Err.Description
150       LogError "frmStatisticsCollection", "FillTests", intEL, strES, sql

End Sub

Private Sub FindTestCode(ByVal TestName As String)

      Dim tb As Recordset
      Dim sql As String
      Dim s As String

10    On Error GoTo FindTestCode_Error

20    sql = "SELECT DISTINCT Code " & _
            "FROM " & Disapline & " " & _
            "WHERE " & LongTestName & " = '" & TestName & "' " & _
            "and inuse=1 "
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql

50    Do While Not tb.EOF
60        With tb
70            TestCode = !Code
80        End With
90        tb.MoveNext
100   Loop
110   Exit Sub

FindTestCode_Error:

      Dim strES As String
      Dim intEL As Integer

120   intEL = Erl
130   strES = Err.Description
140   LogError "frmStatisticsCollection", "FindTestCode", intEL, strES, sql

End Sub
Private Sub optBetween_Click(Index As Integer)

      Dim upto As String

10    On Error GoTo optBetween_Click_Error

20    dtFrom = BetweenDates(Index, upto)
30    dtTo = upto
40    cmdRecalc.Visible = True

50    Exit Sub

optBetween_Click_Error:

      Dim strES As String
      Dim intEL As Integer

60    intEL = Erl
70    strES = Err.Description
80    LogError "frmStatisticsCollection", "optBetween_Click", intEL, strES

End Sub

Public Function BetweenDates(ByVal Index As Integer, _
                             ByRef upto As String) _
                             As String

      Dim From As String
      Dim m As Long

10    On Error GoTo BetweenDates_Error

20    Select Case Index
      Case 0:    'last week
30        From = Format$(DateAdd("ww", -1, Now), "dd/mm/yyyy")
40        upto = Format$(Now, "dd/mm/yyyy")
50    Case 1:    'last month
60        From = Format$(DateAdd("m", -1, Now), "dd/mm/yyyy")
70        upto = Format$(Now, "dd/mm/yyyy")
80    Case 2:    'last fullmonth
90        From = Format$(DateAdd("m", -1, Now), "dd/mm/yyyy")
100       From = "01/" & Mid$(From, 4)
110       upto = DateAdd("m", 1, From)
120       upto = Format$(DateAdd("d", -1, upto), "dd/mm/yyyy")
130   Case 3:    'last quarter
140       From = Format$(DateAdd("q", -1, Now), "dd/mm/yyyy")
150       upto = Format$(Now, "dd/mm/yyyy")
160   Case 4:    'last full quarter
170       From = Format$(DateAdd("q", -1, Now), "dd/mm/yyyy")
180       m = Val(Mid$(From, 4, 2))
190       m = ((m - 1) \ 3) * 3 + 1
200       From = "01/" & Format$(m, "00") & Mid$(From, 6)
210       upto = DateAdd("q", 1, From)
220       upto = Format$(DateAdd("d", -1, upto), "dd/mm/yyyy")
230   Case 5:    'year to date
240       From = "01/01/" & Format$(Now, "yyyy")
250       upto = Format$(Now, "dd/mm/yyyy")
260   Case 6:    'today
270       From = Format$(Now, "dd/mm/yyyy")
280       upto = From
290   End Select
300   BetweenDates = From

310   Exit Function

320   Exit Function

BetweenDates_Error:

      Dim strES As String
      Dim intEL As Integer

330   intEL = Erl
340   strES = Err.Description
350   LogError "frmStatisticsCollection", "BetweenDates", intEL, strES

End Function

Private Sub optFilter_Click(Index As Integer)

10    On Error GoTo optFilter_Click_Error

20    Select Case Index
      Case 0:    'Show GP's List
30        cmbGp.Visible = True
40        cmbClinician.Visible = False
50        cmbWard.Visible = False
60        filterType = 0
70    Case 1:    'Show Clinicians List
80        cmbClinician.Visible = True
90        cmbGp.Visible = False
100       cmbWard.Visible = False
110       filterType = 1
120   Case 2:    'Show Wards List
130       cmbWard.Visible = True
140       cmbGp.Visible = False
150       cmbClinician.Visible = False
160       filterType = 2
170   End Select

180   Exit Sub

optFilter_Click_Error:

      Dim strES As String
      Dim intEL As Integer

190   intEL = Erl
200   strES = Err.Description
210   LogError "frmStatisticsCollection", "optFilter_Click", intEL, strES

End Sub

