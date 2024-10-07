VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frmMicroSurveillanceSearches 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Micro Surveillance Searches"
   ClientHeight    =   12810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15300
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12810
   ScaleWidth      =   15300
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSCheck SSCheckFlu 
      Height          =   735
      Left            =   6960
      TabIndex        =   60
      Top             =   1440
      Width           =   975
      _Version        =   65536
      _ExtentX        =   1720
      _ExtentY        =   1296
      _StockProps     =   78
      Caption         =   "Flu/Covid"
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
   Begin VB.ComboBox cmbClinician 
      Height          =   315
      Left            =   9900
      TabIndex        =   59
      Top             =   1560
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.ComboBox cmbGP 
      Height          =   315
      Left            =   10020
      TabIndex        =   58
      Top             =   1560
      Visible         =   0   'False
      Width           =   4395
   End
   Begin VB.ComboBox cmbWard 
      Height          =   315
      Left            =   9960
      TabIndex        =   57
      Top             =   1560
      Width           =   4455
   End
   Begin VB.CommandButton cmdFilterOrganisms 
      Caption         =   "Filter Organisms"
      Enabled         =   0   'False
      Height          =   495
      Left            =   5640
      TabIndex        =   51
      Top             =   1710
      Width           =   1215
   End
   Begin VB.Frame fraMainTab1 
      BackColor       =   &H00404000&
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
      Height          =   11700
      Left            =   150
      TabIndex        =   1
      Top             =   450
      Width           =   14610
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         Caption         =   "Gp/Clinician"
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
         Left            =   9600
         TabIndex        =   53
         Top             =   180
         Width           =   4890
         Begin VB.OptionButton optGPClinician 
            Caption         =   "Ward"
            Height          =   255
            Index           =   2
            Left            =   3720
            TabIndex        =   56
            Top             =   480
            Width           =   975
         End
         Begin VB.OptionButton optGPClinician 
            Caption         =   "Clinician"
            Height          =   255
            Index           =   1
            Left            =   1980
            TabIndex        =   55
            Top             =   480
            Width           =   975
         End
         Begin VB.OptionButton optGPClinician 
            Caption         =   "GP"
            Height          =   255
            Index           =   0
            Left            =   600
            TabIndex        =   54
            Top             =   480
            Width           =   615
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Sites"
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
         Left            =   14445
         TabIndex        =   44
         Top             =   1890
         Visible         =   0   'False
         Width           =   855
         Begin VB.ListBox lst2 
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
            Height          =   1950
            ItemData        =   "frmMicroSurveillanceSearches.frx":0000
            Left            =   240
            List            =   "frmMicroSurveillanceSearches.frx":0002
            Style           =   1  'Checkbox
            TabIndex        =   45
            Top             =   390
            Width           =   2955
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Sites"
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
         Left            =   14430
         TabIndex        =   42
         Top             =   2685
         Visible         =   0   'False
         Width           =   990
         Begin VB.ListBox List 
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
            ItemData        =   "frmMicroSurveillanceSearches.frx":0004
            Left            =   240
            List            =   "frmMicroSurveillanceSearches.frx":0006
            Style           =   1  'Checkbox
            TabIndex        =   43
            Top             =   360
            Width           =   2955
         End
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
         TabIndex        =   23
         Text            =   "frmMicroSurveillanceSearches.frx":0008
         Top             =   10650
         Visible         =   0   'False
         Width           =   1210
      End
      Begin VB.CommandButton Command 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Index           =   0
         Left            =   405
         TabIndex        =   22
         Top             =   11175
         Visible         =   0   'False
         Width           =   4140
      End
      Begin VB.Frame fraSites 
         Caption         =   "Sites"
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
         TabIndex        =   20
         Top             =   1890
         Width           =   4785
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
            ItemData        =   "frmMicroSurveillanceSearches.frx":0028
            Left            =   240
            List            =   "frmMicroSurveillanceSearches.frx":002A
            Style           =   1  'Checkbox
            TabIndex        =   21
            Top             =   330
            Width           =   4300
         End
      End
      Begin VB.Frame fraOrganismGroup 
         Caption         =   "Organisms"
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
         Left            =   5265
         TabIndex        =   18
         Top             =   1875
         Width           =   9180
         Begin VB.ListBox lstOrganismGroup 
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
            ItemData        =   "frmMicroSurveillanceSearches.frx":002C
            Left            =   210
            List            =   "frmMicroSurveillanceSearches.frx":002E
            Style           =   1  'Checkbox
            TabIndex        =   19
            Top             =   345
            Width           =   8730
         End
      End
      Begin VB.Timer tmrDown 
         Enabled         =   0   'False
         Interval        =   200
         Left            =   7230
         Top             =   10905
      End
      Begin VB.Timer tmrUp 
         Enabled         =   0   'False
         Interval        =   200
         Left            =   7665
         Top             =   10905
      End
      Begin VB.TextBox txtTest 
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
         Left            =   4590
         TabIndex        =   17
         Top             =   11010
         Visible         =   0   'False
         Width           =   2535
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
         TabIndex        =   13
         Top             =   180
         Width           =   7590
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
            TabIndex        =   30
            Top             =   1125
            Width           =   1755
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
            TabIndex        =   29
            Top             =   750
            Width           =   1395
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
            TabIndex        =   28
            Top             =   1125
            Width           =   1275
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
            TabIndex        =   27
            Top             =   750
            Width           =   1095
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
            TabIndex        =   26
            Top             =   390
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
            TabIndex        =   25
            Top             =   1125
            Width           =   1635
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
            TabIndex        =   24
            Top             =   750
            Width           =   1455
         End
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
            Left            =   5220
            MaskColor       =   &H8000000F&
            Picture         =   "frmMicroSurveillanceSearches.frx":0030
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   180
            Width           =   1200
         End
         Begin MSComCtl2.DTPicker dtFrom 
            Height          =   315
            Left            =   180
            TabIndex        =   15
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
            Format          =   116850689
            CurrentDate     =   37019
         End
         Begin MSComCtl2.DTPicker dtTo 
            Height          =   315
            Left            =   1830
            TabIndex        =   16
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
            Format          =   116850689
            CurrentDate     =   37019
         End
      End
      Begin VB.Frame fraResultType 
         Appearance      =   0  'Flat
         Caption         =   "Result Type"
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
         Left            =   7920
         TabIndex        =   10
         Top             =   180
         Width           =   1590
         Begin VB.CommandButton Command 
            Height          =   360
            Index           =   1
            Left            =   1920
            TabIndex        =   35
            Top             =   360
            Visible         =   0   'False
            Width           =   1725
         End
         Begin VB.OptionButton optType 
            Alignment       =   1  'Right Justify
            Caption         =   "All"
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
            Index           =   0
            Left            =   585
            TabIndex        =   33
            Top             =   360
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.OptionButton optType 
            Alignment       =   1  'Right Justify
            Caption         =   "Positive"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   225
            TabIndex        =   32
            Top             =   780
            Width           =   1095
         End
         Begin VB.OptionButton optType 
            Alignment       =   1  'Right Justify
            Caption         =   "Negative"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   225
            TabIndex        =   31
            Top             =   1200
            Width           =   1095
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
            Left            =   2580
            TabIndex        =   11
            Top             =   1065
            Width           =   1095
         End
         Begin VB.Label lblResultsTotal 
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
            Left            =   1860
            TabIndex        =   12
            Top             =   1185
            Width           =   1095
         End
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
         Picture         =   "frmMicroSurveillanceSearches.frx":033A
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   10635
         Width           =   1210
      End
      Begin VB.CommandButton cmdExit 
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
         Left            =   13185
         Picture         =   "frmMicroSurveillanceSearches.frx":0644
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   10635
         Width           =   1210
      End
      Begin TabDlg.SSTab tabResults 
         Height          =   5640
         Left            =   255
         TabIndex        =   3
         Top             =   4815
         Width           =   14175
         _ExtentX        =   25003
         _ExtentY        =   9948
         _Version        =   393216
         Tabs            =   9
         Tab             =   7
         TabsPerRow      =   9
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
         TabCaption(0)   =   "Generaltemp"
         TabPicture(0)   =   "frmMicroSurveillanceSearches.frx":150E
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "fraProgress"
         Tab(0).Control(1)=   "g"
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "C && S Counts"
         TabPicture(1)   =   "frmMicroSurveillanceSearches.frx":152A
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "g2"
         Tab(1).Control(1)=   "fraCSProgressBar"
         Tab(1).ControlCount=   2
         TabCaption(2)   =   "Tab2"
         TabPicture(2)   =   "frmMicroSurveillanceSearches.frx":1546
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "g3"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "Other Counts temp"
         TabPicture(3)   =   "frmMicroSurveillanceSearches.frx":1562
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "g4"
         Tab(3).ControlCount=   1
         TabCaption(4)   =   "Gen"
         TabPicture(4)   =   "frmMicroSurveillanceSearches.frx":157E
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "g5"
         Tab(4).Control(1)=   "fraMainProgressBar"
         Tab(4).ControlCount=   2
         TabCaption(5)   =   "Demographics"
         TabPicture(5)   =   "frmMicroSurveillanceSearches.frx":159A
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "G6"
         Tab(5).ControlCount=   1
         TabCaption(6)   =   "Other Counts"
         TabPicture(6)   =   "frmMicroSurveillanceSearches.frx":15B6
         Tab(6).ControlEnabled=   0   'False
         Tab(6).Control(0)=   "G7"
         Tab(6).ControlCount=   1
         TabCaption(7)   =   "General"
         TabPicture(7)   =   "frmMicroSurveillanceSearches.frx":15D2
         Tab(7).ControlEnabled=   -1  'True
         Tab(7).Control(0)=   "GeneralG"
         Tab(7).Control(0).Enabled=   0   'False
         Tab(7).ControlCount=   1
         TabCaption(8)   =   "GP/Clinician"
         TabPicture(8)   =   "frmMicroSurveillanceSearches.frx":15EE
         Tab(8).ControlEnabled=   0   'False
         Tab(8).Control(0)=   "GPClinician"
         Tab(8).ControlCount=   1
         Begin VB.Frame fraCSProgressBar 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000C&
            BorderStyle     =   0  'None
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
            ForeColor       =   &H80000008&
            Height          =   900
            Left            =   -69870
            TabIndex        =   48
            Top             =   2790
            Visible         =   0   'False
            Width           =   3840
            Begin MSComctlLib.ProgressBar CSProgressBar 
               Height          =   375
               Left            =   -15
               TabIndex        =   49
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
               TabIndex        =   50
               Top             =   135
               Width           =   3855
            End
         End
         Begin VB.Frame fraMainProgressBar 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000C&
            BorderStyle     =   0  'None
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
            ForeColor       =   &H80000008&
            Height          =   900
            Left            =   -69870
            TabIndex        =   38
            Top             =   3285
            Visible         =   0   'False
            Width           =   3840
            Begin MSComctlLib.ProgressBar MainProgressBar 
               Height          =   375
               Left            =   -15
               TabIndex        =   39
               Top             =   390
               Width           =   3855
               _ExtentX        =   6800
               _ExtentY        =   661
               _Version        =   393216
               Appearance      =   1
            End
            Begin VB.Label lblMainProgressBar 
               Alignment       =   2  'Center
               BackColor       =   &H80000009&
               Caption         =   "Fetching Results........."
               Height          =   255
               Left            =   0
               TabIndex        =   40
               Top             =   120
               Width           =   3855
            End
         End
         Begin VB.Frame fraProgress 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000C&
            BorderStyle     =   0  'None
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
            ForeColor       =   &H80000008&
            Height          =   735
            Left            =   -69870
            TabIndex        =   4
            Top             =   3285
            Visible         =   0   'False
            Width           =   3855
            Begin MSComctlLib.ProgressBar prgFetchingResults 
               Height          =   375
               Left            =   0
               TabIndex        =   5
               Top             =   360
               Width           =   3855
               _ExtentX        =   6800
               _ExtentY        =   661
               _Version        =   393216
               Appearance      =   1
            End
            Begin VB.Label lblProgress 
               Alignment       =   2  'Center
               BackColor       =   &H80000009&
               Caption         =   "Fetching Results........."
               Height          =   255
               Left            =   0
               TabIndex        =   6
               Top             =   120
               Width           =   3855
            End
         End
         Begin MSFlexGridLib.MSFlexGrid g 
            Height          =   4215
            Left            =   -74730
            TabIndex        =   7
            Top             =   1230
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
         Begin MSFlexGridLib.MSFlexGrid g2 
            Height          =   4875
            Left            =   -74790
            TabIndex        =   8
            Top             =   705
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
            ScrollBars      =   2
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
         Begin MSFlexGridLib.MSFlexGrid g3 
            Height          =   4245
            Left            =   -74685
            TabIndex        =   34
            Top             =   1680
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
         Begin MSFlexGridLib.MSFlexGrid g4 
            Height          =   4725
            Left            =   -74850
            TabIndex        =   36
            Top             =   1005
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
         Begin MSFlexGridLib.MSFlexGrid g5 
            Height          =   4920
            Left            =   -74835
            TabIndex        =   37
            Top             =   1050
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
         Begin MSFlexGridLib.MSFlexGrid G6 
            Height          =   4725
            Left            =   -74850
            TabIndex        =   41
            Top             =   1050
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
         Begin MSFlexGridLib.MSFlexGrid G7 
            Height          =   4875
            Left            =   -74790
            TabIndex        =   46
            Top             =   705
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
         Begin MSFlexGridLib.MSFlexGrid GeneralG 
            Height          =   4875
            Left            =   210
            TabIndex        =   47
            Top             =   705
            Width           =   13785
            _ExtentX        =   24315
            _ExtentY        =   8599
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
         Begin MSFlexGridLib.MSFlexGrid GPClinician 
            Height          =   4875
            Left            =   -74790
            TabIndex        =   52
            Top             =   705
            Width           =   13785
            _ExtentX        =   24315
            _ExtentY        =   8599
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
      End
   End
   Begin TabDlg.SSTab tabHospitals 
      Height          =   12210
      Left            =   75
      TabIndex        =   0
      Top             =   30
      Width           =   14760
      _ExtentX        =   26035
      _ExtentY        =   21537
      _Version        =   393216
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
      TabPicture(0)   =   "frmMicroSurveillanceSearches.frx":160A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Portloise"
      TabPicture(1)   =   "frmMicroSurveillanceSearches.frx":1626
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tullamore"
      TabPicture(2)   =   "frmMicroSurveillanceSearches.frx":1642
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
   End
End
Attribute VB_Name = "frmMicroSurveillanceSearches"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim calFrom As String
Dim calTo As String
Dim recordCount As Integer
Private AgeGroupFrom(0 To 4) As Long
Private AgeGroupTo(0 To 4) As Long

Private Sub cmdCancel_Click()

10    Unload Me

End Sub

Private Function SiteAndOrganismEnabled(ByVal Site As String, OrganismGroup As String) As Boolean
      Dim i As Integer
10    On Error GoTo SiteAndOrganismEnabled_Error

20    SiteAndOrganismEnabled = True
30    For i = 0 To lstSites.ListCount - 1
40        If Site = lstSites.List(i) And lstSites.Selected(i) = False Then
50            SiteAndOrganismEnabled = False
60            Exit Function
70        End If
80    Next i
90    For i = 0 To lstOrganismGroup.ListCount - 1
100       If OrganismGroup = lstOrganismGroup.List(i) And lstOrganismGroup.Selected(i) = False Then
110           SiteAndOrganismEnabled = False
120           Exit Function
130       End If
140   Next i

150   Exit Function

SiteAndOrganismEnabled_Error:

      Dim strES As String
      Dim intEL As Integer

160   intEL = Erl
170   strES = Err.Description
180   LogError "frmMicroSurveillanceSearches", "SiteAndOrganismEnabled", intEL, strES

End Function

Private Function OrganismAndSiteFilter()
      Dim i As Integer
      Dim J As Integer
      Dim K As Integer
      Dim siteName As String
      Dim good As String
      Dim w As Integer

10    On Error GoTo OrganismAndSiteFilter_Error
20    fraCSProgressBar.Visible = True
30    CSProgressBar.Value = 0
40    recordCount = 0
50    With g5
60        For i = 1 To .Rows - 1
              'General
70            If SiteAndOrganismEnabled(.TextMatrix(i, 15), .TextMatrix(i, 17)) = False Then
80                .RowHeight(i) = 0
90                CSProgressBar.Value = CSProgressBar.Value + 1
100               lblCSProgressBar = "Fetching results ... (" & Int(CSProgressBar.Value * 100 / CSProgressBar.Max) & " %)"
110               If CSProgressBar.Value = 100 Then
120                   CSProgressBar.Value = 0
130               End If
140               lblCSProgressBar.Refresh
150           Else
160               .RowHeight(i) = 240
170               CSProgressBar.Value = CSProgressBar.Value + 1
180               lblCSProgressBar = "Fetching results ... (" & Int(CSProgressBar.Value * 100 / CSProgressBar.Max) & " %)"
190               If CSProgressBar.Value = 100 Then
200                   CSProgressBar.Value = 0
210               End If
220               lblCSProgressBar.Refresh
230           End If
240           If i > 0 And .RowHeight(i) > 0 Then
250               recordCount = recordCount + 1
260           End If
270       Next i
280   End With
290   txtResultTotal.Text = recordCount
300   fraCSProgressBar.Visible = False
310   Exit Function

OrganismAndSiteFilter_Error:

      Dim strES As String
      Dim intEL As Integer

320   intEL = Erl
330   strES = Err.Description
340   LogError "frmMicroSurveillanceSearches", "OrganismAndSiteFilter", intEL, strES

End Function

Private Sub cmdExcel_Click()

      Dim strHeading As String

10    On Error GoTo cmdExcel_Click_Error

20    If tabResults.Tab = 0 Then
30        If g.Rows < 2 Then
40            iMsg "Nothing to export"
50            Exit Sub
60        End If
70        strHeading = "List of All General Comments from Micro Surveillance Searches " & vbCr
80        ExpotToExcell 0, strHeading
90    End If
100   If tabResults.Tab = 1 Then
110       If g2.Rows < 2 Then
120           iMsg "No Demographics to export"
130           Exit Sub
140       End If
150       strHeading = "Counts from Micro Surveillance Searches " & vbCr
160       ExpotToExcell 1, strHeading
170   End If
180   If tabResults.Tab = 2 Then
190       If g3.Rows < 2 Then
200           iMsg "No Results to export"
210           Exit Sub
220       End If
230       strHeading = "List of All General Results from Micro Surveillance Searches " & vbCr
240       ExpotToExcell 2, strHeading
250   End If
260   If tabResults.Tab = 4 Then
270       If g5.Rows < 2 Then
280           iMsg "No Results to export"
290           Exit Sub
300       End If
310       strHeading = "List of All Results from Micro Surveillance Searches " & vbCr
320       ExpotToExcell 5, strHeading
330   End If
340   If tabResults.Tab = 6 Then  'Other Counts
350       If G7.Rows < 2 Then
360           iMsg "No Results to export"
370           Exit Sub
380       End If
390       strHeading = "Other Results from Micro Surveillance Searches " & vbCr
400       ExpotToExcell 6, strHeading
410   End If
420   If tabResults.Tab = 7 Then  'General Demographics
430       If G7.Rows < 2 Then
440           iMsg "No Results to export"
450           Exit Sub
460       End If
470       strHeading = "General Results from Micro Surveillance Searches " & vbCr
480       ExpotToExcell 7, strHeading
490   End If

500   If tabResults.Tab = 8 Then  'Specific Ward
510       If GPClinician.Rows < 2 Then
520           iMsg "No Results to export"
530           Exit Sub
540       End If
550       strHeading = "General Results from " & cmbWard.Text & " " & vbCr
560       ExpotToExcell 8, strHeading
570   End If


580   Exit Sub

cmdExcel_Click_Error:

      Dim strES As String
      Dim intEL As Integer

590   intEL = Erl
600   strES = Err.Description
610   LogError "frmMicroSurveillanceSearches", "cmdExcel_Click", intEL, strES
End Sub
Public Function ExpotToExcell(ByVal TabNumber As Integer, strHeading As String)

10    On Error GoTo ExpotToExcell_Error

20    Select Case TabNumber
      Case 0:    'General
30        ExportFlexGrid Me.g, Me, strHeading
40    Case 1:    'Demographics
50        ExportFlexGrid Me.g2, Me, strHeading
60    Case 2:    'Results
70        ExportFlexGrid Me.g3, Me, strHeading
80    Case 5:    'All Results
90        ExportFlexGrid Me.g5, Me, strHeading
100   Case 6:    'Other Results
110       ExportFlexGrid Me.G7, Me, strHeading
120   Case 7:    'Other Results
130       ExportFlexGrid Me.GeneralG, Me, strHeading
140   Case 8:    'Specific Ward Results
150       ExportFlexGrid Me.GPClinician, Me, strHeading
160   End Select

170   Exit Function

ExpotToExcell_Error:

      Dim strES As String
      Dim intEL As Integer

180   intEL = Erl
190   strES = Err.Description
200   LogError "frmMicroSurveillanceSearches", "ExpotToExcell", intEL, strES
End Function

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
      Dim X As Integer
      Dim intLineCount As Integer
      Dim intColCount As Integer

10    On Error GoTo ExportFlexGrid_Error
20    AmountOfCols = objGrid.Cols
30    For X = 0 To objGrid.Cols - 1
40        If objGrid.ColWidth(X) = 0 Then
50            AmountOfCols = AmountOfCols - 1
60        End If
70    Next X
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
180               .range(.cells(R + 1, 1), .cells(R + 1, AmountOfCols)).MergeCells = True
190               .range(.cells(R + 1, 1), .cells(R + 1, AmountOfCols)).HorizontalAlignment = 3
200               .range(.cells(R + 1, 1), .cells(R + 1, AmountOfCols)).Font.Bold = True
210               .range(.cells(R + 1, 1), .cells(R + 1, AmountOfCols)).Font.Size = 16
220               .range(.cells(R + 1, 1), .cells(R + 1, AmountOfCols)).Borders.Weight = 4
230               objWS.cells(R + 1, 1) = "'" & strTokens(R)
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
360                           .range(.cells(R + 1 + intLineCount, 1), .cells(R + 1 + intLineCount, AmountOfCols)).Font.Bold = True
370                           .range(.cells(R + 1 + intLineCount, 1), .cells(R + 1 + intLineCount, AmountOfCols)).WrapText = True
380                           .range(.cells(R + 1, 1), .cells(R + 1, AmountOfCols)).Borders.Weight = 2
390                       End If
400                       objGrid.row = R
410                       objGrid.Col = c
420                       .cells(filterRows, intColCount) = "'" & objGrid.TextMatrix(R, c)
430                       .cells(filterRows, intColCount).Borders.Weight = 2
440                       intColCount = intColCount + 1
450                   End If
460               Next
470           End If
480       Next
490       objXL.ActiveSheet.PageSetup.LeftMargin = 0.25
500       objXL.ActiveSheet.PageSetup.RightMargin = 0.25
510       objXL.ActiveSheet.PageSetup.Orientation = 2
520       objXL.ActiveSheet.PageSetup.Zoom = 60  ' Reduce to 60% when printing
530       .cells.Columns.AutoFit
540       .cells.Columns(14).ColumnWidth = 50
550       .cells.Columns(15).ColumnWidth = 50
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

Private Sub cmdExit_Click()

10        Unload Me

End Sub

Private Sub cmdFilterOrganisms_Click()

10    On Error GoTo cmdFilterOrganisms_Click_Error

20    FillG7


30    OrganismAndSiteFilter
40    getDemographics

50    Exit Sub

cmdFilterOrganisms_Click_Error:

       Dim strES As String
       Dim intEL As Integer

60     intEL = Erl
70     strES = Err.Description
80     LogError "frmMicroSurveillanceSearches", "cmdFilterOrganisms_Click", intEL, strES

End Sub

Private Sub cmdRecalc_Click()
      Dim tb As Recordset
      Dim sql As String
      Dim s As String

10    On Error GoTo cmdRecalc_Click_Error

20    tabResults.Tab = 1

30    g3.Clear
40    GeneralG.Clear
50    GPClinician.Clear
60    g3.Rows = 3
70    InitGrid2
80    InitGrid3
90    InitGeneralG
100   InitGrid7    ' Other Counts by Age Group
110   InitGPClinician
120   cmdRecalc.Enabled = False
130   calFrom = Format(dtFrom, "dd/MMM/yyyy 00:00:00")
140   calTo = Format(dtTo, "dd/MMM/yyyy 23:59:59")

150   FillGeneralG
160   getDemographics
170   FillG
180   fillGFaeses
190   FillG7    'Other Counts by Age Group
200   cmdRecalc.Enabled = True
210   SortG3

220   Exit Sub

cmdRecalc_Click_Error:

      Dim strES As String
      Dim intEL As Integer

230   intEL = Erl
240   strES = Err.Description
250   LogError "frmMicroSurveillanceSearches", "cmdRecalc_Click", intEL, strES, sql

End Sub
Private Sub FillSitesList()
      Dim sql As String
      Dim tb As Recordset

10    On Error GoTo FillSitesList_Error
20        sql = "SELECT Text FROM Lists WHERE ListType = 'MicroSS' and InUse =1 order by ListOrder"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        If tb.EOF Then
60            Exit Sub
70        End If
80        lstSites.Clear
90        While Not tb.EOF
100         lstSites.AddItem tb!Text & ""
110         lstSites.Selected(lstSites.NewIndex) = False
            'lstSites.Selected(lstSites.NewIndex) = True
120         tb.MoveNext
130       Wend
140   Exit Sub

FillSitesList_Error:

      Dim strES As String
      Dim intEL As Integer

150   intEL = Erl
160   strES = Err.Description
170   LogError "frmMicroSurveillanceSearches", "FillSitesList", intEL, strES
End Sub
Private Sub SortG7()
      Dim R As Integer
      Dim c As Integer
      Dim TR As Integer
      Dim TC As Integer
      Dim VisableCols As Integer
      Dim Y As Integer
      Dim NewC As Integer
      Dim NewRow As Integer
      Dim intcol As Integer
      Dim introw As Integer
      Dim LastRecord As Integer

10    On Error GoTo SortG7_Error
20    G7.Visible = False
30    With G7
40        fraCSProgressBar.Visible = True
50        CSProgressBar.Value = 0
60        For c = 3 To .Cols - 1
70            TR = 0
80            For R = 1 To .Rows - 2
90                TR = TR + Val(.TextMatrix(R, c))
100           Next
110           .TextMatrix(R, c) = TR
120       Next
130       For R = 1 To .Rows - 1
140           TC = 0
150           For c = 3 To (.Cols - 2)
160               .row = R
170               .Col = c
180               If Val(.TextMatrix(R, c)) = 0 Then
190                   .CellBackColor = &H8000000F
200               Else
210                   .CellBackColor = vbYellow
220                   TC = TC + Val(.TextMatrix(R, c))
230               End If
240           Next
250           .TextMatrix(R, c) = TC
260           CSProgressBar.Value = CSProgressBar.Value + 1
270           lblCSProgressBar = "Fetching results ... (" & Int(CSProgressBar.Value * 100 / CSProgressBar.Max) & " %)"
280           If CSProgressBar.Value = 100 Then
290               CSProgressBar.Value = 0
300           End If
310           lblCSProgressBar.Refresh
320       Next
330   End With
340   G7.Visible = True
350   fraCSProgressBar.Visible = False
360   Exit Sub

SortG7_Error:

      Dim strES As String
      Dim intEL As Integer

370   intEL = Erl
380   strES = Err.Description
390   LogError "frmMicroSurveillanceSearches", "SortG7", intEL, strES
End Sub
Private Sub fillGFaeses()
      Dim sql As String
      Dim tb As Recordset
      Dim test As String
      Dim Obs As New Observations
      Dim Ob As Observation
      Dim SiteCriteria As String
      Dim Res() As String
      Dim WCC As String
      Dim ScientistComment As String
      Dim ConsultantComment As String
      Dim CommentExists As String
      Dim i As Integer

10    On Error GoTo FillGFaeces_Error
      '<***********************************X = 33
20    G7.Visible = False
30    For i = 0 To lstSites.ListCount - 1
40        If lstSites.Selected(i) Then
50            SiteCriteria = SiteCriteria & "'" & lstSites.List(i) & "" & "',"
60        End If
70    Next
80    If Trim(SiteCriteria) = "" Then Exit Sub
90    SiteCriteria = Left(SiteCriteria, Len(SiteCriteria) - 1)
100   If optType(0) Then
110       sql = "SELECT     d.SampleID, d.SampleDate, d.RecDate, d.Chart,  D.Ward, COALESCE(D.PatName, '') PatName, D.Clinician, D.Age, D.Addr0, D.DoB, D.GP, D.SEX,d.Clinician, M.Site,M.SiteDetails, f.Organism, f.result, f.ValidatedDateTime " & _
                "FROM         MicroSiteDetails M INNER JOIN " & _
                "                Demographics AS d ON M.SampleID = d.SampleID INNER JOIN " & _
                "                    (SELECT     Faeces_2.SampleID, 'rota' AS Organism, Faeces_2.Rota AS result, PrintValidLog.ValidatedDateTime " & _
                "                      FROM          Faeces AS Faeces_2 INNER JOIN " & _
                "                                             PrintValidLog ON Faeces_2.SampleID = PrintValidLog.SampleID " & _
                "                      WHERE      (Faeces_2.Rota IS NOT NULL) AND (PrintValidLog.Department = N'A') " & _
                "                      Union " & _
                "                      SELECT     Faeces_2.SampleID, 'Adeno' AS Organism, Faeces_2.Adeno AS result, PrintValidLog_7.ValidatedDateTime " & _
                "                      FROM         Faeces AS Faeces_2 INNER JOIN " & _
                "                                            PrintValidLog AS PrintValidLog_7 ON Faeces_2.SampleID = PrintValidLog_7.SampleID " & _
                "                      WHERE     (Faeces_2.Adeno IS NOT NULL) AND (PrintValidLog_7.Department = N'A') " & _
                "                      Union " & _
                "                      SELECT     Faeces_2.SampleID, 'Cryptosporidium' AS Organism, Faeces_2.Cryptosporidium AS result, PrintValidLog_6.ValidatedDateTime " & _
                "                      FROM         Faeces AS Faeces_2 INNER JOIN " & _
                "                                            PrintValidLog AS PrintValidLog_6 ON Faeces_2.SampleID = PrintValidLog_6.SampleID " & _
                "                      WHERE     (Faeces_2.Cryptosporidium IS NOT NULL) AND (PrintValidLog_6.Department = N'A') " & _
                "                      Union " & _
                "                      SELECT     Faeces_2.SampleID, 'GiardiaLambila' AS Organism, Faeces_2.GiardiaLambila AS result, PrintValidLog_5.ValidatedDateTime " & _
                "                      FROM         Faeces AS Faeces_2 INNER JOIN " & _
                "                                            PrintValidLog AS PrintValidLog_5 ON Faeces_2.SampleID = PrintValidLog_5.SampleID " & _
                "                      WHERE     (Faeces_2.GiardiaLambila IS NOT NULL) AND (PrintValidLog_5.Department = N'A') " & _
      "                      Union "
120       sql = sql & "               SELECT     Faeces_2.SampleID, 'GDHDetail' AS Organism, Faeces_2.GDHDetail AS result, PrintValidLog_4.ValidatedDateTime " & _
                "                      FROM         Faeces AS Faeces_2 INNER JOIN " & _
                "                                            PrintValidLog AS PrintValidLog_4 ON Faeces_2.SampleID = PrintValidLog_4.SampleID " & _
                "                      WHERE     (Faeces_2.GDHDetail IS NOT NULL) AND (PrintValidLog_4.Department = N'A') " & _
                "                      Union " & _
                "                      SELECT     Faeces_2.SampleID, 'ToxinAB' AS Organism, Faeces_2.ToxinAB AS result, PrintValidLog_3.ValidatedDateTime " & _
                "                      FROM         Faeces AS Faeces_2 INNER JOIN " & _
                "                                            PrintValidLog AS PrintValidLog_3 ON Faeces_2.SampleID = PrintValidLog_3.SampleID" & _
                "                      WHERE     (Faeces_2.ToxinAB IS NOT NULL) AND (PrintValidLog_3.Department = N'A') " & _
                "                      Union " & _
                "                      SELECT     Faeces_2.SampleID, 'pcr' AS Organism, Faeces_2.PCR AS result, PrintValidLog_2.ValidatedDateTime " & _
                "                      FROM         Faeces AS Faeces_2 INNER JOIN " & _
                "                                            PrintValidLog AS PrintValidLog_2 ON Faeces_2.SampleID = PrintValidLog_2.SampleID " & _
                "                      WHERE     (Faeces_2.PCR IS NOT NULL) AND (PrintValidLog_2.Department = N'A') " & _
                "                      Union " & _
                "                      SELECT     Faeces_2.SampleID, 'PCRDetail' AS Organism, Faeces_2.PCRDetail AS result, PrintValidLog_1.ValidatedDateTime " & _
                "                      FROM         Faeces AS Faeces_2 INNER JOIN " & _
                "                                            PrintValidLog AS PrintValidLog_1 ON Faeces_2.SampleID = PrintValidLog_1.SampleID " & _
                "                      WHERE     (Faeces_2.PCRDetail IS NOT NULL) AND (PrintValidLog_1.Department = N'A')) AS f ON M.SampleID = f.SampleID " & _
                "WHERE     (d.SampleDate BETWEEN '" & calFrom & "' and '" & calTo & "') AND (M.Site IN (" & SiteCriteria & "))"
                
                
130             Set tb = New Recordset
140   RecOpenServer 0, tb, sql
150   CSProgressBar.Value = 0
      Dim Site As String
      Dim range As String
160   Do While Not tb.EOF
170       g3.AddItem Format(tb!SampleDate, "dd mmm yyyy") & vbTab & Format(tb!RecDate, "dd mmm yyyy") & vbTab & Format(tb!ValidatedDateTime, "dd mmm yyyy") _
                     & vbTab & tb!SampleID - SysOptMicroOffset(0) & vbTab & tb!Chart & vbTab & tb!PatName & vbTab & tb!Dob _
                     & vbTab & tb!Age & vbTab & tb!sex & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & tb!Addr0 & vbTab _
                     & tb!GP & vbTab & tb!Ward & vbTab & tb!Site & vbTab & tb!SiteDetails & vbTab & tb!Organism & vbTab & tb!Result & vbTab & "" & vbTab & tb!Clinician _
                     & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & ""
180       tb.MoveNext
          
190       CSProgressBar.Value = CSProgressBar.Value + 1
200       lblCSProgressBar = "Fetching results ... (" & Int(CSProgressBar.Value * 100 / CSProgressBar.Max) & " %)"
210       If CSProgressBar.Value = 100 Then
220           CSProgressBar.Value = 0
230       End If
240       lblCSProgressBar.Refresh
250   Loop

260   End If

270   g5.Visible = True
280   Exit Sub

FillGFaeces_Error:

      Dim strES As String
      Dim intEL As Integer

290   intEL = Erl
300   strES = Err.Description
310   LogError "frmMicroSurveillanceSearches", "FillGFaeces", intEL, strES, sql
End Sub
Private Sub SortG5()
      Dim X As Integer
      Dim c As Integer
      Dim VisableCols As Integer
      Dim Y As Integer
      Dim NewC As Integer
      Dim NewRow As Integer
      Dim intcol As Integer
      Dim introw As Integer
      Dim LastRecord As Integer

10    On Error GoTo SortG5_Error
20    For Y = 0 To g3.Cols - 1
30        g5.ColWidth(Y) = 50
40        If g3.ColWidth(Y) > 0 Then
50            VisableCols = VisableCols + 1
60        End If
70    Next Y
80    NewRow = 0
90    For X = 0 To g3.Rows - 1
100       If g3.RowHeight(X) > 0 Then
110           NewC = 0
120           If X > 0 Then
130               NewRow = NewRow + 1
140               If NewRow > 2 Then
150                   g5.Rows = g5.Rows + 1
160               End If
170           End If
180           For c = 0 To g3.Cols - 1
190               If g3.ColWidth(c) > 0 Or c = 12 Or c = 13 Then
200                   g5.TextMatrix(NewRow, NewC) = g3.TextMatrix(X, c)
210                   NewC = NewC + 1
220                   MainProgressBar.Value = MainProgressBar.Value + 1
230                   lblMainProgressBar = "Fetching results ... (" & Int(MainProgressBar.Value * 100 / MainProgressBar.Max) & " %)"
240                   If MainProgressBar.Value = 100 Then
250                       MainProgressBar.Value = 0
260                   End If
270               End If
280           Next c
290       End If
300   Next X
310   With g5
320       .WordWrap = True
330       For intcol = 0 To .Cols - 1
340           For introw = 0 To .Rows - 1
350               If .ColWidth(intcol) < frmMicroSurveillanceSearches.TextWidth(.TextMatrix(introw, intcol)) + 100 Then
360                   .ColWidth(intcol) = frmMicroSurveillanceSearches.TextWidth(.TextMatrix(introw, intcol)) + 100
370               End If
380           Next
390       Next
400   End With
410   For X = 16 To g3.Cols - 1
420       If g5.TextMatrix(0, X) = "" Then
430           LastRecord = X
440           Exit For
450       End If
460   Next X
470   For Y = X To g5.Cols - 1
480       g5.ColWidth(Y) = 0
490   Next Y
500   For X = 1 To 2
510       g5.ColWidth(X) = 1000
520   Next X
530   g5.ColWidth(3) = 900
540   g5.ColWidth(4) = 850
550   g5.ColWidth(23) = 900
560   g5.ColWidth(24) = 900
      'Make columns 9 to 11 not visable
570   For c = 9 To 11
580       g5.ColWidth(c) = 0
590   Next c
600   txtResultTotal = g5.Rows - 1
610   Exit Sub

SortG5_Error:

      Dim strES As String
      Dim intEL As Integer

620   intEL = Erl
630   strES = Err.Description
640   LogError "frmMicroSurveillanceSearches", "SortG5", intEL, strES
End Sub
Private Sub Command_Click(Index As Integer)
      Dim X As Integer
      Dim c As Integer
      Dim VisableCols As Integer
      Dim Y As Integer
      Dim NewC As Integer
      Dim NewRow As Integer
      Dim intcol As Integer
      Dim introw As Integer
      Dim LastRecord As Integer

10    On Error GoTo Command_Click_Error

20    For Y = 0 To g3.Cols - 1
30        g5.ColWidth(Y) = 50
40        If g3.ColWidth(Y) > 0 Then
50            VisableCols = VisableCols + 1
60        End If
70    Next Y
80    NewRow = 0
90    For X = 0 To g3.Rows - 1
100       If g3.RowHeight(X) > 0 Then
110           NewC = 0
120           If X > 0 Then
130               NewRow = NewRow + 1
140               If NewRow > 2 Then
150                   g5.Rows = g5.Rows + 1
160               End If
170           End If
180           For c = 0 To g3.Cols - 1
190               If g3.ColWidth(c) > 0 Then
200                   g5.TextMatrix(NewRow, NewC) = g3.TextMatrix(X, c)
210                   NewC = NewC + 1
220               End If
230           Next c
240       End If
250   Next X
260   With g5
270       For intcol = 0 To .Cols - 1
280           For introw = 0 To .Rows - 1
290               If .ColWidth(intcol) < frmMicroSurveillanceSearches.TextWidth(.TextMatrix(introw, intcol)) + 100 And intcol <> 1 And intcol <> 2 Then
300                   .ColWidth(intcol) = frmMicroSurveillanceSearches.TextWidth(.TextMatrix(introw, intcol)) + 100
310               End If
320           Next
330       Next
340   End With
350   For X = 12 To g3.Cols - 1
360       If g5.TextMatrix(0, X) = "" Then
370           LastRecord = X
380           Exit For
390       End If
400   Next X
410   For Y = X To g5.Cols - 1
420       g5.ColWidth(Y) = 0
430   Next Y
440   g5.RowHeight(0) = g5.RowHeight(0) * 2
450   g5.ColWidth(1) = 850
460   g5.ColWidth(2) = 700
470   g5.WordWrap = True

480   Exit Sub

Command_Click_Error:

      Dim strES As String
      Dim intEL As Integer

490   intEL = Erl
500   strES = Err.Description
510   LogError "frmMicroSurveillanceSearches", "Command_Click", intEL, strES

End Sub

Private Sub Form_Load()
      Dim n As Integer

10    On Error GoTo Form_Load_Error

20    With frmMicroSurveillanceSearches
30        .Top = (Screen.Height - .Height) / 2
40        .Left = (Screen.Width - .Width) / 2
50        .tabHospitals.TabCaption(0) = HospName(0)
60        .tabHospitals.TabCaption(1) = HospName(1)
70        .tabHospitals.TabCaption(2) = HospName(2)
80        .tabResults.TabVisible(0) = False
90        .tabResults.TabVisible(2) = False '
100       .tabResults.TabVisible(3) = False 'other counts by sample id
110       .tabResults.TabVisible(4) = False
120       .tabResults.TabVisible(5) = False
130   End With
140   For n = 0 To 4
150       AgeGroupFrom(n) = Choose(n + 1, 0, 4 * 365, 14 * 365, 44 * 365, 60 * 365)
160       AgeGroupTo(n) = Choose(n + 1, 4 * 365, 14 * 365, 44 * 365, 60 * 365, 43830)
170   Next
180   dtTo = Format(Now, "dd/MM/yyyy")
190   dtFrom = Format(Now, "dd/MM/yyyy")
200   InitGrid
210   InitGrid2
      '200   InitGrid3
220   InitGrid4 ' other counts by sample
230   InitGrid7 'Other Counts by Age Group
240   InitGPClinician
250   FillSitesList
260   g.Visible = False
270   g5.RowHeight(0) = g5.RowHeight(0) * 2

280   Exit Sub

Form_Load_Error:

      Dim strES As String
      Dim intEL As Integer

290   intEL = Erl
300   strES = Err.Description
310   LogError "frmMicroSurveillanceSearches", "Form_Load", intEL, strES
End Sub

Private Sub g_DblClick()
      Dim R As Long
      Dim c As Long

10    On Error GoTo g_DblClick_Error

20    R = g.MouseRow
30    c = g.MouseCol
40    g.row = R
50    g.Col = c
60    If g.CellBackColor = &H80& Then
70        frmMicroSurveillanceComments.SampleID = g.TextMatrix(R, 0) + SysOptMicroOffset(0)
80        frmMicroSurveillanceComments.Show 1
90    End If

100   Exit Sub

g_DblClick_Error:

      Dim strES As String
      Dim intEL As Integer

110   intEL = Erl
120   strES = Err.Description
130   LogError "frmMicroSurveillanceSearches", "g_DblClick", intEL, strES
End Sub

'---------------------------------------------------------------------------------------
' Procedure : g2_MouseDown
' Author    : Trevor Dunican
' Date      : 11/12/2014
' Purpose   : Collapse each site when clicked
'---------------------------------------------------------------------------------------
'
Private Sub g2_MouseDown(Button As Integer, Shift As Integer, w As Single, z As Single)
      Dim s As String
      Dim n As Integer
      Dim X As Integer
      Dim Y As Integer
      Dim StartIndex As Integer
      Dim EndIndex As Integer
      Dim R As Integer
      Dim c As Integer

10    On Error GoTo g2_MouseDown_Error

20    R = g2.MouseRow
30    c = g2.MouseCol
40    g2.row = R
50    g2.Col = c
60    If c = 0 And R > 0 Then
70        For X = R To g2.Rows - 1
80            If g2.TextMatrix(X, 0) = "" Then Exit For
90            StartIndex = X
100           For Y = StartIndex + 1 To g2.Rows - 1
110               If g2.TextMatrix(Y, 0) <> "" Then
120                   EndIndex = Y - 1
130                   Exit For
140               End If
150           Next Y
160           If StartIndex > EndIndex Then
170               EndIndex = Y - 1
180           End If
190           If g2.RowHeight(StartIndex + 1) > 0 Then
200               For n = StartIndex + 1 To EndIndex
210                   g2.RowHeight(n) = 0
220               Next n
230               Exit Sub
240           Else
250               For n = StartIndex + 1 To EndIndex
260                   g2.RowHeight(n) = 275
270               Next n
280               Exit Sub
290           End If
300       Next X
310   End If

320   Exit Sub

g2_MouseDown_Error:

      Dim strES As String
      Dim intEL As Integer

330   intEL = Erl
340   strES = Err.Description
350   LogError "frmMicroSurveillanceSearches", "g2_MouseDown", intEL, strES

End Sub
Private Sub SortG3()
      Dim samID As String
      Dim n As Integer
      Dim X As Integer
      Dim Y As Integer
      Dim StartIndex As Integer
      Dim EndIndex As Integer
      Dim row As Integer
      Dim column As Integer
      Dim Day As String
      Dim A As Integer
      Dim AntiBody As String
      Dim f As Integer
      Dim FirstCheck As Boolean
      Dim OrganismType As String
      Dim RecordNumber As Integer
      Dim LastRecord As Integer

10    On Error GoTo SortG3_Error

20    prgFetchingResults.Value = 0
30    fraMainProgressBar.Visible = True
40    Day = g3.TextMatrix(1, 0)
50    samID = g3.TextMatrix(1, 1)
60    OrganismType = g3.TextMatrix(1, 17)
70    A = 25
80    FirstCheck = True
90    RecordNumber = 1
      'recordCount = 0
100   For X = 1 To g3.Rows - 1
110       If g3.TextMatrix(X, 1) = samID And g3.TextMatrix(X, 0) = Day And g3.TextMatrix(X, 17) = OrganismType Then
120           AntiBody = g3.TextMatrix(X, 9)
130           For f = 23 To A Step 2
140               If g3.TextMatrix(0, f) = AntiBody Then
                      '***********************************************
150                   g3.TextMatrix(RecordNumber, f) = g3.TextMatrix(X, 10)
160                   g3.TextMatrix(RecordNumber, f + 1) = g3.TextMatrix(X, 11)

170                   Exit For
180               End If
190               If g3.TextMatrix(0, f) = "" And g3.TextMatrix(0, f) <> AntiBody Then
200                   g3.TextMatrix(0, f) = AntiBody    'g3.TextMatrix(n, 9)
210                   g3.TextMatrix(RecordNumber, f) = g3.TextMatrix(X, 10)
220                   g3.TextMatrix(RecordNumber, f + 1) = g3.TextMatrix(X, 11)
230                   A = A + 2
240                   Exit For
250               End If
260           Next f
270           If Not FirstCheck Then
280               g3.RowHeight(X) = 0
290           End If
300           If A > 25 Then
310               FirstCheck = False
320           End If
330       Else
340           Day = g3.TextMatrix(X, 0)
350           samID = g3.TextMatrix(X, 1)
360           OrganismType = g3.TextMatrix(X, 17)
370           FirstCheck = True
380           RecordNumber = X
390           X = X - 1
400       End If
410       If X = g3.Rows - 1 Then
420           Exit For
430       End If
440       MainProgressBar.Value = MainProgressBar.Value + 1
450       lblMainProgressBar = "Fetching results ... (" & Int(MainProgressBar.Value * 100 / MainProgressBar.Max) & " %)"
460       If MainProgressBar.Value = 100 Then
470           MainProgressBar.Value = 0
480       End If
490       lblMainProgressBar.Refresh
500   Next X
510   fraMainProgressBar.Visible = False
520   For n = 23 To g3.Cols - 1
530       If g3.TextMatrix(0, n) = "" Then
540           LastRecord = n
550           Exit For
560       End If
570   Next n
580   For Y = n To g3.Cols - 1
590       g3.ColWidth(Y) = 0
600   Next Y
      'Make columns 9 to 13 not visable
610   For n = 9 To 13
620       g5.ColWidth(n) = 0
630   Next n

640   Exit Sub

SortG3_Error:

      Dim strES As String
      Dim intEL As Integer

650   intEL = Erl
660   strES = Err.Description
670   LogError "frmMicroSurveillanceSearches", "SortG3", intEL, strES
End Sub

Private Sub g5_DblClick()
      Dim R As Long
      Dim c As Long

10    On Error GoTo g5_DblClick_Error

20    R = g5.MouseRow
30    c = g5.MouseCol
40    g5.row = R
50    g5.Col = c
60    If g5.TextMatrix(R, c) = "X" Then
70        frmMicroSurveillanceComments.SampleID = g5.TextMatrix(R, 3) + SysOptMicroOffset(0)
80        frmMicroSurveillanceComments.Show 1
90    End If

100   Exit Sub

g5_DblClick_Error:

      Dim strES As String
      Dim intEL As Integer

110   intEL = Erl
120   strES = Err.Description
130   LogError "frmMicroSurveillanceSearches", "g5_DblClick", intEL, strES
End Sub

Private Sub g5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

10    With g5
20        .ToolTipText = .TextMatrix(.MouseRow, .MouseCol)
30    End With

End Sub



Private Sub lstOrganismGroup_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
10    With lstOrganismGroup
20        .ToolTipText = .Text
30    End With
End Sub

Private Sub lstOrganismGroup_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
10    OrganismAndSiteFilter
End Sub

Private Sub lstSites_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
10    With lstSites
20        .ToolTipText = .Text
30    End With
End Sub

Private Sub lstSites_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

      'calFrom = Format(dtFrom, "dd/MMM/yyyy 00:00:00")
      'calTo = Format(dtTo, "dd/MMM/yyyy 23:59:59")

'10    FillG7
'20    SortG7
'
'30    OrganismAndSiteFilter
'40    getDemographics

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
80    LogError "frmMicroSurveillanceSearches", "optBetween_Click", intEL, strES

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

BetweenDates_Error:

      Dim strES As String
      Dim intEL As Integer

320   intEL = Erl
330   strES = Err.Description
340   LogError "frmMicroSurveillanceSearches", "BetweenDates", intEL, strES

End Function

Private Sub FillG()

      Dim sql As String
      Dim tb As Recordset
      Dim test As String
      Dim Obs As New Observations
      Dim Ob As Observation
      Dim SiteCriteria As String
      Dim Res() As String
      Dim WCC As String
      Dim ScientistComment As String
      Dim ConsultantComment As String
      Dim CommentExists As String
      Dim i As Integer


10    On Error GoTo FillG_Error
      '<***********************************X = 33
20    g.Visible = False
30    For i = 0 To lstSites.ListCount - 1
40        If lstSites.Selected(i) Then
50            SiteCriteria = SiteCriteria & "'" & lstSites.List(i) & "" & "',"
60        End If
70    Next
80    If Trim(SiteCriteria) = "" Then Exit Sub
90    SiteCriteria = Left(SiteCriteria, Len(SiteCriteria) - 1)
100   If optType(0) Then
110       sql = "SELECT I.SampleID, D.Chart, D.Ward, COALESCE(D.PatName, '') PatName, D.Clinician, D.Age, D.Addr0, D.DoB, D.GP, D.SEX, U.WCC, M.Site, M.SiteDetails, I.OrganismGroup, I.OrganismName, S.Result, S.AntibioticCode, " & _
                "S.RSI, I.Qualifier, D.SampleDate, D.RecDate, P.ValidatedDateTime from Isolates I " & _
                "Inner Join MicroSiteDetails M ON I.SampleID = M.SampleID " & _
                "Left Join (SELECT * FROM PrintValidLog WHERE Department = 'D') P ON M.SampleID = P.SampleID " & _
                "Left Join Urine U ON M.SampleID = U.SampleID " & _
                "Inner Join Demographics D ON I.SampleID = D.SampleID " & _
                "Left Join Sensitivities S ON I.SampleID = S.SampleID  And I.IsolateNumber = S.IsolateNumber " & _
                "WHERE D.SampleDate between '" & calFrom & "' and '" & calTo & "' and Site in (" & SiteCriteria & ") " & _
                "ORDER by D.SampleDate DESC, I.SampleID, Site"
120   ElseIf optType(1) Then
130       sql = "SELECT I.SampleID, D.Chart, D.Ward, COALESCE(D.PatName, '') PatName, D.Clinician, D.Age, D.Addr0, D.DoB, D.GP, D.SEX, U.WCC, M.Site, M.SiteDetails, I.OrganismGroup, I.OrganismName, S.Result, S.AntibioticCode, " & _
                "S.RSI, I.Qualifier, D.SampleDate, D.RecDate, P.ValidatedDateTime from Isolates I " & _
                "Inner Join MicroSiteDetails M ON I.SampleID = M.SampleID " & _
                "Left Join (SELECT * FROM PrintValidLog WHERE Department = 'D') P ON M.SampleID = P.SampleID " & _
                "Left Join Urine U ON M.SampleID = U.SampleID " & _
                "Inner Join Demographics D ON I.SampleID = D.SampleID " & _
                "LEft Join Sensitivities S ON I.SampleID = S.SampleID  And I.IsolateNumber = S.IsolateNumber " & _
                "WHERE OrganismGroup <> 'Negative Results' and D.SampleDate between '" & calFrom & "' and '" & calTo & "' and Site in (" & SiteCriteria & ") " & _
                "ORDER by D.SampleDate DESC, I.SampleID, Site"
140   ElseIf optType(2) Then
150       sql = "SELECT I.SampleID, D.Chart, D.Ward, COALESCE(D.PatName, '') PatName, D.Clinician, D.Age, D.Addr0, D.DoB, D.GP, D.SEX, U.WCC, M.Site, M.SiteDetails, I.OrganismGroup, I.OrganismName, S.Result, S.AntibioticCode, " & _
                "S.RSI, I.Qualifier, D.SampleDate, D.RecDate, P.ValidatedDateTime from Isolates I " & _
                "Inner Join MicroSiteDetails M ON I.SampleID = M.SampleID " & _
                "Left Join (SELECT * FROM PrintValidLog WHERE Department = 'D') P ON M.SampleID = P.SampleID " & _
                "Left Join Urine U ON M.SampleID = U.SampleID " & _
                "Inner Join Demographics D ON I.SampleID = D.SampleID " & _
                "Left Join Sensitivities S ON I.SampleID = S.SampleID  And I.IsolateNumber = S.IsolateNumber " & _
                "WHERE OrganismGroup = 'Negative Results' and D.SampleDate between '" & calFrom & "' and '" & calTo & "' and Site in (" & SiteCriteria & ") " & _
                "ORDER by D.SampleDate DESC, I.SampleID, Site"
160   End If
170   Set tb = New Recordset
180   RecOpenServer 0, tb, sql
190   g.Clear
200   InitGrid
210   lstOrganismGroup.Clear
220   g.Rows = 1
230   g3.Rows = 1
240   fraCSProgressBar.Visible = True
250   CSProgressBar.Value = 0
260   Do While Not tb.EOF
270       CommentExists = ""
280       ScientistComment = ""
290       ConsultantComment = ""
300       Res = Split(tb!WCC & "", "|")
310       If UBound(Res) = -1 Then
320           WCC = tb!WCC & ""
330       ElseIf UBound(Res) > 1 Then
340           WCC = Res(0)
350       End If

360       Set Obs = Obs.Load(tb!SampleID, _
                             "MicroCS", "MicroConsultant")
370       If Not Obs Is Nothing Then
380           If Obs.Count > 0 Then
390               For Each Ob In Obs
400                   Select Case Ob.Discipline
                      Case "MicroCS"
410                       ScientistComment = Ob.Comment
420                   Case "MicroConsultant"
430                       ConsultantComment = Ob.Comment
440                   End Select
450                   CommentExists = "X"
460               Next
470           End If
480           g.AddItem tb!SampleID - SysOptMicroOffset(0) & vbTab & tb!Site & vbTab & tb!OrganismGroup & "" _
                        & vbTab & tb!OrganismName & vbTab & tb!Qualifier & vbTab & tb!Result & vbTab & tb!AntibioticCode _
                        & vbTab & tb!RSI & vbTab & Format(tb!SampleDate, "dd mmm yyyy") & ""

490           g3.AddItem Format(tb!SampleDate, "dd mmm yyyy") & vbTab & Format(tb!RecDate, "dd mmm yyyy") & vbTab & Format(tb!ValidatedDateTime, "dd mmm yyyy") _
                         & vbTab & tb!SampleID - SysOptMicroOffset(0) & vbTab & tb!Chart & vbTab & tb!PatName & vbTab & tb!Dob _
                         & vbTab & tb!Age & vbTab & tb!sex & vbTab & tb!AntibioticCode & vbTab & tb!Result & vbTab & tb!RSI & vbTab & tb!Addr0 & vbTab _
                         & tb!GP & vbTab & tb!Ward & vbTab & tb!Site & vbTab & tb!SiteDetails & vbTab & tb!OrganismName & vbTab & "" & vbTab & WCC & vbTab & tb!Clinician _
                         & vbTab & tb!Qualifier & vbTab & "" & vbTab & "" & vbTab & ScientistComment & vbTab & ConsultantComment & ""
500           If CommentExists = "X" Then
510               g.row = g.Rows - 1
520               g.Col = 9
530               g.CellBackColor = &H80&
540               g3.row = g3.Rows - 1
550               g3.Col = 22
560               g3.Text = CommentExists
570           End If
              '*******************************************************************************************
580       End If

590       If Not OrganismExistsInList(tb!OrganismName & "") Then
600           lstOrganismGroup.AddItem tb!OrganismName & ""
610           lstOrganismGroup.Selected(lstOrganismGroup.NewIndex) = True
620       End If
630       CSProgressBar.Value = CSProgressBar.Value + 1
640       lblCSProgressBar = "Fetching results ... (" & Int(CSProgressBar.Value * 100 / CSProgressBar.Max) & " %)"
650       If CSProgressBar.Value = 100 Then
660           CSProgressBar.Value = 0
670       End If
680       lblCSProgressBar.Refresh
690       tb.MoveNext
700   Loop
710   fillGFaeses
720   getDemographics
730   fraCSProgressBar.Visible = False
      Dim introw As Integer
      Dim intcol As Integer
740   With g
750       For intcol = 0 To .Cols - 1
760           For introw = 0 To .Rows - 1
770               If .ColWidth(intcol) < frmMicroSurveillanceSearches.TextWidth(.TextMatrix(introw, intcol)) + 100 Then
780                   .ColWidth(intcol) = frmMicroSurveillanceSearches.TextWidth(.TextMatrix(introw, intcol)) + 100
790               End If
800               If g3.ColWidth(intcol) < frmMicroSurveillanceSearches.TextWidth(g3.TextMatrix(introw, intcol)) + 100 And intcol <> 1 And intcol <> 2 Then
810                   g3.ColWidth(intcol) = frmMicroSurveillanceSearches.TextWidth(g3.TextMatrix(introw, intcol)) + 100
820               End If
830           Next
840       Next
850   End With
860   g.Visible = True

870   Exit Sub

FillG_Error:

      Dim strES As String
      Dim intEL As Integer

880   intEL = Erl
890   strES = Err.Description
900   LogError "frmMicroSurveillanceSearches", "FillG", intEL, strES, sql

End Sub

Private Sub FillGeneralG()

          Dim sql As String
          Dim sql1 As String
          Dim tb As Recordset
          Dim test As String
          Dim Obs As New Observations
          Dim Ob As Observation
          Dim SiteCriteria As String
          Dim Res() As String
          Dim WCC As String
          Dim ScientistComment As String
          Dim ConsultantComment As String
          Dim CommentExists As String
          Dim i As Integer
          Dim LastSampleId As String
          Dim LastOrgGroup As String
          Dim MaxColOfAntibiotics As Integer

10        On Error GoTo FillGeneral_Error
          '<***********************************X = 33
20        g.Visible = False
30        For i = 0 To lstSites.ListCount - 1
40            If lstSites.Selected(i) Then
50                SiteCriteria = SiteCriteria & "'" & lstSites.List(i) & "" & "',"
60            End If
70        Next
80        If Trim(SiteCriteria) = "" Then Exit Sub
90        SiteCriteria = Left(SiteCriteria, Len(SiteCriteria) - 1)
100       If optType(0) Then
110           sql = "    SELECT TOP (100) PERCENT PI.PatMobile, M.SampleID AS SampleId, D.Chart, D.Ward, COALESCE (D.PatName, '') AS PatName, D.Clinician, D.Age, D.Addr0, D.DoB, D.GP, D.Sex, U.WCC, M.Site, M.SiteDetails, " & _
                  "    I.OrganismGroup, I.OrganismName, S.Result, S.AntibioticCode, S.RSI, I.Qualifier, D.SampleDate, D.RecDate, P.ValidatedDateTime, U.Protein, U.Glucose, U.RCC, " & _
                  "    U.Crystals , Faeces.Rota, Faeces.Adeno, Faeces.Cryptosporidium, Faeces.GiardiaLambila, Faeces.GDH, Faeces.ToxinAB, Faeces.PCR , Faeces.Gram , CSFResults.Appearance0, CSFResults.WCC0, "
120           sql = sql & "    (SELECT     Result  FROM GenericResults  WHERE      (TestName = 'FluidAppearance0') AND (SampleID = M.SampleID)) AS CellCount, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'FluidAppearance1') AND (SampleID = M.SampleID)) AS Appearance, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'FluidGram') AND (SampleID = M.SampleID)) AS GramStain1, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'FluidGram(2)') AND (SampleID = M.SampleID)) AS GramStain2, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'FluidZN') AND (SampleID = M.SampleID)) AS ZNstain, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'FluidLeishmans') AND (SampleID = M.SampleID)) AS LeishmansStain, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'FluidWetPrep') AND (SampleID = M.SampleID)) AS WetPrep, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'FluidCrystals') AND (SampleID = M.SampleID)) AS FCrystals, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'CSFHaem0') AND (SampleID = M.SampleID)) AS RCC1, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'CSFHaem1') AND (SampleID = M.SampleID)) AS RCC2, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'CSFHaem2') AND (SampleID = M.SampleID)) AS RCC3, "
130           sql = sql & "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'CSFHaem3') AND (SampleID = M.SampleID)) AS WCC1, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'CSFHaem4') AND (SampleID = M.SampleID)) AS WCC2,  " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'CSFHaem5') AND (SampleID = M.SampleID)) AS WCC3, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'CSFHaem6') AND (SampleID = M.SampleID)) AS Polymorphic1, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'CSFHaem7') AND (SampleID = M.SampleID)) AS Polymorphic2, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'CSFHaem8') AND (SampleID = M.SampleID)) AS Polymorphic3, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'CSFHaem9') AND (SampleID = M.SampleID)) AS Mononucleated1, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'CSFHaem10') AND (SampleID = M.SampleID)) AS Mononucleated2, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'CSFHaem11') AND (SampleID = M.SampleID)) AS Mononucleated3, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'FluidGlucose') AND (SampleID = M.SampleID)) AS FluidGlucose, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'FluidProtein') AND (SampleID = M.SampleID)) AS FluidProtein, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'FluidAlbumin') AND (SampleID = M.SampleID)) AS FluidAlbumin, "
140           sql = sql & "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'FluidGlobulin') AND (SampleID = M.SampleID)) AS FluidGlobulin, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'FluidLDH') AND (SampleID = M.SampleID)) AS FluidLDH, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'FluidAmylase') AND (SampleID = M.SampleID)) AS FluidAmylase, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'CSFGlucose') AND (SampleID = M.SampleID)) AS CSFGlucose, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'CSFProtein') AND (SampleID = M.SampleID)) AS CSFProtein, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'FungalElements') AND (SampleID = M.SampleID)) AS FungalElements, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'PneumococcalAT') AND (SampleID = M.SampleID)) AS PneumococcalAT, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'LegionellaAT') AND (SampleID = M.SampleID)) AS LegionellaAT, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'BATScreen') AND (SampleID = M.SampleID)) AS BATScreen, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'BATScreenComment') AND (SampleID = M.SampleID)) AS BATScreenComment "
150           sql = sql & " FROM  MicroSiteDetails AS M " & _
                  "    INNER JOIN Demographics AS D ON M.SampleID = D.SampleID " & _
                  "    LEFT JOIN PatientIfs AS PI ON D.Chart = PI.Chart " & _
                  "    LEFT OUTER JOIN  (SELECT  *  From PrintValidLog WHERE (Department = 'D') OR (Department = 'A')) AS P ON M.SampleID = P.SampleID " & _
                  "    LEFT OUTER JOIN CSFResults ON M.SampleID = CSFResults.SampleID " & _
                  "    LEFT OUTER JOIN Isolates AS I " & _
                  "    LEFT JOIN Sensitivities AS S ON I.SampleID = S.SampleID AND I.IsolateNumber = S.IsolateNumber ON M.SampleID = I.SampleID " & _
                  "    LEFT OUTER JOIN Urine AS U ON M.SampleID = U.SampleID " & _
                  "    LEFT OUTER JOIN Faeces ON M.SampleID = Faeces.SampleID " & _
                  "    WHERE (D.SampleDate BETWEEN '" & calFrom & "' and '" & calTo & "') AND (M.Site IN (" & SiteCriteria & ")) " & _
                  "    ORDER BY D.SampleDate DESC, I.SampleID, M.Site"
160       ElseIf optType(1) Then
170           sql = "    SELECT TOP (100) PERCENT M.SampleID AS SampleId, D.Chart, D.Ward, COALESCE (D.PatName, '') AS PatName, D.Clinician, D.Age, D.Addr0, D.DoB, D.GP, D.Sex, U.WCC, M.Site, M.SiteDetails, " & _
                  "    I.OrganismGroup, I.OrganismName, S.Result, S.AntibioticCode, S.RSI, I.Qualifier, D.SampleDate, D.RecDate, P.ValidatedDateTime, U.Protein, U.Glucose, U.RCC, " & _
                  "    U.Crystals , Faeces.Rota, Faeces.Adeno, Faeces.Cryptosporidium, Faeces.GiardiaLambila, Faeces.GDH, Faeces.ToxinAB, Faeces.PCR , Faeces.Gram ,CSFResults.Appearance0, CSFResults.WCC0 , "
180           sql = sql & "    (SELECT     Result  FROM GenericResults  WHERE      (TestName = 'FluidAppearance0') AND (SampleID = M.SampleID)) AS CellCount, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'FluidAppearance1') AND (SampleID = M.SampleID)) AS Appearance, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'FluidGram') AND (SampleID = M.SampleID)) AS GramStain1, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'FluidGram(2)') AND (SampleID = M.SampleID)) AS GramStain2, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'FluidZN') AND (SampleID = M.SampleID)) AS ZNstain, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'FluidLeishmans') AND (SampleID = M.SampleID)) AS LeishmansStain, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'FluidWetPrep') AND (SampleID = M.SampleID)) AS WetPrep, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'FluidCrystals') AND (SampleID = M.SampleID)) AS FCrystals, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'CSFHaem0') AND (SampleID = M.SampleID)) AS RCC1, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'CSFHaem1') AND (SampleID = M.SampleID)) AS RCC2, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'CSFHaem2') AND (SampleID = M.SampleID)) AS RCC3, "
190           sql = sql & "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'CSFHaem3') AND (SampleID = M.SampleID)) AS WCC1, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'CSFHaem4') AND (SampleID = M.SampleID)) AS WCC2,  " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'CSFHaem5') AND (SampleID = M.SampleID)) AS WCC3, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'CSFHaem6') AND (SampleID = M.SampleID)) AS Polymorphic1, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'CSFHaem7') AND (SampleID = M.SampleID)) AS Polymorphic2, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'CSFHaem8') AND (SampleID = M.SampleID)) AS Polymorphic3, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'CSFHaem9') AND (SampleID = M.SampleID)) AS Mononucleated1, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'CSFHaem10') AND (SampleID = M.SampleID)) AS Mononucleated2, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'CSFHaem11') AND (SampleID = M.SampleID)) AS Mononucleated3, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'FluidGlucose') AND (SampleID = M.SampleID)) AS FluidGlucose, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'FluidProtein') AND (SampleID = M.SampleID)) AS FluidProtein, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'FluidAlbumin') AND (SampleID = M.SampleID)) AS FluidAlbumin, "
200           sql = sql & "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'FluidGlobulin') AND (SampleID = M.SampleID)) AS FluidGlobulin, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'FluidLDH') AND (SampleID = M.SampleID)) AS FluidLDH, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'FluidAmylase') AND (SampleID = M.SampleID)) AS FluidAmylase, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'CSFGlucose') AND (SampleID = M.SampleID)) AS CSFGlucose, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'CSFProtein') AND (SampleID = M.SampleID)) AS CSFProtein, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'FungalElements') AND (SampleID = M.SampleID)) AS FungalElements, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'PneumococcalAT') AND (SampleID = M.SampleID)) AS PneumococcalAT, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'LegionellaAT') AND (SampleID = M.SampleID)) AS LegionellaAT, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'BATScreen') AND (SampleID = M.SampleID)) AS BATScreen, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'BATScreenComment') AND (SampleID = M.SampleID)) AS BATScreenComment "
210           sql = sql & " FROM  MicroSiteDetails AS M " & _
                  "    INNER JOIN Demographics AS D ON M.SampleID = D.SampleID " & _
                  "    INNER JOIN  (SELECT  *  From PrintValidLog WHERE (Department = 'D') OR (Department = 'A')) AS P ON M.SampleID = P.SampleID " & _
                  "    LEFT OUTER JOIN CSFResults ON P.SampleID = CSFResults.SampleID " & _
                  "    LEFT OUTER JOIN Isolates AS I " & _
                  "    LEFT JOIN Sensitivities AS S ON I.SampleID = S.SampleID AND I.IsolateNumber = S.IsolateNumber ON M.SampleID = I.SampleID " & _
                  "    LEFT OUTER JOIN Urine AS U ON M.SampleID = U.SampleID " & _
                  "    LEFT OUTER JOIN Faeces ON M.SampleID = Faeces.SampleID " & _
                  "    WHERE (OrganismGroup <> 'Negative Results' and D.SampleDate BETWEEN '" & calFrom & "' and '" & calTo & "') AND (M.Site IN (" & SiteCriteria & ")) " & _
                  "    ORDER BY D.SampleDate DESC, I.SampleID, M.Site"
220       ElseIf optType(2) Then
230           sql = "    SELECT TOP (100) PERCENT M.SampleID AS SampleId, D.Chart, D.Ward, COALESCE (D.PatName, '') AS PatName, D.Clinician, D.Age, D.Addr0, D.DoB, D.GP, D.Sex, U.WCC, M.Site, M.SiteDetails, " & _
                  "    I.OrganismGroup, I.OrganismName, S.Result, S.AntibioticCode, S.RSI, I.Qualifier, D.SampleDate, D.RecDate, P.ValidatedDateTime, U.Protein, U.Glucose, U.RCC, " & _
                  "    U.Crystals , Faeces.Rota, Faeces.Adeno, Faeces.Cryptosporidium, Faeces.GiardiaLambila, Faeces.GDH, Faeces.ToxinAB, Faeces.PCR , Faeces.Gram , CSFResults.Appearance0, CSFResults.WCC0 , "
240           sql = sql & "    (SELECT     Result  FROM GenericResults  WHERE      (TestName = 'FluidAppearance0') AND (SampleID = M.SampleID)) AS CellCount, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'FluidAppearance1') AND (SampleID = M.SampleID)) AS Appearance, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'FluidGram') AND (SampleID = M.SampleID)) AS GramStain1, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'FluidGram(2)') AND (SampleID = M.SampleID)) AS GramStain2, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'FluidZN') AND (SampleID = M.SampleID)) AS ZNstain, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'FluidLeishmans') AND (SampleID = M.SampleID)) AS LeishmansStain, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'FluidWetPrep') AND (SampleID = M.SampleID)) AS WetPrep, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'FluidCrystals') AND (SampleID = M.SampleID)) AS FCrystals, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'CSFHaem0') AND (SampleID = M.SampleID)) AS RCC1, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'CSFHaem1') AND (SampleID = M.SampleID)) AS RCC2, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'CSFHaem2') AND (SampleID = M.SampleID)) AS RCC3, "
250           sql = sql & "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'CSFHaem3') AND (SampleID = M.SampleID)) AS WCC1, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'CSFHaem4') AND (SampleID = M.SampleID)) AS WCC2,  " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'CSFHaem5') AND (SampleID = M.SampleID)) AS WCC3, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'CSFHaem6') AND (SampleID = M.SampleID)) AS Polymorphic1, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'CSFHaem7') AND (SampleID = M.SampleID)) AS Polymorphic2, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'CSFHaem8') AND (SampleID = M.SampleID)) AS Polymorphic3, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'CSFHaem9') AND (SampleID = M.SampleID)) AS Mononucleated1, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'CSFHaem10') AND (SampleID = M.SampleID)) AS Mononucleated2, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'CSFHaem11') AND (SampleID = M.SampleID)) AS Mononucleated3, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'FluidGlucose') AND (SampleID = M.SampleID)) AS FluidGlucose, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'FluidProtein') AND (SampleID = M.SampleID)) AS FluidProtein, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'FluidAlbumin') AND (SampleID = M.SampleID)) AS FluidAlbumin, "
260           sql = sql & "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'FluidGlobulin') AND (SampleID = M.SampleID)) AS FluidGlobulin, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'FluidLDH') AND (SampleID = M.SampleID)) AS FluidLDH, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'FluidAmylase') AND (SampleID = M.SampleID)) AS FluidAmylase, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'CSFGlucose') AND (SampleID = M.SampleID)) AS CSFGlucose, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'CSFProtein') AND (SampleID = M.SampleID)) AS CSFProtein, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'FungalElements') AND (SampleID = M.SampleID)) AS FungalElements, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'PneumococcalAT') AND (SampleID = M.SampleID)) AS PneumococcalAT, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'LegionellaAT') AND (SampleID = M.SampleID)) AS LegionellaAT, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'BATScreen') AND (SampleID = M.SampleID)) AS BATScreen, " & _
                  "    (SELECT     Result  From GenericResults  WHERE      (TestName = 'BATScreenComment') AND (SampleID = M.SampleID)) AS BATScreenComment "
270           sql = sql & " FROM  MicroSiteDetails AS M " & _
                  "    INNER JOIN Demographics AS D ON M.SampleID = D.SampleID " & _
                  "    INNER JOIN  (SELECT  *  From PrintValidLog WHERE (Department = 'D') OR (Department = 'A')) AS P ON M.SampleID = P.SampleID " & _
                  "    LEFT JOIN CSFResults ON M.SampleID = CSFResults.SampleID " & _
                  "    LEFT JOIN Isolates AS I " & _
                  "    LEFT JOIN Sensitivities AS S ON I.SampleID = S.SampleID AND I.IsolateNumber = S.IsolateNumber ON M.SampleID = I.SampleID " & _
                  "    LEFT JOIN Urine AS U ON M.SampleID = U.SampleID " & _
                  "    LEFT JOIN Faeces ON M.SampleID = Faeces.SampleID " & _
                  "    WHERE (OrganismGroup = 'Negative Results' and D.SampleDate BETWEEN '" & calFrom & "' and '" & calTo & "') AND (M.Site IN (" & SiteCriteria & ")) " & _
                  "    ORDER BY D.SampleDate DESC, I.SampleID, M.Site"
280       End If
290       Set tb = New Recordset
300       sql1 = "SELECT  distinct top 100   AntibioticCode from (" & sql & ") AS s_1 WHERE AntibioticCode is NOT null"
310       RecOpenServer 0, tb, sql1
320       i = 68
330       With GeneralG
340           Do While Not tb.EOF
350               .ColWidth(i) = 1000: .ColAlignment(i) = flexAlignLeftCenter: .TextMatrix(0, i) = tb!AntibioticCode
360               .ColWidth(i + 1) = 1000: .ColAlignment(i + 1) = flexAlignLeftCenter: .TextMatrix(0, i + 1) = "Result"
370               tb.MoveNext
380               i = i + 2
390           Loop

400           MaxColOfAntibiotics = i
410       End With
420       tb.Close
430       RecOpenServer 0, tb, sql
440       g.Rows = 1
450       g3.Rows = 1
460       fraCSProgressBar.Visible = True
470       CSProgressBar.Value = 0
480       Do While Not tb.EOF
490           CommentExists = ""
500           ScientistComment = ""
510           ConsultantComment = ""
520           If tb!WCC & "" <> "" Then
530               WCC = GetSplitValue(tb!WCC & "", 0)
540           End If
550           Set Obs = Obs.Load(tb!SampleID, _
                                 "MicroCS", "MicroConsultant")
560           If Not Obs Is Nothing Then
570               If Obs.Count > 0 Then
580                   For Each Ob In Obs
590                       Select Case Ob.Discipline
                          Case "MicroCS"
600                           ScientistComment = Ob.Comment
610                       Case "MicroConsultant"
620                           ConsultantComment = Ob.Comment
630                       End Select
640                       CommentExists = "X"
650                   Next
660               End If
                  
670               If tb!SampleID = LastSampleId And tb!OrganismGroup & "" = LastOrgGroup And SSCheckFlu = False Then
                  
680               Else

690                   GeneralG.AddItem Format(tb!SampleDate, "dd mmm yyyy") & vbTab & Format(tb!RecDate, "dd mmm yyyy") & vbTab & Format(tb!ValidatedDateTime, "dd mmm yyyy") _
                                     & vbTab & tb!SampleID - SysOptMicroOffset(0) & vbTab & tb!Chart & vbTab & tb!PatName & vbTab & tb!Dob _
                                     & vbTab & tb!Age & vbTab & tb!sex & vbTab & tb!Addr0 & vbTab & "" & vbTab & tb!PatMobile & vbTab & tb!GP & vbTab & tb!Ward & vbTab & tb!Clinician _
                                     & vbTab & ScientistComment & vbTab & ConsultantComment _
                                     & vbTab & tb!Site & vbTab & tb!SiteDetails & vbTab & tb!OrganismName & vbTab _
                                     & tb!Qualifier & vbTab & WCC & vbTab & GetSplitValue(tb!RCC & "", 0) & vbTab & tb!WCC0 & vbTab & GetSplitValue(tb!RCC & "", 0) _
                                     & vbTab & tb!Appearance0 & vbTab & tb!Gram & vbTab & tb!Glucose & vbTab & tb!Protein & vbTab & tb!Crystals _
                                     & vbTab & GetSplitValue(tb!Rota & "", 0) & vbTab & GetSplitValue(tb!Adeno & "", 0) & vbTab & GetSplitValue(tb!Cryptosporidium & "", 0) _
                                     & vbTab & GetSplitValue(tb!GiardiaLambila & "", 0) & vbTab & GetSplitValue(tb!GDH & "", 0) _
                                     & vbTab & GetSplitValue(tb!ToxinAB & "", 0) & vbTab & GetSplitValue(tb!PCR & "", 0) & " " _
                                     & vbTab & tb!CellCount & "" & vbTab & tb!Appearance & "" & tb!GramStain1 & "" & vbTab & tb!GramStain2 & " " & vbTab & tb!ZNstain & "" & vbTab & tb!LeishmansStain & "" _
                                     & vbTab & tb!WetPrep & "" & vbTab & tb!FCrystals & "" & vbTab & tb!RCC1 & "" & vbTab & tb!RCC2 & "" & tb!RCC3 & "" & vbTab & tb!WCC1 & "" _
                                     & vbTab & tb!WCC2 & "" & vbTab & tb!WCC3 & "" & tb!Polymorphic1 & "" & vbTab & tb!Polymorphic2 & "" & tb!Polymorphic3 & "" & vbTab & tb!Mononucleated1 & "" _
                                     & vbTab & tb!Mononucleated2 & "" & vbTab & tb!Mononucleated3 & "" & vbTab & tb!FluidGlucose & "" & vbTab & tb!FluidProtein & "" & tb!FluidAlbumin & "" & vbTab & tb!FluidGlobulin & "" _
                                     & vbTab & tb!FluidLDH & "" & vbTab & tb!FluidAmylase & "" & tb!CSFGlucose & "" & vbTab & tb!CSFProtein & "" & tb!FungalElements & "" & vbTab & tb!PneumococcalAT & "" _
                                     & vbTab & tb!LegionellaAT & "" & vbTab & tb!BATScreen & "" & tb!BATScreenComment & ""

700                   If optGPClinician(0) Then
710                       If tb!GP <> "" Then
720                           GPClinician.AddItem Format(tb!SampleDate, "dd mmm yyyy") & vbTab & Format(tb!RecDate, "dd mmm yyyy") & vbTab & Format(tb!ValidatedDateTime, "dd mmm yyyy") _
                                                & vbTab & tb!SampleID - SysOptMicroOffset(0) & vbTab & tb!Chart & vbTab & tb!PatName & vbTab & tb!Dob _
                                                & vbTab & tb!Age & vbTab & tb!sex & vbTab & tb!Addr0 & vbTab & "" & vbTab & tb!PatMobile & vbTab & tb!GP & vbTab & tb!Ward & vbTab & tb!Clinician _
                                                & vbTab & ScientistComment & vbTab & ConsultantComment _
                                                & vbTab & tb!Site & vbTab & tb!SiteDetails & vbTab & tb!OrganismName & vbTab _
                                                & tb!Qualifier & vbTab & WCC & vbTab & GetSplitValue(tb!RCC & "", 0) & vbTab & tb!WCC0 & vbTab & GetSplitValue(tb!RCC & "", 0) _
                                                & vbTab & tb!Appearance0 & vbTab & tb!Gram & vbTab & tb!Glucose & vbTab & tb!Protein & vbTab & tb!Crystals _
                                                & vbTab & GetSplitValue(tb!Rota & "", 0) & vbTab & GetSplitValue(tb!Adeno & "", 0) & vbTab & GetSplitValue(tb!Cryptosporidium & "", 0) _
                                                & vbTab & GetSplitValue(tb!GiardiaLambila & "", 0) & vbTab & GetSplitValue(tb!GDH & "", 0) _
                                                & vbTab & GetSplitValue(tb!ToxinAB & "", 0) & vbTab & GetSplitValue(tb!PCR & "", 0) & " " _
                                                & vbTab & tb!CellCount & "" & vbTab & tb!Appearance & "" & tb!GramStain1 & "" & vbTab & tb!GramStain2 & " " & vbTab & tb!ZNstain & "" & vbTab & tb!LeishmansStain & "" _
                                                & vbTab & tb!WetPrep & "" & vbTab & tb!FCrystals & "" & vbTab & tb!RCC1 & "" & vbTab & tb!RCC2 & "" & tb!RCC3 & "" & vbTab & tb!WCC1 & "" _
                                                & vbTab & tb!WCC2 & "" & vbTab & tb!WCC3 & "" & tb!Polymorphic1 & "" & vbTab & tb!Polymorphic2 & "" & tb!Polymorphic3 & "" & vbTab & tb!Mononucleated1 & "" _
                                                & vbTab & tb!Mononucleated2 & "" & vbTab & tb!Mononucleated3 & "" & vbTab & tb!FluidGlucose & "" & vbTab & tb!FluidProtein & "" & tb!FluidAlbumin & "" & vbTab & tb!FluidGlobulin & "" _
                                                & vbTab & tb!FluidLDH & "" & vbTab & tb!FluidAmylase & "" & tb!CSFGlucose & "" & vbTab & tb!CSFProtein & "" & tb!FungalElements & "" & vbTab & tb!PneumococcalAT & "" _
                                                & vbTab & tb!LegionellaAT & "" & vbTab & tb!BATScreen & "" & tb!BATScreenComment & ""

730                       End If
740                   End If
750                   If optGPClinician(1) Then
760                       If tb!Clinician <> "" Then
770                           GPClinician.AddItem Format(tb!SampleDate, "dd mmm yyyy") & vbTab & Format(tb!RecDate, "dd mmm yyyy") & vbTab & Format(tb!ValidatedDateTime, "dd mmm yyyy") _
                                                & vbTab & tb!SampleID - SysOptMicroOffset(0) & vbTab & tb!Chart & vbTab & tb!PatName & vbTab & tb!Dob _
                                                & vbTab & tb!Age & vbTab & tb!sex & vbTab & tb!Addr0 & vbTab & "" & vbTab & tb!PatMobile & vbTab & tb!GP & vbTab & tb!Ward & vbTab & tb!Clinician _
                                                & vbTab & ScientistComment & vbTab & ConsultantComment _
                                                & vbTab & tb!Site & vbTab & tb!SiteDetails & vbTab & tb!OrganismName & vbTab _
                                                & tb!Qualifier & vbTab & WCC & vbTab & GetSplitValue(tb!RCC & "", 0) & vbTab & tb!WCC0 & vbTab & GetSplitValue(tb!RCC & "", 0) _
                                                & vbTab & tb!Appearance0 & vbTab & tb!Gram & vbTab & tb!Glucose & vbTab & tb!Protein & vbTab & tb!Crystals _
                                                & vbTab & GetSplitValue(tb!Rota & "", 0) & vbTab & GetSplitValue(tb!Adeno & "", 0) & vbTab & GetSplitValue(tb!Cryptosporidium & "", 0) _
                                                & vbTab & GetSplitValue(tb!GiardiaLambila & "", 0) & vbTab & GetSplitValue(tb!GDH & "", 0) _
                                                & vbTab & GetSplitValue(tb!ToxinAB & "", 0) & vbTab & GetSplitValue(tb!PCR & "", 0) & " " _
                                                & vbTab & tb!CellCount & "" & vbTab & tb!Appearance & "" & tb!GramStain1 & "" & vbTab & tb!GramStain2 & " " & vbTab & tb!ZNstain & "" & vbTab & tb!LeishmansStain & "" _
                                                & vbTab & tb!WetPrep & "" & vbTab & tb!FCrystals & "" & vbTab & tb!RCC1 & "" & vbTab & tb!RCC2 & "" & tb!RCC3 & "" & vbTab & tb!WCC1 & "" _
                                                & vbTab & tb!WCC2 & "" & vbTab & tb!WCC3 & "" & tb!Polymorphic1 & "" & vbTab & tb!Polymorphic2 & "" & tb!Polymorphic3 & "" & vbTab & tb!Mononucleated1 & "" _
                                                & vbTab & tb!Mononucleated2 & "" & vbTab & tb!Mononucleated3 & "" & vbTab & tb!FluidGlucose & "" & vbTab & tb!FluidProtein & "" & tb!FluidAlbumin & "" & vbTab & tb!FluidGlobulin & "" _
                                                & vbTab & tb!FluidLDH & "" & vbTab & tb!FluidAmylase & "" & tb!CSFGlucose & "" & vbTab & tb!CSFProtein & "" & tb!FungalElements & "" & vbTab & tb!PneumococcalAT & "" _
                                                & vbTab & tb!LegionellaAT & "" & vbTab & tb!BATScreen & "" & tb!BATScreenComment & ""

780                       End If
790                   End If
800                   If optGPClinician(2) Then
810                       If tb!Ward = cmbWard.Text Then
820                           GPClinician.AddItem Format(tb!SampleDate, "dd mmm yyyy") & vbTab & Format(tb!RecDate, "dd mmm yyyy") & vbTab & Format(tb!ValidatedDateTime, "dd mmm yyyy") _
                                                & vbTab & tb!SampleID - SysOptMicroOffset(0) & vbTab & tb!Chart & vbTab & tb!PatName & vbTab & tb!Dob _
                                                & vbTab & tb!Age & vbTab & tb!sex & vbTab & tb!Addr0 & vbTab & "" & vbTab & tb!PatMobile & vbTab & tb!GP & vbTab & tb!Ward & vbTab & tb!Clinician _
                                                & vbTab & ScientistComment & vbTab & ConsultantComment _
                                                & vbTab & tb!Site & vbTab & tb!SiteDetails & vbTab & tb!OrganismName & vbTab _
                                                & tb!Qualifier & vbTab & WCC & vbTab & GetSplitValue(tb!RCC & "", 0) & vbTab & tb!WCC0 & vbTab & GetSplitValue(tb!RCC & "", 0) _
                                                & vbTab & tb!Appearance0 & vbTab & tb!Gram & vbTab & tb!Glucose & vbTab & tb!Protein & vbTab & tb!Crystals _
                                                & vbTab & GetSplitValue(tb!Rota & "", 0) & vbTab & GetSplitValue(tb!Adeno & "", 0) & vbTab & GetSplitValue(tb!Cryptosporidium & "", 0) _
                                                & vbTab & GetSplitValue(tb!GiardiaLambila & "", 0) & vbTab & GetSplitValue(tb!GDH & "", 0) _
                                                & vbTab & GetSplitValue(tb!ToxinAB & "", 0) & vbTab & GetSplitValue(tb!PCR & "", 0) & " " _
                                                & vbTab & tb!CellCount & "" & vbTab & tb!Appearance & "" & tb!GramStain1 & "" & vbTab & tb!GramStain2 & " " & vbTab & tb!ZNstain & "" & vbTab & tb!LeishmansStain & "" _
                                                & vbTab & tb!WetPrep & "" & vbTab & tb!FCrystals & "" & vbTab & tb!RCC1 & "" & vbTab & tb!RCC2 & "" & tb!RCC3 & "" & vbTab & tb!WCC1 & "" _
                                                & vbTab & tb!WCC2 & "" & vbTab & tb!WCC3 & "" & tb!Polymorphic1 & "" & vbTab & tb!Polymorphic2 & "" & tb!Polymorphic3 & "" & vbTab & tb!Mononucleated1 & "" _
                                                & vbTab & tb!Mononucleated2 & "" & vbTab & tb!Mononucleated3 & "" & vbTab & tb!FluidGlucose & "" & vbTab & tb!FluidProtein & "" & tb!FluidAlbumin & "" & vbTab & tb!FluidGlobulin & "" _
                                                & vbTab & tb!FluidLDH & "" & vbTab & tb!FluidAmylase & "" & tb!CSFGlucose & "" & vbTab & tb!CSFProtein & "" & tb!FungalElements & "" & vbTab & tb!PneumococcalAT & "" _
                                                & vbTab & tb!LegionellaAT & "" & vbTab & tb!BATScreen & "" & tb!BATScreenComment & ""
830                       End If
840                   End If
850               End If
860               For i = 68 To MaxColOfAntibiotics
870                   If GeneralG.TextMatrix(0, i) & "" = (tb!AntibioticCode & "") Then
880                       GeneralG.TextMatrix(GeneralG.Rows - 1, i) = IIf(IsNull(tb!RSI), "", IIf(Trim(tb!RSI) = "", "", tb!RSI))    ' & ""
890                       GeneralG.TextMatrix(GeneralG.Rows - 1, i + 1) = IIf(IsNull(tb!Result), "", IIf(Trim(tb!Result) = "", "", tb!Result))  ' & ""
900                   End If
910               Next
920           End If
930           LastSampleId = tb!SampleID & ""
940           LastOrgGroup = tb!OrganismGroup & ""
950           CSProgressBar.Value = CSProgressBar.Value + 1
960           lblCSProgressBar = "Fetching results ... (" & Int(CSProgressBar.Value * 100 / CSProgressBar.Max) & " %)"
970           If CSProgressBar.Value = 100 Then
980               CSProgressBar.Value = 0
990           End If
1000          lblCSProgressBar.Refresh
1010          tb.MoveNext
1020      Loop
1030      fraCSProgressBar.Visible = False
1040      g.Visible = True

1050      Exit Sub

FillGeneral_Error:

          Dim strES As String
          Dim intEL As Integer

1060      intEL = Erl
1070      strES = Err.Description
1080      LogError "frmMicroSurveillanceSearches", "FillGeneralG", intEL, strES, sql

End Sub
Private Sub FillG4()

      Dim sql As String
      Dim tb As Recordset
      Dim test As String
      Dim Obs As New Observations
      Dim Ob As Observation
      Dim SiteCriteria As String
      Dim Res() As String
      Dim WCC As String
      Dim ScientistComment As String
      Dim ConsultantComment As String
      Dim CommentExists As String

10    On Error GoTo FillG4_Error
      '<***********************************X = 33
20    g4.Visible = False

30    If optType(0) Then
40        sql = "SELECT * FROM" & _
                " (SELECT     Demographics.SampleID, Demographics.PatName, Demographics.Age, Demographics.Sex, Demographics.SampleDate, COUNT(Faeces.Rota) AS Rota, COUNT(Faeces.Adeno) AS Adeno, " & _
                "            COUNT(Faeces.OB0) AS OBO, COUNT(Faeces.OB1) AS OB1, COUNT(Faeces.OB2) AS OB2, COUNT(Faeces.ToxinAB) AS ToxinAB, COUNT(Faeces.Cryptosporidium) " & _
                "            AS Cryptosporidium, COUNT(Faeces.HPylori) AS HPylori, COUNT(Faeces.CDiffCulture) AS CDiffCulture, GenericResults.TestName, COUNT(GenericResults.Result) " & _
                "            AS Result, COUNT(Urine.WCC) AS WCC, COUNT(Urine.RCC) AS RCC, COUNT(Urine.Crystals) AS Crystals, COUNT(Urine.Casts) AS Casts " & _
                " FROM         Demographics LEFT OUTER JOIN " & _
                "            Urine ON Demographics.SampleID = Urine.SampleID LEFT OUTER JOIN " & _
                "            Faeces ON Demographics.SampleID = Faeces.SampleID LEFT OUTER JOIN " & _
                "            GenericResults ON Demographics.SampleID = GenericResults.SampleID " & _
                " GROUP BY Demographics.PatName, Demographics.Age, Demographics.Sex, Demographics.SampleID, GenericResults.TestName, Demographics.SampleDate " & _
                " HAVING      (Demographics.SampleDate BETWEEN '" & calFrom & "' and '" & calTo & "')) g " & _
                " Pivot " & _
                " ( SUM(Result) FOR Testname IN ([RedSub],[RSV])) AS pvt"

50    End If
60    Set tb = New Recordset
70    RecOpenServer 0, tb, sql
80    g4.Clear
90    InitGrid4
100   g4.Rows = 1
110   fraMainProgressBar.Visible = True
120   MainProgressBar.Value = 0
130   Do While Not tb.EOF

140           g4.AddItem tb!SampleID - SysOptMicroOffset(0) & vbTab & tb!PatName & vbTab & tb!sex _
                         & vbTab & tb!Age & vbTab & Format(tb!SampleDate, "dd mmm yyyy") & vbTab & tb!Rota & vbTab & tb!Adeno _
                         & vbTab & tb!OBO & vbTab & tb!OB1 & vbTab & tb!OB2 & vbTab & tb!ToxinAB & vbTab & tb!Cryptosporidium & vbTab & tb!HPylori & vbTab _
                         & tb!CDiffCulture & vbTab & tb!WCC & vbTab & tb!RCC & vbTab & tb!Crystals & vbTab & tb!Casts & vbTab & IIf(IsNull(tb!RedSub), 0, tb!RedSub) & vbTab & IIf(IsNull(tb!RSV), 0, tb!RSV)
150   tb.MoveNext
160   Loop
170   fraMainProgressBar.Visible = False
180   g4.Visible = True

190   Exit Sub

FillG4_Error:

      Dim strES As String
      Dim intEL As Integer

200   intEL = Erl
210   strES = Err.Description
220   LogError "frmMicroSurveillanceSearches", "FillG4", intEL, strES, sql

End Sub

Private Sub FillG7()

      Dim sql As String
      Dim tb As Recordset
      Dim test As String
      Dim Obs As New Observations
      Dim Ob As Observation
      Dim SiteCriteria As String
      Dim Res() As String
      Dim WCC As String
      Dim ScientistComment As String
      Dim ConsultantComment As String
      Dim CommentExists As String
      Dim i As Integer


10    On Error GoTo FillG7_Error
      '<***********************************X = 33
20    G7.Visible = False

30    For i = 0 To lstSites.ListCount - 1
40        If lstSites.Selected(i) Then
50            SiteCriteria = SiteCriteria & "'" & lstSites.List(i) & "" & "',"
60        End If
70    Next

80    If Trim(SiteCriteria) = "" Then Exit Sub
90    SiteCriteria = Left(SiteCriteria, Len(SiteCriteria) - 1)
100   If optType(0) Then

         
110       sql = " SELECT MicroSiteDetails.Site, " & _
                "              CASE WHEN FLOOR((CAST(SampleDate AS INTEGER) - CAST(Dob AS INTEGER)) / 365.25) BETWEEN 0 AND 4 THEN '0-4' WHEN FLOOR((CAST(SampleDate AS INTEGER) - CAST(Dob AS INTEGER)) / 365.25) BETWEEN 5 AND 14 THEN '5-14' WHEN FLOOR((CAST(SampleDate AS INTEGER) - CAST(Dob AS INTEGER)) / 365.25) BETWEEN 15 AND 44 THEN '15-44' WHEN FLOOR((CAST(SampleDate AS INTEGER) - CAST(Dob AS INTEGER)) / 365.25) BETWEEN 45 AND 60 THEN '45-60' WHEN FLOOR((CAST(SampleDate AS INTEGER) - CAST(Dob AS INTEGER)) / 365.25) > 60 THEN '60+' END AS RANGE, " & _
                "              Demographics.Sex, COUNT(Faeces.Rota) AS Rota,  COUNT(Faeces.Adeno) AS Adeno, COUNT(Faeces.OB0) AS OBO,  " & _
                "              COUNT(Faeces.OB1) AS OB1,  COUNT(Faeces.OB2) AS OB2,  COUNT(Faeces.ToxinAB) AS ToxinAB, COUNT(Faeces.Cryptosporidium) AS Cryptosporidium, " & _
                "              COUNT(Faeces.HPylori) AS HPylori,  COUNT(Faeces.CDiffCulture) AS CDiffCulture, " & _
                "              COUNT( CASE  WHEN GenericResults.TestName = N'RSV' THEN '1' END) AS RSV,  COUNT(   CASE  WHEN GenericResults.TestName = N'RedSub' THEN '1'  END) AS RedSub, " & _
                "              COUNT(Urine.WCC) AS WCC, COUNT(Urine.RCC) AS RCC, COUNT(Urine.Crystals) AS Crystals, COUNT(Urine.Casts)  As Casts " & _
                " FROM         Demographics LEFT OUTER JOIN  MicroSiteDetails ON Demographics.SampleID = MicroSiteDetails.SampleID " & _
                "                           LEFT OUTER JOIN Urine ON Demographics.SampleID = Urine.SampleID " & _
                "                           LEFT OUTER JOIN Faeces ON Demographics.SampleID = Faeces.SampleID " & _
                "                           LEFT OUTER JOIN GenericResults ON Demographics.SampleID = GenericResults.SampleID " & _
                " WHERE        (Demographics.SampleDate BETWEEN '" & calFrom & "' and '" & calTo & "') AND SITE IN (" & SiteCriteria & ")  " & _
                " GROUP BY      MicroSiteDetails.Site, Demographics.Sex, " & _
                "              CASE WHEN FLOOR((CAST(SampleDate AS INTEGER) - CAST(Dob AS INTEGER)) / 365.25) BETWEEN 0 AND 4 THEN '0-4' " & _
                "                   WHEN FLOOR((CAST(SampleDate AS INTEGER) - CAST(Dob AS INTEGER)) / 365.25) BETWEEN 5 AND 14 THEN '5-14' " & _
                "                   WHEN FLOOR((CAST(SampleDate AS INTEGER) - CAST(Dob AS INTEGER)) / 365.25) BETWEEN 15 AND 44 THEN '15-44' " & _
                "                   WHEN FLOOR((CAST(SampleDate AS INTEGER) - CAST(Dob AS INTEGER)) / 365.25) BETWEEN 45 AND 60 THEN '45-60' " & _
                "                   WHEN FLOOR((CAST(SampleDate AS INTEGER) - CAST(Dob AS INTEGER)) / 365.25) > 60 THEN '60+' End " & _
                " Having (Not (MicroSiteDetails.Site Is Null))"
                

120   Set tb = New Recordset
130   RecOpenServer 0, tb, sql
140   G7.Clear
150   InitGrid7
160   G7.Rows = 1
170   fraCSProgressBar.Visible = True
180   CSProgressBar.Value = 0
      Dim Site As String
      Dim range As String
190   Do While Not tb.EOF
200       If Len(tb!Site) = 0 Or IsNull(tb!Site) Then
210       Else
220           If Site = tb!Site & "" Then
230               Site = ""
240           Else
250               Site = tb!Site & ""
260           End If
              '190   If range = tb!range & "" Then
              '200       If site = "" Then
              '210           range = ""
              '220       Else
              '230           range = tb!range & ""
              '240       End If
              '250   Else
              '260       range = tb!range & ""
              '270   End If
270           G7.AddItem Site & vbTab & tb!range & vbTab & IIf(Len(tb!sex) = 0, "U", tb!sex) _
                         & vbTab & tb!Rota & vbTab & tb!Adeno & vbTab & tb!OBO & vbTab & tb!OB1 & vbTab & tb!OB2 & vbTab & tb!ToxinAB & vbTab & tb!Cryptosporidium & vbTab & tb!HPylori & vbTab _
                         & tb!CDiffCulture & vbTab & tb!WCC & vbTab & tb!RCC & vbTab & tb!Crystals & vbTab & tb!Casts & vbTab & IIf(IsNull(tb!RedSub), 0, tb!RedSub) & vbTab & IIf(IsNull(tb!RSV), 0, tb!RSV)
280           Site = tb!Site & ""
              '300           range = tb!range & ""
290           tb.MoveNext
300       End If
310       CSProgressBar.Value = CSProgressBar.Value + 1
320       lblCSProgressBar = "Fetching results ... (" & Int(CSProgressBar.Value * 100 / CSProgressBar.Max) & " %)"
330       If CSProgressBar.Value = 100 Then
340           CSProgressBar.Value = 0
350       End If
360       lblCSProgressBar.Refresh

370   Loop
380   G7.AddItem "Totals"
390   fraCSProgressBar.Visible = False
400   SortG7
410   End If

420   G7.Visible = True

430   Exit Sub

FillG7_Error:

      Dim strES As String
      Dim intEL As Integer

440   intEL = Erl
450   strES = Err.Description
460   LogError "frmMicroSurveillanceSearches", "FillG7", intEL, strES, sql

End Sub

Private Function OrganismExistsInList(ByVal OrganismName As String) As Boolean
      Dim i As Integer

10    On Error GoTo OrganismExistsInList_Error

20    OrganismExistsInList = False
30    For i = 0 To lstOrganismGroup.ListCount - 1
40        If lstOrganismGroup.List(i) = OrganismName & "" Then
50            OrganismExistsInList = True
60        End If
70    Next i

80    Exit Function

OrganismExistsInList_Error:

      Dim strES As String
      Dim intEL As Integer

90    intEL = Erl
100   strES = Err.Description
110   LogError "frmMicroSurveillanceSearches", "OrganismExistsInList", intEL, strES

End Function

Private Function SiteExistsInList(ByVal siteName As String) As Boolean
      Dim i As Integer

10    On Error GoTo SiteExistsInList_Error

20    SiteExistsInList = False

30    For i = 0 To lstSites.ListCount - 1
40        If lstSites.List(i) = siteName & "" Then
50            SiteExistsInList = True
60        End If
70    Next i

80    Exit Function

SiteExistsInList_Error:

      Dim strES As String
      Dim intEL As Integer

90    intEL = Erl
100   strES = Err.Description
110   LogError "frmMicroSurveillanceSearches", "SiteExistsInList", intEL, strES

End Function

Public Sub ClearFGrid(ByVal g As MSFlexGrid)

10    On Error GoTo ClearFGrid_Error

20    With g
30        .Rows = .FixedRows + 1
40        .AddItem ""
50        .RemoveItem .FixedRows
60        .Visible = False
70    End With

80    Exit Sub

ClearFGrid_Error:

      Dim strES As String
      Dim intEL As Integer

90    intEL = Erl
100   strES = Err.Description
110   LogError "frmMicroSurveillanceSearches", "ClearFGrid", intEL, strES

End Sub

Private Sub InitGrid6()
      Dim i As Integer

10    On Error GoTo InitGrid6_Error

20    With G6
30        .Rows = 2: .Cols = 13
40        .FixedRows = 1: .FixedCols = 1
50        .Rows = 1
60        .SelectionMode = flexSelectionByRow
70        .RowHeight(0) = .RowHeight(0) * 2
80        .TextMatrix(0, 0) = "Sample Date"
90        .TextMatrix(0, 1) = "Chart Numnber"
100       .TextMatrix(0, 2) = "laboratory Number"
110       .TextMatrix(0, 3) = "Patients Name"
120       .TextMatrix(0, 4) = "DOB"
130       .TextMatrix(0, 5) = "Sex"
140       .TextMatrix(0, 6) = "Age"
150       .TextMatrix(0, 7) = "Ward"
160       .TextMatrix(0, 8) = "Doctor"
170       .TextMatrix(0, 9) = "Location"
180       .TextMatrix(0, 10) = "Site"
190       .TextMatrix(0, 11) = "Other Site"
200       .TextMatrix(0, 12) = "Organism"
210       For i = 0 To .Cols - 1
220           .ColWidth(i) = 900
230           .ColAlignment(i) = flexAlignCenterCenter
240       Next i
250       .WordWrap = True
260   End With

270   Exit Sub

InitGrid6_Error:

      Dim strES As String
      Dim intEL As Integer

280   intEL = Erl
290   strES = Err.Description
300   LogError "frmMicroSurveillanceSearches", "InitGrid6", intEL, strES
End Sub
'---------------------------------------------------------------------------------------
' Procedure : InitGrid3
' Author    : Trevor Dunican
' Date      : 19/12/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub InitGrid3()
      Dim i As Integer
      Dim z As Integer

10    On Error GoTo InitGrid3_Error

20    With g3
30        .Rows = 2: .Cols = 198
40        .FixedRows = 1: .FixedCols = 1
50        .Rows = 1
60        .SelectionMode = flexSelectionByRow
70        .RowHeight(0) = .RowHeight(0) * 2
80        .TextMatrix(0, 0) = "Sample Date"
90        .TextMatrix(0, 1) = "Date Received"
100       .TextMatrix(0, 2) = "Date Validated"
110       .TextMatrix(0, 3) = "laboratory Number"
120       .TextMatrix(0, 4) = "Chart Number"
130       .TextMatrix(0, 5) = "Name"
140       .TextMatrix(0, 6) = "DOB"
150       .TextMatrix(0, 7) = "Age"
160       .TextMatrix(0, 8) = "Sex"
170       .TextMatrix(0, 9) = ""
180       .TextMatrix(0, 10) = ""
190       .TextMatrix(0, 11) = ""
200       .TextMatrix(0, 12) = "Address"
210       .TextMatrix(0, 13) = "Doctor"
220       .TextMatrix(0, 14) = "Location"
230       .TextMatrix(0, 15) = "Site"
240       .TextMatrix(0, 16) = "Site Details"
250       .TextMatrix(0, 17) = "Organism"
260   .TextMatrix(0, 18) = "Result"
270       .TextMatrix(0, 19) = "WCC"
280       .TextMatrix(0, 20) = "Clinician"
290       .TextMatrix(0, 21) = "Cell Count"
300       .TextMatrix(0, 22) = "All Sus"
310       .TextMatrix(0, 23) = "C"
320       .TextMatrix(0, 24) = "Scientist Comments"
330       .TextMatrix(0, 25) = "Consultant Comments"
340       .TextMatrix(0, 26) = ""
350       For i = 0 To .Cols - 1
360           .ColWidth(i) = 600
370           .ColAlignment(i) = flexAlignCenterCenter
380       Next i
390       For z = 27 To .Cols - 1 Step 2
400           .TextMatrix(0, z) = "S/I/R"
410           .ColWidth(z) = 600
420       Next z
430       .WordWrap = True
440   End With

450   Exit Sub

InitGrid3_Error:

      Dim strES As String
      Dim intEL As Integer

460   intEL = Erl
470   strES = Err.Description
480   LogError "frmMicroSurveillanceSearches", "InitGrid3", intEL, strES
End Sub
Private Sub InitGeneralG()
      Dim i As Integer
      Dim z As Integer

10    On Error GoTo InitGeneralG_Error

20    With GeneralG
30        .Clear
40        .Rows = 2: .Cols = 200    '68 columns are fix and the 138 are for use as dynamic

50        .FixedRows = 1    '.FixedCols = 1
60        .Rows = 1
70        .SelectionMode = flexSelectionByRow
80        .RowHeight(0) = TextHeight("A") * 2
90        .ColWidth(0) = 1200: .ColAlignment(0) = flexAlignLeftCenter: .TextMatrix(0, 0) = "Sample Date"
100       .ColWidth(1) = 1200: .ColAlignment(1) = flexAlignLeftCenter: .TextMatrix(0, 1) = "Date Received"
110       .ColWidth(2) = 1200: .ColAlignment(2) = flexAlignLeftCenter: .TextMatrix(0, 2) = "Date Validated"
120       .ColWidth(3) = 1200: .ColAlignment(3) = flexAlignLeftCenter: .TextMatrix(0, 3) = "laboratory Number"
130       .ColWidth(4) = 1000: .ColAlignment(4) = flexAlignLeftCenter: .TextMatrix(0, 4) = "Chart Number"
140       .ColWidth(5) = 2000: .ColAlignment(5) = flexAlignLeftCenter: .TextMatrix(0, 5) = "Name"
150       .ColWidth(6) = 1000: .ColAlignment(6) = flexAlignLeftCenter: .TextMatrix(0, 6) = "DOB"
160       .ColWidth(7) = 800: .ColAlignment(7) = flexAlignLeftCenter: .TextMatrix(0, 7) = "Age"
170       .ColWidth(8) = 500: .ColAlignment(8) = flexAlignLeftCenter: .TextMatrix(0, 8) = "Sex"
180       .ColWidth(9) = 3000: .ColAlignment(9) = flexAlignLeftCenter: .TextMatrix(0, 9) = "Address"
190       .ColWidth(9) = 1000: .ColAlignment(10) = flexAlignLeftCenter: .TextMatrix(0, 10) = "EirCode"
200       .ColWidth(9) = 1000: .ColAlignment(11) = flexAlignLeftCenter: .TextMatrix(0, 11) = "Phone"
          
210       .ColWidth(10) = 2000: .ColAlignment(12) = flexAlignLeftCenter: .TextMatrix(0, 12) = "GP"    'Doctor
220       .ColWidth(11) = 2000: .ColAlignment(13) = flexAlignLeftCenter: .TextMatrix(0, 13) = "Location"
230       .ColWidth(12) = 1000: .ColAlignment(14) = flexAlignLeftCenter: .TextMatrix(0, 14) = "Consultant"
240       .ColWidth(13) = 2000: .ColAlignment(15) = flexAlignLeftCenter: .TextMatrix(0, 15) = "Scientist Comments"
250       .ColWidth(14) = 2000: .ColAlignment(16) = flexAlignLeftCenter: .TextMatrix(0, 16) = "Consultant Comments"
260       .ColWidth(15) = 1000: .ColAlignment(17) = flexAlignLeftCenter: .TextMatrix(0, 17) = "Site"
270       .ColWidth(16) = 1000: .ColAlignment(18) = flexAlignLeftCenter: .TextMatrix(0, 18) = "Site Details"
280       .ColWidth(17) = 2000: .ColAlignment(19) = flexAlignLeftCenter: .TextMatrix(0, 19) = "Organism Name"
290       .ColWidth(18) = 1000: .ColAlignment(20) = flexAlignLeftCenter: .TextMatrix(0, 20) = "Result"
          '-----------
300       .ColWidth(19) = 1000: .ColAlignment(21) = flexAlignLeftCenter: .TextMatrix(0, 21) = "WCC"
310       .ColWidth(20) = 1000: .ColAlignment(22) = flexAlignLeftCenter: .TextMatrix(0, 22) = "Cell Count"
320       .ColWidth(21) = 1000: .ColAlignment(23) = flexAlignLeftCenter: .TextMatrix(0, 23) = "WCC/CMM"
330       .ColWidth(22) = 1000: .ColAlignment(24) = flexAlignLeftCenter: .TextMatrix(0, 24) = "RCC/CMM"
340       .ColWidth(23) = 1000: .ColAlignment(25) = flexAlignLeftCenter: .TextMatrix(0, 25) = "Appearance"
350       .ColWidth(24) = 1000: .ColAlignment(26) = flexAlignLeftCenter: .TextMatrix(0, 26) = "Gram Stain"
360       .ColWidth(25) = 1000: .ColAlignment(27) = flexAlignLeftCenter: .TextMatrix(0, 27) = "Glocose mmol/L"
370       .ColWidth(26) = 1000: .ColAlignment(28) = flexAlignLeftCenter: .TextMatrix(0, 27) = "Proteing/L"
380       .ColWidth(27) = 1000: .ColAlignment(29) = flexAlignLeftCenter: .TextMatrix(0, 28) = "Crystal"
390       .ColWidth(28) = 1250: .ColAlignment(30) = flexAlignLeftCenter: .TextMatrix(0, 30) = "Rota"
400       .ColWidth(29) = 1250: .ColAlignment(31) = flexAlignLeftCenter: .TextMatrix(0, 31) = "Adeno"
410       .ColWidth(30) = 1500: .ColAlignment(32) = flexAlignLeftCenter: .TextMatrix(0, 32) = "Crypo"
420       .ColWidth(31) = 1000: .ColAlignment(33) = flexAlignLeftCenter: .TextMatrix(0, 33) = "Giardia"
430       .ColWidth(32) = 1000: .ColAlignment(34) = flexAlignLeftCenter: .TextMatrix(0, 34) = "GDH"
440       .ColWidth(33) = 1270: .ColAlignment(35) = flexAlignLeftCenter: .TextMatrix(0, 35) = "Toxin"
450       .ColWidth(34) = 3680: .ColAlignment(36) = flexAlignLeftCenter: .TextMatrix(0, 36) = "PCR"
460       .ColWidth(35) = 1000: .ColAlignment(37) = flexAlignLeftCenter: .TextMatrix(0, 37) = "F.Cell Count"
470       .ColWidth(36) = 1000: .ColAlignment(38) = flexAlignLeftCenter: .TextMatrix(0, 38) = "F.Appearance"
480       .ColWidth(37) = 1000: .ColAlignment(39) = flexAlignLeftCenter: .TextMatrix(0, 39) = "F.GramStain1"
490       .ColWidth(38) = 1000: .ColAlignment(40) = flexAlignLeftCenter: .TextMatrix(0, 40) = "F.GramStain2"
500       .ColWidth(39) = 1000: .ColAlignment(41) = flexAlignLeftCenter: .TextMatrix(0, 41) = "F.ZN"
510       .ColWidth(40) = 1000: .ColAlignment(42) = flexAlignLeftCenter: .TextMatrix(0, 42) = "F.Leishmans"
520       .ColWidth(41) = 1000: .ColAlignment(43) = flexAlignLeftCenter: .TextMatrix(0, 43) = "F.WetPrep"
530       .ColWidth(42) = 1000: .ColAlignment(44) = flexAlignLeftCenter: .TextMatrix(0, 44) = "F.Crystal"
540       .ColWidth(43) = 1000: .ColAlignment(45) = flexAlignLeftCenter: .TextMatrix(0, 45) = "RCC1"
550       .ColWidth(44) = 1000: .ColAlignment(46) = flexAlignLeftCenter: .TextMatrix(0, 46) = "RCC2"
560       .ColWidth(45) = 1000: .ColAlignment(47) = flexAlignLeftCenter: .TextMatrix(0, 47) = "RCC3"
570       .ColWidth(46) = 1000: .ColAlignment(48) = flexAlignLeftCenter: .TextMatrix(0, 48) = "WCC1"
580       .ColWidth(47) = 1000: .ColAlignment(49) = flexAlignLeftCenter: .TextMatrix(0, 49) = "WCC2"
590       .ColWidth(48) = 1000: .ColAlignment(50) = flexAlignLeftCenter: .TextMatrix(0, 50) = "WCC3"
600       .ColWidth(49) = 1000: .ColAlignment(51) = flexAlignLeftCenter: .TextMatrix(0, 51) = "Polymorphic1"
610       .ColWidth(50) = 1000: .ColAlignment(52) = flexAlignLeftCenter: .TextMatrix(0, 52) = "Polymorphic2"
620       .ColWidth(51) = 1000: .ColAlignment(53) = flexAlignLeftCenter: .TextMatrix(0, 53) = "Polymorphic3"
630       .ColWidth(52) = 1000: .ColAlignment(54) = flexAlignLeftCenter: .TextMatrix(0, 54) = "Mononucleated1"
640       .ColWidth(53) = 1000: .ColAlignment(55) = flexAlignLeftCenter: .TextMatrix(0, 55) = "Mononucleated2"
650       .ColWidth(54) = 1000: .ColAlignment(56) = flexAlignLeftCenter: .TextMatrix(0, 56) = "Mononucleated3"
660       .ColWidth(55) = 1000: .ColAlignment(57) = flexAlignLeftCenter: .TextMatrix(0, 57) = "FluidGlucose"
670       .ColWidth(56) = 1000: .ColAlignment(58) = flexAlignLeftCenter: .TextMatrix(0, 58) = "FluidProtein"
680       .ColWidth(57) = 1000: .ColAlignment(59) = flexAlignLeftCenter: .TextMatrix(0, 59) = "FluidAlbumin"
690       .ColWidth(58) = 1000: .ColAlignment(60) = flexAlignLeftCenter: .TextMatrix(0, 60) = "FluidGlobulin"
700       .ColWidth(59) = 1000: .ColAlignment(61) = flexAlignLeftCenter: .TextMatrix(0, 61) = "FluidLDH"
710       .ColWidth(60) = 1000: .ColAlignment(62) = flexAlignLeftCenter: .TextMatrix(0, 62) = "FluidAmylase"
720       .ColWidth(61) = 1000: .ColAlignment(63) = flexAlignLeftCenter: .TextMatrix(0, 63) = "CSFGlucose"
730       .ColWidth(62) = 1000: .ColAlignment(64) = flexAlignLeftCenter: .TextMatrix(0, 64) = "CSFProtein"
740       .ColWidth(63) = 1250: .ColAlignment(65) = flexAlignLeftCenter: .TextMatrix(0, 65) = "FungalElements"
750       .ColWidth(64) = 1000: .ColAlignment(66) = flexAlignLeftCenter: .TextMatrix(0, 66) = "PneumococcalAT"
760       .ColWidth(65) = 1250: .ColAlignment(67) = flexAlignLeftCenter: .TextMatrix(0, 67) = "LegionellaAT"
770       .ColWidth(66) = 1500: .ColAlignment(68) = flexAlignLeftCenter: .TextMatrix(0, 68) = "BATScreen"
780       .ColWidth(67) = 1000: .ColAlignment(69) = flexAlignLeftCenter: .TextMatrix(0, 69) = "BATScreenComment"
          
790       For i = 70 To .Cols - 1
800           .ColWidth(i) = 0: .ColAlignment(i) = flexAlignLeftCenter: .TextMatrix(0, i) = ""
810       Next

820       .WordWrap = True
830   End With

840   Exit Sub

InitGeneralG_Error:

      Dim strES As String
      Dim intEL As Integer

850   intEL = Erl
860   strES = Err.Description
870   LogError "frmMicroSurveillanceSearches", "InitGeneralG", intEL, strES
End Sub

'---------------------------------------------------------------------------------------
' Procedure : InitGPClinician
' Author    : tduni
' Date      : 22/05/2018
' Purpose   : Micro Surveillance Searches
'---------------------------------------------------------------------------------------
'
Private Sub InitGPClinician()
Dim i As Integer
Dim z As Integer

On Error GoTo InitGPClinician_Error

With GPClinician
    .Clear
    .Rows = 2: .Cols = 200    '68 columns are fix and the 138 are for use as dynamic

    .FixedRows = 1    '.FixedCols = 1
    .Rows = 1
    .SelectionMode = flexSelectionByRow
    .RowHeight(0) = TextHeight("A") * 2
    .ColWidth(0) = 1200: .ColAlignment(0) = flexAlignLeftCenter: .TextMatrix(0, 0) = "Sample Date"
    .ColWidth(1) = 1200: .ColAlignment(1) = flexAlignLeftCenter: .TextMatrix(0, 1) = "Date Received"
    .ColWidth(2) = 1200: .ColAlignment(2) = flexAlignLeftCenter: .TextMatrix(0, 2) = "Date Validated"
    .ColWidth(3) = 1200: .ColAlignment(3) = flexAlignLeftCenter: .TextMatrix(0, 3) = "laboratory Number"
    .ColWidth(4) = 1000: .ColAlignment(4) = flexAlignLeftCenter: .TextMatrix(0, 4) = "Chart Number"
    .ColWidth(5) = 2000: .ColAlignment(5) = flexAlignLeftCenter: .TextMatrix(0, 5) = "Name"
    .ColWidth(6) = 1000: .ColAlignment(6) = flexAlignLeftCenter: .TextMatrix(0, 6) = "DOB"
    .ColWidth(7) = 800: .ColAlignment(7) = flexAlignLeftCenter: .TextMatrix(0, 7) = "Age"
    .ColWidth(8) = 500: .ColAlignment(8) = flexAlignLeftCenter: .TextMatrix(0, 8) = "Sex"
    .ColWidth(9) = 3000: .ColAlignment(9) = flexAlignLeftCenter: .TextMatrix(0, 9) = "Address"
    .ColWidth(9) = 1000: .ColAlignment(10) = flexAlignLeftCenter: .TextMatrix(0, 10) = "EirCode"
    .ColWidth(9) = 1000: .ColAlignment(11) = flexAlignLeftCenter: .TextMatrix(0, 11) = "Phone"
    
    .ColWidth(10) = 2000: .ColAlignment(12) = flexAlignLeftCenter: .TextMatrix(0, 12) = "GP"    'Doctor
    .ColWidth(11) = 2000: .ColAlignment(13) = flexAlignLeftCenter: .TextMatrix(0, 13) = "Location"
    .ColWidth(12) = 1000: .ColAlignment(14) = flexAlignLeftCenter: .TextMatrix(0, 14) = "Consultant"
    .ColWidth(13) = 2000: .ColAlignment(15) = flexAlignLeftCenter: .TextMatrix(0, 15) = "Scientist Comments"
    .ColWidth(14) = 2000: .ColAlignment(16) = flexAlignLeftCenter: .TextMatrix(0, 16) = "Consultant Comments"
    .ColWidth(15) = 1000: .ColAlignment(17) = flexAlignLeftCenter: .TextMatrix(0, 17) = "Site"
    .ColWidth(16) = 1000: .ColAlignment(18) = flexAlignLeftCenter: .TextMatrix(0, 18) = "Site Details"
    .ColWidth(17) = 2000: .ColAlignment(19) = flexAlignLeftCenter: .TextMatrix(0, 19) = "Organism Name"
    .ColWidth(18) = 1000: .ColAlignment(20) = flexAlignLeftCenter: .TextMatrix(0, 20) = "Result"
    '-----------
    .ColWidth(19) = 1000: .ColAlignment(21) = flexAlignLeftCenter: .TextMatrix(0, 21) = "WCC"
    .ColWidth(20) = 1000: .ColAlignment(22) = flexAlignLeftCenter: .TextMatrix(0, 22) = "Cell Count"
    .ColWidth(21) = 1000: .ColAlignment(23) = flexAlignLeftCenter: .TextMatrix(0, 23) = "WCC/CMM"
    .ColWidth(22) = 1000: .ColAlignment(24) = flexAlignLeftCenter: .TextMatrix(0, 24) = "RCC/CMM"
    .ColWidth(23) = 1000: .ColAlignment(25) = flexAlignLeftCenter: .TextMatrix(0, 25) = "Appearance"
    .ColWidth(24) = 1000: .ColAlignment(26) = flexAlignLeftCenter: .TextMatrix(0, 26) = "Gram Stain"
    .ColWidth(25) = 1000: .ColAlignment(27) = flexAlignLeftCenter: .TextMatrix(0, 27) = "Glocose mmol/L"
    .ColWidth(26) = 1000: .ColAlignment(28) = flexAlignLeftCenter: .TextMatrix(0, 27) = "Proteing/L"
    .ColWidth(27) = 1000: .ColAlignment(29) = flexAlignLeftCenter: .TextMatrix(0, 28) = "Crystal"
    .ColWidth(28) = 1250: .ColAlignment(30) = flexAlignLeftCenter: .TextMatrix(0, 30) = "Rota"
    .ColWidth(29) = 1250: .ColAlignment(31) = flexAlignLeftCenter: .TextMatrix(0, 31) = "Adeno"
    .ColWidth(30) = 1500: .ColAlignment(32) = flexAlignLeftCenter: .TextMatrix(0, 32) = "Crypo"
    .ColWidth(31) = 1000: .ColAlignment(33) = flexAlignLeftCenter: .TextMatrix(0, 33) = "Giardia"
    .ColWidth(32) = 1000: .ColAlignment(34) = flexAlignLeftCenter: .TextMatrix(0, 34) = "GDH"
    .ColWidth(33) = 1270: .ColAlignment(35) = flexAlignLeftCenter: .TextMatrix(0, 35) = "Toxin"
    .ColWidth(34) = 3680: .ColAlignment(36) = flexAlignLeftCenter: .TextMatrix(0, 36) = "PCR"
    .ColWidth(35) = 1000: .ColAlignment(37) = flexAlignLeftCenter: .TextMatrix(0, 37) = "F.Cell Count"
    .ColWidth(36) = 1000: .ColAlignment(38) = flexAlignLeftCenter: .TextMatrix(0, 38) = "F.Appearance"
    .ColWidth(37) = 1000: .ColAlignment(39) = flexAlignLeftCenter: .TextMatrix(0, 39) = "F.GramStain1"
    .ColWidth(38) = 1000: .ColAlignment(40) = flexAlignLeftCenter: .TextMatrix(0, 40) = "F.GramStain2"
    .ColWidth(39) = 1000: .ColAlignment(41) = flexAlignLeftCenter: .TextMatrix(0, 41) = "F.ZN"
    .ColWidth(40) = 1000: .ColAlignment(42) = flexAlignLeftCenter: .TextMatrix(0, 42) = "F.Leishmans"
    .ColWidth(41) = 1000: .ColAlignment(43) = flexAlignLeftCenter: .TextMatrix(0, 43) = "F.WetPrep"
    .ColWidth(42) = 1000: .ColAlignment(44) = flexAlignLeftCenter: .TextMatrix(0, 44) = "F.Crystal"
    .ColWidth(43) = 1000: .ColAlignment(45) = flexAlignLeftCenter: .TextMatrix(0, 45) = "RCC1"
    .ColWidth(44) = 1000: .ColAlignment(46) = flexAlignLeftCenter: .TextMatrix(0, 46) = "RCC2"
    .ColWidth(45) = 1000: .ColAlignment(47) = flexAlignLeftCenter: .TextMatrix(0, 47) = "RCC3"
    .ColWidth(46) = 1000: .ColAlignment(48) = flexAlignLeftCenter: .TextMatrix(0, 48) = "WCC1"
    .ColWidth(47) = 1000: .ColAlignment(49) = flexAlignLeftCenter: .TextMatrix(0, 49) = "WCC2"
    .ColWidth(48) = 1000: .ColAlignment(50) = flexAlignLeftCenter: .TextMatrix(0, 50) = "WCC3"
    .ColWidth(49) = 1000: .ColAlignment(51) = flexAlignLeftCenter: .TextMatrix(0, 51) = "Polymorphic1"
    .ColWidth(50) = 1000: .ColAlignment(52) = flexAlignLeftCenter: .TextMatrix(0, 52) = "Polymorphic2"
    .ColWidth(51) = 1000: .ColAlignment(53) = flexAlignLeftCenter: .TextMatrix(0, 53) = "Polymorphic3"
    .ColWidth(52) = 1000: .ColAlignment(54) = flexAlignLeftCenter: .TextMatrix(0, 54) = "Mononucleated1"
    .ColWidth(53) = 1000: .ColAlignment(55) = flexAlignLeftCenter: .TextMatrix(0, 55) = "Mononucleated2"
    .ColWidth(54) = 1000: .ColAlignment(56) = flexAlignLeftCenter: .TextMatrix(0, 56) = "Mononucleated3"
    .ColWidth(55) = 1000: .ColAlignment(57) = flexAlignLeftCenter: .TextMatrix(0, 57) = "FluidGlucose"
    .ColWidth(56) = 1000: .ColAlignment(58) = flexAlignLeftCenter: .TextMatrix(0, 58) = "FluidProtein"
    .ColWidth(57) = 1000: .ColAlignment(59) = flexAlignLeftCenter: .TextMatrix(0, 59) = "FluidAlbumin"
    .ColWidth(58) = 1000: .ColAlignment(60) = flexAlignLeftCenter: .TextMatrix(0, 60) = "FluidGlobulin"
    .ColWidth(59) = 1000: .ColAlignment(61) = flexAlignLeftCenter: .TextMatrix(0, 61) = "FluidLDH"
    .ColWidth(60) = 1000: .ColAlignment(62) = flexAlignLeftCenter: .TextMatrix(0, 62) = "FluidAmylase"
    .ColWidth(61) = 1000: .ColAlignment(63) = flexAlignLeftCenter: .TextMatrix(0, 63) = "CSFGlucose"
    .ColWidth(62) = 1000: .ColAlignment(64) = flexAlignLeftCenter: .TextMatrix(0, 64) = "CSFProtein"
    .ColWidth(63) = 1250: .ColAlignment(65) = flexAlignLeftCenter: .TextMatrix(0, 65) = "FungalElements"
    .ColWidth(64) = 1000: .ColAlignment(66) = flexAlignLeftCenter: .TextMatrix(0, 66) = "PneumococcalAT"
    .ColWidth(65) = 1250: .ColAlignment(67) = flexAlignLeftCenter: .TextMatrix(0, 67) = "LegionellaAT"
    .ColWidth(66) = 1500: .ColAlignment(68) = flexAlignLeftCenter: .TextMatrix(0, 68) = "BATScreen"
    .ColWidth(67) = 1000: .ColAlignment(69) = flexAlignLeftCenter: .TextMatrix(0, 69) = "BATScreenComment"
    
    For i = 70 To .Cols - 1
        .ColWidth(i) = 0: .ColAlignment(i) = flexAlignLeftCenter: .TextMatrix(0, i) = ""
    Next

    .WordWrap = True
End With

Exit Sub

InitGPClinician_Error:

Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "frmMicroSurveillanceSearches", "InitGPClinician", intEL, strES

End Sub

'---------------------------------------------------------------------------------------
' Procedure : InitGrid2
' Author    : Trevor Dunican
' Date      : 28/10/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub InitGrid2()
      Dim i As Integer

10    On Error GoTo InitGrid2_Error

20    With g2
30        .Rows = 3: .Cols = 11
40        .FixedRows = 2: .FixedCols = 3
50        .Rows = 2
60        .SelectionMode = flexSelectionByRow
70        For i = 0 To 1
80            If i = 0 Then
90                .TextMatrix(0, 0) = "Laboratory"
100               .TextMatrix(0, 6) = "Age Group"
110               .TextMatrix(0, 8) = "M/F"
120               .TextMatrix(0, 9) = "Org"
130               .TextMatrix(0, 10) = "Site"
140           Else
150               .Rows = 2
160               .TextMatrix(1, 0) = "Test Preformed"
170               .TextMatrix(1, 1) = "Lab Results"
180               .TextMatrix(1, 2) = "Sex"
190               .TextMatrix(1, 3) = "0  -  4"
200               .TextMatrix(1, 4) = "5 - 14"
210               .TextMatrix(1, 5) = "15 - 44"
220               .TextMatrix(1, 6) = "45   -   60"
230               .TextMatrix(1, 7) = "60+"
240               .TextMatrix(1, 8) = "Total"
250               .TextMatrix(1, 9) = "Total"
260               .TextMatrix(1, 10) = "Total"
270           End If
280       Next i
290       .ColWidth(0) = 2790
300       .ColWidth(1) = 2800
310       .ColWidth(2) = 800
320       .ColWidth(3) = 800
330       .ColWidth(4) = 800
340       .ColWidth(5) = 800
350       .ColWidth(6) = 900
360       .ColWidth(7) = 800
370       .ColWidth(8) = 900
380       .ColWidth(9) = 900
390       .ColWidth(10) = 900
400       .ColAlignment(0) = flexAlignLeftCenter
410       .ColAlignment(1) = flexAlignLeftCenter
420       .ColAlignment(2) = flexAlignLeftCenter
430       .ColAlignment(3) = flexAlignLeftCenter
440       .ColAlignment(4) = flexAlignLeftCenter
450       .ColAlignment(5) = flexAlignLeftCenter
460       .ColAlignment(6) = flexAlignLeftCenter
470       .ColAlignment(7) = flexAlignLeftCenter
480       .ColAlignment(8) = flexAlignLeftCenter
490       .ColAlignment(9) = flexAlignLeftCenter
500       .ColAlignment(10) = flexAlignLeftCenter
510   End With

520   Exit Sub

InitGrid2_Error:

      Dim strES As String
      Dim intEL As Integer

530   intEL = Erl
540   strES = Err.Description
550   LogError "frmMicroSurveillanceSearches", "InitGrid2", intEL, strES
End Sub

'---------------------------------------------------------------------------------------
' Procedure : InitGrid4
' Author    : Farhan Waheed
' Date      : 27/04/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub InitGrid4()
      Dim i As Integer

10    On Error GoTo InitGrid4_Error

20    With g4
30        .Rows = 2: .Cols = 20
40        .FixedRows = 1: .FixedCols = 5
50        .Rows = 1
60        .SelectionMode = flexSelectionByRow


70                .TextMatrix(0, 0) = "SampleID"
80                .TextMatrix(0, 1) = "Patient Name"
90                .TextMatrix(0, 2) = "Sex"
100               .TextMatrix(0, 3) = "Age"
110               .TextMatrix(0, 4) = "SampleDate"
120               .TextMatrix(0, 5) = "Rota"
130               .TextMatrix(0, 6) = "Adeno"
140               .TextMatrix(0, 7) = "OBO"
150               .TextMatrix(0, 8) = "OB1"
160               .TextMatrix(0, 9) = "OB2"
170               .TextMatrix(0, 10) = "ToxinAB"
180               .TextMatrix(0, 11) = "Cryptosporidium"
190               .TextMatrix(0, 12) = "HPylori"
200               .TextMatrix(0, 13) = "CD-Culture"
210               .TextMatrix(0, 14) = "WCC"
220               .TextMatrix(0, 15) = "RCC"
230               .TextMatrix(0, 16) = "Crystals"
240               .TextMatrix(0, 17) = "Casts"
250               .TextMatrix(0, 18) = "RedSub"
260               .TextMatrix(0, 19) = "RSV"

270       .ColWidth(0) = 1500
280       .ColWidth(1) = 2000
290       .ColWidth(2) = 800
300       .ColWidth(3) = 800
310       .ColWidth(4) = 1200
320       .ColWidth(5) = 800
330       .ColWidth(6) = 900
340       .ColWidth(7) = 800
350       .ColWidth(8) = 800
360       .ColWidth(9) = 800
370       .ColWidth(10) = 800
380       .ColWidth(11) = 800
390       .ColWidth(12) = 800
400       .ColWidth(13) = 800
410       .ColWidth(14) = 800
420       .ColWidth(15) = 800
430       .ColWidth(16) = 800
440       .ColWidth(17) = 800
450       .ColWidth(18) = 800
460       .ColWidth(19) = 800
470       .ColAlignment(0) = flexAlignLeftCenter
480       .ColAlignment(1) = flexAlignLeftCenter
490       .ColAlignment(2) = flexAlignLeftCenter
500       .ColAlignment(3) = flexAlignLeftCenter
510       .ColAlignment(4) = flexAlignLeftCenter
520       .ColAlignment(5) = flexAlignLeftCenter
530       .ColAlignment(6) = flexAlignLeftCenter
540       .ColAlignment(7) = flexAlignLeftCenter
550       .ColAlignment(8) = flexAlignLeftCenter
560       .ColAlignment(9) = flexAlignLeftCenter
570       .ColAlignment(10) = flexAlignLeftCenter
580       .ColAlignment(11) = flexAlignLeftCenter
590       .ColAlignment(12) = flexAlignLeftCenter
600       .ColAlignment(13) = flexAlignLeftCenter
610       .ColAlignment(14) = flexAlignLeftCenter
620       .ColAlignment(15) = flexAlignLeftCenter
630       .ColAlignment(16) = flexAlignLeftCenter
640       .ColAlignment(17) = flexAlignLeftCenter
650       .ColAlignment(18) = flexAlignLeftCenter
660       .ColAlignment(19) = flexAlignLeftCenter
670   End With

680   Exit Sub

InitGrid4_Error:

      Dim strES As String
      Dim intEL As Integer

690   intEL = Erl
700   strES = Err.Description
710   LogError "frmMicroSurveillanceSearches", "InitGrid4", intEL, strES
End Sub
'---------------------------------------------------------------------------------------
' Procedure : InitGrid4
' Author    : Farhan Waheed
' Date      : 27/04/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub InitGrid7()
      Dim i As Integer

10    On Error GoTo InitGrid4_Error

20    With G7
30        .Rows = 2: .Cols = 19
40        .FixedRows = 1: .FixedCols = 3
50        .Rows = 1
60        .SelectionMode = flexSelectionByRow


70                .TextMatrix(0, 0) = "Site"
80                .TextMatrix(0, 1) = "Age Group"
90                .TextMatrix(0, 2) = "Sex"
100               .TextMatrix(0, 3) = "Rota"
110               .TextMatrix(0, 4) = "Adeno"
120               .TextMatrix(0, 5) = "OBO"
130               .TextMatrix(0, 6) = "OB1"
140               .TextMatrix(0, 7) = "OB2"
150               .TextMatrix(0, 8) = "ToxAB"
160               .TextMatrix(0, 9) = "Crypto"
170               .TextMatrix(0, 10) = "H.Pyl"
180               .TextMatrix(0, 11) = "CDiff"
190               .TextMatrix(0, 12) = "WCC"
200               .TextMatrix(0, 13) = "RCC"
210               .TextMatrix(0, 14) = "Cryst"
220               .TextMatrix(0, 15) = "Casts"
230               .TextMatrix(0, 16) = "RSub"
240               .TextMatrix(0, 17) = "RSV"
250               .TextMatrix(0, 18) = "Total"

260       .ColWidth(0) = 2000
270       .ColWidth(1) = 1200
280       .ColWidth(2) = 600
290       .ColWidth(3) = 600
300       .ColWidth(4) = 600
310       .ColWidth(5) = 600
320       .ColWidth(6) = 600
330       .ColWidth(7) = 600
340       .ColWidth(8) = 600
350       .ColWidth(9) = 600
360       .ColWidth(10) = 600
370       .ColWidth(11) = 600
380       .ColWidth(12) = 600
390       .ColWidth(13) = 600
400       .ColWidth(14) = 600
410       .ColWidth(15) = 600
420       .ColWidth(16) = 600
430       .ColWidth(17) = 600
440       .ColWidth(18) = 600

450       .ColAlignment(0) = flexAlignLeftCenter
460       .ColAlignment(1) = flexAlignLeftCenter
470       .ColAlignment(2) = flexAlignLeftCenter
480       .ColAlignment(3) = flexAlignLeftCenter
490       .ColAlignment(4) = flexAlignLeftCenter
500       .ColAlignment(5) = flexAlignLeftCenter
510       .ColAlignment(6) = flexAlignLeftCenter
520       .ColAlignment(7) = flexAlignLeftCenter
530       .ColAlignment(8) = flexAlignLeftCenter
540       .ColAlignment(9) = flexAlignLeftCenter
550       .ColAlignment(10) = flexAlignLeftCenter
560       .ColAlignment(11) = flexAlignLeftCenter
570       .ColAlignment(12) = flexAlignLeftCenter
580       .ColAlignment(13) = flexAlignLeftCenter
590       .ColAlignment(14) = flexAlignLeftCenter
600       .ColAlignment(15) = flexAlignLeftCenter
610       .ColAlignment(16) = flexAlignLeftCenter
620       .ColAlignment(17) = flexAlignLeftCenter
630       .ColAlignment(18) = flexAlignLeftCenter

640   End With

650   Exit Sub

InitGrid4_Error:

      Dim strES As String
      Dim intEL As Integer

660   intEL = Erl
670   strES = Err.Description
680   LogError "frmMicroSurveillanceSearches", "InitGrid4", intEL, strES
End Sub
'---------------------------------------------------------------------------------------
' Procedure : InitGrid
' Author    : Trevor Dunican
' Date      : 26/11/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub InitGrid()

10    On Error GoTo InitGrid_Error

20    With g
30        .Rows = 2: .Cols = 10
40        .FixedRows = 1: .FixedCols = 1
50        .Rows = 1
60        .SelectionMode = flexSelectionByRow
70        .TextMatrix(0, 0) = "Sample ID"
80        .TextMatrix(0, 1) = "Site"
90        .TextMatrix(0, 2) = "Organism Group"
100       .TextMatrix(0, 3) = "Organism Name"
110       .TextMatrix(0, 4) = "Qualifier"
120       .TextMatrix(0, 5) = "Result"
130       .TextMatrix(0, 6) = "Antibiotic"
140       .TextMatrix(0, 7) = "RSI"
150       .TextMatrix(0, 8) = "Date"
160       .TextMatrix(0, 9) = "C"
170       .ColWidth(0) = 100
180       .ColWidth(1) = 100
190       .ColWidth(2) = 100
200       .ColWidth(3) = 100
210       .ColWidth(4) = 100
220       .ColWidth(5) = 100
230       .ColWidth(6) = 100
240       .ColWidth(7) = 100
250       .ColWidth(8) = 100
260       .ColWidth(9) = 100
270       .ColAlignment(0) = flexAlignLeftCenter
280       .ColAlignment(1) = flexAlignLeftCenter
290       .ColAlignment(2) = flexAlignLeftCenter
300       .ColAlignment(3) = flexAlignLeftCenter
310       .ColAlignment(4) = flexAlignLeftCenter
320       .ColAlignment(5) = flexAlignLeftCenter
330       .ColAlignment(6) = flexAlignLeftCenter
340       .ColAlignment(7) = flexAlignLeftCenter
350       .ColAlignment(8) = flexAlignLeftCenter
360       .ColAlignment(9) = flexAlignLeftCenter
370   End With

380   Exit Sub

InitGrid_Error:

      Dim strES As String
      Dim intEL As Integer

390   intEL = Erl
400   strES = Err.Description
410   LogError "frmMicroSurveillanceSearches", "InitGrid", intEL, strES

End Sub

Private Function OrganismGroupFilter()
      Dim i As Integer
      Dim J As Integer
      Dim K As Integer
      Dim siteName As String
      Dim good As String
      Dim w As Integer

10    On Error GoTo OrganismGroupFilter_Error

20    fraMainProgressBar.Visible = True
30    For i = 0 To lstOrganismGroup.ListCount - 1
40        If lstOrganismGroup.Selected(i) = False Then
50            For J = 0 To g.Rows - 1
60                If g.TextMatrix(J, 2) = lstOrganismGroup.List(i) Then
70                    g.RowHeight(J) = 0
80                End If
90            Next J
100       End If
110       If lstOrganismGroup.Selected(i) = True Then
120           For J = 0 To g.Rows - 1
130               If g.TextMatrix(J, 2) = lstOrganismGroup.List(i) Then
140                   g.RowHeight(J) = 275
150               End If
160           Next J
170       End If
180       MainProgressBar.Value = MainProgressBar.Value + 1
190       If MainProgressBar.Value = 100 Then
200           MainProgressBar.Value = 0
210       End If
220   Next i
230   recordCount = 0
240   For i = 0 To g.Rows - 1
250       If g.RowHeight(i) > 0 Then
260           recordCount = recordCount + 1
270       End If
280   Next i
290   txtResultTotal.Text = recordCount - 1
300   fraMainProgressBar.Visible = False

310   Exit Function

OrganismGroupFilter_Error:

      Dim strES As String
      Dim intEL As Integer

320   intEL = Erl
330   strES = Err.Description
340   LogError "frmMicroSurveillanceSearches", "OrganismGroupFilter", intEL, strES
End Function

Private Function SitesFilter()
      Dim i As Integer
      Dim J As Integer
      Dim siteName As String
      Dim good As String

10    On Error GoTo SitesFilter_Error

20    fraMainProgressBar.Visible = True
30    For i = 0 To lstSites.ListCount - 1
40        If lstSites.Selected(i) = False Then
50            For J = 0 To g.Rows - 1
60                If g.TextMatrix(J, 1) = lstSites.List(i) Then
70                    g.RowHeight(J) = 0
80                End If
90            Next J
100       End If
110       If lstSites.Selected(i) = True Then
120           For J = 0 To g.Rows - 1
130               If g.TextMatrix(J, 1) = lstSites.List(i) Then
140                   g.RowHeight(J) = 275
150               End If
160           Next J
170       End If
180       MainProgressBar.Value = MainProgressBar.Value + 1
190       If MainProgressBar.Value = 100 Then
200           MainProgressBar.Value = 0
210       End If
220   Next i
230   recordCount = 0
240   For i = 0 To g.Rows - 1
250       If g.RowHeight(i) > 0 Then
260           recordCount = recordCount + 1
270       End If
280   Next i
290   txtResultTotal.Text = recordCount - 1
300   fraMainProgressBar.Visible = False

310   Exit Function

SitesFilter_Error:

      Dim strES As String
      Dim intEL As Integer

320   intEL = Erl
330   strES = Err.Description
340   LogError "frmMicroSurveillanceSearches", "SitesFilter", intEL, strES

End Function

'---------------------------------------------------------------------------------------
' Procedure : getDemographics
' Author    : Trevor Dunican
' Date      : 26/11/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Function getDemographics()
      Dim sql As String
      Dim tb As Recordset
      Dim tbr As Recordset
      Dim tbO As Recordset
      Dim tbS As Recordset
      Dim s As String
      Dim n As Integer
      Dim X As Integer
      Dim sum As Long
      Dim Y As Integer
      Dim SiteAdded As Boolean
      Dim OrgAdded As Boolean
      Dim SiteCriteria As String
      Dim StartIndex As Integer
      Dim EndIndex As Integer
      Dim i As Integer

10    On Error GoTo getDemographics_Error

20    InitGrid2
30    fraCSProgressBar.Visible = True
40    CSProgressBar.Value = 0

50    For i = 0 To lstSites.ListCount - 1
60        If lstSites.Selected(i) Then
70            SiteCriteria = SiteCriteria & "'" & lstSites.List(i) & "" & "',"
80        End If
90    Next

100   If Trim(SiteCriteria) = "" Then Exit Function
110   SiteCriteria = Left(SiteCriteria, Len(SiteCriteria) - 1)
120   sql = "SELECT DISTINCT Site " & _
            "FROM MicroSiteDetails M JOIN Demographics D " & _
            "ON M.SampleID = D.SampleID " & _
            "WHERE D.SampleDate BETWEEN '" & Format$(dtFrom, "yyyymmdd") & " 00:00:00' " & _
            "AND '" & Format$(dtTo, "yyyymmdd") & " 23:59:59' " & _
            "AND Site in (" & SiteCriteria & ")"
130   Set tb = New Recordset
140   RecOpenClient 0, tb, sql
150   If Not tb.EOF Then
160       Do While Not tb.EOF
170           SiteAdded = False
180           sql = "SELECT DISTINCT(OrganismGroup) FROM Isolates I " & _
                    "INNER JOIN Demographics D ON I.SampleID = D.SampleID " & _
                    "INNER JOIN MicroSiteDetails M ON M.SampleID = I.SampleID " & _
                    "WHERE Site = '" & tb!Site & "' " & _
                    "AND D.SampleDate BETWEEN '" & Format$(dtFrom, "yyyymmdd") & " 00:00:00' " & _
                    "AND '" & Format$(dtTo, "yyyymmdd") & " 23:59:59' "
190           If optType(1).Value = True Then
200               sql = sql & "AND OrganismGroup <> 'Negative Results' "
210           ElseIf optType(2).Value = True Then
220               sql = sql & "AND OrganismGroup = 'Negative Results' "
230           End If
240           Set tbO = New Recordset
250           RecOpenServer 0, tbO, sql
260           Do While Not tbO.EOF
270               OrgAdded = False
280               If Not SiteAdded Then
290                   g2.AddItem tb!Site & ""
300                   SiteAdded = True
310               Else
320                   g2.AddItem ""
330               End If
340               If Not OrgAdded Then
350                   g2.TextMatrix(g2.Rows - 1, 1) = tbO!OrganismGroup & ""
360                   OrgAdded = True
370               Else
380                   g2.TextMatrix(g2.Rows - 1, 1) = ""
390               End If
400               g2.TextMatrix(g2.Rows - 1, 2) = "Male"
410               For n = 0 To 4
420                   sql = "SELECT COUNT(*) Total " & _
                            "FROM MicroSiteDetails M INNER JOIN Isolates I ON M.SampleID = I.SampleID " & _
                            "INNER JOIN Demographics D ON M.SampleID = D.SampleID " & _
                            "WHERE D.SampleDate BETWEEN '" & Format$(dtFrom, "yyyymmdd") & " 00:00:00' " & _
                            "AND '" & Format$(dtTo, "yyyymmdd") & " 23:59:59' " & _
                            "AND DATEDIFF(Day, DoB, SampleDate) > " & AgeGroupFrom(n) & " " & _
                            "AND DATEDIFF(Day, DoB, SampleDate) < " & AgeGroupTo(n) & " " & _
                            "AND Site = '" & tb!Site & "' " & _
                            "AND OrganismGroup = '" & tbO!OrganismGroup & "' " & _
                            "AND D.Sex = 'M'"

430                   Set tbr = New Recordset
440                   RecOpenServer 0, tbr, sql
450                   CSProgressBar.Value = CSProgressBar.Value + 1
460                   lblCSProgressBar = "Fetching results ... (" & Int(CSProgressBar.Value * 100 / CSProgressBar.Max) & " %)"
470                   If CSProgressBar.Value = 100 Then
480                       CSProgressBar.Value = 0
490                   End If
500                   lblCSProgressBar.Refresh
510                   If tbr!Total > 0 Then
520                       g2.TextMatrix(g2.Rows - 1, n + 3) = tbr!Total
530                   End If
540               Next
550               If Not SiteAdded Then
560                   g2.AddItem tb!Site & ""
570                   SiteAdded = True
580               Else
590                   g2.AddItem ""
600               End If
610               If Not OrgAdded Then
620                   g2.TextMatrix(g2.Rows - 1, 1) = tbO!OrganismGroup & ""
630                   OrgAdded = True
640               Else
650                   g2.TextMatrix(g2.Rows - 1, 1) = ""
660               End If
670               g2.TextMatrix(g2.Rows - 1, 2) = "Female"
680               For n = 0 To 4
690                   sql = "SELECT COUNT(*) Total " & _
                            "FROM MicroSiteDetails M INNER JOIN Isolates I ON M.SampleID = I.SampleID " & _
                            "INNER JOIN Demographics D ON M.SampleID = D.SampleID " & _
                            "WHERE D.SampleDate BETWEEN '" & Format$(dtFrom, "yyyymmdd") & " 00:00:00' " & _
                            "AND '" & Format$(dtTo, "yyyymmdd") & " 23:59:59' " & _
                            "AND DATEDIFF(Day, DoB, SampleDate) > " & AgeGroupFrom(n) & " " & _
                            "AND DATEDIFF(Day, DoB, SampleDate) < " & AgeGroupTo(n) & " " & _
                            "AND Site = '" & tb!Site & "' " & _
                            "AND OrganismGroup = '" & tbO!OrganismGroup & "' " & _
                            "AND D.Sex = 'F'"
700                   Set tbr = New Recordset
710                   RecOpenServer 0, tbr, sql
720                   CSProgressBar.Value = CSProgressBar.Value + 1
730                   lblCSProgressBar = "Fetching results ... (" & Int(CSProgressBar.Value * 100 / CSProgressBar.Max) & " %)"
740                   If CSProgressBar.Value = 100 Then
750                       CSProgressBar.Value = 0
760                   End If
770                   lblCSProgressBar.Refresh
780                   If tbr!Total > 0 Then
790                       g2.TextMatrix(g2.Rows - 1, n + 3) = tbr!Total
800                   End If
810               Next
820               tbO.MoveNext
830           Loop
840           tb.MoveNext
850       Loop
860   End If
      'Sex Total
870   For n = 2 To g2.Rows - 1
880       sum = 0
890       For X = 3 To 7
900           sum = sum + Val(g2.TextMatrix(n, X))
910       Next
920       g2.TextMatrix(n, 8) = sum
930   Next
      'Organism Total
940   For n = 2 To g2.Rows - 2 Step 2
950       sum = Val(g2.TextMatrix(n, 8)) + Val(g2.TextMatrix(n + 1, 8))
960       g2.TextMatrix(n, 9) = sum
970   Next
980   g2.AddItem ""
      'Site Total
990   sum = 0
1000  For X = 2 To g2.Rows - 1
1010      If g2.TextMatrix(X, 0) <> "" Then
1020          StartIndex = X
1030          For Y = StartIndex + 1 To g2.Rows - 1
1040              If g2.TextMatrix(Y, 0) <> "" Then
1050                  EndIndex = Y - 1
1060                  Exit For
1070              End If
1080          Next Y
1090          If StartIndex > EndIndex Then
1100              EndIndex = Y - 1
1110          End If
1120          For n = StartIndex To EndIndex
1130              sum = sum + Val(g2.TextMatrix(n, 9))
1140          Next n
1150          g2.TextMatrix(StartIndex, 10) = sum
1160          sum = 0
1170      End If
1180  Next X
1190  fraCSProgressBar.Visible = False
1200  Exit Function

getDemographics_Error:

      Dim strES As String
      Dim intEL As Integer

1210  intEL = Erl
1220  strES = Err.Description
1230  LogError "frmMicroSurveillanceSearches", "getDemographics", intEL, strES, sql
End Function

Private Function RealColPos(Col As Integer, grid As MSFlexGrid)
      Dim i As Integer
      Dim merged As Integer

10    On Error GoTo RealColPos_Error

20    With grid
30        i = Col - 1: merged = 0
40        Do While .ColPos(Col) = .ColPos(i)
50            merged = merged + 1
60            i = i - 1
70        Loop
80        If merged > 0 Then
90            RealColPos = .ColPos(Col - merged)
100           Do While merged > 0
110               RealColPos = RealColPos + .ColWidth(Col - merged)
120               merged = merged - 1
130           Loop
140       Else
150           RealColPos = .ColPos(Col)
160       End If
170   End With

180   Exit Function

RealColPos_Error:

      Dim strES As String
      Dim intEL As Integer

190   intEL = Erl
200   strES = Err.Description
210   LogError "frmMicroSurveillanceSearches", "RealColPos", intEL, strES

End Function

Private Sub FillG3()

      Dim sql As String
      Dim tb As Recordset
      Dim test As String
      Dim Obs As New Observations
      Dim SiteCriteria As String

10    On Error GoTo FillG3_Error

20    sql = "SELECT Text FROM Lists WHERE ListType = 'MicroSS'"
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql
50    If tb.EOF Then
60        Exit Sub
70    End If
80    SiteCriteria = ""
90    While Not tb.EOF
100       SiteCriteria = SiteCriteria & "'" & tb!Text & "" & "',"
110       tb.MoveNext
120   Wend
130   If Trim(SiteCriteria) = "" Then Exit Sub
140   SiteCriteria = Left(SiteCriteria, Len(SiteCriteria) - 1)
150   If optType(0) Then
160       sql = "SELECT I.SampleID, D.Chart, D.PatName, D.Age, D.DoB, D.GP, D.SEX, M.Site, I.OrganismGroup, I.OrganismName, S.Result, S.Antibiotic, " & _
                "S.RSI, I.Qualifier, I.RecordDateTime from Isolates I " & _
                "Inner Join MicroSiteDetails M ON I.SampleID = M.SampleID " & _
                "Inner Join Demographics D ON I.SampleID = D.SampleID " & _
                "Inner Join Sensitivities S ON I.SampleID = S.SampleID  And I.IsolateNumber = S.IsolateNumber " & _
                "WHERE I.RecordDateTime between '" & calFrom & "' and '" & calTo & "' and Site in (" & SiteCriteria & ") " & _
                "ORDER by Site, I.RecordDateTime DESC"
170   End If
180   Set tb = New Recordset
190   RecOpenServer 0, tb, sql
200   g3.Rows = 1
210   fraProgress.Visible = True
220   prgFetchingResults.Value = 0
230   Do While Not tb.EOF
240       g3.AddItem Format(tb!RecordDateTime, "dd mmm yyyy") & vbTab & tb!SampleID - SysOptMicroOffset(0) & vbTab & tb!Chart & vbTab & tb!Dob & vbTab & tb!sex _
                     & vbTab & tb!GP & vbTab & tb!Site & vbTab & tb!OrganismGroup _
                     & vbTab & tb!OrganismName & vbTab & tb!Qualifier & vbTab & tb!Result & vbTab & tb!Antibiotic _
                     & vbTab & tb!RSI & ""
250       Set Obs = Obs.Load(tb!SampleID, _
                             "MicroGeneral", "Demographic", "MicroCS", _
                             "MicroConsultant", "CSFFluid", "MicroCDiff")
260       If Not OrganismExistsInList(tb!OrganismGroup & "") Then
270           lstOrganismGroup.AddItem tb!OrganismGroup & ""
280           lstOrganismGroup.Selected(lstOrganismGroup.NewIndex) = True
290       End If
300       If Not SiteExistsInList(tb!Site & "") Then
310           lstSites.AddItem tb!Site & ""
320           lstSites.Selected(lstSites.NewIndex) = True
330       End If
340       MainProgressBar.Value = MainProgressBar.Value + 1
350       lblMainProgressBar = "Fetching results ... (" & Int(MainProgressBar.Value * 100 / MainProgressBar.Max) & " %)"
360       If MainProgressBar.Value = 100 Then
370           MainProgressBar.Value = 0
380       End If
390       lblMainProgressBar.Refresh
400       tb.MoveNext
410   Loop
420   fraProgress.Visible = False

      Dim introw As Integer
      Dim intcol As Integer
430   With g3
440       For intcol = 0 To .Cols - 1
450           For introw = 0 To .Rows - 1
460               If .ColWidth(intcol) < frmMicroSurveillanceSearches.TextWidth(.TextMatrix(introw, intcol)) + 100 And intcol <> 1 And intcol <> 2 Then
470                   .ColWidth(intcol) = frmMicroSurveillanceSearches.TextWidth(.TextMatrix(introw, intcol)) + 100
480               End If
490           Next
500       Next
510   End With

520   Exit Sub

FillG3_Error:

      Dim strES As String
      Dim intEL As Integer

530   intEL = Erl
540   strES = Err.Description
550   LogError "frmMicroSurveillanceSearches", "FillG3", intEL, strES, sql

End Sub

Private Sub optGPClinician_Click(Index As Integer)
If Index = 2 Then
FillGPsClinWard Me, HospName(0)
End If
End Sub
