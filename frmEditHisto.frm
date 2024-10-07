VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEditHisto 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   $"frmEditHisto.frx":0000
   ClientHeight    =   10065
   ClientLeft      =   75
   ClientTop       =   390
   ClientWidth     =   14130
   Icon            =   "frmEditHisto.frx":0116
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10065
   ScaleWidth      =   14130
   ShowInTaskbar   =   0   'False
   Begin ComCtl2.UpDown udNoCopies 
      Height          =   375
      Left            =   13530
      TabIndex        =   200
      Top             =   3623
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   327681
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txtNoCopies 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   13020
      Locked          =   -1  'True
      TabIndex        =   199
      Text            =   "3"
      Top             =   3600
      Width           =   495
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1935
      Left            =   180
      TabIndex        =   125
      Top             =   360
      Width           =   13230
      Begin VB.Frame Frame6 
         Height          =   1890
         Left            =   0
         TabIndex        =   149
         Top             =   0
         Width           =   2790
         Begin VB.TextBox txtSampleID 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   1170
            MaxLength       =   12
            TabIndex        =   0
            Top             =   495
            Width           =   1035
         End
         Begin VB.ComboBox cMRU 
            Height          =   315
            Left            =   585
            TabIndex        =   151
            Top             =   1305
            Width           =   2070
         End
         Begin VB.TextBox txtYear 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   45
            TabIndex        =   150
            Top             =   495
            Width           =   930
         End
         Begin ComCtl2.UpDown UpDown1 
            Height          =   480
            Left            =   2205
            TabIndex        =   152
            Top             =   495
            Width           =   240
            _ExtentX        =   450
            _ExtentY        =   847
            _Version        =   327681
            Value           =   1
            BuddyControl    =   "txtSampleID"
            BuddyDispid     =   196612
            OrigLeft        =   2025
            OrigTop         =   510
            OrigRight       =   2265
            OrigBottom      =   990
            Max             =   999999999
            Min             =   1
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.Label Label19 
            Caption         =   "/"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   1035
            TabIndex        =   157
            Top             =   495
            Width           =   195
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            Caption         =   "Sample ID"
            Height          =   195
            Left            =   720
            TabIndex        =   156
            Top             =   0
            Width           =   735
         End
         Begin VB.Image iRelevant 
            Appearance      =   0  'Flat
            Height          =   480
            Index           =   0
            Left            =   435
            Picture         =   "frmEditHisto.frx":0420
            Top             =   90
            Width           =   480
         End
         Begin VB.Image iRelevant 
            Height          =   480
            Index           =   1
            Left            =   1800
            Picture         =   "frmEditHisto.frx":072A
            Top             =   90
            Width           =   480
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "MRU"
            Height          =   195
            Left            =   135
            TabIndex        =   155
            Top             =   1350
            Width           =   375
         End
         Begin VB.Image imgLast 
            Height          =   255
            Left            =   2295
            Picture         =   "frmEditHisto.frx":0A34
            Stretch         =   -1  'True
            ToolTipText     =   "Find Last Record"
            Top             =   180
            Width           =   405
         End
         Begin VB.Label lblResultOrRequest 
            Alignment       =   2  'Center
            BackColor       =   &H00FF8080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Results"
            Height          =   285
            Left            =   915
            TabIndex        =   154
            Top             =   210
            Width           =   885
         End
         Begin VB.Label lblDisp 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "H"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   2475
            TabIndex        =   153
            Top             =   495
            Width           =   285
         End
      End
      Begin VB.TextBox txtAandE 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   4320
         TabIndex        =   2
         Top             =   570
         Width           =   1245
      End
      Begin VB.TextBox txtNOPAS 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   5580
         TabIndex        =   3
         Top             =   570
         Width           =   1245
      End
      Begin VB.TextBox txtChart 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   2880
         MaxLength       =   8
         TabIndex        =   1
         Top             =   540
         Width           =   1425
      End
      Begin VB.TextBox txtName 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   6840
         MaxLength       =   30
         TabIndex        =   4
         Tag             =   "tName"
         Top             =   540
         Width           =   3495
      End
      Begin VB.TextBox txtDoB 
         Height          =   285
         Left            =   10920
         MaxLength       =   10
         TabIndex        =   6
         Top             =   270
         Width           =   1545
      End
      Begin VB.TextBox txtAge 
         Height          =   285
         Left            =   10920
         MaxLength       =   4
         TabIndex        =   8
         Top             =   630
         Width           =   1545
      End
      Begin VB.TextBox txtSex 
         Height          =   285
         Left            =   10920
         MaxLength       =   6
         TabIndex        =   9
         Top             =   990
         Width           =   1545
      End
      Begin VB.CommandButton bsearch 
         Appearance      =   0  'Flat
         Caption         =   "Se&arch"
         Height          =   345
         Left            =   9660
         TabIndex        =   5
         Top             =   180
         Width           =   675
      End
      Begin VB.CommandButton bDoB 
         Appearance      =   0  'Flat
         Caption         =   "S&earch"
         Height          =   285
         Left            =   12510
         TabIndex        =   7
         Top             =   240
         Width           =   660
      End
      Begin VB.Label lblDemographicComment 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   510
         Left            =   2880
         TabIndex        =   141
         ToolTipText     =   "Demographic Comment"
         Top             =   1350
         Width           =   9645
      End
      Begin VB.Label lblRundate 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   675
         TabIndex        =   135
         Top             =   495
         Width           =   1515
      End
      Begin VB.Label lblAandE 
         Caption         =   "A and E"
         Height          =   225
         Left            =   4545
         TabIndex        =   134
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblNOPAS 
         AutoSize        =   -1  'True
         Caption         =   "NOPAS"
         Height          =   195
         Index           =   0
         Left            =   5610
         TabIndex        =   133
         Top             =   390
         Width           =   555
      End
      Begin VB.Label lblChartNumber 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Monaghan Chart #"
         Height          =   285
         Left            =   2880
         TabIndex        =   132
         ToolTipText     =   "Click to change Location"
         Top             =   270
         Width           =   1425
      End
      Begin VB.Label lAddWardGP 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2880
         TabIndex        =   131
         Top             =   1050
         Width           =   7455
      End
      Begin VB.Label lNoPrevious 
         AutoSize        =   -1  'True
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "No Previous Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   7860
         TabIndex        =   130
         Top             =   210
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.Label lblNameTitle 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         Height          =   195
         Left            =   6840
         TabIndex        =   129
         Top             =   360
         Width           =   420
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         Caption         =   "D.o.B"
         Height          =   195
         Left            =   10470
         TabIndex        =   128
         Top             =   300
         Width           =   405
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "Age"
         Height          =   195
         Left            =   10560
         TabIndex        =   127
         Top             =   660
         Width           =   285
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         Caption         =   "Sex"
         Height          =   195
         Left            =   10590
         TabIndex        =   126
         Top             =   1020
         Width           =   270
      End
   End
   Begin VB.CommandButton cmdSetPrinter 
      Height          =   615
      Left            =   12780
      Picture         =   "frmEditHisto.frx":0E76
      Style           =   1  'Graphical
      TabIndex        =   57
      ToolTipText     =   "Back Colour Normal:Automatic Printer Selection.Back Colour Red:-Forced"
      Top             =   4320
      Width           =   1275
   End
   Begin VB.Timer TimerBar 
      Interval        =   1000
      Left            =   11970
      Top             =   -30
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   135
      Left            =   390
      TabIndex        =   56
      Top             =   120
      Width           =   13545
      _ExtentX        =   23892
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   1
      Max             =   600
      Scrolling       =   1
   End
   Begin VB.CommandButton bPrintHold 
      Caption         =   "Print && Hold"
      Height          =   885
      Left            =   12780
      Picture         =   "frmEditHisto.frx":1180
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   5040
      Width           =   1275
   End
   Begin VB.CommandButton bHistory 
      Caption         =   "&History"
      Height          =   885
      Left            =   12795
      Picture         =   "frmEditHisto.frx":148A
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   7890
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.CommandButton bPrint 
      Caption         =   "&Print"
      Height          =   885
      Left            =   12780
      Picture         =   "frmEditHisto.frx":18CC
      Style           =   1  'Graphical
      TabIndex        =   35
      Tag             =   "bprint"
      Top             =   5985
      Width           =   1275
   End
   Begin VB.CommandButton bFAX 
      Caption         =   "FAX"
      Height          =   885
      Index           =   0
      Left            =   12780
      Picture         =   "frmEditHisto.frx":1BD6
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   6930
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.CommandButton bCancel 
      Appearance      =   0  'Flat
      Caption         =   "Exit"
      Height          =   825
      Left            =   12795
      Picture         =   "frmEditHisto.frx":1EE0
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   8850
      Width           =   1275
   End
   Begin TabDlg.SSTab ssTab1 
      Height          =   7395
      Left            =   180
      TabIndex        =   24
      Top             =   2340
      Width           =   12570
      _ExtentX        =   22172
      _ExtentY        =   13044
      _Version        =   393216
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Demographics"
      TabPicture(0)   =   "frmEditHisto.frx":21EA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "bSaveHold"
      Tab(0).Control(1)=   "bSave"
      Tab(0).Control(2)=   "cmdDemoVal"
      Tab(0).Control(3)=   "Frame7"
      Tab(0).Control(4)=   "Frame5"
      Tab(0).Control(5)=   "Frame4"
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Histology Work Screen"
      TabPicture(1)   =   "frmEditHisto.frx":2206
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lblStatus"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "fraSpec(0)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "fraSpec(1)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "fraSpec(2)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "fraSpec(3)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "fraSpec(4)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "fraSpec(5)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "cmdSaveHisto(1)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "cmdSaveHHold(1)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).ControlCount=   10
      TabCaption(2)   =   "Histology Report"
      TabPicture(2)   =   "frmEditHisto.frx":2222
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "chkNCRI(0)"
      Tab(2).Control(1)=   "cmdViewHistoRep"
      Tab(2).Control(2)=   "cmdSaveHHold(0)"
      Tab(2).Control(3)=   "cmdSaveHisto(0)"
      Tab(2).Control(4)=   "cmdHVal"
      Tab(2).Control(5)=   "txtHistoComment"
      Tab(2).Control(6)=   "t"
      Tab(2).Control(7)=   "lblVal"
      Tab(2).Control(8)=   "lblD"
      Tab(2).Control(9)=   "lblC"
      Tab(2).Control(10)=   "lblB"
      Tab(2).Control(11)=   "lblA"
      Tab(2).Control(12)=   "Label18"
      Tab(2).ControlCount=   13
      TabCaption(3)   =   "Cytology"
      TabPicture(3)   =   "frmEditHisto.frx":223E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txtCytoComment"
      Tab(3).Control(1)=   "chkNCRI(1)"
      Tab(3).Control(2)=   "cmdViewMicroRep"
      Tab(3).Control(3)=   "cmdSaveCHold"
      Tab(3).Control(4)=   "cmdSaveCyto"
      Tab(3).Control(5)=   "c(13)"
      Tab(3).Control(6)=   "cmdCVal"
      Tab(3).Control(7)=   "Frame3"
      Tab(3).Control(8)=   "txtCyto"
      Tab(3).Control(9)=   "SSPanel2"
      Tab(3).Control(10)=   "SSPanel3"
      Tab(3).Control(11)=   "Label9"
      Tab(3).Control(12)=   "iCLocked"
      Tab(3).Control(13)=   "iCUnlocked"
      Tab(3).Control(14)=   "iCKey"
      Tab(3).Control(15)=   "iCDelete"
      Tab(3).Control(16)=   "iCDate(0)"
      Tab(3).Control(17)=   "Label12"
      Tab(3).ControlCount=   18
      Begin VB.TextBox txtCytoComment 
         Height          =   600
         Left            =   -74640
         MultiLine       =   -1  'True
         TabIndex        =   192
         Top             =   5490
         Width           =   11985
      End
      Begin VB.CheckBox chkNCRI 
         Caption         =   "Report to NCRI"
         Height          =   285
         Index           =   1
         Left            =   -74595
         TabIndex        =   191
         Top             =   6300
         Width           =   1500
      End
      Begin VB.CheckBox chkNCRI 
         Caption         =   "Report to NCRI"
         Height          =   285
         Index           =   0
         Left            =   -74370
         TabIndex        =   175
         Top             =   6705
         Width           =   1500
      End
      Begin VB.CommandButton cmdViewHistoRep 
         Caption         =   "View Reports"
         Height          =   870
         Left            =   -67395
         Picture         =   "frmEditHisto.frx":225A
         Style           =   1  'Graphical
         TabIndex        =   169
         Top             =   6390
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.CommandButton cmdViewMicroRep 
         Caption         =   "View Reports"
         Height          =   870
         Left            =   -67575
         Picture         =   "frmEditHisto.frx":2564
         Style           =   1  'Graphical
         TabIndex        =   168
         Top             =   6120
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.CommandButton cmdSaveHHold 
         Caption         =   "Save && &Hold"
         Enabled         =   0   'False
         Height          =   870
         Index           =   1
         Left            =   11385
         Picture         =   "frmEditHisto.frx":286E
         Style           =   1  'Graphical
         TabIndex        =   166
         Top             =   6030
         Width           =   1065
      End
      Begin VB.CommandButton cmdSaveHisto 
         Caption         =   "&Save Details"
         Enabled         =   0   'False
         Height          =   870
         Index           =   1
         Left            =   11385
         Picture         =   "frmEditHisto.frx":2B78
         Style           =   1  'Graphical
         TabIndex        =   165
         Top             =   5130
         Width           =   1095
      End
      Begin VB.CommandButton cmdSaveCHold 
         Caption         =   "Save && &Hold"
         Enabled         =   0   'False
         Height          =   870
         Left            =   -63705
         Picture         =   "frmEditHisto.frx":2E82
         Style           =   1  'Graphical
         TabIndex        =   164
         Top             =   6120
         Width           =   1155
      End
      Begin VB.CommandButton cmdSaveCyto 
         Caption         =   "&Save Details"
         Enabled         =   0   'False
         Height          =   870
         Left            =   -64875
         Picture         =   "frmEditHisto.frx":318C
         Style           =   1  'Graphical
         TabIndex        =   163
         Top             =   6120
         Width           =   1095
      End
      Begin VB.CommandButton cmdSaveHHold 
         Caption         =   "Save && &Hold"
         Enabled         =   0   'False
         Height          =   870
         Index           =   0
         Left            =   -63750
         Picture         =   "frmEditHisto.frx":3496
         Style           =   1  'Graphical
         TabIndex        =   162
         Top             =   6390
         Width           =   1155
      End
      Begin VB.CommandButton cmdSaveHisto 
         Caption         =   "&Save Details"
         Enabled         =   0   'False
         Height          =   870
         Index           =   0
         Left            =   -64920
         Picture         =   "frmEditHisto.frx":37A0
         Style           =   1  'Graphical
         TabIndex        =   161
         Top             =   6360
         Width           =   1095
      End
      Begin VB.CommandButton bSaveHold 
         Caption         =   "Save && &Hold"
         Enabled         =   0   'False
         Height          =   870
         Left            =   -64965
         Picture         =   "frmEditHisto.frx":3AAA
         Style           =   1  'Graphical
         TabIndex        =   160
         Top             =   5220
         Width           =   1155
      End
      Begin VB.CommandButton bSave 
         Caption         =   "&Save Details"
         Enabled         =   0   'False
         Height          =   870
         Left            =   -66090
         Picture         =   "frmEditHisto.frx":3DB4
         Style           =   1  'Graphical
         TabIndex        =   159
         Top             =   5220
         Width           =   1095
      End
      Begin VB.ComboBox c 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   13
         Left            =   -71730
         TabIndex        =   170
         Top             =   450
         Width           =   6600
      End
      Begin VB.CommandButton cmdDemoVal 
         Caption         =   "Validate"
         Height          =   870
         Left            =   -67080
         Picture         =   "frmEditHisto.frx":40BE
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   5220
         Width           =   945
      End
      Begin VB.Frame Frame7 
         Caption         =   "Date"
         Height          =   2805
         Left            =   -67125
         TabIndex        =   143
         Top             =   540
         Width           =   3285
         Begin MSComCtl2.DTPicker dtRunDate 
            Height          =   315
            Left            =   945
            TabIndex        =   20
            Top             =   1860
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            _Version        =   393216
            Format          =   272039937
            CurrentDate     =   36942
         End
         Begin MSComCtl2.DTPicker dtSampleDate 
            Height          =   315
            Left            =   945
            TabIndex        =   16
            Top             =   315
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            _Version        =   393216
            Format          =   272039937
            CurrentDate     =   36942
         End
         Begin MSMask.MaskEdBox tSampleTime 
            Height          =   315
            Left            =   2355
            TabIndex        =   17
            ToolTipText     =   "Time of Sample"
            Top             =   315
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSComCtl2.DTPicker dtRecDate 
            Height          =   315
            Left            =   945
            TabIndex        =   18
            Top             =   1080
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            _Version        =   393216
            Format          =   272039937
            CurrentDate     =   36942
         End
         Begin MSMask.MaskEdBox tRecTime 
            Height          =   315
            Left            =   2355
            TabIndex        =   19
            ToolTipText     =   "Time of Sample"
            Top             =   1080
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Sample"
            Height          =   195
            Index           =   2
            Left            =   360
            TabIndex        =   146
            Top             =   360
            Width           =   525
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Run"
            Height          =   195
            Index           =   2
            Left            =   585
            TabIndex        =   145
            Top             =   1920
            Width           =   300
         End
         Begin VB.Image iRunDate 
            Height          =   330
            Index           =   0
            Left            =   990
            Picture         =   "frmEditHisto.frx":43C8
            Stretch         =   -1  'True
            ToolTipText     =   "Previous Day"
            Top             =   2190
            Width           =   480
         End
         Begin VB.Image iRunDate 
            Height          =   330
            Index           =   1
            Left            =   1890
            Picture         =   "frmEditHisto.frx":480A
            Stretch         =   -1  'True
            ToolTipText     =   "Next Day"
            Top             =   2190
            Width           =   480
         End
         Begin VB.Image iSampleDate 
            Height          =   330
            Index           =   0
            Left            =   975
            Picture         =   "frmEditHisto.frx":4C4C
            Stretch         =   -1  'True
            ToolTipText     =   "Previous Day"
            Top             =   630
            Width           =   480
         End
         Begin VB.Image iSampleDate 
            Height          =   330
            Index           =   1
            Left            =   1845
            Picture         =   "frmEditHisto.frx":508E
            Stretch         =   -1  'True
            ToolTipText     =   "Next Day"
            Top             =   630
            Width           =   480
         End
         Begin VB.Image iToday 
            Height          =   330
            Index           =   0
            Left            =   1500
            Picture         =   "frmEditHisto.frx":54D0
            Stretch         =   -1  'True
            ToolTipText     =   "Set to Today"
            Top             =   2190
            Width           =   360
         End
         Begin VB.Image iToday 
            Height          =   330
            Index           =   1
            Left            =   1455
            Picture         =   "frmEditHisto.frx":5912
            Stretch         =   -1  'True
            ToolTipText     =   "Set to Today"
            Top             =   630
            Width           =   360
         End
         Begin VB.Image iToday 
            Height          =   330
            Index           =   2
            Left            =   1455
            Picture         =   "frmEditHisto.frx":5D54
            Stretch         =   -1  'True
            ToolTipText     =   "Set to Today"
            Top             =   1410
            Width           =   360
         End
         Begin VB.Image iRecDate 
            Height          =   330
            Index           =   1
            Left            =   1845
            Picture         =   "frmEditHisto.frx":6196
            Stretch         =   -1  'True
            ToolTipText     =   "Next Day"
            Top             =   1410
            Width           =   480
         End
         Begin VB.Image iRecDate 
            Height          =   330
            Index           =   0
            Left            =   975
            Picture         =   "frmEditHisto.frx":65D8
            Stretch         =   -1  'True
            ToolTipText     =   "Previous Day"
            Top             =   1410
            Width           =   480
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Received"
            Height          =   195
            Index           =   0
            Left            =   195
            TabIndex        =   144
            Top             =   1140
            Width           =   690
         End
      End
      Begin VB.CommandButton cmdCVal 
         Caption         =   "Validate"
         Height          =   870
         Left            =   -66225
         Picture         =   "frmEditHisto.frx":6A1A
         Style           =   1  'Graphical
         TabIndex        =   142
         Top             =   6120
         Width           =   1275
      End
      Begin VB.CommandButton cmdHVal 
         Caption         =   "Validate"
         Height          =   870
         Left            =   -66090
         Picture         =   "frmEditHisto.frx":6D24
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   6390
         Width           =   1140
      End
      Begin VB.TextBox txtHistoComment 
         Height          =   600
         Left            =   -74685
         MultiLine       =   -1  'True
         TabIndex        =   32
         Top             =   5670
         Width           =   11985
      End
      Begin VB.Frame Frame5 
         Height          =   735
         Left            =   -66045
         TabIndex        =   136
         Top             =   3600
         Width           =   1455
         Begin VB.OptionButton cRooH 
            Alignment       =   1  'Right Justify
            Caption         =   "Routine"
            Height          =   195
            Index           =   0
            Left            =   420
            TabIndex        =   138
            Top             =   180
            Width           =   885
         End
         Begin VB.OptionButton cRooH 
            Alignment       =   1  'Right Justify
            Caption         =   "Out of Hours"
            Height          =   195
            Index           =   1
            Left            =   90
            TabIndex        =   137
            Top             =   420
            Width           =   1215
         End
      End
      Begin VB.Frame fraSpec 
         Caption         =   "Specimen F"
         Height          =   2475
         Index           =   5
         Left            =   7590
         TabIndex        =   91
         Top             =   4560
         Visible         =   0   'False
         Width           =   3735
         Begin VB.TextBox txtHist 
            Height          =   465
            Index           =   5
            Left            =   1950
            MaxLength       =   300
            TabIndex        =   110
            Top             =   1680
            Width           =   1635
         End
         Begin VB.ComboBox cmbStain 
            Height          =   315
            Index           =   5
            Left            =   210
            Style           =   2  'Dropdown List
            TabIndex        =   98
            Top             =   1740
            Width           =   1665
         End
         Begin MSFlexGridLib.MSFlexGrid grdSpec 
            Height          =   1335
            Index           =   5
            Left            =   180
            TabIndex        =   92
            Top             =   210
            Width           =   3525
            _ExtentX        =   6218
            _ExtentY        =   2355
            _Version        =   393216
            Cols            =   50
            FormatString    =   "        "
         End
         Begin MSFlexGridLib.MSFlexGrid grdComm 
            Height          =   1155
            Index           =   5
            Left            =   360
            TabIndex        =   123
            Top             =   300
            Width           =   3285
            _ExtentX        =   5794
            _ExtentY        =   2037
            _Version        =   393216
            Cols            =   15
            FormatString    =   "        "
         End
         Begin VB.TextBox tInput 
            Height          =   285
            Index           =   5
            Left            =   720
            TabIndex        =   117
            Top             =   870
            Width           =   615
         End
      End
      Begin VB.Frame fraSpec 
         Caption         =   "Specimen E"
         Height          =   2475
         Index           =   4
         Left            =   7590
         TabIndex        =   89
         Top             =   2160
         Visible         =   0   'False
         Width           =   3735
         Begin VB.TextBox txtHist 
            Height          =   465
            Index           =   4
            Left            =   1950
            MaxLength       =   300
            TabIndex        =   109
            Top             =   1920
            Width           =   1635
         End
         Begin VB.ComboBox cmbStain 
            Height          =   315
            Index           =   4
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   97
            Top             =   1980
            Width           =   1665
         End
         Begin MSFlexGridLib.MSFlexGrid grdSpec 
            Height          =   1605
            Index           =   4
            Left            =   45
            TabIndex        =   90
            Top             =   270
            Width           =   3525
            _ExtentX        =   6218
            _ExtentY        =   2831
            _Version        =   393216
            Cols            =   50
            FormatString    =   "        "
         End
         Begin MSFlexGridLib.MSFlexGrid grdComm 
            Height          =   1155
            Index           =   4
            Left            =   45
            TabIndex        =   122
            Top             =   315
            Width           =   3285
            _ExtentX        =   5794
            _ExtentY        =   2037
            _Version        =   393216
            Cols            =   15
            FormatString    =   "        "
         End
         Begin VB.TextBox tInput 
            Height          =   285
            Index           =   4
            Left            =   690
            TabIndex        =   116
            Top             =   870
            Width           =   615
         End
      End
      Begin VB.Frame fraSpec 
         Caption         =   "Specimen D"
         Height          =   2475
         Index           =   3
         Left            =   3825
         TabIndex        =   87
         Top             =   4590
         Visible         =   0   'False
         Width           =   3735
         Begin VB.TextBox txtHist 
            Height          =   465
            Index           =   3
            Left            =   1920
            MaxLength       =   300
            TabIndex        =   108
            Top             =   1650
            Width           =   1635
         End
         Begin VB.ComboBox cmbStain 
            Height          =   315
            Index           =   3
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   96
            Top             =   1740
            Width           =   1665
         End
         Begin MSFlexGridLib.MSFlexGrid grdSpec 
            Height          =   1365
            Index           =   3
            Left            =   120
            TabIndex        =   88
            Top             =   210
            Width           =   3525
            _ExtentX        =   6218
            _ExtentY        =   2408
            _Version        =   393216
            Cols            =   50
            FormatString    =   "        "
         End
         Begin VB.TextBox tInput 
            Height          =   285
            Index           =   3
            Left            =   990
            TabIndex        =   115
            Top             =   960
            Width           =   615
         End
         Begin MSFlexGridLib.MSFlexGrid grdComm 
            Height          =   1155
            Index           =   3
            Left            =   330
            TabIndex        =   121
            Top             =   330
            Width           =   3285
            _ExtentX        =   5794
            _ExtentY        =   2037
            _Version        =   393216
            Cols            =   15
            FormatString    =   "        "
         End
      End
      Begin VB.Frame fraSpec 
         Caption         =   "Specimen C"
         Height          =   2475
         Index           =   2
         Left            =   3870
         TabIndex        =   85
         Top             =   2160
         Visible         =   0   'False
         Width           =   3735
         Begin VB.TextBox txtHist 
            Height          =   465
            Index           =   2
            Left            =   1890
            MaxLength       =   300
            TabIndex        =   107
            Top             =   1920
            Width           =   1635
         End
         Begin VB.ComboBox cmbStain 
            Height          =   315
            Index           =   2
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   95
            Top             =   1980
            Width           =   1665
         End
         Begin MSFlexGridLib.MSFlexGrid grdSpec 
            Height          =   1575
            Index           =   2
            Left            =   90
            TabIndex        =   86
            Top             =   270
            Width           =   3525
            _ExtentX        =   6218
            _ExtentY        =   2778
            _Version        =   393216
            Cols            =   50
            FormatString    =   "        "
         End
         Begin VB.TextBox tInput 
            Height          =   285
            Index           =   2
            Left            =   660
            TabIndex        =   114
            Top             =   930
            Width           =   615
         End
         Begin MSFlexGridLib.MSFlexGrid grdComm 
            Height          =   1155
            Index           =   2
            Left            =   240
            TabIndex        =   120
            Top             =   540
            Width           =   3285
            _ExtentX        =   5794
            _ExtentY        =   2037
            _Version        =   393216
            Cols            =   15
            FormatString    =   "        "
         End
      End
      Begin VB.Frame fraSpec 
         Caption         =   "Specimen B"
         Height          =   2475
         Index           =   1
         Left            =   150
         TabIndex        =   83
         Top             =   4560
         Visible         =   0   'False
         Width           =   3735
         Begin VB.TextBox txtHist 
            Height          =   465
            Index           =   1
            Left            =   1980
            MaxLength       =   300
            TabIndex        =   106
            Top             =   1620
            Width           =   1635
         End
         Begin VB.ComboBox cmbStain 
            Height          =   315
            Index           =   1
            Left            =   150
            Style           =   2  'Dropdown List
            TabIndex        =   94
            Top             =   1740
            Width           =   1665
         End
         Begin MSFlexGridLib.MSFlexGrid grdSpec 
            Height          =   1365
            Index           =   1
            Left            =   120
            TabIndex        =   84
            Top             =   210
            Width           =   3525
            _ExtentX        =   6218
            _ExtentY        =   2408
            _Version        =   393216
            Cols            =   50
            FormatString    =   "        "
         End
         Begin VB.TextBox tInput 
            Height          =   285
            Index           =   1
            Left            =   2010
            TabIndex        =   113
            Top             =   900
            Width           =   615
         End
         Begin MSFlexGridLib.MSFlexGrid grdComm 
            Height          =   1155
            Index           =   1
            Left            =   330
            TabIndex        =   119
            Top             =   420
            Width           =   3285
            _ExtentX        =   5794
            _ExtentY        =   2037
            _Version        =   393216
            Cols            =   15
            FormatString    =   "        "
         End
      End
      Begin VB.Frame fraSpec 
         Caption         =   "Specimen A"
         Height          =   2475
         Index           =   0
         Left            =   180
         TabIndex        =   81
         Top             =   2160
         Visible         =   0   'False
         Width           =   3735
         Begin VB.TextBox txtHist 
            Height          =   465
            Index           =   0
            Left            =   1980
            MaxLength       =   300
            TabIndex        =   105
            Top             =   1920
            Width           =   1635
         End
         Begin VB.ComboBox cmbStain 
            Height          =   315
            Index           =   0
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   93
            Top             =   1980
            Width           =   1665
         End
         Begin MSFlexGridLib.MSFlexGrid grdSpec 
            Height          =   1575
            Index           =   0
            Left            =   150
            TabIndex        =   82
            Top             =   270
            Width           =   3525
            _ExtentX        =   6218
            _ExtentY        =   2778
            _Version        =   393216
            Cols            =   50
            FormatString    =   "        "
         End
         Begin VB.TextBox tInput 
            Height          =   285
            Index           =   0
            Left            =   960
            TabIndex        =   112
            Top             =   1110
            Width           =   615
         End
         Begin MSFlexGridLib.MSFlexGrid grdComm 
            Height          =   1155
            Index           =   0
            Left            =   240
            TabIndex        =   118
            Top             =   480
            Width           =   3285
            _ExtentX        =   5794
            _ExtentY        =   2037
            _Version        =   393216
            Cols            =   15
            FormatString    =   "        "
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Nature of Specimen"
         Height          =   1635
         Left            =   150
         TabIndex        =   66
         Top             =   630
         Width           =   12030
         Begin VB.ComboBox c 
            Height          =   315
            Index           =   5
            Left            =   6375
            Sorted          =   -1  'True
            TabIndex        =   30
            Top             =   1080
            Visible         =   0   'False
            Width           =   3975
         End
         Begin VB.ComboBox c 
            Height          =   315
            Index           =   4
            Left            =   6375
            Sorted          =   -1  'True
            TabIndex        =   29
            Top             =   720
            Visible         =   0   'False
            Width           =   3975
         End
         Begin ComCtl2.UpDown upBlck 
            Height          =   315
            Index           =   0
            Left            =   4815
            TabIndex        =   74
            Top             =   390
            Visible         =   0   'False
            Width           =   240
            _ExtentX        =   450
            _ExtentY        =   556
            _Version        =   327681
            Enabled         =   -1  'True
         End
         Begin VB.ComboBox c 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   0
            Left            =   420
            Sorted          =   -1  'True
            TabIndex        =   25
            Top             =   330
            Visible         =   0   'False
            Width           =   4005
         End
         Begin VB.ComboBox c 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   1
            Left            =   420
            Sorted          =   -1  'True
            TabIndex        =   26
            Top             =   690
            Visible         =   0   'False
            Width           =   4005
         End
         Begin VB.ComboBox c 
            Height          =   315
            Index           =   2
            Left            =   420
            Sorted          =   -1  'True
            TabIndex        =   27
            Top             =   1050
            Visible         =   0   'False
            Width           =   4005
         End
         Begin VB.ComboBox c 
            Height          =   315
            Index           =   3
            Left            =   6375
            Sorted          =   -1  'True
            TabIndex        =   28
            Top             =   360
            Visible         =   0   'False
            Width           =   3975
         End
         Begin ComCtl2.UpDown upBlck 
            Height          =   315
            Index           =   1
            Left            =   4815
            TabIndex        =   76
            Top             =   750
            Visible         =   0   'False
            Width           =   240
            _ExtentX        =   450
            _ExtentY        =   556
            _Version        =   327681
            Enabled         =   -1  'True
         End
         Begin ComCtl2.UpDown upBlck 
            Height          =   315
            Index           =   2
            Left            =   4815
            TabIndex        =   78
            Top             =   1080
            Visible         =   0   'False
            Width           =   240
            _ExtentX        =   450
            _ExtentY        =   556
            _Version        =   327681
            Enabled         =   -1  'True
         End
         Begin ComCtl2.UpDown upBlck 
            Height          =   315
            Index           =   3
            Left            =   10785
            TabIndex        =   80
            Top             =   330
            Visible         =   0   'False
            Width           =   240
            _ExtentX        =   450
            _ExtentY        =   556
            _Version        =   327681
            Enabled         =   -1  'True
         End
         Begin ComCtl2.UpDown upBlck 
            Height          =   315
            Index           =   4
            Left            =   10785
            TabIndex        =   99
            Top             =   720
            Visible         =   0   'False
            Width           =   240
            _ExtentX        =   450
            _ExtentY        =   556
            _Version        =   327681
            Enabled         =   -1  'True
         End
         Begin ComCtl2.UpDown upBlck 
            Height          =   315
            Index           =   5
            Left            =   10785
            TabIndex        =   101
            Top             =   1080
            Visible         =   0   'False
            Width           =   240
            _ExtentX        =   450
            _ExtentY        =   556
            _Version        =   327681
            Enabled         =   -1  'True
         End
         Begin ComCtl2.UpDown upFS 
            Height          =   315
            Index           =   0
            Left            =   5490
            TabIndex        =   178
            Top             =   405
            Visible         =   0   'False
            Width           =   240
            _ExtentX        =   450
            _ExtentY        =   556
            _Version        =   327681
            Enabled         =   -1  'True
         End
         Begin ComCtl2.UpDown upFS 
            Height          =   315
            Index           =   1
            Left            =   5490
            TabIndex        =   179
            Top             =   765
            Visible         =   0   'False
            Width           =   240
            _ExtentX        =   450
            _ExtentY        =   556
            _Version        =   327681
            Enabled         =   -1  'True
         End
         Begin ComCtl2.UpDown upFS 
            Height          =   315
            Index           =   2
            Left            =   5490
            TabIndex        =   180
            Top             =   1125
            Visible         =   0   'False
            Width           =   240
            _ExtentX        =   450
            _ExtentY        =   556
            _Version        =   327681
            Enabled         =   -1  'True
         End
         Begin ComCtl2.UpDown upFS 
            Height          =   315
            Index           =   3
            Left            =   11475
            TabIndex        =   181
            Top             =   315
            Visible         =   0   'False
            Width           =   240
            _ExtentX        =   450
            _ExtentY        =   556
            _Version        =   327681
            Enabled         =   -1  'True
         End
         Begin ComCtl2.UpDown upFS 
            Height          =   315
            Index           =   4
            Left            =   11475
            TabIndex        =   182
            Top             =   720
            Visible         =   0   'False
            Width           =   240
            _ExtentX        =   450
            _ExtentY        =   556
            _Version        =   327681
            Enabled         =   -1  'True
         End
         Begin ComCtl2.UpDown upFS 
            Height          =   315
            Index           =   5
            Left            =   11475
            TabIndex        =   183
            Top             =   1080
            Visible         =   0   'False
            Width           =   240
            _ExtentX        =   450
            _ExtentY        =   556
            _Version        =   327681
            Enabled         =   -1  'True
         End
         Begin VB.Label Label8 
            Caption         =   "FS"
            Height          =   195
            Index           =   1
            Left            =   11115
            TabIndex        =   190
            Top             =   135
            Width           =   285
         End
         Begin VB.Label lblFS 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Index           =   5
            Left            =   11070
            TabIndex        =   189
            Top             =   1080
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label lblFS 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Index           =   4
            Left            =   11070
            TabIndex        =   188
            Top             =   720
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label lblFS 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Index           =   3
            Left            =   11070
            TabIndex        =   187
            Top             =   360
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label lblFS 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Index           =   2
            Left            =   5085
            TabIndex        =   186
            Top             =   1080
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label lblFS 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Index           =   1
            Left            =   5085
            TabIndex        =   185
            Top             =   765
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label lblFS 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Index           =   0
            Left            =   5085
            TabIndex        =   184
            Top             =   405
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label Label8 
            Caption         =   "FS"
            Height          =   195
            Index           =   0
            Left            =   5130
            TabIndex        =   177
            Top             =   180
            Width           =   285
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "[F]"
            Height          =   195
            Index           =   5
            Left            =   6135
            TabIndex        =   104
            Top             =   1140
            Visible         =   0   'False
            Width           =   180
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "[E]"
            Height          =   195
            Index           =   4
            Left            =   6135
            TabIndex        =   103
            Top             =   780
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Label lblBlock 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Index           =   5
            Left            =   10455
            TabIndex        =   102
            Top             =   1080
            Width           =   315
         End
         Begin VB.Label lblBlock 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Index           =   4
            Left            =   10455
            TabIndex        =   100
            Top             =   720
            Width           =   315
         End
         Begin VB.Label lblBlock 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Index           =   3
            Left            =   10455
            TabIndex        =   79
            Top             =   330
            Width           =   315
         End
         Begin VB.Label lblBlock 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Index           =   2
            Left            =   4485
            TabIndex        =   77
            Top             =   1080
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.Label lblBlock 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Index           =   1
            Left            =   4485
            TabIndex        =   75
            Top             =   750
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.Label lblBlock 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Index           =   0
            Left            =   4485
            TabIndex        =   73
            Top             =   390
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.Label Label13 
            Caption         =   "Blocks"
            Height          =   255
            Index           =   3
            Left            =   10365
            TabIndex        =   72
            Top             =   135
            Width           =   585
         End
         Begin VB.Label Label13 
            Caption         =   "Blocks"
            Height          =   255
            Index           =   0
            Left            =   4455
            TabIndex        =   71
            Top             =   180
            Width           =   585
         End
         Begin VB.Label Label5 
            Caption         =   " [A]"
            Height          =   225
            Index           =   0
            Left            =   120
            TabIndex        =   70
            Top             =   390
            Visible         =   0   'False
            Width           =   360
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "[B]"
            Height          =   195
            Index           =   1
            Left            =   180
            TabIndex        =   69
            Top             =   750
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "[C]"
            Height          =   195
            Index           =   2
            Left            =   180
            TabIndex        =   68
            Top             =   1110
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "[D]"
            Height          =   195
            Index           =   3
            Left            =   6135
            TabIndex        =   67
            Top             =   420
            Visible         =   0   'False
            Width           =   210
         End
      End
      Begin VB.TextBox t 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4500
         Left            =   -74685
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   31
         Top             =   855
         Width           =   12000
      End
      Begin VB.Frame Frame3 
         Height          =   855
         Left            =   -73515
         TabIndex        =   60
         Top             =   720
         Width           =   10935
         Begin VB.ComboBox c 
            Height          =   315
            Index           =   9
            Left            =   6390
            TabIndex        =   174
            Top             =   450
            Width           =   4215
         End
         Begin VB.ComboBox c 
            Height          =   315
            Index           =   8
            Left            =   1800
            TabIndex        =   173
            Top             =   450
            Width           =   4215
         End
         Begin VB.ComboBox c 
            Height          =   315
            Index           =   6
            Left            =   1800
            TabIndex        =   171
            Top             =   135
            Width           =   4215
         End
         Begin VB.ComboBox c 
            Height          =   315
            Index           =   7
            Left            =   6390
            TabIndex        =   172
            Top             =   135
            Width           =   4215
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Nature of Specimen [A]"
            Height          =   225
            Left            =   60
            TabIndex        =   64
            Top             =   180
            Width           =   1710
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "[B]"
            Height          =   195
            Left            =   6120
            TabIndex        =   63
            Top             =   180
            Width           =   195
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "[C]"
            Height          =   195
            Left            =   1575
            TabIndex        =   62
            Top             =   495
            Width           =   195
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "[D]"
            Height          =   195
            Left            =   6120
            TabIndex        =   61
            Top             =   495
            Width           =   210
         End
      End
      Begin VB.TextBox txtCyto 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3615
         Left            =   -73515
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   176
         Top             =   1620
         Width           =   10905
      End
      Begin VB.Frame Frame4 
         Height          =   6225
         Left            =   -74595
         TabIndex        =   38
         Top             =   540
         Width           =   7335
         Begin VB.CommandButton cmdCopyTo 
            Caption         =   "++ cc ++"
            Height          =   960
            Left            =   6840
            TabIndex        =   167
            Top             =   2745
            Width           =   375
         End
         Begin VB.ComboBox cmbHospital 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   900
            TabIndex        =   12
            Top             =   2655
            Width           =   5850
         End
         Begin VB.ComboBox cmbGP 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   900
            TabIndex        =   15
            Top             =   3870
            Width           =   5850
         End
         Begin VB.ComboBox cmbClinician 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   900
            TabIndex        =   14
            Top             =   3465
            Width           =   5850
         End
         Begin VB.TextBox taddress 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   1
            Left            =   900
            MaxLength       =   30
            TabIndex        =   11
            Top             =   2205
            Width           =   5850
         End
         Begin VB.TextBox taddress 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   0
            Left            =   900
            MaxLength       =   30
            TabIndex        =   10
            Top             =   1830
            Width           =   5850
         End
         Begin VB.ComboBox cmbWard 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   900
            TabIndex        =   13
            Top             =   3060
            Width           =   5850
         End
         Begin VB.ComboBox cClDetails 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   900
            Sorted          =   -1  'True
            TabIndex        =   23
            Top             =   5310
            Width           =   5850
         End
         Begin VB.TextBox txtDemographicComment 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   990
            Left            =   900
            MaxLength       =   160
            MultiLine       =   -1  'True
            TabIndex        =   22
            Top             =   4275
            Width           =   5850
         End
         Begin VB.Label lblNOPAS 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Index           =   1
            Left            =   4245
            TabIndex        =   147
            Top             =   270
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Hospital"
            Height          =   195
            Left            =   270
            TabIndex        =   124
            Top             =   2745
            Width           =   570
         End
         Begin VB.Label Label32 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "GP"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   615
            TabIndex        =   54
            Top             =   3960
            Width           =   225
         End
         Begin VB.Label Label31 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Clinician"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   255
            TabIndex        =   53
            Top             =   3555
            Width           =   585
         End
         Begin VB.Label Label30 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Comments"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   105
            TabIndex        =   52
            Top             =   4320
            Width           =   735
         End
         Begin VB.Label Label25 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Address"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   315
            TabIndex        =   51
            Top             =   1890
            Width           =   570
         End
         Begin VB.Label Label24 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Ward"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   450
            TabIndex        =   50
            Top             =   3150
            Width           =   390
         End
         Begin VB.Label Label23 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Sex"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   4110
            TabIndex        =   49
            Top             =   1200
            Width           =   270
         End
         Begin VB.Label Label22 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Age"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   2940
            TabIndex        =   48
            Top             =   1200
            Width           =   285
         End
         Begin VB.Label Label21 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "D.o.B"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   480
            TabIndex        =   47
            Top             =   1185
            Width           =   405
         End
         Begin VB.Label Label17 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Name"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   450
            TabIndex        =   46
            Top             =   810
            Width           =   420
         End
         Begin VB.Label Label16 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Chart #"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   405
            TabIndex        =   45
            Top             =   330
            Width           =   525
         End
         Begin VB.Label Label36 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Cl Details"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   180
            TabIndex        =   44
            Top             =   5400
            Width           =   660
         End
         Begin VB.Label lChart 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1065
            TabIndex        =   43
            Top             =   300
            Width           =   1485
         End
         Begin VB.Label lName 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1065
            TabIndex        =   42
            Top             =   690
            Width           =   5385
         End
         Begin VB.Label lDoB 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1020
            TabIndex        =   41
            Top             =   1185
            Width           =   1515
         End
         Begin VB.Label lAge 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   3360
            TabIndex        =   40
            Top             =   1200
            Width           =   585
         End
         Begin VB.Label lSex 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   4440
            TabIndex        =   39
            Top             =   1200
            Width           =   705
         End
         Begin VB.Label Label35 
            Caption         =   "Nopas"
            Height          =   285
            Left            =   3735
            TabIndex        =   148
            Top             =   330
            Visible         =   0   'False
            Width           =   615
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   405
         Left            =   -74310
         TabIndex        =   58
         Top             =   4380
         Visible         =   0   'False
         Width           =   555
         _Version        =   65536
         _ExtentX        =   979
         _ExtentY        =   714
         _StockProps     =   15
         Caption         =   "2"
         ForeColor       =   -2147483640
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   4
         Begin VB.Image iCPrint 
            Height          =   330
            Index           =   1
            Left            =   45
            Picture         =   "frmEditHisto.frx":702E
            ToolTipText     =   "Print 2 Copies"
            Top             =   30
            Visible         =   0   'False
            Width           =   360
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   405
         Left            =   -74850
         TabIndex        =   59
         Top             =   4380
         Visible         =   0   'False
         Width           =   525
         _Version        =   65536
         _ExtentX        =   926
         _ExtentY        =   714
         _StockProps     =   15
         Caption         =   "1"
         ForeColor       =   -2147483640
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Begin VB.Image iCPrint 
            Height          =   330
            Index           =   0
            Left            =   180
            Picture         =   "frmEditHisto.frx":71B8
            ToolTipText     =   "Print"
            Top             =   30
            Visible         =   0   'False
            Width           =   360
         End
      End
      Begin VB.Label lblVal 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -72795
         TabIndex        =   198
         Top             =   6660
         Width           =   4830
      End
      Begin VB.Label lblD 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   -65415
         TabIndex        =   197
         Top             =   495
         Width           =   2895
      End
      Begin VB.Label lblC 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   -68430
         TabIndex        =   196
         Top             =   495
         Width           =   2985
      End
      Begin VB.Label lblB 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   -71625
         TabIndex        =   195
         Top             =   495
         Width           =   3165
      End
      Begin VB.Label lblA 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   -74910
         TabIndex        =   194
         Top             =   495
         Width           =   3210
      End
      Begin VB.Label Label9 
         Caption         =   "Comment"
         Height          =   240
         Left            =   -74640
         TabIndex        =   193
         Top             =   5265
         Width           =   1140
      End
      Begin VB.Label Label18 
         Caption         =   "Comment"
         Height          =   240
         Left            =   -74685
         TabIndex        =   140
         Top             =   5355
         Width           =   1140
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   4920
         TabIndex        =   111
         Top             =   360
         Width           =   1875
      End
      Begin VB.Image iCLocked 
         Height          =   480
         Left            =   -74595
         Picture         =   "frmEditHisto.frx":7342
         ToolTipText     =   "Use Key to Unlock"
         Top             =   3015
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image iCUnlocked 
         Height          =   480
         Left            =   -74520
         Picture         =   "frmEditHisto.frx":7784
         ToolTipText     =   "Validate & Lock"
         Top             =   3090
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image iCKey 
         Height          =   480
         Left            =   -74070
         Picture         =   "frmEditHisto.frx":7BC6
         ToolTipText     =   "Unlock"
         Top             =   3090
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image iCDelete 
         Height          =   330
         Left            =   -74520
         Picture         =   "frmEditHisto.frx":8008
         ToolTipText     =   "Delete Selection"
         Top             =   3960
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Image iCDate 
         Height          =   330
         Index           =   0
         Left            =   -74520
         Picture         =   "frmEditHisto.frx":8192
         ToolTipText     =   "Insert The Current Date"
         Top             =   3570
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Comment"
         Height          =   195
         Left            =   -72420
         TabIndex        =   65
         Top             =   510
         Width           =   660
      End
   End
   Begin VB.CommandButton bViewBB 
      Caption         =   "Transfusion Details"
      Height          =   885
      Left            =   6840
      Picture         =   "frmEditHisto.frx":831C
      Style           =   1  'Graphical
      TabIndex        =   139
      Top             =   5535
      Width           =   1320
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   158
      Top             =   9780
      Width           =   14130
      _ExtentX        =   24924
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4480
            MinWidth        =   4480
            Object.ToolTipText     =   "User Name"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "11/22/2023"
            Object.ToolTipText     =   "Todays Date"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   5292
            MinWidth        =   5292
            Object.ToolTipText     =   "Demographic Check"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Rundate"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6068
            MinWidth        =   6068
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Bevel           =   0
            Object.Width           =   12347
            MinWidth        =   12347
            Text            =   "Custom Software Ltd"
            TextSave        =   "Custom Software Ltd"
         EndProperty
      EndProperty
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
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "No Of Copies"
      Height          =   195
      Left            =   13020
      TabIndex        =   201
      Top             =   3360
      Width           =   945
   End
End
Attribute VB_Name = "frmEditHisto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mNewRecord As Boolean

Private PreviousCyto As Boolean
Private PreviousHisto As Boolean


Private HistoLoaded As Boolean
Private CytoLoaded As Boolean

Private Activated As Boolean

Private HDate As String

Private pPrintToPrinter As String

Private Sub Clear_HistoWork()

          Dim n As Long
          Dim f As Long
          Dim sql As String
          Dim tb As New Recordset

10        On Error GoTo Clear_HistoWork_Error

20        StatusBar1.Panels(5).Text = ""
30        txtHist(0) = ""
40        lblStatus = ""
50        HDate = ""
60        For n = 0 To 5
70            c(n).ListIndex = 0
80            With grdSpec(n)
90                .Cols = 1
100               .Rows = 2
110               .AddItem ""
120               .RemoveItem 1
130           End With
140           fraSpec(n).Visible = False
150           With grdComm(n)
160               .Rows = 2
170               .AddItem ""
180               .RemoveItem 1
190           End With
200           lblBlock(n).Caption = ""
210           lblFS(n).Caption = ""
220       Next

230       For n = 0 To 5
240           fraSpec(n).Visible = False
250           cmbStain(n).Clear
260       Next

270       For f = 0 To 5
280           grdSpec(f).ColWidth(0) = 1000
290           grdSpec(f).TextMatrix(1, 0) = "Pieces"
300           grdSpec(f).Cols = 1
310       Next

320       For n = 0 To 5
330           fraSpec(n).Caption = "Specimen " & chr$(Asc("A") + n)
340       Next

350       sql = "SELECT * from lists WHERE listtype = 'SH' "
360       Set tb = New Recordset
370       RecOpenServer 0, tb, sql

380       Do While Not tb.EOF
390           For n = 0 To 5
400               cmbStain(n).AddItem Trim(tb!Text)
410           Next
420           tb.MoveNext
430       Loop



440       Exit Sub

Clear_HistoWork_Error:

          Dim strES As String
          Dim intEL As Integer

450       Screen.MousePointer = 0

460       intEL = Erl
470       strES = Err.Description
480       LogError "frmEditHisto", "Clear_HistoWork", intEL, strES, sql


End Sub
Private Sub bcancel_Click()

10        On Error GoTo bCancel_Click_Error

20        pBar = 0

30        Unload Me

40        Exit Sub

bCancel_Click_Error:

          Dim strES As String
          Dim intEL As Integer

50        Screen.MousePointer = 0

60        intEL = Erl
70        strES = Err.Description
80        LogError "frmEditHisto", "bcancel_Click", intEL, strES


End Sub


Private Sub bDoB_Click()

10        On Error GoTo bDoB_Click_Error

20        pBar = 0

30        With frmPatHistoryNew
40            .oHD(1) = True
50            .oFor(2) = True
60            .txtName = txtDoB
70            .FromEdit = True
80            .EditScreen = Me
90            .bsearch = True
100           If Not .NoPreviousDetails Then
110               .Show 1
120           Else
130               FlashNoPrevious
140           End If
150       End With

160       Exit Sub

bDoB_Click_Error:

          Dim strES As String
          Dim intEL As Integer

170       Screen.MousePointer = 0

180       intEL = Erl
190       strES = Err.Description
200       LogError "frmEditHisto", "bDoB_Click", intEL, strES


End Sub

Private Sub bFAX_Click(Index As Integer)

10        On Error GoTo bFAX_Click_Error

20        pBar = 0

30        Exit Sub

bFAX_Click_Error:

          Dim strES As String
          Dim intEL As Integer

40        Screen.MousePointer = 0

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmEditHisto", "bFAX_Click", intEL, strES


End Sub


Private Sub bHistory_Click()

10        On Error GoTo bHistory_Click_Error

20        pBar = 0

30        iMsg "Under Development"




40        Exit Sub

bHistory_Click_Error:

          Dim strES As String
          Dim intEL As Integer

50        Screen.MousePointer = 0

60        intEL = Erl
70        strES = Err.Description
80        LogError "frmEditHisto", "bHistory_Click", intEL, strES


End Sub




Private Sub bprint_Click()

          Dim tb As New Recordset
          Dim sql As String
          Dim Yadd As Long
          Dim Dept As String
          Dim pTime As String
          Dim tx As Single

10        On Error GoTo bprint_Click_Error

20        pBar = 0

30        Yadd = Val(Swap_Year(txtYear)) * 1000

40        If Trim$(txtSex) = "" Then
50            If iMsg("Sex not entered." & vbCrLf & "Do you want to enter sex now?", vbQuestion + vbYesNo) = vbYes Then
60                Exit Sub
70            End If
80        End If

90        If Trim$(txtSampleID) = "" Then
100           iMsg "Must have Lab Number.", vbCritical
110           Exit Sub
120       End If

130       If Len(cmbWard) = 0 Then
140           iMsg "Must have Ward entry.", vbCritical
150           Exit Sub
160       End If

170       If Trim$(cmbWard) = "GP" Then
180           If Len(cmbGP) = 0 Then
190               iMsg "Must have Ward or GP entry.", vbCritical
200               Exit Sub
210           End If
220       End If

230       If SaveDemographics_Click = False Then Exit Sub

240       If lblDisp = "C" Then
250           SaveCytology
260           Dept = "Y"
270           Yadd = SysOptCytoOffset(0) + (Val(Swap_Year(txtYear)) * 1000)
280       Else
290           SaveHistoWork
300           SaveHistology
310           Dept = "P"
320           Yadd = SysOptHistoOffset(0) + (Val(Swap_Year(txtYear)) * 1000)
330       End If


340       LogTimeOfPrinting txtSampleID, Dept
350       pTime = Format(Now, "dd/MMM/yyyy hh:mm:ss")
360       sql = "SELECT * FROM PrintPending WHERE " & _
                "Department = '" & Dept & "' " & _
                "AND SampleID = '" & txtSampleID + Yadd & "' " & _
                "AND hYear = '" & txtYear & "' " & _
                "AND pTime = '" & pTime & "'"
370       Set tb = New Recordset
380       RecOpenClient 0, tb, sql
390       If tb.EOF Then
400           tb.AddNew
410       End If
420       tb!SampleID = txtSampleID + Yadd
430       tb!Ward = cmbWard
440       tb!Clinician = cmbClinician
450       tb!GP = cmbGP
460       tb!Department = Dept
470       tb!Initiator = UserName
480       tb!UsePrinter = pPrintToPrinter
490       tb!Hyear = txtYear
500       tb!pTime = pTime
510       tb!NoOfCopies = Val(txtNoCopies)
520       tb.Update

530       txtSampleID = txtSampleID + 1

540       LoadAllDetails

550       txtSampleID.SetFocus

560       tx = Timer: Do While Timer - tx < 1: Loop

570       Exit Sub

bprint_Click_Error:

          Dim strES As String
          Dim intEL As Integer

580       Screen.MousePointer = 0

590       intEL = Erl
600       strES = Err.Description
610       LogError "frmEditHisto", "bPrint_Click", intEL, strES

End Sub




Private Sub bPrintHold_Click()
          Dim Dept As String
          Dim Yadd As Long
          Dim sql As String
          Dim tb As Recordset
          Dim pTime As String
          Dim tx As Single


10        On Error GoTo bPrintHold_Click_Error

20        pBar = 0

30        If Trim$(txtSex) = "" Then
40            If iMsg("Sex not entered." & vbCrLf & "Do you want to enter sex now?", vbQuestion + vbYesNo) = vbYes Then
50                Exit Sub
60            End If
70        End If

80        If Trim$(txtSampleID) = "" Then
90            iMsg "Must have Lab Number.", vbCritical
100           Exit Sub
110       End If

120       If Trim$(cmbWard) = "" Then
130           iMsg "Must have Ward entry.", vbCritical
140           Exit Sub
150       End If

160       If Trim$(cmbWard) = "GP" Then
170           If Trim$(cmbGP) = "" Then
180               iMsg "Must have Ward or GP entry.", vbCritical
190               Exit Sub
200           End If
210       End If

220       If SaveDemographics_Click = False Then Exit Sub

230       If lblDisp = "C" Then
240           SaveCytology
250           Dept = "Y"
260           Yadd = SysOptCytoOffset(0) + (Val(Swap_Year(txtYear)) * 1000)
270       Else
280           SaveHistoWork
290           SaveHistology
300           Dept = "P"
310           Yadd = SysOptHistoOffset(0) + (Val(Swap_Year(txtYear)) * 1000)
320       End If



330       LogTimeOfPrinting txtSampleID + Yadd, Dept
340       pTime = Format(Now, "dd/MMM/yyyy hh:mm:ss")
350       sql = "SELECT * FROM PrintPending WHERE " & _
                "Department = '" & Dept & "' " & _
                "AND SampleID = '" & txtSampleID + Yadd & "' " & _
                "AND hYear = '" & txtYear & "' " & _
                "AND pTime = '" & pTime & "'"
360       Set tb = New Recordset
370       RecOpenClient 0, tb, sql
380       If tb.EOF Then
390           tb.AddNew
400       End If
410       tb!SampleID = txtSampleID + Yadd
420       tb!Department = Dept
430       tb!Initiator = UserName
440       tb!UsePrinter = pPrintToPrinter
450       tb!Hyear = txtYear
460       tb!Ward = cmbWard
470       tb!Clinician = cmbClinician
480       tb!GP = cmbGP
490       tb!pTime = pTime
500       tb!NoOfCopies = Val(txtNoCopies)
510       tb.Update

520       txtSampleID.SetFocus

530       tx = Timer: Do While Timer - tx < 1: Loop

540       Exit Sub

bPrintHold_Click_Error:

          Dim strES As String
          Dim intEL As Integer

550       Screen.MousePointer = 0

560       intEL = Erl
570       strES = Err.Description
580       LogError "frmEditHisto", "bPrintHold_Click", intEL, strES


End Sub




Private Sub SaveHistoWork()
          Dim tb As New Recordset
          Dim sql As String
          Dim n As Long
          Dim X As Long
          Dim b As Long
          Dim Yadd As Long

10        On Error GoTo SaveHistoWork_Error

20        Yadd = Val(Swap_Year(txtYear)) * 1000

30        If HDate = "" Then HDate = Format(Now, "dd/MMM/yyyy")

40        sql = "DELETE from HistoSpecimen WHERE sampleid = '" & txtSampleID + SysOptHistoOffset(0) + Yadd & "' and hyear = '" & txtYear & "'"
50        Cnxn(0).Execute sql



60        sql = "SELECT * from HistoSpecimen WHERE sampleid = '" & txtSampleID + SysOptHistoOffset(0) + Yadd & "' and hyear = '" & txtYear & "'"
70        Set tb = New Recordset
80        RecOpenServer 0, tb, sql

90        For n = 0 To 5
100           If Trim(c(n)) <> "" Then
110               sql = "Insert Into HistoSpecimen (SampleID, Specimen, type, blocks, Rundate, Remark, HYear, Status, fs) " & _
                        "Values (@SampleID0, '@Specimen1', '@type2', '@blocks3', '@Rundate4', '@Remark5', '@HYear6', '@Status7', '@fs8') "

120               sql = Replace(sql, "@SampleID0", txtSampleID + SysOptHistoOffset(0) + Yadd)
130               sql = Replace(sql, "@Specimen1", n)
140               sql = Replace(sql, "@type2", AddTicks(c(n)))
150               sql = Replace(sql, "@blocks3", lblBlock(n).Caption)
160               sql = Replace(sql, "@Rundate4", Format(HDate, "dd/MMM/yyyy"))
170               sql = Replace(sql, "@Remark5", AddTicks(txtHist(n)))
180               sql = Replace(sql, "@HYear6", txtYear)
190               sql = Replace(sql, "@Status7", lblStatus & "")
200               sql = Replace(sql, "@fs8", lblFS(n).Caption)

210               Cnxn(0).Execute sql


220           End If
230       Next

240       RecClose tb

250       sql = "DELETE from HistoBlock WHERE sampleid = " & txtSampleID + SysOptHistoOffset(0) + Yadd & " and hyear = '" & txtYear & "'"
260       Cnxn(0).Execute sql

270       sql = "SELECT * from HistoBlock WHERE sampleid = '" & txtSampleID + SysOptHistoOffset(0) + Yadd & "' and hyear = '" & txtYear & "'"
280       Set tb = New Recordset
290       RecOpenServer 0, tb, sql

300       For n = 0 To 5
310           If c(n) <> "" Then
320               For X = 1 To Val(grdSpec(n).Cols - 1)
330                   sql = "Insert Into HistoBlock (SampleID, Specimen, Block, Pieces, Type, HYear, PiComm) " & _
                            "Values (@SampleID0, '@Specimen1', @Block2, @Pieces3, '@Type4', '@HYear5', '@PiComm6') "

340                   sql = Replace(sql, "@SampleID0", txtSampleID + SysOptHistoOffset(0) + Yadd)
350                   sql = Replace(sql, "@Specimen1", n)
360                   sql = Replace(sql, "@Block2", X)
370                   If Trim(grdSpec(n).TextMatrix(1, X)) <> "" Then
380                       sql = Replace(sql, "@Pieces3", grdSpec(n).TextMatrix(1, X))
390                   Else
400                       sql = Replace(sql, "@Pieces3", "NULL")
410                   End If
420                   sql = Replace(sql, "@Type4", c(n))
430                   sql = Replace(sql, "@HYear5", txtYear)

440                   Cnxn(0).Execute sql

450               Next
460           End If
470       Next

480       RecClose tb

490       sql = "DELETE from histostain WHERE sampleid = '" & txtSampleID + SysOptHistoOffset(0) + Yadd & "' and hyear = '" & txtYear & "'"
500       Cnxn(0).Execute sql


510       sql = "SELECT * from HistoStain WHERE sampleid = '" & txtSampleID + SysOptHistoOffset(0) + Yadd & "' and hyear = '" & txtYear & "'"
520       Set tb = New Recordset
530       RecOpenServer 0, tb, sql

540       For n = 0 To 5
550           If c(n) <> "" And grdSpec(n).Rows > 2 Then
560               For X = 1 To Val(lblBlock(n))
570                   For b = 2 To grdSpec(n).Rows - 1
580                       sql = "Insert Into HistoStain (SampleID, Stain, Result, Block, Grid, Specimen, HYear, ResComm) " & _
                                "Values (@SampleID0, '@Stain1', '@Result2', @Block3, @Grid4, @Specimen5, '@HYear6', '@ResComm7') "

590                       sql = Replace(sql, "@SampleID0", txtSampleID + SysOptHistoOffset(0) + Yadd)
600                       sql = Replace(sql, "@Stain1", grdSpec(n).TextMatrix(b, 0))
610                       sql = Replace(sql, "@Result2", grdSpec(n).TextMatrix(b, X))
620                       sql = Replace(sql, "@Block3", X)
630                       sql = Replace(sql, "@Grid4", b - 1)
640                       sql = Replace(sql, "@Specimen5", n)
650                       sql = Replace(sql, "@HYear6", txtYear)
                          'sql = Replace(sql, "@ResComm7", grdComm(n).TextMatrix(b, x))

660                       Cnxn(0).Execute sql

670                   Next
680               Next
690           End If
700       Next

710       RecClose tb

720       Exit Sub

SaveHistoWork_Error:

          Dim strES As String
          Dim intEL As Integer

730       Screen.MousePointer = 0

740       intEL = Erl
750       strES = Err.Description
760       LogError "frmEditHisto", "SaveHistoWork", intEL, strES, sql


End Sub


Private Sub SaveHistology()
          Dim tb As New Recordset
          Dim sql As String
          Dim Yadd As Long

10        On Error GoTo SaveHistology_Error

20        Yadd = Val(Swap_Year(txtYear)) * 1000

30        SaveHistoWork

40        sql = "SELECT * FROM Historesults WHERE " & _
                "SampleID = '" & txtSampleID + SysOptHistoOffset(0) + Yadd & "' " & _
                "AND hyear = '" & txtYear & "'"
50        Set tb = New Recordset
60        RecOpenServer 0, tb, sql

70        If tb.EOF Then
80            tb.AddNew
90        End If

100       tb!SampleID = txtSampleID + SysOptHistoOffset(0) + Yadd
110       tb!Hyear = txtYear
120       tb!histocomment = c(3)
130       tb!NatureOfSpecimen = c(0)
140       tb!natureofspecimenB = Trim(c(1))
150       tb!natureofspecimenC = c(2)
160       tb!natureofspecimenD = c(3)
170       tb!natureofspecimene = c(4)
180       tb!natureofspecimenf = c(5)
190       tb!historeport = (Trim(T))
200       tb!ncri = chkNCRI(0).Value
210       tb.Update

220       Exit Sub

SaveHistology_Error:

          Dim strES As String
          Dim intEL As Integer

230       Screen.MousePointer = 0

240       intEL = Erl
250       strES = Err.Description
260       LogError "frmEditHisto", "SaveHistology", intEL, strES


End Sub


Private Sub SaveCytology()

          Dim tb As New Recordset
          Dim sql As String
          Dim Yadd As Long

10        On Error GoTo SaveCytology_Error

20        Yadd = Val(Swap_Year(txtYear)) * 1000

30        sql = "SELECT * FROM Demographics WHERE " & _
                "SampleID = '" & txtSampleID + SysOptCytoOffset(0) + Yadd & "' " & _
                "AND hYear = '" & txtYear & "'"

40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        If tb.EOF Then
70            tb.AddNew
80        End If
90        tb!SampleID = txtSampleID + SysOptCytoOffset(0) + Yadd
100       tb!Hyear = txtYear
110       tb.Update

120       sql = "SELECT * FROM CytoResults WHERE " & _
                "SampleID = '" & txtSampleID + SysOptCytoOffset(0) + Yadd & "' " & _
                "AND hYear = '" & txtYear & "'"
130       Set tb = New Recordset
140       RecOpenServer 0, tb, sql

150       If tb.EOF Then
160           tb.AddNew
170       End If

180       tb!SampleID = txtSampleID + SysOptCytoOffset(0) + Yadd
190       tb!Hyear = txtYear
200       tb!cytocomment = c(13).Text
210       tb!NatureOfSpecimen = c(6).Text
220       tb!natureofspecimenB = c(7).Text
230       tb!natureofspecimenC = c(8).Text
240       tb!natureofspecimenD = c(9).Text
250       tb!cytoreport = txtCyto
260       tb!ncri = chkNCRI(1).Value
270       tb.Update

280       Exit Sub

SaveCytology_Error:

          Dim strES As String
          Dim intEL As Integer

290       Screen.MousePointer = 0

300       intEL = Erl
310       strES = Err.Description
320       LogError "frmEditHisto", "SaveCytology", intEL, strES, sql

End Sub


Private Function SaveDemographics_Click() As Boolean


10        On Error GoTo SaveDemographics_Click_Error

20        pBar = 0

30        SaveDemographics_Click = False

40        If Trim$(txtSex) = "" Then
50            If iMsg("Sex not entered." & vbCrLf & "Do you want to enter sex now?", vbQuestion + vbYesNo) = vbYes Then
60                Exit Function
70            End If
80        End If

90        If Trim$(txtSampleID) = "" Then
100           iMsg "Must have Lab Number.", vbCritical
110           Exit Function
120       End If

130       If Trim$(txtName) <> "" Then
140           If Trim$(cmbWard) = "" Then
150               iMsg "Must have Ward entry.", vbCritical
160               Exit Function
170           End If

180           If Trim$(cmbWard) = "GP" Then
190               If Trim$(cmbGP) = "" Then
200                   iMsg "Must have GP entry.", vbCritical
210                   Exit Function
220               End If
230           End If
240       End If

250       If dtRunDate < dtSampleDate Then
260           iMsg "Sample Date After Run Date. Please Amend!"
270           Exit Function
280       End If

290       If dtRunDate < dtRecDate Then
300           iMsg "Rec. Date After Run Date. Please Amend!"
310           Exit Function
320       End If

330       If dtRecDate < dtSampleDate Then
340           iMsg "Sample Date After Rec. Date. Please Amend!"
350           Exit Function
360       End If



370       SaveDemographics
380       UPDATEMRU

390       SaveDemographics_Click = True


400       Exit Function

SaveDemographics_Click_Error:

          Dim strES As String
          Dim intEL As Integer

410       Screen.MousePointer = 0

420       intEL = Erl
430       strES = Err.Description
440       LogError "frmEditHisto", "SaveDemographics_Click", intEL, strES


End Function



Private Sub SaveInc_Click()

10        On Error GoTo SaveInc_Click_Error

20        pBar = 0

30        If Trim$(txtSex) = "" Then
40            If iMsg("Sex not entered." & vbCrLf & "Do you want to enter sex now?", vbQuestion + vbYesNo) = vbYes Then
50                Exit Sub
60            End If
70        End If

80        If Trim$(txtSampleID) = "" Then
90            iMsg "Must have Lab Number.", vbCritical
100           Exit Sub
110       End If

120       If Trim$(txtName) <> "" Then
130           If Trim$(cmbWard) = "" Then
140               iMsg "Must have Ward entry.", vbCritical
150               Exit Sub
160           End If

170           If Trim$(cmbWard) = "GP" Then
180               If Trim$(cmbGP) = "" Then
190                   iMsg "Must have GP entry.", vbCritical
200                   Exit Sub
210               End If
220           End If
230       End If

          '240   If lblChartNumber.BackColor = vbRed And Trim$(txtChart) <> "" Then
          '250     If iMsg("Confirm this Patient has" & vbCrLf & _
           '                lblChartNumber.Caption, vbYesNo + vbQuestion, , vbRed, 12) = vbNo Then
          '260       Exit Sub
          '270     End If
          '280   End If


240       If SaveDemographics_Click = False Then Exit Sub
250       UPDATEMRU


260       txtSampleID = Format$(Val(txtSampleID) + 1)
270       LoadAllDetails


280       Exit Sub

SaveInc_Click_Error:

          Dim strES As String
          Dim intEL As Integer

290       Screen.MousePointer = 0

300       intEL = Erl
310       strES = Err.Description
320       LogError "frmEditHisto", "SaveInc_Click", intEL, strES


End Sub

Private Sub bSave_Click()

10        On Error GoTo bSave_Click_Error

20        SaveComments
30        SaveInc_Click

40        bsave.Enabled = False
50        bSaveHold.Enabled = False

60        txtSampleID.SetFocus

70        Exit Sub

bSave_Click_Error:

          Dim strES As String
          Dim intEL As Integer

80        Screen.MousePointer = 0

90        intEL = Erl
100       strES = Err.Description
110       LogError "frmEditHisto", "bSave_Click", intEL, strES


End Sub
Private Sub SaveComments()

          Dim Obs As New Observations
          Dim Yadd As Long
          Dim SampleID As Long

10        On Error GoTo SaveComments_Error

20        txtSampleID = Format(Val(txtSampleID))
30        If Val(txtSampleID) = 0 Then Exit Sub

40        Yadd = Val(Swap_Year(txtYear)) * 1000
50        SampleID = txtSampleID + Yadd

60        If lblDisp = "C" Then
70            SampleID = SampleID + SysOptCytoOffset(0)
80            Obs.Save SampleID, True, _
                       "Demographic", Trim$(txtDemographicComment), _
                       "Cytology", Trim$(txtCytoComment)
90        Else
100           SampleID = SampleID + SysOptHistoOffset(0)
110           Obs.Save SampleID, True, _
                       "Demographic", Trim$(txtDemographicComment), _
                       "Histology", Trim$(txtHistoComment)
120       End If

130       Exit Sub

SaveComments_Error:

          Dim strES As String
          Dim intEL As Integer

140       intEL = Erl
150       strES = Err.Description
160       LogError "frmEditHisto", "SaveComments", intEL, strES

End Sub
Private Sub bSaveHold_Click()

10        On Error GoTo bSaveHold_Click_Error

20        If SaveDemographics_Click = False Then Exit Sub
30        SaveComments

40        bsave.Enabled = False
50        bSaveHold.Enabled = False

60        txtSampleID.SetFocus

70        Exit Sub

bSaveHold_Click_Error:

          Dim strES As String
          Dim intEL As Integer

80        Screen.MousePointer = 0

90        intEL = Erl
100       strES = Err.Description
110       LogError "frmEditHisto", "bSaveHold_Click", intEL, strES


End Sub

Private Sub bsearch_Click()

10        On Error GoTo bsearch_Click_Error

20        pBar = 0

30        With frmPatHistoryNew
40            .oHD(1) = True
50            .oFor(0) = True
60            .txtName = txtName
70            .FromEdit = True
80            .EditScreen = Me
90            .bsearch = True
100           If Not .NoPreviousDetails Then
110               .Show 1
120           Else
130               FlashNoPrevious
140           End If
150       End With

160       Exit Sub

bsearch_Click_Error:

          Dim strES As String
          Dim intEL As Integer

170       Screen.MousePointer = 0

180       intEL = Erl
190       strES = Err.Description
200       LogError "frmEditHisto", "bsearch_Click", intEL, strES


End Sub


Private Sub c_Change(Index As Integer)

10        On Error GoTo c_Change_Error

20        If Index < 6 Then
30            If c(Index) <> "" Then
40                lblBlock(Index).Visible = True
50                upBlck(Index).Visible = True
60                lblFS(Index).Visible = True
70                upFS(Index).Visible = True
80                fraSpec(Index).Visible = True
90                fraSpec(Index).Caption = "Specimen " & chr$(Asc("A") + Index) & " - " & c(Index)
100           Else
110               lblBlock(Index).Visible = False
120               upBlck(Index).Visible = False
130               lblFS(Index).Visible = False
140               upFS(Index).Visible = False
150               fraSpec(Index).Visible = False
160               fraSpec(Index).Caption = "Specimen " & chr$(Asc("A") + Index)
170           End If
180           lblBlock(Index).Caption = ""
190           grdSpec(Index).Cols = 1
200       ElseIf Index = 13 Then
210           cmdSaveCyto.Enabled = True
220           cmdSaveCHold.Enabled = True
230       End If


240       If Index > 5 And Index < 10 Then
250           cmdSaveCyto.Enabled = True
260           cmdSaveCHold.Enabled = True
270       End If

280       Exit Sub

c_Change_Error:

          Dim strES As String
          Dim intEL As Integer

290       Screen.MousePointer = 0

300       intEL = Erl
310       strES = Err.Description
320       LogError "frmEditHisto", "c_Change", intEL, strES


End Sub

Private Sub c_Click(Index As Integer)

10        On Error GoTo c_Click_Error

20        If Index < 6 Then
30            If c(Index) <> "" Then
40                lblBlock(Index).Visible = True
50                upBlck(Index).Visible = True
60                lblFS(Index).Visible = True
70                upFS(Index).Visible = True
80                fraSpec(Index).Visible = True
90                fraSpec(Index).Caption = "Specimen " & chr$(Asc("A") + Index) & " - " & c(Index)
100           Else
110               lblBlock(Index).Visible = False
120               upBlck(Index).Visible = False
130               lblFS(Index).Visible = False
140               upFS(Index).Visible = False
150               fraSpec(Index).Visible = False
160               fraSpec(Index).Caption = "Specimen " & chr$(Asc("A") + Index)
170           End If
180           grdSpec(Index).Cols = 1
190           lblBlock(Index).Caption = ""
200       End If
210       If Index = 13 Then
220           cmdSaveCyto.Enabled = True
230           cmdSaveCHold.Enabled = True
240       Else
250           cmdSaveHisto(0).Enabled = True
260           cmdSaveHHold(0).Enabled = True
270           cmdSaveHisto(1).Enabled = True
280           cmdSaveHHold(1).Enabled = True
290       End If

300       If Index > 5 And Index < 10 Then
310           cmdSaveCyto.Enabled = True
320           cmdSaveCHold.Enabled = True
330       End If

340       Exit Sub

c_Click_Error:

          Dim strES As String
          Dim intEL As Integer

350       Screen.MousePointer = 0

360       intEL = Erl
370       strES = Err.Description
380       LogError "frmEditHisto", "c_Click", intEL, strES

End Sub

Private Sub c_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

10        On Error GoTo c_KeyDown_Error

20        If Index < 6 Then
30            If c(Index) <> "" Then
40                lblBlock(Index).Visible = True
50                upBlck(Index).Visible = True
60                lblFS(Index).Visible = True
70                upFS(Index).Visible = True
80                fraSpec(Index).Visible = True
90                fraSpec(Index).Caption = "Specimen " & chr$(Asc("A") + Index) & " - " & c(Index)
100           Else
110               lblBlock(Index).Visible = False
120               upBlck(Index).Visible = False
130               lblFS(Index).Visible = False
140               upFS(Index).Visible = False
150               fraSpec(Index).Visible = False
160               fraSpec(Index).Caption = "Specimen " & chr$(Asc("A") + Index)
170           End If
180           grdSpec(Index).Cols = 1
190           lblBlock(Index).Caption = ""
200       End If

210       cmdSaveHisto(0).Enabled = True
220       cmdSaveHHold(0).Enabled = True
230       cmdSaveHisto(1).Enabled = True
240       cmdSaveHHold(1).Enabled = True

250       If Index > 5 And Index < 10 Then
260           cmdSaveCyto.Enabled = True
270           cmdSaveCHold.Enabled = True
280       End If

290       Exit Sub

c_KeyDown_Error:

          Dim strES As String
          Dim intEL As Integer

300       Screen.MousePointer = 0

310       intEL = Erl
320       strES = Err.Description
330       LogError "frmEditHisto", "c_KeyDown", intEL, strES


End Sub


Private Sub cClDetails_Click()

10        On Error GoTo cClDetails_Click_Error

20        bsave.Enabled = True
30        bSaveHold.Enabled = True

40        Exit Sub

cClDetails_Click_Error:

          Dim strES As String
          Dim intEL As Integer

50        Screen.MousePointer = 0

60        intEL = Erl
70        strES = Err.Description
80        LogError "frmEditHisto", "cClDetails_Click", intEL, strES


End Sub

Private Sub cClDetails_LostFocus()

          Dim tb As New Recordset
          Dim sql As String

10        On Error GoTo cClDetails_LostFocus_Error

20        pBar = 0

30        If Trim$(cClDetails) = "" Then Exit Sub




40        sql = "SELECT * from lists WHERE listtype = 'CD' and text = '" & cClDetails & "'"
50        Set tb = New Recordset
60        RecOpenServer 0, tb, sql
70        If Not tb.EOF Then
80            cClDetails = Trim(tb!Text)
90        End If


100       Exit Sub

cClDetails_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

110       Screen.MousePointer = 0

120       intEL = Erl
130       strES = Err.Description
140       LogError "frmEditHisto", "cClDetails_LostFocus", intEL, strES


End Sub





Private Sub chkNCRI_Click(Index As Integer)

10        On Error GoTo chkNCRI_Click_Error

20        If Index = 0 Then
30            cmdSaveHisto(0).Enabled = True
40            cmdSaveHHold(0).Enabled = True
50            cmdSaveHisto(1).Enabled = True
60            cmdSaveHHold(1).Enabled = True
70        Else
80            cmdSaveCyto.Enabled = True
90            cmdSaveCHold.Enabled = True
100       End If

110       Exit Sub

chkNCRI_Click_Error:

          Dim strES As String
          Dim intEL As Integer

120       Screen.MousePointer = 0

130       intEL = Erl
140       strES = Err.Description
150       LogError "frmEditHisto", "chkNCRI_Click", intEL, strES


End Sub

Private Sub cmbClinician_Change()
10        On Error GoTo cmbClinician_Change_Error

20        SetWardClinGP

30        Exit Sub

cmbClinician_Change_Error:

          Dim strES As String
          Dim intEL As Integer

40        Screen.MousePointer = 0

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmEditHisto", "cmbClinician_Change", intEL, strES


End Sub

Private Sub cmbClinician_Click()

10        On Error GoTo cmbClinician_Click_Error

20        bsave.Enabled = True
30        bSaveHold.Enabled = True

40        Exit Sub

cmbClinician_Click_Error:

          Dim strES As String
          Dim intEL As Integer

50        Screen.MousePointer = 0

60        intEL = Erl
70        strES = Err.Description
80        LogError "frmEditHisto", "cmbClinician_Click", intEL, strES


End Sub

Private Sub cmbClinician_KeyPress(KeyAscii As Integer)

10        On Error GoTo cmbClinician_KeyPress_Error

20        bsave.Enabled = True
30        bSaveHold.Enabled = True

40        Exit Sub

cmbClinician_KeyPress_Error:

          Dim strES As String
          Dim intEL As Integer

50        Screen.MousePointer = 0

60        intEL = Erl
70        strES = Err.Description
80        LogError "frmEditHisto", "cmbClinician_KeyPress", intEL, strES


End Sub


Private Sub cmbClinician_LostFocus()

10        On Error GoTo cmbClinician_LostFocus_Error

20        pBar = 0
30        cmbClinician = QueryKnown("Clin", cmbClinician, cmbHospital)

40        Exit Sub

cmbClinician_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

50        Screen.MousePointer = 0

60        intEL = Erl
70        strES = Err.Description
80        LogError "frmEditHisto", "cmbClinician_LostFocus", intEL, strES


End Sub

Private Sub cmbGP_Change()

10        On Error GoTo cmbGP_Change_Error

20        SetWardClinGP

30        If Trim$(cmbGP) <> "" Then
40            cmbWard = "GP"
50        End If

60        Exit Sub

cmbGP_Change_Error:

          Dim strES As String
          Dim intEL As Integer

70        Screen.MousePointer = 0

80        intEL = Erl
90        strES = Err.Description
100       LogError "frmEditHisto", "cmbGP_Change", intEL, strES


End Sub
Private Sub SetWardClinGP()

          Dim GPAddr As String

10        On Error GoTo SetWardClinGP_Error

20        GPAddr = AddressOfGP(cmbGP)

30        lAddWardGP = Trim$(taddress(0)) & " " & Trim$(taddress(1)) & " : " & cmbWard & " : " & cmbGP & ":" & GPAddr & " " & cmbClinician

40        Exit Sub

SetWardClinGP_Error:

          Dim strES As String
          Dim intEL As Integer

50        Screen.MousePointer = 0

60        intEL = Erl
70        strES = Err.Description
80        LogError "frmEditHisto", "SetWardClinGP", intEL, strES

End Sub

Private Sub cmbGP_Click()

10        On Error GoTo cmbGP_Click_Error

20        pBar = 0

30        SetWardClinGP

40        cmbWard = "GP"
50        bsave.Enabled = True
60        bSaveHold.Enabled = True

70        Exit Sub

cmbGP_Click_Error:

          Dim strES As String
          Dim intEL As Integer

80        Screen.MousePointer = 0

90        intEL = Erl
100       strES = Err.Description
110       LogError "frmEditHisto", "cmbGP_Click", intEL, strES

End Sub


Private Sub cmbGP_KeyPress(KeyAscii As Integer)

10        On Error GoTo cmbGP_KeyPress_Error

20        bsave.Enabled = True
30        bSaveHold.Enabled = True

40        Exit Sub

cmbGP_KeyPress_Error:

          Dim strES As String
          Dim intEL As Integer

50        Screen.MousePointer = 0

60        intEL = Erl
70        strES = Err.Description
80        LogError "frmEditHisto", "cmbGP_KeyPress", intEL, strES


End Sub


Private Sub cmbGP_LostFocus()

10        On Error GoTo cmbGP_LostFocus_Error

20        cmbGP = QueryKnown("GP", cmbGP, cmbHospital)

30        Exit Sub

cmbGP_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

40        Screen.MousePointer = 0

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmEditHisto", "cmbGP_LostFocus", intEL, strES


End Sub


Private Sub CheckDepartments()

10        On Error GoTo CheckDepartments_Error
20        If Trim$(txtSampleID) <> "" And Trim$(txtYear) <> "" Then
30            If SysOptDeptHisto(0) = True And lblDisp = "H" Then
40                If AreHistoResultsPresent(txtSampleID, txtYear) = 1 Then
50                    SSTab1.TabCaption(1) = "<<Histology Work Screen>>"
60                    SSTab1.TabCaption(2) = "<<Histology Report>>"
70                    SSTab1.TabVisible(3) = False
80                End If
90            End If

100           If SysOptDeptCyto(0) = True And lblDisp = "C" Then
110               If AreCytoResultsPresent(txtSampleID, txtYear) = 1 Then
120                   SSTab1.TabCaption(3) = "<<Cytology>>"
130                   SSTab1.TabVisible(1) = False
140                   SSTab1.TabVisible(2) = False
150               End If
160           End If
170       End If

180       Exit Sub

CheckDepartments_Error:

          Dim strES As String
          Dim intEL As Integer

190       Screen.MousePointer = 0

200       intEL = Erl
210       strES = Err.Description
220       LogError "frmEditHisto", "CheckDepartments", intEL, strES


End Sub




Private Sub cmbHospital_Click()


10        On Error GoTo cmbHospital_Click_Error

20        FillGPsClinWard Me, cmbHospital

30        bsave.Enabled = True
40        bSaveHold.Enabled = True

50        Exit Sub

cmbHospital_Click_Error:

          Dim strES As String
          Dim intEL As Integer

60        Screen.MousePointer = 0

70        intEL = Erl
80        strES = Err.Description
90        LogError "frmEditHisto", "cmbHospital_Click", intEL, strES


End Sub

Private Sub cmbHospital_LostFocus()

          Dim n As Long

10        On Error GoTo cmbHospital_LostFocus_Error

20        For n = 0 To cmbHospital.ListCount
30            If UCase(cmbHospital) = UCase(Left(cmbHospital.List(n), Len(cmbHospital))) Then
40                cmbHospital.ListIndex = n
50            End If
60        Next


70        Exit Sub

cmbHospital_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

80        Screen.MousePointer = 0

90        intEL = Erl
100       strES = Err.Description
110       LogError "frmEditHisto", "cmbHospital_LostFocus", intEL, strES


End Sub

Private Sub cmbStain_Click(Index As Integer)

10        On Error GoTo cmbStain_Click_Error

20        grdSpec(Index).AddItem cmbStain(Index)
30        grdComm(Index).AddItem ""
40        cmdSaveHisto(0).Enabled = True
50        cmdSaveHHold(0).Enabled = True
60        cmdSaveHisto(1).Enabled = True
70        cmdSaveHHold(1).Enabled = True

80        Exit Sub

cmbStain_Click_Error:

          Dim strES As String
          Dim intEL As Integer

90        Screen.MousePointer = 0

100       intEL = Erl
110       strES = Err.Description
120       LogError "frmEditHisto", "cmbStain_Click", intEL, strES


End Sub

Private Sub cmdCopyTo_Click()
          Dim s As String

10        On Error GoTo cmdCopyTo_Click_Error

20        s = cmbWard & " " & cmbClinician & " " & cmbGP
30        s = Trim$(s)

40        frmCopyTo.EditScreen = Me
50        frmCopyTo.lblOriginal = s
60        If lblDisp = "H" Then
70            frmCopyTo.lblSampleID = txtSampleID + SysOptHistoOffset(0) + Val(Swap_Year(txtYear) * 1000)
80        Else
90            frmCopyTo.lblSampleID = txtSampleID + SysOptCytoOffset(0) + Val(Swap_Year(txtYear) * 1000)
100       End If
110       frmCopyTo.Show 1

120       CheckCC

130       Exit Sub

cmdCopyTo_Click_Error:

          Dim strES As String
          Dim intEL As Integer

140       Screen.MousePointer = 0

150       intEL = Erl
160       strES = Err.Description
170       LogError "frmEditHisto", "cmdCopyTo_Click", intEL, strES

End Sub
Private Sub CheckCC()

          Dim sql As String
          Dim tb As Recordset



10        On Error GoTo CheckCC_Error

20        cmdCopyTo.Caption = "cc"
30        cmdCopyTo.Font.Bold = False
40        cmdCopyTo.BackColor = &H8000000F

50        If Trim$(txtSampleID) = "" Then Exit Sub

60        If lblDisp = "H" Then
70            sql = "Select * from SendCopyTo where " & _
                    "SampleID = '" & txtSampleID + SysOptHistoOffset(0) + Val(Swap_Year(txtYear) * 1000) & "'"
80        Else
90            sql = "Select * from SendCopyTo where " & _
                    "SampleID = '" & txtSampleID + SysOptCytoOffset(0) + Val(Swap_Year(txtYear) * 1000) & "'"
100       End If
110       Set tb = New Recordset
120       RecOpenServer 0, tb, sql
130       If Not tb.EOF Then
140           cmdCopyTo.Caption = "++ cc ++"
150           cmdCopyTo.Font.Bold = True
160           cmdCopyTo.BackColor = &H8080FF
170       End If





180       Exit Sub

CheckCC_Error:

          Dim strES As String
          Dim intEL As Integer

190       Screen.MousePointer = 0

200       intEL = Erl
210       strES = Err.Description
220       LogError "frmEditHisto", "CheckCC", intEL, strES, sql


End Sub

Private Sub cmdCVal_Click()
          Dim Samp As Long
          Dim sql As String



10        On Error GoTo cmdCVal_Click_Error

20        Samp = txtSampleID + SysOptCytoOffset(0) + Val(Swap_Year(txtYear) * 1000)

30        If UserMemberOf = "Users" Then
40            iMsg "You are not allowed to change validation status.", vbInformation
50            Exit Sub
60        End If

70        If cmdCVal.Caption = "Validate" Then
80            If UserMemberOf = "Secretaries" Then
90                iMsg "You are not allowed to validate results.", vbInformation
100               Exit Sub
110           End If
120           If Trim(txtCyto & "") = "" Then
130               iMsg "No Report to Validate!"
140               Exit Sub
150           End If
160           sql = "UPDATE demographics set cytovalid = 1 WHERE sampleid = " & Samp & ""
170           Cnxn(0).Execute sql
180           sql = "UPDATE cytoresults set ValidDate = '" & Format(Now, "dd/MMM/yyyy hh:mm") & "', username = '" & UserCode & "' where Sampleid = " & Samp & ""
190           Cnxn(0).Execute sql
200           bSave_Click
210           SaveCytology
220           SaveComments
230           LockCRecord
240       Else
250           If UCase(iBOX("Enter Password to Unvalidate", , , True)) = UserPass Then
260               sql = "UPDATE cytoresults set username = '" & UserCode & "' where Sampleid = " & Samp & ""
270               Cnxn(0).Execute sql
280               sql = "UPDATE demographics set cytovalid = 0 WHERE sampleid = " & Samp & ""
290               Cnxn(0).Execute sql
300               UnlockCRecord
310           End If
320       End If

330       Exit Sub

cmdCVal_Click_Error:

          Dim strES As String
          Dim intEL As Integer

340       Screen.MousePointer = 0

350       intEL = Erl
360       strES = Err.Description
370       LogError "frmEditHisto", "cmdCVal_Click", intEL, strES, sql


End Sub
Private Sub cmdHVal_Click()

          Dim Samp As Long
          Dim sql As String

10        On Error GoTo cmdHVal_Click_Error

20        Samp = txtSampleID + SysOptHistoOffset(0) + Val(Swap_Year(txtYear) * 1000)

30        If UserMemberOf = "Users" Then
40            iMsg "You are not allowed to change validation status.", vbInformation
50            Exit Sub
60        End If

70        If cmdHVal.Caption = "Validate" Then
80            If UserMemberOf = "Secretaries" Then
90                iMsg "You are not allowed to validate results.", vbInformation
100               Exit Sub
110           End If
120           If Trim(T) = "" Then
130               iMsg "No report to validate!"
140               Exit Sub
150           End If
160           sql = "UPDATE demographics set histovalid = 1 where Sampleid = " & Samp & ""
170           Cnxn(0).Execute sql
180           sql = "UPDATE historesults set ValidDate = '" & Format(Now, "dd/MMM/yyyy hh:mm") & "', username = '" & UserCode & "' where Sampleid = " & Samp & ""
190           Cnxn(0).Execute sql
200           bSave_Click
210           SaveHistoWork
220           SaveHistology
230           SaveComments
240           LockRecord
250       Else
260           If UCase(iBOX("Enter Password to Unvalidate", , , True)) = UserPass Then
270               sql = "UPDATE historesults set username = '' where Sampleid = " & Samp & ""
280               Cnxn(0).Execute sql
290               sql = "UPDATE demographics set histovalid = 0 WHERE sampleid = " & Samp & ""
300               Cnxn(0).Execute sql
310               UnlockRecord
320           End If
330       End If

340       Exit Sub

cmdHVal_Click_Error:

          Dim strES As String
          Dim intEL As Integer

350       Screen.MousePointer = 0

360       intEL = Erl
370       strES = Err.Description
380       LogError "frmEditHisto", "cmdHVal_Click", intEL, strES, sql

End Sub

Private Sub cmdSaveCHold_Click()

10        On Error GoTo cmdSaveCHold_Click_Error

20        SaveCytology
30        SaveComments

40        cmdSaveCyto.Enabled = False
50        cmdSaveCHold.Enabled = False


60        Exit Sub

cmdSaveCHold_Click_Error:

          Dim strES As String
          Dim intEL As Integer

70        Screen.MousePointer = 0

80        intEL = Erl
90        strES = Err.Description
100       LogError "frmEditHisto", "cmdSaveCHold_Click", intEL, strES


End Sub

Private Sub cmdSaveHHold_Click(Index As Integer)

10        On Error GoTo cmdSaveHHold_Click_Error

20        SaveHistoWork
30        SaveHistology
40        SaveComments

50        cmdSaveHisto(0).Enabled = False
60        cmdSaveHHold(0).Enabled = False
70        cmdSaveHisto(1).Enabled = False
80        cmdSaveHHold(1).Enabled = False

90        txtSampleID.SetFocus

100       Exit Sub

cmdSaveHHold_Click_Error:

          Dim strES As String
          Dim intEL As Integer

110       Screen.MousePointer = 0

120       intEL = Erl
130       strES = Err.Description
140       LogError "frmEditHisto", "cmdSaveHHold_Click", intEL, strES


End Sub

Private Sub cmdSaveHisto_Click(Index As Integer)

10        On Error GoTo cmdSaveHisto_Click_Error

20        If bsave.Enabled = True Then
30            SaveDemographics
40        End If

50        SaveHistoWork
60        SaveHistology
70        SaveComments

80        txtSampleID = Format$(Val(txtSampleID) + 1)

90        LoadAllDetails

100       SSTab1.Tab = 0

110       bsave.Enabled = False
120       bSaveHold.Enabled = False
130       cmdSaveHisto(0).Enabled = False
140       cmdSaveHHold(0).Enabled = False
150       cmdSaveHisto(1).Enabled = False
160       cmdSaveHHold(1).Enabled = False

170       txtSampleID.SetFocus

180       Exit Sub

cmdSaveHisto_Click_Error:

          Dim strES As String
          Dim intEL As Integer

190       Screen.MousePointer = 0

200       intEL = Erl
210       strES = Err.Description
220       LogError "frmEditHisto", "cmdSaveHisto_Click", intEL, strES


End Sub

Private Sub cmdSetPrinter_Click()

10        On Error GoTo cmdSetPrinter_Click_Error

20        Set frmForcePrinter.f = frmEditHisto
30        frmForcePrinter.Show 1

40        If pPrintToPrinter = "Automatic SELECTion" Then
50            pPrintToPrinter = ""
60        End If

70        If pPrintToPrinter <> "" Then
80            cmdSetPrinter.BackColor = vbRed
90            cmdSetPrinter.ToolTipText = "Print Forced to " & pPrintToPrinter
100       Else
110           cmdSetPrinter.BackColor = vbButtonFace
120           pPrintToPrinter = ""
130           cmdSetPrinter.ToolTipText = "Printer SELECTed Automatically"
140       End If

150       Exit Sub

cmdSetPrinter_Click_Error:

          Dim strES As String
          Dim intEL As Integer

160       Screen.MousePointer = 0

170       intEL = Erl
180       strES = Err.Description
190       LogError "frmEditHisto", "cmdSetPrinter_Click", intEL, strES


End Sub



Private Sub cmdViewHistoRep_Click()

          Dim Yadd

10        On Error GoTo cmdViewHistoRep_Click_Error

20        Yadd = Val(Swap_Year(txtYear)) * 1000


30        frmRFT.SampleID = txtSampleID + SysOptHistoOffset(0) + Yadd
40        frmRFT.Dept = "P"
50        frmRFT.Show 1

60        Exit Sub

cmdViewHistoRep_Click_Error:

          Dim strES As String
          Dim intEL As Integer

70        Screen.MousePointer = 0

80        intEL = Erl
90        strES = Err.Description
100       LogError "frmEditHisto", "cmdViewHistoRep_Click", intEL, strES


End Sub

Private Sub cmdViewMicroRep_Click()
          Dim Yadd

10        On Error GoTo cmdViewMicroRep_Click_Error

20        Yadd = Val(Swap_Year(txtYear)) * 1000


30        frmRFT.SampleID = txtSampleID + SysOptCytoOffset(0) + Yadd
40        frmRFT.Dept = "Y"
50        frmRFT.Show 1

60        Exit Sub

cmdViewMicroRep_Click_Error:

          Dim strES As String
          Dim intEL As Integer

70        Screen.MousePointer = 0

80        intEL = Erl
90        strES = Err.Description
100       LogError "frmEditHisto", "cmdViewMicroRep_Click", intEL, strES


End Sub

Private Sub cMRU_Click()

10        On Error GoTo cMRU_Click_Error

20        txtYear = Left(cMRU, 4)
30        If Mid(cMRU, 5, 1) = "H" Then lblDisp = "H" Else lblDisp = "C"
40        txtSampleID = Mid(cMRU, 6, Len(cMRU) - 5)

50        LoadAllDetails

60        bsave.Enabled = False

70        Exit Sub

cMRU_Click_Error:

          Dim strES As String
          Dim intEL As Integer

80        Screen.MousePointer = 0

90        intEL = Erl
100       strES = Err.Description
110       LogError "frmEditHisto", "cMRU_Click", intEL, strES


End Sub

Private Sub cMRU_KeyPress(KeyAscii As Integer)

10        On Error GoTo cMRU_KeyPress_Error

20        KeyAscii = 0

30        Exit Sub

cMRU_KeyPress_Error:

          Dim strES As String
          Dim intEL As Integer

40        Screen.MousePointer = 0

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmEditHisto", "cMRU_KeyPress", intEL, strES


End Sub

Private Sub cmbWard_Change()

10        On Error GoTo cmbWard_Change_Error

20        SetWardClinGP

30        Exit Sub

cmbWard_Change_Error:

          Dim strES As String
          Dim intEL As Integer

40        Screen.MousePointer = 0

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmEditHisto", "cmbWard_Change", intEL, strES


End Sub

Private Sub cmbWard_Click()

10        On Error GoTo cmbWard_Click_Error

20        SetWardClinGP

30        bsave.Enabled = True
40        bSaveHold.Enabled = True

50        Exit Sub

cmbWard_Click_Error:

          Dim strES As String
          Dim intEL As Integer

60        Screen.MousePointer = 0

70        intEL = Erl
80        strES = Err.Description
90        LogError "frmEditHisto", "cmbWard_Click", intEL, strES

End Sub


Private Sub cmbWard_KeyPress(KeyAscii As Integer)

10        On Error GoTo cmbWard_KeyPress_Error

20        bsave.Enabled = True
30        bSaveHold.Enabled = True

40        Exit Sub

cmbWard_KeyPress_Error:

          Dim strES As String
          Dim intEL As Integer

50        Screen.MousePointer = 0

60        intEL = Erl
70        strES = Err.Description
80        LogError "frmEditHisto", "cmbWard_KeyPress", intEL, strES


End Sub


Private Sub cmbWard_LostFocus()
          Dim Found As Boolean
          Dim tb As New Recordset
          Dim sql As String



10        On Error GoTo cmbWard_LostFocus_Error

20        If Trim$(cmbWard) = "" Then
30            cmbWard = "GP"
40            Exit Sub
50        End If

60        Found = False

70        sql = "SELECT * from wards WHERE (text = '" & AddTicks(cmbWard) & "' or code = '" & AddTicks(cmbWard) & "') and hospitalcode = '" & ListCodeFor("HO", cmbHospital) & "' And InUse = '1'"

80        Set tb = New Recordset
90        RecOpenServer 0, tb, sql
100       If Not tb.EOF Then
110           cmbWard = Trim(tb!Text)
120           Found = True
130       End If

140       If Not Found Then
150           cmbWard = "GP"
160       End If




170       Exit Sub

cmbWard_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

180       Screen.MousePointer = 0

190       intEL = Erl
200       strES = Err.Description
210       LogError "frmEditHisto", "cmbWard_LostFocus", intEL, strES


End Sub

Private Sub cRooH_Click(Index As Integer)


10        On Error GoTo cRooH_Click_Error

20        bsave.Enabled = True
30        bSaveHold.Enabled = True

40        Exit Sub

cRooH_Click_Error:

          Dim strES As String
          Dim intEL As Integer

50        Screen.MousePointer = 0

60        intEL = Erl
70        strES = Err.Description
80        LogError "frmEditHisto", "cRooH_Click", intEL, strES


End Sub

Private Sub dtRecDate_CloseUp()

10        On Error GoTo dtRecDate_CloseUp_Error

20        pBar = 0

30        bsave.Enabled = True
40        bSaveHold.Enabled = True

50        Exit Sub

dtRecDate_CloseUp_Error:

          Dim strES As String
          Dim intEL As Integer

60        Screen.MousePointer = 0

70        intEL = Erl
80        strES = Err.Description
90        LogError "frmEditHisto", "dtRecDate_CloseUp", intEL, strES

End Sub

Private Sub dtRunDate_CloseUp()

10        On Error GoTo dtRunDate_CloseUp_Error

20        pBar = 0

30        bsave.Enabled = True
40        bSaveHold.Enabled = True


50        Exit Sub

dtRunDate_CloseUp_Error:

          Dim strES As String
          Dim intEL As Integer

60        Screen.MousePointer = 0

70        intEL = Erl
80        strES = Err.Description
90        LogError "frmEditHisto", "dtRunDate_CloseUp", intEL, strES


End Sub

Private Sub dtSampleDate_CloseUp()

10        On Error GoTo dtSampleDate_CloseUp_Error

20        pBar = 0

30        bsave.Enabled = True
40        bSaveHold.Enabled = True


50        Exit Sub

dtSampleDate_CloseUp_Error:

          Dim strES As String
          Dim intEL As Integer

60        Screen.MousePointer = 0

70        intEL = Erl
80        strES = Err.Description
90        LogError "frmEditHisto", "dtSampleDate_CloseUp", intEL, strES


End Sub

Sub FillComments()
          Dim sql As String
          Dim tb As New Recordset

10        On Error GoTo FillComments_Error

20        Screen.MousePointer = 11

30        c(13).Clear
40        c(0).Clear


50        sql = "SELECT * from lists WHERE listtype = 'CI'"
60        Set tb = New Recordset
70        RecOpenServer 0, tb, sql
80        Do While Not tb.EOF
90            c(13).AddItem Trim(tb!Text)
100           tb.MoveNext
110       Loop
120       c(13).ListIndex = -1

130       Screen.MousePointer = 0

140       Exit Sub

FillComments_Error:

          Dim strES As String
          Dim intEL As Integer

150       Screen.MousePointer = 0

160       intEL = Erl
170       strES = Err.Description
180       LogError "frmEditHisto", "FillComments", intEL, strES


End Sub
Private Sub FillLists()

          Dim HospCode As String
          Dim sql As String
          Dim tb As New Recordset

10        On Error GoTo FillLists_Error

20        Screen.MousePointer = 11

30        HospCode = ListCodeFor("HO", HospName(0))

40        FillGPsClinWard Me, HospName(0)

50        cmbHospital.Clear

60        sql = "SELECT * from lists WHERE listtype = 'HO'"
70        Set tb = New Recordset
80        RecOpenServer 0, tb, sql
90        Do While Not tb.EOF
100           cmbHospital.AddItem Trim(tb!Text)
110           tb.MoveNext
120       Loop

130       sql = "SELECT * from lists WHERE listtype = 'CD'"
140       Set tb = New Recordset
150       RecOpenServer 0, tb, sql
160       Do While Not tb.EOF
170           cClDetails.AddItem Trim(tb!Text)
180           tb.MoveNext
190       Loop

200       cmbHospital.ListIndex = -1
210       cClDetails.ListIndex = -1

220       Exit Sub

FillLists_Error:

          Dim strES As String
          Dim intEL As Integer

230       Screen.MousePointer = 0

240       intEL = Erl
250       strES = Err.Description
260       LogError "frmEditHisto", "FillLists", intEL, strES


End Sub

Private Sub FillMRU()

          Dim sql As String
          Dim tb As New Recordset

10        On Error GoTo FillMRU_Error

20        sql = "SELECT top 10 * from HMRU WHERE " & _
                "UserCode = '" & UserCode & "' " & _
                "Order by DateTime desc"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql

50        cMRU.Clear
60        Do While Not tb.EOF
70            cMRU.AddItem Trim$(tb!SampleID & "")
80            tb.MoveNext
90        Loop
100       If cMRU.ListCount > 0 Then
110           cMRU = ""
120       End If

130       Exit Sub

FillMRU_Error:

          Dim strES As String
          Dim intEL As Integer

140       Screen.MousePointer = 0

150       intEL = Erl
160       strES = Err.Description
170       LogError "frmEditHisto", "FillMRU", intEL, strES


End Sub

Private Sub FlashNoPrevious()

          Dim T As Single
          Dim n As Long

10        On Error GoTo FlashNoPrevious_Error

20        For n = 1 To 5
30            lNoPrevious.Visible = True
40            lNoPrevious.Refresh
50            T = Timer
60            Do While Timer - T < 0.1: DoEvents: Loop
70            lNoPrevious.Visible = False
80            lNoPrevious.Refresh
90            T = Timer
100           Do While Timer - T < 0.1: DoEvents: Loop
110       Next

120       Exit Sub

FlashNoPrevious_Error:

          Dim strES As String
          Dim intEL As Integer

130       Screen.MousePointer = 0

140       intEL = Erl
150       strES = Err.Description
160       LogError "frmEditHisto", "FlashNoPrevious", intEL, strES


End Sub

Private Sub Form_Activate()

10        On Error GoTo Form_Activate_Error

20        TimerBar.Enabled = True
30        pBar = 0
40        UpDown1.Max = 99999999
50        Set_Font Me

60        Exit Sub

Form_Activate_Error:

          Dim strES As String
          Dim intEL As Integer

70        Screen.MousePointer = 0

80        intEL = Erl
90        strES = Err.Description
100       LogError "frmEditHisto", "Form_Activate", intEL, strES


End Sub

Private Sub Form_Deactivate()

10        On Error GoTo Form_Deactivate_Error

20        pBar = 0
30        TimerBar.Enabled = False

40        Exit Sub

Form_Deactivate_Error:

          Dim strES As String
          Dim intEL As Integer

50        Screen.MousePointer = 0

60        intEL = Erl
70        strES = Err.Description
80        LogError "frmEditHisto", "Form_Deactivate", intEL, strES


End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
          Dim s As String

10        On Error GoTo Form_KeyDown_Error

20        If SSTab1.Tab = 1 Then
30            If KeyCode = vbKeyD Then
40                If Shift And vbAltMask Then
50                    s = Format(Now, "dd/mm/yyyy")
60                    T.SelText = s
70                    KeyCode = 0
80                End If
90            End If
100           If KeyCode = vbKeyW Then
110               If Shift And vbAltMask Then
120                   T.SetFocus
130                   KeyCode = 0
140               End If
150           End If
160       ElseIf SSTab1.Tab = 2 Then
170           If KeyCode = vbKeyD Then
180               If Shift And vbAltMask Then
190                   s = Format(Now, "dd/mm/yyyy")
200                   txtCyto.SelText = s
210                   KeyCode = 0
220               End If
230           End If
240           If KeyCode = vbKeyW Then
250               If Shift And vbAltMask Then
260                   txtCyto.SetFocus
270                   KeyCode = 0
280               End If
290           End If
300       End If


310       Exit Sub

Form_KeyDown_Error:

          Dim strES As String
          Dim intEL As Integer

320       Screen.MousePointer = 0

330       intEL = Erl
340       strES = Err.Description
350       LogError "frmEditHisto", "Form_KeyDown", intEL, strES


End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

10        On Error GoTo Form_KeyPress_Error

20        pBar = 0

30        Exit Sub

Form_KeyPress_Error:

          Dim strES As String
          Dim intEL As Integer

40        Screen.MousePointer = 0

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmEditHisto", "Form_KeyPress", intEL, strES


End Sub

Sub FillNatureOfSpecimens()

          Dim sql As String
          Dim tb As New Recordset
          Dim n As Long

10        On Error GoTo FillNatureOfSpecimens_Error

20        Screen.MousePointer = 11

30        For n = 0 To 9
40            c(n).Clear
50        Next

60        sql = "SELECT * from lists WHERE listtype = 'NA'"
70        Set tb = New Recordset
80        RecOpenServer 0, tb, sql

90        Do While Not tb.EOF
100           For n = 0 To 9
110               c(n).AddItem StrConv(Trim(tb!Text), vbProperCase)
120           Next
130           tb.MoveNext
140       Loop

150       For n = 0 To 9
160           c(n).AddItem ""
170           c(n).ListIndex = -1
180       Next

190       Screen.MousePointer = 0

200       Exit Sub

FillNatureOfSpecimens_Error:

          Dim strES As String
          Dim intEL As Integer

210       Screen.MousePointer = 0

220       intEL = Erl
230       strES = Err.Description
240       LogError "frmEditHisto", "FillNatureOfSpecimens", intEL, strES, sql


End Sub
Private Sub Form_Load()
          Dim n As Long

10        On Error GoTo Form_Load_Error

20        n = n + 1

30        If SysOptDeptHisto(0) = False Then
40            SSTab1.TabVisible(1) = False
50            SSTab1.TabVisible(2) = False
60        Else
70            n = n + 2
80        End If

90        If SysOptDeptCyto(0) = False Then SSTab1.TabVisible(3) = False Else n = n + 1

100       SSTab1.TabsPerRow = n

110       StatusBar1.Panels(1).Text = UserName

120       bPrintHold.Visible = False
130       bPrint.Visible = False

140       For n = 0 To SysOptHistoSamps(0)
150           c(n).Visible = True
160           Label5(n).Visible = True
170           lblBlock(n).Visible = True
180           upBlck(n).Visible = True
190           upFS(n).Visible = True
200           lblFS(n).Visible = True
210       Next

220       cmdViewHistoRep.Visible = SysOptRTFView(0)
230       cmdViewMicroRep.Visible = SysOptRTFView(0)

          '235   If TestSys Then EnableTestMode Me

240       pPrintToPrinter = GetSetting("Netacquire", "Histology", "Printer", "")

250       FillComments
260       FillLists
270       FillNatureOfSpecimens
280       Clear_HistoWork



290       With lblChartNumber
300           .BackColor = &H8000000F
310           .ForeColor = vbBlack
320           Select Case HospName(0)
              Case "PORTLAOISE", "TULLAMORE"
330               .Caption = initial2upper(HospName(0)) & " Chart #"
340               lblAandE.Visible = False
350               lblNOPAS(0).Visible = False
                  '      lblNOPAS(1).Visible = False
360               txtAandE.Visible = False
370               txtNOPAS.Visible = False
380               lblNameTitle.Left = 4350
390               txtName.Left = 4350
400               txtName.Width = 6015
410           Case "MULLINGAR"
420               .Caption = initial2upper(HospName(0)) & " Chart #"
430               lblAandE.Visible = True
440               lblNOPAS(0).Visible = False
450               lblNOPAS(1).Visible = False
460               txtAandE.Visible = True
470               txtNOPAS.Visible = False
480               txtAandE.Width = 2000
490               lblNameTitle.Left = 3550
500               txtName.Left = 3550
510               txtName.Width = 4000
520           End Select
530       End With


540       dtRunDate = Format$(Now, "dd/mm/yyyy")
550       dtSampleDate = Format$(Now, "dd/mm/yyyy")
560       dtRecDate = Format$(Now, "dd/mm/yyyy")

570       UpDown1.Max = 999999

580       txtYear = GetSetting("NetAcquire", "StartUp", "LastYear", "2000")
590       lblDisp = "H"    'GetSetting("NetAcquire", "StartUp", "LastDisp", "H")
600       txtSampleID = GetSetting("NetAcquire", "StartUp", "LastHisto", "1")


610       If lblDisp = "C" Then
620           lblDisp = "C"
630           SSTab1.TabVisible(1) = False
640           SSTab1.TabVisible(2) = False
650           SSTab1.TabVisible(3) = True
660       Else
670           lblDisp = "H"
680           SSTab1.TabVisible(1) = True
690           SSTab1.TabVisible(2) = True
700           SSTab1.TabVisible(3) = False
710       End If

720       LoadAllDetails

730       If UserMemberOf = "Secretarys" Then
740           For n = 1 To 4
750               SSTab1.TabVisible(n) = False
760           Next
770       Else
780           bSaveHold.Enabled = False
790           bsave.Enabled = False
800           cmdSaveHisto(0).Enabled = False
810           cmdSaveHHold(0).Enabled = False
820           cmdSaveHisto(1).Enabled = False
830           cmdSaveHHold(1).Enabled = False
840           cmdSaveCyto.Enabled = False
850           cmdSaveCHold.Enabled = False
860       End If

870       Activated = False

880       Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

890       Screen.MousePointer = 0

900       intEL = Erl
910       strES = Err.Description
920       LogError "frmEditHisto", "Form_Load", intEL, strES


End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        On Error GoTo Form_MouseMove_Error

20        pBar = 0

30        Exit Sub

Form_MouseMove_Error:

          Dim strES As String
          Dim intEL As Integer

40        Screen.MousePointer = 0

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmEditHisto", "Form_MouseMove", intEL, strES


End Sub

Private Sub Form_Paint()

          Dim TabNumber As Long

10        On Error GoTo Form_Paint_Error

20        If Activated Then Exit Sub

30        Activated = True

40        TabNumber = Val(GetSetting("NetAcquire", "StartUp", "Histology", "0"))

50        If SSTab1.TabVisible(TabNumber) Then
60            SSTab1.Tab = TabNumber
70        Else
80            SSTab1.Tab = 0
90        End If


100       Exit Sub

Form_Paint_Error:

          Dim strES As String
          Dim intEL As Integer

110       Screen.MousePointer = 0

120       intEL = Erl
130       strES = Err.Description
140       LogError "frmEditHisto", "Form_Paint", intEL, strES


End Sub

Private Sub Form_Unload(Cancel As Integer)


10        On Error GoTo Form_Unload_Error

20        SaveSetting "NetAcquire", "StartUp", "LastHisto", txtSampleID
30        SaveSetting "NetAcquire", "StartUp", "LastYear", txtYear
40        SaveSetting "NetAcquire", "StartUp", "LastDisp", lblDisp

50        SaveSetting "NetAcquire", "StartUp", "Histology", CStr(SSTab1.Tab)


60        pPrintToPrinter = ""

70        Activated = False

80        Exit Sub

Form_Unload_Error:

          Dim strES As String
          Dim intEL As Integer

90        Screen.MousePointer = 0

100       intEL = Erl
110       strES = Err.Description
120       LogError "frmEditHisto", "Form_Unload", intEL, strES


End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        On Error GoTo Frame1_MouseMove_Error

20        pBar = 0

30        Exit Sub

Frame1_MouseMove_Error:

          Dim strES As String
          Dim intEL As Integer

40        Screen.MousePointer = 0

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmEditHisto", "Frame1_MouseMove", intEL, strES


End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        On Error GoTo Frame2_MouseMove_Error

20        pBar = 0

30        Exit Sub

Frame2_MouseMove_Error:

          Dim strES As String
          Dim intEL As Integer

40        Screen.MousePointer = 0

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmEditHisto", "Frame2_MouseMove", intEL, strES


End Sub

Private Sub grdSpec_Click(Index As Integer)

10        On Error GoTo grdSpec_Click_Error

20        If grdSpec(Index).RowSel > 0 Then
30            If InStr(grdSpec(Index).TextMatrix(grdSpec(Index).RowSel, grdSpec(Index).ColSel), "Pieces") Then Exit Sub
40            If InStr(grdSpec(Index).TextMatrix(grdSpec(Index).RowSel, grdSpec(Index).ColSel), "Blk") Then Exit Sub
50            If grdSpec(Index).ColSel > 0 Then
60                tinput(Index).Text = grdSpec(Index).TextMatrix(grdSpec(Index).RowSel, grdSpec(Index).ColSel)
70                tinput(Index).SetFocus
80                Exit Sub
90            End If
100       End If

110       Exit Sub

grdSpec_Click_Error:

          Dim strES As String
          Dim intEL As Integer

120       Screen.MousePointer = 0

130       intEL = Erl
140       strES = Err.Description
150       LogError "frmEditHisto", "grdSpec_Click", intEL, strES

End Sub

Private Sub grdSpec_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

10        On Error GoTo grdSpec_KeyDown_Error

20        If KeyCode = 116 Then
30            If Index < 6 Then
40                grdSpec(Index + 1).SetFocus
50            Else
60                grdSpec(0).SetFocus
70            End If
80        End If



90        Exit Sub

grdSpec_KeyDown_Error:

          Dim strES As String
          Dim intEL As Integer

100       Screen.MousePointer = 0

110       intEL = Erl
120       strES = Err.Description
130       LogError "frmEditHisto", "grdSpec_KeyDown", intEL, strES


End Sub

Private Sub grdSpec_KeyPress(Index As Integer, KeyAscii As Integer)


10        On Error GoTo grdSpec_KeyPress_Error

20        If KeyAscii = 116 Then
30            If Index < 6 Then
40                grdSpec(Index + 1).SetFocus
50            Else
60                grdSpec(0).SetFocus
70            End If
80        End If

90        Exit Sub

grdSpec_KeyPress_Error:

          Dim strES As String
          Dim intEL As Integer

100       Screen.MousePointer = 0

110       intEL = Erl
120       strES = Err.Description
130       LogError "frmEditHisto", "grdSpec_KeyPress", intEL, strES


End Sub

Private Sub grdSpec_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

10        On Error GoTo grdSpec_MouseMove_Error

20        grdSpec(Index).ToolTipText = ""
30        If grdSpec(Index).CellBackColor = vbYellow Then
40            grdSpec(Index).ToolTipText = grdComm(Index).TextMatrix(grdSpec(Index).MouseRow, grdSpec(Index).MouseCol)
50        End If


60        Exit Sub

grdSpec_MouseMove_Error:

          Dim strES As String
          Dim intEL As Integer

70        Screen.MousePointer = 0

80        intEL = Erl
90        strES = Err.Description
100       LogError "frmEditHisto", "grdSpec_MouseMove", intEL, strES


End Sub



Private Sub imgLast_Click()

          Dim sql As String
          Dim tb As New Recordset
          Dim strDept As String
          Dim Yadd As Long


10        On Error GoTo imgLast_Click_Error

20        On Error GoTo imgLast_Click_Error

30        Yadd = Val(Swap_Year(Trim(Format(Now, "YYYY")))) * 1000

40        Select Case SSTab1.Tab
          Case 0:

50            bsave.Enabled = False
60            bsave.Enabled = False
70            cmdSaveHisto(0).Enabled = False
80            cmdSaveHHold(0).Enabled = False
90            cmdSaveHisto(1).Enabled = False
100           cmdSaveHHold(1).Enabled = False
110           cmdSaveCyto.Enabled = False
120           cmdSaveCHold.Enabled = False
130           If lblDisp = "H" Then
140               strDept = "Histo"
150           Else
160               strDept = "Cyto"
170           End If
180       Case 1: strDept = "Histo"
190       Case 2: strDept = "Histo"
200       Case 3: strDept = "Cyto"
210       End Select

220       sql = "SELECT top 1 SampleID, hYear from " & strDept & "Results "
230       If lblDisp = "H" Then
240           sql = sql & "WHERE sampleid < 40000000"
250       End If
260       sql = sql & "Order by hYear desc, SampleID desc"


270       Set tb = New Recordset
280       RecOpenServer 0, tb, sql
290       If Not tb.EOF Then
300           If lblDisp = "H" Then
310               txtSampleID = tb!SampleID - (SysOptHistoOffset(0) + Yadd)
320           Else
330               txtSampleID = tb!SampleID - (SysOptCytoOffset(0) + Yadd)
340           End If
350           txtYear = Trim(tb!Hyear)
360       End If



370       LoadAllDetails

380       bsave.Enabled = False
390       bsave.Enabled = False
400       cmdSaveHisto(0).Enabled = False
410       cmdSaveHHold(0).Enabled = False
420       cmdSaveHisto(1).Enabled = False
430       cmdSaveHHold(1).Enabled = False
440       cmdSaveCyto.Enabled = False
450       cmdSaveCHold.Enabled = False



460       Exit Sub

imgLast_Click_Error:

          Dim strES As String
          Dim intEL As Integer

470       intEL = Erl
480       strES = Err.Description
490       LogError "frmEditHisto", "imgLast_Click", intEL, strES, sql





End Sub

Private Sub iRecDate_Click(Index As Integer)

10        On Error GoTo iRecDate_Click_Error

20        If Index = 0 Then
30            dtRecDate = DateAdd("d", -1, dtRecDate)
40        Else
50            If DateDiff("d", dtRecDate, Now) > 0 Then
60                dtRecDate = DateAdd("d", 1, dtRecDate)
70            End If
80        End If

90        If dtRecDate < dtSampleDate Then
100           iMsg "Rundate less than sampledate!"
110           dtRecDate = dtSampleDate
120       End If

130       Exit Sub

iRecDate_Click_Error:

          Dim strES As String
          Dim intEL As Integer

140       Screen.MousePointer = 0

150       intEL = Erl
160       strES = Err.Description
170       LogError "frmEditHisto", "iRecDate_Click", intEL, strES


End Sub

Private Sub irelevant_Click(Index As Integer)

          Dim sql As String
          Dim tb As New Recordset
          Dim strDept As String
          Dim strDirection As String
          Dim strArrow As String
          Dim Yadd As Long



10        On Error GoTo irelevant_Click_Error

20        If bsave.Enabled = True Then
30            If iMsg("Not Saved. Do you wish to continue", vbYesNo) = vbNo Then
40                Exit Sub
50            End If
60        End If

70        Select Case SSTab1.Tab
          Case 0:

80            LoadAllDetails
90            bSaveHold.Enabled = False
100           bsave.Enabled = False
110           cmdSaveHisto(0).Enabled = False
120           cmdSaveHHold(0).Enabled = False
130           cmdSaveHisto(1).Enabled = False
140           cmdSaveHHold(1).Enabled = False
150           cmdSaveCyto.Enabled = False
160           cmdSaveCHold.Enabled = False
170           If lblDisp = "H" Then
180               strDept = "Histo"
190               Yadd = SysOptHistoOffset(0)
200           Else
210               strDept = "Cyto"
220               Yadd = SysOptCytoOffset(0)
230           End If
240       Case 1, 2:
250           strDept = "Histo"
260           Yadd = SysOptHistoOffset(0)
270       Case 3:
280           strDept = "Cyto"
290           Yadd = SysOptCytoOffset(0)
300       End Select

310       strDirection = IIf(Index = 0, "Desc", "Asc")
320       strArrow = IIf(Index = 0, "<", ">")

330       Yadd = Yadd + Val(Swap_Year(txtYear)) * 1000


340       If lblResultOrRequest = "Results" Then
350           sql = "SELECT top 1 SampleID from " & strDept & "Results WHERE " & _
                    "SampleID " & strArrow & " " & txtSampleID + Yadd & " and hyear = '" & txtYear & "' " & _
                    "Order by SampleID " & strDirection
360       Else
370           sql = "SELECT top 1 SampleID from " & strDept & "Requests WHERE " & _
                    "SampleID " & strArrow & " " & txtSampleID + Yadd & " and hyear = '" & txtYear & "' " & _
                    "Order by SampleID " & strDirection
380       End If



390       Set tb = New Recordset
400       RecOpenServer 0, tb, sql
410       If Not tb.EOF Then
420           txtSampleID = Val(tb!SampleID & "") - Yadd
430           If txtSampleID < 1 Then txtSampleID = 1
440       End If

450       LoadAllDetails


460       Exit Sub

irelevant_Click_Error:

          Dim strES As String
          Dim intEL As Integer

470       Screen.MousePointer = 0

480       intEL = Erl
490       strES = Err.Description
500       LogError "frmEditHisto", "irelevant_Click", intEL, strES


End Sub

Private Sub iRunDate_Click(Index As Integer)



10        On Error GoTo iRunDate_Click_Error

20        If Index = 0 Then
30            dtRunDate = DateAdd("d", -1, dtRunDate)
40        Else
50            If DateDiff("d", dtRunDate, Now) > 0 Then
60                dtRunDate = DateAdd("d", 1, dtRunDate)
70            End If
80        End If


90        If dtRunDate < dtSampleDate Then
100           iMsg "Rundate less than sampledate!"
110           dtRunDate = dtSampleDate
120       End If


130       bsave.Enabled = True
140       bSaveHold.Enabled = True




150       Exit Sub

iRunDate_Click_Error:

          Dim strES As String
          Dim intEL As Integer

160       Screen.MousePointer = 0

170       intEL = Erl
180       strES = Err.Description
190       LogError "frmEditHisto", "iRunDate_Click", intEL, strES


End Sub



Private Sub iSampleDate_Click(Index As Integer)

10        On Error GoTo iSampleDate_Click_Error

20        If Index = 0 Then
30            dtSampleDate = DateAdd("d", -1, dtSampleDate)
40        Else
50            If DateDiff("d", dtSampleDate, Now) > 0 Then
60                dtSampleDate = DateAdd("d", 1, dtSampleDate)
70            End If
80        End If


90        bsave.Enabled = True
100       bSaveHold.Enabled = True


110       Exit Sub

iSampleDate_Click_Error:

          Dim strES As String
          Dim intEL As Integer

120       Screen.MousePointer = 0

130       intEL = Erl
140       strES = Err.Description
150       LogError "frmEditHisto", "iSampleDate_Click", intEL, strES


End Sub

Private Sub iToday_Click(Index As Integer)



10        On Error GoTo iToday_Click_Error

20        If Index = 0 Then
30            dtRunDate = Format$(Now, "dd/mm/yyyy")
40        ElseIf Index = 1 Then
50            If DateDiff("d", dtRunDate, Now) > 0 Then
60                dtSampleDate = dtRunDate
70            Else
80                dtSampleDate = Format$(Now, "dd/mm/yyyy")
90            End If
100       ElseIf Index = 2 Then
110           If DateDiff("d", dtRunDate, Now) > 0 Then
120               dtRecDate = dtRunDate
130           Else
140               dtRecDate = Format$(Now, "dd/mm/yyyy")
150           End If
160       End If

170       bsave.Enabled = True
180       bSaveHold.Enabled = True



190       Exit Sub

iToday_Click_Error:

          Dim strES As String
          Dim intEL As Integer

200       Screen.MousePointer = 0

210       intEL = Erl
220       strES = Err.Description
230       LogError "frmEditHisto", "iToday_Click", intEL, strES


End Sub





Private Sub lblChartNumber_Click()

10        On Error GoTo lblChartNumber_Click_Error

20        With lblChartNumber
30            If InStr(.Caption, HospName(0)) = 0 Then
40                .BackColor = vbRed
50                .ForeColor = vbYellow
60            Else
70                .BackColor = &H8000000F
80                .ForeColor = vbBlack
90            End If

100       End With

110       If Trim$(txtChart) <> "" Then
120           LoadPatientFromChart Me, True
130           bsave.Enabled = True
140           bSaveHold.Enabled = True
150       End If

160       Exit Sub

lblChartNumber_Click_Error:

          Dim strES As String
          Dim intEL As Integer

170       Screen.MousePointer = 0

180       intEL = Erl
190       strES = Err.Description
200       LogError "frmEditHisto", "lblChartNumber_Click", intEL, strES


End Sub


Private Sub LoadAllDetails()

10        On Error GoTo LoadAllDetails_Error

20        HistoLoaded = False
30        CytoLoaded = False


40        ClearAll

50        LoadDemographics
60        CheckDepartments
70        LoadComments
80        Select Case SSTab1.Tab
          Case 0:
90            If lblDisp = "H" Then
100               LoadHisto
110               HistoLoaded = True
120           Else
130               LoadHisto
140               HistoLoaded = True
150           End If

160       Case 1: LoadHisto
170           HistoLoaded = True
180       Case 2: LoadHisto
190           HistoLoaded = True
200       Case 3: LoadCyto
210           CytoLoaded = True
220       End Select

230       SetViewHistory

240       cmdSaveHisto(0).Enabled = False
250       cmdSaveHHold(0).Enabled = False
260       cmdSaveHisto(1).Enabled = False
270       cmdSaveHHold(1).Enabled = False
280       cmdSaveCyto.Enabled = False
290       cmdSaveCHold.Enabled = False

300       Exit Sub

LoadAllDetails_Error:

          Dim strES As String
          Dim intEL As Integer

310       Screen.MousePointer = 0

320       intEL = Erl
330       strES = Err.Description
340       LogError "frmEditHisto", "LoadAllDetails", intEL, strES


End Sub
Private Sub LoadComments()

          Dim Ob As Observation
          Dim Obs As Observations
          Dim lngSampleID As Long
          Dim Yadd As Long

10        On Error GoTo LoadComments_Error

20        txtHistoComment = ""
30        txtDemographicComment = ""
40        lblDemographicComment = ""

50        If Trim$(txtSampleID) = "" Then Exit Sub

60        Yadd = Val(Swap_Year(txtYear)) * 1000

70        If lblDisp = "H" Then
80            lngSampleID = txtSampleID + SysOptHistoOffset(0) + Yadd
90        Else
100           lngSampleID = txtSampleID + SysOptCytoOffset(0) + Yadd
110       End If

120       Set Obs = New Observations
130       Set Obs = Obs.Load(lngSampleID, "Demographic", "Histology", "Cytology")
140       If Not Obs Is Nothing Then
150           For Each Ob In Obs
160               Select Case UCase$(Ob.Discipline)
                  Case "HISTOLOGY": txtHistoComment = Split_Comm(Ob.Comment)
170               Case "CYTOLOGY": txtCytoComment = Split_Comm(Ob.Comment)
180               Case "DEMOGRAPHIC": txtDemographicComment = Split_Comm(Ob.Comment)
190                   lblDemographicComment = txtDemographicComment
200               End Select
210           Next
220       End If

230       Exit Sub

LoadComments_Error:

          Dim strES As String
          Dim intEL As Integer

240       intEL = Erl
250       strES = Err.Description
260       LogError "frmEditHisto", "LoadComments", intEL, strES

End Sub

Public Sub LoadHisto()

          Dim Deltatb As Recordset
          Dim sql As String
          Dim Yadd As Long




10        On Error GoTo LoadHisto_Error

20        Yadd = Val(Swap_Year(txtYear)) * 1000


30        SSTab1.TabCaption(1) = "Histology Work Screen"
40        SSTab1.TabCaption(2) = "Histology Report"
50        chkNCRI(0).Value = 0

          'get date & run number of previous record
60        PreviousHisto = False
70        If txtChart <> "" Then
80            sql = "SELECT top 1 SampleID, RunDate from Demographics WHERE " & _
                    "Chart = '" & txtChart & "' " & _
                    "and SampleID < '" & txtSampleID + SysOptHistoOffset(0) + Yadd & "' " & _
                    "order by SampleID desc"
90            Set Deltatb = New Recordset
100           RecOpenServer 0, Deltatb, sql
110           If Not Deltatb.EOF Then
120               PreviousHisto = True
130           End If
140       End If

150       LoadRecord

160       LoadHistoWork


170       cmdSaveHisto(0).Enabled = False
180       cmdSaveHHold(0).Enabled = False
190       cmdSaveHisto(1).Enabled = False
200       cmdSaveHHold(1).Enabled = False





210       Exit Sub

LoadHisto_Error:

          Dim strES As String
          Dim intEL As Integer

220       Screen.MousePointer = 0

230       intEL = Erl
240       strES = Err.Description
250       LogError "frmEditHisto", "LoadHisto", intEL, strES, sql


End Sub

Private Sub LoadCyto()

          Dim Deltatb As Recordset
          Dim sql As String
          Dim Yadd

10        On Error GoTo LoadCyto_Error

20        Yadd = Val(Swap_Year(txtYear)) * 1000

30        StatusBar1.Panels(5).Text = ""

40        SSTab1.TabCaption(1) = "Histology Work Screen"
50        chkNCRI(1).Value = 0
          'get date & run number of previous record
60        PreviousHisto = False
70        If txtChart <> "" Then
80            sql = "SELECT top 1 demographics.SampleID, demographics.RunDate from Demographics, cytoresults WHERE " & _
                    "demographics.Chart = '" & txtChart & "' " & _
                    "and demographics.SampleID < '" & txtSampleID + SysOptCytoOffset(0) + Yadd & "' and demographics.sampleid = cytoresults.sampleid " & _
                    "order by demographics.SampleID desc"
90            Set Deltatb = New Recordset
100           RecOpenServer 0, Deltatb, sql
110           If Not Deltatb.EOF Then
120               PreviousCyto = True
130           End If
140       End If

150       loadCrecord


160       Exit Sub

LoadCyto_Error:

          Dim strES As String
          Dim intEL As Integer

170       intEL = Erl
180       strES = Err.Description
190       LogError "frmEditHisto", "LoadCyto", intEL, strES, sql


End Sub



Private Sub LoadDemographics()

          Dim sql As String
          Dim tb As New Recordset
          Dim SampleDate As String
          Dim Yadd As Long


10        On Error GoTo LoadDemographics_Error

20        If Trim$(txtSampleID) = "" Then Exit Sub

30        cmdDemoVal.Caption = "Validate"

40        Screen.MousePointer = 11

50        Yadd = Val(Swap_Year(txtYear)) * 1000

60        If lblDisp = "H" Then
70            sql = "SELECT * FROM Demographics WHERE " & _
                    "SampleID = '" & txtSampleID + SysOptHistoOffset(0) + Yadd & "'"
80        Else
90            sql = "SELECT * FROM Demographics WHERE " & _
                    "SampleID = '" & txtSampleID + SysOptCytoOffset(0) + Yadd & "'"
100       End If

110       Set tb = New Recordset
120       RecOpenServer 0, tb, sql
130       If tb.EOF Then
140           Set_Demo True
150           mNewRecord = True
160           dtRunDate = Format$(Now, "dd/mm/yyyy")
170           dtRecDate = Format$(Now, "dd/mm/yyyy")
180           dtSampleDate = Format$(Now, "dd/mm/yyyy")
190           txtChart = ""
200           txtName = ""
210           taddress(0) = ""
220           taddress(1) = ""
230           txtSex = ""
240           txtDoB = ""
250           txtAge = ""
260           cmbClinician = ""
270           cmbGP = ""
280           cmbWard = "GP"
290           cClDetails = ""
              'cCat.ListIndex = 0
300           txtDemographicComment = ""
310           tSampleTime.Mask = ""
320           tSampleTime.Text = ""
330           tSampleTime.Mask = "##:##"
340           tRecTime.Mask = ""
350           tRecTime.Text = ""
360           tRecTime.Mask = "##:##"
370           lblChartNumber.Caption = initial2upper(HospName(0)) & " Chart #"
380           lblChartNumber.BackColor = &H8000000F
390           lblChartNumber.ForeColor = vbBlack
400           txtAandE = ""
410           txtNOPAS = ""
420           lDoB = ""
430           lAge = ""
440           lSex = ""
450           cmbHospital = ""
460       Else
470           If Trim$(tb!Hospital & "") <> "" Then
480               lblChartNumber = Trim$(tb!Hospital) & " Chart #"
490               If UCase(tb!Hospital) = UCase(HospName(0)) Then
500                   lblChartNumber.BackColor = &H8000000F
510                   lblChartNumber.ForeColor = vbBlack
520               Else
530                   lblChartNumber.BackColor = vbRed
540                   lblChartNumber.ForeColor = vbYellow
550               End If
560           Else
570               lblChartNumber.Caption = initial2upper(HospName(0)) & " Chart #"
580               lblChartNumber.BackColor = &H8000000F
590               lblChartNumber.ForeColor = vbBlack
600           End If
610           If IsDate(tb!SampleDate) Then
620               dtSampleDate = Format$(tb!SampleDate, "dd/mm/yyyy")
630           Else
640               dtSampleDate = Format$(Now, "dd/mm/yyyy")
650           End If
660           If IsDate(tb!Rundate) Then
670               dtRunDate = Format$(tb!Rundate, "dd/mm/yyyy")
680           Else
690               dtRunDate = Format$(Now, "dd/mm/yyyy")
700           End If
710           StatusBar1.Panels(4).Text = dtRunDate
720           txtYear = Trim(tb!Hyear & "")
730           mNewRecord = False
740           If Trim$(tb!RooH & "") <> "" Then cRooH(0) = tb!RooH
750           If Trim$(tb!RooH & "") <> "" Then cRooH(1) = Not tb!RooH
760           txtChart = Trim(tb!Chart & "")
770           txtName = Trim(tb!PatName & "")
780           taddress(0) = Trim(tb!Addr0 & "")
790           taddress(1) = Trim(tb!Addr1 & "")
800           If Trim(tb!Hospital) & "" <> "" Then cmbHospital = tb!Hospital Else cmbHospital = initial2upper(HospName(0))
810           Select Case Left$(Trim$(UCase$(tb!sex & "")), 1)
              Case "M": txtSex = "Male"
820           Case "F": txtSex = "Female"
830           Case Else: txtSex = ""
840           End Select
850           lSex = txtSex
860           txtDoB = Format$(tb!Dob, "dd/mm/yyyy")
870           lAge = txtDoB
880           txtAge = tb!Age & ""
890           lAge = txtAge
900           cmbClinician = tb!Clinician & ""
910           cmbGP = tb!GP & ""
920           cmbWard = tb!Ward & ""
930           cClDetails = tb!ClDetails & ""
              '  If cCat <> "" Then cCat = tb!Category Else cCat = "Default"
940           If IsDate(tb!SampleDate) Then
950               dtSampleDate = Format$(tb!SampleDate, "dd/mm/yyyy")
960               If Format$(tb!SampleDate, "hh:mm") <> "00:00" Then
970                   tSampleTime = Format$(tb!SampleDate, "hh:mm")
980               Else
990                   tSampleTime.Mask = ""
1000                  tSampleTime.Text = ""
1010                  tSampleTime.Mask = "##:##"
1020              End If
1030          Else
1040              dtSampleDate = Format$(Now, "dd/mm/yyyy")
1050              tSampleTime.Mask = ""
1060              tSampleTime.Text = ""
1070              tSampleTime.Mask = "##:##"
1080          End If
1090          If IsDate(tb!RecDate & "") Then
1100              dtRecDate = Format$(tb!RecDate, "dd/mm/yyyy")
1110              If Format$(tb!RecDate, "hh:mm") <> "00:00" Then
1120                  tRecTime = Format$(tb!RecDate, "hh:mm")
1130              Else
1140                  tRecTime.Mask = ""
1150                  tRecTime.Text = ""
1160                  tRecTime.Mask = "##:##"
1170              End If
1180          Else
1190              dtRecDate = dtRunDate
1200              tRecTime.Mask = ""
1210              tRecTime.Text = ""
1220              tRecTime.Mask = "##:##"
1230          End If
1240          If SysOptDemoVal(0) = True Then
1250              If tb!Valid = True Then
1260                  cmdDemoVal.Caption = "VALID"
1270                  Set_Demo False
1280              Else
1290                  cmdDemoVal.Caption = "Validate"
1300                  Set_Demo True
1310              End If
1320          End If
1330      End If
1340      bsave.Enabled = False
1350      bSaveHold.Enabled = False

1360      CheckCC

1370      Screen.MousePointer = 0



1380      Exit Sub

LoadDemographics_Error:

          Dim strES As String
          Dim intEL As Integer

1390      intEL = Erl
1400      strES = Err.Description
1410      LogError "frmEditHisto", "LoadDemographics", intEL, strES, sql


End Sub

Public Property Let PrintToPrinter(ByVal strNewValue As String)
Attribute PrintToPrinter.VB_HelpID = 580

10        On Error GoTo PrintToPrinter_Error

20        pPrintToPrinter = strNewValue

30        Exit Property

PrintToPrinter_Error:

          Dim strES As String
          Dim intEL As Integer

40        Screen.MousePointer = 0

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmEditHisto", "PrintToPrinter", intEL, strES


End Property

Public Property Get PrintToPrinter() As String

10        On Error GoTo PrintToPrinter_Error

20        PrintToPrinter = pPrintToPrinter

30        Exit Property

PrintToPrinter_Error:

          Dim strES As String
          Dim intEL As Integer

40        Screen.MousePointer = 0

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmEditHisto", "PrintToPrinter", intEL, strES


End Property

Private Sub SaveDemographics()

          Dim sql As String
          Dim Hosp As String
          Dim Yadd As Long

10        On Error GoTo SaveDemographics_Error

20        Yadd = Val(Swap_Year(txtYear)) * 1000


30        If Trim$(tSampleTime) <> "__:__" Then
40            If Not IsDate(tSampleTime) Then
50                iMsg "Invalid Time", vbExclamation
60                Exit Sub
70            End If
80        End If

90        Hosp = HospName(0)

          'Created on 08/10/2010 15:11:40
          'Autogenerated by SQL Scripting

100       sql = "If Exists(Select 1 From Demographics " & _
                "Where SampleID = @SampleID0 ) " & _
                "Begin " & _
                "Update Demographics Set " & _
                "PatName = '@PatName2', Age = '@Age3', Sex = '@Sex4', RunDate = '@RunDate7', DoB = '@DoB8', Addr0 = '@Addr09', Addr1 = '@Addr110', Ward = '@Ward11', Clinician = '@Clinician12', GP = '@GP13', SampleDate = '@SampleDate14', ClDetails = '@ClDetails20', Hospital = '@Hospital21', RooH = @RooH23, Chart = '@Chart33', RecDate = '@RecDate38', HYear = '@HYear45' " & _
                "Where SampleID = @SampleID0  " & _
                "End  " & _
                "Else " & _
                "Begin  " & _
                "Insert Into Demographics (SampleID, PatName, Age, Sex, RunDate, DoB, Addr0, Addr1, Ward, Clinician, GP, SampleDate, ClDetails, Hospital, RooH, Chart, RecDate, HYear) Values (@SampleID0, '@PatName2', '@Age3', '@Sex4', '@RunDate7', '@DoB8', '@Addr09', '@Addr110', '@Ward11', '@Clinician12', '@GP13', '@SampleDate14', '@ClDetails20', '@Hospital21', @RooH23, '@Chart33', '@RecDate38', '@HYear45') " & _
                "End"

110       If lblDisp = "H" Then
120           sql = Replace(sql, "@SampleID0", txtSampleID + SysOptHistoOffset(0) + Yadd)
130       Else
140           sql = Replace(sql, "@SampleID0", txtSampleID + SysOptCytoOffset(0) + Yadd)
150       End If
160       sql = Replace(sql, "@PatName2", AddTicks(Trim$(txtName)))
170       sql = Replace(sql, "@Age3", txtAge)
180       sql = Replace(sql, "@Sex4", Left$(txtSex, 1))
190       sql = Replace(sql, "@RunDate7", Format$(dtRunDate, "dd/mmm/yyyy"))
200       If IsDate(txtDoB) Then
210           sql = Replace(sql, "@DoB8", Format$(txtDoB, "dd/mmm/yyyy"))
220       Else
230           sql = Replace(sql, "'@DoB8'", "NULL")
240       End If
250       sql = Replace(sql, "@Addr09", AddTicks(taddress(0)))
260       sql = Replace(sql, "@Addr110", AddTicks(taddress(1)))
270       sql = Replace(sql, "@Ward11", AddTicks(StrConv(Left$(cmbWard, 50), vbProperCase)))
280       sql = Replace(sql, "@Clinician12", AddTicks(Left$(cmbClinician, 50)))
290       sql = Replace(sql, "@GP13", AddTicks(Left$(cmbGP, 50)))
300       If IsDate(tSampleTime) Then
310           sql = Replace(sql, "@SampleDate14", Format$(dtSampleDate, "dd/mmm/yyyy") & " " & Format$(tSampleTime, "hh:mm"))
320       Else
330           sql = Replace(sql, "@SampleDate14", Format$(dtSampleDate, "dd/mmm/yyyy"))
340       End If
350       sql = Replace(sql, "@ClDetails20", AddTicks(Left$(cClDetails, 50)))
360       sql = Replace(sql, "@Hospital21", AddTicks(cmbHospital))
370       sql = Replace(sql, "@RooH23", IIf(cRooH(0), 1, 0))
380       sql = Replace(sql, "@Chart33", txtChart)
390       If IsDate(tRecTime) Then
400           If Format$(dtRecDate, "yyyy/mm/dd") <= Format$(dtSampleDate, "yyyy/mm/dd") Then
410               sql = Replace(sql, "@RecDate38", Format$(dtSampleDate, "dd/mmm/yyyy") & " " & Format$(tRecTime, "hh:mm"))
420           Else
430               sql = Replace(sql, "@RecDate38", Format$(dtRecDate, "dd/mmm/yyyy") & " " & Format$(tRecTime, "hh:mm"))
440           End If
450       Else
460           sql = Replace(sql, "@RecDate38", Format$(dtRecDate, "dd/mmm/yyyy") & " " & Format$(Now, "hh:mm"))
470       End If

480       sql = Replace(sql, "@HYear45", txtYear)

490       Cnxn(0).Execute sql



          'If lblDisp = "H" Then
          '    sql = "SELECT * FROM Demographics WHERE " & _
               '          "SampleID = '" & txtSampleID + SysOptHistoOffset(0) + Yadd & "'"
          'Else
          '    sql = "SELECT * FROM Demographics WHERE " & _
               '          "SampleID = '" & txtSampleID + SysOptCytoOffset(0) + Yadd & "'"
          'End If
          'Set tb = New Recordset
          'RecOpenServer 0, tb, sql
          'If tb.EOF Then
          '    tb.AddNew
          '    If lblDisp = "H" Then
          '        tb!SampleID = txtSampleID + SysOptHistoOffset(0) + Yadd
          '    End If
          '    If lblDisp = "C" Then
          '        tb!SampleID = txtSampleID + SysOptCytoOffset(0) + Yadd
          '    End If
          'End If
          '
          'tb!RooH = cRooH(0)
          '
          'tb!Rundate = Format$(dtRunDate, "dd/mmm/yyyy")
          'If IsDate(tSampleTime) Then
          '    tb!SampleDate = Format$(dtSampleDate, "dd/mmm/yyyy") & " " & Format$(tSampleTime, "hh:mm")
          'Else
          '    tb!SampleDate = Format$(dtSampleDate, "dd/mmm/yyyy")
          'End If
          'If IsDate(tRecTime) Then
          '    If Format$(dtRecDate, "yyyy/mm/dd") <= Format$(dtSampleDate, "yyyy/mm/dd") Then
          '        tb!RecDate = Format$(dtSampleDate, "dd/mmm/yyyy") & " " & Format$(tRecTime, "hh:mm")
          '    Else
          '        tb!RecDate = Format$(dtRecDate, "dd/mmm/yyyy") & " " & Format$(tRecTime, "hh:mm")
          '    End If
          'Else
          '    tb!RecDate = Format$(dtRecDate, "dd/mmm/yyyy") & " " & Format$(Now, "hh:mm")
          'End If
          '
          'tb!Chart = txtChart
          'tb!PatName = Trim$(txtName)
          'If IsDate(txtDoB) Then
          '    tb!Dob = Format$(txtDoB, "dd/mmm/yyyy")
          'Else
          '    tb!Dob = Null
          'End If
          'tb!Hyear = txtYear
          ''If cCat = "Default" Then tb!Category = "" Else tb!Category = cCat
          'tb!Age = txtAge
          'tb!sex = Left$(txtSex, 1)
          'tb!Addr0 = tAddress(0)
          'tb!Addr1 = tAddress(1)
          'tb!Ward = StrConv(Left$(cmbWard, 50), vbProperCase)
          'tb!Clinician = Left$(cmbClinician, 50)
          'tb!GP = Left$(cmbGP, 50)
          'tb!ClDetails = Left$(cClDetails, 50)
          'tb!Hospital = cmbHospital
          'tb.Update

500       LogTimeOfPrinting txtSampleID, "D"

510       Screen.MousePointer = 0



520       Exit Sub

SaveDemographics_Error:

          Dim strES As String
          Dim intEL As Integer

530       intEL = Erl
540       strES = Err.Description
550       LogError "frmEditHisto", "SaveDemographics", intEL, strES, sql


End Sub

Private Sub SetViewHistory()

'Select Case SSTab1.Tab
'  Case 0: bHistory.Visible = False
'  Case 1: bHistory.Visible = PreviousHisto
'  Case 2: bHistory.Visible = PreviousCyto
'End Select

End Sub

Private Sub lblDisp_Click()

10        On Error GoTo lblDisp_Click_Error

20        If lblDisp = "H" Then
30            lblDisp = "C"
40            SSTab1.TabVisible(1) = False
50            SSTab1.TabVisible(2) = False
60            SSTab1.TabVisible(3) = True
70        Else
80            lblDisp = "H"
90            SSTab1.TabVisible(1) = True
100           SSTab1.TabVisible(2) = True
110           SSTab1.TabVisible(3) = False
120       End If

130       LoadAllDetails

140       Exit Sub

lblDisp_Click_Error:

          Dim strES As String
          Dim intEL As Integer

150       Screen.MousePointer = 0

160       intEL = Erl
170       strES = Err.Description
180       LogError "frmEditHisto", "lblDisp_Click", intEL, strES


End Sub

Private Sub lblStatus_Click()

10        On Error GoTo lblStatus_Click_Error

20        If lblStatus = "Cut Up" Then
30            lblStatus = "Stain"
40        ElseIf lblStatus = "Stain" Then
50            lblStatus = "Specials"
60        ElseIf lblStatus = "Specials" Then
70            lblStatus = "Completed"
80        ElseIf lblStatus = "Completed" Then
90            lblStatus = ""
100       ElseIf lblStatus = "" Then
110           lblStatus = "Cut Up"
120       End If

130       cmdSaveHisto(0).Enabled = True
140       cmdSaveHHold(0).Enabled = True
150       cmdSaveHisto(1).Enabled = True
160       cmdSaveHHold(1).Enabled = True

170       Exit Sub

lblStatus_Click_Error:

          Dim strES As String
          Dim intEL As Integer

180       Screen.MousePointer = 0

190       intEL = Erl
200       strES = Err.Description
210       LogError "frmEditHisto", "lblStatus_Click", intEL, strES


End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)

      'SELECT Case PreviousTab
      '  Case 0
      '    If bSave.Enabled Then
      '      If iMsg("Demographic Details have changed!" & vbCrLf & "Save?", vbQuestion + vbYesNo) = vbYes Then
      '        bSave_Click
      '      End If
      '    End If
      '  Case 1
      '    If bSaveHisto.Enabled Then
      '      If iMsg("Histology Details have changed!" & vbCrLf & "Save?", vbQuestion + vbYesNo) = vbYes Then
      '        bSaveHisto_Click
      '      End If
      '    End If
      '  Case 2
      '    If bSaveCyto.Enabled Then
      '      If iMsg("Cytology Details have changed!" & vbCrLf & "Save?", vbQuestion + vbYesNo) = vbYes Then
      '        bSaveCyto_Click
      '      End If
      '    End If
      'End SELECT

10        On Error GoTo SSTab1_Click_Error

20        Select Case SSTab1.Tab
          Case 0:    'Demographics
30            If bsave.Enabled Then
40                If iMsg("Demographic Details have changed!" & vbCrLf & "Save?", vbQuestion + vbYesNo) = vbYes Then
50                    bSave_Click
60                End If
70            End If

80        Case 1:    'Histology
90            If cmdSaveHisto(0).Enabled Then
100               If iMsg("Histology Details have changed!" & vbCrLf & "Save?", vbQuestion + vbYesNo) = vbYes Then
110                   cmdSaveHHold_Click 0
120               End If
130           End If
140           If Not HistoLoaded Then
150               LoadHisto
160               HistoLoaded = True
170           End If

180       Case 2:    'Histology
190           If cmdSaveHisto(0).Enabled Then
200               If iMsg("Histology Details have changed!" & vbCrLf & "Save?", vbQuestion + vbYesNo) = vbYes Then
210                   cmdSaveHHold_Click 0
220               End If
230           End If
240           If Not HistoLoaded Then
250               LoadHisto
260               HistoLoaded = True
270           End If

280       Case 3:    'Cytology
290           If cmdSaveCyto.Enabled Then
300               If iMsg("Cytology Details have changed!" & vbCrLf & "Save?", vbQuestion + vbYesNo) = vbYes Then
310                   cmdSaveCHold_Click
320               End If
330           End If
340           If Not CytoLoaded Then
350               LoadCyto
360               CytoLoaded = True
370           End If


380       End Select


390       bPrintHold.Visible = True
400       bPrint.Visible = True

410       Select Case SSTab1.Tab
          Case 0
420           bPrintHold.Visible = False
430           bPrint.Visible = False
440       End Select


450       SetViewHistory

460       Exit Sub

SSTab1_Click_Error:

          Dim strES As String
          Dim intEL As Integer

470       Screen.MousePointer = 0

480       intEL = Erl
490       strES = Err.Description
500       LogError "frmEditHisto", "SSTab1_Click", intEL, strES


End Sub

Private Sub SSTab1_KeyDown(KeyCode As Integer, Shift As Integer)
          Dim s As String

10        On Error GoTo SSTab1_KeyDown_Error

20        If SSTab1.Tab = 2 Then
30            If KeyCode = vbKeyD Then
40                If Shift And vbAltMask Then
50                    s = Format(Now, "dd/mm/yyyy")
60                    T.SelText = s
70                    KeyCode = 0
80                End If
90            End If
100           If KeyCode = vbKeyW Then
110               If Shift And vbAltMask Then
120                   T.SetFocus
130                   KeyCode = 0
140               End If
150           End If
160       ElseIf SSTab1.Tab = 3 Then
170           If KeyCode = vbKeyD Then
180               If Shift And vbAltMask Then
190                   s = Format(Now, "dd/mm/yyyy")
200                   txtCyto.SelText = s
210                   KeyCode = 0
220               End If
230           End If
240           If KeyCode = vbKeyW Then
250               If Shift And vbAltMask Then
260                   txtCyto.SetFocus
270                   KeyCode = 0
280               End If
290           End If
300       End If

310       Exit Sub

SSTab1_KeyDown_Error:

          Dim strES As String
          Dim intEL As Integer

320       Screen.MousePointer = 0

330       intEL = Erl
340       strES = Err.Description
350       LogError "frmEditHisto", "SSTab1_KeyDown", intEL, strES

End Sub

Private Sub SSTab1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        On Error GoTo SSTab1_MouseMove_Error

20        pBar = 0

30        Exit Sub

SSTab1_MouseMove_Error:

          Dim strES As String
          Dim intEL As Integer

40        Screen.MousePointer = 0

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmEditHisto", "SSTab1_MouseMove", intEL, strES


End Sub



Private Sub T_KeyDown(KeyCode As Integer, Shift As Integer)
          Dim s As String

10        On Error GoTo T_KeyDown_Error

20        If T.Locked = True Then Exit Sub
30        pBar = 0
40        If KeyCode = vbKeyD Then
50            If Shift And vbAltMask Then
60                s = Format(Now, "dd/mm/yyyy")
70                T.SelText = s
80                KeyCode = 0
90            End If
100       End If


110       Exit Sub

T_KeyDown_Error:

          Dim strES As String
          Dim intEL As Integer

120       Screen.MousePointer = 0

130       intEL = Erl
140       strES = Err.Description
150       LogError "frmEditHisto", "T_KeyDown", intEL, strES


End Sub

Private Sub t_KeyUp(KeyCode As Integer, Shift As Integer)

10        On Error GoTo t_KeyUp_Error

20        If T.Locked = True Then Exit Sub
30        pBar = 0
40        cmdSaveHisto(0).Enabled = True
50        cmdSaveHHold(0).Enabled = True
60        cmdSaveHisto(1).Enabled = True
70        cmdSaveHHold(1).Enabled = True

80        Exit Sub

t_KeyUp_Error:

          Dim strES As String
          Dim intEL As Integer

90        Screen.MousePointer = 0

100       intEL = Erl
110       strES = Err.Description
120       LogError "frmEditHisto", "t_KeyUp", intEL, strES


End Sub

Private Sub taddress_Change(Index As Integer)

10        On Error GoTo taddress_Change_Error

20        SetWardClinGP

30        Exit Sub

taddress_Change_Error:

          Dim strES As String
          Dim intEL As Integer

40        Screen.MousePointer = 0

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmEditHisto", "taddress_Change", intEL, strES


End Sub

Private Sub taddress_KeyPress(Index As Integer, KeyAscii As Integer)
10        On Error GoTo taddress_KeyPress_Error

20        pBar = 0
30        bsave.Enabled = True
40        bSaveHold.Enabled = True

50        Exit Sub

taddress_KeyPress_Error:

          Dim strES As String
          Dim intEL As Integer

60        Screen.MousePointer = 0

70        intEL = Erl
80        strES = Err.Description
90        LogError "frmEditHisto", "taddress_KeyPress", intEL, strES


End Sub

Private Sub taddress_LostFocus(Index As Integer)

10        On Error GoTo taddress_LostFocus_Error

20        taddress(Index) = StrConv(taddress(Index), vbProperCase)

30        Exit Sub

taddress_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

40        Screen.MousePointer = 0

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmEditHisto", "taddress_LostFocus", intEL, strES


End Sub

Private Sub txtage_Change()

10        On Error GoTo txtage_Change_Error

20        lAge = txtAge

30        Exit Sub

txtage_Change_Error:

          Dim strES As String
          Dim intEL As Integer

40        Screen.MousePointer = 0

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmEditHisto", "txtage_Change", intEL, strES


End Sub

Private Sub txtAge_KeyPress(KeyAscii As Integer)

10        On Error GoTo txtAge_KeyPress_Error

20        bsave.Enabled = True
30        bSaveHold.Enabled = True

40        Exit Sub

txtAge_KeyPress_Error:

          Dim strES As String
          Dim intEL As Integer

50        Screen.MousePointer = 0

60        intEL = Erl
70        strES = Err.Description
80        LogError "frmEditHisto", "txtAge_KeyPress", intEL, strES


End Sub


Private Sub txtchart_Change()

10        On Error GoTo txtchart_Change_Error

20        lChart = txtChart

30        Exit Sub

txtchart_Change_Error:

          Dim strES As String
          Dim intEL As Integer

40        Screen.MousePointer = 0

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmEditHisto", "txtchart_Change", intEL, strES


End Sub

Private Sub txtChart_KeyPress(KeyAscii As Integer)

10        On Error GoTo txtChart_KeyPress_Error

20        'If txtChart.Locked Then Exit Sub

30        bsave.Enabled = True
40        bSaveHold.Enabled = True

50        Exit Sub

txtChart_KeyPress_Error:

          Dim strES As String
          Dim intEL As Integer

60        Screen.MousePointer = 0

70        intEL = Erl
80        strES = Err.Description
90        LogError "frmEditHisto", "txtChart_KeyPress", intEL, strES


End Sub


Private Sub txtchart_LostFocus()

10        On Error GoTo txtchart_LostFocus_Error

20        If Trim$(txtChart) = "" Then Exit Sub
30        If Trim$(txtName) <> "" Then Exit Sub

40        LoadPatientFromChart Me, True

50        If txtName = "" Then
60            LoadDemo txtChart
70        End If

80        Exit Sub

txtchart_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

90        Screen.MousePointer = 0

100       intEL = Erl
110       strES = Err.Description
120       LogError "frmEditHisto", "txtchart_LostFocus", intEL, strES


End Sub


Private Sub txtCyto_KeyDown(KeyCode As Integer, Shift As Integer)

          Dim s As String
10        On Error GoTo txtCyto_KeyDown_Error

20        If txtCyto.Locked = True Then Exit Sub
30        pBar = 0
40        If KeyCode = vbKeyD Then
50            If Shift And vbAltMask Then
60                s = Format(Now, "dd/mm/yyyy")
70                txtCyto.SelText = s
80                KeyCode = 0
90            End If
100       End If

110       Exit Sub

txtCyto_KeyDown_Error:

          Dim strES As String
          Dim intEL As Integer

120       Screen.MousePointer = 0

130       intEL = Erl
140       strES = Err.Description
150       LogError "frmEditHisto", "txtCyto_KeyDown", intEL, strES


End Sub



Private Sub txtCytoComment_Change()

10        On Error GoTo txtCytoComment_Change_Error

20        pBar = 0

30        Exit Sub

txtCytoComment_Change_Error:

          Dim strES As String
          Dim intEL As Integer

40        Screen.MousePointer = 0

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmEditHisto", "txtCytoComment_Change", intEL, strES


End Sub

Private Sub txtCytoComment_KeyDown(KeyCode As Integer, Shift As Integer)


          Dim s As Variant
          Dim n As Long
          Dim z As Long
          Dim tb As New Recordset
          Dim sql As String


10        On Error GoTo txtCytoComment_KeyDown_Error

20        pBar = 0

30        If KeyCode = 113 Then

40            n = txtCytoComment.SelStart

50            z = 2
60            s = Mid(txtCytoComment, n - z, z + 1)
70            z = 3
80            If ListText("CI", s) <> "" Then
90                s = ListText("CI", s)
100           Else
110               s = ""
120           End If

130           If s = "" Then
140               z = 1
150               s = Mid(txtCytoComment, n - z, z + 1)
160               z = 2
170               If ListText("CI", s) <> "" Then
180                   s = ListText("CI", s)
190               Else
200                   s = ""
210               End If
220           End If

230           If s = "" Then
240               z = 1
250               s = Mid(txtCytoComment, n, z)

260               If ListText("CI", s) <> "" Then
270                   s = ListText("CI", s)
280               End If
290           End If

300           txtCytoComment = Left(txtCytoComment, n - (z))
310           txtCytoComment = txtCytoComment & s

320           txtCytoComment.SelStart = Len(txtCytoComment)

330       ElseIf KeyCode = 114 Then

340           sql = "SELECT * from lists WHERE listtype = 'CI'"
350           Set tb = New Recordset
360           RecOpenServer 0, tb, sql
370           Do While Not tb.EOF
380               s = Trim(tb!Text)
390               frmMessages.lstComm.AddItem s
400               tb.MoveNext
410           Loop

420           Set frmMessages.f = Me
430           Set frmMessages.T = txtCytoComment
440           frmMessages.Show 1

450       End If

460       If txtCytoComment.Locked = False Then
470           cmdSaveCyto.Enabled = True
480           cmdSaveCHold.Enabled = True
490       End If



500       Exit Sub

txtCytoComment_KeyDown_Error:

          Dim strES As String
          Dim intEL As Integer

510       Screen.MousePointer = 0

520       intEL = Erl
530       strES = Err.Description
540       LogError "frmEditHisto", "txtCytoComment_KeyDown", intEL, strES


End Sub

Private Sub txtDemographicComment_KeyDown(KeyCode As Integer, Shift As Integer)

          Dim s As Variant
          Dim n As Long
          Dim sql As String
          Dim tb As New Recordset
          Dim z As Long

10        On Error GoTo txtDemographicComment_KeyDown_Error

20        If KeyCode = 113 Then

30            n = txtDemographicComment.SelStart

40            z = 2
50            s = Mid(txtDemographicComment, n - z, z + 1)
60            z = 3
70            If ListText("DE", s) <> "" Then
80                s = ListText("DE", s)
90            Else
100               s = ""
110           End If

120           If s = "" Then
130               z = 1
140               s = Mid(txtDemographicComment, n - z, z + 1)
150               z = 2
160               If ListText("DE", s) <> "" Then
170                   s = ListText("DE", s)
180               Else
190                   s = ""
200               End If
210           End If

220           If s = "" Then
230               z = 1
240               s = Mid(txtDemographicComment, n, z)

250               If ListText("DE", s) <> "" Then
260                   s = ListText("DE", s)
270               End If
280           End If

290           txtDemographicComment = Left(txtDemographicComment, n - (z))
300           txtDemographicComment = txtDemographicComment & s

310           txtDemographicComment.SelStart = Len(txtDemographicComment)

320           bsave.Enabled = True
330           bSaveHold.Enabled = True

340       ElseIf KeyCode = 114 Then

350           sql = "SELECT * from lists WHERE listtype = 'DE'"
360           Set tb = New Recordset
370           RecOpenServer 0, tb, sql
380           Do While Not tb.EOF
390               s = Trim(tb!Text)
400               frmMessages.lstComm.AddItem s
410               tb.MoveNext
420           Loop

430           Set frmMessages.f = frmEditAll
440           Set frmMessages.T = txtDemographicComment
450           frmMessages.Show 1

460           bsave.Enabled = True
470           bSaveHold.Enabled = True

480       End If

490       Exit Sub

txtDemographicComment_KeyDown_Error:

          Dim strES As String
          Dim intEL As Integer

500       Screen.MousePointer = 0

510       intEL = Erl
520       strES = Err.Description
530       LogError "frmEditHisto", "txtDemographicComment_KeyDown", intEL, strES

End Sub


Private Sub txtDoB_Change()

10        On Error GoTo txtDoB_Change_Error

20        lDoB = txtDoB

30        Exit Sub

txtDoB_Change_Error:

          Dim strES As String
          Dim intEL As Integer

40        Screen.MousePointer = 0

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmEditHisto", "txtDoB_Change", intEL, strES


End Sub

Private Sub txtDoB_KeyPress(KeyAscii As Integer)

10        On Error GoTo txtDoB_KeyPress_Error

20        If txtDoB.Locked Then Exit Sub

30        bsave.Enabled = True
40        bSaveHold.Enabled = True

50        Exit Sub

txtDoB_KeyPress_Error:

          Dim strES As String
          Dim intEL As Integer

60        Screen.MousePointer = 0

70        intEL = Erl
80        strES = Err.Description
90        LogError "frmEditHisto", "txtDoB_KeyPress", intEL, strES


End Sub

Private Sub txtDoB_LostFocus()

10        On Error GoTo txtDoB_LostFocus_Error

20        If txtDoB.Locked Then Exit Sub

30        txtDoB = Convert62Date(txtDoB, BACKWARD)
40        txtDoB = Format(txtDoB, "dd/MM/yyyy")
50        txtAge = CalcAge(txtDoB, dtSampleDate)

60        Exit Sub

txtDoB_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

70        Screen.MousePointer = 0

80        intEL = Erl
90        strES = Err.Description
100       LogError "frmEditHisto", "txtDoB_LostFocus", intEL, strES


End Sub

Private Sub TimerBar_Timer()

10        On Error GoTo TimerBar_Timer_Error

20        pBar = pBar + 1

30        If pBar = pBar.Max Then
40            Unload Me
50            Exit Sub
60        End If

70        Exit Sub

TimerBar_Timer_Error:

          Dim strES As String
          Dim intEL As Integer

80        Screen.MousePointer = 0

90        intEL = Erl
100       strES = Err.Description
110       LogError "frmEditHisto", "TimerBar_Timer", intEL, strES


End Sub


Private Sub tInput_Change(Index As Integer)

10        On Error GoTo tInput_Change_Error

20        If InStr(grdSpec(Index).TextMatrix(grdSpec(Index).RowSel, grdSpec(Index).ColSel), "Pieces") Then Exit Sub
30        If InStr(grdSpec(Index).TextMatrix(grdSpec(Index).RowSel, grdSpec(Index).ColSel), "Blk") Then Exit Sub

40        tinput(Index).SelStart = Len(tinput(Index))

50        grdSpec(Index).TextMatrix(grdSpec(Index).RowSel, grdSpec(Index).ColSel) = Trim(tinput(Index))

60        Exit Sub

tInput_Change_Error:

          Dim strES As String
          Dim intEL As Integer

70        Screen.MousePointer = 0

80        intEL = Erl
90        strES = Err.Description
100       LogError "frmEditHisto", "tInput_Change", intEL, strES


End Sub

Private Sub tInput_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
          Dim Comm As String

10        On Error GoTo tInput_KeyDown_Error

20        Comm = ""

30        If KeyCode = 113 Then

40            Comm = iBOX("Comment!")
50            If Comm <> "" Then
60                grdComm(Index).TextMatrix(grdSpec(Index).RowSel, grdSpec(Index).ColSel) = Comm
70                grdSpec(Index).Col = grdSpec(Index).ColSel
80                grdSpec(Index).Row = grdSpec(Index).RowSel
90                grdSpec(Index).CellBackColor = vbYellow
100           End If
110       ElseIf KeyCode = 117 Then
120           If Index = 6 Then
130               tinput(0).SetFocus
140           Else
150               If fraSpec(Index + 1).Visible = True Then
160                   tinput(Index + 1).SetFocus
170               End If
180           End If
190       ElseIf KeyCode = 118 Then
200           If Index = 0 Then
210               tinput(0).SetFocus
220           Else
230               If fraSpec(Index - 1).Visible = True Then
240                   tinput(Index - 1).SetFocus
250               End If
260           End If
270       End If

280       cmdSaveHisto(0).Enabled = True
290       cmdSaveHHold(0).Enabled = True
300       cmdSaveHisto(1).Enabled = True
310       cmdSaveHHold(1).Enabled = True

320       Exit Sub

tInput_KeyDown_Error:

          Dim strES As String
          Dim intEL As Integer

330       Screen.MousePointer = 0

340       intEL = Erl
350       strES = Err.Description
360       LogError "frmEditHisto", "tInput_KeyDown", intEL, strES


End Sub



Private Sub txtHist_KeyPress(Index As Integer, KeyAscii As Integer)


10        On Error GoTo txtHist_KeyPress_Error

20        cmdSaveHisto(0).Enabled = True
30        cmdSaveHHold(0).Enabled = True
40        cmdSaveHisto(1).Enabled = True
50        cmdSaveHHold(1).Enabled = True

60        Exit Sub

txtHist_KeyPress_Error:

          Dim strES As String
          Dim intEL As Integer

70        Screen.MousePointer = 0

80        intEL = Erl
90        strES = Err.Description
100       LogError "frmEditHisto", "txtHist_KeyPress", intEL, strES


End Sub

Private Sub txtName_Change()

10        On Error GoTo txtName_Change_Error

20        lName = txtName

30        Exit Sub

txtName_Change_Error:

          Dim strES As String
          Dim intEL As Integer

40        Screen.MousePointer = 0

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmEditHisto", "txtName_Change", intEL, strES


End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)

10        On Error GoTo txtName_KeyPress_Error

20        If txtName.Locked Then Exit Sub

30        bsave.Enabled = True
40        bSaveHold.Enabled = True

50        Exit Sub

txtName_KeyPress_Error:

          Dim strES As String
          Dim intEL As Integer

60        Screen.MousePointer = 0

70        intEL = Erl
80        strES = Err.Description
90        LogError "frmEditHisto", "txtName_KeyPress", intEL, strES


End Sub

Private Sub txtname_LostFocus()


          Dim strName As String
          Dim strSex As String

10        On Error GoTo txtname_LostFocus_Error

20        strName = txtName
30        strSex = txtSex

40        NameLostFocus strName, strSex

50        txtName = strName
60        txtSex = strSex

70        Exit Sub

txtname_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

80        Screen.MousePointer = 0

90        intEL = Erl
100       strES = Err.Description
110       LogError "frmEditHisto", "txtname_LostFocus", intEL, strES


End Sub


Private Sub tSampleTime_KeyPress(KeyAscii As Integer)

10        On Error GoTo tSampleTime_KeyPress_Error

20        bsave.Enabled = True
30        bSaveHold.Enabled = True

40        Exit Sub

tSampleTime_KeyPress_Error:

          Dim strES As String
          Dim intEL As Integer

50        Screen.MousePointer = 0

60        intEL = Erl
70        strES = Err.Description
80        LogError "frmEditHisto", "tSampleTime_KeyPress", intEL, strES


End Sub

Private Sub txtSampleID_GotFocus()

10        On Error GoTo txtSampleID_GotFocus_Error

20        txtSampleID.SelStart = Len(txtSampleID)

30        Exit Sub

txtSampleID_GotFocus_Error:

          Dim strES As String
          Dim intEL As Integer

40        Screen.MousePointer = 0

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmEditHisto", "txtSampleID_GotFocus", intEL, strES


End Sub

Private Sub txtSampleID_KeyPress(KeyAscii As Integer)

10        On Error GoTo txtSampleID_KeyPress_Error

20        KeyAscii = VI(KeyAscii, Numeric_Only)

30        Exit Sub

txtSampleID_KeyPress_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmEditHisto", "txtSampleID_KeyPress", intEL, strES


End Sub

Private Sub txtSex_Change()

10        On Error GoTo txtSex_Change_Error

20        If txtDoB.Locked Then Exit Sub

30        lSex = txtSex

40        Exit Sub

txtSex_Change_Error:

          Dim strES As String
          Dim intEL As Integer

50        Screen.MousePointer = 0

60        intEL = Erl
70        strES = Err.Description
80        LogError "frmEditHisto", "txtSex_Change", intEL, strES


End Sub

Private Sub txtsex_Click()

10        On Error GoTo txtsex_Click_Error

20        If txtSex.Locked Then Exit Sub

30        Select Case Trim$(txtSex)
          Case "": txtSex = "Male"
40        Case "Male": txtSex = "Female"
50        Case "Female": txtSex = ""
60        Case Else: txtSex = ""
70        End Select

80        bsave.Enabled = True
90        bSaveHold.Enabled = True

100       Exit Sub

txtsex_Click_Error:

          Dim strES As String
          Dim intEL As Integer

110       Screen.MousePointer = 0

120       intEL = Erl
130       strES = Err.Description
140       LogError "frmEditHisto", "txtsex_Click", intEL, strES


End Sub

Private Sub txtsex_KeyPress(KeyAscii As Integer)

10        On Error GoTo txtsex_KeyPress_Error

20        KeyAscii = 0
30        txtsex_Click

40        Exit Sub

txtsex_KeyPress_Error:

          Dim strES As String
          Dim intEL As Integer

50        Screen.MousePointer = 0

60        intEL = Erl
70        strES = Err.Description
80        LogError "frmEditHisto", "txtsex_KeyPress", intEL, strES


End Sub

Private Sub txtSex_LostFocus()

          Dim ForeName As String

10        On Error GoTo txtSex_LostFocus_Error

20        ForeName = ParseForeName(txtName)
30        If ForeName = "" Then Exit Sub
40        If Trim$(txtSex) = "" Then Exit Sub

50        Exit Sub

txtSex_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

60        Screen.MousePointer = 0

70        intEL = Erl
80        strES = Err.Description
90        LogError "frmEditHisto", "txtSex_LostFocus", intEL, strES


End Sub


Private Sub txtAandE_LostFocus()

10        On Error GoTo txtAandE_LostFocus_Error

20        If Trim(txtName) = "" Then
30            LoadDemo Trim(txtAandE)
40        End If

50        Exit Sub

txtAandE_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

60        Screen.MousePointer = 0

70        intEL = Erl
80        strES = Err.Description
90        LogError "frmEditHisto", "txtAandE_Lostfocus", intEL, strES


End Sub

Private Sub txtDemographicComment_KeyPress(KeyAscii As Integer)

10        On Error GoTo txtDemographicComment_KeyPress_Error

20        bsave.Enabled = True
30        bSaveHold.Enabled = True

40        Exit Sub

txtDemographicComment_KeyPress_Error:

          Dim strES As String
          Dim intEL As Integer

50        Screen.MousePointer = 0

60        intEL = Erl
70        strES = Err.Description
80        LogError "frmEditHisto", "txtDemographicComment_KeyPress", intEL, strES


End Sub

Private Sub txtDemographicComment_LostFocus()

          Dim s As Variant
          Dim n As Long
          Dim sql As String
          Dim tb As New Recordset

10        On Error GoTo txtDemographicComment_LostFocus_Error

20        If Trim$(txtDemographicComment) = "" Then Exit Sub

30        s = Split(txtDemographicComment, " ")

40        For n = 0 To UBound(s)
50            sql = "SELECT Text FROM Lists WHERE ListType = 'DE' AND Text = '" & AddTicks(s(n)) & "'"
60            Set tb = New Recordset
70            RecOpenServer 0, tb, sql
80            If Not tb.EOF Then
90                s(n) = Trim(tb!Text)
100           End If
110       Next

120       txtDemographicComment = Join(s, " ")
130       lblDemographicComment = txtDemographicComment

140       Exit Sub

txtDemographicComment_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

150       Screen.MousePointer = 0

160       intEL = Erl
170       strES = Err.Description
180       LogError "frmEditHisto", "txtDemographicComment_LostFocus", intEL, strES, sql

End Sub

Private Sub txtNoPas_LostFocus()

10        On Error GoTo txtNoPas_LostFocus_Error

20        If Trim(txtName) = "" Then
30            LoadDemo txtNOPAS
40        End If

50        Exit Sub

txtNoPas_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

60        Screen.MousePointer = 0

70        intEL = Erl
80        strES = Err.Description
90        LogError "frmEditHisto", "txtNoPas_LostFocus", intEL, strES


End Sub
Private Sub LoadDemo(ByVal IDNumber As String)

          Dim tb As New Recordset
          Dim sql As String
          Dim IDType As String


10        On Error GoTo LoadDemo_Error

20        IDType = CheckDemographics(IDNumber)
30        If IDType = "" Then
              'clearpatient
40            Exit Sub
50        End If

60        sql = "SELECT * from patientifs WHERE " & _
                IDType & " = '" & IDNumber & "' "

70        Set tb = New Recordset
80        RecOpenServer 0, tb, sql
90        If tb.EOF = True Then
              '   clearpatient
100       Else
110           If tb!Chart & "" = "" Then txtChart = tb!Mrn & "" Else txtChart = tb!Chart & ""
120           txtAandE = tb!AandE & ""
130           txtNOPAS = tb!NOPAS & ""
140           txtName = tb!PatName & ""
150           If Not IsNull(tb!Dob) Then
160               lDoB = Format(tb!Dob, "DD/MM/YYYY")
170           Else
180               lDoB = ""
190           End If
200           lAge = CalcAge(tb!Dob & "", dtSampleDate)
210           Select Case tb!sex & ""
              Case "M": lSex = "Male"
220           Case "F": lSex = "Female"
230           Case Else: lSex = ""
240           End Select
250           txtSex = lSex
260           txtAge = lAge
270           txtDoB = lDoB
280           taddress(0) = tb!Address0 & ""
290           taddress(1) = tb!Address1 & ""
300           cmbWard.Text = tb!Ward & ""
310           cmbClinician.Text = tb!Clinician & ""
320       End If
330       tb.Close



340       Exit Sub

LoadDemo_Error:

          Dim strES As String
          Dim intEL As Integer

350       Screen.MousePointer = 0

360       intEL = Erl
370       strES = Err.Description
380       LogError "frmEditHisto", "LoadDemo", intEL, strES, sql


End Sub

Private Sub txtSampleID_LostFocus()

10        On Error GoTo txtSampleID_LostFocus_Error

20        If Trim$(txtSampleID) = "" Then Exit Sub

30        If Val(txtSampleID) < 1 Or Trim$(txtSampleID) = "" Or Val(txtSampleID) > (2 ^ 31) - 1 Then
40            txtSampleID = ""
50            txtSampleID.SetFocus
60            Exit Sub
70        End If

80        txtSampleID = Val(txtSampleID)
90        txtSampleID = Int(txtSampleID)


100       LoadAllDetails

110       bsave.Enabled = False
120       bSaveHold.Enabled = False
130       cmdSaveHisto(0).Enabled = False
140       cmdSaveHHold(0).Enabled = False
150       cmdSaveHisto(1).Enabled = False
160       cmdSaveHHold(1).Enabled = False
170       cmdSaveCyto.Enabled = False
180       cmdSaveCHold.Enabled = False

190       Exit Sub

txtSampleID_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

200       Screen.MousePointer = 0

210       intEL = Erl
220       strES = Err.Description
230       LogError "frmEditHisto", "txtSampleID_LostFocus", intEL, strES


End Sub

Private Sub UPDATEMRU()

          Dim sql As String
          Dim tb As New Recordset
          Dim n As Long
          Dim Found As Boolean
          Dim NewMRU(0 To 9, 0 To 1) As String
          '(x,0) SampleID
          '(x,1) DateTime

10        On Error GoTo UPDATEMRU_Error

20        sql = "SELECT top 10 * from HMRU WHERE " & _
                "UserCode = '" & UserCode & "' " & _
                "Order by DateTime desc"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql

50        n = -1
60        Do While Not tb.EOF
70            n = n + 1
80            NewMRU(n, 0) = Trim$(tb!SampleID)
90            NewMRU(n, 1) = tb!Datetime
100           tb.MoveNext
110       Loop

120       Found = False
130       For n = 0 To 9
140           If txtSampleID = NewMRU(n, 0) Then
150               sql = "UPDATE HMRU " & _
                        "Set DateTime = '" & Format$(Now, "dd/mmm/yyyy hh:mm:ss") & "' " & _
                        "WHERE SampleID = '" & txtYear & lblDisp & txtSampleID & "' " & _
                        "and UserCode = '" & UserCode & "'"
160               Cnxn(0).Execute sql
170               Found = True
180               Exit For
190           End If
200       Next

210       If Not Found Then
220           sql = "DELETE from HMRU WHERE " & _
                    "UserCode = '" & UserCode & "'"
230           Cnxn(0).Execute sql
240           For n = 0 To 8
250               If NewMRU(n, 0) <> "" Then
260                   sql = "INSERT into HMRU " & _
                            "(SampleID, DateTime, UserCode ) VALUES " & _
                            "('" & NewMRU(n, 0) & "', " & _
                            "'" & Format$(NewMRU(n, 1), "dd/mmm/yyyy hh:mm:ss") & "', " & _
                            "'" & UserCode & "')"
270                   Cnxn(0).Execute sql
280               End If
290           Next
300           sql = "INSERT into HMRU " & _
                    "(SampleID, DateTime, UserCode ) VALUES " & _
                    "('" & txtYear & lblDisp & txtSampleID & "', " & _
                    "'" & Format$(Now, "dd/mmm/yyyy hh:mm:ss") & "', " & _
                    "'" & UserCode & "')"
310           Cnxn(0).Execute sql
320       End If

330       FillMRU

340       Exit Sub

UPDATEMRU_Error:

          Dim strES As String
          Dim intEL As Integer

350       Screen.MousePointer = 0

360       intEL = Erl
370       strES = Err.Description
380       LogError "frmEditHisto", "UPDATEMRU", intEL, strES, sql


End Sub

Private Sub txtYear_KeyPress(KeyAscii As Integer)

10        On Error GoTo txtYear_KeyPress_Error

20        KeyAscii = VI(KeyAscii, Numeric_Only)


30        Exit Sub

txtYear_KeyPress_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmEditHisto", "txtYear_KeyPress", intEL, strES


End Sub

Private Sub txtYear_LostFocus()

10        On Error GoTo txtYear_LostFocus_Error

20        If Trim$(txtSampleID) = "" Then Exit Sub

30        LoadAllDetails

40        bSaveHold.Enabled = False
50        bsave.Enabled = False
60        cmdSaveHisto(0).Enabled = False
70        cmdSaveHHold(0).Enabled = False
80        cmdSaveHisto(1).Enabled = False
90        cmdSaveHHold(1).Enabled = False
100       cmdSaveCyto.Enabled = False
110       cmdSaveCHold.Enabled = False

120       Exit Sub

txtYear_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

130       Screen.MousePointer = 0

140       intEL = Erl
150       strES = Err.Description
160       LogError "frmEditHisto", "txtYear_LostFocus", intEL, strES


End Sub

Private Sub udNoCopies_DownClick()
10        txtNoCopies = Val(txtNoCopies) - 1
20        If Val(txtNoCopies) < 1 Then txtNoCopies = "1"
End Sub

Private Sub udNoCopies_UpClick()
10        txtNoCopies = Val(txtNoCopies) + 1
20        If Val(txtNoCopies) > 9 Then txtNoCopies = "9"

End Sub

Private Sub upBlck_DownClick(Index As Integer)

10        On Error GoTo upBlck_DownClick_Error

20        If Val(lblBlock(Index)) > 0 Then
30            lblBlock(Index) = Val(lblBlock(Index)) - 1
40            If lblBlock(Index) = "0" Then
50                lblBlock(Index) = ""
60            End If
70            grdSpec(Index).Cols = Val(lblBlock(Index)) + 1
80            cmdSaveHisto(0).Enabled = True
90            cmdSaveHHold(0).Enabled = True
100           cmdSaveHisto(1).Enabled = True
110           cmdSaveHHold(1).Enabled = True
120       Else
              '  If lblBlock(Index) = "0" Then
              '    lblBlock(Index) = ""
              '  End If
130           grdSpec(Index).Cols = 1
140           cmdSaveHisto(0).Enabled = True
150           cmdSaveHHold(0).Enabled = True
160           cmdSaveHisto(1).Enabled = True
170           cmdSaveHHold(1).Enabled = True
180       End If

190       Exit Sub

upBlck_DownClick_Error:

          Dim strES As String
          Dim intEL As Integer

200       Screen.MousePointer = 0

210       intEL = Erl
220       strES = Err.Description
230       LogError "frmEditHisto", "upBlck_DownClick", intEL, strES


End Sub

Private Sub upBlck_UpClick(Index As Integer)

10        On Error GoTo upBlck_UpClick_Error

20        lblBlock(Index) = Val(lblBlock(Index)) + 1
30        grdSpec(Index).Cols = Val(lblBlock(Index)) + 1
40        grdSpec(Index).ColWidth(Val(lblBlock(Index))) = 600
50        grdSpec(Index).TextMatrix(0, Val(lblBlock(Index))) = "Blk " & Val(lblBlock(Index))
60        fraSpec(Index).Visible = True

70        cmdSaveHisto(0).Enabled = True
80        cmdSaveHHold(0).Enabled = True
90        cmdSaveHisto(1).Enabled = True
100       cmdSaveHHold(1).Enabled = True


110       Exit Sub

upBlck_UpClick_Error:

          Dim strES As String
          Dim intEL As Integer

120       Screen.MousePointer = 0

130       intEL = Erl
140       strES = Err.Description
150       LogError "frmEditHisto", "upBlck_UpClick", intEL, strES


End Sub

Private Sub UpDown1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        On Error GoTo UpDown1_MouseUp_Error

20        pBar = 0

30        LoadAllDetails

40        bSaveHold.Enabled = False
50        bsave.Enabled = False
60        cmdSaveHisto(0).Enabled = False
70        cmdSaveHHold(0).Enabled = False
80        cmdSaveHisto(1).Enabled = False
90        cmdSaveHHold(1).Enabled = False
100       cmdSaveCyto.Enabled = False
110       cmdSaveCHold.Enabled = False

120       Exit Sub

UpDown1_MouseUp_Error:

          Dim strES As String
          Dim intEL As Integer

130       Screen.MousePointer = 0

140       intEL = Erl
150       strES = Err.Description
160       LogError "frmEditHisto", "UpDown1_MouseUp", intEL, strES


End Sub

Private Sub t_KeyPress(KeyAscii As Integer)

          Dim Code As String
          Dim phrase As String
          Dim NewT As String
          Dim NewINSERTionPoint As Long
          Dim tb As New Recordset
          Dim sql As String

10        On Error GoTo t_KeyPress_Error

20        If T.Locked = True Then Exit Sub
30        pBar = 0
40        If KeyAscii = Asc(" ") Then    'possible end of code
50            Code = lastword()
60            If Trim(Code) <> "" And Len(Code) < 5 Then    'possible code
70                sql = "SELECT * from lists WHERE listtype = 'PH' and code = '" & AddTicks(Code) & "'"
80                Set tb = New Recordset
90                RecOpenServer 0, tb, sql
100               If Not tb.EOF Then
110                   phrase = Trim(tb!Text)
120                   NewT = Left(T, T.SelStart - Len(Code))
130                   NewT = NewT & phrase
140                   NewINSERTionPoint = Len(NewT)
150                   NewT = NewT & Mid$(T, T.SelStart + 1)
160                   T = NewT
170                   T.SelStart = NewINSERTionPoint
180                   KeyAscii = 0
190               End If
200           End If
210       End If

220       cmdSaveHisto(0).Enabled = True
230       cmdSaveHHold(0).Enabled = True
240       cmdSaveHisto(1).Enabled = True
250       cmdSaveHHold(1).Enabled = True

260       Exit Sub

t_KeyPress_Error:

          Dim strES As String
          Dim intEL As Integer

270       Screen.MousePointer = 0

280       intEL = Erl
290       strES = Err.Description
300       LogError "frmEditHisto", "t_KeyPress", intEL, strES, sql

End Sub

Private Function lastword() As String

          Dim n As Long
          Dim Code As String
          Dim character As String

10        On Error GoTo lastword_Error

20        For n = T.SelStart To 1 Step -1
30            character = Mid$(T, n, 1)
40            If character = " " Then Exit For
50            If character = chr$(10) Then Exit For
60            Code = character & Code
70        Next
80        lastword = Code

90        Exit Function

lastword_Error:

          Dim strES As String
          Dim intEL As Integer

100       Screen.MousePointer = 0

110       intEL = Erl
120       strES = Err.Description
130       LogError "frmEditHisto", "lastword", intEL, strES


End Function

Private Function lastwordC() As String

          Dim n As Long
          Dim Code As String
          Dim character As String

10        On Error GoTo lastwordC_Error

20        For n = txtCyto.SelStart To 1 Step -1
30            character = Mid$(txtCyto, n, 1)
40            If character = " " Then Exit For
50            If character = chr$(10) Then Exit For
60            Code = character & Code
70        Next
80        lastwordC = Code

90        Exit Function

lastwordC_Error:

          Dim strES As String
          Dim intEL As Integer

100       Screen.MousePointer = 0

110       intEL = Erl
120       strES = Err.Description
130       LogError "frmEditHisto", "lastwordC", intEL, strES


End Function


Private Sub txtCyto_KeyPress(KeyAscii As Integer)

          Dim sql As String
          Dim tb As New Recordset
          Dim Code As String
          Dim phrase As String
          Dim NewT As String
          Dim NewINSERTionPoint As Long

10        On Error GoTo txtCyto_KeyPress_Error

20        If txtCyto.Locked = True Then Exit Sub
30        pBar = 0

40        If KeyAscii = Asc(" ") Then    'possible end of code
50            Code = lastwordC()
60            If Trim(Code) <> "" And Len(Code) < 5 Then    'possible code
70                sql = "SELECT * from lists WHERE listtype = 'PH' and code = '" & AddTicks(Code) & "'"
80                Set tb = New Recordset
90                RecOpenServer 0, tb, sql
100               If Not tb.EOF Then
110                   phrase = Trim(tb!Text)
120                   NewT = Left(txtCyto, txtCyto.SelStart - Len(Code))
130                   NewT = NewT & phrase
140                   NewINSERTionPoint = Len(NewT)
150                   NewT = NewT & Mid$(txtCyto, txtCyto.SelStart + 1)
160                   txtCyto = NewT
170                   txtCyto.SelStart = NewINSERTionPoint
180                   KeyAscii = 0
190               End If
200           End If
210       End If

220       cmdSaveCyto.Enabled = True
230       cmdSaveCHold.Enabled = True

240       Exit Sub

txtCyto_KeyPress_Error:

          Dim strES As String
          Dim intEL As Integer

250       Screen.MousePointer = 0

260       intEL = Erl
270       strES = Err.Description
280       LogError "frmEditHisto", "txtCyto_KeyPress", intEL, strES, sql


End Sub

Private Sub ClearAll()

          Dim n As Long
10        On Error GoTo ClearAll_Error

20        lblA = ""
30        lblB = ""
40        lblC = ""
50        lblD = ""

60        lblVal = ""
70        For n = 3 To 9
80            c(n) = ""
90        Next
100       T = ""
110       txtCyto = ""

120       Exit Sub

ClearAll_Error:

          Dim strES As String
          Dim intEL As Integer

130       Screen.MousePointer = 0

140       intEL = Erl
150       strES = Err.Description
160       LogError "frmEditHisto", "ClearAll", intEL, strES


End Sub
Private Sub LoadRecord()

          Dim tb As New Recordset
          Dim sql As String
          Dim Yadd As Long
          Dim Valid As Boolean

10        On Error GoTo LoadRecord_Error

20        ClearAll

30        If Trim$(txtSampleID) = "" Then Exit Sub

40        Valid = False
50        Yadd = Val(Swap_Year(txtYear)) * 1000

60        sql = "SELECT * FROM Demographics WHERE " & _
                "SampleID = '" & txtSampleID + SysOptHistoOffset(0) + Yadd & "' " & _
                "AND hYear = '" & txtYear & "'"

70        Set tb = New Recordset
80        RecOpenServer 0, tb, sql

90        If tb.EOF Then
100           UnlockRecord
110       Else
120           If Not IsNull(tb!histovalid) And tb!histovalid Then
130               Valid = True
140               LockRecord
150           Else
160               UnlockRecord
170           End If
180       End If

190       Clear_HistoWork

200       sql = "SELECT * FROM Historesults WHERE " & _
                "SampleID = '" & txtSampleID + SysOptHistoOffset(0) + Yadd & "' " & _
                "AND hyear = '" & txtYear & "'"
210       Set tb = New Recordset
220       RecOpenServer 0, tb, sql

230       If Not tb.EOF Then
240           c(13) = tb!histocomment & ""
250           c(0) = tb!NatureOfSpecimen & ""
260           c(1) = tb!natureofspecimenB & ""
270           c(2) = tb!natureofspecimenC & ""
280           c(3) = tb!natureofspecimenD & ""
290           c(4) = tb!natureofspecimene & ""
300           c(5) = tb!natureofspecimenf & ""
310           lblA = "A: " & c(0)
320           lblB = "B: " & c(1)
330           lblC = "C: " & c(2)
340           lblD = "D: " & c(3)
350           T = tb!historeport & ""
360           If Trim(tb!validdate & "") <> "" Then StatusBar1.Panels(5).Text = "Validated on " & Format(tb!validdate, "dd/MMM/yyyy hh:mm")
370           chkNCRI(0).Value = IIf(IsNull(tb!ncri), 0, tb!ncri)
380           SSTab1.TabCaption(2) = "<< Histology Report >>"
390           If Valid = True Then lblVal = "Validated by " & tb!UserName & ""
400       End If

410       Exit Sub

LoadRecord_Error:

          Dim strES As String
          Dim intEL As Integer

420       intEL = Erl
430       strES = Err.Description
440       LogError "frmEditHisto", "LoadRecord", intEL, strES, sql

End Sub
Sub LockRecord()
          Dim n As Long

10        On Error GoTo LockRecord_Error

20        Frame2.Enabled = False
30        For n = 0 To 5
40            fraSpec(n).Enabled = False
50        Next
60        txtHistoComment.Locked = True
70        T.Locked = True
80        chkNCRI(0).Enabled = False
90        cmdHVal.Caption = "VALID"

100       Exit Sub

LockRecord_Error:

          Dim strES As String
          Dim intEL As Integer

110       Screen.MousePointer = 0

120       intEL = Erl
130       strES = Err.Description
140       LogError "frmEditHisto", "LockRecord", intEL, strES


End Sub
Sub LockCRecord()

10        On Error GoTo LockCRecord_Error

20        txtCyto.Locked = True
30        cmdCVal.Caption = "VALID"
40        c(13).Enabled = False
50        Frame3.Enabled = False
60        chkNCRI(1).Enabled = False
70        txtCytoComment.Locked = True

80        Exit Sub

LockCRecord_Error:

          Dim strES As String
          Dim intEL As Integer

90        Screen.MousePointer = 0

100       intEL = Erl
110       strES = Err.Description
120       LogError "frmEditHisto", "LockCRecord", intEL, strES


End Sub
Sub UnlockRecord()
          Dim n As Long

10        On Error GoTo UnlockRecord_Error

20        Frame2.Enabled = True
30        For n = 0 To 5
40            fraSpec(n).Enabled = True
50        Next

60        txtHistoComment.Locked = False
70        T.Locked = False
80        cmdHVal.Caption = "Validate"
90        chkNCRI(0).Enabled = True

100       Exit Sub

UnlockRecord_Error:

          Dim strES As String
          Dim intEL As Integer

110       Screen.MousePointer = 0

120       intEL = Erl
130       strES = Err.Description
140       LogError "frmEditHisto", "UnlockRecord", intEL, strES


End Sub
Sub UnlockCRecord()

10        On Error GoTo UnlockCRecord_Error

20        txtCyto.Locked = False
30        cmdCVal.Caption = "Validate"
40        c(13).Enabled = True
50        Frame3.Enabled = True
60        c(9).Enabled = True
70        c(8).Enabled = True
80        c(7).Enabled = True
90        c(6).Enabled = True
100       chkNCRI(1).Enabled = True
110       txtCytoComment.Locked = False

120       Exit Sub

UnlockCRecord_Error:

          Dim strES As String
          Dim intEL As Integer

130       Screen.MousePointer = 0

140       intEL = Erl
150       strES = Err.Description
160       LogError "frmEditHisto", "UnlockCRecord", intEL, strES


End Sub

Sub CLockRecord()

10        On Error GoTo CLockRecord_Error

20        iCUnlocked.Visible = False
30        iCLocked.Visible = True
40        iCKey.Visible = True
50        T.Locked = False
60        c(13).Enabled = False
70        Frame3.Enabled = False
80        c(9).Enabled = False
90        c(8).Enabled = False
100       c(7).Enabled = False
110       c(6).Enabled = False
120       txtCytoComment.Locked = True

130       Exit Sub

CLockRecord_Error:

          Dim strES As String
          Dim intEL As Integer

140       Screen.MousePointer = 0

150       intEL = Erl
160       strES = Err.Description
170       LogError "frmEditHisto", "CLockRecord", intEL, strES


End Sub


Private Function CheckDemographics(ByVal TrialID As String) _
        As String

          Dim sn As New Recordset
          Dim sql As String
          Dim n As Long
          Dim pName(1 To 4) As String
          Dim pAddress(1 To 4) As String
          Dim pDoB(1 To 4) As String
          Dim IDFound(1 To 4) As Boolean
          Dim Found As Long
          Dim f As Form

10        On Error GoTo CheckDemographics_Error

20        If TrialID = "" Then Exit Function


30        Set sn = New Recordset
40        With sn
50            Found = 0
60            For n = 1 To 4
70                IDFound(n) = False
80                sql = "SELECT * from patientifs WHERE " & _
                        Choose(n, "CHART", "NOPAS", "MRN", "AandE") & " = '" & TrialID & "'"
90                RecOpenServer 0, sn, sql
100               If Not .EOF Then
110                   Do While Not sn.EOF
120                       IDFound(n) = True
130                       Found = Found + 1
140                       pName(n) = !PatName
150                       If Not IsNull(!Dob) Then pDoB(n) = Format(!Dob, "dd/MM/yyyy")
160                       pAddress(n) = !Address0 & " " & !Address1 & ""
170                       sn.MoveNext
180                   Loop
190               End If
200               .Close
210           Next
220       End With

230       If Found = 0 Then
240           CheckDemographics = ""
250       ElseIf Found = 1 Then
260           For n = 1 To 4
270               If IDFound(n) Then
280                   CheckDemographics = Choose(n, "CHART", "NOPAS", "MRN", "AandE")
290                   Exit For
300               End If
310           Next
320       Else
330           Set f = New frmDemogCheck
340           With f
350               For n = 1 To 4
360                   If IDFound(n) Then
370                       .bSelect(n).Visible = True
380                       .lName(n) = pName(n)
390                       .lAddress(n) = pAddress(n)
400                       .lDoB(n) = pDoB(n)
410                   End If
420               Next
430               .Show 1
440               CheckDemographics = .IDType
450           End With
460           Unload f
470           Set f = Nothing
480       End If


490       Exit Function

CheckDemographics_Error:

          Dim strES As String
          Dim intEL As Integer

500       Screen.MousePointer = 0

510       intEL = Erl
520       strES = Err.Description
530       LogError "frmEditHisto", "CheckDemographics", intEL, strES, sql


End Function



Private Sub LoadHistoWork()
          Dim tb As New Recordset
          Dim sql As String
          Dim n As Long
          Dim X As Long
          Dim stainfound As Boolean
          Dim z As Long
          Dim Rw As Long
          Dim Yadd As Long



10        On Error GoTo LoadHistoWork_Error

20        If Trim$(txtSampleID) = "" Then Exit Sub

30        Yadd = Val(Swap_Year(txtYear)) * 1000
40        sql = "SELECT * from HistoSpecimen WHERE sampleid = '" & txtSampleID + SysOptHistoOffset(0) + Yadd & "' and hyear = '" & txtYear & "' order by specimen asc "
50        Set tb = New Recordset
60        RecOpenServer 0, tb, sql

70        Do While Not tb.EOF
80            If tb!Rundate & "" <> "" Then HDate = tb!Rundate
90            SSTab1.TabCaption(1) = "<< Histology Work Screen >>"
100           lblStatus = Trim$(tb!Status & "")
110           For n = 0 To 5
120               If tb!specimen = n Then
130                   fraSpec(n).Caption = fraSpec(n).Caption & " - " & Trim(tb!Type)
140                   For X = 0 To c(n).ListCount
150                       If UCase(Trim(c(n).List(X))) = UCase(Trim(tb!Type)) Then
160                           c(n).ListIndex = X
170                       End If
180                   Next
190                   lblFS(n) = Trim(tb!FS & "")
200                   txtHist(n) = Trim(tb!Remark & "")
210                   lblBlock(n) = Trim(tb!blocks)
220               End If
230           Next
240           tb.MoveNext
250       Loop

260       RecClose tb


270       sql = "SELECT * from HistoBlock WHERE sampleid = '" & txtSampleID + SysOptHistoOffset(0) + Yadd & "' and hyear = '" & txtYear & "' order by specimen, block"
280       Set tb = New Recordset
290       RecOpenServer 0, tb, sql

300       Do While Not tb.EOF
310           For n = 0 To 5
320               If Trim(c(n)) = Trim(tb!Type) Then
330                   For X = 1 To Val(lblBlock(n))
340                       If tb!block = X Then

350                           fraSpec(n).Visible = True
360                           grdSpec(n).Cols = Val(lblBlock(n)) + 1
370                           grdComm(n).Cols = grdSpec(n).Cols
380                           grdSpec(n).ColWidth(X) = 600
390                           grdSpec(n).TextMatrix(0, X) = "Blk " & X
400                           grdSpec(n).TextMatrix(1, 0) = "Pieces"
410                           grdSpec(n).TextMatrix(1, X) = tb!pieces & ""
420                           grdComm(n).TextMatrix(1, X) = Trim$(tb!picomm & "")
430                       End If
440                   Next
450               End If
460           Next
470           tb.MoveNext
480       Loop

490       RecClose tb


500       sql = "SELECT * from HistoStain WHERE sampleid = '" & txtSampleID + SysOptHistoOffset(0) + Yadd & "' order by block,grid"
510       Set tb = New Recordset
520       RecOpenServer 0, tb, sql

530       Do While Not tb.EOF
540           For n = 0 To 5
550               If n = tb!specimen Then
560                   For X = 1 To grdSpec(n).Cols
570                       If X = tb!block Then
580                           stainfound = False
590                           For z = 2 To grdSpec(n).Rows - 1
600                               If grdSpec(n).TextMatrix(z, 0) = tb!stain Then
610                                   stainfound = True
620                                   Rw = z
630                               End If
640                           Next
650                           If Not stainfound Then
660                               grdSpec(n).AddItem tb!stain
670                               grdComm(n).AddItem ""
680                           End If
690                           If Rw > 0 Then
700                               grdSpec(n).TextMatrix(Rw, X) = Trim(tb!Result)
710                               grdComm(n).TextMatrix(Rw, X) = Trim$(tb!ResComm & "")
720                           Else
730                               grdSpec(n).TextMatrix(grdSpec(n).Rows - 1, X) = Trim(tb!Result)
740                               grdComm(n).TextMatrix(grdComm(n).Rows - 1, X) = Trim$(tb!ResComm & "")
750                           End If
760                       End If
770                   Next
780               End If
790           Next
800           tb.MoveNext
810       Loop
820       RecClose tb



          'For n = 0 To 5
          '  grdComm(n).Cols = z + 1
          '  grdSpec(n).Cols = z + 1
          '  For x = 1 To grdComm(n).Rows - 1
          '    For z = 0 To z
          '      If grdComm(n).TextMatrix(x, z) <> "" Then
          '        grdSpec(n).Row = x
          '        grdSpec(n).Col = z
          '        grdSpec(n).CellBackColor = vbYellow
          '      End If
          '    Next
          '  Next
          'Next


830       Exit Sub

LoadHistoWork_Error:

          Dim strES As String
          Dim intEL As Integer

840       Screen.MousePointer = 0

850       intEL = Erl
860       strES = Err.Description
870       LogError "frmEditHisto", "LoadHistoWork", intEL, strES, sql


End Sub

Private Sub txtHistoComment_Change()

10        On Error GoTo txtHistoComment_Change_Error

20        pBar = 0

30        Exit Sub

txtHistoComment_Change_Error:

          Dim strES As String
          Dim intEL As Integer

40        Screen.MousePointer = 0

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmEditHisto", "txtHistoComment_Change", intEL, strES


End Sub

Private Sub txtHistoComment_KeyDown(KeyCode As Integer, Shift As Integer)
          Dim s As Variant
          Dim n As Long
          Dim z As Long
          Dim tb As New Recordset
          Dim sql As String




10        On Error GoTo txtHistoComment_KeyDown_Error

20        If KeyCode = 113 Then

30            n = txtHistoComment.SelStart

40            z = 2
50            s = Mid(txtHistoComment, n - z, z + 1)
60            z = 3
70            If ListText("HI", s) <> "" Then
80                s = ListText("HI", s)
90            Else
100               s = ""
110           End If

120           If s = "" Then
130               z = 1
140               s = Mid(txtHistoComment, n - z, z + 1)
150               z = 2
160               If ListText("HI", s) <> "" Then
170                   s = ListText("HI", s)
180               Else
190                   s = ""
200               End If
210           End If

220           If s = "" Then
230               z = 1
240               s = Mid(txtHistoComment, n, z)

250               If ListText("HI", s) <> "" Then
260                   s = ListText("HI", s)
270               End If
280           End If

290           txtHistoComment = Left(txtHistoComment, n - (z))
300           txtHistoComment = txtHistoComment & s

310           txtHistoComment.SelStart = Len(txtHistoComment)

320       ElseIf KeyCode = 114 Then

330           sql = "SELECT * from lists WHERE listtype = 'HI'"
340           Set tb = New Recordset
350           RecOpenServer 0, tb, sql
360           Do While Not tb.EOF
370               s = Trim(tb!Text)
380               frmMessages.lstComm.AddItem s
390               tb.MoveNext
400           Loop

410           Set frmMessages.f = Me
420           Set frmMessages.T = txtHistoComment
430           frmMessages.Show 1

440       End If

450       cmdSaveHisto(0).Enabled = True
460       cmdSaveHHold(0).Enabled = True
470       cmdSaveHisto(1).Enabled = True
480       cmdSaveHHold(1).Enabled = True




490       Exit Sub

txtHistoComment_KeyDown_Error:

          Dim strES As String
          Dim intEL As Integer

500       intEL = Erl
510       strES = Err.Description
520       LogError "frmEditHisto", "txtHistoComment_KeyDown", intEL, strES


End Sub


Private Sub loadCrecord()

          Dim tb As New Recordset
          Dim sql As String
          Dim Yadd As Long

10        On Error GoTo loadCrecord_Error

20        ClearAll

30        Yadd = Val(Swap_Year(txtYear)) * 1000

40        sql = "SELECT * FROM Demographics WHERE " & _
                "SampleID = '" & txtSampleID + SysOptCytoOffset(0) + Yadd & "' " & _
                "AND hYear = '" & txtYear & "'"

50        Set tb = New Recordset
60        RecOpenServer 0, tb, sql

70        If tb.EOF Then
80            UnlockCRecord
90        Else
100           If Not IsNull(tb!cytovalid) And tb!cytovalid Then
110               LockCRecord
120           Else
130               UnlockCRecord
140           End If
150       End If

160       sql = "SELECT * FROM CytoResults WHERE " & _
                "SampleID = '" & txtSampleID + SysOptCytoOffset(0) + Yadd & "' " & _
                "AND hYear = '" & txtYear & "'"
170       Set tb = New Recordset
180       RecOpenServer 0, tb, sql

190       If Not tb.EOF Then
200           c(13) = tb!cytocomment & ""
210           c(6) = tb!NatureOfSpecimen & ""
220           c(7) = tb!natureofspecimenB & ""
230           c(8) = tb!natureofspecimenC & ""
240           c(9) = tb!natureofspecimenD & ""
250           txtCyto = tb!cytoreport & ""
260           If Trim(tb!validdate & "") <> "" Then StatusBar1.Panels(5).Text = Format(tb!validdate, "dd/MMM/yyyy hh:mm")
270           If Trim(tb!ncri & "") <> "" Then chkNCRI(1).Value = tb!ncri
280       End If

290       Exit Sub

loadCrecord_Error:

          Dim strES As String
          Dim intEL As Integer

300       intEL = Erl
310       strES = Err.Description
320       LogError "frmEditHisto", "loadCrecord", intEL, strES, sql

End Sub

Private Sub Set_Demo(ByVal Demo As Boolean)

10        On Error GoTo Set_Demo_Error

20        Frame4.Enabled = Demo
30        Frame5.Enabled = Demo
40        Frame7.Enabled = Demo

50        'txtChart.Locked = Not Demo
60        txtAandE.Locked = Not Demo
70        txtNOPAS.Locked = Not Demo
80        txtName.Locked = Not Demo
90        txtDoB.Locked = Not Demo
100       txtAge.Locked = Not Demo
110       txtSex.Locked = Not Demo

120       If Demo = False Then
130           StatusBar1.Panels(3).Text = "Demographics Validated"
140           StatusBar1.Panels(3).Bevel = sbrInset
150       Else
160           StatusBar1.Panels(3).Text = "Check Demographics"
170           StatusBar1.Panels(3).Bevel = sbrRaised
180       End If


190       Exit Sub

Set_Demo_Error:

          Dim strES As String
          Dim intEL As Integer

200       Screen.MousePointer = 0

210       intEL = Erl
220       strES = Err.Description
230       LogError "frmEditHisto", "Set_Demo", intEL, strES


End Sub

Private Sub cmdDemoVal_Click()

          Dim sql As String
          Dim tb As New Recordset
          Dim SampleID As String
          Dim Yadd As Long

10        On Error GoTo cmdDemoVal_Click_Error

20        Yadd = Val(Swap_Year(txtYear)) * 1000

30        If lblDisp = "H" Then
40            SampleID = txtSampleID + SysOptHistoOffset(0) + Yadd
50        Else
60            SampleID = txtSampleID + SysOptCytoOffset(0) + Yadd
70        End If

80        If cmdDemoVal.Caption = "Validate" Then
90            sql = "SELECT * FROM Demographics WHERE " & _
                    "SampleID = '" & SampleID & "'"
100           Set tb = New Recordset
110           RecOpenServer 0, tb, sql
120           If Not tb.EOF Then
130               sql = "UPDATE demographics set valid = 1, " & _
                        "username = '" & UserName & "' WHERE " & _
                        "sampleid = '" & SampleID & "'"
140               Cnxn(0).Execute sql
150               Set_Demo False
160               cmdDemoVal.Caption = "VALID"
170           End If
180       Else
190           If UCase(iBOX("Enter password to unValidate ?", , , True)) = UserPass Then
200               sql = "SELECT * FROM Demographics WHERE " & _
                        "SampleID = '" & SampleID & "'"
210               Set tb = New Recordset
220               RecOpenServer 0, tb, sql
230               If Not tb.EOF Then
240                   sql = "UPDATE demographics set valid = 0, " & _
                            "username = '" & UserName & "' WHERE " & _
                            "sampleid = '" & SampleID & "'"
250                   Cnxn(0).Execute sql
260                   Set_Demo True
270                   cmdDemoVal.Caption = "Validate"
280               End If
290           End If
300       End If

310       If SaveDemographics_Click = False Then Exit Sub

320       Exit Sub

cmdDemoVal_Click_Error:

          Dim strES As String
          Dim intEL As Integer

330       Screen.MousePointer = 0

340       intEL = Erl
350       strES = Err.Description
360       LogError "frmEditHisto", "cmdDemoVal_Click", intEL, strES, sql


End Sub




Private Sub cmdSaveCyto_Click()

10        On Error GoTo cmdSaveCyto_Click_Error

20        SaveCytology
30        SaveComments

40        txtSampleID = Format$(Val(txtSampleID) + 1)
50        LoadAllDetails

60        Exit Sub

cmdSaveCyto_Click_Error:

          Dim strES As String
          Dim intEL As Integer

70        Screen.MousePointer = 0

80        intEL = Erl
90        strES = Err.Description
100       LogError "frmEditHisto", "cmdSaveCyto_Click", intEL, strES


End Sub



Private Sub upFS_DownClick(Index As Integer)

10        On Error GoTo upFS_DownClick_Error

20        If Val(lblFS(Index)) > 0 Then
30            lblFS(Index) = Val(lblFS(Index)) - 1
40            If lblFS(Index) = "0" Then
50                lblFS(Index) = ""
60            End If
70            grdSpec(Index).Cols = Val(lblFS(Index)) + 1
80            cmdSaveHisto(0).Enabled = True
90            cmdSaveHHold(0).Enabled = True
100           cmdSaveHisto(1).Enabled = True
110           cmdSaveHHold(1).Enabled = True
120       Else
              '  If lblBlock(Index) = "0" Then
              '    lblBlock(Index) = ""
              '  End If
130           grdSpec(Index).Cols = 1
140           cmdSaveHisto(0).Enabled = True
150           cmdSaveHHold(0).Enabled = True
160           cmdSaveHisto(1).Enabled = True
170           cmdSaveHHold(1).Enabled = True
180       End If


190       Exit Sub

upFS_DownClick_Error:

          Dim strES As String
          Dim intEL As Integer

200       Screen.MousePointer = 0

210       intEL = Erl
220       strES = Err.Description
230       LogError "frmEditHisto", "upFS_DownClick", intEL, strES


End Sub

Private Sub upFS_UpClick(Index As Integer)


10        On Error GoTo upFS_UpClick_Error

20        lblFS(Index) = Val(lblFS(Index)) + 1

30        cmdSaveHisto(0).Enabled = True
40        cmdSaveHHold(0).Enabled = True
50        cmdSaveHisto(1).Enabled = True
60        cmdSaveHHold(1).Enabled = True

70        Exit Sub

upFS_UpClick_Error:

          Dim strES As String
          Dim intEL As Integer

80        Screen.MousePointer = 0

90        intEL = Erl
100       strES = Err.Description
110       LogError "frmEditHisto", "upFS_UpClick", intEL, strES


End Sub
