VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "comct232.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmEditMicrobiologyNew 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Microbiology"
   ClientHeight    =   10215
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   14625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10215
   ScaleWidth      =   14625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdHealthLink 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12240
      Picture         =   "frmEditBacteriologyNew.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   517
      Top             =   8760
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.TextBox txtWhoPrinted 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8505
      Locked          =   -1  'True
      TabIndex        =   508
      Text            =   "Printed By"
      Top             =   9480
      Visible         =   0   'False
      Width           =   4200
   End
   Begin VB.TextBox txtWhoValidated 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4245
      Locked          =   -1  'True
      TabIndex        =   507
      Text            =   "Validated By"
      Top             =   9480
      Visible         =   0   'False
      Width           =   4200
   End
   Begin VB.Frame Frame6 
      Caption         =   "Copies"
      Height          =   645
      Left            =   13140
      TabIndex        =   484
      Top             =   2610
      Width           =   1305
      Begin VB.TextBox txtNoCopies 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   485
         Text            =   "1"
         Top             =   240
         Width           =   270
      End
      Begin ComCtl2.UpDown udNoCopies 
         Height          =   360
         Left            =   405
         TabIndex        =   486
         Top             =   210
         Width           =   240
         _ExtentX        =   450
         _ExtentY        =   635
         _Version        =   327681
         Value           =   1
         BuddyControl    =   "txtNoCopies"
         BuddyDispid     =   196613
         OrigLeft        =   420
         OrigTop         =   210
         OrigRight       =   675
         OrigBottom      =   585
         Max             =   9
         Min             =   1
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label lblInterim 
         Alignment       =   2  'Center
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "I"
         Height          =   255
         Left            =   990
         TabIndex        =   488
         ToolTipText     =   "Print Interim Report"
         Top             =   270
         Width           =   240
      End
      Begin VB.Label lblFinal 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "F"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   750
         TabIndex        =   487
         ToolTipText     =   "Print Final Report"
         Top             =   270
         Width           =   240
      End
   End
   Begin VB.TextBox txtWhoSaved 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   180
      Locked          =   -1  'True
      TabIndex        =   383
      Text            =   "Saved By"
      Top             =   9480
      Visible         =   0   'False
      Width           =   4000
   End
   Begin VB.Frame fraSampleID 
      Height          =   1425
      Left            =   225
      TabIndex        =   275
      Top             =   180
      Width           =   2595
      Begin VB.TextBox txtSampleID 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   13.5
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   150
         MaxLength       =   8
         TabIndex        =   0
         Top             =   540
         Width           =   2145
      End
      Begin VB.ComboBox cMRU 
         Height          =   315
         Left            =   570
         TabIndex        =   277
         Text            =   "cMRU"
         Top             =   1050
         Width           =   1605
      End
      Begin VB.ComboBox cmbSiteSearch 
         Height          =   315
         Left            =   540
         TabIndex        =   276
         Text            =   "cmbSiteSearch"
         Top             =   210
         Width           =   1185
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   480
         Left            =   2280
         TabIndex        =   278
         Top             =   540
         Width           =   240
         _ExtentX        =   450
         _ExtentY        =   847
         _Version        =   327681
         Value           =   1
         BuddyControl    =   "txtSampleID"
         BuddyDispid     =   196618
         OrigLeft        =   1920
         OrigTop         =   540
         OrigRight       =   2160
         OrigBottom      =   1020
         Max             =   999999999
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Sample ID"
         Height          =   195
         Index           =   92
         Left            =   720
         TabIndex        =   280
         Top             =   30
         Width           =   735
      End
      Begin VB.Image iRelevant 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   0
         Left            =   60
         Picture         =   "frmEditBacteriologyNew.frx":08CA
         Top             =   120
         Width           =   480
      End
      Begin VB.Image iRelevant 
         Height          =   480
         Index           =   1
         Left            =   1740
         Picture         =   "frmEditBacteriologyNew.frx":0BD4
         Top             =   120
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "MRU"
         Height          =   195
         Index           =   91
         Left            =   150
         TabIndex        =   279
         Top             =   1110
         Width           =   375
      End
      Begin VB.Image imgLast 
         Height          =   300
         Left            =   2250
         Picture         =   "frmEditBacteriologyNew.frx":0EDE
         Stretch         =   -1  'True
         ToolTipText     =   "Find Last Record"
         Top             =   150
         Width           =   300
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   274
      Top             =   9840
      Width           =   14625
      _ExtentX        =   25797
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   5
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "User Name"
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            TextSave        =   "07/10/2024"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Todays Date"
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   4410
            MinWidth        =   4410
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Demographic Check"
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Run Date"
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   2
            Bevel           =   0
            Object.Width           =   12347
            MinWidth        =   12347
            Text            =   "Custom Software Ltd"
            TextSave        =   "Custom Software Ltd"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton bViewBB 
      Caption         =   "Transfusion Details"
      Enabled         =   0   'False
      Height          =   840
      Left            =   14730
      Picture         =   "frmEditBacteriologyNew.frx":1320
      Style           =   1  'Graphical
      TabIndex        =   159
      Top             =   1950
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.CommandButton cmdViewReports 
      Caption         =   "Reports"
      Height          =   690
      Left            =   13170
      Picture         =   "frmEditBacteriologyNew.frx":1BEA
      Style           =   1  'Graphical
      TabIndex        =   158
      ToolTipText     =   "View Printed && Faxed Reports"
      Top             =   3990
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.CommandButton cmdArchive 
      BackColor       =   &H0000FFFF&
      Caption         =   "Archived Entries"
      Height          =   795
      Left            =   13170
      Picture         =   "frmEditBacteriologyNew.frx":1EF4
      Style           =   1  'Graphical
      TabIndex        =   150
      Top             =   1020
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.CommandButton cmdPhone 
      Caption         =   "Phone Results"
      Height          =   765
      Left            =   13170
      Picture         =   "frmEditBacteriologyNew.frx":21FE
      Style           =   1  'Graphical
      TabIndex        =   97
      ToolTipText     =   "Log as Phoned"
      Top             =   210
      Width           =   1275
   End
   Begin VB.CommandButton cmdSaveHold 
      Caption         =   "Save && Hold"
      Enabled         =   0   'False
      Height          =   645
      Left            =   13170
      Picture         =   "frmEditBacteriologyNew.frx":2640
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   5580
      Width           =   1275
   End
   Begin VB.CommandButton cmdValidateMicro 
      Caption         =   "&Validate"
      Height          =   765
      Left            =   13200
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmEditBacteriologyNew.frx":2A82
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   7020
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.CommandButton cmdSaveMicro 
      Caption         =   "&Save Details"
      Enabled         =   0   'False
      Height          =   705
      Left            =   13170
      Picture         =   "frmEditBacteriologyNew.frx":2EC4
      Style           =   1  'Graphical
      TabIndex        =   89
      Top             =   6240
      Width           =   1275
   End
   Begin VB.CommandButton cmdSetPrinter 
      Height          =   675
      Left            =   13170
      Picture         =   "frmEditBacteriologyNew.frx":352E
      Style           =   1  'Graphical
      TabIndex        =   56
      ToolTipText     =   "Back Colour Normal:Automatic Printer Selection.Back Colour Red:-Forced"
      Top             =   1860
      Width           =   1275
   End
   Begin VB.Timer TimerBar 
      Interval        =   1000
      Left            =   11970
      Top             =   0
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   135
      Left            =   180
      TabIndex        =   51
      Top             =   0
      Width           =   12765
      _ExtentX        =   22516
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   1
      Max             =   600
      Scrolling       =   1
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   2130
      Index           =   0
      Left            =   180
      TabIndex        =   30
      Top             =   135
      Width           =   12765
      Begin VB.CommandButton cmdPrintBarcode 
         Height          =   530
         Left            =   2700
         Picture         =   "frmEditBacteriologyNew.frx":3838
         Style           =   1  'Graphical
         TabIndex        =   518
         Top             =   780
         Width           =   500
      End
      Begin VB.CommandButton cmdPatientNotePad 
         Height          =   500
         Left            =   2700
         Picture         =   "frmEditBacteriologyNew.frx":3C18
         Style           =   1  'Graphical
         TabIndex        =   516
         Tag             =   "bprint"
         Top             =   240
         Width           =   500
      End
      Begin VB.CommandButton cmdReleasetoWard 
         Caption         =   "Release to Ward/GP"
         Height          =   480
         Left            =   1380
         Style           =   1  'Graphical
         TabIndex        =   512
         Top             =   1575
         Width           =   1200
      End
      Begin VB.CommandButton cmdReleaseReport 
         Caption         =   "Release to Consultant"
         Height          =   480
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   511
         Top             =   1575
         Width           =   1200
      End
      Begin VB.CommandButton cmdDartViewer 
         Height          =   390
         Left            =   7245
         Picture         =   "frmEditBacteriologyNew.frx":44E2
         Style           =   1  'Graphical
         TabIndex        =   290
         Top             =   150
         Width           =   375
      End
      Begin VB.TextBox txtAandE 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   4860
         TabIndex        =   2
         Tag             =   "A and E Number"
         ToolTipText     =   "A & E Number"
         Top             =   570
         Width           =   1635
      End
      Begin VB.TextBox txtName 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   6510
         MaxLength       =   30
         TabIndex        =   3
         Tag             =   "tName"
         Top             =   570
         Width           =   4395
      End
      Begin VB.CommandButton cmdAddToConsultantList 
         Caption         =   "Remove from  Consultant List"
         Height          =   300
         Left            =   0
         TabIndex        =   96
         Top             =   1260
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.ComboBox cmbConsultantVal 
         Height          =   315
         Left            =   570
         TabIndex        =   94
         Text            =   "cmbConsultantVal"
         Top             =   1710
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.TextBox txtChart 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3360
         MaxLength       =   8
         TabIndex        =   1
         Top             =   570
         Width           =   1485
      End
      Begin VB.TextBox txtDoB 
         Height          =   285
         Left            =   11490
         MaxLength       =   10
         TabIndex        =   4
         Top             =   450
         Width           =   1155
      End
      Begin VB.TextBox txtAge 
         Height          =   285
         Left            =   11490
         MaxLength       =   4
         TabIndex        =   15
         Top             =   765
         Width           =   1155
      End
      Begin VB.TextBox txtSex 
         Height          =   285
         Left            =   11490
         MaxLength       =   6
         TabIndex        =   5
         Top             =   1080
         Width           =   1185
      End
      Begin VB.CommandButton bsearch 
         Appearance      =   0  'Flat
         Caption         =   "Searc&h"
         Height          =   345
         Left            =   9450
         TabIndex        =   14
         Top             =   180
         Width           =   675
      End
      Begin VB.CommandButton bDoB 
         Appearance      =   0  'Flat
         Caption         =   "Search"
         Height          =   285
         Left            =   11985
         TabIndex        =   31
         Top             =   135
         Width           =   645
      End
      Begin VB.Label lblAandE 
         Caption         =   "A and E"
         Height          =   225
         Left            =   5085
         TabIndex        =   281
         Top             =   300
         Width           =   615
      End
      Begin VB.Label lblNameTitle 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         Height          =   195
         Left            =   6600
         TabIndex        =   157
         Top             =   360
         Width           =   450
      End
      Begin VB.Label lblABsInUse 
         BorderStyle     =   1  'Fixed Single
         Height          =   645
         Left            =   10500
         TabIndex        =   73
         Top             =   1395
         Width           =   2235
      End
      Begin VB.Label Label44 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2730
         TabIndex        =   72
         Top             =   1710
         Width           =   7545
      End
      Begin VB.Label lblSiteDetails 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2730
         TabIndex        =   68
         Top             =   1350
         Width           =   7545
      End
      Begin VB.Label lblChartNumber 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Monaghan Chart #"
         Height          =   285
         Left            =   3360
         TabIndex        =   55
         ToolTipText     =   "Click to change Location"
         Top             =   270
         Width           =   1485
      End
      Begin VB.Label lblAddWardGP 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3360
         TabIndex        =   54
         Top             =   1050
         Width           =   7545
      End
      Begin VB.Label lNoPrevious 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "No Previous Details"
         ForeColor       =   &H0000FFFF&
         Height          =   450
         Left            =   8010
         TabIndex        =   52
         Top             =   120
         Visible         =   0   'False
         Width           =   960
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "D.o.B"
         Height          =   195
         Index           =   88
         Left            =   11040
         TabIndex        =   34
         Top             =   480
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Age"
         Height          =   195
         Index           =   89
         Left            =   11130
         TabIndex        =   33
         Top             =   795
         Width           =   285
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Sex"
         Height          =   195
         Index           =   90
         Left            =   11160
         TabIndex        =   32
         Top             =   1110
         Width           =   270
      End
   End
   Begin VB.CommandButton bHistory 
      Caption         =   "&History"
      Height          =   675
      Left            =   13170
      Picture         =   "frmEditBacteriologyNew.frx":4DAC
      Style           =   1  'Graphical
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   7920
      Width           =   1275
   End
   Begin VB.CommandButton bPrint 
      Caption         =   "&Print"
      Height          =   615
      Left            =   13170
      Picture         =   "frmEditBacteriologyNew.frx":51EE
      Style           =   1  'Graphical
      TabIndex        =   19
      Tag             =   "bprint"
      Top             =   3300
      Width           =   1275
   End
   Begin VB.CommandButton bFAX 
      Caption         =   "FAX"
      Height          =   825
      Index           =   0
      Left            =   13170
      Picture         =   "frmEditBacteriologyNew.frx":5858
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   4710
      Width           =   1275
   End
   Begin VB.CommandButton bCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   645
      Left            =   13170
      Picture         =   "frmEditBacteriologyNew.frx":5C9A
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   8820
      Width           =   1275
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7185
      Left            =   180
      TabIndex        =   17
      Top             =   2280
      Width           =   12795
      _ExtentX        =   22569
      _ExtentY        =   12674
      _Version        =   393216
      Style           =   1
      Tabs            =   14
      Tab             =   12
      TabsPerRow      =   14
      TabHeight       =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Demographics"
      TabPicture(0)   =   "frmEditBacteriologyNew.frx":6304
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdViewExternal"
      Tab(0).Control(1)=   "cmdDemoVal"
      Tab(0).Control(2)=   "cmdCopyFromPrevious"
      Tab(0).Control(3)=   "Frame14"
      Tab(0).Control(4)=   "Frame13"
      Tab(0).Control(5)=   "Frame12"
      Tab(0).Control(6)=   "cmdOrderTests"
      Tab(0).Control(7)=   "cmdSaveInc"
      Tab(0).Control(8)=   "Frame4"
      Tab(0).Control(9)=   "cmdSaveDemographics"
      Tab(0).Control(10)=   "fraDate"
      Tab(0).Control(11)=   "Frame5"
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "Urine"
      TabPicture(1)   =   "frmEditBacteriologyNew.frx":6320
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdIQ200Repeats"
      Tab(1).Control(1)=   "fraIQ200"
      Tab(1).Control(2)=   "fraValid(1)"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Identification"
      TabPicture(2)   =   "frmEditBacteriologyNew.frx":633C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdLock(2)"
      Tab(2).Control(1)=   "cmdObserva(0)"
      Tab(2).Control(2)=   "FrameExtras(4)"
      Tab(2).Control(3)=   "FrameExtras(3)"
      Tab(2).Control(4)=   "FrameExtras(2)"
      Tab(2).Control(5)=   "FrameExtras(1)"
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "Faeces"
      TabPicture(3)   =   "frmEditBacteriologyNew.frx":6358
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cmdObserva(2)"
      Tab(3).Control(1)=   "SSPanel1"
      Tab(3).Control(2)=   "grdDay(1)"
      Tab(3).Control(3)=   "udHistoricalFaecesView"
      Tab(3).Control(4)=   "grdDay(2)"
      Tab(3).Control(5)=   "grdDay(3)"
      Tab(3).Control(6)=   "lblViewOrganism"
      Tab(3).Control(7)=   "Label1(58)"
      Tab(3).ControlCount=   8
      TabCaption(4)   =   "C && S"
      TabPicture(4)   =   "frmEditBacteriologyNew.frx":6374
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "imgSquareCross"
      Tab(4).Control(1)=   "imgSquareTick"
      Tab(4).Control(2)=   "lblPrinted(4)"
      Tab(4).Control(3)=   "lblValid(4)"
      Tab(4).Control(4)=   "fraValid(4)"
      Tab(4).Control(5)=   "grdAB(1)"
      Tab(4).Control(6)=   "grdAB(2)"
      Tab(4).Control(7)=   "grdAB(3)"
      Tab(4).Control(8)=   "grdAB(4)"
      Tab(4).ControlCount=   9
      TabCaption(5)   =   "FOB"
      TabPicture(5)   =   "frmEditBacteriologyNew.frx":6390
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "fraValid(5)"
      Tab(5).Control(1)=   "lblValid(5)"
      Tab(5).Control(2)=   "lblPrinted(5)"
      Tab(5).ControlCount=   3
      TabCaption(6)   =   "Rota/Adeno"
      TabPicture(6)   =   "frmEditBacteriologyNew.frx":63AC
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "fraValid(6)"
      Tab(6).Control(1)=   "lblValid(6)"
      Tab(6).Control(2)=   "lblPrinted(6)"
      Tab(6).ControlCount=   3
      TabCaption(7)   =   "Red/Sub"
      TabPicture(7)   =   "frmEditBacteriologyNew.frx":63C8
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "fraValid(7)"
      Tab(7).Control(1)=   "cmdLock(7)"
      Tab(7).Control(2)=   "lblValid(7)"
      Tab(7).Control(3)=   "lblPrinted(7)"
      Tab(7).ControlCount=   4
      TabCaption(8)   =   "RSV"
      TabPicture(8)   =   "frmEditBacteriologyNew.frx":63E4
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "lblPrinted(8)"
      Tab(8).Control(1)=   "lblValid(8)"
      Tab(8).Control(2)=   "cmdLock(8)"
      Tab(8).Control(3)=   "fraValid(8)"
      Tab(8).ControlCount=   4
      TabCaption(9)   =   "Fluids"
      TabPicture(9)   =   "frmEditBacteriologyNew.frx":6400
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "fraValid(9)"
      Tab(9).Control(1)=   "lblPrinted(9)"
      Tab(9).Control(2)=   "lblValid(9)"
      Tab(9).ControlCount=   3
      TabCaption(10)  =   "C.diff"
      TabPicture(10)  =   "frmEditBacteriologyNew.frx":641C
      Tab(10).ControlEnabled=   0   'False
      Tab(10).Control(0)=   "txtCDiffMSC"
      Tab(10).Control(1)=   "fraValid(10)"
      Tab(10).Control(2)=   "Label2"
      Tab(10).Control(3)=   "lblValidatedBy"
      Tab(10).Control(4)=   "lblValid(10)"
      Tab(10).Control(5)=   "lblPrinted(10)"
      Tab(10).ControlCount=   6
      TabCaption(11)  =   "O/P"
      TabPicture(11)  =   "frmEditBacteriologyNew.frx":6438
      Tab(11).ControlEnabled=   0   'False
      Tab(11).Control(0)=   "fraValid(11)"
      Tab(11).Control(1)=   "lblValid(11)"
      Tab(11).Control(2)=   "lblPrinted(11)"
      Tab(11).ControlCount=   3
      TabCaption(12)  =   "Blood Culture"
      TabPicture(12)  =   "frmEditBacteriologyNew.frx":6454
      Tab(12).ControlEnabled=   -1  'True
      Tab(12).Control(0)=   "lblPrinted(12)"
      Tab(12).Control(0).Enabled=   0   'False
      Tab(12).Control(1)=   "cmdLock(12)"
      Tab(12).Control(1).Enabled=   0   'False
      Tab(12).Control(2)=   "fraBC"
      Tab(12).Control(2).Enabled=   0   'False
      Tab(12).ControlCount=   3
      TabCaption(13)  =   "H.Pylori"
      TabPicture(13)  =   "frmEditBacteriologyNew.frx":6470
      Tab(13).ControlEnabled=   0   'False
      Tab(13).Control(0)=   "lblPrinted(13)"
      Tab(13).Control(1)=   "lblValid(13)"
      Tab(13).Control(2)=   "fraValid(13)"
      Tab(13).ControlCount=   3
      Begin VB.CommandButton cmdLock 
         Caption         =   "Unlock Result"
         Height          =   975
         Index           =   2
         Left            =   -63390
         Picture         =   "frmEditBacteriologyNew.frx":648C
         Style           =   1  'Graphical
         TabIndex        =   514
         Top             =   3375
         Width           =   1035
      End
      Begin VB.TextBox txtCDiffMSC 
         BackColor       =   &H00FFFFFF&
         Height          =   1005
         Left            =   -73740
         MultiLine       =   -1  'True
         TabIndex        =   503
         Text            =   "frmEditBacteriologyNew.frx":68CE
         Top             =   5280
         Width           =   6705
      End
      Begin VB.CommandButton cmdIQ200Repeats 
         BackColor       =   &H0086C0FF&
         Caption         =   "View Repeat"
         Height          =   735
         Left            =   -63900
         Picture         =   "frmEditBacteriologyNew.frx":68E9
         Style           =   1  'Graphical
         TabIndex        =   489
         Top             =   5460
         Width           =   1185
      End
      Begin VB.Frame fraIQ200 
         Caption         =   "Urinalysis"
         Height          =   4455
         Left            =   -66840
         TabIndex        =   482
         Top             =   690
         Width           =   4545
         Begin VB.CommandButton cmdDeleteIQ200 
            Height          =   375
            Left            =   4020
            Picture         =   "frmEditBacteriologyNew.frx":71B3
            Style           =   1  'Graphical
            TabIndex        =   490
            ToolTipText     =   "Remove IQ200 results for this sample"
            Top             =   300
            Width           =   375
         End
         Begin MSFlexGridLib.MSFlexGrid grdIQ200 
            Height          =   4065
            Left            =   120
            TabIndex        =   483
            Top             =   270
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   7170
            _Version        =   393216
            Rows            =   1
            Cols            =   4
            FixedCols       =   0
            ForeColor       =   -2147483635
            BackColorFixed  =   -2147483647
            ForeColorFixed  =   -2147483624
            ScrollTrack     =   -1  'True
            FormatString    =   "|<Name                         |<Result            |<Unit     "
         End
      End
      Begin VB.CommandButton cmdObserva 
         Caption         =   " Observa"
         Height          =   1035
         Index           =   2
         Left            =   -65760
         Picture         =   "frmEditBacteriologyNew.frx":773D
         Style           =   1  'Graphical
         TabIndex        =   470
         Top             =   5730
         Width           =   1035
      End
      Begin VB.CommandButton cmdObserva 
         Caption         =   " Observa"
         Height          =   1155
         Index           =   0
         Left            =   -63390
         Picture         =   "frmEditBacteriologyNew.frx":8607
         Style           =   1  'Graphical
         TabIndex        =   459
         Top             =   2100
         Width           =   1035
      End
      Begin VB.Frame fraBC 
         Height          =   4755
         Left            =   360
         TabIndex        =   452
         Top             =   780
         Width           =   10665
         Begin VB.CommandButton cmdBloodCulture 
            BackColor       =   &H0080FFFF&
            Caption         =   "Identification"
            Height          =   945
            Index           =   1
            Left            =   3660
            Style           =   1  'Graphical
            TabIndex        =   515
            Top             =   3660
            Width           =   2385
         End
         Begin VB.CommandButton cmdVitek 
            Caption         =   "Order on Vitek"
            Height          =   495
            Index           =   2
            Left            =   8580
            TabIndex        =   473
            Top             =   2460
            Width           =   1605
         End
         Begin VB.CommandButton cmdVitek 
            Caption         =   "Order on Vitek"
            Height          =   495
            Index           =   1
            Left            =   8580
            TabIndex        =   472
            Top             =   1830
            Width           =   1605
         End
         Begin VB.CommandButton cmdVitek 
            Caption         =   "Order on Vitek"
            Height          =   495
            Index           =   0
            Left            =   8580
            TabIndex        =   471
            Top             =   1200
            Width           =   1605
         End
         Begin VB.CommandButton cmdBloodCulture 
            BackColor       =   &H0080FFFF&
            Caption         =   "Culture and Sensitivities"
            Height          =   945
            Index           =   0
            Left            =   6120
            Style           =   1  'Graphical
            TabIndex        =   461
            Top             =   3660
            Width           =   2385
         End
         Begin MSFlexGridLib.MSFlexGrid gBC 
            Height          =   3135
            Left            =   120
            TabIndex        =   453
            Top             =   450
            Width           =   8385
            _ExtentX        =   14790
            _ExtentY        =   5530
            _Version        =   393216
            Cols            =   6
            FixedCols       =   0
            RowHeightMin    =   600
            BackColor       =   -2147483634
            ForeColor       =   -2147483635
            BackColorFixed  =   -2147483647
            ForeColorFixed  =   -2147483624
            WordWrap        =   -1  'True
            AllowBigSelection=   0   'False
            FocusRect       =   0
            HighLight       =   0
            GridLines       =   3
            GridLinesFixed  =   3
            ScrollBars      =   0
            AllowUserResizing=   1
            FormatString    =   "^Run Date/Time |^Bottle Number |^Type of Test  |^Result                  |^Hours to Detection |^Valid  "
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label lblBcStatus 
            Caption         =   "Status"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   180
            TabIndex        =   513
            Top             =   180
            Width           =   8235
         End
      End
      Begin VB.CommandButton cmdLock 
         Caption         =   "Unlock Result"
         Height          =   795
         Index           =   12
         Left            =   11250
         Picture         =   "frmEditBacteriologyNew.frx":94D1
         Style           =   1  'Graphical
         TabIndex        =   451
         Top             =   2490
         Width           =   1425
      End
      Begin VB.Frame FrameExtras 
         Caption         =   "Organism 4"
         ForeColor       =   &H00C000C0&
         Height          =   6495
         Index           =   4
         Left            =   -66000
         TabIndex        =   424
         Top             =   480
         Width           =   2505
         Begin VB.TextBox txtCoagulase 
            Height          =   285
            Index           =   4
            Left            =   900
            TabIndex        =   433
            Tag             =   "Coa"
            Top             =   1530
            Width           =   1515
         End
         Begin VB.TextBox txtCatalase 
            Height          =   285
            Index           =   4
            Left            =   900
            TabIndex        =   432
            Tag             =   "Cat"
            Top             =   1830
            Width           =   1515
         End
         Begin VB.TextBox txtOxidase 
            Height          =   285
            Index           =   4
            Left            =   900
            TabIndex        =   431
            Tag             =   "Oxi"
            Top             =   2130
            Width           =   1515
         End
         Begin VB.TextBox txtReincubation 
            Height          =   285
            Index           =   4
            Left            =   900
            TabIndex        =   430
            Tag             =   "Rei"
            Top             =   2430
            Width           =   1515
         End
         Begin VB.ComboBox cmbGram 
            Height          =   315
            Index           =   4
            Left            =   900
            Sorted          =   -1  'True
            TabIndex        =   429
            Top             =   270
            Width           =   1515
         End
         Begin VB.ComboBox cmbWetPrep 
            Height          =   315
            Index           =   4
            Left            =   900
            TabIndex        =   428
            Top             =   900
            Width           =   1515
         End
         Begin VB.TextBox txtNotes 
            Height          =   3645
            Index           =   4
            Left            =   90
            MultiLine       =   -1  'True
            TabIndex        =   427
            Top             =   2730
            Width           =   2355
         End
         Begin VB.TextBox txtIndole 
            Height          =   285
            Index           =   4
            Left            =   900
            TabIndex        =   426
            Top             =   1230
            Width           =   1515
         End
         Begin VB.TextBox txtZN 
            Height          =   285
            Index           =   4
            Left            =   900
            TabIndex        =   425
            Top             =   600
            Width           =   1515
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Gram Stain"
            Height          =   195
            Index           =   60
            Left            =   90
            TabIndex        =   441
            Top             =   330
            Width           =   780
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Coagulase"
            Height          =   195
            Index           =   59
            Left            =   120
            TabIndex        =   440
            Top             =   1560
            Width           =   750
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Catalase"
            Height          =   195
            Index           =   57
            Left            =   255
            TabIndex        =   439
            Top             =   1860
            Width           =   615
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Oxidase"
            Height          =   195
            Index           =   56
            Left            =   300
            TabIndex        =   438
            Top             =   2160
            Width           =   570
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Wet Prep"
            Height          =   195
            Index           =   55
            Left            =   195
            TabIndex        =   437
            Top             =   930
            Width           =   675
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Reinc"
            Height          =   195
            Index           =   54
            Left            =   450
            TabIndex        =   436
            Top             =   2430
            Width           =   420
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Indole"
            Height          =   195
            Index           =   4
            Left            =   435
            TabIndex        =   435
            Top             =   1260
            Width           =   435
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "ZN Stain"
            Height          =   195
            Index           =   4
            Left            =   225
            TabIndex        =   434
            Top             =   630
            Width           =   630
         End
      End
      Begin VB.Frame FrameExtras 
         Caption         =   "Organism 3"
         ForeColor       =   &H00C000C0&
         Height          =   6495
         Index           =   3
         Left            =   -68640
         TabIndex        =   406
         Top             =   480
         Width           =   2505
         Begin VB.TextBox txtCoagulase 
            Height          =   285
            Index           =   3
            Left            =   900
            TabIndex        =   415
            Tag             =   "Coa"
            Top             =   1530
            Width           =   1515
         End
         Begin VB.TextBox txtCatalase 
            Height          =   285
            Index           =   3
            Left            =   900
            TabIndex        =   414
            Tag             =   "Cat"
            Top             =   1830
            Width           =   1515
         End
         Begin VB.TextBox txtOxidase 
            Height          =   285
            Index           =   3
            Left            =   900
            TabIndex        =   413
            Tag             =   "Oxi"
            Top             =   2130
            Width           =   1515
         End
         Begin VB.TextBox txtReincubation 
            Height          =   285
            Index           =   3
            Left            =   900
            TabIndex        =   412
            Tag             =   "Rei"
            Top             =   2430
            Width           =   1515
         End
         Begin VB.ComboBox cmbGram 
            Height          =   315
            Index           =   3
            Left            =   900
            Sorted          =   -1  'True
            TabIndex        =   411
            Top             =   270
            Width           =   1515
         End
         Begin VB.ComboBox cmbWetPrep 
            Height          =   315
            Index           =   3
            Left            =   900
            TabIndex        =   410
            Top             =   900
            Width           =   1515
         End
         Begin VB.TextBox txtNotes 
            Height          =   3645
            Index           =   3
            Left            =   90
            MultiLine       =   -1  'True
            TabIndex        =   409
            Top             =   2730
            Width           =   2355
         End
         Begin VB.TextBox txtIndole 
            Height          =   285
            Index           =   3
            Left            =   900
            TabIndex        =   408
            Top             =   1230
            Width           =   1515
         End
         Begin VB.TextBox txtZN 
            Height          =   285
            Index           =   3
            Left            =   900
            TabIndex        =   407
            Top             =   600
            Width           =   1515
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Gram Stain"
            Height          =   195
            Index           =   53
            Left            =   90
            TabIndex        =   423
            Top             =   330
            Width           =   780
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Coagulase"
            Height          =   195
            Index           =   52
            Left            =   120
            TabIndex        =   422
            Top             =   1560
            Width           =   750
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Catalase"
            Height          =   195
            Index           =   51
            Left            =   255
            TabIndex        =   421
            Top             =   1860
            Width           =   615
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Oxidase"
            Height          =   195
            Index           =   50
            Left            =   300
            TabIndex        =   420
            Top             =   2160
            Width           =   570
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Wet Prep"
            Height          =   195
            Index           =   49
            Left            =   195
            TabIndex        =   419
            Top             =   930
            Width           =   675
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Reinc"
            Height          =   195
            Index           =   48
            Left            =   450
            TabIndex        =   418
            Top             =   2430
            Width           =   420
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Indole"
            Height          =   195
            Index           =   3
            Left            =   435
            TabIndex        =   417
            Top             =   1260
            Width           =   435
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "ZN Stain"
            Height          =   195
            Index           =   3
            Left            =   255
            TabIndex        =   416
            Top             =   630
            Width           =   630
         End
      End
      Begin VB.Frame FrameExtras 
         Caption         =   "Organism 2"
         ForeColor       =   &H00C000C0&
         Height          =   6495
         Index           =   2
         Left            =   -71280
         TabIndex        =   388
         Top             =   480
         Width           =   2505
         Begin VB.TextBox txtCoagulase 
            Height          =   285
            Index           =   2
            Left            =   900
            TabIndex        =   397
            Tag             =   "Coa"
            Top             =   1530
            Width           =   1515
         End
         Begin VB.TextBox txtCatalase 
            Height          =   285
            Index           =   2
            Left            =   900
            TabIndex        =   396
            Tag             =   "Cat"
            Top             =   1830
            Width           =   1515
         End
         Begin VB.TextBox txtOxidase 
            Height          =   285
            Index           =   2
            Left            =   900
            TabIndex        =   395
            Tag             =   "Oxi"
            Top             =   2130
            Width           =   1515
         End
         Begin VB.TextBox txtReincubation 
            Height          =   285
            Index           =   2
            Left            =   900
            TabIndex        =   394
            Tag             =   "Rei"
            Top             =   2430
            Width           =   1515
         End
         Begin VB.ComboBox cmbGram 
            Height          =   315
            Index           =   2
            Left            =   900
            Sorted          =   -1  'True
            TabIndex        =   393
            Top             =   270
            Width           =   1515
         End
         Begin VB.ComboBox cmbWetPrep 
            Height          =   315
            Index           =   2
            Left            =   900
            TabIndex        =   392
            Top             =   900
            Width           =   1515
         End
         Begin VB.TextBox txtNotes 
            Height          =   3645
            Index           =   2
            Left            =   90
            MultiLine       =   -1  'True
            TabIndex        =   391
            Top             =   2730
            Width           =   2355
         End
         Begin VB.TextBox txtIndole 
            Height          =   285
            Index           =   2
            Left            =   900
            TabIndex        =   390
            Top             =   1230
            Width           =   1515
         End
         Begin VB.TextBox txtZN 
            Height          =   285
            Index           =   2
            Left            =   900
            TabIndex        =   389
            Top             =   600
            Width           =   1515
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Gram Stain"
            Height          =   195
            Index           =   47
            Left            =   90
            TabIndex        =   405
            Top             =   330
            Width           =   780
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Coagulase"
            Height          =   195
            Index           =   46
            Left            =   120
            TabIndex        =   404
            Top             =   1560
            Width           =   750
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Catalase"
            Height          =   195
            Index           =   45
            Left            =   255
            TabIndex        =   403
            Top             =   1860
            Width           =   615
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Oxidase"
            Height          =   195
            Index           =   44
            Left            =   300
            TabIndex        =   402
            Top             =   2160
            Width           =   570
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Wet Prep"
            Height          =   195
            Index           =   43
            Left            =   195
            TabIndex        =   401
            Top             =   930
            Width           =   675
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Reinc"
            Height          =   195
            Index           =   42
            Left            =   450
            TabIndex        =   400
            Top             =   2430
            Width           =   420
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Indole"
            Height          =   195
            Index           =   2
            Left            =   435
            TabIndex        =   399
            Top             =   1260
            Width           =   435
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "ZN Stain"
            Height          =   195
            Index           =   2
            Left            =   255
            TabIndex        =   398
            Top             =   630
            Width           =   630
         End
      End
      Begin VB.Frame fraValid 
         BorderStyle     =   0  'None
         Caption         =   "Frame9"
         Height          =   2505
         Index           =   8
         Left            =   -71940
         TabIndex        =   377
         Top             =   1800
         Width           =   4035
         Begin VB.Frame fraRSV 
            Caption         =   "RSV"
            Height          =   1275
            Left            =   150
            TabIndex        =   378
            Top             =   570
            Width           =   3615
            Begin VB.Label lblRSV 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   13.5
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Left            =   780
               TabIndex        =   379
               Top             =   480
               Width           =   2205
            End
         End
      End
      Begin VB.CommandButton cmdViewExternal 
         Caption         =   "View/Edit Externals"
         Height          =   945
         Left            =   -63540
         Picture         =   "frmEditBacteriologyNew.frx":9913
         Style           =   1  'Graphical
         TabIndex        =   368
         Tag             =   "bOrder"
         Top             =   5310
         Width           =   1035
      End
      Begin VB.Frame fraValid 
         BorderStyle     =   0  'None
         Caption         =   "Frame6"
         Height          =   4005
         Index           =   7
         Left            =   -71310
         TabIndex        =   360
         Top             =   1350
         Width           =   2655
         Begin VB.Frame fraRedSub 
            Caption         =   "Reducing Substances"
            Height          =   2475
            Left            =   270
            TabIndex        =   361
            Top             =   420
            Width           =   2175
            Begin VB.CheckBox chkRS 
               Caption         =   "2 %"
               Height          =   195
               Index           =   5
               Left            =   720
               TabIndex        =   367
               Top             =   1950
               Width           =   555
            End
            Begin VB.CheckBox chkRS 
               Caption         =   "1 %"
               Height          =   195
               Index           =   4
               Left            =   720
               TabIndex        =   366
               Top             =   1650
               Width           =   555
            End
            Begin VB.CheckBox chkRS 
               Caption         =   "0.75 %"
               Height          =   195
               Index           =   3
               Left            =   720
               TabIndex        =   365
               Top             =   1350
               Width           =   795
            End
            Begin VB.CheckBox chkRS 
               Caption         =   "0.5 %"
               Height          =   195
               Index           =   2
               Left            =   720
               TabIndex        =   364
               Top             =   1050
               Width           =   705
            End
            Begin VB.CheckBox chkRS 
               Caption         =   "0.25 %"
               Height          =   195
               Index           =   1
               Left            =   720
               TabIndex        =   363
               Top             =   750
               Width           =   795
            End
            Begin VB.CheckBox chkRS 
               Caption         =   "0 %"
               Height          =   195
               Index           =   0
               Left            =   720
               TabIndex        =   362
               Top             =   450
               Width           =   555
            End
         End
      End
      Begin VB.Frame fraValid 
         BorderStyle     =   0  'None
         Height          =   6735
         Index           =   9
         Left            =   -74850
         TabIndex        =   298
         Top             =   360
         Width           =   11385
         Begin VB.Frame fraCSF 
            Height          =   6585
            Left            =   150
            TabIndex        =   300
            Top             =   60
            Width           =   9945
            Begin VB.ComboBox cmbZN 
               Height          =   315
               Left            =   1380
               TabIndex        =   457
               Text            =   "cmbZN"
               Top             =   1560
               Width           =   2745
            End
            Begin VB.Frame Frame3 
               Caption         =   "BAT Screen"
               Height          =   1665
               Index           =   2
               Left            =   7260
               TabIndex        =   446
               Top             =   4560
               Width           =   2625
               Begin VB.TextBox txtBATComments 
                  Height          =   585
                  Left            =   90
                  MaxLength       =   50
                  TabIndex        =   450
                  Top             =   1005
                  Width           =   2385
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "BAT Screen Result"
                  Height          =   195
                  Index           =   25
                  Left            =   90
                  TabIndex        =   449
                  Top             =   240
                  Width           =   1365
               End
               Begin VB.Label lblBATResult 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   255
                  Left            =   90
                  TabIndex        =   448
                  Top             =   435
                  Width           =   2385
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Comments"
                  Height          =   195
                  Index           =   24
                  Left            =   90
                  TabIndex        =   447
                  Top             =   810
                  Width           =   735
               End
            End
            Begin VB.Frame Frame3 
               Caption         =   "Antigen Tests"
               Height          =   825
               Index           =   3
               Left            =   7260
               TabIndex        =   372
               Top             =   3570
               Width           =   2625
               Begin VB.Label lblLegionellaAT 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   255
                  Left            =   1470
                  TabIndex        =   376
                  Top             =   480
                  Width           =   1035
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Legionella AT"
                  Height          =   195
                  Index           =   9
                  Left            =   450
                  TabIndex        =   375
                  Top             =   510
                  Width           =   975
               End
               Begin VB.Label lblPneuAT 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   255
                  Left            =   1470
                  TabIndex        =   374
                  Top             =   210
                  Width           =   1035
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Pneumococcal AT"
                  Height          =   195
                  Index           =   5
                  Left            =   120
                  TabIndex        =   373
                  Top             =   240
                  Width           =   1320
               End
            End
            Begin VB.Frame Frame3 
               Caption         =   "KOH Preparation"
               Height          =   825
               Index           =   1
               Left            =   4800
               TabIndex        =   369
               Top             =   3570
               Width           =   2385
               Begin VB.CheckBox chkFungal 
                  Caption         =   "No Fungal Elements Seen"
                  Height          =   255
                  Index           =   1
                  Left            =   120
                  TabIndex        =   371
                  Top             =   480
                  Width           =   2145
               End
               Begin VB.CheckBox chkFungal 
                  Caption         =   "Fungal Elements Seen"
                  Height          =   225
                  Index           =   0
                  Left            =   120
                  TabIndex        =   370
                  Top             =   240
                  Width           =   2175
               End
            End
            Begin VB.ComboBox cmbFluidAppearance 
               Height          =   315
               Index           =   1
               Left            =   1380
               TabIndex        =   359
               Text            =   "cmbFluidAppearance"
               Top             =   510
               Width           =   2745
            End
            Begin VB.ComboBox cmbFluidCrystals 
               Height          =   315
               Left            =   1380
               TabIndex        =   358
               Text            =   "cmbFluidCrystals"
               Top             =   2640
               Width           =   2745
            End
            Begin VB.ComboBox cmbFluidWetPrep 
               Height          =   315
               Left            =   1380
               TabIndex        =   357
               Text            =   "cmbFluidWetPrep"
               Top             =   2280
               Width           =   2745
            End
            Begin VB.TextBox txtFluidComment 
               Height          =   1605
               Left            =   1380
               TabIndex        =   333
               Text            =   "txtFluidComment"
               Top             =   4890
               Width           =   5745
            End
            Begin VB.ComboBox cmbFluidGram 
               Height          =   315
               Index           =   0
               Left            =   1380
               TabIndex        =   332
               Text            =   "cmbFluidGram"
               Top             =   840
               Width           =   2745
            End
            Begin VB.ComboBox cmbFluidGram 
               Height          =   315
               Index           =   1
               Left            =   1380
               TabIndex        =   331
               Text            =   "cmbFluidGram"
               Top             =   1200
               Width           =   2745
            End
            Begin VB.ComboBox cmbFluidLeishmans 
               Height          =   315
               Left            =   1380
               TabIndex        =   330
               Text            =   "cmbFluidLeishmans"
               Top             =   1920
               Width           =   2745
            End
            Begin VB.TextBox txtHaem 
               Height          =   285
               Index           =   0
               Left            =   1380
               MaxLength       =   50
               TabIndex        =   329
               Top             =   3480
               Width           =   885
            End
            Begin VB.TextBox txtHaem 
               Height          =   285
               Index           =   1
               Left            =   2310
               MaxLength       =   50
               TabIndex        =   328
               Top             =   3480
               Width           =   885
            End
            Begin VB.TextBox txtHaem 
               Height          =   285
               Index           =   2
               Left            =   3240
               MaxLength       =   50
               TabIndex        =   327
               Top             =   3480
               Width           =   885
            End
            Begin VB.TextBox txtHaem 
               Height          =   285
               Index           =   3
               Left            =   1380
               MaxLength       =   50
               TabIndex        =   326
               Top             =   3810
               Width           =   885
            End
            Begin VB.TextBox txtHaem 
               Height          =   285
               Index           =   4
               Left            =   2310
               MaxLength       =   50
               TabIndex        =   325
               Top             =   3810
               Width           =   885
            End
            Begin VB.TextBox txtHaem 
               Height          =   285
               Index           =   5
               Left            =   3240
               MaxLength       =   50
               TabIndex        =   324
               Top             =   3810
               Width           =   885
            End
            Begin VB.TextBox txtHaem 
               Height          =   285
               Index           =   6
               Left            =   1380
               MaxLength       =   50
               TabIndex        =   323
               Top             =   4140
               Width           =   885
            End
            Begin VB.TextBox txtHaem 
               Height          =   285
               Index           =   7
               Left            =   2310
               MaxLength       =   50
               TabIndex        =   322
               Top             =   4140
               Width           =   885
            End
            Begin VB.TextBox txtHaem 
               Height          =   285
               Index           =   8
               Left            =   3240
               MaxLength       =   50
               TabIndex        =   321
               Top             =   4140
               Width           =   885
            End
            Begin VB.TextBox txtHaem 
               Height          =   285
               Index           =   9
               Left            =   1380
               MaxLength       =   50
               TabIndex        =   320
               Top             =   4470
               Width           =   885
            End
            Begin VB.TextBox txtHaem 
               Height          =   285
               Index           =   10
               Left            =   2310
               MaxLength       =   50
               TabIndex        =   319
               Top             =   4470
               Width           =   885
            End
            Begin VB.TextBox txtHaem 
               Height          =   285
               Index           =   11
               Left            =   3240
               MaxLength       =   50
               TabIndex        =   318
               Top             =   4470
               Width           =   885
            End
            Begin VB.Frame Frame3 
               Caption         =   "External"
               Height          =   3495
               Index           =   0
               Left            =   4800
               TabIndex        =   302
               Top             =   0
               Width           =   5085
               Begin VB.TextBox txtBioResult 
                  Height          =   285
                  Index           =   7
                  Left            =   2820
                  TabIndex        =   352
                  Top             =   1650
                  Width           =   1245
               End
               Begin VB.TextBox txtBioResult 
                  Height          =   285
                  Index           =   6
                  Left            =   2820
                  TabIndex        =   351
                  Top             =   1350
                  Width           =   1245
               End
               Begin VB.CheckBox chkBio 
                  Caption         =   "Protein"
                  Height          =   195
                  Index           =   7
                  Left            =   4110
                  TabIndex        =   350
                  Top             =   1680
                  Width           =   795
               End
               Begin VB.CheckBox chkBio 
                  Caption         =   "Glucose"
                  Height          =   195
                  Index           =   6
                  Left            =   4110
                  TabIndex        =   349
                  Top             =   1380
                  Width           =   885
               End
               Begin VB.TextBox txtInHouseSID 
                  Height          =   285
                  Left            =   990
                  TabIndex        =   316
                  Top             =   390
                  Width           =   1545
               End
               Begin VB.CheckBox chkBio 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Amylase"
                  Height          =   195
                  Index           =   5
                  Left            =   300
                  TabIndex        =   315
                  Top             =   2880
                  Width           =   885
               End
               Begin VB.CheckBox chkBio 
                  Alignment       =   1  'Right Justify
                  Caption         =   "LDH"
                  Height          =   195
                  Index           =   4
                  Left            =   540
                  TabIndex        =   314
                  Top             =   2580
                  Width           =   645
               End
               Begin VB.CheckBox chkBio 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Globulin"
                  Height          =   195
                  Index           =   3
                  Left            =   300
                  TabIndex        =   313
                  Top             =   2280
                  Width           =   885
               End
               Begin VB.CheckBox chkBio 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Albumin"
                  Height          =   195
                  Index           =   2
                  Left            =   330
                  TabIndex        =   312
                  Top             =   1980
                  Width           =   855
               End
               Begin VB.CheckBox chkBio 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Tot Protein"
                  Height          =   195
                  Index           =   1
                  Left            =   90
                  TabIndex        =   311
                  Top             =   1680
                  Width           =   1095
               End
               Begin VB.CheckBox chkBio 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Glucose"
                  Height          =   195
                  Index           =   0
                  Left            =   300
                  TabIndex        =   310
                  Top             =   1380
                  Width           =   885
               End
               Begin VB.CommandButton cmdOrderInHouse 
                  Caption         =   "Order Tests"
                  Height          =   375
                  Left            =   2850
                  Style           =   1  'Graphical
                  TabIndex        =   309
                  Top             =   2280
                  Width           =   1245
               End
               Begin VB.TextBox txtBioResult 
                  Height          =   285
                  Index           =   0
                  Left            =   1290
                  TabIndex        =   308
                  Top             =   1350
                  Width           =   1245
               End
               Begin VB.TextBox txtBioResult 
                  Height          =   285
                  Index           =   1
                  Left            =   1290
                  TabIndex        =   307
                  Top             =   1650
                  Width           =   1245
               End
               Begin VB.TextBox txtBioResult 
                  Height          =   285
                  Index           =   2
                  Left            =   1290
                  TabIndex        =   306
                  Top             =   1950
                  Width           =   1245
               End
               Begin VB.TextBox txtBioResult 
                  Height          =   285
                  Index           =   3
                  Left            =   1290
                  TabIndex        =   305
                  Top             =   2250
                  Width           =   1245
               End
               Begin VB.TextBox txtBioResult 
                  Height          =   285
                  Index           =   4
                  Left            =   1290
                  TabIndex        =   304
                  Top             =   2550
                  Width           =   1245
               End
               Begin VB.TextBox txtBioResult 
                  Height          =   285
                  Index           =   5
                  Left            =   1290
                  TabIndex        =   303
                  Top             =   2850
                  Width           =   1245
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00FFFF00&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "Fluids"
                  Height          =   255
                  Index           =   118
                  Left            =   990
                  TabIndex        =   354
                  Top             =   960
                  Width           =   1545
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H0080FF80&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "CSF"
                  Height          =   255
                  Index           =   117
                  Left            =   2820
                  TabIndex        =   353
                  Top             =   960
                  Width           =   1470
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  Caption         =   "In House Sample ID"
                  Height          =   405
                  Index           =   115
                  Left            =   180
                  TabIndex        =   317
                  Top             =   330
                  Width           =   735
               End
            End
            Begin VB.ComboBox cmbFluidAppearance 
               Height          =   315
               Index           =   0
               Left            =   1380
               TabIndex        =   301
               Text            =   "cmbFluidAppearance"
               Top             =   150
               Width           =   2745
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Appearance"
               Height          =   195
               Index           =   23
               Left            =   450
               TabIndex        =   479
               Top             =   570
               Width           =   870
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "ZN Stain"
               Height          =   195
               Index           =   4
               Left            =   690
               TabIndex        =   458
               Top             =   1620
               Width           =   630
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Crystals"
               Height          =   195
               Index           =   120
               Left            =   780
               TabIndex        =   356
               Top             =   2670
               Width           =   540
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Wet Prep"
               Height          =   195
               Index           =   119
               Left            =   645
               TabIndex        =   355
               Top             =   2340
               Width           =   675
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Gram Stain"
               Height          =   195
               Index           =   0
               Left            =   540
               TabIndex        =   348
               Top             =   900
               Width           =   780
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Gram Stain"
               Height          =   195
               Index           =   1
               Left            =   540
               TabIndex        =   347
               Top             =   1260
               Width           =   780
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Leishman's Stain"
               Height          =   195
               Index           =   2
               Left            =   135
               TabIndex        =   346
               Top             =   2010
               Width           =   1185
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Specimen"
               Height          =   195
               Index           =   6
               Left            =   615
               TabIndex        =   345
               Top             =   3120
               Width           =   705
            End
            Begin VB.Label Label1 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "     1          2          3"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Index           =   7
               Left            =   1380
               TabIndex        =   344
               Top             =   3030
               Width           =   2730
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "WCC"
               Height          =   195
               Index           =   10
               Left            =   945
               TabIndex        =   343
               Top             =   3855
               Width           =   375
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "RCC"
               Height          =   195
               Index           =   8
               Left            =   990
               TabIndex        =   342
               Top             =   3525
               Width           =   330
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Polymorphic"
               Height          =   195
               Index           =   11
               Left            =   465
               TabIndex        =   341
               Top             =   4185
               Width           =   855
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Mononucleated"
               Height          =   195
               Index           =   12
               Left            =   210
               TabIndex        =   340
               Top             =   4515
               Width           =   1110
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Comment"
               Height          =   195
               Index           =   13
               Left            =   660
               TabIndex        =   339
               Top             =   4950
               Width           =   660
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "/cmm"
               Height          =   195
               Index           =   111
               Left            =   4230
               TabIndex        =   338
               Top             =   3525
               Width           =   405
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "%"
               Height          =   195
               Index           =   112
               Left            =   4230
               TabIndex        =   337
               Top             =   4515
               Width           =   120
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "%"
               Height          =   195
               Index           =   113
               Left            =   4230
               TabIndex        =   336
               Top             =   4185
               Width           =   120
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "/cmm"
               Height          =   195
               Index           =   114
               Left            =   4230
               TabIndex        =   335
               Top             =   3855
               Width           =   405
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Cell Count"
               Height          =   195
               Index           =   116
               Left            =   600
               TabIndex        =   334
               Top             =   210
               Width           =   720
            End
         End
         Begin VB.CommandButton cmdLock 
            Caption         =   "Unlock Result"
            Height          =   945
            Index           =   9
            Left            =   10140
            Picture         =   "frmEditBacteriologyNew.frx":9D55
            Style           =   1  'Graphical
            TabIndex        =   299
            Top             =   2670
            Width           =   1035
         End
      End
      Begin MSFlexGridLib.MSFlexGrid grdAB 
         Height          =   3945
         Index           =   4
         Left            =   -66300
         TabIndex        =   291
         Top             =   1830
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   6959
         _Version        =   393216
         Cols            =   10
         FixedCols       =   0
         BackColor       =   -2147483624
         ForeColor       =   -2147483635
         BackColorFixed  =   -2147483647
         ForeColorFixed  =   -2147483624
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   0
         GridLines       =   2
         FormatString    =   "<Antibiotic             |^S/R|>Rprt|<Result  |<Date/Time            |<Operator||||"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid grdAB 
         Height          =   3945
         Index           =   3
         Left            =   -69060
         TabIndex        =   292
         Top             =   1830
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   6959
         _Version        =   393216
         Cols            =   10
         FixedCols       =   0
         BackColor       =   -2147483624
         ForeColor       =   -2147483635
         BackColorFixed  =   -2147483647
         ForeColorFixed  =   -2147483624
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   0
         GridLines       =   2
         FormatString    =   "<Antibiotic             |^S/R|>Rprt|<Result  |<Date/Time            |<Operator||||"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid grdAB 
         Height          =   3945
         Index           =   2
         Left            =   -71790
         TabIndex        =   293
         Top             =   1830
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   6959
         _Version        =   393216
         Cols            =   10
         FixedCols       =   0
         BackColor       =   -2147483624
         ForeColor       =   -2147483635
         BackColorFixed  =   -2147483647
         ForeColorFixed  =   -2147483624
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   0
         GridLines       =   2
         FormatString    =   "<Antibiotic             |^S/R|>Rprt|<Result  |<Date/Time            |<Operator||||"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid grdAB 
         Height          =   3945
         Index           =   1
         Left            =   -74520
         TabIndex        =   294
         Top             =   1830
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   6959
         _Version        =   393216
         Cols            =   10
         FixedCols       =   0
         BackColor       =   -2147483624
         ForeColor       =   -2147483635
         BackColorFixed  =   -2147483647
         ForeColorFixed  =   -2147483624
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   0
         GridLines       =   2
         FormatString    =   "<Antibiotic             |^S/R|>Rprt|<Result  |<Date/Time            |<Operator||||"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CommandButton cmdDemoVal 
         Caption         =   "&Validate"
         Height          =   735
         Left            =   -69120
         Picture         =   "frmEditBacteriologyNew.frx":A197
         Style           =   1  'Graphical
         TabIndex        =   273
         Top             =   6300
         Width           =   1035
      End
      Begin VB.Frame fraValid 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   1785
         Index           =   13
         Left            =   -71550
         TabIndex        =   269
         Top             =   2220
         Width           =   5895
         Begin VB.CommandButton cmdLock 
            Caption         =   "Unlock Result"
            Height          =   885
            Index           =   13
            Left            =   4230
            Picture         =   "frmEditBacteriologyNew.frx":A4A1
            Style           =   1  'Graphical
            TabIndex        =   272
            Top             =   480
            Width           =   1485
         End
         Begin VB.Frame fraHPylori 
            Caption         =   "H.Pylori Antigen Test"
            Height          =   1575
            Left            =   90
            TabIndex        =   270
            Top             =   120
            Width           =   3945
            Begin VB.Label lblHPylori 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   180
               TabIndex        =   271
               Top             =   630
               Width           =   3585
            End
         End
      End
      Begin VB.Frame fraValid 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   4245
         Index           =   11
         Left            =   -72660
         TabIndex        =   261
         Top             =   1620
         Width           =   8055
         Begin VB.Frame fraOP 
            Caption         =   "Ova / Parasites"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3945
            Left            =   60
            TabIndex        =   263
            Top             =   120
            Width           =   6195
            Begin VB.ComboBox cmbOva 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Index           =   0
               Left            =   510
               TabIndex        =   266
               Top             =   2040
               Width           =   5055
            End
            Begin VB.ComboBox cmbOva 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Index           =   1
               Left            =   510
               TabIndex        =   265
               Top             =   2610
               Width           =   5055
            End
            Begin VB.ComboBox cmbOva 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Index           =   2
               Left            =   510
               TabIndex        =   264
               Top             =   3180
               Width           =   5055
            End
            Begin VB.Label lblGiardia 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2760
               TabIndex        =   505
               Top             =   1320
               Width           =   2805
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Giardia Lambila"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   28
               Left            =   600
               TabIndex        =   504
               Top             =   1350
               Width           =   1650
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Cryptosporidium"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   110
               Left            =   540
               TabIndex        =   268
               Top             =   810
               Width           =   1710
            End
            Begin VB.Label lblCrypto 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2760
               TabIndex        =   267
               Top             =   780
               Width           =   2805
            End
         End
         Begin VB.CommandButton cmdLock 
            Caption         =   "Unlock Result"
            Height          =   795
            Index           =   11
            Left            =   6510
            Picture         =   "frmEditBacteriologyNew.frx":A8E3
            Style           =   1  'Graphical
            TabIndex        =   262
            Top             =   1230
            Width           =   1425
         End
      End
      Begin VB.Frame fraValid 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   3525
         Index           =   10
         Left            =   -73920
         TabIndex        =   256
         Top             =   1380
         Width           =   11385
         Begin VB.Frame fraCDiff 
            Caption         =   "C.diff"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3435
            Left            =   180
            TabIndex        =   258
            Top             =   60
            Width           =   9555
            Begin VB.ComboBox cmbPCR 
               Height          =   315
               Left            =   3720
               TabIndex        =   510
               Text            =   "cmbPCR"
               Top             =   2745
               Width           =   5295
            End
            Begin VB.ComboBox cmbGDH 
               Height          =   315
               Left            =   3720
               TabIndex        =   509
               Text            =   "cmbGDH"
               Top             =   2190
               Width           =   5295
            End
            Begin VB.Label lblTitle 
               AutoSize        =   -1  'True
               Caption         =   "PCR"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   7
               Left            =   780
               TabIndex        =   502
               Top             =   2745
               Width           =   495
            End
            Begin VB.Label lblTitle 
               AutoSize        =   -1  'True
               Caption         =   "GDH"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   6
               Left            =   720
               TabIndex        =   501
               Top             =   2190
               Width           =   555
            End
            Begin VB.Label lblPCR 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   1470
               TabIndex        =   500
               Top             =   2715
               Width           =   2115
            End
            Begin VB.Label lblGDH 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   1470
               TabIndex        =   499
               Top             =   2160
               Width           =   2115
            End
            Begin VB.Label lblCDiffCulture 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   1470
               TabIndex        =   456
               Top             =   810
               Width           =   4995
            End
            Begin VB.Label lblTitle 
               AutoSize        =   -1  'True
               Caption         =   "Culture"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   5
               Left            =   510
               TabIndex        =   455
               Top             =   840
               Width           =   765
            End
            Begin VB.Label lblTitle 
               AutoSize        =   -1  'True
               Caption         =   "Toxin A/B"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   4
               Left            =   270
               TabIndex        =   260
               Top             =   1515
               Width           =   1005
            End
            Begin VB.Label lblToxinA 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   1470
               TabIndex        =   259
               Top             =   1485
               Width           =   4995
            End
         End
         Begin VB.CommandButton cmdLock 
            Caption         =   "Unlock Result"
            Height          =   795
            Index           =   10
            Left            =   9960
            Picture         =   "frmEditBacteriologyNew.frx":AD25
            Style           =   1  'Graphical
            TabIndex        =   257
            Top             =   480
            Width           =   1365
         End
      End
      Begin VB.Frame fraValid 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   2805
         Index           =   6
         Left            =   -72120
         TabIndex        =   249
         Top             =   1470
         Width           =   6285
         Begin VB.CommandButton cmdLock 
            Caption         =   "Unlock Result"
            Height          =   795
            Index           =   6
            Left            =   4710
            Picture         =   "frmEditBacteriologyNew.frx":B167
            Style           =   1  'Graphical
            TabIndex        =   255
            Top             =   1050
            Width           =   1425
         End
         Begin VB.Frame fraRotaAdeno 
            Caption         =   "Rota/Adeno"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2325
            Left            =   60
            TabIndex        =   250
            Top             =   270
            Width           =   4485
            Begin VB.TextBox txtAdeno 
               Alignment       =   2  'Center
               BackColor       =   &H8000000F&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Left            =   945
               TabIndex        =   252
               Top             =   1440
               Width           =   3105
            End
            Begin VB.TextBox txtRota 
               Alignment       =   2  'Center
               BackColor       =   &H8000000F&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Left            =   945
               TabIndex        =   251
               Top             =   660
               Width           =   3105
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Adeno"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   40
               Left            =   120
               TabIndex        =   254
               Top             =   1440
               Width           =   705
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Rota"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   39
               Left            =   270
               TabIndex        =   253
               Top             =   690
               Width           =   525
            End
         End
      End
      Begin VB.Frame fraValid 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   2685
         Index           =   5
         Left            =   -71910
         TabIndex        =   243
         Top             =   1440
         Width           =   6285
         Begin VB.CommandButton cmdLock 
            Caption         =   "Unlock Result"
            Height          =   795
            Index           =   5
            Left            =   4650
            Picture         =   "frmEditBacteriologyNew.frx":B5A9
            Style           =   1  'Graphical
            TabIndex        =   248
            Top             =   930
            Width           =   1425
         End
         Begin VB.Frame fraFOB 
            Caption         =   "Occult Blood"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2265
            Left            =   240
            TabIndex        =   244
            Top             =   240
            Width           =   4245
            Begin VB.Label lblFOB 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Index           =   2
               Left            =   780
               TabIndex        =   247
               Top             =   1560
               Width           =   2595
            End
            Begin VB.Label lblFOB 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Index           =   1
               Left            =   780
               TabIndex        =   246
               Top             =   1020
               Width           =   2595
            End
            Begin VB.Label lblFOB 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Index           =   0
               Left            =   780
               TabIndex        =   245
               Top             =   480
               Width           =   2595
            End
         End
      End
      Begin VB.Frame fraValid 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   6795
         Index           =   4
         Left            =   -74970
         TabIndex        =   197
         Top             =   330
         Width           =   12165
         Begin VB.CommandButton cmdObserva 
            Caption         =   " Observa"
            Height          =   1155
            Index           =   1
            Left            =   11070
            Picture         =   "frmEditBacteriologyNew.frx":B9EB
            Style           =   1  'Graphical
            TabIndex        =   460
            Top             =   210
            Width           =   1035
         End
         Begin VB.CommandButton cmdLock 
            Caption         =   "Unlock Result"
            Height          =   975
            Index           =   4
            Left            =   11190
            Picture         =   "frmEditBacteriologyNew.frx":C8B5
            Style           =   1  'Graphical
            TabIndex        =   242
            Top             =   3270
            Width           =   705
         End
         Begin VB.Frame fraCS 
            BorderStyle     =   0  'None
            Caption         =   "fraCS"
            Height          =   6765
            Left            =   60
            TabIndex        =   198
            Top             =   180
            Width           =   11265
            Begin VB.CommandButton cmdDelete 
               Height          =   345
               Index           =   4
               Left            =   8250
               Picture         =   "frmEditBacteriologyNew.frx":CCF7
               Style           =   1  'Graphical
               TabIndex        =   477
               ToolTipText     =   "Remove this Organism"
               Top             =   330
               Width           =   375
            End
            Begin VB.CommandButton cmdDelete 
               Height          =   345
               Index           =   3
               Left            =   5490
               Picture         =   "frmEditBacteriologyNew.frx":D6F9
               Style           =   1  'Graphical
               TabIndex        =   476
               ToolTipText     =   "Remove this Organism"
               Top             =   330
               Width           =   375
            End
            Begin VB.CommandButton cmdDelete 
               Height          =   345
               Index           =   2
               Left            =   2760
               Picture         =   "frmEditBacteriologyNew.frx":E0FB
               Style           =   1  'Graphical
               TabIndex        =   475
               ToolTipText     =   "Remove this Organism"
               Top             =   300
               Width           =   375
            End
            Begin VB.CommandButton cmdDelete 
               Height          =   345
               Index           =   1
               Left            =   30
               Picture         =   "frmEditBacteriologyNew.frx":EAFD
               Style           =   1  'Graphical
               TabIndex        =   474
               ToolTipText     =   "Remove this Organism"
               Top             =   300
               Width           =   375
            End
            Begin VB.CommandButton cmdReportNone 
               Height          =   285
               Index           =   4
               Left            =   10830
               Picture         =   "frmEditBacteriologyNew.frx":F4FF
               Style           =   1  'Graphical
               TabIndex        =   469
               ToolTipText     =   "Make All Non-Reportable"
               Top             =   1650
               Width           =   285
            End
            Begin VB.CommandButton cmdReportAll 
               Height          =   285
               Index           =   4
               Left            =   10830
               Picture         =   "frmEditBacteriologyNew.frx":F7D5
               Style           =   1  'Graphical
               TabIndex        =   468
               ToolTipText     =   "Make All Reportable"
               Top             =   1350
               Width           =   285
            End
            Begin VB.CommandButton cmdReportNone 
               Height          =   285
               Index           =   3
               Left            =   8070
               Picture         =   "frmEditBacteriologyNew.frx":FAAB
               Style           =   1  'Graphical
               TabIndex        =   467
               ToolTipText     =   "Make All Non-Reportable"
               Top             =   1650
               Width           =   285
            End
            Begin VB.CommandButton cmdReportAll 
               Height          =   285
               Index           =   3
               Left            =   8070
               Picture         =   "frmEditBacteriologyNew.frx":FD81
               Style           =   1  'Graphical
               TabIndex        =   466
               ToolTipText     =   "Make All Reportable"
               Top             =   1350
               Width           =   285
            End
            Begin VB.CommandButton cmdReportNone 
               Height          =   285
               Index           =   2
               Left            =   5340
               Picture         =   "frmEditBacteriologyNew.frx":10057
               Style           =   1  'Graphical
               TabIndex        =   465
               ToolTipText     =   "Make All Non-Reportable"
               Top             =   1650
               Width           =   285
            End
            Begin VB.CommandButton cmdReportAll 
               Height          =   285
               Index           =   2
               Left            =   5340
               Picture         =   "frmEditBacteriologyNew.frx":1032D
               Style           =   1  'Graphical
               TabIndex        =   464
               ToolTipText     =   "Make All Reportable"
               Top             =   1350
               Width           =   285
            End
            Begin VB.CommandButton cmdReportAll 
               Height          =   285
               Index           =   1
               Left            =   2610
               Picture         =   "frmEditBacteriologyNew.frx":10603
               Style           =   1  'Graphical
               TabIndex        =   463
               ToolTipText     =   "Make All Reportable"
               Top             =   1350
               Width           =   285
            End
            Begin VB.CommandButton cmdReportNone 
               Height          =   285
               Index           =   1
               Left            =   2610
               Picture         =   "frmEditBacteriologyNew.frx":108D9
               Style           =   1  'Graphical
               TabIndex        =   462
               ToolTipText     =   "Make All Non-Reportable"
               Top             =   1650
               Width           =   285
            End
            Begin VB.CheckBox chkNonReportable 
               DownPicture     =   "frmEditBacteriologyNew.frx":10BAF
               Height          =   345
               Index           =   3
               Left            =   8250
               Picture         =   "frmEditBacteriologyNew.frx":11139
               Style           =   1  'Graphical
               TabIndex        =   445
               ToolTipText     =   "Check to make culture non-reportable"
               Top             =   660
               Width           =   375
            End
            Begin VB.CheckBox chkNonReportable 
               DownPicture     =   "frmEditBacteriologyNew.frx":116C3
               Height          =   345
               Index           =   2
               Left            =   5490
               Picture         =   "frmEditBacteriologyNew.frx":11C4D
               Style           =   1  'Graphical
               TabIndex        =   444
               ToolTipText     =   "Check to make culture non-reportable"
               Top             =   660
               Width           =   375
            End
            Begin VB.CheckBox chkNonReportable 
               DownPicture     =   "frmEditBacteriologyNew.frx":121D7
               Height          =   345
               Index           =   1
               Left            =   2760
               Picture         =   "frmEditBacteriologyNew.frx":12761
               Style           =   1  'Graphical
               TabIndex        =   443
               ToolTipText     =   "Check to make culture non-reportable"
               Top             =   660
               Width           =   375
            End
            Begin VB.CheckBox chkNonReportable 
               DownPicture     =   "frmEditBacteriologyNew.frx":12CEB
               Height          =   345
               Index           =   0
               Left            =   30
               Picture         =   "frmEditBacteriologyNew.frx":13275
               Style           =   1  'Graphical
               TabIndex        =   442
               ToolTipText     =   "Check to make culture non-reportable"
               Top             =   660
               Width           =   375
            End
            Begin VB.ComboBox cmbConC 
               Height          =   315
               Left            =   5910
               TabIndex        =   199
               Text            =   "cmbConC"
               Top             =   5970
               Visible         =   0   'False
               Width           =   4395
            End
            Begin VB.ComboBox cmbMSC 
               Height          =   315
               Left            =   450
               TabIndex        =   202
               Text            =   "cmbMSC"
               Top             =   5970
               Visible         =   0   'False
               Width           =   4455
            End
            Begin VB.ComboBox cmbQualifier 
               Height          =   315
               Index           =   4
               Left            =   8640
               TabIndex        =   228
               Text            =   "cmbQualifier"
               Top             =   990
               Width           =   2205
            End
            Begin VB.ComboBox cmbQualifier 
               Height          =   315
               Index           =   3
               Left            =   5880
               TabIndex        =   227
               Text            =   "cmbQualifier"
               Top             =   990
               Width           =   2205
            End
            Begin VB.ComboBox cmbQualifier 
               Height          =   315
               Index           =   2
               Left            =   3150
               TabIndex        =   226
               Text            =   "cmbQualifier"
               Top             =   990
               Width           =   2205
            End
            Begin VB.ComboBox cmbQualifier 
               Height          =   315
               Index           =   1
               Left            =   420
               TabIndex        =   225
               Text            =   "cmbQualifier"
               Top             =   990
               Width           =   2205
            End
            Begin VB.CommandButton cmdUseSecondary 
               Height          =   525
               Index           =   4
               Left            =   8220
               Picture         =   "frmEditBacteriologyNew.frx":137FF
               Style           =   1  'Graphical
               TabIndex        =   224
               ToolTipText     =   "Use Secondary Lists"
               Top             =   3630
               Width           =   375
            End
            Begin VB.CommandButton cmdRemoveSecondary 
               Height          =   525
               Index           =   4
               Left            =   8220
               Picture         =   "frmEditBacteriologyNew.frx":13B09
               Style           =   1  'Graphical
               TabIndex        =   223
               ToolTipText     =   "Remove Secondary Lists"
               Top             =   3090
               Width           =   375
            End
            Begin VB.CommandButton cmdUseSecondary 
               Height          =   525
               Index           =   3
               Left            =   5460
               Picture         =   "frmEditBacteriologyNew.frx":13E13
               Style           =   1  'Graphical
               TabIndex        =   222
               ToolTipText     =   "Use Secondary Lists"
               Top             =   3630
               Width           =   375
            End
            Begin VB.CommandButton cmdRemoveSecondary 
               Height          =   525
               Index           =   3
               Left            =   5460
               Picture         =   "frmEditBacteriologyNew.frx":1411D
               Style           =   1  'Graphical
               TabIndex        =   221
               ToolTipText     =   "Remove Secondary Lists"
               Top             =   3090
               Width           =   375
            End
            Begin VB.CommandButton cmdUseSecondary 
               Height          =   525
               Index           =   2
               Left            =   2730
               Picture         =   "frmEditBacteriologyNew.frx":14427
               Style           =   1  'Graphical
               TabIndex        =   220
               ToolTipText     =   "Use Secondary Lists"
               Top             =   3630
               Width           =   375
            End
            Begin VB.CommandButton cmdRemoveSecondary 
               Height          =   525
               Index           =   2
               Left            =   2730
               Picture         =   "frmEditBacteriologyNew.frx":14731
               Style           =   1  'Graphical
               TabIndex        =   219
               ToolTipText     =   "Remove Secondary Lists"
               Top             =   3090
               Width           =   375
            End
            Begin VB.CommandButton cmdUseSecondary 
               Height          =   525
               Index           =   1
               Left            =   0
               Picture         =   "frmEditBacteriologyNew.frx":14A3B
               Style           =   1  'Graphical
               TabIndex        =   218
               ToolTipText     =   "Use Secondary Lists"
               Top             =   3630
               Width           =   375
            End
            Begin VB.CommandButton cmdRemoveSecondary 
               Height          =   525
               Index           =   1
               Left            =   0
               Picture         =   "frmEditBacteriologyNew.frx":14D45
               Style           =   1  'Graphical
               TabIndex        =   217
               ToolTipText     =   "Remove Secondary Lists"
               Top             =   3090
               Width           =   375
            End
            Begin VB.ComboBox cmbABSelect 
               BackColor       =   &H00FFFFC0&
               Height          =   315
               Index           =   4
               IntegralHeight  =   0   'False
               Left            =   8640
               TabIndex        =   216
               Text            =   "cmbABSelect"
               Top             =   5280
               Width           =   2205
            End
            Begin VB.ComboBox cmbABSelect 
               BackColor       =   &H00FFFFC0&
               Height          =   315
               Index           =   3
               IntegralHeight  =   0   'False
               Left            =   5880
               TabIndex        =   215
               Text            =   "cmbABSelect"
               Top             =   5280
               Width           =   2205
            End
            Begin VB.ComboBox cmbABSelect 
               BackColor       =   &H00FFFFC0&
               Height          =   315
               Index           =   2
               IntegralHeight  =   0   'False
               Left            =   3120
               TabIndex        =   214
               Text            =   "cmbABSelect"
               Top             =   5310
               Width           =   2235
            End
            Begin VB.ComboBox cmbABSelect 
               BackColor       =   &H00FFFFC0&
               Height          =   315
               Index           =   1
               IntegralHeight  =   0   'False
               Left            =   420
               TabIndex        =   213
               Text            =   "cmbABSelect"
               Top             =   5280
               Width           =   2205
            End
            Begin VB.ComboBox cmbOrgName 
               BackColor       =   &H00C0E0FF&
               Height          =   315
               Index           =   4
               Left            =   8640
               TabIndex        =   212
               Text            =   "cmbOrgName"
               Top             =   660
               Width           =   2205
            End
            Begin VB.ComboBox cmbOrgName 
               BackColor       =   &H00C0E0FF&
               Height          =   315
               Index           =   3
               Left            =   5880
               TabIndex        =   211
               Text            =   "cmbOrgName"
               Top             =   660
               Width           =   2205
            End
            Begin VB.ComboBox cmbOrgName 
               BackColor       =   &H00C0E0FF&
               Height          =   315
               Index           =   2
               Left            =   3150
               TabIndex        =   210
               Text            =   "cmbOrgName"
               Top             =   660
               Width           =   2205
            End
            Begin VB.ComboBox cmbOrgName 
               BackColor       =   &H00C0E0FF&
               Height          =   315
               Index           =   1
               Left            =   420
               TabIndex        =   209
               Text            =   "cmbOrgName"
               Top             =   660
               Width           =   2205
            End
            Begin VB.ComboBox cmbOrgGroup 
               BackColor       =   &H0000FFFF&
               Height          =   315
               Index           =   4
               Left            =   8910
               TabIndex        =   208
               Text            =   "cmbOrgGroup"
               Top             =   330
               Width           =   1935
            End
            Begin VB.ComboBox cmbOrgGroup 
               BackColor       =   &H0000FFFF&
               Height          =   315
               Index           =   1
               Left            =   690
               TabIndex        =   207
               Text            =   "cmbOrgGroup"
               Top             =   330
               Width           =   1935
            End
            Begin VB.ComboBox cmbOrgGroup 
               BackColor       =   &H0000FFFF&
               Height          =   315
               Index           =   2
               Left            =   3420
               TabIndex        =   206
               Text            =   "cmbOrgGroup"
               Top             =   330
               Width           =   1935
            End
            Begin VB.ComboBox cmbOrgGroup 
               BackColor       =   &H0000FFFF&
               Height          =   315
               Index           =   3
               Left            =   6150
               TabIndex        =   205
               Text            =   "cmbOrgGroup"
               Top             =   330
               Width           =   1935
            End
            Begin VB.TextBox txtMSC 
               Height          =   1005
               Left            =   420
               MaxLength       =   2000
               MultiLine       =   -1  'True
               TabIndex        =   204
               Text            =   "frmEditBacteriologyNew.frx":1504F
               Top             =   5610
               Width           =   4485
            End
            Begin VB.CommandButton cmdMSC 
               Caption         =   "..."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   4920
               TabIndex        =   203
               ToolTipText     =   "Choose a comment from a list"
               Top             =   5940
               Width           =   435
            End
            Begin VB.TextBox txtConC 
               Height          =   1005
               Left            =   5880
               MaxLength       =   2000
               MultiLine       =   -1  'True
               TabIndex        =   201
               Text            =   "frmEditBacteriologyNew.frx":1506A
               Top             =   5610
               Width           =   4425
            End
            Begin VB.CommandButton cmdConC 
               Caption         =   "..."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   10320
               TabIndex        =   200
               ToolTipText     =   "Choose a comment from a list"
               Top             =   5940
               Width           =   435
            End
            Begin VB.Label lblCf 
               Alignment       =   2  'Center
               BackColor       =   &H000000FF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Cystic fibrosis patient"
               ForeColor       =   &H0000FFFF&
               Height          =   285
               Left            =   8640
               TabIndex        =   380
               Top             =   30
               Visible         =   0   'False
               Width           =   2190
            End
            Begin VB.Label lblSetAllR 
               AutoSize        =   -1  'True
               BackColor       =   &H008080FF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "R"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Index           =   4
               Left            =   8340
               TabIndex        =   241
               ToolTipText     =   "Set All Resistant"
               Top             =   4170
               Width           =   270
            End
            Begin VB.Label lblSetAllS 
               AutoSize        =   -1  'True
               BackColor       =   &H0080FF80&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "S"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Index           =   4
               Left            =   8340
               TabIndex        =   240
               ToolTipText     =   "Set All Sensitive"
               Top             =   4560
               Width           =   255
            End
            Begin VB.Label lblSetAllR 
               AutoSize        =   -1  'True
               BackColor       =   &H008080FF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "R"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Index           =   3
               Left            =   5580
               TabIndex        =   239
               ToolTipText     =   "Set All Resistant"
               Top             =   4170
               Width           =   270
            End
            Begin VB.Label lblSetAllS 
               AutoSize        =   -1  'True
               BackColor       =   &H0080FF80&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "S"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Index           =   3
               Left            =   5580
               TabIndex        =   238
               ToolTipText     =   "Set All Sensitive"
               Top             =   4560
               Width           =   255
            End
            Begin VB.Label lblSetAllR 
               AutoSize        =   -1  'True
               BackColor       =   &H008080FF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "R"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Index           =   2
               Left            =   2850
               TabIndex        =   237
               ToolTipText     =   "Set All Resistant"
               Top             =   4170
               Width           =   270
            End
            Begin VB.Label lblSetAllS 
               AutoSize        =   -1  'True
               BackColor       =   &H0080FF80&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "S"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Index           =   2
               Left            =   2850
               TabIndex        =   236
               ToolTipText     =   "Set All Sensitive"
               Top             =   4560
               Width           =   255
            End
            Begin VB.Label Label46 
               Alignment       =   2  'Center
               BackColor       =   &H00C0C0FF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "4"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   3
               Left            =   8640
               TabIndex        =   235
               Top             =   330
               Width           =   270
            End
            Begin VB.Label Label46 
               Alignment       =   2  'Center
               BackColor       =   &H00C0C0FF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "3"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   2
               Left            =   5880
               TabIndex        =   234
               Top             =   330
               Width           =   270
            End
            Begin VB.Label Label46 
               Alignment       =   2  'Center
               BackColor       =   &H00C0C0FF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "2"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   1
               Left            =   3150
               TabIndex        =   233
               Top             =   330
               Width           =   270
            End
            Begin VB.Label Label46 
               Alignment       =   2  'Center
               BackColor       =   &H00C0C0FF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "1"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   0
               Left            =   420
               TabIndex        =   232
               Top             =   330
               Width           =   270
            End
            Begin VB.Label lblSetAllS 
               AutoSize        =   -1  'True
               BackColor       =   &H0080FF80&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "S"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Index           =   1
               Left            =   120
               TabIndex        =   231
               ToolTipText     =   "Set All Sensitive"
               Top             =   4560
               Width           =   255
            End
            Begin VB.Label lblSetAllR 
               AutoSize        =   -1  'True
               BackColor       =   &H008080FF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "R"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Index           =   1
               Left            =   120
               TabIndex        =   230
               ToolTipText     =   "Set All Resistant"
               Top             =   4170
               Width           =   270
            End
            Begin VB.Label lblCells 
               Alignment       =   2  'Center
               BackColor       =   &H0080C0FF&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   390
               TabIndex        =   229
               Top             =   30
               Width           =   7665
            End
         End
      End
      Begin VB.Frame fraValid 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   6045
         Index           =   1
         Left            =   -74160
         TabIndex        =   169
         Top             =   420
         Width           =   10125
         Begin VB.TextBox txtUrineComment 
            Height          =   1185
            Left            =   30
            MultiLine       =   -1  'True
            TabIndex        =   195
            Top             =   4620
            Width           =   7065
         End
         Begin VB.Frame fraMicroscopy 
            Caption         =   "Microscopy"
            Height          =   3405
            Left            =   60
            TabIndex        =   178
            Top             =   270
            Width           =   2955
            Begin VB.CommandButton cmdDeleteMicroscopy 
               Height          =   345
               Left            =   2550
               Picture         =   "frmEditBacteriologyNew.frx":15080
               Style           =   1  'Graphical
               TabIndex        =   478
               ToolTipText     =   "Remove all microscopy results"
               Top             =   120
               Width           =   375
            End
            Begin VB.TextBox txtRCC 
               Height          =   285
               Left            =   1560
               MaxLength       =   10
               TabIndex        =   187
               Top             =   1290
               Width           =   1200
            End
            Begin VB.TextBox txtWCC 
               Height          =   285
               Left            =   1560
               MaxLength       =   10
               TabIndex        =   186
               Top             =   960
               Width           =   780
            End
            Begin VB.ComboBox cmbCasts 
               Height          =   315
               Left            =   750
               TabIndex        =   185
               Text            =   "cmbCasts"
               Top             =   1920
               Width           =   2025
            End
            Begin VB.ComboBox cmbCrystals 
               Height          =   315
               Left            =   750
               TabIndex        =   184
               Text            =   "cmbCrystals"
               Top             =   1590
               Width           =   2025
            End
            Begin VB.ComboBox cmbMisc 
               Height          =   315
               Index           =   0
               Left            =   750
               TabIndex        =   183
               Text            =   "cmbMisc"
               Top             =   2250
               Width           =   2025
            End
            Begin VB.ComboBox cmbMisc 
               Height          =   315
               Index           =   1
               Left            =   750
               TabIndex        =   182
               Text            =   "cmbMisc"
               Top             =   2580
               Width           =   2025
            End
            Begin VB.ComboBox cmbMisc 
               Height          =   315
               Index           =   2
               Left            =   750
               TabIndex        =   181
               Text            =   "cmbMisc"
               Top             =   2910
               Width           =   2025
            End
            Begin VB.TextBox txtBacteria 
               Height          =   285
               Left            =   1560
               TabIndex        =   180
               Top             =   630
               Visible         =   0   'False
               Width           =   1185
            End
            Begin VB.CommandButton cmdNADMicro 
               Caption         =   "NAD"
               Height          =   345
               Left            =   180
               TabIndex        =   179
               Top             =   600
               Width           =   615
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Crystals"
               Height          =   195
               Index           =   17
               Left            =   180
               TabIndex        =   194
               Top             =   1650
               Width           =   540
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Casts"
               Height          =   195
               Index           =   18
               Left            =   330
               TabIndex        =   193
               Top             =   1980
               Width           =   390
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Misc"
               Height          =   195
               Index           =   19
               Left            =   390
               TabIndex        =   192
               Top             =   2310
               Width           =   330
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "RCC"
               Height          =   195
               Index           =   16
               Left            =   1170
               TabIndex        =   191
               Top             =   1320
               Width           =   330
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "WCC"
               Height          =   195
               Index           =   15
               Left            =   1110
               TabIndex        =   190
               Top             =   1020
               Width           =   375
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Bacteria"
               Height          =   195
               Index           =   14
               Left            =   900
               TabIndex        =   189
               Top             =   660
               Visible         =   0   'False
               Width           =   585
            End
            Begin VB.Label Label36 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               Caption         =   "/cmm"
               Height          =   195
               Index           =   0
               Left            =   2340
               TabIndex        =   188
               Top             =   1020
               Width           =   435
            End
         End
         Begin VB.Frame fraPregnancy 
            Caption         =   "Pregnancy"
            Height          =   1245
            Left            =   3240
            TabIndex        =   172
            Top             =   270
            Width           =   3855
            Begin VB.TextBox txtPregnancy 
               ForeColor       =   &H8000000D&
               Height          =   285
               Left            =   1560
               MaxLength       =   20
               MultiLine       =   -1  'True
               TabIndex        =   174
               ToolTipText     =   "P-Positive N-Negative E-Equivocal I-Inconclusive U-Unsuitable"
               Top             =   420
               Width           =   2055
            End
            Begin VB.TextBox txtHCGLevel 
               Height          =   285
               Left            =   1560
               MaxLength       =   5
               TabIndex        =   173
               Top             =   750
               Width           =   1545
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "IU/L"
               Height          =   195
               Index           =   22
               Left            =   3120
               TabIndex        =   177
               Top             =   780
               Width           =   330
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               Caption         =   "Pregnancy Test"
               Height          =   195
               Index           =   20
               Left            =   360
               TabIndex        =   176
               Top             =   450
               Width           =   1155
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               Caption         =   "HCG Level"
               Height          =   195
               Index           =   21
               Left            =   720
               TabIndex        =   175
               Top             =   780
               Width           =   795
            End
         End
         Begin VB.CommandButton cmdLock 
            Caption         =   "Unlock Result"
            Height          =   735
            Index           =   1
            Left            =   7500
            Picture         =   "frmEditBacteriologyNew.frx":15A82
            Style           =   1  'Graphical
            TabIndex        =   171
            Top             =   5070
            Width           =   1425
         End
         Begin VB.ComboBox cmbUrineComment 
            Height          =   315
            Left            =   1920
            TabIndex        =   170
            Text            =   "cmbUrineComment"
            Top             =   4320
            Width           =   5175
         End
         Begin VB.Label lblPrinted 
            Alignment       =   2  'Center
            BackColor       =   &H0000FFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "P"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   13.5
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400040&
            Height          =   435
            Index           =   1
            Left            =   9630
            TabIndex        =   481
            ToolTipText     =   "Sample has been Printed"
            Top             =   5370
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.Label lblValid 
            Alignment       =   2  'Center
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "V"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   13.5
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   1
            Left            =   9120
            TabIndex        =   480
            Top             =   5370
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Urine Specimen Comment"
            Height          =   195
            Index           =   33
            Left            =   30
            TabIndex        =   196
            Top             =   4350
            Width           =   1830
         End
      End
      Begin VB.CommandButton cmdLock 
         Caption         =   "Unlock Result"
         Height          =   795
         Index           =   8
         Left            =   -67770
         Picture         =   "frmEditBacteriologyNew.frx":15EC4
         Style           =   1  'Graphical
         TabIndex        =   155
         Top             =   2550
         Width           =   1425
      End
      Begin VB.CommandButton cmdLock 
         Caption         =   "Un&lock Result"
         Height          =   795
         Index           =   7
         Left            =   -68280
         Picture         =   "frmEditBacteriologyNew.frx":16306
         Style           =   1  'Graphical
         TabIndex        =   153
         Top             =   3300
         Width           =   1425
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   6030
         Left            =   -74970
         TabIndex        =   104
         Top             =   870
         Width           =   7065
         _Version        =   65536
         _ExtentX        =   12462
         _ExtentY        =   10636
         _StockProps     =   15
         BackColor       =   13160660
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   3
         BorderWidth     =   1
         BevelInner      =   2
         Outline         =   -1  'True
         FloodShowPct    =   0   'False
         Begin VB.Frame Frame2 
            Caption         =   "Day 2"
            Height          =   2250
            Index           =   1
            Left            =   150
            TabIndex        =   129
            Top             =   2100
            Width           =   6765
            Begin VB.ComboBox cmbDay2 
               Height          =   315
               Index           =   53
               Left            =   4830
               TabIndex        =   493
               Text            =   "cmbDay2"
               Top             =   1800
               Width           =   1755
            End
            Begin VB.ComboBox cmbDay2 
               Height          =   315
               Index           =   52
               Left            =   2970
               TabIndex        =   492
               Text            =   "cmbDay2"
               Top             =   1800
               Width           =   1755
            End
            Begin VB.ComboBox cmbDay2 
               Height          =   315
               Index           =   51
               Left            =   1110
               TabIndex        =   491
               Text            =   "cmbDay2"
               Top             =   1800
               Width           =   1755
            End
            Begin VB.ComboBox cmbDay2 
               Height          =   315
               Index           =   41
               Left            =   1110
               TabIndex        =   284
               Text            =   "cmbDay2"
               Top             =   675
               Width           =   1755
            End
            Begin VB.ComboBox cmbDay2 
               Height          =   315
               Index           =   42
               Left            =   2970
               TabIndex        =   283
               Text            =   "cmbDay2"
               Top             =   675
               Width           =   1755
            End
            Begin VB.ComboBox cmbDay2 
               Height          =   315
               Index           =   43
               Left            =   4830
               TabIndex        =   282
               Text            =   "cmbDay2"
               Top             =   675
               Width           =   1755
            End
            Begin VB.ComboBox cmbDay2 
               Height          =   315
               Index           =   33
               Left            =   4830
               TabIndex        =   138
               Text            =   "cmbDay2"
               Top             =   1425
               Width           =   1755
            End
            Begin VB.ComboBox cmbDay2 
               Height          =   315
               Index           =   32
               Left            =   2970
               TabIndex        =   137
               Text            =   "cmbDay2"
               Top             =   1425
               Width           =   1755
            End
            Begin VB.ComboBox cmbDay2 
               Height          =   315
               Index           =   31
               Left            =   1110
               TabIndex        =   136
               Text            =   "cmbDay2"
               Top             =   1425
               Width           =   1755
            End
            Begin VB.ComboBox cmbDay2 
               Height          =   315
               Index           =   23
               Left            =   4830
               TabIndex        =   135
               Text            =   "cmbDay2"
               Top             =   1050
               Width           =   1755
            End
            Begin VB.ComboBox cmbDay2 
               Height          =   315
               Index           =   22
               Left            =   2970
               TabIndex        =   134
               Text            =   "cmbDay2"
               Top             =   1050
               Width           =   1755
            End
            Begin VB.ComboBox cmbDay2 
               Height          =   315
               Index           =   21
               Left            =   1110
               TabIndex        =   133
               Text            =   "cmbDay2"
               Top             =   1050
               Width           =   1755
            End
            Begin VB.ComboBox cmbDay2 
               Height          =   315
               Index           =   13
               Left            =   4830
               TabIndex        =   132
               Text            =   "cmbDay2"
               Top             =   300
               Width           =   1755
            End
            Begin VB.ComboBox cmbDay2 
               Height          =   315
               Index           =   12
               Left            =   2970
               TabIndex        =   131
               Text            =   "cmbDay2"
               Top             =   300
               Width           =   1755
            End
            Begin VB.ComboBox cmbDay2 
               Height          =   315
               Index           =   11
               Left            =   1110
               TabIndex        =   130
               Text            =   "cmbDay2"
               Top             =   300
               Width           =   1755
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "STEC"
               Height          =   195
               Index           =   26
               Left            =   360
               TabIndex        =   494
               Top             =   1860
               Width           =   420
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "DCA"
               Height          =   195
               Index           =   100
               Left            =   450
               TabIndex        =   285
               Top             =   735
               Width           =   330
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Organism 3"
               Height          =   195
               Index           =   106
               Left            =   5280
               TabIndex        =   144
               Top             =   0
               Width           =   795
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Organism 2"
               Height          =   195
               Index           =   105
               Left            =   3390
               TabIndex        =   143
               Top             =   0
               Width           =   795
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Organism 1"
               Height          =   195
               Index           =   104
               Left            =   1620
               TabIndex        =   142
               Top             =   0
               Width           =   795
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "CAMP"
               Height          =   195
               Index           =   102
               Left            =   330
               TabIndex        =   141
               Top             =   1485
               Width           =   450
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "CROMO "
               Height          =   195
               Index           =   101
               Left            =   135
               TabIndex        =   140
               Top             =   1110
               Width           =   645
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "XLD"
               Height          =   195
               Index           =   99
               Left            =   465
               TabIndex        =   139
               Top             =   360
               Width           =   315
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Day 3"
            Height          =   1350
            Index           =   2
            Left            =   150
            TabIndex        =   121
            Top             =   4485
            Width           =   6765
            Begin VB.ComboBox cmbDay3 
               Height          =   315
               Index           =   4
               Left            =   1110
               TabIndex        =   288
               Text            =   "cmbDay3"
               Top             =   765
               Width           =   1755
            End
            Begin VB.ComboBox cmbDay3 
               Height          =   315
               Index           =   5
               Left            =   2970
               TabIndex        =   287
               Text            =   "cmbDay3"
               Top             =   765
               Width           =   1755
            End
            Begin VB.ComboBox cmbDay3 
               Height          =   315
               Index           =   6
               Left            =   4830
               TabIndex        =   286
               Text            =   "cmbDay3"
               Top             =   765
               Width           =   1755
            End
            Begin VB.ComboBox cmbDay3 
               Height          =   315
               Index           =   3
               Left            =   4830
               TabIndex        =   124
               Text            =   "cmbDay3"
               Top             =   390
               Width           =   1755
            End
            Begin VB.ComboBox cmbDay3 
               Height          =   315
               Index           =   2
               Left            =   2970
               TabIndex        =   123
               Text            =   "cmbDay3"
               Top             =   390
               Width           =   1755
            End
            Begin VB.ComboBox cmbDay3 
               Height          =   315
               Index           =   1
               Left            =   1110
               TabIndex        =   122
               Text            =   "cmbDay3"
               Top             =   390
               Width           =   1755
            End
            Begin VB.Label Label34 
               AutoSize        =   -1  'True
               Caption         =   "CROMO "
               Height          =   195
               Index           =   3
               Left            =   120
               TabIndex        =   289
               Top             =   825
               Width           =   645
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Organism 3"
               Height          =   195
               Index           =   109
               Left            =   5280
               TabIndex        =   128
               Top             =   0
               Width           =   795
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Organism 2"
               Height          =   195
               Index           =   108
               Left            =   3390
               TabIndex        =   127
               Top             =   0
               Width           =   795
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Organism 1"
               Height          =   195
               Index           =   107
               Left            =   1620
               TabIndex        =   126
               Top             =   0
               Width           =   795
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "CAMP"
               Height          =   195
               Index           =   103
               Left            =   315
               TabIndex        =   125
               Top             =   450
               Width           =   450
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Day 1"
            Height          =   1845
            Index           =   0
            Left            =   150
            TabIndex        =   105
            Top             =   150
            Width           =   6765
            Begin VB.ComboBox cmbDay1 
               Height          =   315
               Index           =   11
               Left            =   1110
               TabIndex        =   106
               Text            =   "cmbDay1"
               Top             =   300
               Width           =   1755
            End
            Begin VB.ComboBox cmbDay1 
               Height          =   315
               Index           =   43
               Left            =   4830
               TabIndex        =   497
               Text            =   "cmbDay1"
               Top             =   1425
               Width           =   1755
            End
            Begin VB.ComboBox cmbDay1 
               Height          =   315
               Index           =   42
               Left            =   2970
               TabIndex        =   496
               Text            =   "cmbDay1"
               Top             =   1425
               Width           =   1755
            End
            Begin VB.ComboBox cmbDay1 
               Height          =   315
               Index           =   41
               Left            =   1110
               TabIndex        =   495
               Text            =   "cmbDay1"
               Top             =   1425
               Width           =   1755
            End
            Begin VB.ComboBox cmbDay1 
               Height          =   315
               Index           =   33
               Left            =   4830
               TabIndex        =   114
               Text            =   "cmbDay1"
               Top             =   1050
               Width           =   1755
            End
            Begin VB.ComboBox cmbDay1 
               Height          =   315
               Index           =   32
               Left            =   2970
               TabIndex        =   113
               Text            =   "cmbDay1"
               Top             =   1050
               Width           =   1755
            End
            Begin VB.ComboBox cmbDay1 
               Height          =   315
               Index           =   31
               Left            =   1110
               TabIndex        =   112
               Text            =   "cmbDay1"
               Top             =   1050
               Width           =   1755
            End
            Begin VB.ComboBox cmbDay1 
               Height          =   315
               Index           =   23
               Left            =   4830
               TabIndex        =   111
               Text            =   "cmbDay1"
               Top             =   675
               Width           =   1755
            End
            Begin VB.ComboBox cmbDay1 
               Height          =   315
               Index           =   22
               Left            =   2970
               TabIndex        =   110
               Text            =   "cmbDay1"
               Top             =   675
               Width           =   1755
            End
            Begin VB.ComboBox cmbDay1 
               Height          =   315
               Index           =   21
               Left            =   1110
               TabIndex        =   109
               Text            =   "cmbDay1"
               Top             =   675
               Width           =   1755
            End
            Begin VB.ComboBox cmbDay1 
               Height          =   315
               Index           =   13
               Left            =   4830
               TabIndex        =   108
               Text            =   "cmbDay1"
               Top             =   300
               Width           =   1755
            End
            Begin VB.ComboBox cmbDay1 
               Height          =   315
               Index           =   12
               Left            =   2970
               TabIndex        =   107
               Text            =   "cmbDay1"
               Top             =   300
               Width           =   1755
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "STEC"
               Height          =   195
               Index           =   27
               Left            =   360
               TabIndex        =   498
               Top             =   1485
               Width           =   420
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Organism 3"
               Height          =   195
               Index           =   95
               Left            =   5280
               TabIndex        =   120
               Top             =   0
               Width           =   795
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Organism 2"
               Height          =   195
               Index           =   94
               Left            =   3390
               TabIndex        =   119
               Top             =   0
               Width           =   795
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Organism 1"
               Height          =   195
               Index           =   93
               Left            =   1620
               TabIndex        =   118
               Top             =   0
               Width           =   795
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "SMAC"
               Height          =   195
               Index           =   98
               Left            =   330
               TabIndex        =   117
               Top             =   1110
               Width           =   450
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "DCA"
               Height          =   195
               Index           =   97
               Left            =   450
               TabIndex        =   116
               Top             =   735
               Width           =   330
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "XLD"
               Height          =   195
               Index           =   96
               Left            =   465
               TabIndex        =   115
               Top             =   360
               Width           =   315
            End
         End
      End
      Begin MSFlexGridLib.MSFlexGrid grdDay 
         Height          =   1485
         Index           =   1
         Left            =   -67890
         TabIndex        =   101
         Top             =   1110
         Width           =   5610
         _ExtentX        =   9895
         _ExtentY        =   2619
         _Version        =   393216
         Cols            =   5
         BackColor       =   -2147483624
         ForeColor       =   -2147483635
         BackColorFixed  =   -2147483647
         ForeColorFixed  =   -2147483624
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   0
         AllowUserResizing=   1
         FormatString    =   "<Date/Time   |<XLD               |<DCA            |<SMAC           |<Technician      "
      End
      Begin ComCtl2.UpDown udHistoricalFaecesView 
         Height          =   285
         Left            =   -64860
         TabIndex        =   100
         Top             =   780
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   503
         _Version        =   327681
         Value           =   1
         BuddyControl    =   "lblViewOrganism"
         BuddyDispid     =   196813
         OrigLeft        =   10290
         OrigTop         =   780
         OrigRight       =   10695
         OrigBottom      =   1020
         Max             =   3
         Min             =   1
         Orientation     =   1
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.CommandButton cmdCopyFromPrevious 
         BackColor       =   &H00FF80FF&
         Caption         =   "Copy all Details from Sample # 123456789"
         Height          =   285
         Left            =   -74670
         Style           =   1  'Graphical
         TabIndex        =   95
         Top             =   420
         Visible         =   0   'False
         Width           =   5265
      End
      Begin VB.Frame FrameExtras 
         Caption         =   "Organism 1"
         ForeColor       =   &H00C000C0&
         Height          =   6495
         Index           =   1
         Left            =   -73920
         TabIndex        =   75
         Top             =   480
         Width           =   2505
         Begin VB.TextBox txtZN 
            Height          =   285
            Index           =   1
            Left            =   900
            TabIndex        =   386
            Top             =   600
            Width           =   1515
         End
         Begin VB.TextBox txtIndole 
            Height          =   285
            Index           =   1
            Left            =   900
            TabIndex        =   384
            Top             =   1230
            Width           =   1515
         End
         Begin VB.TextBox txtNotes 
            Height          =   3645
            Index           =   1
            Left            =   90
            MultiLine       =   -1  'True
            TabIndex        =   76
            Top             =   2730
            Width           =   2355
         End
         Begin VB.ComboBox cmbWetPrep 
            Height          =   315
            Index           =   1
            Left            =   900
            TabIndex        =   82
            Top             =   900
            Width           =   1515
         End
         Begin VB.ComboBox cmbGram 
            Height          =   315
            Index           =   1
            Left            =   900
            Sorted          =   -1  'True
            TabIndex        =   81
            Top             =   270
            Width           =   1515
         End
         Begin VB.TextBox txtReincubation 
            Height          =   285
            Index           =   1
            Left            =   900
            TabIndex        =   80
            Tag             =   "Rei"
            Top             =   2430
            Width           =   1515
         End
         Begin VB.TextBox txtOxidase 
            Height          =   285
            Index           =   1
            Left            =   900
            TabIndex        =   79
            Tag             =   "Oxi"
            Top             =   2130
            Width           =   1515
         End
         Begin VB.TextBox txtCatalase 
            Height          =   285
            Index           =   1
            Left            =   900
            TabIndex        =   78
            Tag             =   "Cat"
            Top             =   1830
            Width           =   1515
         End
         Begin VB.TextBox txtCoagulase 
            Height          =   285
            Index           =   1
            Left            =   900
            TabIndex        =   77
            Tag             =   "Coa"
            Top             =   1530
            Width           =   1515
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "ZN Stain"
            Height          =   195
            Index           =   0
            Left            =   225
            TabIndex        =   387
            Top             =   630
            Width           =   630
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Indole"
            Height          =   195
            Index           =   0
            Left            =   435
            TabIndex        =   385
            Top             =   1260
            Width           =   435
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Reinc"
            Height          =   195
            Index           =   38
            Left            =   450
            TabIndex        =   88
            Top             =   2430
            Width           =   420
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Wet Prep"
            Height          =   195
            Index           =   34
            Left            =   195
            TabIndex        =   87
            Top             =   930
            Width           =   675
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Oxidase"
            Height          =   195
            Index           =   37
            Left            =   300
            TabIndex        =   86
            Top             =   2160
            Width           =   570
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Catalase"
            Height          =   195
            Index           =   36
            Left            =   255
            TabIndex        =   85
            Top             =   1860
            Width           =   615
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Coagulase"
            Height          =   195
            Index           =   35
            Left            =   120
            TabIndex        =   84
            Top             =   1560
            Width           =   750
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Gram Stain"
            Height          =   195
            Index           =   3
            Left            =   90
            TabIndex        =   83
            Top             =   330
            Width           =   780
         End
      End
      Begin VB.Frame Frame14 
         Caption         =   "Clinical Details"
         Height          =   1815
         Left            =   -69420
         TabIndex        =   69
         Top             =   4440
         Width           =   5535
         Begin VB.TextBox txtClinDetails 
            Height          =   1095
            Left            =   300
            MultiLine       =   -1  'True
            TabIndex        =   71
            Top             =   600
            Width           =   5025
         End
         Begin VB.ComboBox cmbClinDetails 
            Height          =   315
            Left            =   300
            TabIndex        =   70
            Top             =   270
            Width           =   5025
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "Patients Current Antibiotics"
         Height          =   1035
         Left            =   -69300
         TabIndex        =   64
         Top             =   3240
         Width           =   5415
         Begin VB.CommandButton cmdABsInUse 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   4800
            Style           =   1  'Graphical
            TabIndex        =   67
            Top             =   420
            Width           =   375
         End
         Begin VB.ListBox lstABsInUse 
            Height          =   735
            IntegralHeight  =   0   'False
            ItemData        =   "frmEditBacteriologyNew.frx":16748
            Left            =   180
            List            =   "frmEditBacteriologyNew.frx":1674A
            TabIndex        =   66
            ToolTipText     =   "Click to remove entry"
            Top             =   240
            Width           =   4545
         End
         Begin VB.ComboBox cmbABsInUse 
            Height          =   315
            Left            =   180
            TabIndex        =   65
            Text            =   "cmbABsInUse"
            Top             =   420
            Visible         =   0   'False
            Width           =   4545
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Site"
         Height          =   765
         Left            =   -69300
         TabIndex        =   60
         Top             =   2430
         Width           =   5415
         Begin VB.TextBox txtSiteDetails 
            Height          =   315
            Left            =   2130
            TabIndex        =   62
            Top             =   270
            Width           =   3195
         End
         Begin VB.ComboBox cmbSite 
            Height          =   315
            Left            =   120
            TabIndex        =   61
            Text            =   "cmbSite"
            Top             =   270
            Width           =   1965
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Site Details"
            Height          =   195
            Index           =   87
            Left            =   2190
            TabIndex        =   63
            Top             =   30
            Width           =   795
         End
      End
      Begin VB.CommandButton cmdOrderTests 
         Caption         =   "Order Tests"
         Height          =   945
         Left            =   -63540
         Picture         =   "frmEditBacteriologyNew.frx":1674C
         Style           =   1  'Graphical
         TabIndex        =   53
         Tag             =   "bOrder"
         Top             =   2520
         Width           =   1035
      End
      Begin VB.CommandButton cmdSaveInc 
         Caption         =   "&Save"
         Height          =   735
         Left            =   -64920
         Picture         =   "frmEditBacteriologyNew.frx":16A56
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   6300
         Width           =   1035
      End
      Begin VB.Frame Frame4 
         Height          =   5655
         Left            =   -74670
         TabIndex        =   35
         Top             =   600
         Width           =   5265
         Begin VB.CheckBox chkPregnant 
            Alignment       =   1  'Right Justify
            Caption         =   "Pregnant"
            Height          =   225
            Left            =   4020
            TabIndex        =   295
            Top             =   300
            Width           =   945
         End
         Begin VB.CommandButton cmdCopyTo 
            BackColor       =   &H008080FF&
            Caption         =   "++ cc ++"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   4800
            Style           =   1  'Graphical
            TabIndex        =   149
            ToolTipText     =   "Copy To"
            Top             =   2160
            Width           =   375
         End
         Begin VB.CheckBox chkPenicillin 
            Alignment       =   1  'Right Justify
            Caption         =   "Penicillin Allergy"
            Height          =   225
            Left            =   3540
            TabIndex        =   90
            Top             =   540
            Width           =   1425
         End
         Begin VB.ComboBox cmbHospital 
            Height          =   315
            Left            =   900
            TabIndex        =   8
            Top             =   2160
            Width           =   3915
         End
         Begin VB.ComboBox cmbDemogComment 
            Height          =   315
            Left            =   900
            TabIndex        =   59
            Text            =   "cmbDemogComment"
            Top             =   3570
            Width           =   3915
         End
         Begin VB.ComboBox cmbGP 
            Height          =   315
            Left            =   900
            TabIndex        =   11
            Text            =   "cmbGP"
            Top             =   3180
            Width           =   3915
         End
         Begin VB.ComboBox cmbClinician 
            Height          =   315
            Left            =   900
            TabIndex        =   10
            Text            =   "cmbClinician"
            Top             =   2850
            Width           =   3915
         End
         Begin VB.TextBox tAddress 
            Height          =   285
            Index           =   1
            Left            =   750
            MaxLength       =   30
            TabIndex        =   7
            Top             =   1770
            Width           =   4215
         End
         Begin VB.TextBox tAddress 
            Height          =   285
            Index           =   0
            Left            =   750
            MaxLength       =   30
            TabIndex        =   6
            Top             =   1500
            Width           =   4215
         End
         Begin VB.ComboBox cmbWard 
            Height          =   315
            Left            =   900
            TabIndex        =   9
            Text            =   "cmbWard"
            Top             =   2490
            Width           =   3915
         End
         Begin VB.TextBox txtDemographicComment 
            Height          =   1515
            Left            =   900
            MultiLine       =   -1  'True
            TabIndex        =   16
            Top             =   3900
            Width           =   3885
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            Caption         =   "Hospital"
            Height          =   195
            Index           =   100
            Left            =   270
            TabIndex        =   74
            Top             =   2220
            Width           =   570
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            Caption         =   "GP"
            Height          =   195
            Index           =   103
            Left            =   630
            TabIndex        =   50
            Top             =   3270
            Width           =   225
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            Caption         =   "Clinician"
            Height          =   195
            Index           =   102
            Left            =   255
            TabIndex        =   49
            Top             =   2880
            Width           =   585
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            Caption         =   "Comments"
            Height          =   195
            Index           =   104
            Left            =   120
            TabIndex        =   48
            Top             =   3630
            Width           =   735
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            Caption         =   "Address"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   47
            Top             =   1530
            Width           =   570
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            Caption         =   "Ward"
            Height          =   195
            Index           =   101
            Left            =   450
            TabIndex        =   46
            Top             =   2550
            Width           =   390
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Sex"
            Height          =   195
            Index           =   84
            Left            =   3930
            TabIndex        =   45
            Top             =   1200
            Width           =   270
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Age"
            Height          =   195
            Index           =   83
            Left            =   2760
            TabIndex        =   44
            Top             =   1200
            Width           =   285
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            Caption         =   "D.o.B"
            Height          =   195
            Index           =   3
            Left            =   210
            TabIndex        =   43
            Top             =   1230
            Width           =   405
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            Caption         =   "Name"
            Height          =   195
            Index           =   1
            Left            =   210
            TabIndex        =   42
            Top             =   810
            Width           =   420
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            Caption         =   "Chart #"
            Height          =   195
            Index           =   0
            Left            =   90
            TabIndex        =   41
            Top             =   330
            Width           =   525
         End
         Begin VB.Label lblChart 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   750
            TabIndex        =   40
            Top             =   300
            Width           =   1485
         End
         Begin VB.Label lblName 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   750
            TabIndex        =   39
            Top             =   780
            Width           =   4215
         End
         Begin VB.Label lblDoB 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   750
            TabIndex        =   38
            Top             =   1170
            Width           =   1515
         End
         Begin VB.Label lblAge 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   3180
            TabIndex        =   37
            Top             =   1170
            Width           =   585
         End
         Begin VB.Label lblSex 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   4260
            TabIndex        =   36
            Top             =   1170
            Width           =   705
         End
      End
      Begin VB.CommandButton cmdSaveDemographics 
         Caption         =   "Save && &Hold"
         Enabled         =   0   'False
         Height          =   735
         Left            =   -67020
         Picture         =   "frmEditBacteriologyNew.frx":170C0
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   6300
         Width           =   1035
      End
      Begin VB.Frame fraDate 
         Caption         =   "Sample Date"
         Height          =   1845
         Left            =   -69300
         TabIndex        =   23
         Top             =   600
         Width           =   4395
         Begin MSComCtl2.DTPicker dtRecDate 
            Height          =   315
            Left            =   2310
            TabIndex        =   92
            Top             =   270
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            _Version        =   393216
            Format          =   188874753
            CurrentDate     =   38078
         End
         Begin MSComCtl2.DTPicker dtRunDate 
            Height          =   315
            Left            =   1410
            TabIndex        =   24
            Top             =   1140
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            _Version        =   393216
            Format          =   188874753
            CurrentDate     =   36942
         End
         Begin MSComCtl2.DTPicker dtSampleDate 
            Height          =   315
            Left            =   90
            TabIndex        =   12
            Top             =   270
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            _Version        =   393216
            Format          =   188874753
            CurrentDate     =   36942
         End
         Begin MSMask.MaskEdBox tSampleTime 
            Height          =   315
            Left            =   1470
            TabIndex        =   25
            ToolTipText     =   "Time of Sample"
            Top             =   270
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
         Begin MSMask.MaskEdBox tRecTime 
            Height          =   315
            Left            =   3690
            TabIndex        =   145
            ToolTipText     =   "Time of Sample"
            Top             =   270
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
         Begin VB.Label lblDateError 
            Alignment       =   2  'Center
            BackColor       =   &H000000FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Date Sequence Error"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   675
            Left            =   3270
            TabIndex        =   381
            Top             =   1170
            Visible         =   0   'False
            Width           =   1125
         End
         Begin VB.Image iRecDate 
            Height          =   330
            Index           =   1
            Left            =   3180
            Picture         =   "frmEditBacteriologyNew.frx":17502
            Stretch         =   -1  'True
            ToolTipText     =   "Next Day"
            Top             =   600
            Width           =   480
         End
         Begin VB.Image iRecDate 
            Height          =   330
            Index           =   0
            Left            =   2310
            Picture         =   "frmEditBacteriologyNew.frx":17944
            Stretch         =   -1  'True
            ToolTipText     =   "Previous Day"
            Top             =   600
            Width           =   480
         End
         Begin VB.Image iToday 
            Height          =   330
            Index           =   2
            Left            =   2790
            Picture         =   "frmEditBacteriologyNew.frx":17D86
            Stretch         =   -1  'True
            ToolTipText     =   "Set to Today"
            Top             =   600
            Width           =   360
         End
         Begin VB.Label lblDate 
            Alignment       =   2  'Center
            Caption         =   "Received in Lab"
            Height          =   255
            Index           =   0
            Left            =   2400
            TabIndex        =   91
            Top             =   0
            Width           =   1500
         End
         Begin VB.Image iToday 
            Height          =   330
            Index           =   1
            Left            =   570
            Picture         =   "frmEditBacteriologyNew.frx":181C8
            Stretch         =   -1  'True
            ToolTipText     =   "Set to Today"
            Top             =   600
            Width           =   360
         End
         Begin VB.Image iToday 
            Height          =   330
            Index           =   0
            Left            =   1890
            Picture         =   "frmEditBacteriologyNew.frx":1860A
            Stretch         =   -1  'True
            ToolTipText     =   "Set to Today"
            Top             =   1470
            Width           =   360
         End
         Begin VB.Image iSampleDate 
            Height          =   330
            Index           =   1
            Left            =   960
            Picture         =   "frmEditBacteriologyNew.frx":18A4C
            Stretch         =   -1  'True
            ToolTipText     =   "Next Day"
            Top             =   600
            Width           =   480
         End
         Begin VB.Image iSampleDate 
            Height          =   330
            Index           =   0
            Left            =   90
            Picture         =   "frmEditBacteriologyNew.frx":18E8E
            Stretch         =   -1  'True
            ToolTipText     =   "Previous Day"
            Top             =   600
            Width           =   480
         End
         Begin VB.Image iRunDate 
            Height          =   330
            Index           =   1
            Left            =   2280
            Picture         =   "frmEditBacteriologyNew.frx":192D0
            Stretch         =   -1  'True
            ToolTipText     =   "Next Day"
            Top             =   1470
            Width           =   480
         End
         Begin VB.Image iRunDate 
            Height          =   330
            Index           =   0
            Left            =   1380
            Picture         =   "frmEditBacteriologyNew.frx":19712
            Stretch         =   -1  'True
            ToolTipText     =   "Previous Day"
            Top             =   1470
            Width           =   480
         End
         Begin VB.Label lblDate 
            Alignment       =   1  'Right Justify
            Caption         =   "Run Date"
            Height          =   225
            Index           =   1
            Left            =   450
            TabIndex        =   26
            Top             =   1170
            Width           =   930
         End
      End
      Begin VB.Frame Frame5 
         Height          =   795
         Left            =   -64920
         TabIndex        =   27
         Top             =   600
         Width           =   1365
         Begin VB.OptionButton cRooH 
            Alignment       =   1  'Right Justify
            Caption         =   "Out of Hours"
            Height          =   195
            Index           =   1
            Left            =   90
            TabIndex        =   28
            Top             =   480
            Width           =   1215
         End
         Begin VB.OptionButton cRooH 
            Alignment       =   1  'Right Justify
            Caption         =   "Routine"
            Height          =   195
            Index           =   0
            Left            =   420
            TabIndex        =   29
            Top             =   240
            Width           =   885
         End
      End
      Begin MSFlexGridLib.MSFlexGrid grdDay 
         Height          =   1485
         Index           =   2
         Left            =   -67890
         TabIndex        =   102
         Top             =   2940
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   2619
         _Version        =   393216
         Cols            =   6
         BackColor       =   -2147483624
         ForeColor       =   -2147483635
         BackColorFixed  =   -2147483647
         ForeColorFixed  =   -2147483624
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   0
         AllowUserResizing=   1
         FormatString    =   "<Date/Time   |<XLD          |<DCA          |<CROMO     |<CAMP        |<Technician      "
      End
      Begin MSFlexGridLib.MSFlexGrid grdDay 
         Height          =   1035
         Index           =   3
         Left            =   -67890
         TabIndex        =   103
         Top             =   4680
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   1826
         _Version        =   393216
         Cols            =   4
         BackColor       =   -2147483624
         ForeColor       =   -2147483635
         BackColorFixed  =   -2147483647
         ForeColorFixed  =   -2147483624
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   0
         AllowUserResizing=   1
         FormatString    =   "<Date/Time   |<CAMP                     |<CROMO                 |<Technician      "
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "CDiff Comment"
         Height          =   195
         Left            =   -73740
         TabIndex        =   506
         Top             =   5040
         Width           =   1050
      End
      Begin VB.Label lblPrinted 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Printed"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   360
         Index           =   12
         Left            =   11250
         TabIndex        =   454
         ToolTipText     =   "Sample has been Printed"
         Top             =   2070
         Visible         =   0   'False
         Width           =   1410
      End
      Begin VB.Label lblValidatedBy 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Validated By"
         Height          =   285
         Left            =   -64110
         TabIndex        =   382
         Top             =   30
         Visible         =   0   'False
         Width           =   1875
      End
      Begin VB.Label lblPrinted 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Printed"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   360
         Index           =   9
         Left            =   -63450
         TabIndex        =   297
         ToolTipText     =   "Sample has been Printed"
         Top             =   3630
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.Label lblValid 
         Alignment       =   2  'Center
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Valid"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   9
         Left            =   -63450
         TabIndex        =   296
         Top             =   3060
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.Label lblValid 
         Alignment       =   2  'Center
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Valid"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   13
         Left            =   -64620
         TabIndex        =   168
         Top             =   2700
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.Label lblPrinted 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Printed"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   360
         Index           =   13
         Left            =   -64620
         TabIndex        =   167
         ToolTipText     =   "Sample has been Printed"
         Top             =   3060
         Visible         =   0   'False
         Width           =   1410
      End
      Begin VB.Label lblValid 
         Alignment       =   2  'Center
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "V"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   13.5
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   4
         Left            =   -62760
         TabIndex        =   166
         Top             =   3660
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Label lblValid 
         Alignment       =   2  'Center
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Valid"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   11
         Left            =   -64110
         TabIndex        =   165
         Top             =   2850
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.Label lblValid 
         Alignment       =   2  'Center
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Valid"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   10
         Left            =   -63990
         TabIndex        =   164
         Top             =   660
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.Label lblValid 
         Alignment       =   2  'Center
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Valid"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   8
         Left            =   -66180
         TabIndex        =   163
         Top             =   2550
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.Label lblValid 
         Alignment       =   2  'Center
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Valid"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   7
         Left            =   -68280
         TabIndex        =   162
         Top             =   2130
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.Label lblValid 
         Alignment       =   2  'Center
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Valid"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   6
         Left            =   -65520
         TabIndex        =   161
         Top             =   2520
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.Label lblValid 
         Alignment       =   2  'Center
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Valid"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   5
         Left            =   -65280
         TabIndex        =   160
         Top             =   2370
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.Label lblPrinted 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Printed"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   360
         Index           =   8
         Left            =   -66180
         TabIndex        =   156
         ToolTipText     =   "Sample has been Printed"
         Top             =   2910
         Visible         =   0   'False
         Width           =   1410
      End
      Begin VB.Label lblPrinted 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Printed"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   360
         Index           =   7
         Left            =   -68280
         TabIndex        =   154
         ToolTipText     =   "Sample has been Printed"
         Top             =   2520
         Visible         =   0   'False
         Width           =   1410
      End
      Begin VB.Label lblPrinted 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Printed"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   360
         Index           =   6
         Left            =   -65520
         TabIndex        =   152
         ToolTipText     =   "Sample has been Printed"
         Top             =   2910
         Visible         =   0   'False
         Width           =   1410
      End
      Begin VB.Label lblPrinted 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Printed"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   360
         Index           =   5
         Left            =   -65280
         TabIndex        =   151
         ToolTipText     =   "Sample has been Printed"
         Top             =   2730
         Visible         =   0   'False
         Width           =   1410
      End
      Begin VB.Label lblPrinted 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "P"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   13.5
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   405
         Index           =   4
         Left            =   -62760
         TabIndex        =   148
         ToolTipText     =   "Sample has been Printed"
         Top             =   4110
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Label lblPrinted 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Printed"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   360
         Index           =   10
         Left            =   -63990
         TabIndex        =   147
         ToolTipText     =   "Sample has been Printed"
         Top             =   1020
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.Label lblPrinted 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Printed"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   360
         Index           =   11
         Left            =   -64110
         TabIndex        =   146
         ToolTipText     =   "Sample has been Printed"
         Top             =   3240
         Visible         =   0   'False
         Width           =   1410
      End
      Begin VB.Label lblViewOrganism 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -65250
         TabIndex        =   99
         Top             =   810
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Historical View of Organism"
         Height          =   195
         Index           =   58
         Left            =   -67200
         TabIndex        =   98
         Top             =   810
         Width           =   1920
      End
      Begin VB.Image imgSquareTick 
         Height          =   225
         Left            =   -63870
         Picture         =   "frmEditBacteriologyNew.frx":19B54
         Top             =   420
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Image imgSquareCross 
         Height          =   225
         Left            =   -63660
         Picture         =   "frmEditBacteriologyNew.frx":19E2A
         Top             =   420
         Visible         =   0   'False
         Width           =   210
      End
   End
   Begin VB.Image imgHGreen 
      Height          =   510
      Left            =   16980
      Picture         =   "frmEditBacteriologyNew.frx":1A100
      Top             =   360
      Width           =   480
   End
   Begin VB.Image imgHRed 
      Height          =   510
      Left            =   17220
      Picture         =   "frmEditBacteriologyNew.frx":1AE02
      Top             =   1080
      Width           =   480
   End
   Begin VB.Label lblNOPAS 
      AutoSize        =   -1  'True
      Caption         =   "NOPAS"
      Height          =   195
      Left            =   15150
      TabIndex        =   93
      Top             =   1080
      Width           =   555
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuListsMessages 
         Caption         =   "&Messages"
      End
      Begin VB.Menu mnuConsultantList 
         Caption         =   "&Consultant List"
      End
      Begin VB.Menu mnuListsSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuLists 
      Caption         =   "&Lists"
      Visible         =   0   'False
      Begin VB.Menu mnuListsUrine 
         Caption         =   "&Urine"
         Begin VB.Menu mnuListsUrineSub 
            Caption         =   "&Bacteria"
            Index           =   0
         End
         Begin VB.Menu mnuListsUrineSub 
            Caption         =   "&WCC"
            Index           =   1
         End
         Begin VB.Menu mnuListsUrineSub 
            Caption         =   "&RCC"
            Index           =   2
         End
         Begin VB.Menu mnuListsUrineSub 
            Caption         =   "Cr&ystals"
            Index           =   3
         End
         Begin VB.Menu mnuListsUrineSub 
            Caption         =   "&Casts"
            Index           =   4
         End
         Begin VB.Menu mnuListsUrineSub 
            Caption         =   "&Miscellaneous"
            Index           =   5
         End
         Begin VB.Menu mnuListsUrineSub 
            Caption         =   "-"
            Index           =   6
         End
         Begin VB.Menu mnuListsUrineSub 
            Caption         =   "Pre&gnancy"
            Index           =   7
         End
         Begin VB.Menu mnuListsUrineSub 
            Caption         =   "HCG &Level"
            Index           =   8
         End
      End
      Begin VB.Menu mnuListsIdentification 
         Caption         =   "Identification"
         Begin VB.Menu mnuListsIdentificationSub 
            Caption         =   "Gram Stains"
            Index           =   0
         End
         Begin VB.Menu mnuListsIdentificationSub 
            Caption         =   "Wet Prep"
            Index           =   1
         End
      End
      Begin VB.Menu mnuListsFaeces 
         Caption         =   "&Faeces"
         Begin VB.Menu mnuListsFaecesSub 
            Caption         =   "&XLD"
            Index           =   0
         End
         Begin VB.Menu mnuListsFaecesSub 
            Caption         =   "&DCA"
            Index           =   1
         End
         Begin VB.Menu mnuListsFaecesSub 
            Caption         =   "&SMAC"
            Index           =   2
         End
         Begin VB.Menu mnuListsFaecesSub 
            Caption         =   "&CROMO"
            Index           =   3
         End
         Begin VB.Menu mnuListsFaecesSub 
            Caption         =   "C&AMP"
            Index           =   4
         End
      End
      Begin VB.Menu mnuListsTitles 
         Caption         =   "&Titles"
         Begin VB.Menu mnuListsTitlesSub 
            Caption         =   "&FOB"
            Index           =   0
         End
         Begin VB.Menu mnuListsTitlesSub 
            Caption         =   "H.Pylori"
            Index           =   1
         End
         Begin VB.Menu mnuListsTitlesSub 
            Caption         =   "C. Diff &Culture"
            Index           =   2
         End
         Begin VB.Menu mnuListsTitlesSub 
            Caption         =   "C. Diff &ToxinAB"
            Index           =   3
         End
         Begin VB.Menu mnuListsTitlesSub 
            Caption         =   "&Rota"
            Index           =   4
         End
         Begin VB.Menu mnuListsTitlesSub 
            Caption         =   "&Adeno"
            Index           =   5
         End
         Begin VB.Menu mnuListsTitlesSub 
            Caption         =   "&RSV"
            Index           =   6
         End
         Begin VB.Menu mnuListsTitlesSub 
            Caption         =   "&Cryptosporidium"
            Index           =   7
         End
         Begin VB.Menu mnuListsTitlesSub 
            Caption         =   "OP Co&mments"
            Index           =   8
         End
      End
      Begin VB.Menu mnuListsFluids 
         Caption         =   "&Fluids"
         Begin VB.Menu mnuListsFluidsSub 
            Caption         =   "&Appearance"
            Index           =   0
         End
         Begin VB.Menu mnuListsFluidsSub 
            Caption         =   "Cell Co&unt"
            Index           =   1
         End
         Begin VB.Menu mnuListsFluidsSub 
            Caption         =   "&Gram Stain"
            Index           =   2
         End
         Begin VB.Menu mnuListsFluidsSub 
            Caption         =   "&ZN Stain"
            Index           =   3
         End
         Begin VB.Menu mnuListsFluidsSub 
            Caption         =   "&Leishman's Stain"
            Index           =   4
         End
         Begin VB.Menu mnuListsFluidsSub 
            Caption         =   "&Wet Prep"
            Index           =   5
         End
         Begin VB.Menu mnuListsFluidsSub 
            Caption         =   "&Crystals"
            Index           =   6
         End
         Begin VB.Menu mnuListsFluidsSub 
            Caption         =   "&Sites"
            Index           =   7
         End
      End
      Begin VB.Menu mnuCandS 
         Caption         =   "C && S"
         Begin VB.Menu mnuCandSSub 
            Caption         =   "Microbiology Sites"
            Index           =   0
         End
         Begin VB.Menu mnuCandSSub 
            Caption         =   "Organism Groups"
            Index           =   1
         End
         Begin VB.Menu mnuCandSSub 
            Caption         =   "Organisms"
            Index           =   2
         End
         Begin VB.Menu mnuCandSSub 
            Caption         =   "Antibiotics"
            Index           =   3
         End
         Begin VB.Menu mnuCandSSub 
            Caption         =   "Antibiotic Panels"
            Index           =   4
         End
         Begin VB.Menu mnuCandSSub 
            Caption         =   "Micro Setup"
            Index           =   5
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnuListsTabSetup 
         Caption         =   "Tab S&etup"
      End
   End
End
Attribute VB_Name = "frmEditMicrobiologyNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mNewRecord As Boolean

Private Activated As Boolean

Private pPrintToPrinter As String

Private UrineLoaded As Boolean
Private IdentLoaded As Boolean
Private FaecesLoaded As Boolean
Private HPyloriLoaded As Boolean
Private CSLoaded As Boolean
Private FOBLoaded As Boolean
Private RotaAdenoLoaded As Boolean
Private CdiffLoaded As Boolean
Private OPLoaded As Boolean
Private IdentificationLoaded As Boolean
Private FluidsLoaded As Boolean

Private SampleIDWithOffset As Double

Dim ListBacteria() As String
Dim ListPregnancy() As String
Dim ListOrganism() As String

Dim ListRCC() As ListColour
Dim ListWCC() As ListColour
Dim ListFOB() As ListColour
Dim ListHPylori() As ListColour
Dim ListCDiffCulture() As ListColour
Dim ListCDiffToxinAB() As ListColour
Dim ListGDH() As ListColour
Dim ListPCR() As ListColour
Dim ListRota() As ListColour
Dim ListAdeno() As ListColour
Dim ListRSV() As ListColour
Dim ListCrypto() As ListColour
Dim ListGiardia() As ListColour
Dim ListGDHDetail() As ListColour
Dim ListPCRDetail() As ListColour

Private ForceSaveability As Boolean
Private IQ200ResultsExist As Boolean

Private pForcedSID As Double

Private BacTek3DInUse As Boolean
Private ObservaInUse As Boolean
Private ReleaseToHealthlinkOnValidate As Boolean

Private Type AntibioticList
    Name As String
    Priority As Integer
End Type
Private Type ExclusionABList
    Antibiotic() As AntibioticList
    SelectedABCount As Integer
    AutoReportABCount As Integer
    ExclusionABCount As Integer
End Type


Private Sub CheckExternals()

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo CheckExternals_Error

20        sql = "SELECT COUNT (*) AS Tot FROM MicroExternalResults WHERE " & _
                "SampleID = '" & Val(txtSampleID) + SysOptMicroOffset(0) & "'"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        If tb!Tot > 0 Then
60            cmdViewExternal.BackColor = vbYellow
70        Else
80            cmdViewExternal.BackColor = vbButtonFace
90        End If

100       Exit Sub

CheckExternals_Error:

          Dim strES As String
          Dim intEL As Integer

110       intEL = Erl
120       strES = Err.Description
130       LogError "frmEditMicrobiologyNew", "CheckExternals", intEL, strES, sql

End Sub

Private Sub CheckObserva()

          Dim tb As Recordset
          Dim sql As String
          Dim n As Integer

10        On Error GoTo CheckObserva_Error

20        If Val(txtSampleID) = 0 Then Exit Sub

30        sql = "SELECT COUNT(*) Tot FROM BactOrders WHERE " & _
                "SampleID = '" & txtSampleID & "'"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        If tb!Tot > 0 Then
70            For n = 0 To 2
80                cmdObserva(n).Caption = "Requested"
90            Next
100       Else
110           For n = 0 To 2
120               cmdObserva(n).Caption = "Observa"
130           Next
140       End If

150       Exit Sub

CheckObserva_Error:

          Dim strES As String
          Dim intEL As Integer

160       intEL = Erl
170       strES = Err.Description
180       LogError "frmEditMicrobiologyNew", "CheckObserva", intEL, strES, sql

End Sub

Private Function DemographicsHaveChanged(ByVal SID As String) As Boolean

          Dim sql As String
          Dim tb As Recordset
          Dim RetVal As Boolean

10        RetVal = False

20        sql = "Select * from Demographics where " & _
                "SampleID = '" & SID & "'"

30        Set tb = New Recordset
40        RecOpenClient 0, tb, sql
50        If Not tb.EOF Then

60            If tb!RooH <> cRooH(0) Then RetVal = True

70            If IsDate(tRecTime) Then
80                If Format$(tb!RecDate, "HH:nn") <> Format$(tRecTime, "HH:nn") Then RetVal = True
90            End If
100           If Format$(tb!RecDate, "dd/MM/yyyy") <> Format$(dtRecDate, "dd/MM/yyyy") Then RetVal = True

110           If Format$(tb!Rundate, "dd/MM/yyyy") <> Format$(dtRunDate, "dd/MM/yyyy") Then RetVal = True

120           If IsDate(tSampleTime) Then
130               If Format$(tb!SampleDate, "HH:nn") <> Format$(tSampleTime, "HH:nn") Then RetVal = True
140           End If
150           If Format$(tb!SampleDate, "dd/MM/yyyy") <> Format$(dtSampleDate, "dd/MM/yyyy") Then RetVal = True

160           If Trim$(tb!Chart & "") <> Trim$(txtChart) Then RetVal = True
170           If Trim$(tb!PatName & "") <> Trim$(txtName) Then RetVal = True

180           If IsDate(tb!Dob & "") Then
190               If Format$(tb!Dob, "dd/MM/yyyy") <> Format$(txtDoB, "dd/MM/yyyy") Then RetVal = True
200           Else
210               If IsDate(txtDoB) Then RetVal = True
220           End If

230           If Trim$(tb!AandE & "") <> Trim$(txtAandE) Then RetVal = True
240           If Trim$(tb!Age & "") <> Trim$(txtAge) Then RetVal = True
250           If Trim$(tb!sex & "") <> Left$(txtSex, 1) Then RetVal = True
260           If Trim$(tb!Addr0 & "") <> Trim$(taddress(0)) Then RetVal = True
270           If Trim$(tb!Addr1 & "") <> Trim$(taddress(1)) Then RetVal = True
280           If Trim$(tb!Ward & "") <> Trim$(cmbWard) Then RetVal = True
290           If Trim$(tb!Clinician & "") <> Trim$(cmbClinician) Then RetVal = True
300           If Trim$(tb!GP & "") <> Trim$(cmbGP) Then RetVal = True
310           If Trim$(tb!ClDetails & "") <> Trim$(txtClinDetails) Then RetVal = True
320           If Trim$(tb!Hospital & "") <> Trim$(cmbHospital) Then RetVal = True
330           If tb!Pregnant <> chkPregnant Then RetVal = True
340           If tb!PenicillinAllergy <> chkPenicillin.Value Then RetVal = True

350       Else
360           RetVal = True    'must be a new record
370       End If
380       DemographicsHaveChanged = RetVal

End Function

Private Sub FillNameDemographics()

          Dim sql As String
          Dim tb As Recordset

10        On Error GoTo FillNameDemographics_Error

20        sql = "SELECT UserName, DateTimeDemographics FROM Demographics WHERE " & _
                "SampleID = '" & Val(txtSampleID) + SysOptMicroOffset(0) & "'"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        If Not tb.EOF Then
60            If Trim$(tb!UserName & "") <> "" Then
70                txtWhoSaved.Text = "Saved By " & Trim(tb!UserName & "") & " On " & tb!DateTimeDemographics
80                txtWhoSaved.Visible = True
90            End If
100       End If

110       Exit Sub

FillNameDemographics_Error:

          Dim strES As String
          Dim intEL As Integer

120       intEL = Erl
130       strES = Err.Description
140       LogError "frmEditMicrobiologyNew", "FillNameDemographics", intEL, strES, sql

End Sub

Private Sub FillNameFaeces(ByVal TabNo As Integer)

          Dim sql As String
          Dim tb As Recordset
          Dim Dept As String

10        On Error GoTo FillNameFaeces_Error

20        sql = "SELECT UserName FROM Faeces WHERE " & _
                "SampleID = '" & Val(txtSampleID) + SysOptMicroOffset(0) & "'"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        If Not tb.EOF Then
60            If Trim$(tb!UserName & "") <> "" Then
70                txtWhoSaved.Text = "Saved By " & tb!UserName
80                txtWhoSaved.Visible = True
90            End If
100       End If

110       Select Case TabNo
          Case 5: Dept = "F"
120       Case 6: Dept = "A"
130       Case 10: Dept = "G"
140       Case 11: Dept = "O"
150       Case 13: Dept = "Y"

160       End Select

170       sql = "SELECT * FROM PrintValidLog WHERE " & _
                "SampleID = '" & Val(txtSampleID) + SysOptMicroOffset(0) & "' AND Department = '" & Dept & "'"
180       Set tb = New Recordset
190       RecOpenServer 0, tb, sql
200       If Not tb.EOF Then
210           If Trim$(tb!ValidatedBy & "") <> "" Then
220               txtWhoValidated.Text = "Validated By " & Trim(tb!ValidatedBy & "") & " On " & tb!ValidatedDateTime
230               txtWhoValidated.Visible = True
240           End If
250           If Trim$(tb!PrintedBy & "") <> "" Then
260               txtWhoPrinted.Text = "Printed By " & Trim(tb!PrintedBy & "") & " On " & tb!PrintedDateTime
270               txtWhoPrinted.Visible = True
280           End If
290       End If


300       Exit Sub

FillNameFaeces_Error:

          Dim strES As String
          Dim intEL As Integer

310       intEL = Erl
320       strES = Err.Description
330       LogError "frmEditMicrobiologyNew", "FillNameFaeces", intEL, strES, sql

End Sub

Private Sub FillNameR(ByVal TestName As String)

          Dim sql As String
          Dim tb As Recordset
          Dim Dept As String

10        On Error GoTo FillNameR_Error

20        sql = "SELECT UserName FROM GenericResults WHERE " & _
                "SampleID = '" & Val(txtSampleID) + SysOptMicroOffset(0) & "' " & _
                "AND TestName = '" & TestName & "'"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        If Not tb.EOF Then
60            If Trim$(tb!UserName & "") <> "" Then
70                txtWhoSaved.Text = "Saved By " & tb!UserName
80                txtWhoSaved.Visible = True
90            End If
100       End If

110       If TestName = "RedSub" Then
120           Dept = "R"
130       ElseIf TestName = "RSV" Then
140           Dept = "V"
150       End If

160       sql = "SELECT * FROM PrintValidLog WHERE " & _
                "SampleID = '" & Val(txtSampleID) + SysOptMicroOffset(0) & "' AND Department = '" & Dept & "'"
170       Set tb = New Recordset
180       RecOpenServer 0, tb, sql
190       If Not tb.EOF Then
200           If Trim$(tb!ValidatedBy & "") <> "" Then
210               txtWhoValidated.Text = "Validated By " & Trim(tb!ValidatedBy & "") & " On " & tb!ValidatedDateTime
220               txtWhoValidated.Visible = True
230           End If
240           If Trim$(tb!PrintedBy & "") <> "" Then
250               txtWhoPrinted.Text = "Printed By " & Trim(tb!PrintedBy & "") & " On " & tb!PrintedDateTime
260               txtWhoPrinted.Visible = True
270           End If
280       End If


290       Exit Sub

FillNameR_Error:

          Dim strES As String
          Dim intEL As Integer

300       intEL = Erl
310       strES = Err.Description
320       LogError "frmEditMicrobiologyNew", "FillNameR", intEL, strES, sql

End Sub


Private Sub FillNameFluids()

          Dim sql As String
          Dim tb As Recordset

10        On Error GoTo FillNameFluids_Error

20        sql = "SELECT UserName FROM GenericResults WHERE " & _
                "SampleID = '" & Val(txtSampleID) + SysOptMicroOffset(0) & "' " & _
                "AND (TestName LIKE 'Fluid%' OR " & _
                "     TestName LIKE 'CSF%' OR " & _
                "     TestName = 'PneumococcalAT' OR " & _
                "     TestName = 'LegionellaAT' OR " & _
                "     TestName = 'FungalElements' )"

30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        If Not tb.EOF Then
60            If Trim$(tb!UserName & "") <> "" Then
70                txtWhoSaved.Text = "Saved By " & tb!UserName
80                txtWhoSaved.Visible = True
90            End If
100       End If


110       sql = "SELECT * FROM PrintValidLog WHERE " & _
                "SampleID = '" & Val(txtSampleID) + SysOptMicroOffset(0) & "' AND Department = 'G'"
120       Set tb = New Recordset
130       RecOpenServer 0, tb, sql
140       If Not tb.EOF Then
150           If Trim$(tb!ValidatedBy & "") <> "" Then
160               txtWhoValidated.Text = "Validated By " & Trim(tb!ValidatedBy & "") & " On " & tb!ValidatedDateTime
170               txtWhoValidated.Visible = True
180           End If
190           If Trim$(tb!PrintedBy & "") <> "" Then
200               txtWhoPrinted.Text = "Printed By " & Trim(tb!PrintedBy & "") & " On " & tb!PrintedDateTime
210               txtWhoPrinted.Visible = True
220           End If
230       End If


240       Exit Sub

FillNameFluids_Error:

          Dim strES As String
          Dim intEL As Integer

250       intEL = Erl
260       strES = Err.Description
270       LogError "frmEditMicrobiologyNew", "FillNameFluids", intEL, strES, sql

End Sub

Private Function FormatCSFResult(ByVal Result As String) As String

          Dim RetVal As String

10        RetVal = ""

20        If IsNumeric(Result) Then
30            If Val(Result) < 1 Then
40                RetVal = Format$(Result & "", "0.0##")
50            Else
60                RetVal = Result
70            End If
80        Else
90            RetVal = Result
100       End If

110       FormatCSFResult = RetVal

End Function

Private Function IsCystic(ByVal SampleIDWithOffset As Double) As Boolean

          Dim sql As String
          Dim tb As Recordset

10        On Error GoTo IsCystic_Error

20        IsCystic = False

30        sql = "SELECT COUNT (*) Tot FROM Observations WHERE " & _
                "SampleID = " & SampleIDWithOffset & " " & _
                "AND Comment = 'CYSTIC FIBROSIS PATIENT'"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        If tb!Tot > 0 Then
70            IsCystic = True
80        End If

90        Exit Function

IsCystic_Error:

          Dim strES As String
          Dim intEL As Integer

100       intEL = Erl
110       strES = Err.Description
120       LogError "frmEditMicrobiologyNew", "IsCystic", intEL, strES, sql

End Function

Private Function IsVitekOrdered(ByVal ANF As String) As Integer
      'returns -1 not ordered
      '         0 requested
      '         1 Programmed
      '         2 Resulted

          Dim tb As Recordset
          Dim sql As String
          Dim RetVal As Integer

10        On Error GoTo IsVitekOrdered_Error

20        If Trim$(ANF) = "" Then
30            IsVitekOrdered = -1
40            Exit Function
50        End If

60        sql = "SELECT Programmed FROM BactOrders WHERE " & _
                "SampleID = '" & txtSampleID & "' " & _
                "AND TestRequested = '" & ANF & "'"
70        Set tb = New Recordset
80        RecOpenServer 0, tb, sql
90        If tb.EOF Then
100           RetVal = -1
110       Else
120           RetVal = tb!Programmed
130       End If

140       sql = "SELECT COUNT(*) Cnt FROM Isolates WHERE " & _
                "SampleID = '" & Val(txtSampleID) + SysOptMicroOffset(0) & "' " & _
                "AND (IsolateNumber = "
150       If ANF = "A" Then
160           sql = sql & "1 OR IsolateNumber = 2)"
170       ElseIf ANF = "N" Then
180           sql = sql & "3 OR IsolateNumber = 4)"
190       ElseIf ANF = "F" Then
200           sql = sql & "5 OR IsolateNumber = 6)"
210       End If
220       Set tb = New Recordset
230       RecOpenServer 0, tb, sql
240       If tb!Cnt > 0 Then
250           RetVal = 2
260       End If

270       IsVitekOrdered = RetVal

280       Exit Function

IsVitekOrdered_Error:

          Dim strES As String
          Dim intEL As Integer

290       intEL = Erl
300       strES = Err.Description
310       LogError "frmEditMicrobiologyNew", "IsVitekOrdered", intEL, strES, sql

End Function

Private Function LoadFluids() As Boolean
      'returns true if anything loaded

          Dim n As Integer

10        On Error GoTo LoadFluids_Error

20        LoadFluids = False

30        cmdLock(9).Visible = False
40        fraCSF.Enabled = True
50        cmbFluidAppearance(0) = ""
60        cmbFluidAppearance(1) = ""
70        cmbFluidGram(0) = ""
80        cmbFluidGram(1) = ""
90        cmbFluidLeishmans = ""
100       cmbFluidWetPrep = ""
110       cmbFluidCrystals = ""
120       cmbZN = ""

130       For n = 0 To 11
140           txtHaem(n) = ""
150       Next
          'txtFluidComment = ""
160       txtInHouseSID = ""
170       For n = 0 To 7
180           chkBio(n).Value = 0
190           txtBioResult(n) = ""
200       Next

210       lblPneuAT = ""
220       lblLegionellaAT = ""
230       chkFungal(0).Value = 0
240       chkFungal(1).Value = 0
250       lblBATResult.Caption = ""
260       txtBATComments = ""

270       LoadFluids = LoadFluidBio

280       Exit Function

LoadFluids_Error:

          Dim strES As String
          Dim intEL As Integer

290       intEL = Erl
300       strES = Err.Description
310       LogError "frmEditMicrobiologyNew", "LoadFluids", intEL, strES

End Function

Private Sub LoadListFluidGram()

          Dim sql As String
          Dim tb As Recordset

10        On Error GoTo LoadListFluidGram_Error

20        cmbFluidGram(0).Clear
30        cmbFluidGram(1).Clear

40        sql = "SELECT * FROM Lists WHERE " & _
                "ListType = 'CG' " & _
                "ORDER BY ListOrder"
50        Set tb = New Recordset
60        RecOpenServer 0, tb, sql
70        Do While Not tb.EOF
80            cmbFluidGram(0).AddItem Trim(tb!Text & "")
90            cmbFluidGram(1).AddItem Trim(tb!Text & "")
100           tb.MoveNext
110       Loop

120       FixComboWidth cmbFluidGram(0)
130       FixComboWidth cmbFluidGram(1)

140       Exit Sub

LoadListFluidGram_Error:

          Dim strES As String
          Dim intEL As Integer

150       intEL = Erl
160       strES = Err.Description
170       LogError "frmEditMicrobiologyNew", "LoadListFluidGram", intEL, strES, sql

End Sub

Private Sub LoadListFluidCrystals()

          Dim sql As String
          Dim tb As Recordset

10        On Error GoTo LoadListFluidCrystals_Error

20        cmbFluidCrystals.Clear

30        sql = "SELECT * FROM Lists WHERE " & _
                "ListType = 'FC' " & _
                "ORDER BY ListOrder"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        Do While Not tb.EOF
70            cmbFluidCrystals.AddItem Trim(tb!Text & "")
80            tb.MoveNext
90        Loop

100       FixComboWidth cmbFluidCrystals

110       Exit Sub

LoadListFluidCrystals_Error:

          Dim strES As String
          Dim intEL As Integer

120       intEL = Erl
130       strES = Err.Description
140       LogError "frmEditMicrobiologyNew", "LoadListFluidCrystals", intEL, strES, sql

End Sub


Private Sub LoadListFluidLeishman()

          Dim sql As String
          Dim tb As Recordset

10        On Error GoTo LoadListFluidLeishman_Error

20        cmbFluidLeishmans.Clear

30        sql = "SELECT * FROM Lists WHERE " & _
                "ListType = 'CL' " & _
                "ORDER BY ListOrder"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        Do While Not tb.EOF
70            cmbFluidLeishmans.AddItem Trim(tb!Text & "")
80            tb.MoveNext
90        Loop

100       FixComboWidth cmbFluidLeishmans

110       Exit Sub

LoadListFluidLeishman_Error:

          Dim strES As String
          Dim intEL As Integer

120       intEL = Erl
130       strES = Err.Description
140       LogError "frmEditMicrobiologyNew", "LoadListFluidLeishman", intEL, strES, sql

End Sub


Private Sub LoadListFluidZN()

          Dim sql As String
          Dim tb As Recordset

10        On Error GoTo LoadListFluidZN_Error

20        cmbZN.Clear

30        sql = "SELECT * FROM Lists WHERE " & _
                "ListType = 'FluidZN' " & _
                "ORDER BY ListOrder"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        Do While Not tb.EOF
70            cmbZN.AddItem Trim(tb!Text & "")
80            tb.MoveNext
90        Loop

100       FixComboWidth cmbZN

110       Exit Sub

LoadListFluidZN_Error:

          Dim strES As String
          Dim intEL As Integer

120       intEL = Erl
130       strES = Err.Description
140       LogError "frmEditMicrobiologyNew", "LoadListFluidZN", intEL, strES, sql

End Sub

Private Sub LoadListFluidWetPrep()

          Dim sql As String
          Dim tb As Recordset

10        On Error GoTo LoadListFluidWetPrep_Error

20        cmbFluidWetPrep.Clear

30        sql = "SELECT * FROM Lists WHERE " & _
                "ListType = 'FW' " & _
                "ORDER BY ListOrder"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        Do While Not tb.EOF
70            cmbFluidWetPrep.AddItem Trim(tb!Text & "")
80            tb.MoveNext
90        Loop

100       FixComboWidth cmbFluidWetPrep

110       Exit Sub

LoadListFluidWetPrep_Error:

          Dim strES As String
          Dim intEL As Integer

120       intEL = Erl
130       strES = Err.Description
140       LogError "frmEditMicrobiologyNew", "LoadListFluidWetPrep", intEL, strES, sql

End Sub
Private Sub SaveFluid(ByVal cmb As ComboBox, ByVal TestName As String)

          Dim sql As String
          Dim tb As Recordset

10        On Error GoTo SaveFluid_Error

20        If cmb.Text <> "" Then
30            sql = "SELECT * FROM GenericResults WHERE " & _
                    "SampleID = '" & Val(txtSampleID) + SysOptMicroOffset(0) & "' " & _
                    "AND TestName = '" & TestName & "'"
40            Set tb = New Recordset
50            RecOpenClient 0, tb, sql
60            If tb.EOF Then
70                tb.AddNew
80            End If
90            tb!SampleID = Val(txtSampleID) + SysOptMicroOffset(0)
100           tb!TestName = TestName
110           tb!Result = cmb.Text
120           tb!UserName = UserName
130           tb.Update
140       Else
150           sql = "DELETE FROM GenericResults WHERE " & _
                    "SampleID = '" & Val(txtSampleID) + SysOptMicroOffset(0) & "' " & _
                    "AND TestName = '" & TestName & "'"
160           Cnxn(0).Execute sql
170       End If

180       Exit Sub

SaveFluid_Error:

          Dim strES As String
          Dim intEL As Integer

190       intEL = Erl
200       strES = Err.Description
210       LogError "frmEditMicrobiologyNew", "SaveFluid", intEL, strES, sql

End Sub

Private Sub SaveFluids()

          Dim n As Integer
          Dim TestName As String

10        On Error GoTo SaveFluids_Error

20        For n = 0 To 7
30            TestName = Choose(n + 1, "FluidGlucose", "FluidProtein", "FluidAlbumin", _
                                "FluidGlobulin", "FluidLDH", "FluidAmylase", _
                                "CSFGlucose", "CSFProtein")
40            SaveGenericResult TestName, txtBioResult(n)
50        Next

60        For n = 0 To 11
70            SaveGenericResult "CSFHaem" & Format$(n), txtHaem(n)
80        Next

90        SaveGenericResult "PneumococcalAT", lblPneuAT
100       SaveGenericResult "LegionellaAT", lblLegionellaAT

110       If chkFungal(0).Value = 1 Then
120           SaveGenericResult "FungalElements", "Seen"
130       ElseIf chkFungal(1).Value = 1 Then
140           SaveGenericResult "FungalElements", "Not Seen"
150       Else
160           SaveGenericResult "FungalElements", ""
170       End If

180       SaveGenericResult "BATScreen", lblBATResult
190       SaveGenericResult "BATScreenComment", txtBATComments

200       SaveFluid cmbFluidAppearance(0), "FluidAppearance0"
210       SaveFluid cmbFluidAppearance(1), "FluidAppearance1"
220       SaveFluid cmbFluidGram(0), "FluidGram"
230       SaveFluid cmbFluidGram(1), "FluidGram(2)"
240       SaveFluid cmbFluidLeishmans, "FluidLeishmans"
250       SaveFluid cmbFluidWetPrep, "FluidWetPrep"
260       SaveFluid cmbFluidCrystals, "FluidCrystals"
270       SaveFluid cmbZN, "FluidZN"

280       Exit Sub

SaveFluids_Error:

          Dim strES As String
          Dim intEL As Integer

290       intEL = Erl
300       strES = Err.Description
310       LogError "frmEditMicrobiologyNew", "SaveFluids", intEL, strES

End Sub

Private Sub SaveGenericResult(ByVal TestName As String, _
                              ByVal Result As String)

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo SaveGenericResults_Error

20        sql = "SELECT * FROM GenericResults WHERE " & _
                "SampleID = '" & Val(txtSampleID) + SysOptMicroOffset(0) & "' " & _
                "AND TestName = '" & TestName & "'"
30        Set tb = New Recordset
40        RecOpenClient 0, tb, sql

50        If Trim$(Result) <> "" Then
60            If tb.EOF Then
70                tb.AddNew
80            End If
90            tb!SampleID = Val(txtSampleID) + SysOptMicroOffset(0)
100           tb!TestName = TestName
110           tb!Result = Result
120           tb!UserName = UserName
130           tb.Update
140       Else
150           If Not tb.EOF Then
160               sql = "DELETE FROM GenericResults WHERE " & _
                        "SampleID = '" & Val(txtSampleID) + SysOptMicroOffset(0) & "' " & _
                        "AND TestName = '" & TestName & "'"
170               Cnxn(0).Execute sql
180           End If
190       End If

200       Exit Sub

SaveGenericResults_Error:

          Dim strES As String
          Dim intEL As Integer

210       intEL = Erl
220       strES = Err.Description
230       LogError "frmEditMicrobiologyNew", "SaveGenericResults", intEL, strES, sql

End Sub


Private Sub SaveMicroSiteDetails()

          Dim tb As Recordset
          Dim sql As String
          Dim DetailsPresent As Boolean
          Dim AddRecord As Boolean
          Dim AbsInUse(0 To 3) As String
          Dim n As Integer

10        On Error GoTo SaveMicroSiteDetails_Error

20        For n = 0 To 3
30            AbsInUse(n) = ""
40        Next
50        For n = 0 To lstABsInUse.ListCount - 1
60            If n < 4 Then
70                AbsInUse(n) = lstABsInUse.List(n)
80            End If
90        Next

100       AddRecord = False

110       DetailsPresent = False
120       If Trim$(cmbSite & txtSiteDetails & AbsInUse(0) & AbsInUse(1) & AbsInUse(2) & AbsInUse(3)) <> "" Then
130           DetailsPresent = True
140       End If

150       sql = "Select * from MicroSiteDetails where " & _
                "SampleID = '" & SampleIDWithOffset & "'"
160       Set tb = New Recordset
170       RecOpenClient 0, tb, sql
180       If tb.EOF Then
190           If DetailsPresent Then
200               tb.AddNew
210               AddRecord = True
220           End If
230       Else
240           If tb!Site & "" <> cmbSite Or _
                 tb!SiteDetails & "" <> txtSiteDetails Or _
                 tb!PCA0 & "" <> AbsInUse(0) Or _
                 tb!PCA1 & "" <> AbsInUse(1) Or _
                 tb!PCA2 & "" <> AbsInUse(2) Or _
                 tb!PCA3 & "" <> AbsInUse(3) Then
250               AddRecord = True
260           End If
270       End If
280       If AddRecord Then
290           tb!SampleID = SampleIDWithOffset
300           tb!Site = cmbSite
310           tb!SiteDetails = txtSiteDetails
320           tb!UserName = UserName
330           For n = 0 To 3
340               tb("PCA" & Format(n)) = AbsInUse(n)
350           Next
360           tb.Update
370       End If

380       Exit Sub

SaveMicroSiteDetails_Error:

          Dim strES As String
          Dim intEL As Integer

390       intEL = Erl
400       strES = Err.Description
410       LogError "frmEditMicrobiologyNew", "SaveMicroSiteDetails", intEL, strES, sql

End Sub

Private Sub SetComboWidths()

          Dim n As Integer

10        For n = 1 To 4
20            SetComboDropDownWidth cmbOrgGroup(n)
30            SetComboDropDownWidth cmbOrgName(n)
40            SetComboDropDownWidth cmbQualifier(n)
50        Next

End Sub
Private Sub ShowWhoSaved(ByVal Index As Integer)

10        txtWhoSaved.Visible = False
20        txtWhoValidated.Visible = False
30        txtWhoPrinted.Visible = False

40        Select Case Index
          Case 0: FillNameDemographics
50        Case 1: FillNameUrine
60        Case 2: FillNameUrineIdent
70        Case 3: FillNameFaecesWorkSheet
80        Case 4: FillNameSensitivities
90        Case 5, 6: FillNameFaeces Index
100       Case 7: FillNameR "RedSub"
110       Case 8: FillNameR "RSV"
120       Case 9: FillNameFluids
130       Case 10, 11, 13: FillNameFaeces Index
140       End Select

End Sub

Private Sub FillNameUrine()

          Dim sql As String
          Dim tb As Recordset

10        On Error GoTo FillNameUrine_Error


20        sql = "SELECT UserName FROM Urine WHERE " & _
                "SampleID = '" & Val(txtSampleID) + SysOptMicroOffset(0) & "'"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        If Not tb.EOF Then
60            If Trim$(tb!UserName & "") <> "" Then
70                txtWhoSaved.Text = "Saved By " & tb!UserName
80                txtWhoSaved.Visible = True
90            End If
100       End If

110       sql = "SELECT * FROM PrintValidLog WHERE " & _
                "SampleID = '" & Val(txtSampleID) + SysOptMicroOffset(0) & "' AND Department = 'U'"
120       Set tb = New Recordset
130       RecOpenServer 0, tb, sql
140       If Not tb.EOF Then
150           If Trim$(tb!ValidatedBy & "") <> "" Then
160               txtWhoValidated.Text = "Validated By " & Trim(tb!ValidatedBy & "") & " On " & tb!ValidatedDateTime
170               txtWhoValidated.Visible = True
180           End If
190           If Trim$(tb!PrintedBy & "") <> "" Then
200               txtWhoPrinted.Text = "Printed By " & Trim(tb!PrintedBy & "") & " On " & tb!PrintedDateTime
210               txtWhoPrinted.Visible = True
220           End If
230       End If


240       Exit Sub

FillNameUrine_Error:

          Dim strES As String
          Dim intEL As Integer

250       intEL = Erl
260       strES = Err.Description
270       LogError "frmEditMicrobiologyNew", "FillNameUrine", intEL, strES, sql

End Sub

Private Sub FillNameUrineIdent()

          Dim sql As String
          Dim tb As Recordset

10        On Error GoTo FillNameUrineIdent_Error

20        sql = "SELECT UserName FROM UrineIdent WHERE " & _
                "SampleID = '" & Val(txtSampleID) + SysOptMicroOffset(0) & "'"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        If Not tb.EOF Then
60            If Trim$(tb!UserName & "") <> "" Then
70                txtWhoSaved.Text = "Saved By " & tb!UserName
80                txtWhoSaved.Visible = True
90            End If
100       End If

110       Exit Sub

FillNameUrineIdent_Error:

          Dim strES As String
          Dim intEL As Integer

120       intEL = Erl
130       strES = Err.Description
140       LogError "frmEditMicrobiologyNew", "FillNameUrineIdent", intEL, strES, sql

End Sub

Private Sub FillNameFaecesWorkSheet()

          Dim sql As String
          Dim tb As Recordset

10        On Error GoTo FillNameFaecesWorkSheet_Error

20        sql = "SELECT Operator FROM FaecesWorkSheet WHERE " & _
                "SampleID = '" & Val(txtSampleID) + SysOptMicroOffset(0) & "'"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        If Not tb.EOF Then
60            If Trim$(tb!Operator & "") <> "" Then
70                txtWhoSaved.Text = "Saved By " & tb!Operator
80                txtWhoSaved.Visible = True
90            End If
100       End If

110       Exit Sub

FillNameFaecesWorkSheet_Error:

          Dim strES As String
          Dim intEL As Integer

120       intEL = Erl
130       strES = Err.Description
140       LogError "frmEditMicrobiologyNew", "FillNameFaecesWorkSheet", intEL, strES, sql

End Sub

Private Sub FillNameSensitivities()

          Dim sql As String
          Dim tb As Recordset

10        On Error GoTo FillNameSensitivities_Error

20        sql = "SELECT TOP 1 UserName FROM Sensitivities WHERE " & _
                "SampleID = '" & Val(txtSampleID) + SysOptMicroOffset(0) & "' " & _
                "AND COALESCE(UserName, '') <> ''"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        If Not tb.EOF Then
60            txtWhoSaved.Text = "Saved By " & tb!UserName
70            txtWhoSaved.Visible = True
80        End If

90        sql = "SELECT * FROM PrintValidLog WHERE " & _
                "SampleID = '" & Val(txtSampleID) + SysOptMicroOffset(0) & "' AND Department = 'D'"
100       Set tb = New Recordset
110       RecOpenServer 0, tb, sql
120       If Not tb.EOF Then
130           If Trim$(tb!ValidatedBy & "") <> "" Then
140               txtWhoValidated.Text = "Validated By " & Trim(tb!ValidatedBy & "") & " On " & tb!ValidatedDateTime
150               txtWhoValidated.Visible = True
160           End If
170           If Trim$(tb!PrintedBy & "") <> "" Then
180               txtWhoPrinted.Text = "Printed By " & Trim(tb!PrintedBy & "") & " On " & tb!PrintedDateTime
190               txtWhoPrinted.Visible = True
200           End If
210       End If


220       Exit Sub

FillNameSensitivities_Error:

          Dim strES As String
          Dim intEL As Integer

230       intEL = Erl
240       strES = Err.Description
250       LogError "frmEditMicrobiologyNew", "FillNameSensitivities", intEL, strES, sql

End Sub

Private Sub chkBio_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)

          Dim n As Integer

End Sub


Private Sub chkFungal_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)

On Error GoTo chkFungal_MouseUp_Error
10        chkFungal(Abs(Index - 1)).Value = 0
20        ShowUnlock 9




Exit Sub

chkFungal_MouseUp_Error:
Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "frmEditMicrobiologyNew", "chkFungal_MouseUp", intEL, strES

End Sub

Private Sub chkNonReportable_Click(Index As Integer)

10        cmdSaveMicro.Enabled = True
20        cmdSaveHold.Enabled = True

End Sub

Private Sub chkPenicillin_Click()

10        If chkPenicillin.Value = 1 And InStr(txtClinDetails, "Penicillin") = 0 Then
20            txtClinDetails = txtClinDetails & " Allergic to Penicillin;"
30        End If

40        cmdSaveDemographics.Enabled = True

End Sub

Private Sub cmbClinDetails_KeyPress(KeyAscii As Integer)
'10        KeyAscii = AutoComplete(cmbClinDetails, KeyAscii, False)

End Sub

Private Sub cmbDay1_LostFocus(Index As Integer)
'10        cmbDay1(Index).Text = QueryCombo(cmbDay1(Index))
End Sub

Private Sub cmbDay2_LostFocus(Index As Integer)
'10        cmbDay2(Index).Text = QueryCombo(cmbDay2(Index))
End Sub

Private Sub cmbDay3_LostFocus(Index As Integer)
'10        cmbDay3(Index).Text = QueryCombo(cmbDay3(Index))
End Sub

Private Sub cmbFluidAppearance_Click(Index As Integer)

10        cmdSaveHold.Enabled = True
20        cmdSaveMicro.Enabled = True

30        ShowUnlock 9

End Sub


Private Sub cmbFluidAppearance_KeyPress(Index As Integer, KeyAscii As Integer)

10        cmdSaveHold.Enabled = True
20        cmdSaveMicro.Enabled = True

30        ShowUnlock 9

End Sub


Private Sub cmbFluidCrystals_Click()

10        cmdSaveHold.Enabled = True
20        cmdSaveMicro.Enabled = True

30        ShowUnlock 9

End Sub


Private Sub cmbFluidCrystals_KeyPress(KeyAscii As Integer)

10        cmdSaveHold.Enabled = True
20        cmdSaveMicro.Enabled = True

30        ShowUnlock 9

End Sub


Private Sub cmbFluidGram_Click(Index As Integer)

10        cmdSaveHold.Enabled = True
20        cmdSaveMicro.Enabled = True

30        ShowUnlock 9

End Sub


Private Sub cmbFluidGram_KeyPress(Index As Integer, KeyAscii As Integer)

10        cmdSaveHold.Enabled = True
20        cmdSaveMicro.Enabled = True

30        ShowUnlock 9

End Sub


Private Sub cmbFluidLeishmans_Click()

10        cmdSaveHold.Enabled = True
20        cmdSaveMicro.Enabled = True

30        ShowUnlock 9

End Sub


Private Sub cmbFluidLeishmans_KeyPress(KeyAscii As Integer)

10        cmdSaveHold.Enabled = True
20        cmdSaveMicro.Enabled = True

30        ShowUnlock 9

End Sub


Private Sub cmbFluidWetPrep_Click()

10        cmdSaveHold.Enabled = True
20        cmdSaveMicro.Enabled = True

30        ShowUnlock 9

End Sub


Private Sub cmbFluidWetPrep_KeyPress(KeyAscii As Integer)

10        cmdSaveHold.Enabled = True
20        cmdSaveMicro.Enabled = True

30        ShowUnlock 9

End Sub


Private Sub cmbHospital_KeyPress(KeyAscii As Integer)

10        On Error GoTo cmbHospital_KeyPress_Error

      '20        KeyAscii = AutoComplete(cmbHospital, KeyAscii, False)

20        Exit Sub

cmbHospital_KeyPress_Error:

          Dim strES As String
          Dim intEL As Integer

30        intEL = Erl
40        strES = Err.Description
50        LogError "frmEditMicrobiologyNew", "cmbHospital_KeyPress", intEL, strES

End Sub

Private Sub cmbHospital_LostFocus()

          Dim n As Integer

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

80        intEL = Erl
90        strES = Err.Description
100       LogError "frmEditMicrobiologyNew", "cmbHospital_LostFocus", intEL, strES

End Sub


Private Sub cmbSiteSearch_KeyPress(KeyAscii As Integer)

10        KeyAscii = 0

End Sub


Private Sub cmbZN_Click()

10        cmdSaveHold.Enabled = True
20        cmdSaveMicro.Enabled = True

30        ShowUnlock 9

End Sub


Private Sub cmbZN_KeyPress(KeyAscii As Integer)

10        cmdSaveHold.Enabled = True
20        cmdSaveMicro.Enabled = True

30        ShowUnlock 9

End Sub


Private Sub cmdArchive_Click()

10        With frmAuditMicro
20            .SampleID = Val(txtSampleID)
30            .Show 1
40        End With

End Sub



Private Sub cmdBloodCulture_Click(Index As Integer)

On Error GoTo cmdBloodCulture_Click_Error
Select Case Index
    Case 0
        With frmEditBloodCulture
            .lblSampleID = txtSampleID
            .lblName = txtName
            .lblChart = txtChart
            .lblAandE = txtAandE
            .lblClinician = cmbClinician
            .lblDoB = txtDoB
            .lblAge = lblAge
            .lblSex = lblSex
            .lblWard = cmbWard
            .lblGP = cmbGP
            .lblABsInUse = lblABsInUse
            .Show 1
        End With
    Case 1
        With frmEditIdentification
            .lblSampleID = txtSampleID
            .lblName = txtName
            .lblChart = txtChart
            .lblAandE = txtAandE
            .lblClinician = cmbClinician
            .lblDoB = txtDoB
            .lblAge = lblAge
            .lblSex = lblSex
            .lblWard = cmbWard
            .lblGP = cmbGP
            .lblABsInUse = lblABsInUse
            .Show 1
        End With
End Select
Exit Sub

cmdBloodCulture_Click_Error:
Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "frmEditMicrobiologyNew", "cmdBloodCulture_Click", intEL, strES

End Sub

Private Sub cmdDelete_Click(Index As Integer)

          Dim s As String
          Dim sql As String

10        On Error GoTo cmdDelete_Click_Error

20        s = "You are about to remove this Organism and all its Sensitivities." & vbCrLf & _
              "You will not be able to undo this action." & vbCrLf & vbCrLf & _
              "Do you want to proceed?"
30        If iMsg(s, vbQuestion + vbYesNo, "Confirmation Required", vbRed) = vbYes Then
40            If UCase$(iBOX("Please enter your password.", "Confirmation Required", , True)) = UCase$(UserPass) Then

50                sql = "DELETE FROM Sensitivities WHERE " & _
                        "SampleID = '" & Val(txtSampleID) + SysOptMicroOffset(0) & "' " & _
                        "AND IsolateNumber = '" & Index & "'"
60                Cnxn(0).Execute sql

70                sql = "DELETE FROM Isolates WHERE " & _
                        "SampleID = '" & Val(txtSampleID) + SysOptMicroOffset(0) & "' " & _
                        "AND IsolateNumber = '" & Index & "'"
80                Cnxn(0).Execute sql

90                LoadAllDetails

100           Else
110               iMsg "Incorrect Password", vbInformation
120           End If

130       Else
140           iMsg "Action cancelled", vbInformation
150       End If

160       Exit Sub

cmdDelete_Click_Error:

          Dim strES As String
          Dim intEL As Integer

170       intEL = Erl
180       strES = Err.Description
190       LogError "frmEditMicrobiologyNew", "cmdDelete_Click", intEL, strES, sql, "Index=" & Format$(Index)

End Sub

Private Sub cmdDeleteIQ200_Click()

          Dim sql As String

10        On Error GoTo cmdDeleteIQ200_Click_Error

20        If iMsg("This will delete current IQ200 results, Do you want to proceed?", vbQuestion + vbYesNo) = vbYes Then
30            If UCase(iBOX("Password required", , , True)) = UserPass Then
40                sql = "DELETE From IQ200 Where SampleID = '" & txtSampleID + SysOptMicroOffset(0) & "'"
50                Cnxn(0).Execute sql
60                LoadUrine
70            End If
80        End If

90        Exit Sub

cmdDeleteIQ200_Click_Error:

          Dim strES As String
          Dim intEL As Integer

100       intEL = Erl
110       strES = Err.Description
120       LogError "frmEditMicrobiologyNew", "cmdDeleteIQ200_Click", intEL, strES

End Sub

Private Sub cmdDeleteMicroscopy_Click()

          Dim s As String
          Dim sql As String

10        On Error GoTo cmdDeleteMicroscopy_Click_Error

20        s = "You are about to remove all Microscopy results for this Sample." & vbCrLf & _
              "You will not be able to undo this action." & vbCrLf & vbCrLf & _
              "Do you want to proceed?"
30        If iMsg(s, vbQuestion + vbYesNo, "Confirmation Required", vbRed) = vbYes Then
40            If UCase$(iBOX("Please enter your password.", "Confirmation Required", , True)) = UCase$(UserPass) Then

50                sql = "DELETE FROM Urine WHERE " & _
                        "SampleID = '" & Val(txtSampleID) + SysOptMicroOffset(0) & "'"
60                Cnxn(0).Execute sql

70                LoadAllDetails

80            Else
90                iMsg "Incorrect Password", vbInformation
100           End If

110       Else
120           iMsg "Action cancelled", vbInformation
130       End If

140       Exit Sub

cmdDeleteMicroscopy_Click_Error:

          Dim strES As String
          Dim intEL As Integer

150       intEL = Erl
160       strES = Err.Description
170       LogError "frmEditMicrobiologyNew", "cmdDeleteMicroscopy_Click", intEL, strES, sql


End Sub


Private Sub cmdHealthLink_Click()

Dim SID As String

On Error GoTo cmdHealthLink_Click_Error


If cmdReleaseReport.BackColor = vbGreen Then
    iMsg "Report is released to consultant, cannot be released to healthlink until authorised"
    Exit Sub
End If

SID = Format$(Val(txtSampleID) + SysOptMicroOffset(0))

With cmdHealthLink
    If .Picture = imgHGreen.Picture Then
        Set .Picture = imgHRed.Picture
        ReleaseMicro SID, 0
    Else
        If lblInterim.BackColor = vbGreen Then
            ReleaseMicro SID, 1
        Else
            ReleaseMicro SID, 2
        End If
        Set .Picture = imgHGreen.Picture
        
    End If
End With

Exit Sub
cmdHealthLink_Click_Error:
   
LogError "frmEditMicrobiologyNew", "cmdHealthLink_Click", Erl, Err.Description



End Sub
Private Sub cmdIQ200Repeats_Click()

10        With frmViewIQ200Repeat
20            .SampleID = Val(txtSampleID) + SysOptMicroOffset(0)
30            .Show 1
40        End With

50        LoadIQ200

End Sub





Private Sub cmdObserva_Click(Index As Integer)

      '0 - Identification
      '1 - C & S
      '2 - Faeces

10        On Error GoTo cmdObserva_Click_Error

20        If Index = 0 Then
30            If UCase$(cmbSite) = "BLOOD CULTURE" Then
40                SSTab1.Tab = 12
50                Exit Sub
60            End If
70        End If

80        OrderOnObserva txtSampleID

90        Exit Sub

cmdObserva_Click_Error:

          Dim strES As String
          Dim intEL As Integer

100       intEL = Erl
110       strES = Err.Description
120       LogError "frmEditMicrobiologyNew", "cmdObserva_Click", intEL, strES

End Sub

Private Sub cmdOrderInHouse_Click()

          Dim Ordered As Boolean
          Dim n As Integer

10        If Trim$(txtInHouseSID) = "" Then
20            iMsg "Enter In House Sample ID", vbCritical, "NetAcquire - Error"
30            Exit Sub
40        End If

50        Ordered = False

60        For n = 0 To 7
70            If chkBio(n).Value = 1 Then
80                OrderBio
90                OrderDemographics
100               LogExternals
110               Ordered = True
120               Exit For
130           End If
140       Next

150       If Not Ordered Then
160           iMsg "Nothing to do!", vbExclamation
170           cmdOrderInHouse.Caption = "Order Tests"
180           cmdOrderInHouse.BackColor = vbButtonFace
190       Else
200           cmdOrderInHouse.Caption = "Ordered"
210           cmdOrderInHouse.BackColor = vbYellow
220       End If

End Sub

Private Sub LogExternals()

          Dim sql As String
          Dim tb As New Recordset

10        On Error GoTo LogExternals_Error

20        sql = "SELECT * FROM MicroExternals WHERE " & _
                "MicroSID = '" & Val(txtSampleID) + SysOptMicroOffset(0) & "'"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        If tb.EOF Then
60            tb.AddNew
70        End If
80        tb!MicroSID = Val(txtSampleID) + SysOptMicroOffset(0)
90        tb!InHouseSID = txtInHouseSID
100       tb!OrderGlu = chkBio(0) = 1
110       tb!OrderTP = chkBio(1) = 1
120       tb!OrderAlb = chkBio(2) = 1
130       tb!OrderGlo = chkBio(3) = 1
140       tb!OrderLDH = chkBio(4) = 1
150       tb!OrderAmy = chkBio(5) = 1
160       tb!OrderCSFGlu = chkBio(6) = 1
170       tb!OrderCSFTP = chkBio(7) = 1
180       tb!UserName = UserName
190       tb.Update

200       Exit Sub

LogExternals_Error:

          Dim strES As String
          Dim intEL As Integer

210       intEL = Erl
220       strES = Err.Description
230       LogError "frmEditMicrobiologyNew", "LogExternals", intEL, strES, sql

End Sub

Private Function LoadFluidBio() As Boolean

          Dim tb As Recordset
          Dim sql As String
          Dim br As BIEResult
          Dim BRs As New BIEResults

10        On Error GoTo LoadFluidBio_Error

20        LoadFluidBio = False

30        cmdOrderInHouse.Caption = "Order Tests"
40        cmdOrderInHouse.BackColor = vbButtonFace

50        sql = "SELECT * FROM MicroExternals WHERE " & _
                "MicroSID = '" & Val(txtSampleID) + SysOptMicroOffset(0) & "'"
60        Set tb = New Recordset
70        RecOpenServer 0, tb, sql
80        If Not tb.EOF Then
90            LoadFluidBio = True
100           txtInHouseSID = tb!InHouseSID
110           cmdOrderInHouse.Caption = "Ordered"
120           cmdOrderInHouse.BackColor = vbYellow
130           chkBio(0) = IIf(tb!OrderGlu <> 0, 1, 0)
140           chkBio(1) = IIf(tb!OrderTP <> 0, 1, 0)
150           chkBio(2) = IIf(tb!OrderAlb <> 0, 1, 0)
160           chkBio(3) = IIf(tb!OrderGlo <> 0, 1, 0)
170           chkBio(4) = IIf(tb!OrderLDH <> 0, 1, 0)
180           chkBio(5) = IIf(tb!OrderAmy <> 0, 1, 0)
190           chkBio(6) = IIf(tb!OrderCSFGlu <> 0, 1, 0)
200           chkBio(7) = IIf(tb!OrderCSFTP <> 0, 1, 0)
210       End If

220       Set BRs = BRs.Load("Bio", txtInHouseSID, "Results", gVALID, gDONTCARE, 0, "", "")
230       If BRs.Count <> 0 Then
240           LoadFluidBio = True
250           For Each br In BRs
260               Select Case br.LongName
                  Case "Glucose": txtBioResult(0) = br.Result
270               Case "Protein": txtBioResult(1) = br.Result
280               Case "Albumin": txtBioResult(2) = br.Result
290               Case "Globulin": txtBioResult(3) = br.Result
300               Case "LDH": txtBioResult(4) = br.Result
310               Case "Amylase": txtBioResult(5) = br.Result
320               Case "CSF Glucose": txtBioResult(6) = br.Result
330               Case "CSF Protein": txtBioResult(7) = br.Result
340               End Select
350           Next
360       End If

370       sql = "SELECT * FROM GenericResults WHERE " & _
                "SampleID = '" & Val(txtSampleID) + SysOptMicroOffset(0) & "' " & _
                "AND Result IS NOT NULL"
380       Set tb = New Recordset
390       RecOpenServer 0, tb, sql
400       Do While Not tb.EOF
410           LoadFluidBio = True
420           Select Case tb!TestName

              Case "PneumococcalAT": lblPneuAT = tb!Result & ""
430           Case "LegionellaAT": lblLegionellaAT = tb!Result & ""
440           Case "FungalElements":
450               If tb!Result & "" = "Seen" Then
460                   chkFungal(0).Value = 1
470               ElseIf tb!Result & "" = "Not Seen" Then
480                   chkFungal(1).Value = 1
490               Else
500                   chkFungal(0).Value = 0
510                   chkFungal(1).Value = 0
520               End If
530           Case "BATScreen": lblBATResult = tb!Result & ""
540           Case "BATScreenComment": txtBATComments = tb!Result & ""
550           Case "FluidGlucose": txtBioResult(0) = FormatCSFResult(tb!Result)
560           Case "FluidProtein": txtBioResult(1) = FormatCSFResult(tb!Result)
570           Case "FluidAlbumin": txtBioResult(2) = FormatCSFResult(tb!Result)
580           Case "FluidGlobulin": txtBioResult(3) = FormatCSFResult(tb!Result)
590           Case "FluidLDH": txtBioResult(4) = FormatCSFResult(tb!Result)
600           Case "FluidAmylase": txtBioResult(5) = FormatCSFResult(tb!Result)
610           Case "CSFGlucose": txtBioResult(6) = FormatCSFResult(tb!Result)
620           Case "CSFProtein": txtBioResult(7) = FormatCSFResult(tb!Result)

630           Case "FluidAppearance0": cmbFluidAppearance(0) = tb!Result & ""
640           Case "FluidAppearance1": cmbFluidAppearance(1) = tb!Result & ""
650           Case "FluidGram": cmbFluidGram(0) = tb!Result & ""
660           Case "FluidGram(2)": cmbFluidGram(1) = tb!Result & ""
670           Case "FluidLeishmans": cmbFluidLeishmans = tb!Result & ""
680           Case "FluidZN": cmbZN = tb!Result & ""
690           Case "FluidWetPrep": cmbFluidWetPrep = tb!Result & ""
700           Case "FluidCrystals": cmbFluidCrystals = tb!Result & ""

710           Case "CSFHaem0": txtHaem(0) = tb!Result & ""
720           Case "CSFHaem1": txtHaem(1) = tb!Result & ""
730           Case "CSFHaem2": txtHaem(2) = tb!Result & ""
740           Case "CSFHaem3": txtHaem(3) = tb!Result & ""
750           Case "CSFHaem4": txtHaem(4) = tb!Result & ""
760           Case "CSFHaem5": txtHaem(5) = tb!Result & ""
770           Case "CSFHaem6": txtHaem(6) = tb!Result & ""
780           Case "CSFHaem7": txtHaem(7) = tb!Result & ""
790           Case "CSFHaem8": txtHaem(8) = tb!Result & ""
800           Case "CSFHaem9": txtHaem(9) = tb!Result & ""
810           Case "CSFHaem10": txtHaem(10) = tb!Result & ""
820           Case "CSFHaem11": txtHaem(11) = tb!Result & ""
830           End Select
840           tb.MoveNext
850       Loop

860       Exit Function

LoadFluidBio_Error:

          Dim strES As String
          Dim intEL As Integer

870       intEL = Erl
880       strES = Err.Description
890       LogError "frmEditMicrobiologyNew", "LoadFluidBio", intEL, strES, sql

End Function
Private Sub OrderDemographics()

          Dim tb As Recordset
          Dim tbM As Recordset
          Dim sql As String
          Dim fld As Field

10        On Error GoTo OrderDemographics_Error

20        sql = "Select * from Demographics where " & _
                "SampleID = '" & Val(txtSampleID) + SysOptMicroOffset(0) & "'"
30        Set tbM = New Recordset
40        RecOpenClient 0, tbM, sql
50        If tbM.EOF Then Exit Sub

60        sql = "SELECT * FROM Demographics WHERE SampleID = '" & txtInHouseSID & "'"
70        Set tb = New Recordset
80        RecOpenClient 0, tb, sql
90        If tb.EOF Then
100           tb.AddNew
110       End If
120       For Each fld In tb.Fields
130           Select Case UCase(fld.Name)
              Case "SAMPLEID":
140               tb!SampleID = Val(txtInHouseSID)
150           Case "ADDR0":
160               tb!Addr0 = "Micro Lab"
170           Case "ADDR1":
180               tb!Addr1 = "Micro Lab"
190           Case "WARD":
200               tb!Ward = "Micro Lab"
210           Case "CLINICIAN":
220               tb!Clinician = "Micro Lab"
230           Case "GP":
240               tb!GP = "Micro Lab"
250           Case Else:
260               tb.Fields(fld.Name).Value = tbM.Fields(fld.Name).Value
270           End Select
280       Next
290       tb.Update

300       Exit Sub

OrderDemographics_Error:

          Dim strES As String
          Dim intEL As Integer

310       intEL = Erl
320       strES = Err.Description
330       LogError "frmEditMicrobiologyNew", "OrderDemographics", intEL, strES, sql

End Sub

Private Sub OrderBio()

          Dim n As Long
          Dim Code As String
          Dim sql As String

10        On Error GoTo OrderBio_Error

20        Cnxn(0).Execute ("DELETE from BioRequests WHERE " & _
                           "SampleID = '" & txtInHouseSID & "' " & _
                           "and Programmed = 0")

30        For n = 0 To 7
40            If chkBio(n).Value = 1 Then
50                Code = CodeForLongName(Choose(n + 1, "Glucose", "Protein", _
                                                "Albumin", "Globulin", "LDH", "Amylase", _
                                                "CSF Glucose", "CSF Protein"))

60                sql = "INSERT into BioRequests " & _
                        "(SampleID, Code, DateTime, SampleType, Programmed, AnalyserID,Hospital) VALUES " & _
                        "('" & txtInHouseSID & "', " & _
                        "'" & Code & "', " & _
                        "'" & Format$(Now, "yyyyMMdd HH:mm") & "', " & _
                        "'" & IIf(n < 6, "S", "C") & "', " & _
                        "'0', " & _
                        "'" & BioAnalyserIDForCode(Code) & "', " & _
                        "'" & GetHospitalName(Code, "Bio") & "')" 'added Hospital Name Trevor:
70                Cnxn(0).Execute sql
80            End If
90        Next

100       Exit Sub

OrderBio_Error:

          Dim strES As String
          Dim intEL As Integer

110       intEL = Erl
120       strES = Err.Description
130       LogError "frmEditMicrobiologyNew", "OrderBio", intEL, strES, sql

End Sub


Private Sub cmdDemoVal_Click()

          Dim Validating As Boolean

10        On Error GoTo cmdDemoVal_Click_Error

20        Validating = cmdDemoVal.Caption = "&Validate"



30        If Validating Then
40            If Not EntriesOK(txtSampleID, txtName, txtSex, cmbWard.Text, cmbGP.Text) Then
50                Exit Sub
60            End If
70            cmdSaveDemographics_Click
80        End If

90        ValidateDemographics Validating

100       If Validating Then
110           txtSampleID = Format$(Val(txtSampleID) + 1)
120           txtSampleID.SelStart = 0
130           txtSampleID.SelLength = Len(txtSampleID)
140           txtSampleID.SetFocus
150       End If

160       LoadAllDetails

170       cmdSaveInc.Enabled = False
180       cmdSaveDemographics.Enabled = False

190       Exit Sub

cmdDemoVal_Click_Error:

          Dim strES As String
          Dim intEL As Integer

200       intEL = Erl
210       strES = Err.Description
220       LogError "frmEditMicrobiologyNew", "cmdDemoVal_Click", intEL, strES

End Sub

Private Sub FillDemographicComments()

          Dim sql As String
          Dim tb As Recordset

10        On Error GoTo FillDemographicComments_Error

20        cmbDemogComment.Clear

30        sql = "Select * from Lists where " & _
                "ListType = 'DE'"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        Do While Not tb.EOF
70            cmbDemogComment.AddItem tb!Text & ""
80            tb.MoveNext
90        Loop

100       FixComboWidth cmbDemogComment
110       Exit Sub

FillDemographicComments_Error:

          Dim strES As String
          Dim intEL As Integer

120       intEL = Erl
130       strES = Err.Description
140       LogError "frmEditMicrobiologyNew", "FillDemographicComments", intEL, strES, sql

End Sub

Private Sub FillUrineComments()

          Dim sql As String
          Dim tb As Recordset

10        On Error GoTo FillUrineComments_Error

20        cmbUrineComment.Clear

30        sql = "Select * from Lists where " & _
                "ListType = 'UC' ORDER BY ListOrder"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        Do While Not tb.EOF
70            cmbUrineComment.AddItem tb!Text & ""
80            tb.MoveNext
90        Loop

100       FixComboWidth cmbUrineComment

110       Exit Sub

FillUrineComments_Error:

          Dim strES As String
          Dim intEL As Integer

120       intEL = Erl
130       strES = Err.Description
140       LogError "frmEditMicrobiologyNew", "FillUrineComments", intEL, strES, sql

End Sub

Private Function CheckTimes() As Boolean

          Dim strTime As String

          'returns true if ok
10        On Error GoTo CheckTimes_Error

20        CheckTimes = True

30        If HospName(0) <> "Monaghan" Then Exit Function

40        If Not IsDate(tSampleTime) Then
50            If InStr(txtDemographicComment, "Sample Time Unknown.") = 0 Then
60                If iMsg("Is Sample Time unknown?", vbQuestion + vbYesNo) = vbYes Then
70                    txtDemographicComment = txtDemographicComment & " Sample Time Unknown."
80                Else
90                    strTime = iTIME("Sample Time?")
100                   If IsDate(strTime) Then
110                       tSampleTime = strTime
120                   Else
130                       CheckTimes = False
140                       Exit Function
150                   End If
160               End If
170           End If
180       End If

190       If Not IsDate(tRecTime) Then
200           strTime = iTIME("Received Time?")
210           If IsDate(strTime) Then
220               tRecTime = strTime
230           Else
240               CheckTimes = False
250               Exit Function
260           End If
270       End If

280       Exit Function

CheckTimes_Error:

          Dim strES As String
          Dim intEL As Integer

290       intEL = Erl
300       strES = Err.Description
310       LogError "frmEditMicrobiologyNew", "CheckTimes", intEL, strES


End Function

Private Sub ClearIndividualFaeces()

          Dim n As Integer

10        On Error GoTo ClearIndividualFaeces_Error

20        For n = 0 To 2
30            lblFOB(n) = ""
40            lblFOB(n).BackColor = vbButtonFace
50        Next

60        txtRota = ""
70        txtRota.BackColor = vbButtonFace
80        txtAdeno = ""
90        txtAdeno.BackColor = vbButtonFace

100       lblToxinA = ""
110       lblToxinA.BackColor = vbButtonFace

120       lblCDiffCulture = ""
130       lblCDiffCulture.BackColor = vbButtonFace

140       lblGDH = ""
150       lblGDH.BackColor = vbButtonFace

160       cmbGDH = ""

170       lblPCR = ""
180       lblPCR.BackColor = vbButtonFace

190       cmbPCR = ""

200       lblCrypto = ""
210       lblCrypto.BackColor = vbButtonFace

220       lblGiardia = ""
230       lblGiardia.BackColor = vbButtonFace
240       For n = 0 To 2
250           cmbOva(n) = ""
260       Next

270       lblHPylori = ""
280       lblHPylori.BackColor = vbButtonFace

290       Exit Sub

ClearIndividualFaeces_Error:

          Dim strES As String
          Dim intEL As Integer

300       intEL = Erl
310       strES = Err.Description
320       LogError "frmEditMicrobiologyNew", "ClearIndividualFaeces", intEL, strES


End Sub

'Private Sub EnableCopyFrom()
'
'      Dim sql As String
'      Dim tb As Recordset
'      Dim PrevSID As Long
'
'10    On Error GoTo EnableCopyFrom_Error
'
'20    cmdCopyFromPrevious.Visible = False
'
'30    If sysOptAllowCopyDemographics(0) = False Then
'40      Exit Sub
'50    End If
'
'60    If Trim$(txtName) <> "" Or txtDoB <> "" Then
'70      Exit Sub
'80    End If
'
'90    PrevSID = SysOptMicroOffset(0) + Val(txtSampleID) - 1
'
'100   sql = "Select PatName from Demographics where " & _
 '            "SampleID = " & PrevSID & " " & _
 '            "and PatName <> '' " & _
 '            "and PatName is not null " & _
 '            "and DoB is not null"
'110   Set tb = New Recordset
'120   RecOpenServer 0, tb, sql
'130   If Not tb.EOF Then
'140     cmdCopyFromPrevious.Caption = "Copy All Details from Sample # " & _
 '                                      Format$(PrevSID - SysOptMicroOffset(0)) & _
 '                                      " Name " & tb!PatName
'150     cmdCopyFromPrevious.Visible = True
'160   End If
'
'170   Exit Sub
'
'EnableCopyFrom_Error:
'
'      Dim strES As String
'      Dim intEL As Integer
'
'180   intEL = Erl
'190   strES = Err.Description
'200   LogError "frmEditMicrobiologyNew", "EnableCopyFrom", intEL, strES, sql
'
'
'End Sub

Private Sub ClearIdent()

          Dim Index As Integer

10        On Error GoTo ClearIdent_Error

20        For Index = 1 To 4
30            cmbGram(Index) = ""
40            txtZN(Index) = ""
50            cmbWetPrep(Index) = ""
60            txtIndole(Index) = ""
70            txtCoagulase(Index) = ""
80            txtCatalase(Index) = ""
90            txtOxidase(Index) = ""
100           txtNotes(Index) = ""
110           txtReincubation(Index) = ""
120       Next

130       Exit Sub

ClearIdent_Error:

          Dim strES As String
          Dim intEL As Integer

140       intEL = Erl
150       strES = Err.Description
160       LogError "frmEditMicrobiologyNew", "ClearIdent", intEL, strES


End Sub

Private Sub ClearUrine()

10        On Error GoTo ClearUrine_Error

20        txtBacteria = ""

30        txtPregnancy = ""
40        txtHCGLevel = ""

50        txtWCC = "": txtWCC.BackColor = vbButtonFace: txtWCC.ForeColor = vbBlack
60        txtRCC = "": txtRCC.BackColor = vbButtonFace: txtRCC.ForeColor = vbBlack
70        cmbCrystals = ""
80        cmbCasts = ""
90        cmbMisc(0) = ""
100       cmbMisc(1) = ""
110       cmbMisc(2) = ""

120       Exit Sub

ClearUrine_Error:

          Dim strES As String
          Dim intEL As Integer

130       intEL = Erl
140       strES = Err.Description
150       LogError "frmEditMicrobiologyNew", "ClearUrine", intEL, strES

End Sub


Private Sub cmbSiteEffects()

          Dim f As Form

10        On Error GoTo cmbSiteEffects_Error

20        GetTabsFromSetUp

30        cmdOrderTests.Enabled = False

40        If InStr(1, cmbSite, "Faeces") > 0 Then

50            OrderFaeces
60            txtSiteDetails = ""

70        ElseIf cmbSite = "Urine" Then

80            Set f = frmMicroOrderUrine
90            f.txtSampleID = txtSampleID
100           f.Show 1
110           txtSiteDetails = f.SiteDetails
120           If f.chkUrine(2) Then
130               SSTab1.TabVisible(7) = True
140           End If
150           Unload f
160           Set f = Nothing

170       End If

180       lblSiteDetails = cmbSite & " " & txtSiteDetails

190       cmdSaveDemographics.Enabled = True
200       cmdSaveInc.Enabled = True

210       Exit Sub

cmbSiteEffects_Error:

          Dim strES As String
          Dim intEL As Integer

220       intEL = Erl
230       strES = Err.Description
240       LogError "frmEditMicrobiologyNew", "cmbSiteEffects", intEL, strES

End Sub

Private Sub FillABSelect(ByVal Index As Integer)

          Dim tb As Recordset
          Dim sql As String
          Dim n As Integer
          Dim ExcludeList As String
          Dim T As Single

10        On Error GoTo FillABSelect_Error

20        cmbABSelect(Index).Clear

30        ExcludeList = ""
40        For n = 1 To grdAB(Index).Rows - 1
50            ExcludeList = ExcludeList & _
                            "AntibioticName <> '" & LTrim(RTrim(grdAB(Index).TextMatrix(n, 0))) & "' and "
60        Next
70        ExcludeList = Left$(ExcludeList, Len(ExcludeList) - 4)

80        sql = "SELECT DISTINCT RTRIM(AntibioticName) AS AntibioticName, ListOrder " & _
                "FROM Antibiotics WHERE " & _
                ExcludeList & _
                "ORDER BY ListOrder"

90        Set tb = New Recordset
100       RecOpenServer 0, tb, sql
110       T = Timer
120       With cmbABSelect(Index)
130           Do While Not tb.EOF
140               .AddItem tb!AntibioticName & ""
150               tb.MoveNext
160           Loop
170       End With
180       Debug.Print Timer - T, "FillABSelectFillG"

190       Exit Sub

FillABSelect_Error:

          Dim strES As String
          Dim intEL As Integer

200       intEL = Erl
210       strES = Err.Description
220       LogError "frmEditMicrobiologyNew", "FillABSelect", intEL, strES, sql

End Sub

Private Sub FillCurrentABs()

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo FillCurrentABs_Error

20        cmbABsInUse.Clear

30        sql = "Select distinct AntibioticName, ListOrder " & _
                "from Antibiotics " & _
                "order by ListOrder"
40        Set tb = New Recordset
50        RecOpenClient 0, tb, sql
60        Do While Not tb.EOF
70            cmbABsInUse.AddItem Trim$(tb!AntibioticName & "")
80            tb.MoveNext
90        Loop

100       Exit Sub

FillCurrentABs_Error:

          Dim strES As String
          Dim intEL As Integer

110       intEL = Erl
120       strES = Err.Description
130       LogError "frmEditMicrobiologyNew", "FillCurrentABs", intEL, strES, sql


End Sub

Private Sub FillForConsultantValidation()

          Dim sql As String
          Dim tb As Recordset
          Dim SID As Double

10        On Error GoTo FillForConsultantValidation_Error

20        cmdAddToConsultantList.Caption = "Add to Consultant List"

30        cmbConsultantVal.Clear

40        sql = "Select * from ConsultantList " & _
                "Order by SampleID"
50        Set tb = New Recordset
60        RecOpenServer 0, tb, sql
70        Do While Not tb.EOF
80            SID = Val(tb!SampleID) - SysOptMicroOffset(0)
90            cmbConsultantVal.AddItem Format$(SID)
100           If SID = Val(txtSampleID) Then
110               cmdAddToConsultantList.Caption = "Remove from Consultant List"
120           End If
130           tb.MoveNext
140       Loop

150       Exit Sub

FillForConsultantValidation_Error:

          Dim strES As String
          Dim intEL As Integer

160       intEL = Erl
170       strES = Err.Description
180       LogError "frmEditMicrobiologyNew", "FillForConsultantValidation", intEL, strES, sql


End Sub

Private Sub FillHistoricalFaeces()

          Dim n As Integer
          Dim sql As String
          Dim tb As Recordset
          Dim strSelect As String
          Dim s As String
          Dim strPrevious(1 To 3) As String

10        On Error GoTo FillHistoricalFaeces_Error

20        For n = 1 To 3
30            grdDay(n).Rows = 2
40            grdDay(n).AddItem ""
50            grdDay(n).RemoveItem 1
60        Next

70        Select Case lblViewOrganism.Caption
          Case "1": strSelect = "Day111 as XLD, Day121 as DCA, Day131 as SMAC, " & _
                                "Day211 as XLDS, Day221 as CROMO2, Day231 as CAMP2,  Day241 as DCA2, " & _
                                "Day31 as CAMP3, Day34 as CROMO3"

80        Case "2": strSelect = "Day112 as XLD, Day122 as DCA, Day132 as SMAC, " & _
                                "Day212 as XLDS, Day222 as CROMO2, Day232 as CAMP2,  Day242 as DCA2, " & _
                                "Day32 as CAMP3, Day35 as CROMO3"

90        Case "3": strSelect = "Day113 as XLD, Day123 as DCA, Day133 as SMAC, " & _
                                "Day213 as XLDS, Day223 as CROMO2, Day233 as CAMP2,  Day243 as DCA2, " & _
                                "Day33 as CAMP3, Day36 as CROMO3"

              '  Case "2": strSelect = "Day112 as XLD, Day122 as DCA, Day132 as SMAC, " & _
                 '                        "Day212 as XLDS, Day222 as DCAS, Day232 as Preston, Day32 as CCDA,"
              '  Case "3": strSelect = "Day113 as XLD, Day123 as DCA, Day133 as SMAC, " & _
                 '                        "Day213 as XLDS, Day223 as DCAS, Day233 as Preston, Day33 as CCDA"
100       End Select

110       sql = "Select Operator, " & strSelect & " from FaecesWorksheet where " & _
                "SampleID = '" & Val(txtSampleID) + SysOptMicroOffset(0) & "'"
120       Set tb = New Recordset
130       RecOpenServer 0, tb, sql
140       If Not tb.EOF Then

              'Day 1
150           s = Trim$(tb!XLD & "") & vbTab & Trim$(tb!DCA & "") & vbTab & Trim$(tb!SMAC & "")
160           If Len(s) > 2 Then
170               strPrevious(1) = s
180               s = "Current" & vbTab & s & vbTab & tb!Operator & ""
190               grdDay(1).AddItem s
200           Else
210               strPrevious(1) = vbTab & vbTab
220           End If

              'Day 2
230           s = Trim$(tb!XLDS & "") & vbTab & Trim$(tb!DCA2 & "") & vbTab & Trim$(tb!CROMO2 & "") & vbTab & Trim$(tb!CAMP2 & "")
240           If Len(s) > 2 Then
250               strPrevious(2) = s
260               s = "Current" & vbTab & s & vbTab & tb!Operator & ""
270               grdDay(2).AddItem s
280           Else
290               strPrevious(2) = vbTab & vbTab
300           End If

              'Day 3
              'CAMP & CROMO
310           s = Trim$(tb!CAMP3 & "") & vbTab & Trim$(tb!CROMO3 & "")
320           strPrevious(3) = s
330           If Len(s) > 0 Then
340               s = "Current" & vbTab & s & vbTab & tb!Operator & ""
350               grdDay(3).AddItem s
360           End If

              ''
370           sql = "Select Operator, TimeOfRecord, " & strSelect & " from ArcFaecesWorksheet where " & _
                    "SampleID = '" & Val(txtSampleID) + SysOptMicroOffset(0) & "' " & _
                    "Order by TimeOfRecord desc"
380           Set tb = New Recordset
390           RecOpenServer 0, tb, sql
400           Do While Not tb.EOF
                  'Day 1
410               s = Trim$(tb!XLD & "") & vbTab & Trim$(tb!DCA & "") & vbTab & Trim$(tb!SMAC & "")
420               If Len(s) > 2 Then
430                   If s <> strPrevious(1) Then
440                       strPrevious(1) = s
450                       s = Format$(tb!TimeOfRecord, "dd/MM hh:mm") & vbTab & s & vbTab & tb!Operator & ""
460                       grdDay(1).AddItem s
470                   End If
480               Else
490                   strPrevious(1) = vbTab & vbTab
500               End If

                  'Day 2
510               s = Trim$(tb!XLDS & "") & vbTab & Trim$(tb!DCA2 & "") & vbTab & Trim$(tb!CROMO2 & "") & vbTab & Trim$(tb!CAMP2 & "")
520               If Len(s) > 2 Then
530                   If s <> strPrevious(2) Then
540                       strPrevious(2) = s
550                       s = Format$(tb!TimeOfRecord, "dd/MM hh:mm") & vbTab & s & vbTab & tb!Operator & ""
560                       grdDay(2).AddItem s
570                   End If
580               Else
590                   strPrevious(2) = vbTab & vbTab
600               End If

                  'Day 3
610               s = Trim$(tb!CAMP3 & "") & vbTab & Trim$(tb!CROMO3 & "")
620               If Len(s) > 0 Then
630                   If s <> strPrevious(3) Then
640                       strPrevious(3) = s
650                       s = Format$(tb!TimeOfRecord, "dd/MM hh:mm") & vbTab & s & vbTab & tb!Operator & ""
660                       grdDay(3).AddItem s
670                   End If
680               Else
690                   strPrevious(3) = ""
700               End If

710               tb.MoveNext

720           Loop
730       End If

740       For n = 1 To 3
750           If grdDay(n).Rows > 2 Then
760               grdDay(n).RemoveItem 1
770           End If
780       Next

790       Exit Sub

FillHistoricalFaeces_Error:

          Dim strES As String
          Dim intEL As Integer

800       intEL = Erl
810       strES = Err.Description
820       LogError "frmEditMicrobiologyNew", "FillHistoricalFaeces", intEL, strES, sql


End Sub

Private Sub FillMSandConsultantComment()

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo FillMSandConsultantComment_Error

20        cmbConC.Clear
30        cmbMSC.Clear

40        sql = "Select * from Lists where " & _
                "ListType = 'BA' " & _
                "ORDER BY ListOrder"
50        Set tb = New Recordset
60        RecOpenServer 0, tb, sql
70        Do While Not tb.EOF
80            cmbMSC.AddItem tb!Text & ""
90            cmbConC.AddItem tb!Text & ""
100           tb.MoveNext
110       Loop

120       Exit Sub

FillMSandConsultantComment_Error:

          Dim strES As String
          Dim intEL As Integer

130       intEL = Erl
140       strES = Err.Description
150       LogError "frmEditMicrobiologyNew", "FillMSandConsultantComment", intEL, strES, sql


End Sub

Private Sub FillOrgNames(ByVal Index As Integer)

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo FillOrgNames_Error

20        cmbOrgName(Index).Clear

30        If cmbOrgGroup(Index).Text = "Negative Results" Then
40            sql = "Select * from Organisms where " & _
                    "GroupName = '" & cmbOrgGroup(Index).Text & "' " & _
                    "AND Site = '" & cmbSite & "' " & _
                    "order by ListOrder"
50        Else
60            sql = "Select Distinct Name, ListOrder from Organisms where " & _
                    "GroupName = '" & cmbOrgGroup(Index).Text & "' " & _
                    "order by ListOrder"
70        End If
80        Set tb = New Recordset
90        RecOpenClient 0, tb, sql
100       Do While Not tb.EOF
110           cmbOrgName(Index).AddItem tb!Name & ""
120           tb.MoveNext
130       Loop

140       SetComboWidths

150       Exit Sub

FillOrgNames_Error:

          Dim strES As String
          Dim intEL As Integer

160       intEL = Erl
170       strES = Err.Description
180       LogError "frmEditMicrobiologyNew", "FillOrgNames", intEL, strES, sql


End Sub


Private Sub GetSampleIDWithOffset()

10        On Error GoTo GetSampleIDWithOffset_Error

20        SampleIDWithOffset = Val(txtSampleID) + SysOptMicroOffset(0)

30        Exit Sub

GetSampleIDWithOffset_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmEditMicrobiologyNew", "GetSampleIDWithOffset", intEL, strES

End Sub


Private Sub FillLists()

10        On Error GoTo FillLists_Error

20        FillGPsClinWard Me, HospName(0)

30        FillCastsCrystalsMiscSite
40        FillFaecesLists
50        FillDemographicComments
60        FillUrineComments
70        LoadListFluidAppearance
80        LoadListFluidCellCount
90        LoadListFluidGram
100       LoadListFluidLeishman
110       LoadListFluidZN
120       LoadListFluidWetPrep
130       LoadListFluidCrystals
140       LoadListQualifier
150       LoadListGeneric cmbGDH, "CDiffGDHDetail"
160       LoadListGeneric cmbPCR, "CDiffPCRDetail"


170       Exit Sub

FillLists_Error:

          Dim strES As String
          Dim intEL As Integer

180       intEL = Erl
190       strES = Err.Description
200       LogError "frmEditMicrobiologyNew", "FillLists", intEL, strES

End Sub



Private Function IsChild() As Boolean

10        On Error GoTo IsChild_Error

20        IsChild = False

30        If Not IsDate(txtDoB) Then Exit Function

40        If DateDiff("yyyy", txtDoB, Now) < 15 Then
50            IsChild = True
60        End If

70        Exit Function

IsChild_Error:

          Dim strES As String
          Dim intEL As Integer

80        intEL = Erl
90        strES = Err.Description
100       LogError "frmEditMicrobiologyNew", "IsChild", intEL, strES

End Function

Private Function IsPregnant() As Boolean

10        If chkPregnant = 1 Then
20            IsPregnant = True
30        Else
40            IsPregnant = False
50        End If

End Function

Private Function IsOutPatient() As Boolean

10        IsOutPatient = False

End Function

Private Sub LoadComments()

          Dim Ob As Observation
          Dim Obs As Observations

10        On Error GoTo LoadComments_Error

20        txtUrineComment = ""
30        txtDemographicComment = ""
40        txtMSC = "Medical Scientist Comments"
50        txtConC = "Consultant Comments"
60        txtFluidComment = ""
70        txtCDiffMSC = ""

80        If Trim$(txtSampleID) = "" Then Exit Sub

90        Set Obs = New Observations
100       Set Obs = Obs.Load(SampleIDWithOffset, _
                             "MicroGeneral", "Demographic", "MicroCS", _
                             "MicroConsultant", "CSFFluid", "MicroCDiff")
110       If Not Obs Is Nothing Then
120           For Each Ob In Obs
130               Select Case UCase$(Ob.Discipline)
                  Case "MICROGENERAL": txtUrineComment = Ob.Comment
140               Case "DEMOGRAPHIC": txtDemographicComment = Ob.Comment
150               Case "MICROCS": txtMSC = Ob.Comment
160               Case "MICROCONSULTANT": txtConC = Ob.Comment
170               Case "CSFFLUID": txtFluidComment = Ob.Comment
180               Case "MICROCDIFF": txtCDiffMSC = Ob.Comment
190               End Select
200           Next
210       End If
220       If txtMSC = "" Then
230           txtMSC = "Medical Scientist Comments"
240       End If

250       If txtConC = "" Then
260           txtConC = "Consultant Comments"
270       End If



280       Exit Sub

LoadComments_Error:

          Dim strES As String
          Dim intEL As Integer

290       intEL = Erl
300       strES = Err.Description
310       LogError "frmEditMicrobiologyNew", "LoadComments", intEL, strES

End Sub

Private Function LoadFaeces() As Boolean
      'Returns true if Faeces results present

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo LoadFaeces_Error

20        ClearFaeces

30        LoadFaeces = False

40        sql = "Select * from FaecesWorkSheet where " & _
                "SampleID = '" & SampleIDWithOffset & "'"
50        Set tb = New Recordset
60        RecOpenServer 0, tb, sql

70        If Not tb.EOF Then
80            LoadFaeces = True
              'Day 1
              'XLD
90            cmbDay1(11) = tb!Day111 & ""
100           cmbDay1(12) = tb!Day112 & ""
110           cmbDay1(13) = tb!Day113 & ""

              'DCA
120           cmbDay1(21) = tb!Day121 & ""
130           cmbDay1(22) = tb!Day122 & ""
140           cmbDay1(23) = tb!Day123 & ""
              'SMAC
150           cmbDay1(31) = tb!Day131 & ""
160           cmbDay1(32) = tb!Day132 & ""
170           cmbDay1(33) = tb!Day133 & ""
              'STEC
180           cmbDay1(41) = tb!Day141 & ""
190           cmbDay1(42) = tb!Day142 & ""
200           cmbDay1(43) = tb!Day143 & ""

              'Day2
              'XLD
210           cmbDay2(11) = tb!Day211 & ""
220           cmbDay2(12) = tb!Day212 & ""
230           cmbDay2(13) = tb!Day213 & ""
              'CROMO
240           cmbDay2(21) = tb!Day221 & ""
250           cmbDay2(22) = tb!Day222 & ""
260           cmbDay2(23) = tb!Day223 & ""
              'SMAC
270           cmbDay2(31) = tb!Day231 & ""
280           cmbDay2(32) = tb!Day232 & ""
290           cmbDay2(33) = tb!Day233 & ""
              'DCA
300           cmbDay2(41) = tb!Day241 & ""
310           cmbDay2(42) = tb!Day242 & ""
320           cmbDay2(43) = tb!Day243 & ""
              'STEC
330           cmbDay2(51) = tb!Day251 & ""
340           cmbDay2(52) = tb!Day252 & ""
350           cmbDay2(53) = tb!Day253 & ""

              'Day 3
              'CAMP
360           cmbDay3(1) = tb!Day31 & ""
370           cmbDay3(2) = tb!Day32 & ""
380           cmbDay3(3) = tb!Day33 & ""
              'CROMO
390           cmbDay3(4) = tb!Day34 & ""
400           cmbDay3(5) = tb!Day35 & ""
410           cmbDay3(6) = tb!Day36 & ""

420       End If

430       Exit Function

LoadFaeces_Error:

          Dim strES As String
          Dim intEL As Integer

440       intEL = Erl
450       strES = Err.Description
460       LogError "frmEditMicrobiologyNew", "LoadFaeces", intEL, strES, sql


End Function

Private Sub LoadListBacteria()

          Dim sql As String
          Dim tb As Recordset

10        On Error GoTo LoadListBacteria_Error

20        ReDim ListBacteria(0 To 0) As String
30        ListBacteria(0) = ""

40        sql = "SELECT * FROM Lists WHERE " & _
                "ListType = 'BB' " & _
                "ORDER BY ListOrder"
50        Set tb = New Recordset
60        RecOpenServer 0, tb, sql
70        Do While Not tb.EOF
80            ReDim Preserve ListBacteria(0 To UBound(ListBacteria) + 1)
90            ListBacteria(UBound(ListBacteria)) = tb!Text & ""
100           tb.MoveNext
110       Loop

120       Exit Sub

LoadListBacteria_Error:

          Dim strES As String
          Dim intEL As Integer

130       intEL = Erl
140       strES = Err.Description
150       LogError "frmEditMicrobiologyNew", "LoadListBacteria", intEL, strES, sql


End Sub
Private Sub LoadListPregnancy()

          Dim sql As String
          Dim tb As Recordset

10        On Error GoTo LoadListPregnancy_Error

20        ReDim ListPregnancy(0 To 0) As String
30        ListPregnancy(0) = ""

40        sql = "SELECT * FROM Lists WHERE " & _
                "ListType = 'PG' " & _
                "ORDER BY ListOrder"
50        Set tb = New Recordset
60        RecOpenServer 0, tb, sql
70        Do While Not tb.EOF
80            ReDim Preserve ListPregnancy(0 To UBound(ListPregnancy) + 1)
90            ListPregnancy(UBound(ListPregnancy)) = tb!Text & ""
100           tb.MoveNext
110       Loop

120       Exit Sub

LoadListPregnancy_Error:

          Dim strES As String
          Dim intEL As Integer

130       intEL = Erl
140       strES = Err.Description
150       LogError "frmEditMicrobiologyNew", "LoadListPregnancy", intEL, strES, sql


End Sub


Private Sub LoadListFluidCellCount()

          Dim sql As String
          Dim tb As Recordset

10        On Error GoTo LoadListFluidCellCount_Error

20        cmbFluidAppearance(0).Clear

30        sql = "SELECT * FROM Lists WHERE " & _
                "ListType = 'CC' " & _
                "ORDER BY ListOrder"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        Do While Not tb.EOF
70            cmbFluidAppearance(0).AddItem Trim(tb!Text & "")
80            tb.MoveNext
90        Loop

100       FixComboWidth cmbFluidAppearance(0)

110       Exit Sub

LoadListFluidCellCount_Error:

          Dim strES As String
          Dim intEL As Integer

120       intEL = Erl
130       strES = Err.Description
140       LogError "frmEditMicrobiologyNew", "LoadListFluidCellCount", intEL, strES, sql

End Sub

Private Sub LoadListFluidAppearance()

          Dim sql As String
          Dim tb As Recordset

10        On Error GoTo LoadListFluidAppearance_Error

20        cmbFluidAppearance(1).Clear

30        sql = "SELECT * FROM Lists WHERE " & _
                "ListType = 'FA' " & _
                "ORDER BY ListOrder"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        Do While Not tb.EOF
70            cmbFluidAppearance(1).AddItem Trim(tb!Text & "")
80            tb.MoveNext
90        Loop

100       FixComboWidth cmbFluidAppearance(1)

110       Exit Sub

LoadListFluidAppearance_Error:

          Dim strES As String
          Dim intEL As Integer

120       intEL = Erl
130       strES = Err.Description
140       LogError "frmEditMicrobiologyNew", "LoadListFluidAppearance", intEL, strES, sql

End Sub

Private Function LoadLockStatus(ByVal Index As Integer) As Boolean

'Returns True if locked

    Dim tb As Recordset
    Dim sql As String
    Dim RetVal As Boolean
    Dim i As Integer

10  On Error GoTo LoadLockStatus_Error

20  RetVal = False
30  LoadLockStatus = False

40  If Index = 0 Or Index = 3 Or Index = 12 Then Exit Function

50  sql = "Select * from LockStatus where " & _
          "SampleID = '" & SampleIDWithOffset & "' " & _
          "AND DeptIndex = '" & Index & "'"
60  Set tb = New Recordset
70  RecOpenServer 0, tb, sql

80  cmdLock(Index).Caption = "&Lock Result"
90  cmdLock(Index).Picture = frmMain.ImageList2.ListImages("Key").Picture

100 If Not tb.EOF Then
110     If tb!Lock Then
120         cmdLock(Index).Caption = "Un&Lock Result"
130         cmdLock(Index).Picture = frmMain.ImageList2.ListImages("Locked").Picture
140         RetVal = True
150     End If
160 End If

170 Select Case Index
    Case 2:
180     For i = 1 To 4
190         FrameExtras(i).Enabled = Not RetVal
200     Next
210 Case 7: fraRedSub.Enabled = Not RetVal
220 Case 8: fraRSV.Enabled = Not RetVal
230 Case 11: fraOP.Enabled = Not RetVal
240 Case 10: fraCDiff.Enabled = Not RetVal
250 Case 6: fraRotaAdeno.Enabled = Not RetVal
260 Case 5: fraFOB.Enabled = Not RetVal
270 Case 1: fraMicroscopy.Enabled = Not RetVal
280     fraPregnancy.Enabled = Not RetVal
290     txtUrineComment.Enabled = Not RetVal
300 Case 4: fraCS.Enabled = Not RetVal
310 Case 13: fraHPylori.Enabled = Not RetVal
320 End Select

330 LoadLockStatus = RetVal

340 Exit Function

LoadLockStatus_Error:

    Dim strES As String
    Dim intEL As Integer

350 intEL = Erl
360 strES = Err.Description
370 LogError "frmEditMicrobiologyNew", "LoadLockStatus", intEL, strES, sql

End Function

Private Sub LockFraCS(ByVal LockIt As Boolean)

10        cmdLock(1).Visible = Not LockIt
20        cmdLock(4).Visible = Not LockIt
30        cmdLock(5).Visible = Not LockIt
40        cmdLock(6).Visible = Not LockIt
50        cmdLock(7).Visible = Not LockIt
60        cmdLock(8).Visible = Not LockIt
70        cmdLock(9).Visible = Not LockIt
80        cmdLock(10).Visible = Not LockIt
90        cmdLock(11).Visible = Not LockIt

100       fraCS.Enabled = Not LockIt

End Sub

Private Sub MoveCursorToSaveButton()

          Dim T As Single

10        On Error GoTo MoveCursorToSaveButton_Error

20        T = Timer

30        SetCursorPos (cmdSaveMicro.Left + (cmdSaveMicro.Width / 2)) / Screen.TwipsPerPixelX, _
                       (cmdSaveMicro.Top + cmdSaveMicro.Height) / Screen.TwipsPerPixelY

40        cmdSaveMicro.BackColor = vbYellow

50        Do While Timer - T < 0.5
60            DoEvents
70        Loop

80        cmdSaveMicro.BackColor = vbButtonFace

90        Exit Sub

MoveCursorToSaveButton_Error:

          Dim strES As String
          Dim intEL As Integer

100       intEL = Erl
110       strES = Err.Description
120       LogError "frmEditMicrobiologyNew", "MoveCursorToSaveButton", intEL, strES

End Sub

Private Function QueryCEF() As Boolean

          Dim grd As Integer
          Dim Y As Integer
          Dim s As String
          Dim FoundSens As Boolean
          Dim FoundCEF As Boolean
          Dim FoundResults As Boolean

10        On Error GoTo QueryCEF_Error

20        QueryCEF = False

30        If UCase(cmbSite) = "URINE" Then
40            FoundSens = False
50            FoundCEF = False
60            FoundResults = False

70            For grd = 1 To 4
80                If grdAB(grd).TextMatrix(1, 0) <> "" Then
90                    FoundResults = True
100                   For Y = 1 To grdAB(grd).Rows - 1
110                       grdAB(grd).Col = 0
120                       grdAB(grd).Row = Y
130                       If grdAB(grd).Font.Bold = False Then
140                           If grdAB(grd).CellBackColor = 0 Then
150                               If grdAB(grd).TextMatrix(Y, 1) = "S" Then
160                                   FoundSens = True
170                                   Exit For
180                               End If
190                           End If
200                       End If
210                   Next
220                   If Not FoundSens Then
230                       For Y = 1 To grdAB(grd).Rows - 1
240                           If grdAB(grd).TextMatrix(Y, 0) = "Cefuroxime" Then
250                               grdAB(grd).Col = 2
260                               grdAB(grd).Row = Y
270                               If grdAB(grd).CellPicture = imgSquareTick.Picture Then
280                                   FoundCEF = True
290                                   Exit For
300                               End If
310                           End If
320                       Next
330                       If FoundCEF Then
340                           Exit For
350                       End If
360                   End If
370               End If
380           Next
              '380     If FoundResults And (Not FoundSens) And (Not FoundCEF) Then
390           If FoundResults And (Not FoundSens) And FoundCEF Then
400               s = "No First line Antibiotics are Sensitive!" & vbCrLf & _
                      "Do you wish to report Cefuroxime?"
410               If iMsg(s, vbQuestion + vbYesNo) = vbYes Then
420                   QueryCEF = True
430               End If
440           End If

450       End If

460       Exit Function

QueryCEF_Error:

          Dim strES As String
          Dim intEL As Integer

470       intEL = Erl
480       strES = Err.Description
490       LogError "frmEditMicrobiologyNew", "QueryCEF", intEL, strES

End Function

Private Function QueryGent() As Boolean

          Dim grd As Integer
          Dim Y As Integer
          Dim s As String
          Dim Reported As Boolean
          Dim FoundResults As Boolean

10        On Error GoTo QueryGent_Error

20        QueryGent = False
30        FoundResults = False
40        If UCase(cmbWard) = "PAEDIATRICS" Or UCase$(cmbWard) = "SPECIAL CARE BABY UNIT" Then
50            Reported = False
60            For grd = 1 To 4
70                If grdAB(grd).TextMatrix(1, 0) <> "" Then
80                    FoundResults = True
90                    For Y = 1 To grdAB(grd).Rows - 1
100                       If grdAB(grd).TextMatrix(Y, 0) = "Gentamicin" Then
110                           grdAB(grd).Row = Y
120                           grdAB(grd).Col = 2
130                           If grdAB(grd).CellPicture = imgSquareTick.Picture Then
140                               Reported = True
150                               Exit For
160                           End If
170                       End If
180                   Next
190                   If Reported Then
200                       Exit For
210                   End If
220               End If
230           Next
240           If FoundResults And Not Reported Then
250               s = "This Isolate is from a Patient" & vbCrLf & _
                      "in " & cmbWard & "." & vbCrLf & _
                      "Do you wish to report Gentamicin?"
260               If iMsg(s, vbQuestion + vbYesNo) = vbYes Then
270                   QueryGent = True
280               End If
290           End If
300       End If

310       Exit Function

QueryGent_Error:

          Dim strES As String
          Dim intEL As Integer

320       intEL = Erl
330       strES = Err.Description
340       LogError "frmEditMicrobiologyNew", "QueryGent", intEL, strES

End Function

 Sub SaveMicro()



10        On Error GoTo SaveMicro_Error


20        pBar = 0

30        GetSampleIDWithOffset

40        SaveDemographics

50        Select Case SSTab1.Tab
          Case 1: SaveUrine
60        Case 2: SaveIdent
70        Case 3: SaveFaeces
80            FillHistoricalFaeces
90        Case 4: ApplyExclusionABRule
100           SaveIsolates
110           SaveSensitivities gNOCHANGE
120           AdjustOrganism
130       Case 5: SaveFOB
140       Case 6: SaveRotaAdeno
150       Case 7: SaveRedSub
160       Case 8: SaveRSV
170       Case 9: SaveFluids
180       Case 10: SaveCdiff
190       Case 11: SaveOP
200       Case 13: SaveHPylori
210       End Select

220       SaveComments
230       UPDATEMRU txtSampleID, cMRU

240       cmdSaveMicro.Enabled = False
250       cmdSaveHold.Enabled = False
260       cmbSite.Enabled = False

270       Exit Sub

SaveMicro_Error:

          Dim strES As String
          Dim intEL As Integer

280       intEL = Erl
290       strES = Err.Description
300       LogError "frmEditMicrobiologyNew", "SaveMicro", intEL, strES

End Sub

Private Sub SaveFaecalTabs()

10        pBar = 0

20        GetSampleIDWithOffset

30        SaveFOB
40        SaveRotaAdeno
50        SaveCdiff
60        SaveOP
70        SaveHPylori
80        SaveRSV

90        cmdSaveMicro.Enabled = False
100       cmdSaveHold.Enabled = False

End Sub

Private Sub ShowPrintValidFlags()

          Dim sql As String
          Dim tb As Recordset

10        On Error GoTo ShowPrintValidFlags_Error

20        GetSampleIDWithOffset

30        lblValid(1).Visible = False
40        lblValid(4).Visible = False
50        lblValid(5).Visible = False
60        lblValid(6).Visible = False
70        lblValid(7).Visible = False
80        lblValid(8).Visible = False
90        lblValid(10).Visible = False
100       lblValid(11).Visible = False
110       lblValid(13).Visible = False

120       fraHPylori.Enabled = True
130       fraCS.Enabled = True
140       fraMicroscopy.Enabled = True
150       fraPregnancy.Enabled = True
160       cmbUrineComment.Enabled = True
170       txtUrineComment.Enabled = True

180       sql = "SELECT * FROM PrintValidLog WHERE " & _
                "SampleID = '" & SampleIDWithOffset & "'"
190       Set tb = New Recordset
200       RecOpenServer 0, tb, sql
210       Do While Not tb.EOF
220           Select Case tb!Department & ""
              Case "U"
230               If tb!Valid = 1 Then
240                   lblValid(1).Visible = True
250                   fraMicroscopy.Enabled = False
260                   fraPregnancy.Enabled = False
270                   cmbUrineComment.Enabled = False
280                   txtUrineComment.Enabled = False
290               End If
300               If tb!Printed = 1 Then
310                   lblPrinted(1).Visible = True
320               End If

330           Case "D"
340               If tb!Valid = 1 Then
350                   lblValid(4).Visible = True
360                   fraCS.Enabled = False
370               End If
380               If tb!Printed = 1 Then
390                   lblPrinted(4).Visible = True
400               End If

410           Case "F"
420               If tb!Valid = 1 Then
430                   lblValid(5).Visible = True
440               End If
450               If tb!Printed = 1 Then
460                   lblPrinted(5).Visible = True
470               End If

480           Case "A"
490               If tb!Valid = 1 Then
500                   lblValid(6).Visible = True
510               End If
520               If tb!Printed = 1 Then
530                   lblPrinted(6).Visible = True
540               End If

550           Case "R"
560               If tb!Valid = 1 Then
570                   lblValid(7).Visible = True
580               End If
590               If tb!Printed = 1 Then
600                   lblPrinted(7).Visible = True
610               End If

620           Case "V"
630               If tb!Valid = 1 Then
640                   lblValid(8).Visible = True
650               End If
660               If tb!Printed = 1 Then
670                   lblPrinted(8).Visible = True
680               End If

690           Case "G"
700               If tb!Valid = 1 Then
710                   lblValid(10).Visible = True
720               End If
730               If tb!Printed = 1 Then
740                   lblPrinted(10).Visible = True
750               End If

760           Case "O"
770               If tb!Valid = 1 Then
780                   lblValid(11).Visible = True
790               End If
800               If tb!Printed = 1 Then
810                   lblPrinted(11).Visible = True
820               End If

830           Case "Y"
840               If tb!Valid = 1 Then
850                   fraHPylori.Enabled = False
860                   lblValid(13).Visible = True
870               End If
880               If tb!Printed = 1 Then
890                   lblPrinted(13).Visible = True
900               End If

910           End Select
920           tb.MoveNext
930       Loop

940       Exit Sub

ShowPrintValidFlags_Error:

          Dim strES As String
          Dim intEL As Integer

950       intEL = Erl
960       strES = Err.Description
970       LogError "frmEditMicrobiologyNew", "ShowPrintValidFlags", intEL, strES, sql

End Sub
Private Function IsCSValid() As Boolean

          Dim sql As String
          Dim tb As Recordset

10        On Error GoTo IsCSValid_Error

20        GetSampleIDWithOffset

30        IsCSValid = False

40        sql = "SELECT * FROM PrintValidLog WHERE " & _
                "SampleID = '" & SampleIDWithOffset & "' " & _
                "AND Department = 'D'"
50        Set tb = New Recordset
60        RecOpenServer 0, tb, sql
70        If tb.EOF Then
80            IsCSValid = True
90        Else
100           If tb!Valid = 1 Then
110               IsCSValid = True
120           End If
130       End If

140       Exit Function

IsCSValid_Error:

          Dim strES As String
          Dim intEL As Integer

150       intEL = Erl
160       strES = Err.Description
170       LogError "frmEditMicrobiologyNew", "IsCSValid", intEL, strES, sql

End Function

Private Sub LoadPrintValid(ByVal Dept As String, ByRef v As String, ByRef P As String)

          Dim sql As String
          Dim tb As Recordset

10        On Error GoTo LoadPrintValid_Error

20        GetSampleIDWithOffset

30        v = ""
40        P = ""

50        sql = "SELECT COALESCE(Printed, 0) AS Printed, COALESCE(Valid, 0) AS Valid " & _
                "FROM PrintValidLog WHERE " & _
                "SampleID = '" & SampleIDWithOffset & "' " & _
                "AND Department = '" & Dept & "'"
60        Set tb = New Recordset
70        RecOpenServer 0, tb, sql
80        If Not tb.EOF Then
90            If tb!Valid Then v = "V"
100           If tb!Printed Then P = "P"
110       End If

120       Exit Sub

LoadPrintValid_Error:

          Dim strES As String
          Dim intEL As Integer

130       intEL = Erl
140       strES = Err.Description
150       LogError "frmEditMicrobiologyNew", "LoadPrintValid", intEL, strES, sql

End Sub

Private Sub ShowUnlock(ByVal Index As Integer)

10        On Error GoTo ShowUnlock_Error

20        cmdSaveMicro.Enabled = True
30        cmdSaveHold.Enabled = True
40        cmdLock(Index).Visible = True
50        cmdLock(Index).Caption = "&Lock Result"
60        cmdLock(Index).Picture = frmMain.ImageList2.ListImages("Key").Picture

70        Exit Sub

ShowUnlock_Error:

          Dim strES As String
          Dim intEL As Integer

80        intEL = Erl
90        strES = Err.Description
100       LogError "frmEditMicrobiologyNew", "ShowUnlock", intEL, strES

End Sub

Private Sub UpdateLockStatus(ByVal SampleID As Double, _
                             ByVal LockIt As Boolean, _
                             ByVal Index As Integer)

          Dim tb As Recordset
          Dim sql As String
          Dim i As Integer

10        On Error GoTo UpdateLockStatus_Error

20        sql = "Select * from LockStatus where " & _
                "SampleID = " & SampleIDWithOffset & " " & _
                "AND DeptIndex = '" & Index & "'"
30        Set tb = New Recordset
40        RecOpenClient 0, tb, sql
50        If tb.EOF Then
60            tb.AddNew
70        End If
80        tb!SampleID = SampleID
90        tb!DeptIndex = Index
100       tb!Lock = LockIt
110       tb.Update

120       If LockIt Then
130           cmdLock(Index).Caption = "Un&Lock Result"
140           cmdLock(Index).Picture = frmMain.ImageList2.ListImages("Locked").Picture
150       Else
160           cmdLock(Index).Caption = "&Lock Result"
170           cmdLock(Index).Picture = frmMain.ImageList2.ListImages("Key").Picture
180       End If

190       Select Case Index
          Case 1: fraMicroscopy.Enabled = Not LockIt
200           fraPregnancy.Enabled = Not LockIt
210           txtUrineComment.Enabled = Not LockIt
220       Case 2:
230           For i = 1 To 4
240               FrameExtras(i).Enabled = Not LockIt
250           Next
260       Case 4: fraCS.Enabled = Not LockIt
270       Case 5: fraFOB.Enabled = Not LockIt
280       Case 6: fraRotaAdeno.Enabled = Not LockIt
290       Case 7: fraRedSub.Enabled = Not LockIt
300       Case 8: fraRSV.Enabled = Not LockIt
310       Case 9: fraCSF.Enabled = Not LockIt
320       Case 11: fraOP.Enabled = Not LockIt
330       Case 10: fraCDiff.Enabled = Not LockIt
340       Case 13: fraHPylori.Enabled = Not LockIt
350       End Select

360       Exit Sub

UpdateLockStatus_Error:

          Dim strES As String
          Dim intEL As Integer

370       intEL = Erl
380       strES = Err.Description
390       LogError "frmEditMicrobiologyNew", "UpdateLockStatus", intEL, strES, sql

End Sub
Private Function LoadOP() As Boolean
      'Returns true if OP results present

          Dim tb As Recordset
          Dim sql As String
          Dim n As Integer
          Dim Found As Boolean
          Dim R() As String

10        On Error GoTo LoadOP_Error

20        cmdLock(11).Visible = False
30        fraOP.Enabled = True
40        lblCrypto = ""
50        lblCrypto.BackColor = &H8000000F
60        lblGiardia = ""
70        lblGiardia.BackColor = &H8000000F
80        For n = 0 To 2
90            cmbOva(n) = ""
100       Next

110       Found = False

120       sql = "Select * from Faeces where " & _
                "SampleID = '" & SampleIDWithOffset & "'"
130       Set tb = New Recordset
140       RecOpenServer 0, tb, sql

150       If Not tb.EOF Then
160           R = Split(tb!Cryptosporidium & "", "|")
170           If UBound(R) = -1 Then
180               lblCrypto = ""
190           ElseIf UBound(R) > 1 Then
200               lblCrypto = R(0)
210               lblCrypto.ForeColor = R(1)
220               lblCrypto.BackColor = R(2)
230               Found = True
240           Else
250               lblCrypto = R(0)
260               Found = True
270           End If

280           R = Split(tb!GiardiaLambila & "", "|")
290           If UBound(R) = -1 Then
300               lblGiardia = ""
310           ElseIf UBound(R) > 1 Then
320               lblGiardia = R(0)
330               lblGiardia.ForeColor = R(1)
340               lblGiardia.BackColor = R(2)
350               Found = True
360           Else
370               lblGiardia = R(0)
380               Found = True
390           End If
400           For n = 0 To 2
410               cmbOva(n) = Trim$(tb("OP" & Format(n)) & "")
420               If cmbOva(n) <> "" Then
430                   Found = True
440               End If
450           Next

460       End If

470       If Found Then
480           cmdLock(11).Visible = True
490           If LoadLockStatus(11) Then
500               fraOP.Enabled = False
510           End If
520           LoadOP = True
530       End If

540       Exit Function

LoadOP_Error:

          Dim strES As String
          Dim intEL As Integer

550       intEL = Erl
560       strES = Err.Description
570       LogError "frmEditMicrobiologyNew", "LoadOP", intEL, strES, sql

End Function


Private Function LoadCDiff() As Boolean
      'Returns true if Cdiff results present

          Dim tb As Recordset
          Dim sql As String
          Dim Found As Boolean
          Dim c() As String
          Dim T() As String
          Dim g() As String
          Dim P() As String

10        On Error GoTo LoadCDiff_Error

20        Found = False

30        cmdLock(10).Visible = False
40        fraCDiff.Enabled = True
50        lblToxinA = ""
60        lblToxinA.BackColor = vbButtonFace
70        lblCDiffCulture = ""
80        lblCDiffCulture.BackColor = vbButtonFace
90        lblGDH = ""
100       lblGDH.BackColor = vbButtonFace
110       cmbGDH = ""
120       lblPCR = ""
130       lblPCR.BackColor = vbButtonFace
140       cmbPCR = ""

150       LoadCDiff = False

160       sql = "SELECT ToxinAB, CDiffCulture, GDH, PCR, COALESCE(GDHDetail,'') GDHDetail, COALESCE(PCRDetail,'') PCRDetail FROM Faeces WHERE " & _
                "SampleID = '" & SampleIDWithOffset & "' " & _
                "AND (    LTRIM(RTRIM(ToxinAB)) <> '' " & _
                "OR      LTRIM(RTRIM(GDH)) <> '' " & _
                "OR      LTRIM(RTRIM(PCR)) <> '' " & _
                "      OR LTRIM(RTRIM(CDiffCulture)) <> '' )"

170       Set tb = New Recordset
180       RecOpenServer 0, tb, sql

190       If Not tb.EOF Then
200           T = Split(tb!ToxinAB & "", "|")
210           If UBound(T) = -1 Then
220               lblToxinA = ""
230           ElseIf UBound(T) > 1 Then
240               lblToxinA = T(0)
250               lblToxinA.ForeColor = T(1)
260               lblToxinA.BackColor = T(2)
270           Else
280               lblToxinA = T(0)
290           End If

300           c = Split(tb!CDiffCulture & "", "|")
310           If UBound(c) = -1 Then
320               lblCDiffCulture = ""
330           ElseIf UBound(c) > 1 Then
340               lblCDiffCulture = c(0)
350               lblCDiffCulture.ForeColor = c(1)
360               lblCDiffCulture.BackColor = c(2)
370           Else
380               lblCDiffCulture = c(0)
390           End If

400           g = Split(tb!GDH & "", "|")
410           If UBound(g) = -1 Then
420               lblGDH = ""
430           ElseIf UBound(g) > 1 Then
440               lblGDH = g(0)
450               lblGDH.ForeColor = g(1)
460               lblGDH.BackColor = g(2)
470           Else
480               lblGDH = g(0)
490           End If
500           cmbGDH = tb!GDHDetail

510           P = Split(tb!PCR & "", "|")
520           If UBound(P) = -1 Then
530               lblPCR = ""
540           ElseIf UBound(P) > 1 Then
550               lblPCR = P(0)
560               lblPCR.ForeColor = P(1)
570               lblPCR.BackColor = P(2)
580           Else
590               lblPCR = P(0)
600           End If

610           cmbPCR = tb!PCRDetail

620           cmdLock(10).Visible = True
630           If LoadLockStatus(10) Then
640               fraCDiff.Enabled = False
650           End If
660           LoadCDiff = True
670       End If

680       Exit Function

LoadCDiff_Error:

          Dim strES As String
          Dim intEL As Integer

690       intEL = Erl
700       strES = Err.Description
710       LogError "frmEditMicrobiologyNew", "LoadCDiff", intEL, strES, sql

End Function


Private Function LoadRotaAdeno() As Boolean
      'Returns true if Rota/Adeno results present

          Dim tb As Recordset
          Dim sql As String
          Dim Found As Boolean
          Dim R() As String
          Dim A() As String

10        On Error GoTo LoadRotaAdeno_Error

20        Found = False

30        cmdLock(6).Visible = False
40        fraRotaAdeno.Enabled = True
50        txtRota = ""
60        txtRota.BackColor = &H8000000F
70        txtAdeno = ""
80        txtAdeno.BackColor = &H8000000F

90        LoadRotaAdeno = False

100       sql = "Select Rota, Adeno from Faeces where " & _
                "SampleID = '" & SampleIDWithOffset & "'"
110       Set tb = New Recordset
120       RecOpenServer 0, tb, sql

130       If Not tb.EOF Then
140           R = Split(tb!Rota & "", "|")
150           If UBound(R) = -1 Then
160               txtRota = ""
170           ElseIf UBound(R) > 1 Then
180               txtRota = R(0)
190               txtRota.ForeColor = R(1)
200               txtRota.BackColor = R(2)
210               Found = True
220           Else
230               txtRota = R(0)
240               Found = True
250           End If

260           A = Split(tb!Adeno & "", "|")
270           If UBound(A) = -1 Then
280               txtAdeno = ""
290           ElseIf UBound(A) > 1 Then
300               txtAdeno = A(0)
310               txtAdeno.ForeColor = A(1)
320               txtAdeno.BackColor = A(2)
330               Found = True
340           Else
350               txtAdeno = A(0)
360               Found = True
370           End If

380       End If

390       If Found Then
400           cmdLock(6).Visible = True
410           If LoadLockStatus(6) Then
420               fraRotaAdeno.Enabled = False
430           End If
440           LoadRotaAdeno = True
450       End If

460       Exit Function

LoadRotaAdeno_Error:

          Dim strES As String
          Dim intEL As Integer

470       intEL = Erl
480       strES = Err.Description
490       LogError "frmEditMicrobiologyNew", "LoadRotaAdeno", intEL, strES, sql

End Function

Private Function LoadRedSub() As Boolean
      'Returns true if Reducing Substances results present

          Dim tb As Recordset
          Dim sql As String
          Dim n As Integer

10        On Error GoTo LoadRedSub_Error

20        cmdLock(7).Visible = False
30        fraRedSub.Enabled = True
40        For n = 0 To 5
50            chkRS(n).Value = 0
60        Next
70        LoadRedSub = False

80        sql = "Select * from GenericResults where " & _
                "SampleID = '" & SampleIDWithOffset & "' " & _
                "and TestName = 'RedSub'"
90        Set tb = New Recordset
100       RecOpenServer 0, tb, sql

110       If Not tb.EOF Then

120           For n = 0 To 5
130               If chkRS(n).Caption = tb!Result Then
140                   chkRS(n).Value = 1
150                   Exit For
160               End If
170           Next

180           cmdLock(7).Visible = True
190           If LoadLockStatus(7) Then
200               fraRedSub.Enabled = False
210           End If

220           LoadRedSub = True

230       End If

240       Exit Function

LoadRedSub_Error:

          Dim strES As String
          Dim intEL As Integer

250       intEL = Erl
260       strES = Err.Description
270       LogError "frmEditMicrobiologyNew", "LoadRedSub", intEL, strES, sql


End Function

Private Function LoadRSV() As Boolean
      'Returns true if RSV results present

          Dim tb As Recordset
          Dim sql As String
          Dim R() As String

10        On Error GoTo LoadRSV_Error

20        cmdLock(8).Visible = False
30        fraRSV.Enabled = True
40        lblRSV.Caption = ""
50        lblRSV.BackColor = &H8000000F

60        LoadRSV = False

70        sql = "Select * from GenericResults where " & _
                "SampleID = '" & SampleIDWithOffset & "' " & _
                "and TestName = 'RSV'"
80        Set tb = New Recordset
90        RecOpenServer 0, tb, sql

100       If Not tb.EOF Then
110           R = Split(tb!Result & "", "|")
120           If UBound(R) = -1 Then
130               lblRSV = ""
140           ElseIf UBound(R) > 1 Then
150               lblRSV = R(0)
160               lblRSV.ForeColor = R(1)
170               lblRSV.BackColor = R(2)
180           Else
190               lblRSV = R(0)
200           End If

210           cmdLock(8).Visible = True
220           If LoadLockStatus(8) Then
230               fraRSV.Enabled = False
240           End If

250           LoadRSV = True

260       End If

270       Exit Function

LoadRSV_Error:

          Dim strES As String
          Dim intEL As Integer

280       intEL = Erl
290       strES = Err.Description
300       LogError "frmEditMicrobiologyNew", "LoadRSV", intEL, strES, sql


End Function

Private Function LoadHPylori() As Boolean
      'LoadHPylori = LoadFaecesHPylori()
      'Exit Function
      'Returns true if HPylori results present

          Dim tb As Recordset
          Dim sql As String
          Dim h() As String

10        On Error GoTo LoadHPylori_Error

20        cmdLock(13).Visible = False
30        fraHPylori.Enabled = True
40        lblHPylori.Caption = ""
50        lblHPylori.BackColor = &H8000000F

60        LoadHPylori = False
          'lblEnteredBy.Visible = False

70        sql = "SELECT HPylori, UserName FROM Faeces WHERE " & _
                "SampleID = '" & SampleIDWithOffset & "' and COALESCE(HPylori, '') <> ''"
80        Set tb = New Recordset
90        RecOpenServer 0, tb, sql

100       If Not tb.EOF Then

              'If Trim$(tb!UserName & "") <> "" Then
              '  lblEnteredBy.Caption = "Entered By" & vbCrLf & tb!UserName
              '  lblEnteredBy.Visible = True
              'End If
110           h = Split(tb!HPylori & "", "|")
120           If UBound(h) = -1 Then
130               lblHPylori = ""
140           ElseIf UBound(h) > 1 Then
150               lblHPylori = h(0)
160               lblHPylori.ForeColor = h(1)
170               lblHPylori.BackColor = h(2)
180               LoadHPylori = True
190           Else
200               lblHPylori = h(0)
210           End If

220           cmdLock(13).Visible = True
230           If LoadLockStatus(13) Then
240               fraHPylori.Enabled = False
250           End If

260       End If

270       Exit Function

LoadHPylori_Error:

          Dim strES As String
          Dim intEL As Integer

280       intEL = Erl
290       strES = Err.Description
300       LogError "frmEditMicrobiologyNew", "LoadHPylori", intEL, strES, sql

End Function


Private Function LoadFOB() As Boolean
      'Returns true if FOB results present

          Dim tb As Recordset
          Dim sql As String
          Dim n As Integer
          Dim Found As Boolean
          Dim FOB() As String

10        On Error GoTo LoadFOB_Error

20        Found = False
30        cmdLock(5).Visible = False
40        fraFOB.Enabled = True
50        For n = 0 To 2
60            lblFOB(n) = ""
70            lblFOB(n).BackColor = &H8000000F
80        Next

90        LoadFOB = False

100       sql = "SELECT OB0, OB1, OB2 FROM Faeces WHERE " & _
                "SampleID = '" & SampleIDWithOffset & "' " & _
                "AND ( COALESCE(OB0, '') <> '' " & _
                "    OR COALESCE(OB1, '') <> '' " & _
                "    OR COALESCE(OB2, '') <> '' )"

110       Set tb = New Recordset
120       RecOpenServer 0, tb, sql

130       If Not tb.EOF Then
140           Found = True
150           For n = 0 To 2
160               If Trim$(tb("OB" & Format(n)) & "") <> "" Then
170                   FOB = Split(tb("OB" & Format(n)), "|")
180                   If UBound(FOB) = -1 Then
190                       lblFOB(n) = ""
200                   ElseIf UBound(FOB) > 1 Then
210                       lblFOB(n) = FOB(0)
220                       lblFOB(n).ForeColor = FOB(1)
230                       lblFOB(n).BackColor = FOB(2)
240                       LoadFOB = True
250                   Else
260                       lblFOB(n) = FOB(0)
270                   End If
280               End If
290           Next

300       End If

310       If Found Then
320           cmdLock(5).Visible = True
330           If LoadLockStatus(5) Then
340               fraFOB.Enabled = False
350           End If
360           LoadFOB = True
370       Else
380           LoadFOB = False
390       End If

400       Exit Function

LoadFOB_Error:

          Dim strES As String
          Dim intEL As Integer

410       intEL = Erl
420       strES = Err.Description
430       LogError "frmEditMicrobiologyNew", "LoadFOB", intEL, strES, sql


End Function

Private Sub LoadSensitivitiesForced(ByVal Index As Integer)

          Dim tb As Recordset
          Dim sql As String
          Dim Report As Boolean

10        On Error GoTo LoadSensitivitiesForced_Error

20        sql = "SELECT LTRIM(RTRIM(A.AntibioticName)) AS AntibioticName, " & _
                "S.Report, S.RSI, S.CPOFlag, S.Result, S.RunDateTime, S.UserName " & _
                "FROM Sensitivities S, Antibiotics A " & _
                "WHERE SampleID = '" & SampleIDWithOffset & "' " & _
                "AND IsolateNumber = '" & Index & "' " & _
                "AND S.AntibioticCode = A.Code " & _
                "AND S.Forced = 1"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        Do While Not tb.EOF
60            With grdAB(Index)
70                .AddItem tb!AntibioticName & vbTab & _
                           tb!RSI & vbTab & _
                           tb!CPOFlag & vbTab & _
                           tb!Result & vbTab & _
                           Format(tb!RunDateTime, "dd/mm/yy hh:mm") & _
                           tb!UserName & ""
80                .Row = .Rows - 1
90                .Col = 2
100               If IsNull(tb!Report) Then
110                   Set .CellPicture = Me.Picture
120               Else
130                   Set .CellPicture = IIf(tb!Report, imgSquareTick.Picture, imgSquareCross.Picture)
140               End If

150               .Col = 0
160               .CellBackColor = &HFFFFC0

170               tb.MoveNext
180           End With
190       Loop

200       Exit Sub

LoadSensitivitiesForced_Error:

          Dim strES As String
          Dim intEL As Integer

210       intEL = Erl
220       strES = Err.Description
230       LogError "frmEditMicrobiologyNew", "LoadSensitivitiesForced", intEL, strES, sql

End Sub

Private Sub LoadSensitivitiesSecondary(ByVal Index As Integer)

          Dim tb As Recordset
          Dim sql As String
          Dim Report As Boolean

10        sql = "SELECT LTRIM(RTRIM(A.AntibioticName)) AS AntibioticName, " & _
                "S.Report, S.RSI, S.CPOFlag, S.Result, S.RunDateTime, S.UserName " & _
                "FROM Sensitivities S, Antibiotics A " & _
                "WHERE SampleID = '" & SampleIDWithOffset & "' " & _
                "AND IsolateNumber = '" & Index & "' " & _
                "AND S.AntibioticCode = A.Code " & _
                "AND S.Secondary = 1"
20        Set tb = New Recordset
30        RecOpenServer 0, tb, sql
40        Do While Not tb.EOF
50            With grdAB(Index)
60                .AddItem tb!AntibioticName & vbTab & _
                           tb!RSI & vbTab & _
                           tb!CPOFlag & vbTab & _
                           tb!Result & vbTab & _
                           Format(tb!RunDateTime, "dd/mm/yy hh:mm") & _
                           tb!UserName & ""
70                .Row = .Rows - 1
80                .Col = 2
90                If IsNull(tb!Report) Then
100                   Set .CellPicture = Me.Picture
110               Else
120                   Set .CellPicture = IIf(tb!Report, imgSquareTick.Picture, imgSquareCross.Picture)
130               End If

140               .Col = 0
150               .CellFontBold = True

160               tb.MoveNext
170           End With
180       Loop

190       Exit Sub

LoadSensitivitiesForced_Error:

          Dim strES As String
          Dim intEL As Integer

200       intEL = Erl
210       strES = Err.Description
220       LogError "frmEditMicrobiologyNew", "LoadSensitivitiesForced", intEL, strES, sql

End Sub

Private Function LoadIdent() As Integer
'Returns number of Isolates Loaded

    Dim tb As Recordset
    Dim sql As String
    Dim n As Integer
    Dim Max As Integer

    On Error GoTo LoadIdent_Error

    ClearIdent

    Max = 0


    For n = 1 To 4
        sql = "Select * from UrineIdent where " & _
              "SampleID = '" & SampleIDWithOffset & "' " & _
              "and Isolate = " & n
        Set tb = New Recordset
        RecOpenClient 0, tb, sql

        If Not tb.EOF Then
            Max = n + 1
            cmbGram(n) = tb!Gram & ""
            txtZN(n) = tb!ZN & ""
            cmbWetPrep(n) = tb!WetPrep & ""
            txtIndole(n) = tb!Indole & ""
            txtCoagulase(n) = tb!Coagulase & ""
            txtCatalase(n) = tb!Catalase & ""
            txtOxidase(n) = tb!Oxidase & ""
            txtReincubation(n) = tb!Reincubation & ""
            txtNotes(n) = Trim$(tb!Notes & "")
        End If
    Next

    LoadIdent = Max
    Call LoadLockStatus(2)
    cmdSaveMicro.Enabled = False
    cmdSaveHold.Enabled = False

    Exit Function

LoadIdent_Error:

    Dim strES As String
    Dim intEL As Integer

    intEL = Erl
    strES = Err.Description
    LogError "frmEditMicrobiologyNew", "LoadIdent", intEL, strES, sql

End Function

Private Function LoadIsolates() As Boolean
      'returns true if loaded
          Dim tb As Recordset
          Dim sql As String
          Dim intIsolate As Integer

10        On Error GoTo LoadIsolates_Error

20        LoadIsolates = False

30        For intIsolate = 1 To 4
40            cmbOrgGroup(intIsolate) = ""
50            cmbOrgName(intIsolate) = ""
60            cmbQualifier(intIsolate) = ""
70            chkNonReportable(intIsolate - 1).Value = 0
80        Next

90        sql = "SELECT * FROM Isolates WHERE " & _
                "SampleID = '" & SampleIDWithOffset & "' " & _
                "AND IsolateNumber < 5"
100       Set tb = New Recordset
110       RecOpenClient 0, tb, sql
120       Do While Not tb.EOF
130           LoadIsolates = True
140           intIsolate = tb!IsolateNumber
150           cmbOrgGroup(intIsolate) = tb!OrganismGroup & ""

160           FillOrgNames intIsolate

170           cmbOrgName(intIsolate) = tb!OrganismName & ""
180           cmbQualifier(intIsolate) = tb!Qualifier & ""
190           chkNonReportable(intIsolate - 1) = IIf(IsNull(tb!NonReportable), 0, tb!NonReportable)
200           tb.MoveNext
210       Loop

220       Exit Function

LoadIsolates_Error:

          Dim strES As String
          Dim intEL As Integer

230       intEL = Erl
240       strES = Err.Description
250       LogError "frmEditMicrobiologyNew", "LoadIsolates", intEL, strES, sql
260       LoadIsolates = False

End Function


Private Function LoadUrine() As Boolean
      'Returns true if Urine Results Present

          Dim sql As String
          Dim tb As Recordset
          Dim U() As String
          Dim i As Integer
          Dim s() As String

10        On Error GoTo LoadUrine_Error

20        ClearUrine

30        LoadUrine = False

40        cmdLock(1).Visible = False
50        fraMicroscopy.Enabled = True
60        fraPregnancy.Enabled = True
70        txtUrineComment.Enabled = True

80        sql = "SELECT * FROM Urine WHERE " & _
                "SampleID = '" & SampleIDWithOffset & "'"
90        Set tb = New Recordset
100       RecOpenServer 0, tb, sql
110       If Not tb.EOF Then
120           Select Case tb!Pregnancy & ""
              Case "P": txtPregnancy = "Positive"
130           Case "N": txtPregnancy = "Negative"
140           Case "I": txtPregnancy = "Inconclusive"
150           Case "E": txtPregnancy = "Equivocal"
160           Case Else: txtPregnancy = tb!Pregnancy & ""
170           End Select
180           txtBacteria = Trim$(tb!Bacteria & "")

190           U = Split(tb!WCC & "", "|")
200           If UBound(U) = -1 Then
210               txtWCC = ""
220           ElseIf UBound(U) > 1 Then
230               txtWCC = U(0)
240               txtWCC.ForeColor = U(1)
250               txtWCC.BackColor = U(2)
260           Else
270               txtWCC = U(0)
280           End If

290           U = Split(tb!RCC & "", "|")
300           If UBound(U) = -1 Then
310               txtRCC = ""
320           ElseIf UBound(U) > 1 Then
330               txtRCC = U(0)
340               txtRCC.ForeColor = U(1)
350               txtRCC.BackColor = U(2)
360           Else
370               txtRCC = U(0)
380           End If

390           txtHCGLevel = Trim$(tb!HCGLevel & "")
400           cmbCrystals = Trim$(tb!Crystals & "")
410           cmbCasts = Trim$(tb!Casts & "")
420           cmbMisc(0) = Trim$(tb!Misc0 & "")
430           cmbMisc(1) = Trim$(tb!Misc1 & "")
440           cmbMisc(2) = Trim$(tb!Misc2 & "")

450           If Trim$(tb!WCC & "") <> "" Then
460               LoadUrine = True
470           ElseIf Trim$(tb!RCC & "") <> "" Then
480               LoadUrine = True
490           ElseIf Trim$(tb!Pregnancy & "") <> "" Then
500               LoadUrine = True
510           ElseIf Trim$(tb!Bacteria & "") <> "" Then
520               LoadUrine = True
530           ElseIf Trim$(tb!HCGLevel & "") <> "" Then
540               LoadUrine = True
550           ElseIf Trim$(tb!BenceJones & "") <> "" Then
560               LoadUrine = True
570           ElseIf Trim$(tb!SG & "") <> "" Then
580               LoadUrine = True
590           ElseIf Trim$(tb!FatGlobules & "") <> "" Then
600               LoadUrine = True
610           ElseIf Trim$(tb!pH & "") <> "" Then
620               LoadUrine = True
630           ElseIf Trim$(tb!Protein & "") <> "" Then
640               LoadUrine = True
650           ElseIf Trim$(tb!Glucose & "") <> "" Then
660               LoadUrine = True
670           ElseIf Trim$(tb!Ketones & "") <> "" Then
680               LoadUrine = True
690           ElseIf Trim$(tb!Urobilinogen & "") <> "" Then
700               LoadUrine = True
710           ElseIf Trim$(tb!Bilirubin & "") <> "" Then
720               LoadUrine = True
730           ElseIf Trim$(tb!BloodHb & "") <> "" Then
740               LoadUrine = True
750           ElseIf Trim$(tb!Crystals & "") <> "" Then
760               LoadUrine = True
770           ElseIf Trim$(tb!Casts & "") <> "" Then
780               LoadUrine = True
790           ElseIf Trim$(tb!Misc0 & "") <> "" Then
800               LoadUrine = True
810           ElseIf Trim$(tb!Misc1 & "") <> "" Then
820               LoadUrine = True
830           ElseIf Trim$(tb!Misc2 & "") <> "" Then
840               LoadUrine = True
850           End If

860           cmdLock(1).Visible = True
870           If LoadLockStatus(1) Then
880               fraMicroscopy.Enabled = False
890               fraPregnancy.Enabled = False
900               txtUrineComment.Enabled = False
910           End If



920       End If
930       If SysOptShowIQ200(0) = True Then
940           LoadIQ200
950           If IQ200ResultsExist Then
960               ReDim s(0 To 2) As String
                  '960       txtBacteria.Enabled = False
                  '970       txtWCC.Enabled = False
                  '980       txtRCC.Enabled = False
                  '990       cmbCrystals.Enabled = False
                  '1000      cmbCasts.Enabled = False
                  '1010      For i = 0 To 2
                  '1020          cmbMisc(i).Enabled = False
                  '1030      Next
                  '1040      cmdNADMicro.Enabled = False
970               For i = 1 To grdIQ200.Rows - 1
980                   If grdIQ200.TextMatrix(i, 0) = "RBC" Then
990                       s(0) = "Red Cells " & grdIQ200.TextMatrix(i, 2)
1000                  End If
1010                  If grdIQ200.TextMatrix(i, 0) = "WBC" Then
1020                      s(1) = "White Cells " & grdIQ200.TextMatrix(i, 2)
1030                  End If
1040                  If InStr(grdIQ200.TextMatrix(i, 1), "Epithelial") > 0 Then
1050                      s(2) = "Epithelial Cells " & grdIQ200.TextMatrix(i, 2)
1060                  End If
1070              Next
1080              lblCells.Caption = s(0) & "     " & s(1) & "     " & s(2)
1090          Else
1100              txtBacteria.Enabled = True
1110              txtWCC.Enabled = True
1120              txtRCC.Enabled = True
1130              cmbCrystals.Enabled = True
1140              cmbCasts.Enabled = True
1150              For i = 0 To 2
1160                  cmbMisc(i).Enabled = True
1170              Next
1180              cmdNADMicro.Enabled = True
1190          End If
1200      End If

1210      Exit Function

LoadUrine_Error:

          Dim strES As String
          Dim intEL As Integer

1220      intEL = Erl
1230      strES = Err.Description
1240      LogError "frmEditMicrobiologyNew", "LoadUrine", intEL, strES, sql

End Function

Private Sub LoadIQ200()
          Dim sql As String
          Dim tb As Recordset
          Dim s As String

10        On Error GoTo LoadIQ200_Error
'grdIQ200.Rows
          ClearIQ200


20        IQ200ResultsExist = False

30        sql = "SELECT * FROM IQ200 " & _
                "WHERE Sampleid = '" & SampleIDWithOffset & "' " & _
                "AND Result <> '[none]'"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        Do While Not tb.EOF
70            s = tb!TestCode & vbTab & tb!LongName & vbTab & tb!Result & vbTab & tb!Unit & ""
80            grdIQ200.AddItem s
90            tb.MoveNext
100       Loop

110       If grdIQ200.Rows > 2 Then
120           grdIQ200.RemoveItem 1
130           IQ200ResultsExist = True
140       End If

150       sql = "SELECT COUNT(*) Tot FROM IQ200Repeats WHERE " & _
                "SampleID = '" & SampleIDWithOffset & "' " & _
                "AND Result <> '[none]'"
160       Set tb = New Recordset
170       RecOpenServer 0, tb, sql
180       If tb!Tot > 0 Then
190           cmdIQ200Repeats.Enabled = True
200           cmdIQ200Repeats.BackColor = &H86C0FF
210       Else
220           cmdIQ200Repeats.Enabled = False
230           cmdIQ200Repeats.BackColor = &H8000000F
240       End If


250       Exit Sub

LoadIQ200_Error:

          Dim strES As String
          Dim intEL As Integer

260       intEL = Erl
270       strES = Err.Description
280       LogError "frmEditMicrobiologyNew", "LoadIQ200", intEL, strES, sql


End Sub

Private Sub OrderFaeces()

          Dim f As Form

10        On Error GoTo OrderFaeces_Error

20        Set f = New frmMicroOrderFaecesNew
30        With f
40            .txtSampleID = txtSampleID
50            .Show 1
60        End With
70        Unload f
80        Set f = Nothing

90        LoadFaecalOrders

100       cmdOrderTests.Enabled = True

110       Exit Sub

OrderFaeces_Error:

          Dim strES As String
          Dim intEL As Integer

120       intEL = Erl
130       strES = Err.Description
140       LogError "frmEditMicrobiologyNew", "OrderFaeces", intEL, strES
End Sub

Private Sub LoadFaecalOrders()

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo LoadFaecalOrders_Error

20        SampleIDWithOffset = Val(txtSampleID) + SysOptMicroOffset(0)

30        sql = "SELECT " & _
                "COALESCE(CS, 0) CS, " & _
                "COALESCE(ssScreen, 0) ssScreen, " & _
                "COALESCE(Campylobacter, 0) Campylobacter, " & _
                "COALESCE(Coli0157, 0) Coli0157, " & _
                "COALESCE(Cryptosporidium, 0) Cryptosporidium, " & _
                "COALESCE(Rota, 0) Rota, " & _
                "COALESCE(Adeno, 0) Adeno, " & _
                "COALESCE(OB0, 0) OB0, " & _
                "COALESCE(OB1, 0) OB1, " & _
                "COALESCE(OB2, 0) OB2, " & _
                "COALESCE(OP, 0) OP, " & _
                "COALESCE(ToxinAB, 0) ToxinAB, " & _
                "COALESCE(CDiff, 0) CDiffCulture, " & _
                "COALESCE(GDH, 0) GDH, " & _
                "COALESCE(PCR, 0) PCR, " & _
                "COALESCE(HPylori, 0) HPylori, " & _
                "COALESCE(RedSub, 0) RedSub ," & _
                "COALESCE(GL, 0) GL " & _
                "FROM FaecalRequests where " & _
                "SampleID = '" & SampleIDWithOffset & "' "
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        If Not tb.EOF Then

70            SSTab1.TabVisible(6) = (tb!Rota Or tb!Adeno = 1)   'rota/adeno
80            SSTab1.TabVisible(5) = (tb!OB0 Or tb!OB1 Or tb!OB2 = 1)    'fob
90            SSTab1.TabVisible(11) = (tb!OP Or tb!Cryptosporidium = 1 Or tb!GL = 1)  'OP or Cryptosporidium
100           SSTab1.TabVisible(10) = (tb!ToxinAB = 1) Or (tb!CDiffCulture = 1) Or (tb!GDH = 1) Or (tb!PCR = 1)
110           SSTab1.TabVisible(13) = tb!HPylori = 1
120           SSTab1.TabVisible(7) = tb!RedSub = 1
130       End If

140       Exit Sub

LoadFaecalOrders_Error:

          Dim strES As String
          Dim intEL As Integer

150       intEL = Erl
160       strES = Err.Description
170       LogError "frmEditMicrobiologyNew", "LoadFaecalOrders", intEL, strES, sql

End Sub

Private Sub PrintThis(Optional PrintAction As String)

          Dim tb As Recordset
          Dim sql As String
          Dim FinalOrInterim As String

10        On Error GoTo PrintThis_Error

20        pBar = 0
30        GetSampleIDWithOffset
'40        If Not EntriesOK(txtSampleID, txtName, txtSex, cmbWard.Text, cmbGP.Text) Then
'50            Exit Sub
'60        End If

70        If Not CheckTimes() Then Exit Sub

80        SaveDemographics

90        sql = "Select * from PrintPending where " & _
                "Department = 'N' " & _
                "and PrintAction = '" & PrintAction & "' " & _
                "and SampleID = '" & SampleIDWithOffset & "'"
100       Set tb = New Recordset
110       RecOpenClient 0, tb, sql
120       If tb.EOF Then
130           tb.AddNew
140       End If
150       tb!SampleID = SampleIDWithOffset
160       tb!Ward = cmbWard
170       tb!Clinician = cmbClinician
180       tb!GP = cmbGP
190       tb!Department = "N"
200       tb!Initiator = UserName
210       tb!UsePrinter = pPrintToPrinter
220       tb!NoOfCopies = Val(txtNoCopies)
230       FinalOrInterim = "F"
240       If lblInterim.BackColor = vbGreen Or PrintAction = "SaveTemp" Then ' Or PrintAction = "SaveFinal"  Then
250           FinalOrInterim = "I"
260       End If
270       tb!FinalInterim = FinalOrInterim
280       tb!pTime = Now
290       tb!PrintAction = PrintAction
300       tb.Update

310       Exit Sub

PrintThis_Error:

          Dim strES As String
          Dim intEL As Integer

320       intEL = Erl
330       strES = Err.Description
340       LogError "frmEditMicrobiologyNew", "PrintThis", intEL, strES, sql

End Sub

Private Sub SaveComments()

      Dim tb As Recordset
      Dim sql As String
      Dim blnCommentsPresent As Boolean
      Dim blnUpdate As Boolean
      Dim Obs As New Observations

10    On Error GoTo SaveComments_Error

20    If Trim$(txtSampleID) = "" Then Exit Sub

30    If InStr(UCase$(txtDemographicComment), "CYSTIC FIBROSIS PATIENT") > 0 Then
40        If iMsg("Confirmation Required." & vbCrLf & _
                  "Is this is a Cystic Fibrosis Patient?", vbQuestion + vbYesNo, , vbRed) = vbNo Then
50            txtDemographicComment = ""
60        Else
70            LogError "frmEditMicrobiologyNew", "SaveComments", 70, "Cystic Fibrosis Confirmed by " & UserName
80        End If
90    End If

100   Obs.Save SampleIDWithOffset, True, _
               "Demographic", Trim$(txtDemographicComment), _
               "MicroGeneral", Trim$(txtUrineComment), _
               "CSFFluid", txtFluidComment
110   If SSTab1.TabVisible(4) = True Then
120       If txtMSC = "Medical Scientist Comments" Or Trim$(txtMSC) = "" Then
130           Obs.Save SampleIDWithOffset, True, "MicroCS", ""
140       Else
150           Obs.Save SampleIDWithOffset, True, "MicroCS", Trim$(txtMSC)
160       End If
170   End If
180   If Trim$(txtCDiffMSC) = "" Then
190       Obs.Save SampleIDWithOffset, True, "MICROCDIFF", ""
200   Else
210       Obs.Save SampleIDWithOffset, True, "MICROCDIFF", Trim$(txtCDiffMSC)
220   End If

230   If txtConC = "Consultant Comments" Or Trim$(txtConC) = "" Then
240       Obs.Save SampleIDWithOffset, True, "MicroConsultant", ""
250   Else
260       Obs.Save SampleIDWithOffset, True, "MicroConsultant", Trim$(txtConC)
270   End If

280   Exit Sub

SaveComments_Error:

      Dim strES As String
      Dim intEL As Integer

290   intEL = Erl
300   strES = Err.Description
310   LogError "frmEditMicrobiologyNew", "SaveComments", intEL, strES

End Sub

Private Sub SaveFaeces()

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo SaveFaeces_Error

20        sql = "Select * from FaecesWorkSheet where " & _
                "SampleID = '" & SampleIDWithOffset & "'"
30        Set tb = New Recordset
40        RecOpenClient 0, tb, sql
50        If tb.EOF Then
60            tb.AddNew
70        End If

80        tb!SampleID = SampleIDWithOffset
90        tb!TimeOfRecord = Format$(Now, "dd/mmm/yyyy hh:mm:ss")
100       tb!Operator = UserName

          'Day 1
          'XLD
110       tb!Day111 = cmbDay1(11)
120       tb!Day112 = cmbDay1(12)
130       tb!Day113 = cmbDay1(13)
          'DCA
140       tb!Day121 = cmbDay1(21)
150       tb!Day122 = cmbDay1(22)
160       tb!Day123 = cmbDay1(23)
          'SMAC
170       tb!Day131 = cmbDay1(31)
180       tb!Day132 = cmbDay1(32)
190       tb!Day133 = cmbDay1(33)
          'STEC
200       tb!Day141 = cmbDay1(41)
210       tb!Day142 = cmbDay1(42)
220       tb!Day143 = cmbDay1(43)

          'Day2
          'XLD
230       tb!Day211 = cmbDay2(11)
240       tb!Day212 = cmbDay2(12)
250       tb!Day213 = cmbDay2(13)

          'DCA
260       tb!Day241 = cmbDay2(41)
270       tb!Day242 = cmbDay2(42)
280       tb!Day243 = cmbDay2(43)

          'CROMO
290       tb!Day221 = cmbDay2(21)
300       tb!Day222 = cmbDay2(22)
310       tb!Day223 = cmbDay2(23)
          'CAMP
320       tb!Day231 = cmbDay2(31)
330       tb!Day232 = cmbDay2(32)
340       tb!Day233 = cmbDay2(33)
          'STEC
350       tb!Day251 = cmbDay2(51)
360       tb!Day252 = cmbDay2(52)
370       tb!Day253 = cmbDay2(53)

          'Day 3
          'CAMP
380       tb!Day31 = cmbDay3(1)
390       tb!Day32 = cmbDay3(2)
400       tb!Day33 = cmbDay3(3)

          'CROMO
410       tb!Day34 = cmbDay3(4)
420       tb!Day35 = cmbDay3(5)
430       tb!Day36 = cmbDay3(6)

440       tb.Update

450       Exit Sub

SaveFaeces_Error:

          Dim strES As String
          Dim intEL As Integer

460       intEL = Erl
470       strES = Err.Description
480       LogError "frmEditMicrobiologyNew", "SaveFaeces", intEL, strES, sql


End Sub

Private Sub SaveOP()

          Dim tb As Recordset
          Dim sql As String
          Dim n As Integer

10        On Error GoTo SaveOp_Error

20        sql = "Select * from Faeces where " & _
                "SampleID = '" & SampleIDWithOffset & "'"
30        Set tb = New Recordset
40        RecOpenClient 0, tb, sql

50        If tb.EOF Then
60            tb.AddNew
70        End If
80        tb!SampleID = SampleIDWithOffset
90        If Trim$(lblCrypto) = "" Then
100           tb!Cryptosporidium = Null
110       Else
120           tb!Cryptosporidium = lblCrypto & "|" & lblCrypto.ForeColor & "|" & lblCrypto.BackColor
130       End If
140       If Trim$(lblGiardia) = "" Then
150           tb!GiardiaLambila = Null
160       Else
170           tb!GiardiaLambila = lblGiardia & "|" & lblGiardia.ForeColor & "|" & lblGiardia.BackColor
180       End If
190       For n = 0 To 2
200           tb("OP" & Format(n)) = cmbOva(n)
210       Next
220       tb!UserName = UserName
230       tb.Update

240       Exit Sub

SaveOp_Error:

          Dim strES As String
          Dim intEL As Integer

250       intEL = Erl
260       strES = Err.Description
270       LogError "frmEditMicrobiologyNew", "SaveOp", intEL, strES, sql

End Sub


Private Sub SaveCdiff()

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo SaveCdiff_Error

20        sql = "Select * from Faeces where " & _
                "SampleID = '" & SampleIDWithOffset & "'"
30        Set tb = New Recordset
40        RecOpenClient 0, tb, sql

50        If tb.EOF Then
60            tb.AddNew
70        End If
80        tb!SampleID = SampleIDWithOffset
90        If Trim$(lblToxinA) = "" Then
100           tb!ToxinAB = Null
110       Else
120           tb!ToxinAB = lblToxinA & "|" & lblToxinA.ForeColor & "|" & lblToxinA.BackColor
130       End If
140       If Trim$(lblCDiffCulture) = "" Then
150           tb!CDiffCulture = Null
160       Else
170           tb!CDiffCulture = lblCDiffCulture & "|" & lblCDiffCulture.ForeColor & "|" & lblCDiffCulture.BackColor
180       End If
190       If Trim$(lblGDH) = "" Then
200           tb!GDH = Null
210       Else
220           tb!GDH = lblGDH & "|" & lblGDH.ForeColor & "|" & lblGDH.BackColor
230       End If

240       If Trim$(cmbGDH) = "" Then
250           tb!GDHDetail = Null
260       Else
270           tb!GDHDetail = cmbGDH
280       End If

290       If Trim$(lblPCR) = "" Then
300           tb!PCR = Null
310       Else
320           tb!PCR = lblPCR & "|" & lblPCR.ForeColor & "|" & lblPCR.BackColor
330       End If

340       If Trim$(cmbPCR) = "" Then
350           tb!PCRDetail = Null
360       Else
370           tb!PCRDetail = cmbPCR
380       End If

390       tb!UserName = UserName

400       tb.Update

410       Exit Sub

SaveCdiff_Error:

          Dim strES As String
          Dim intEL As Integer

420       intEL = Erl
430       strES = Err.Description
440       LogError "frmEditMicrobiologyNew", "SaveCdiff", intEL, strES, sql


End Sub


Private Sub SaveRotaAdeno()

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo SaveRotaAdeno_Error

20        sql = "Select * from Faeces where " & _
                "SampleID = '" & SampleIDWithOffset & "'"
30        Set tb = New Recordset
40        RecOpenClient 0, tb, sql

50        If tb.EOF Then
60            tb.AddNew
70        End If
80        tb!SampleID = SampleIDWithOffset

90        If Trim$(txtRota) = "" Then
100           tb!Rota = Null
110       Else
120           tb!Rota = txtRota & "|" & txtRota.ForeColor & "|" & txtRota.BackColor
130       End If

140       If Trim$(txtAdeno) = "" Then
150           tb!Adeno = Null
160       Else
170           tb!Adeno = txtAdeno & "|" & txtAdeno.ForeColor & "|" & txtAdeno.BackColor
180       End If

190       tb!UserName = UserName
200       tb.Update

210       Exit Sub

SaveRotaAdeno_Error:

          Dim strES As String
          Dim intEL As Integer

220       intEL = Erl
230       strES = Err.Description
240       LogError "frmEditMicrobiologyNew", "SaveRotaAdeno", intEL, strES, sql

End Sub


Private Sub SaveFOB()

          Dim tb As Recordset
          Dim sql As String
          Dim n As Integer

10        On Error GoTo SaveFOB_Error

20        sql = "Select * from Faeces where " & _
                "SampleID = '" & SampleIDWithOffset & "'"
30        Set tb = New Recordset
40        RecOpenClient 0, tb, sql

50        If tb.EOF Then
60            tb.AddNew
70        End If
80        tb!SampleID = SampleIDWithOffset
90        For n = 0 To 2
100           If Trim$(lblFOB(n)) = "" Then
110               tb("OB" & Format(n)) = Null
120           Else
130               tb("OB" & Format(n)) = lblFOB(n) & "|" & lblFOB(n).ForeColor & "|" & lblFOB(n).BackColor
140           End If
150       Next
160       tb!UserName = UserName
170       tb.Update

180       Exit Sub

SaveFOB_Error:

          Dim strES As String
          Dim intEL As Integer

190       intEL = Erl
200       strES = Err.Description
210       LogError "frmEditMicrobiologyNew", "SaveFOB", intEL, strES, sql


End Sub

Private Sub SaveRSV()

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo SaveRSV_Error

20        If Trim$(lblRSV.Caption) <> "" Then
30            sql = "Select * from GenericResults where " & _
                    "SampleID = '" & SampleIDWithOffset & "' " & _
                    "and TestName = 'RSV'"
40            Set tb = New Recordset
50            RecOpenClient 0, tb, sql

60            If tb.EOF Then
70                tb.AddNew
80            End If
90            tb!SampleID = SampleIDWithOffset
100           tb!TestName = "RSV"
110           tb!Result = lblRSV.Caption & "|" & lblRSV.ForeColor & "|" & lblRSV.BackColor
120           tb!UserName = UserName
130           tb.Update
140       Else

150           sql = "Delete from GenericResults where " & _
                    "SampleID = '" & SampleIDWithOffset & "' " & _
                    "and TestName = 'RSV'"
160           Cnxn(0).Execute sql
170       End If

180       Exit Sub

SaveRSV_Error:

          Dim strES As String
          Dim intEL As Integer

190       intEL = Erl
200       strES = Err.Description
210       LogError "frmEditMicrobiologyNew", "SaveRSV", intEL, strES, sql

End Sub



Private Sub SaveHPylori()

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo SaveHPylori_Error

20        sql = "Select * from Faeces where " & _
                "SampleID = '" & SampleIDWithOffset & "'"

30        Set tb = New Recordset
40        RecOpenClient 0, tb, sql
50        If tb.EOF Then
60            tb.AddNew
70        End If
80        tb!SampleID = SampleIDWithOffset
90        If Trim$(lblHPylori) = "" Then
100           tb!HPylori = Null
110       Else
120           tb!HPylori = lblHPylori.Caption & "|" & lblHPylori.ForeColor & "|" & lblHPylori.BackColor
130       End If
140       tb!UserName = UserName
150       tb.Update

160       Exit Sub

SaveHPylori_Error:

          Dim strES As String
          Dim intEL As Integer

170       intEL = Erl
180       strES = Err.Description
190       LogError "frmEditMicrobiologyNew", "SaveHPylori", intEL, strES, sql

End Sub
Private Sub SaveRedSub()

          Dim tb As Recordset
          Dim sql As String
          Dim n As Integer
          Dim Result As String

10        On Error GoTo SaveRedSub_Error

20        Result = ""
30        For n = 0 To 5
40            If chkRS(n).Value = 1 Then
50                Result = chkRS(n).Caption
60                Exit For
70            End If
80        Next

90        If Result = "" Then

100           sql = "Delete from GenericResults where " & _
                    "SampleID = '" & SampleIDWithOffset & "' " & _
                    "and TestName = 'RedSub'"
110           Cnxn(0).Execute sql
120       Else
130           sql = "Select * from GenericResults where " & _
                    "SampleID = '" & SampleIDWithOffset & "' " & _
                    "and TestName = 'RedSub'"
140           Set tb = New Recordset
150           RecOpenClient 0, tb, sql

160           If tb.EOF Then
170               tb.AddNew
180           End If
190           tb!SampleID = SampleIDWithOffset
200           tb!TestName = "RedSub"
210           tb!Result = Result
220           tb!UserName = UserName
230           tb.Update
240       End If

250       Exit Sub

SaveRedSub_Error:

          Dim strES As String
          Dim intEL As Integer

260       intEL = Erl
270       strES = Err.Description
280       LogError "frmEditMicrobiologyNew", "SaveRedSub", intEL, strES, sql


End Sub

Private Sub SaveIdent()

          Dim tb As Recordset
          Dim sql As String
          Dim n As Integer

10        On Error GoTo SaveIdent_Error

20        For n = 1 To 4
30            If IdentIsSaveable(n) Then
40                sql = "Select * from UrineIdent where " & _
                        "SampleID = '" & SampleIDWithOffset & "' " & _
                        "and Isolate = " & n
50                Set tb = New Recordset
60                RecOpenClient 0, tb, sql

70                If tb.EOF Then
80                    tb.AddNew
90                End If
100               tb!SampleID = SampleIDWithOffset
110               tb!Isolate = n
120               tb!SampleID = SampleIDWithOffset
130               tb!Gram = cmbGram(n)
140               tb!ZN = txtZN(n)
150               tb!WetPrep = cmbWetPrep(n)
160               tb!Indole = txtIndole(n)
170               tb!Coagulase = txtCoagulase(n)
180               tb!Catalase = txtCatalase(n)
190               tb!Oxidase = txtOxidase(n)
200               tb!Reincubation = txtReincubation(n)
210               tb!Notes = txtNotes(n)
220               tb!UserName = UserName
230               tb.Update
240           Else
250               sql = "Delete from UrineIdent where " & _
                        "SampleID = '" & SampleIDWithOffset & "' " & _
                        "and Isolate = " & n
260               Cnxn(0).Execute sql
270           End If

280       Next

290       Exit Sub

SaveIdent_Error:

          Dim strES As String
          Dim intEL As Integer

300       intEL = Erl
310       strES = Err.Description
320       LogError "frmEditMicrobiologyNew", "SaveIdent", intEL, strES, sql

End Sub
Private Sub SaveIsolates()

          Dim tb As Recordset
          Dim sql As String
          Dim intIsolate As Integer

10        On Error GoTo SaveIsolates_Error

20        For intIsolate = 1 To 4
30            If cmbOrgGroup(intIsolate) <> "" Then
40                sql = "Select * from Isolates where " & _
                        "SampleID = '" & SampleIDWithOffset & "' " & _
                        "and IsolateNumber = '" & intIsolate & "'"
50                Set tb = New Recordset
60                RecOpenClient 0, tb, sql
70                If tb.EOF Then
80                    tb.AddNew
90                End If
100               tb!SampleID = SampleIDWithOffset
110               tb!IsolateNumber = intIsolate
120               tb!OrganismGroup = cmbOrgGroup(intIsolate)
130               tb!OrganismName = cmbOrgName(intIsolate)
140               tb!Qualifier = cmbQualifier(intIsolate)
150               tb!UserName = UserName
160               tb!NonReportable = chkNonReportable(intIsolate - 1).Value
170               tb.Update

180           Else

190               sql = "Delete from Isolates where " & _
                        "SampleID = '" & SampleIDWithOffset & "' " & _
                        "and IsolateNumber = '" & intIsolate & "'"
200               Cnxn(0).Execute sql

210               sql = "Delete from Sensitivities where " & _
                        "SampleID = '" & SampleIDWithOffset & "' " & _
                        "and IsolateNumber = '" & intIsolate & "'"
220               Cnxn(0).Execute sql

230           End If
240       Next

250       Exit Sub

SaveIsolates_Error:

          Dim strES As String
          Dim intEL As Integer

260       intEL = Erl
270       strES = Err.Description
280       LogError "frmEditMicrobiologyNew", "SaveIsolates", intEL, strES, sql


End Sub

Private Sub SaveUrine()

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo SaveUrine_Error

20        sql = "Select * from Urine where " & _
                "SampleID = '" & SampleIDWithOffset & "'"
30        Set tb = New Recordset
40        RecOpenClient 0, tb, sql
50        If tb.EOF Then
60            tb.AddNew
70        End If
80        tb!SampleID = SampleIDWithOffset
90        tb!Pregnancy = txtPregnancy
100       tb!Bacteria = txtBacteria
110       tb!HCGLevel = txtHCGLevel
120       tb!WCC = txtWCC.Text & "|" & txtWCC.ForeColor & "|" & txtWCC.BackColor
130       tb!RCC = txtRCC.Text & "|" & txtRCC.ForeColor & "|" & txtRCC.BackColor
140       tb!Crystals = cmbCrystals
150       tb!Casts = cmbCasts
160       tb!Misc0 = cmbMisc(0)
170       tb!Misc1 = cmbMisc(1)
180       tb!Misc2 = cmbMisc(2)
          'tb!Valid = Validate
190       tb!UserName = UserName
200       tb.Update

210       Exit Sub

SaveUrine_Error:

          Dim strES As String
          Dim intEL As Integer

220       intEL = Erl
230       strES = Err.Description
240       LogError "frmEditMicrobiologyNew", "SaveUrine", intEL, strES, sql

End Sub

Private Sub SetAsForced(ByVal intIndex As Integer, _
                        ByVal strABName As String, _
                        ByVal blnReport As Boolean)

          Dim tb As Recordset
          Dim sql As String
          Dim SID As Double
          Dim Rpt As Integer

10        On Error GoTo SetAsForced_Error

20        Rpt = IIf(blnReport, 1, 0)
30        SID = SysOptMicroOffset(0) + Val(txtSampleID)

40        sql = "IF EXISTS (SELECT * FROM ForcedABReport WHERE " & _
                "           ABName = '" & strABName & "' " & _
                "           AND [Index] = " & intIndex & " " & _
                "           AND SampleID = " & SID & ") " & _
                "  UPDATE ForcedABReport " & _
                "  SET ABName = '" & strABName & "', " & _
                "  Report = " & Rpt & " " & _
                "  WHERE ABName = '" & strABName & "' " & _
                "  AND [Index] = " & intIndex & " " & _
                "  AND SampleID = " & SID & " " & _
                "ELSE " & _
                "  INSERT INTO ForcedABReport " & _
                "  (SampleID, ABName, Report, [Index]) VALUES " & _
                "  (" & SID & ", " & _
                "   '" & strABName & "', " & _
                "   '" & Rpt & "', " & _
                "   '" & intIndex & "')"
50        Cnxn(0).Execute sql

          '
          '20    sql = "Select * from ForcedABReport where " & _
           '            "ABName = '" & strABName & "' " & _
           '            "and [Index] = " & intIndex & " " & _
           '            "and SampleID = " & SysOptMicroOffset(0) + Val(txtSampleID)
          '30    Set tb = New Recordset
          '40    RecOpenServer 0, tb, sql
          '50    If tb.EOF Then
          '60        tb.AddNew
          '70    End If
          '80    tb!SampleID = SysOptMicroOffset(0) + Val(txtSampleID)
          '90    tb!ABName = strABName
          '100   tb!Report = blnReport
          '110   tb!Index = intIndex
          '120   tb.Update

60        Exit Sub

SetAsForced_Error:

          Dim strES As String
          Dim intEL As Integer

70        intEL = Erl
80        strES = Err.Description
90        LogError "frmEditMicrobiologyNew", "SetAsForced", intEL, strES, sql

End Sub

Private Sub chkPregnant_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

10        If chkPregnant.Value = 1 And InStr(txtClinDetails, "Pregnant") = 0 Then
20            txtClinDetails = txtClinDetails & " Pregnant;"
30        End If

40        cmdSaveDemographics.Enabled = True

End Sub

Private Sub chkRS_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)

          Dim n As Integer
          Dim intOriginal As Integer

10        intOriginal = chkRS(Index).Value

20        For n = 0 To 5
30            chkRS(n).Value = 0
40        Next

50        chkRS(Index).Value = intOriginal

60        ShowUnlock 7

End Sub


Private Sub cmbABSelect_Click(Index As Integer)

          Dim sql As String
          Dim tb As Recordset
          Dim Y As Integer

10        On Error GoTo cmbABSelect_Click_Error

20        grdAB(Index).AddItem cmbABSelect(Index).Text
30        grdAB(Index).Row = grdAB(Index).Rows - 1
40        grdAB(Index).Col = 0
50        grdAB(Index).CellBackColor = &HFFFFC0
60        grdAB(Index).Col = 2
70        Set grdAB(Index).CellPicture = Me.Picture

80        sql = "Select distinct * from Sensitivities as S, Antibiotics as A where " & _
                "SampleID = '" & SampleIDWithOffset & "' " & _
                "and IsolateNumber = '" & Index & "' " & _
                "and S.AntibioticCode = A.Code " & _
                "and AntibioticName = '" & cmbABSelect(Index).Text & "'"
90        Set tb = New Recordset
100       RecOpenClient 0, tb, sql
110       If Not tb.EOF Then

120           With grdAB(Index)
130               Y = .Rows - 1
140               .Row = Y
150               .TextMatrix(Y, 1) = tb!RSI & ""
160               .TextMatrix(Y, 2) = tb!CPOFlag & ""
170               .TextMatrix(Y, 3) = tb!Result & ""
180               .TextMatrix(Y, 4) = Format(tb!RunDateTime, "dd/mm/yy hh:mm")
190               .TextMatrix(Y, 5) = tb!UserName & ""
200               .Col = 2
210               If IsNull(tb!Report) Then
220                   Set .CellPicture = Me.Picture
230               Else
240                   Set .CellPicture = IIf(tb!Report, imgSquareTick.Picture, imgSquareCross.Picture)
250               End If
260           End With

270       End If

280       cmbABSelect(Index) = ""

290       FillABSelect Index

300       cmdSaveMicro.Enabled = True
310       cmdSaveHold.Enabled = True

320       Exit Sub

cmbABSelect_Click_Error:

          Dim strES As String
          Dim intEL As Integer

330       intEL = Erl
340       strES = Err.Description
350       LogError "frmEditMicrobiologyNew", "cmbABSelect_Click", intEL, strES, sql


End Sub

Private Sub cmbABSelect_KeyPress(Index As Integer, KeyAscii As Integer)

10        KeyAscii = 0

End Sub

Private Sub cmbABsInUse_Click()

          Dim n As Integer

10        lstABsInUse.AddItem cmbABsInUse
20        cmbABsInUse.Visible = False
30        lstABsInUse.Visible = True

40        lblABsInUse = ""
50        For n = 0 To lstABsInUse.ListCount - 1
60            lblABsInUse = lblABsInUse & lstABsInUse.List(n) & " "
70        Next

End Sub


Private Sub cmbCasts_Click()

10        ShowUnlock 1

End Sub


Private Sub cmbCasts_LostFocus()

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo cmbCasts_LostFocus_Error

20        sql = "Select * from Lists where " & _
                "ListType = 'CA' " & _
                "and Code = '" & UCase(cmbCasts) & "'"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        If Not tb.EOF Then
60            cmbCasts = tb!Text & ""
70        End If

80        Exit Sub

cmbCasts_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

90        intEL = Erl
100       strES = Err.Description
110       LogError "frmEditMicrobiologyNew", "cmbCasts_LostFocus", intEL, strES, sql


End Sub

Private Sub cmbConC_Click()

10        If txtConC = "Consultant Comments" Then
20            txtConC = ""
30        End If

40        txtConC = txtConC & cmbConC
50        txtConC.SetFocus
60        txtConC.SelStart = Len(txtConC)
70        cmbConC = ""
80        cmbConC.Visible = False

90        cmdSaveMicro.Enabled = True
100       cmdSaveHold.Enabled = True

End Sub


Private Sub cmbConC_LostFocus()

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo cmbConC_LostFocus_Error

20        If cmbConC <> "" Then
30            sql = "SELECT * FROM Lists WHERE " & _
                    "ListType = 'BA' " & _
                    "AND Code = '" & cmbConC & "'"
40            Set tb = New Recordset
50            RecOpenServer 0, tb, sql
60            If Not tb.EOF Then
70                cmbConC = tb!Text & ""
80            End If
90        End If

100       If txtConC() = "Consultant Comments" Then
110           txtConC = cmbConC
120       Else
130           txtConC = txtConC & cmbConC
140       End If

150       cmbConC.Visible = False
160       cmbConC = ""

170       Exit Sub

cmbConC_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

180       intEL = Erl
190       strES = Err.Description
200       LogError "frmEditMicrobiologyNew", "cmbConC_LostFocus", intEL, strES, sql


End Sub


Private Sub cmbConsultantVal_Click()

10        txtSampleID = cmbConsultantVal
20        txtSampleID = Format$(Val(txtSampleID))
30        If txtSampleID = 0 Then Exit Sub

40        GetSampleIDWithOffset

50        LoadAllDetails

60        cmdSaveDemographics.Enabled = False
70        cmdSaveInc.Enabled = False
80        cmdSaveMicro.Enabled = False
90        cmdSaveHold.Enabled = False

End Sub

Private Sub cmbCrystals_Click()

10        ShowUnlock 1

End Sub


Private Sub cmbCrystals_LostFocus()

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo cmbCrystals_LostFocus_Error

20        sql = "Select * from Lists where " & _
                "ListType = 'CR' " & _
                "and Code = '" & UCase(cmbCrystals) & "'"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        If Not tb.EOF Then
60            cmbCrystals = tb!Text & ""
70        End If

80        Exit Sub

cmbCrystals_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

90        intEL = Erl
100       strES = Err.Description
110       LogError "frmEditMicrobiologyNew", "cmbCrystals_LostFocus", intEL, strES, sql


End Sub

Private Sub cmbDay1_Click(Index As Integer)

10        cmdSaveMicro.Enabled = True
20        cmdSaveHold.Enabled = True

End Sub


Private Sub cmbDay1_KeyPress(Index As Integer, KeyAscii As Integer)

      '10        KeyAscii = AutoComplete(cmbDay1(Index), KeyAscii, False)

10        cmdSaveMicro.Enabled = True
20        cmdSaveHold.Enabled = True

End Sub

Private Sub cmbDay2_Click(Index As Integer)

10        cmdSaveMicro.Enabled = True
20        cmdSaveHold.Enabled = True

End Sub


Private Sub cmbDay2_KeyPress(Index As Integer, KeyAscii As Integer)

      '10        KeyAscii = AutoComplete(cmbDay2(Index), KeyAscii, False)

10        cmdSaveMicro.Enabled = True
20        cmdSaveHold.Enabled = True

End Sub

Private Sub cmbDay3_Click(Index As Integer)



10        cmdSaveMicro.Enabled = True
20        cmdSaveHold.Enabled = True

End Sub


Private Sub cmbDay3_KeyPress(Index As Integer, KeyAscii As Integer)

      '10        KeyAscii = AutoComplete(cmbDay3(Index), KeyAscii, False)
10        cmdSaveMicro.Enabled = True
20        cmdSaveHold.Enabled = True

End Sub

Private Sub cmbDemogComment_Click()

10        txtDemographicComment = Trim$(txtDemographicComment & " " & cmbDemogComment)
20        cmbDemogComment = ""

30        cmdSaveDemographics.Enabled = True
40        cmdSaveInc.Enabled = True

End Sub


Private Sub cmbDemogComment_LostFocus()

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo cmbDemogComment_LostFocus_Error

20        sql = "Select * from Lists where " & _
                "ListType = 'DE' " & _
                "and Code = '" & cmbDemogComment & "'"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        If Not tb.EOF Then
60            txtDemographicComment = Trim$(txtDemographicComment & " " & tb!Text & "")
70        Else
80            txtDemographicComment = Trim$(txtDemographicComment & " " & cmbDemogComment)
90        End If
100       cmbDemogComment = ""

110       Exit Sub

cmbDemogComment_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

120       intEL = Erl
130       strES = Err.Description
140       LogError "frmEditMicrobiologyNew", "cmbDemogComment_LostFocus", intEL, strES, sql


End Sub

Private Sub cmbGram_Click(Index As Integer)

10        cmdSaveMicro.Enabled = True
20        cmdSaveHold.Enabled = True

End Sub


Private Sub cmbGram_LostFocus(Index As Integer)

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo cmbGram_LostFocus_Error

20        sql = "Select * from Lists where " & _
                "ListType = 'GS' " & _
                "and Code = '" & cmbGram(Index) & "'"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        If Not tb.EOF Then
60            cmbGram(Index) = tb!Text & ""
70        End If

80        Exit Sub

cmbGram_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

90        intEL = Erl
100       strES = Err.Description
110       LogError "frmEditMicrobiologyNew", "cmbGram_LostFocus", intEL, strES, sql


End Sub


Private Sub cmbHospital_Click()

10        FillGPsClinWard Me, cmbHospital

20        cmdSaveDemographics.Enabled = True
30        cmdSaveInc.Enabled = True

End Sub


Private Sub cmbMisc_Click(Index As Integer)

10        ShowUnlock 1

End Sub


Private Sub cmbMisc_LostFocus(Index As Integer)

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo cmbMisc_LostFocus_Error

20        sql = "Select * from Lists where " & _
                "ListType = 'MI' " & _
                "and Code = '" & cmbMisc(Index) & "'"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        If Not tb.EOF Then
60            cmbMisc(Index) = tb!Text & ""
70        End If

80        Exit Sub

cmbMisc_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

90        intEL = Erl
100       strES = Err.Description
110       LogError "frmEditMicrobiologyNew", "cmbMisc_LostFocus", intEL, strES, sql


End Sub

Private Sub cmbMSC_Click()

10        If txtMSC() = "Medical Scientist Comments" Then
20            txtMSC = ""
30        End If

40        txtMSC = txtMSC & cmbMSC
50        txtMSC.SetFocus
60        txtMSC.SelStart = Len(txtMSC)
70        cmbMSC = ""
80        cmbMSC.Visible = False

90        cmdSaveMicro.Enabled = True
100       cmdSaveHold.Enabled = True

End Sub


Private Sub cmbMSC_LostFocus()

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo cmbMSC_LostFocus_Error

20        If cmbMSC <> "" Then
30            sql = "SELECT * FROM Lists WHERE " & _
                    "ListType = 'BA' " & _
                    "AND Code = '" & AddTicks(cmbMSC) & "'"
40            Set tb = New Recordset
50            RecOpenServer 0, tb, sql
60            If Not tb.EOF Then
70                cmbMSC = tb!Text & ""
80            End If
90        End If

100       If txtMSC() = "Medical Scientist Comments" Then
110           txtMSC = cmbMSC
120       Else
130           txtMSC = txtMSC & cmbMSC
140       End If

150       cmbMSC.Visible = False
160       cmbMSC = ""

170       Exit Sub

cmbMSC_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

180       intEL = Erl
190       strES = Err.Description
200       LogError "frmEditMicrobiologyNew", "cmbMSC_LostFocus", intEL, strES, sql


End Sub


Private Sub cmbOrgGroup_Click(Index As Integer)

10        FillAbGrid Index
20        FillABSelect Index
30        FillOrgNames Index

40        cmdSaveMicro.Enabled = True
50        cmdSaveHold.Enabled = True
60        grdAB(Index).Visible = True

End Sub

Private Sub cmbOrgGroup_LostFocus(Index As Integer)

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo cmbOrgGroup_LostFocus_Error

20        sql = "Select * from Lists where " & _
                "ListType = 'OR' " & _
                "and Code = '" & cmbOrgGroup(Index) & "'"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        If Not tb.EOF Then
60            cmbOrgGroup(Index) = tb!Text & ""
70        End If

80        Exit Sub

cmbOrgGroup_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

90        intEL = Erl
100       strES = Err.Description
110       LogError "frmEditMicrobiologyNew", "cmbOrgGroup_LostFocus", intEL, strES, sql


End Sub


Private Sub cmbOrgName_Click(Index As Integer)

10        cmdSaveMicro.Enabled = True
20        cmdSaveHold.Enabled = True

End Sub

Private Sub cmbOrgName_LostFocus(Index As Integer)

          Dim tb As Recordset
          Dim sql As String
          Dim MicroEColiComment As String

          Dim Obs As New Observations

10        On Error GoTo cmbOrgName_LostFocus_Error

20        sql = "SELECT Name FROM Organisms WHERE " & _
                "Code = '" & AddTicks(cmbOrgName(Index)) & "'"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        If Not tb.EOF Then
60            cmbOrgName(Index) = tb!Name & ""
70        End If

80        MicroEColiComment = GetOptionSetting("MicrobiologyEColi0157Comment", "")
90        If MicroEColiComment <> "" And _
             InStr(UCase$(cmbOrgName(Index)), "COLI") > 0 And _
             InStr(cmbOrgName(Index), "157") > 0 And _
             InStr(txtMSC, MicroEColiComment) = 0 Then

100           Obs.Save SampleIDWithOffset, False, "MicroCS", MicroEColiComment

110           LoadComments

120       End If

130       Exit Sub

cmbOrgName_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

140       intEL = Erl
150       strES = Err.Description
160       LogError "frmEditMicrobiologyNew", "cmbOrgName_LostFocus", intEL, strES, sql

End Sub


Private Sub cmbOva_Click(Index As Integer)
10        cmdSaveHold.Enabled = True
20        cmdSaveMicro.Enabled = True

30        ShowUnlock 11

End Sub

Private Sub cmbOva_KeyPress(Index As Integer, KeyAscii As Integer)

10        KeyAscii = 0

End Sub


Private Sub cmbQualifier_Click(Index As Integer)

10        cmdSaveMicro.Enabled = True
20        cmdSaveHold.Enabled = True

End Sub


Private Sub cmbSite_Change()

10        lblSiteDetails = cmbSite & " " & txtSiteDetails

20        cmdOrderTests.Enabled = False
30        If InStr(1, cmbSite, "Faeces") > 0 Then
40            cmdOrderTests.Enabled = True
50        ElseIf cmbSite = "Urine" Then
60            cmdOrderTests.Enabled = True
70        Else
80            cmdOrderTests.Enabled = True
90        End If

100       cmbSiteSearch = cmbSite

110       If InStr(1, UCase(cmbSite), "MRSA") > 0 Then
120           txtNoCopies = 2
130       Else
140           txtNoCopies = 1
150       End If

End Sub

Private Sub cmbSite_LostFocus()

          Dim Found As Boolean
          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo cmbSite_LostFocus_Error

20        If Trim$(cmbSite) = "" Then
30            Exit Sub
40        End If

50        Found = False
60        sql = "SELECT COUNT(*) Tot FROM Lists WHERE " & _
                "ListType = 'SI' " & _
                "AND ( Text = '" & cmbSite & "' " & _
                "      OR Code = '" & cmbSite & "')"
70        Set tb = New Recordset
80        RecOpenServer 0, tb, sql
90        If tb!Tot = 0 Then
100           cmbSite = ""
110           txtSiteDetails = ""
120       End If

130       cmbSiteSearch = cmbSite

140       Exit Sub

cmbSite_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

150       intEL = Erl
160       strES = Err.Description
170       LogError "frmEditMicrobiologyNew", "cmbSite_LostFocus", intEL, strES, sql

End Sub

Private Sub cmbUrineComment_LostFocus()

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo cmbUrineComment_LostFocus_Error

20        sql = "Select * from Lists where " & _
                "ListType = 'UC' " & _
                "and Code = '" & cmbUrineComment & "'"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        If Not tb.EOF Then
60            txtUrineComment = Trim$(txtUrineComment & " " & tb!Text & "")
70        Else
80            txtUrineComment = Trim$(txtUrineComment & " " & cmbUrineComment)
90        End If
100       cmbUrineComment = ""

110       cmdSaveMicro.Enabled = True
120       cmdSaveHold.Enabled = True

130       Exit Sub

cmbUrineComment_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

140       intEL = Erl
150       strES = Err.Description
160       LogError "frmEditMicrobiologyNew", "cmbUrineComment_LostFocus", intEL, strES, sql

End Sub


Private Sub cmbWetPrep_Click(Index As Integer)

10        cmdSaveMicro.Enabled = True
20        cmdSaveHold.Enabled = True

End Sub


Private Sub cmbWetPrep_LostFocus(Index As Integer)

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo cmbWetPrep_LostFocus_Error

20        sql = "Select * from Lists where " & _
                "ListType = 'WP' " & _
                "and Code = '" & cmbWetPrep(Index) & "'"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        If Not tb.EOF Then
60            cmbWetPrep(Index) = tb!Text & ""
70        End If

80        Exit Sub

cmbWetPrep_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

90        intEL = Erl
100       strES = Err.Description
110       LogError "frmEditMicrobiologyNew", "cmbWetPrep_LostFocus", intEL, strES, sql


End Sub


Private Sub cmdABsInUse_Click()

10        lstABsInUse.Visible = False
20        cmbABsInUse.Visible = True
30        cmbABsInUse.SetFocus

40        cmdSaveDemographics.Enabled = True
50        cmdSaveInc.Enabled = True

End Sub

Private Sub cmdConC_Click()

10        cmbConC.Visible = True
20        cmbConC.SetFocus

End Sub

Private Sub cmdCopyTo_Click()

          Dim s As String

10        s = cmbWard & " " & cmbClinician
20        s = Trim$(s) & " " & cmbGP
30        s = Trim$(s)

40        frmCopyTo.EditScreen = Me
50        frmCopyTo.lblOriginal = s
60        frmCopyTo.lblSampleID = txtSampleID + SysOptMicroOffset(0)
70        frmCopyTo.Show 1

80        CheckCC

End Sub

Private Sub CheckCC()

          Dim sql As String
          Dim tb As Recordset

10        On Error GoTo CheckCC_Error

20        cmdCopyTo.Caption = "cc"
30        cmdCopyTo.Font.Bold = False
40        cmdCopyTo.BackColor = &H8000000F

50        If Trim$(txtSampleID) = "" Then Exit Sub

60        sql = "Select * from SendCopyTo where " & _
                "SampleID = '" & SysOptMicroOffset(0) + Val(txtSampleID) & "'"
70        Set tb = New Recordset
80        RecOpenServer 0, tb, sql
90        If Not tb.EOF Then
100           cmdCopyTo.Caption = "++ cc ++"
110           cmdCopyTo.Font.Bold = True
120           cmdCopyTo.BackColor = &H8080FF
130       End If

140       Exit Sub

CheckCC_Error:

          Dim strES As String
          Dim intEL As Integer

150       intEL = Erl
160       strES = Err.Description
170       LogError "frmEditMicrobiology", "CheckCC", intEL, strES, sql

End Sub


Private Sub cmdDartViewer_Click()

10        On Error GoTo cmdDartViewer_Click_Error

20        Shell "C:\Program Files\The PlumTree Group\Dartviewer\Dartviewer.exe " & Format(txtSampleID, "000000"), vbNormalFocus

30        Exit Sub

cmdDartViewer_Click_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmEditMicrobiologyNew", "cmdDartViewer_Click", intEL, strES

End Sub

Private Sub cmdLock_Click(Index As Integer)

          Dim LockIt As Boolean

10        LockIt = cmdLock(Index).Caption = "&Lock Result"
20        UpdateLockStatus Val(txtSampleID) + SysOptMicroOffset(0), LockIt, Index

End Sub

Private Sub cmdMSC_Click()

10        cmbMSC.Visible = True
20        cmbMSC.SetFocus

End Sub

Private Sub cmdNADMicro_Click()

10        txtBacteria = "Nil"
20        txtWCC = "Nil"
30        txtRCC = "Nil"
40        cmbCrystals = "None seen"
50        cmbCasts = "None seen"
60        cmbMisc(0) = "-"

70        ShowUnlock 1

End Sub

Private Sub cmdPatientNotePad_Click()
10    On Error GoTo cmdPatientNotePad_Click_Error

20    frmPatientNotePad.SampleID = txtSampleID
30    frmPatientNotePad.Caller = "Microbiology"
40    frmPatientNotePad.Show 1

50    Exit Sub
cmdPatientNotePad_Click_Error:
         
60    LogError "frmEditMicrobiologyNew", "cmdPatientNotePad_Click", Erl, Err.Description


End Sub

Private Sub cmdPhone_Click()

10        With frmPhoneLog
20            .SampleID = Val(txtSampleID) + SysOptMicroOffset(0)
30            .Caller = "Micro"
40            If cmbGP <> "" Then
50                .GP = cmbGP
60                .WardOrGP = "GP"
70            Else
80                .GP = cmbWard
90                .WardOrGP = "Ward"
100           End If
110           .Show 1
120       End With

130       CheckIfPhoned

End Sub

Private Sub CheckIfPhoned()

10        On Error GoTo CheckIfPhoned_Error

20        If CheckPhoneLog(txtSampleID + SysOptMicroOffset(0)) Then
30            cmdPhone.BackColor = vbYellow
40            cmdPhone.Caption = "Results Phoned"
50            cmdPhone.ToolTipText = "Results Phoned"
60        Else
70            cmdPhone.BackColor = &H8000000F
80            cmdPhone.Caption = "Phone Results"
90            cmdPhone.ToolTipText = "Phone Results"
100       End If

110       Exit Sub

CheckIfPhoned_Error:

          Dim strES As String
          Dim intEL As Integer

120       intEL = Erl
130       strES = Err.Description
140       LogError "frmEditMicrobiologyNew", "CheckIfPhoned", intEL, strES

End Sub

'Private Sub CheckIfPhoned()
'
'      Dim s As String
'      Dim PhLog As PhoneLog
'      Dim sql As String
'      Dim tb As Recordset
'      Dim OBS As New Observations
'      Dim Title As String
'
'10    On Error GoTo CheckIfPhoned_Error
'
'20    PhLog = CheckPhoneLog1(Val(txtSampleID) + SysOptMicroOffset(0))
'30    If PhLog.SampleID <> 0 Then
'40        Title = Trim(PhLog.Title & " " & PhLog.PersonName)
'50        If Title <> "" Then Title = " (" & Title & ") "
'60        cmdPhone.BackColor = vbYellow
'70        cmdPhone.Caption = "Results Phoned"
'80        cmdPhone.ToolTipText = "Results Phoned"
'90        If InStr(txtDemographicComment.Text, "Results Phoned") = 0 Then
'100           s = "Results Phoned to " & PhLog.PhonedTo & Title & PhLog.PersonName & " at " & _
 '                  Format$(PhLog.Datetime, "hh:mm") & " on " & Format$(PhLog.Datetime, "dd/MM/yyyy") & _
 '                  " by " & PhLog.PhonedBy & "."
'110           If Trim$(txtDemographicComment.Text) = "" Then
'120               txtDemographicComment.Text = s
'130           Else
'140               txtDemographicComment.Text = txtDemographicComment.Text & ". " & s
'150           End If
'
'160           OBS.Save PhLog.SampleID, True, "Demographic", txtDemographicComment
'
'170       End If
'180   Else
'190       cmdPhone.BackColor = &H8000000F
'200       cmdPhone.Caption = "Phone Results"
'210       cmdPhone.ToolTipText = "Phone Results"
'220   End If
'
'230   Exit Sub
'
'CheckIfPhoned_Error:
'
'      Dim strES As String
'      Dim intEL As Integer
'
'240   intEL = Erl
'250   strES = Err.Description
'260   LogError "frmEditMicrobiologyNew", "CheckIfPhoned", intEL, strES, sql
'
'End Sub



Private Sub cmdPrintBarcode_Click()
10    On Error GoTo cmdPrintBarcode_Click_Error


20    If Trim$(txtSampleID) = "" Or Trim(txtName) = "" Then
30            Exit Sub
40        End If

50        With frmPrintBarcodeLabel

60            .lblDepartment = "Microbiology"
70            .lblSampleID.Caption = txtSampleID.Text
80            .lblPatName.Caption = txtName
90            .lblSampleDate.Caption = Format(dtSampleDate & " " & tSampleTime, "dd/MM/yyyy HH:mm:ss")
100           .lblAgeSexDoB = "A/S/DOB: " & _
                          IIf(txtAge = "", "", Left(txtAge, Len(txtAge) - 2)) & " " & _
                          IIf(txtSex = "", "", Left(txtSex, 1)) & " " & _
                          IIf(txtDoB = "", "", Format(txtDoB, "yyyyMMdd"))
110           .Show 1
120       End With

130   Exit Sub

cmdPrintBarcode_Click_Error:
         
140   LogError "frmEditMicrobiologyNew", "cmdPrintBarcode_Click", Erl, Err.Description
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdReleaseReport_Click
' Author    : XPMUser
' Date      : 2/19/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmdReleaseReport_Click()
      Dim sql As String
      'Dim tb As Recordset
      Dim pSampleID As String
10    On Error GoTo cmdReleaseReport_Click_Error


'20    If cmdReleasetoWard.BackColor = vbGreen Then
'30        iMsg "Report is released to the ward, cannot be released to consultant"
'40        Exit Sub
'50    End If

60    frmValidateAll.SampleIDToValidate = SampleIDWithOffset
70    frmValidateAll.Show 1


80    If cmdReleaseReport.BackColor = vbCyan Then
90        If iMsg("Do you want to re-release report to Consultant", vbQuestion + vbYesNo, , vbRed) = vbYes Then
100           UpdateConsultantList Val(txtSampleID) + SysOptMicroOffset(0), "Micro", ReleasedToConsultant, 0, 0
110           pSampleID = SysOptMicroOffset(0) + Val(txtSampleID)
120           If SampleAddedtoConsultantList(pSampleID, "Micro") Then
130               RemoveReport 0, pSampleID, "N", 0
140               Call PrintThis("SaveTemp")
150               cmdReleaseReport.BackColor = vbCyan
160               cmdReleaseReport.Caption = "Release to Consultant"
170           End If
180       End If
190   Else
200       UpdateConsultantList Val(txtSampleID) + SysOptMicroOffset(0), "Micro", ReleasedToConsultant, 0, 0
210       pSampleID = SysOptMicroOffset(0) + Val(txtSampleID)
220       If SampleAddedtoConsultantList(pSampleID, "Micro") Then
230           RemoveReport 0, pSampleID, "N", 0
240           Call PrintThis("SaveTemp")
250           cmdReleaseReport.BackColor = vbCyan
260           cmdReleaseReport.Caption = "Release to Consultant"
270       End If
280   End If
290   txtSampleID = Format$(Val(txtSampleID) + 1)
300   GetSampleIDWithOffset
310   LoadAllDetails

320   Exit Sub


cmdReleaseReport_Click_Error:

      Dim strES As String
      Dim intEL As Integer

330   intEL = Erl
340   strES = Err.Description
350   LogError "frmEditMicrobiologyNew", "cmdReleaseReport_Click", intEL, strES
End Sub

Private Sub cmdReleasetoWard_Click()

      Dim SID As String
      Dim pSampleID As String
      Dim sql As String

10    On Error GoTo cmdReleasetoWard_Click_Error

20    If cmbWard = "GP" Then
          'release to healthlink
30        If cmdReleaseReport.BackColor = vbCyan Then
40            iMsg "Report is released to consultant, cannot be released to healthlink until authorised"
50            Exit Sub
60        End If

70        SID = Format$(Val(txtSampleID) + SysOptMicroOffset(0))

80        With cmdHealthLink
90            If .Picture = imgHGreen.Picture Then
100               Set .Picture = imgHRed.Picture
110               ReleaseMicro SID, 0
120           Else
130               If lblInterim.BackColor = vbGreen Then
140                   ReleaseMicro SID, 1
150               Else
160                   ReleaseMicro SID, 2
170               End If
180               Set .Picture = imgHGreen.Picture

190           End If
200       End With


210   End If

220   pSampleID = SysOptMicroOffset(0) + Val(txtSampleID)
230   If lblFinal.BackColor = vbGreen Then
240       If SampleRelasedtoConsultant(pSampleID, "Micro") Then
250           iMsg "This report is being reviewed by consultant and cannot be released to the ward as a final report.", vbInformation
260           Exit Sub
270       End If
280   End If
290   frmValidateAll.SampleIDToValidate = SampleIDWithOffset
300   frmValidateAll.Show 1
310   If cmdReleasetoWard.BackColor = vbGreen Then
320       If iMsg("Do you want to rerelease report to ward" & vbCrLf & vbCrLf & _
                  "*** Do you need to  release this report to Consultants Queue? ***", vbQuestion + vbYesNo, , vbRed) = vbYes Then



330           Call PrintThis("SaveFinal")
340           cmdReleasetoWard.BackColor = vbGreen
              'cmdReleasetoWard.Caption = "Release to Ward"
350       End If
360   Else

370       Call PrintThis("SaveFinal")
380       UpdateConsultantList Val(txtSampleID) + SysOptMicroOffset(0), "Micro", ReleasedToWard, 0, 0

390       cmdReleasetoWard.BackColor = vbGreen
          'cmdReleasetoWard.Caption = "Release to Ward"
400   End If

410   Exit Sub


cmdReleasetoWard_Click_Error:

      Dim strES As String
      Dim intEL As Integer

420   intEL = Erl
430   strES = Err.Description
440   LogError "frmEditMicrobiologyNew", "cmdReleaseToWard_Click", intEL, strES, sql
End Sub

Private Sub cmdRemoveSecondary_Click(Index As Integer)

          Dim tb As Recordset
          Dim sql As String
          Dim n As Integer
          Dim Found As Boolean
          Dim ABName As String
          Dim intABs As Integer

10        On Error GoTo cmdRemoveSecondary_Click_Error

20        grdAB(Index).Col = 0
30        For n = grdAB(Index).Rows - 1 To 1 Step -1
40            grdAB(Index).Row = n
50            If grdAB(Index).CellFontBold = True Then
60                DeleteSensitivity Index, grdAB(Index).TextMatrix(n, 0)
70                If n = 1 Then
80                    grdAB(Index).AddItem ""
90                End If
100               grdAB(Index).RemoveItem n
110           End If
120       Next

130       FillABSelect Index

140       Exit Sub

150       sql = "Select Distinct A.AntibioticName, D.ListOrder, " & _
                "A.AllowIfPregnant, A.AllowIfOutPatient, A.AllowIfChild " & _
                "from ABDefinitions as D, Antibiotics as A where " & _
                "D.OrganismGroup = '" & cmbOrgGroup(Index) & "' " & _
                "and D.Site = '" & cmbSite & "' " & _
                "and D.PriSec = 'S' " & _
                "and D.AntibioticName = A.AntibioticName " & _
                "order by D.ListOrder"
160       Set tb = New Recordset
170       RecOpenServer 0, tb, sql
180       If tb.EOF Then
190           sql = "Select Distinct A.AntibioticName, D.ListOrder, " & _
                    "A.AllowIfPregnant, A.AllowIfOutPatient, A.AllowIfChild " & _
                    "from ABDefinitions as D, Antibiotics as A where " & _
                    "D.OrganismGroup = '" & cmbOrgGroup(Index) & "' " & _
                    "and D.PriSec = 'S' and D.Site = 'Generic' " & _
                    "and D.AntibioticName = A.AntibioticName " & _
                    "order by D.ListOrder"
200           Set tb = New Recordset
210           RecOpenServer 0, tb, sql
220           If tb.EOF Then
230               Exit Sub
240           End If
250       End If
260       Do While Not tb.EOF

270           Found = False
280           ABName = Trim$(tb!AntibioticName & "")
290           For n = 1 To grdAB(Index).Rows - 1
300               If Trim$(grdAB(Index).TextMatrix(n, 0)) = ABName Then
310                   Found = True
320                   For intABs = 0 To lstABsInUse.ListCount - 1
330                       If lstABsInUse.List(intABs) = ABName Then
340                           Found = False
350                       End If
360                   Next
370                   Exit For
380               End If
390           Next

400           If Found Then
410               If grdAB(Index).Rows = 2 Then
420                   grdAB(Index).AddItem ""
430               End If
440               grdAB(Index).RemoveItem n
450           End If

460           tb.MoveNext
470       Loop

480       Exit Sub

cmdRemoveSecondary_Click_Error:

          Dim strES As String
          Dim intEL As Integer

490       intEL = Erl
500       strES = Err.Description
510       LogError "frmEditMicrobiology", "cmdRemoveSecondary_Click", intEL, strES, sql

End Sub

Private Sub cmdReportAll_Click(Index As Integer)

          Dim n As Integer

10        With grdAB(Index)
20            .Col = 2
30            For n = 1 To .Rows - 1
40                If .TextMatrix(n, 0) <> "" Then
50                    .Row = n
60                    If Trim$(.TextMatrix(n, 1)) = "" Then
70                        Set .CellPicture = imgSquareCross.Picture
80                    Else
90                        Set .CellPicture = imgSquareTick.Picture
100                   End If
110               End If
120           Next
130       End With

140       cmdSaveHold.Enabled = True
150       cmdSaveMicro.Enabled = True

End Sub

Private Sub cmdReportNone_Click(Index As Integer)

          Dim n As Integer

10        With grdAB(Index)
20            .Col = 2
30            For n = 1 To .Rows - 1
40                If .TextMatrix(n, 0) <> "" Then
50                    .Row = n
60                    Set .CellPicture = imgSquareCross.Picture
70                End If
80            Next
90        End With

100       cmdSaveHold.Enabled = True
110       cmdSaveMicro.Enabled = True

End Sub


Private Sub cmdSaveHold_Click()

10        SaveMicro

End Sub

Private Sub cmdUseSecondary_Click(Index As Integer)

          Dim tb As Recordset
          Dim sql As String
          Dim n As Integer
          Dim Found As Boolean
          Dim ABName As String
          Dim ABCode As String
          Dim tbC As Recordset
          Dim Res As String
          Dim RSI As String
          Dim RunDateTime As String
          Dim Operator As String

10        On Error GoTo cmdUseSecondary_Click_Error

20        sql = "Select Distinct A.AntibioticName, D.ListOrder, " & _
                "A.AllowIfPregnant, A.AllowIfOutPatient, A.AllowIfChild " & _
                "from ABDefinitions as D, Antibiotics as A where " & _
                "D.OrganismGroup = '" & cmbOrgGroup(Index) & "' " & _
                "and D.Site = '" & cmbSite & "' " & _
                "and D.PriSec = 'S' " & _
                "and D.AntibioticName = A.AntibioticName " & _
                "order by D.ListOrder"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        If tb.EOF Then
60            sql = "Select Distinct A.AntibioticName, D.ListOrder, " & _
                    "A.AllowIfPregnant, A.AllowIfOutPatient, A.AllowIfChild " & _
                    "from ABDefinitions as D, Antibiotics as A where " & _
                    "D.OrganismGroup = '" & cmbOrgGroup(Index) & "' " & _
                    "and (D.Site = 'Generic' or D.Site is Null ) and D.PriSec = 'S' " & _
                    "and D.AntibioticName = A.AntibioticName " & _
                    "order by D.ListOrder"
70            Set tb = New Recordset
80            RecOpenServer 0, tb, sql
90            If tb.EOF Then
100               Exit Sub
110           End If
120       End If
130       Do While Not tb.EOF

140           Found = False
150           ABName = Trim$(tb!AntibioticName & "")
160           ABCode = AntibioticCodeFor(ABName)
170           sql = "Select * from Sensitivities where " & _
                    "SampleID = '" & SysOptMicroOffset(0) + txtSampleID & "' " & _
                    "and IsolateNumber = '" & Index & "' " & _
                    "and AntibioticCode = '" & ABCode & "'"
180           Set tbC = New Recordset
190           RecOpenServer 0, tbC, sql
200           If Not tbC.EOF Then
210               RSI = tbC!RSI & ""
220               Res = tbC!Result & ""
230               RunDateTime = Format(tbC!RunDateTime, "dd/mm/yy hh:mm")
240               Operator = tbC!UserName & ""
250           Else
260               RSI = ""
270               Res = ""
280               RunDateTime = ""
290               Operator = ""
300           End If

310           For n = 1 To grdAB(Index).Rows - 1
320               If Trim$(grdAB(Index).TextMatrix(n, 0)) = ABName Then
330                   Found = True
340                   Exit For
350               End If
360           Next

370           If Not Found Then
380               grdAB(Index).AddItem ABName & vbTab & _
                                       RSI & vbTab & _
                                       vbTab & _
                                       Res & vbTab & _
                                       RunDateTime & vbTab & Operator
390               grdAB(Index).Row = grdAB(Index).Rows - 1
400               grdAB(Index).Col = 0
410               grdAB(Index).CellFontBold = True
420               grdAB(Index).Col = 2
430               If IsChild() And Not tb!AllowIfChild Then
440                   Set grdAB(Index).CellPicture = imgSquareCross.Picture
450                   grdAB(Index).TextMatrix(grdAB(Index).Row, 2) = "C"
460               ElseIf IsPregnant() And Not tb!AllowIfPregnant Then
470                   Set grdAB(Index).CellPicture = imgSquareCross.Picture
480                   grdAB(Index).TextMatrix(grdAB(Index).Row, 2) = "P"
490               ElseIf IsOutPatient() And Not tb!AllowIfOutPatient Then
500                   Set grdAB(Index).CellPicture = imgSquareCross.Picture
510                   grdAB(Index).TextMatrix(grdAB(Index).Row, 2) = "O"
520               Else
530                   Set grdAB(Index).CellPicture = imgSquareCross.Picture
540               End If
550           End If

560           tb.MoveNext
570       Loop

580       FillABSelect Index

590       cmdSaveMicro.Enabled = True
600       cmdSaveHold.Enabled = True

610       Exit Sub

cmdUseSecondary_Click_Error:

          Dim strES As String
          Dim intEL As Integer

620       intEL = Erl
630       strES = Err.Description
640       LogError "frmEditMicrobiology", "cmdUseSecondary_Click", intEL, strES, sql

End Sub

Private Sub cmdViewExternal_Click()

10        With frmEditMicroExternals
20            .txtSampleID = txtSampleID
30            .lblSampleDate = dtSampleDate
              
40            .Show 1
50        End With

60        CheckExternals

End Sub

Private Sub cmdViewReports_Click()

10        On Error GoTo cmdViewReports_Click_Error

20        frmRFT.SampleID = Val(txtSampleID) + SysOptMicroOffset(0)
30        frmRFT.Dept = "N"
40        frmRFT.Show 1

50        Exit Sub

cmdViewReports_Click_Error:

          Dim strES As String
          Dim intEL As Integer

60        intEL = Erl
70        strES = Err.Description
80        LogError "frmEditMicrobiologyNew", "cmdViewReports_Click", intEL, strES


End Sub

Private Sub cmdVitek_Click(Index As Integer)

          Dim tb As Recordset
          Dim sql As String
          Dim ANF As String

10        On Error GoTo cmdVitek_Click_Error

20        If txtName = "" Or txtDoB = "" Then
30            iMsg "Patient name and date of birth must be provided before ordering to vitek", vbInformation
40            Exit Sub
50        End If

60        Select Case UCase$(gBC.TextMatrix(Index + 1, 2))
          Case GetOptionSetting("BcAerobicBottle", "BSA"): ANF = "A"
70        Case GetOptionSetting("BcAnarobicBottle", "BSN"): ANF = "N"
80        Case GetOptionSetting("BcFanBottle", "BFA"): ANF = "F"
90        Case Else: Exit Sub
100       End Select

          'Created on 08/10/2010 12:01:03
          'Autogenerated by SQL Scripting

110       sql = "IF EXISTS (SELECT 1 FROM BactOrders " & _
                "           WHERE SampleID = @SampleID0 " & _
                "           AND TestRequested = '@TestRequested2' ) " & _
                "  UPDATE BactOrders " & _
                "  SET SampleID = @SampleID0, " & _
                "  Analyser = '@Analyser1', " & _
                "  TestRequested = '@TestRequested2', " & _
                "  Programmed = @Programmed3, " & _
                "  DateTimeOfRecord = getdate() " & _
                "  WHERE SampleID = @SampleID0 " & _
                "  AND TestRequested = '@TestRequested2' " & _
                "ELSE " & _
                "  INSERT INTO BactOrders " & _
                "  (SampleID, Analyser, TestRequested, Programmed, DateTimeOfRecord) VALUES " & _
                "  (@SampleID0, '@Analyser1', '@TestRequested2', @Programmed3, getdate()) "

120       sql = Replace(sql, "@SampleID0", txtSampleID)
130       sql = Replace(sql, "@Analyser1", "Observa")
140       sql = Replace(sql, "@TestRequested2", ANF)
150       sql = Replace(sql, "@Programmed3", 0)

160       Cnxn(0).Execute sql

          'sql = "SELECT * FROM BactOrders WHERE " & _
           '      "SampleID = '" & txtSampleID & "' " & _
           '      "AND TestRequested = '" & ANF & "'"
          '
          'Set tb = New Recordset
          'RecOpenServer 0, tb, sql
          'If tb.EOF Then
          '    tb.AddNew
          'End If
          'tb!SampleID = txtSampleID
          'tb!Analyser = "Observa"
          'tb!TestRequested = ANF
          'tb!Programmed = 0
          'tb!DateTimeOfRecord = Now
          'tb.Update

170       cmdVitek(Index).Caption = "Requested"
180       cmdVitek(Index).Enabled = False

190       Exit Sub

cmdVitek_Click_Error:

          Dim strES As String
          Dim intEL As Integer

200       intEL = Erl
210       strES = Err.Description
220       LogError "frmEditMicrobiologyNew", "cmdVitek_Click", intEL, strES, sql

End Sub

Private Sub cmdWardWarning_Click()

10    On Error GoTo cmdWardWarning_Click_Error

20    MsgBox "You have a differnt version of micro report in the ward. " & vbCrLf & _
              "Please re-release report", vbExclamation, "NetAcquire"

30    Exit Sub
cmdWardWarning_Click_Error:
         
40    LogError "frmEditMicrobiologyNew", "cmdWardWarning_Click", Erl, Err.Description
End Sub

Private Sub cMRU_GotFocus()

10        On Error GoTo cMRU_GotFocus_Error

20        If cmdSaveDemographics.Enabled Or cmdSaveInc.Enabled Then
30            If iMsg("Save Details", vbQuestion + vbYesNo) = vbYes Then
40                GetSampleIDWithOffset
50                SaveDemographics
60                cmdSaveDemographics.Enabled = False
70                cmdSaveInc.Enabled = False
80            End If
90        End If

100       Exit Sub

cMRU_GotFocus_Error:

          Dim strES As String
          Dim intEL As Integer

110       intEL = Erl
120       strES = Err.Description
130       LogError "frmEditMicrobiologyNew", "cMRU_GotFocus", intEL, strES


End Sub

Private Sub cmdAddToConsultantList_Click()

          Dim sql As String
          Dim tb As Recordset

10        On Error GoTo cmdAddToConsultantList_Click_Error


20        Select Case Left$(cmdAddToConsultantList.Caption, 3)
          Case "Add":

              '30            sql = "If Exists(Select 1 From ConsultantList " & _
               '                    "Where SampleID = @SampleID ) " & _
               '                    "Begin " & _
               '                    "Update ConsultantList Set " & _
               '                    "SampleID = @SampleID " & _
               '                    "Where SampleID = @SampleID  " & _
               '                    "End  " & _
               '                    "Else " & _
               '                    "Begin  " & _
               '                    "Insert Into ConsultantList (SampleID) Values (@SampleID) " & _
               '                    "End"


30            sql = "If Exists(Select 1 From ConsultantList " & _
                    "Where SampleID = @SampleID ) " & _
                    "Begin " & _
                    "Update ConsultantList Set " & _
                    "SampleID = @SampleID, " & _
                    "Department = 'Micro', " & _
                    "Status  = '0', " & _
                    "Username = '" & UserName & "'" & _
                    " Where SampleID = @SampleID  " & _
                    "End  " & _
                    "Else " & _
                    "Begin  " & _
                    "Insert Into ConsultantList (SampleID,Department,Status,Username) Values (@SampleID,'Micro','0','" & UserName & "') " & _
                    "End"


40            sql = Replace(sql, "@SampleID", SysOptMicroOffset(0) + Val(txtSampleID))

50            Cnxn(0).Execute sql
60            cmdAddToConsultantList.Caption = "Remove from Consultant List"

70            If cmdReleaseReport.BackColor = vbGreen Then
80                RemoveReport 0, SysOptMicroOffset(0) + Val(txtSampleID), "N", 1
90                Call PrintThis("SaveTemp")
100               cmdReleaseReport.BackColor = vbCyan
110               cmdReleaseReport.Caption = "Release to Consultant"
120           End If
130       Case "Rem":
140           sql = "Delete from ConsultantList " & _
                    "where SampleID = '" & SysOptMicroOffset(0) + Val(txtSampleID) & "'"
150           Cnxn(0).Execute sql
160           cmdAddToConsultantList.Caption = "Add to Consultant List"
170           If cmdReleaseReport.BackColor = vbCyan Then
180               RemoveReport 0, SysOptMicroOffset(0) + Val(txtSampleID), "N", 0
190               cmdReleaseReport.BackColor = vbButtonFace
200               cmdReleaseReport.Caption = "Release to Consultant"
210           End If

220       End Select

          'FillForConsultantValidation

230       Exit Sub

cmdAddToConsultantList_Click_Error:

          Dim strES As String
          Dim intEL As Integer

240       intEL = Erl
250       strES = Err.Description
260       LogError "frmEditMicrobiologyNew", "cmdAddToConsultantList_Click", intEL, strES, sql


End Sub





Private Sub cRooH_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)

10        cmdSaveDemographics.Enabled = True
20        cmdSaveInc.Enabled = True

End Sub

Private Sub dtRecDate_CloseUp()

10        pBar = 0

20        cmdSaveDemographics.Enabled = True
30        cmdSaveInc.Enabled = True

End Sub


Private Sub dtRecDate_LostFocus()

10        SetDatesColour Me

End Sub

Private Sub dtRunDate_LostFocus()

10        SetDatesColour Me

End Sub

Private Sub dtSampleDate_LostFocus()

10        SetDatesColour Me

End Sub

Private Sub fraDate_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

10        pBar = 0

End Sub

Private Sub grdAB_Click(Index As Integer)

          Dim s As String
          Dim RSI As Boolean

10        On Error GoTo grdAB_Click_Error

20        If Not fraCS.Enabled Then Exit Sub

30        cmdSaveMicro.Enabled = True
40        cmdSaveHold.Enabled = True

50        With grdAB(Index)
60            If .MouseRow = 0 Then Exit Sub

70            If .CellBackColor = &HFFFFC0 Then
80                .Enabled = False
90                If iMsg("Remove " & Trim$(.Text) & " from List?", vbQuestion + vbYesNo) = vbYes Then
100                   DeleteSensitivity Index, .TextMatrix(.Row, 0)
110                   .RemoveItem .Row
120                   FillABSelect Index
130               End If
140               .Enabled = True
150           ElseIf .Col = 1 Then
160               s = Trim$(.TextMatrix(.Row, 1))
170               Select Case s
                  Case "": s = "R": RSI = True
180               Case "R": s = "S": RSI = True
190               Case "S": s = "I": RSI = True
200               Case "I": s = "": RSI = False
210               Case Else: s = "": RSI = False
220               End Select
230               .TextMatrix(.Row, 1) = s
240               If cmbOrgName(Index) = "Staphylococcus aureus" And UCase(.TextMatrix(.Row, 0)) = "OXACILLIN" And s = "R" Then
250                   cmbOrgName(Index) = "Staphylococcus aureus (MRSA)"
260               ElseIf cmbOrgName(Index) = "Staphylococcus aureus (MRSA)" And UCase(.TextMatrix(.Row, 0)) = "OXACILLIN" And s <> "R" Then
270                   cmbOrgName(Index) = "Staphylococcus aureus"
280               End If
290               .Col = 2
300               If RSI Then
310                   If AutoReportAB(Index) Then
320                       Set .CellPicture = imgSquareTick.Picture
330                   Else
340                       Set .CellPicture = imgSquareCross.Picture
350                   End If
360               Else
370                   Set .CellPicture = Nothing
380               End If
390           ElseIf .Col = 2 Then
400               If .CellPicture = imgSquareTick.Picture Then
410                   Set .CellPicture = imgSquareCross.Picture
420                   SetAsForced Index, .TextMatrix(.Row, 0), False
430               Else
440                   If .TextMatrix(.Row, 2) = "C" Then
450                       If MsgBox("Report " & .TextMatrix(.Row, 0) & " on a Child?", vbQuestion + vbYesNo) = vbNo Then
460                           Exit Sub
470                       End If
480                   ElseIf .TextMatrix(.Row, 2) = "P" Then
490                       If MsgBox("Report " & .TextMatrix(.Row, 0) & " for Pregnant Patient?", vbQuestion + vbYesNo) = vbNo Then
500                           Exit Sub
510                       End If
520                   ElseIf .TextMatrix(.Row, 2) = "O" Then
530                       If MsgBox("Report " & .TextMatrix(.Row, 0) & " for an Out-Patient?", vbQuestion + vbYesNo) = vbNo Then
540                           Exit Sub
550                       End If
560                   End If
570                   Set .CellPicture = imgSquareTick.Picture
580                   SetAsForced Index, .TextMatrix(.Row, 0), True
590               End If
600           End If


610           .LeftCol = 0

620       End With

630       Exit Sub

grdAB_Click_Error:

          Dim strES As String
          Dim intEL As Integer

640       intEL = Erl
650       strES = Err.Description
660       LogError "frmEditMicrobiologyNew", "grdAB_Click", intEL, strES

End Sub
Private Sub DeleteSensitivity(ByVal Index As Integer, ByVal ABName As String)

          Dim ABCode As String
          Dim sql As String
          Dim tb As Recordset

10        On Error GoTo DeleteSensitivity_Error

20        ABCode = AntibioticCodeFor(ABName)
30        sql = "DELETE FROM Sensitivities WHERE " & _
                "SampleID = '" & Val(txtSampleID) + SysOptMicroOffset(0) & "' " & _
                "and IsolateNumber = '" & Index & "' " & _
                "and AntibioticCode = '" & ABCode & "'"
40        Cnxn(0).Execute sql

50        Exit Sub

DeleteSensitivity_Error:

          Dim strES As String
          Dim intEL As Integer

60        intEL = Erl
70        strES = Err.Description
80        LogError "frmEditMicrobiologyNew", "DeleteSensitivity", intEL, strES, sql

End Sub

Private Sub grdAB_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)

10        pBar = 0

End Sub

Private Sub imgLast_Click()

          Dim sql As String
          Dim tb As Recordset

10        On Error GoTo imgLast_Click_Error

20        If cmdSaveDemographics.Enabled Or cmdSaveInc.Enabled Then
30            If iMsg("Save Details", vbQuestion + vbYesNo) = vbYes Then
40                GetSampleIDWithOffset
50                SaveDemographics
60                cmdSaveDemographics.Enabled = False
70                cmdSaveInc.Enabled = False
80            End If
90        End If

100       GetSampleIDWithOffset

110       sql = "SELECT TOP 1 SampleID FROM MicroSiteDetails WHERE " & _
                "Site like '" & cmbSiteSearch & "' " & _
                "ORDER BY SampleID Desc"

120       Set tb = New Recordset
130       RecOpenClient 0, tb, sql
140       If Not tb.EOF Then
150           txtSampleID = Val(tb!SampleID & "") - SysOptMicroOffset(0)
160       End If

170       GetSampleIDWithOffset
180       LoadAllDetails

190       cmdSaveDemographics.Enabled = False
200       cmdSaveInc.Enabled = False
210       cmdSaveMicro.Enabled = False
220       cmdSaveHold.Enabled = False

230       Exit Sub

imgLast_Click_Error:

          Dim strES As String
          Dim intEL As Integer

240       intEL = Erl
250       strES = Err.Description
260       LogError "frmEditMicrobiologyNew", "imgLast_Click", intEL, strES, sql

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

90        SetDatesColour Me

100       cmdSaveInc.Enabled = True
110       cmdSaveDemographics.Enabled = True

120       Exit Sub

iRecDate_Click_Error:

          Dim strES As String
          Dim intEL As Integer

130       intEL = Erl
140       strES = Err.Description
150       LogError "frmEditMicrobiologyNew", "iRecDate_Click", intEL, strES


End Sub

Private Sub iRelevant_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)

10        On Error GoTo iRelevant_MouseMove_Error

20        If cmdSaveMicro.Enabled Then
30            MoveCursorToSaveButton
40        End If

50        Exit Sub

iRelevant_MouseMove_Error:

          Dim strES As String
          Dim intEL As Integer

60        intEL = Erl
70        strES = Err.Description
80        LogError "frmEditMicrobiologyNew", "iRelevant_MouseMove", intEL, strES


End Sub

Private Sub lblBATResult_Click()

10        On Error GoTo lblBATResult_Click_Error

20        Select Case lblBATResult.Caption
          Case "": lblBATResult.Caption = "Positive"
30        Case "Positive": lblBATResult.Caption = "Negative"
40        Case "Negative": lblBATResult.Caption = "Inclusive"
50        Case "Inclusive": lblBATResult.Caption = "No Sample Received"
60        Case Else: lblBATResult.Caption = ""
70        End Select

          'txtBATComments.Enabled = (lblBATResult.Caption = "No Sample Received")

80        ShowUnlock 9

90        Exit Sub

lblBATResult_Click_Error:

          Dim strES As String
          Dim intEL As Integer

100       intEL = Erl
110       strES = Err.Description
120       LogError "frmEditMicrobiologyNew", "lblBATResult_Click", intEL, strES

End Sub

Private Sub lblCDiffCulture_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

10        On Error GoTo lblCDiffCulture_MouseUp_Error

20        CycleLabel ListCDiffCulture(), lblCDiffCulture

30        cmdSaveHold.Enabled = True
40        cmdSaveMicro.Enabled = True

50        ShowUnlock 10

60        Exit Sub

lblCDiffCulture_MouseUp_Error:

          Dim strES As String
          Dim intEL As Integer

70        intEL = Erl
80        strES = Err.Description
90        LogError "frmEditMicrobiologyNew", "lblCDiffCulture_MouseUp", intEL, strES

End Sub


Private Sub lblCrypto_Click()

10        On Error GoTo lblCrypto_Click_Error

20        CycleLabel ListCrypto(), lblCrypto

30        cmdSaveHold.Enabled = True
40        cmdSaveMicro.Enabled = True

50        ShowUnlock 11

60        Exit Sub

lblCrypto_Click_Error:

          Dim strES As String
          Dim intEL As Integer

70        intEL = Erl
80        strES = Err.Description
90        LogError "frmEditMicrobiologyNew", "lblCrypto_Click", intEL, strES

End Sub

Private Sub lblFinal_Click()

10        With lblFinal
20            .BackColor = vbGreen
30            .FontBold = True
40        End With

50        With lblInterim
60            .BackColor = &H8000000F
70            .FontBold = False
80        End With

End Sub

Private Sub lblFOB_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)

10        CycleLabel ListFOB(), lblFOB(Index)

20        cmdSaveHold.Enabled = True
30        cmdSaveMicro.Enabled = True

40        ShowUnlock 5

End Sub


Private Sub lblGDH_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

          Dim MicroGDHComment As String
          Dim Obs As New Observations

10        On Error GoTo lblGDH_MouseUp_Error

20        CycleLabel ListGDH, lblGDH

30        MicroGDHComment = GetOptionSetting("MicrobiologyCDiffGDHNegativeComment", "")
40        If MicroGDHComment <> "" And _
             InStr(1, UCase(lblGDH.Caption), "NEGATIVE") > 0 And _
             InStr(txtCDiffMSC, MicroGDHComment) = 0 Then

50            Obs.Save SampleIDWithOffset, False, "MICROCDIFF", MicroGDHComment

60            LoadComments

70        End If

80        cmdSaveHold.Enabled = True
90        cmdSaveMicro.Enabled = True

100       ShowUnlock 10

110       Exit Sub

lblGDH_MouseUp_Error:

          Dim strES As String
          Dim intEL As Integer

120       intEL = Erl
130       strES = Err.Description
140       LogError "frmEditMicrobiologyNew", "lblGDH_MouseUp", intEL, strES

End Sub

Private Sub lblGiardia_Click()

10        On Error GoTo lblGiardia_Click_Error

20        CycleLabel ListGiardia(), lblGiardia

30        cmdSaveHold.Enabled = True
40        cmdSaveMicro.Enabled = True

50        ShowUnlock 11

60        Exit Sub

lblGiardia_Click_Error:

          Dim strES As String
          Dim intEL As Integer

70        intEL = Erl
80        strES = Err.Description
90        LogError "frmEditMicrobiologyNew", "lblGiardia_Click", intEL, strES

End Sub

Private Sub lblHPylori_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

10        CycleLabel ListHPylori(), lblHPylori

20        cmdSaveHold.Enabled = True
30        cmdSaveMicro.Enabled = True

40        ShowUnlock 13

End Sub


Private Sub lblInterim_Click()

10        With lblInterim
20            .BackColor = vbGreen
30            .FontBold = True
40        End With

50        With lblFinal
60            .BackColor = &H8000000F
70            .FontBold = False
80        End With

End Sub

Private Sub lblLegionellaAT_Click()

10        Select Case lblLegionellaAT.Caption
          Case "": lblLegionellaAT.Caption = "Negative"
20        Case "Negative": lblLegionellaAT.Caption = "Positive"
30        Case Else: lblLegionellaAT.Caption = ""
40        End Select

50        ShowUnlock 9

End Sub

Private Sub lblPCR_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

          Dim MicroPCRComment As String
          Dim Obs As New Observations

10        On Error GoTo lblPCR_MouseUp_Error


20        CycleLabel ListPCR, lblPCR

30        MicroPCRComment = GetOptionSetting("MicrobiologyCDiffPCRNegativeComment", "")
40        If MicroPCRComment <> "" And _
             InStr(1, UCase(lblPCR.Caption), "TOXIN NEGATIVE") > 0 And _
             InStr(txtCDiffMSC, MicroPCRComment) = 0 Then

50            Obs.Save SampleIDWithOffset, False, "MICROCDIFF", MicroPCRComment

60            LoadComments

70        End If


80        cmdSaveHold.Enabled = True
90        cmdSaveMicro.Enabled = True

100       ShowUnlock 10


110       Exit Sub

lblPCR_MouseUp_Error:

          Dim strES As String
          Dim intEL As Integer

120       intEL = Erl
130       strES = Err.Description
140       LogError "frmEditMicrobiologyNew", "lblPCR_MouseUp", intEL, strES

End Sub

Private Sub lblPneuAT_Click()

10        Select Case lblPneuAT.Caption
          Case "": lblPneuAT.Caption = "Negative"
20        Case "Negative": lblPneuAT.Caption = "Positive"
30        Case Else: lblPneuAT.Caption = ""
40        End Select

50        ShowUnlock 9

End Sub

Private Sub lblRSV_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

10        On Error GoTo lblRSV_MouseUp_Error

20        CycleLabel ListRSV(), lblRSV

30        ShowUnlock 8

40        Exit Sub

lblRSV_MouseUp_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmEditMicrobiologyNew", "lblRSV_MouseUp", intEL, strES

End Sub


Private Sub lblSetAllR_Click(Index As Integer)

          Dim Y As Integer

10        On Error GoTo lblSetAllR_Click_Error

20        With grdAB(Index)
30            .Col = 2
40            For Y = 1 To .Rows - 1
50                If .TextMatrix(Y, 0) <> "" Then
60                    .TextMatrix(Y, 1) = "R"
70                    .Row = Y
80                    Set .CellPicture = imgSquareTick.Picture
90                End If
100           Next
110       End With


120       cmdSaveMicro.Enabled = True
130       cmdSaveHold.Enabled = True

140       Exit Sub

lblSetAllR_Click_Error:

          Dim strES As String
          Dim intEL As Integer

150       intEL = Erl
160       strES = Err.Description
170       LogError "frmEditMicrobiologyNew", "lblSetAllR_Click", intEL, strES


End Sub

Private Sub lblSetAllS_Click(Index As Integer)

          Dim Y As Integer

10        On Error GoTo lblSetAllS_Click_Error

20        With grdAB(Index)
30            .Col = 2
40            For Y = 1 To .Rows - 1
50                If .TextMatrix(Y, 0) <> "" Then
60                    .TextMatrix(Y, 1) = "S"
70                    .Row = Y
80                    Set .CellPicture = imgSquareCross.Picture
90                End If
100           Next
110       End With

120       cmdSaveMicro.Enabled = True
130       cmdSaveHold.Enabled = True

140       Exit Sub

lblSetAllS_Click_Error:

          Dim strES As String
          Dim intEL As Integer

150       intEL = Erl
160       strES = Err.Description
170       LogError "frmEditMicrobiologyNew", "lblSetAllS_Click", intEL, strES


End Sub


Private Sub lblToxinA_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

10        On Error GoTo lblToxinA_MouseUp_Error

20        CycleLabel ListCDiffToxinAB, lblToxinA

30        cmdSaveHold.Enabled = True
40        cmdSaveMicro.Enabled = True

50        ShowUnlock 10

60        Exit Sub

lblToxinA_MouseUp_Error:

          Dim strES As String
          Dim intEL As Integer

70        intEL = Erl
80        strES = Err.Description
90        LogError "frmEditMicrobiologyNew", "lblToxinA_MouseUp", intEL, strES

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
130               FlashNoPrevious lNoPrevious
140           End If
150       End With

160       Exit Sub

bDoB_Click_Error:

          Dim strES As String
          Dim intEL As Integer

170       intEL = Erl
180       strES = Err.Description
190       LogError "frmEditMicrobiologyNew", "bDoB_Click", intEL, strES

End Sub

Private Sub bFAX_Click(Index As Integer)

          Dim tb As New Recordset
          Dim sql As String
          Dim FaxNumber As String
          Dim Disp As String




10        On Error GoTo bFAX_Click_Error

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

120       If Len(cmbWard) = 0 Then
130           iMsg "Must have Ward entry.", vbCritical
140           Exit Sub
150       End If

160       If UCase(Trim$(cmbWard)) = "GP" Then
170           If Len(cmbGP) = 0 Then
180               iMsg "Must have Ward or GP entry.", vbCritical
190               Exit Sub
200           End If
210       End If


220       If UCase(cmbWard) = "GP" Then
230           sql = "SELECT * from GPS WHERE text = '" & cmbGP & "' and hospitalcode = '" & ListCodeFor("HO", cmbHospital) & "'"
240           Set tb = New Recordset
250           RecOpenServer 0, tb, sql
260           If Not tb.EOF Then
270               FaxNumber = Trim$(tb!FAX & "")
280           End If
290       Else
300           sql = "SELECT * from wards WHERE text = '" & cmbWard & "' and hospitalcode = '" & ListCodeFor("HO", cmbHospital) & "'"
310           Set tb = New Recordset
320           RecOpenServer 0, tb, sql
330           If Not tb.EOF Then
340               FaxNumber = Trim$(tb!FAX & "")
350           End If
360       End If


370       FaxNumber = iBOX("Faxnumber ", , FaxNumber)

380       FaxNumber = Trim(FaxNumber)

390       If Trim(FaxNumber) = "" Then
400           iMsg "No Fax Number Entered!" & vbCrLf & "Fax Cancelled!"
410           Exit Sub
420       End If


430       If Not IsNumeric(FaxNumber) Then
440           iMsg "Incorrect Fax Number Entered!" & vbCrLf & "Fax Cancelled!"
450           Exit Sub
460       End If



470       If Len(FaxNumber) < 4 Then
480           iMsg "Incorrect Fax Number Entered!" & vbCrLf & "Fax Cancelled!"
490           Exit Sub
500       End If

510       SaveDemographics



          'LogTimeOfPrinting txtSampleID, Choose(sstab1.Tab, "H", "B", "C", "E", "Q", "I", "")
520       sql = "SELECT * FROM PrintPending WHERE " & _
                "Department = 'N' " & _
                "AND SampleID = '" & SampleIDWithOffset & "' " & _
                "AND FaxNumber <> ''"
530       Set tb = New Recordset
540       RecOpenClient 0, tb, sql
550       If tb.EOF Then
560           tb.AddNew
570       End If
580       tb!SampleID = SampleIDWithOffset
590       tb!Department = "N"
600       tb!Initiator = UserName
610       tb!Ward = cmbWard
620       tb!Clinician = cmbClinician
630       tb!GP = cmbGP
640       tb!UsePrinter = "Fax"
650       tb!FaxNumber = FaxNumber
660       tb.Update



670       Exit Sub

bFAX_Click_Error:

          Dim strES As String
          Dim intEL As Integer

680       intEL = Erl
690       strES = Err.Description
700       LogError "frmEditMicrobiologyNew", "bFAX_Click", intEL, strES, sql


End Sub

Private Sub bHistory_Click()

10        pBar = 0

20        With frmMicroReport
30            .PatChart = txtChart
40            .PatName = txtName
50            .PatDoB = txtDoB
60            .PatSex = Trim$(Left$(txtSex & " ", 1))
70            .PatWard = cmbWard
80            .PatClinician = cmbClinician
90            .PatGP = cmbGP
100           .Show 1
110       End With

End Sub



Private Sub cmbSite_Click()

10        If UCase(cmbSite) = "BLOOD CULTURE" Then
20            If BacTek3DInUse Then
30                If iMsg("Order Blood Culture?", vbQuestion + vbYesNo) = vbYes Then
40                    If (txtName = "" Or txtDoB = "" Or txtSex = "") Then
50                        iMsg "Must have patient Name, DoB, and Sex"
60                        cmbSite.ListIndex = -1
70                        Exit Sub
80                    End If
90                    SaveDemographics
100                   SaveBloodCultureRequest

110               End If
120           End If
130       End If

140       cmbSiteEffects

150       cmbSiteSearch = cmbSite



End Sub

Private Sub SaveBloodCultureRequest()

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo SaveBloodCultureRequest_Error

20        sql = "SELECT * FROM BloodCultureRequests WHERE " & _
                "SampleID = '" & txtSampleID & "'"
30        Set tb = New Recordset
40        RecOpenClient 0, tb, sql
50        If tb.EOF Then
60            tb.AddNew
70        End If
80        tb!SampleID = txtSampleID
90        tb!RequestedDateTime = Format$(Now, "dd/MMM/yyyy HH:nn:ss")
100       tb!RequestedBy = UserName
110       tb!Programmed = 0
120       tb.Update

130       Exit Sub

SaveBloodCultureRequest_Error:

          Dim strES As String
          Dim intEL As Integer

140       intEL = Erl
150       strES = Err.Description
160       LogError "frmEditMicrobiologyNew", "SaveBloodCultureRequest", intEL, strES, sql

End Sub

Private Sub cmdOrderTests_Click()

          Dim f As Form

10        pBar = 0

20        If cmbSite = "Urine" Then

30            Set f = frmMicroOrderUrine

40            f.txtSampleID = txtSampleID
50            f.Show 1
60            If f.SiteDetails <> "" Then
70                txtSiteDetails = f.SiteDetails
80            End If
90            If f.chkUrine(2) Then
100               SSTab1.TabVisible(7) = True
110           End If
120           Unload f
130           Set f = Nothing

140       ElseIf InStr(1, cmbSite, "Faeces") Then

150           OrderFaeces

160       Else

170           With frmMicroOrders
180               .txtSampleID = txtSampleID
190               .Show 1
200           End With

210       End If

End Sub


Private Sub bprint_Click()
      Dim pSampleID As String
10    On Error GoTo bprint_Click_Error

20    pSampleID = SysOptMicroOffset(0) + Val(txtSampleID)
30    If lblFinal.BackColor = vbGreen Then
40        If SampleRelasedtoConsultant(pSampleID, "Micro") Then
50            iMsg "This report is being reviewed by consultant and cannot be released to the ward as a final report.", vbInformation
60            Exit Sub
70        End If
80    End If
90    If Not EntriesOK(txtSampleID, txtName, txtSex, cmbWard.Text, cmbGP.Text) Then
100       Exit Sub
110   End If
120   If cmdDemoVal.Caption = "&Validate" Then
130       If iMsg("Demographics are not validated. Do you want to validate now?", vbQuestion + vbYesNo) = vbYes Then
140           Exit Sub
150       Else
160           ValidateDemographics True
170       End If
180   End If

190   SaveMicro

200   frmValidateAll.SampleIDToValidate = SampleIDWithOffset
210   frmValidateAll.Show 1


      'If Not IsCSValid() Then
      '
      '    If iMsg("C/S Results not Validated!" & vbCrLf & _
           '            "Validate C/S now ?", vbQuestion + vbYesNo) = vbYes Then
      '
      '        SaveSensitivities gYES
      '
      '    Else
      '        Exit Sub
      '    End If
      '
      'End If


220   PrintThis


230   txtSampleID = Format$(Val(txtSampleID) + 1)
240   GetSampleIDWithOffset
250   LoadAllDetails

260   Exit Sub

bprint_Click_Error:

      Dim strES As String
      Dim intEL As Integer

270   intEL = Erl
280   strES = Err.Description
290   LogError "frmEditMicrobiologyNew", "bPrint_Click", intEL, strES

End Sub

Private Sub SaveDemographics()

          Dim sql As String
          Dim tb As Recordset

10        On Error GoTo SaveDemographics_Error

20        txtSampleID = Format(Val(txtSampleID))
30        If Val(txtSampleID) = 0 Then Exit Sub

40        SaveComments

50        If Trim$(tSampleTime) <> "__:__" Then
60            If Not IsDate(tSampleTime) Then
70                iMsg "Invalid Time", vbExclamation
80                Exit Sub
90            End If
100       End If

110       SaveMicroSiteDetails

120       If DemographicsHaveChanged(SampleIDWithOffset) Then
130           sql = "Select * from Demographics where " & _
                    "SampleID = '" & SampleIDWithOffset & "'"

140           Set tb = New Recordset
150           RecOpenClient 0, tb, sql
160           If tb.EOF Then
170               tb.AddNew
180           Else

190               If GetOptionSetting("MicroConfirmChangeName", "True") = "True" Then
200                   If Trim$(tb!PatName & "") <> "" And _
                         Trim$(UCase$(tb!PatName & "")) <> Trim$(UCase$(txtName)) Then
210                       If FlagMessage("Name", tb!PatName, txtName, SampleIDWithOffset) Then
220                           txtName = Trim$(tb!PatName & "")
230                       End If
240                   End If
250               End If

260               If GetOptionSetting("MicroConfirmChangeDoB", "True") = "True" Then
270                   If Not IsNull(tb!Dob) Then
280                       If Format(tb!Dob, "dd/mm/yyyy") <> Format(txtDoB, "dd/mm/yyyy") Then
290                           If FlagMessage("DoB", tb!Dob, txtDoB, SampleIDWithOffset) Then
300                               txtDoB = Format(tb!Dob, "dd/mm/yyyy")
310                           End If
320                       End If
330                   End If
340               End If

350               If GetOptionSetting("MicroConfirmChangeChart", "True") = "True" Then
360                   If Trim$(tb!Chart & "") <> "" And Trim$(UCase$(tb!Chart & "")) <> Trim$(UCase$(txtChart)) Then
370                       If FlagMessage("Chart", tb!Chart, txtChart, SampleIDWithOffset) Then
380                           txtChart = tb!Chart & ""
390                       End If
400                   End If
410               End If

420               If GetOptionSetting("MicroConfirmChangeWard", "True") = "True" Then
430                   If Trim$(tb!Ward & "") <> "" And Trim$(UCase$(tb!Ward & "")) <> Trim$(UCase$(cmbWard)) Then
440                       If FlagMessage("Ward", tb!Ward, cmbWard, SampleIDWithOffset) Then
450                           cmbWard = tb!Ward & ""
460                       End If
470                   End If
480               End If

490               If GetOptionSetting("MicroConfirmChangeClinician", "True") = "True" Then
500                   If Trim$(tb!Clinician & "") <> "" And Trim$(UCase$(tb!Clinician & "")) <> Trim$(UCase$(cmbClinician)) Then
510                       If FlagMessage("Clinician", tb!Clinician, cmbClinician, SampleIDWithOffset) Then
520                           cmbClinician = tb!Clinician & ""
530                       End If
540                   End If
550               End If

560           End If

570           tb!RooH = cRooH(0)

580           If IsDate(tRecTime) Then
590               tb!RecDate = Format$(dtRecDate & " " & tRecTime, "dd/MMM/yyyy HH:nn")
600           Else
610               tb!RecDate = Format$(dtRecDate & " " & Format$(Now, "HH:nn"), "dd/MMM/yyyy HH:nn")
620           End If
630           tb!Rundate = Format$(dtRunDate, "dd/mmm/yyyy")
640           If IsDate(tSampleTime) Then
650               tb!SampleDate = Format$(dtSampleDate, "dd/mmm/yyyy") & " " & Format$(tSampleTime, "hh:mm")
660           Else
670               tb!SampleDate = Format$(dtSampleDate, "dd/mmm/yyyy")
680           End If
690           tb!SampleID = SampleIDWithOffset
700           tb!Chart = txtChart
710           tb!PatName = Trim$(txtName)
720           If IsDate(txtDoB) Then
730               tb!Dob = Format$(txtDoB, "dd/mmm/yyyy")
740           Else
750               tb!Dob = Null
760           End If
770           tb!AandE = txtAandE
780           tb!Age = txtAge
790           tb!sex = Left$(txtSex, 1)
800           tb!Addr0 = taddress(0)
810           tb!Addr1 = taddress(1)
820           tb!Ward = Left$(cmbWard, 50)
830           tb!Clinician = Left$(cmbClinician, 50)
840           tb!GP = Left$(cmbGP, 50)
850           tb!ClDetails = Left$(txtClinDetails, 50)
860           tb!Hospital = cmbHospital
870           tb!Pregnant = chkPregnant
880           tb!Operator = Left$(UserName, 20)
890           tb!PenicillinAllergy = chkPenicillin.Value = 1
900           tb.Update

910           LogTimeOfPrinting SampleIDWithOffset, "D"
920       End If

930       Exit Sub

SaveDemographics_Error:

          Dim strES As String
          Dim intEL As Integer

940       intEL = Erl
950       strES = Err.Description
960       LogError "frmEditMicrobiology", "SaveDemographics", intEL, strES, sql

End Sub

Private Sub cmdSaveMicro_Click()

10        pBar = 0


20        If CheckValidStatus(SSTab1.Tab) Then
30            iMsg "Validated Result - Cannot Save", vbCritical
40            Exit Sub
50        End If

60        GetSampleIDWithOffset
70        SaveMicro

          'Select Case SSTab1.Tab
          '    Case 1: SaveUrine
          '    Case 2: SaveIdent
          '    Case 3: SaveFaeces
          '    Case 4: SaveIsolates
          '        SaveSensitivities gNOCHANGE
          '    Case 5: SaveFOB
          '    Case 6: SaveRotaAdeno
          '    Case 7: SaveRedSub
          '    Case 9: SaveFluids
          '    Case 10: SaveCdiff
          '    Case 11: SaveOP
          '    Case 14: SaveRSV
          'End Select
          '
          'SaveComments
          'UPDATEMRU txtSampleID, cMRU

80        txtSampleID = Format$(Val(txtSampleID) + 1)

90        GetSampleIDWithOffset

100       cmdSaveMicro.Enabled = False
110       cmdSaveHold.Enabled = False
120       LoadAllDetails

End Sub

Private Sub SaveSensitivities(ByVal Validate As Integer)

          Dim tb As Recordset
          Dim sql As String
          Dim intOrg As Integer
          Dim n As Integer
          Dim ABCode As String
          Dim ReportCounter As Integer

10        On Error GoTo SaveSensitivities_Error

20        ReportCounter = 0

30        For intOrg = 1 To 4

40            With grdAB(intOrg)

50                For n = 1 To .Rows - 1
60                    If .TextMatrix(n, 0) <> "" Then
70                        ABCode = AntibioticCodeFor(.TextMatrix(n, 0))
80                        sql = "Select * from Sensitivities where " & _
                                "SampleID = '" & SampleIDWithOffset & "' " & _
                                "and IsolateNumber = '" & intOrg & "' " & _
                                "and AntibioticCode = '" & ABCode & "'"
90                        Set tb = New Recordset
100                       RecOpenClient 0, tb, sql
110                       If tb.EOF Then
120                           tb.AddNew
130                           tb!Rundate = Format(Now, "dd/mmm/yyyy")
140                           tb!RunDateTime = Format(Now, "dd/mmm/yyyy hh:mm")
150                           tb!UserName = UserName
160                       End If
170                       tb!SampleID = SampleIDWithOffset
180                       tb!IsolateNumber = intOrg
190                       tb!AntibioticCode = ABCode
200                       tb!Antibiotic = .TextMatrix(n, 0)
210                       tb!RSI = .TextMatrix(n, 1)
220                       tb!CPOFlag = .TextMatrix(n, 2)
230                       tb!Result = .TextMatrix(n, 3)
240                       tb!Organism = cmbOrgName(intOrg)

250                       .Row = n
260                       .Col = 0
270                       If .CellFontBold = True Then
280                           tb!Secondary = 1
290                       Else
300                           tb!Secondary = 0
310                       End If
320                       If .CellBackColor = &HFFFFC0 Then
330                           tb!Forced = 1
340                       Else
350                           tb!Forced = 0
360                       End If
370                       .Col = 2

                          '        If .CellPicture = 0 Then
                          '          If .TextMatrix(n, 1) = "R" Then
                          '            tb!Report = 1
                          '          ElseIf .TextMatrix(n, 1) = "S" Then
                          '            ReportCounter = ReportCounter + 1
                          '            If ReportCounter < 4 Then
                          '              tb!Report = 1
                          '            Else
                          '              tb!Report = 0
                          '            End If
                          '          Else
                          '            tb!Report = Null
                          '          End If
                          '        Else
380                       If .CellPicture = imgSquareTick.Picture Then
390                           tb!Report = 1
400                       ElseIf .CellPicture = imgSquareCross.Picture Then
410                           tb!Report = 0
420                       Else
430                           tb!Report = Null
440                       End If
                          '        End If
450                       tb.Update
460                   End If
470               Next
480           End With

490       Next

500       If Validate = gYES Then
510           sql = "Update Sensitivities " & _
                    "Set Valid = 1, " & _
                    "AuthoriserCode = '" & UserName & "' " & _
                    "where SampleID = '" & SampleIDWithOffset & "'"
520           Cnxn(0).Execute sql
530       ElseIf Validate = gNO Then
540           sql = "Update Sensitivities " & _
                    "Set Valid = 0, " & _
                    "AuthoriserCode = NULL " & _
                    "where SampleID = '" & SampleIDWithOffset & "'"
550           Cnxn(0).Execute sql
560       End If

570       Exit Sub

SaveSensitivities_Error:

          Dim strES As String
          Dim intEL As Integer

580       intEL = Erl
590       strES = Err.Description
600       LogError "frmEditMicrobiology", "SaveSensitivities", intEL, strES, sql

End Sub


Private Function LoadSensitivities() As Integer
      'Returns number of isolates

          Dim tb As Recordset
          Dim sql As String
          Dim intIsolate As Integer
          Dim Rows As Integer
          Dim s As String

10        On Error GoTo LoadSensitivities_Error

20        LoadSensitivities = 0

30        lblCf.Visible = False
40        If IsCystic(SampleIDWithOffset) Then
50            lblCf.Visible = True
60        End If
70        fraCS.Enabled = True

80        For intIsolate = 1 To 4

90            sql = "IF NOT EXISTS(SELECT * FROM ABDefinitions " & _
                    "              WHERE Site = '" & cmbSite & "' " & _
                    "              AND OrganismGroup = '" & cmbOrgGroup(intIsolate) & "')" & _
                    "  INSERT INTO ABDefinitions " & _
                    "  SELECT AntibioticName, OrganismGroup, '" & cmbSite & "' Site, ListOrder, PriSec, AutoReport, AutoReportIf, AutoPriority " & _
                    "  FROM ABDefinitions " & _
                    "  WHERE Site = 'Generic' " & _
                    "  AND OrganismGroup = '" & cmbOrgGroup(intIsolate) & "'"
100           Cnxn(0).Execute sql

110           With grdAB(intIsolate)
120               .Rows = 2
130               .AddItem ""
140               .RemoveItem 1

150               sql = "SELECT B.AntibioticName, S.Report, S.CPOFlag, S.RSI, S.RunDateTime, S.UserName, S.Result, " & _
                        "B.AutoReport, B.AutoReportIf , B.AutoPriority, B.ListOrder " & _
                        "FROM " & _
                        "(SELECT AntibioticName, Listorder, COALESCE(AutoReport, 0) AutoReport, " & _
                        "COALESCE(AutoReportIf,'') AutoReportIf, COALESCE(AutoPriority,0) AutoPriority " & _
                        "FROM ABDefinitions WHERE Site = '" & cmbSite & "' " & _
                        "AND OrganismGroup = '" & cmbOrgGroup(intIsolate) & "' AND PriSec = 'P' ) B " & _
                        "LEFT OUTER JOIN " & _
                        "(Select * from Sensitivities WHERE SampleID = '" & SampleIDWithOffset & "' " & _
                        "AND IsolateNumber = '" & intIsolate & "' " & _
                        "AND COALESCE(Forced, 0) = 0 AND COALESCE(Secondary, 0) = 0) S ON B.AntibioticName = S.Antibiotic " & _
                        "ORDER BY B.ListOrder"

                  '        "SELECT S.Antibiotic, S.Report, S.CPOFlag, S.RSI, " & _
                           '              "S.RunDateTime, S.UserName, S.Result, B.AutoReport, B.AutoReportIf, B.AutoPriority, B.ListOrder " & _
                           '              "FROM Sensitivities S LEFT OUTER JOIN Antibiotics A " & _
                           '              "ON S.AntibioticCode = A.Code " & _
                           '              "INNER JOIN (SELECT AntibioticName, Listorder, " & _
                           '              "            COALESCE(AutoReport, 0) AutoReport, COALESCE(AutoReportIf,'') AutoReportIf, " & _
                           '              "            COALESCE(AutoPriority,0) AutoPriority FROM ABDefinitions " & _
                           '              "            WHERE Site = '" & cmbSite & "' " & _
                           '              "            AND OrganismGroup = '" & cmbOrgGroup(intIsolate) & "')  B " & _
                           '              "ON S.Antibiotic = B.AntibioticName " & _
                           '              "WHERE SampleID = '" & SampleIDWithOffset & "' " & _
                           '              "AND IsolateNumber = '" & intIsolate & "' " & _
                           '              "AND COALESCE(Forced, 0) = 0 " & _
                           '              "AND COALESCE(Secondary, 0) = 0 " & _
                           '              "ORDER BY B.ListOrder"

160               Set tb = New Recordset
170               RecOpenServer 0, tb, sql
                  '    End If
180               If Not tb.EOF Then
190                   LoadSensitivities = intIsolate
200               End If
210               Do While Not tb.EOF
220                   .AddItem tb!AntibioticName & vbTab & _
                               tb!RSI & vbTab & _
                               tb!CPOFlag & vbTab & _
                               tb!Result & vbTab & _
                               Format(tb!RunDateTime, "dd/mm/yy hh:mm") & vbTab & _
                               tb!UserName & "" & vbTab & _
                               tb!AutoReport & vbTab & _
                               tb!AutoReportIf & vbTab & _
                               tb!AutoPriority & vbTab & _
                               tb!ListOrder


230                   .Row = .Rows - 1
240                   .Col = 2
250                   If IsNull(tb!Report) Then
                          'Apply rule for autoreporting here
260                       If AutoReportAB(intIsolate) Then
270                           Set .CellPicture = imgSquareTick.Picture
280                       Else
290                           Set .CellPicture = Me.Picture
300                       End If
310                   Else
320                       Set .CellPicture = IIf(tb!Report, imgSquareTick.Picture, imgSquareCross.Picture)
330                   End If

340                   tb.MoveNext
350               Loop
360               If .Rows > 2 Then
370                   .RemoveItem 1
380               End If
390           End With

400           LoadSensitivitiesForced intIsolate
410           LoadSensitivitiesSecondary intIsolate

420           FillABSelect intIsolate
430       Next

440       LockFraCS False

450       If LoadLockStatus(4) Then
460           fraCS.Enabled = False
470       End If

480       Exit Function

LoadSensitivities_Error:

          Dim strES As String
          Dim intEL As Integer

490       intEL = Erl
500       strES = Err.Description
510       LogError "frmEditMicrobiologyNew", "LoadSensitivities", intEL, strES, sql

End Function

Private Sub cmdSaveDemographics_Click()

10        pBar = 0

20        If Not EntriesOK(txtSampleID, txtName, txtSex, cmbWard.Text, cmbGP.Text) Then
30            Exit Sub
40        End If

50        If Not CheckTimes() Then Exit Sub

60        cmdSaveDemographics.Caption = "Saving"

70        GetSampleIDWithOffset

80        SaveDemographics
90        UPDATEMRU txtSampleID, cMRU
100       txtSampleID.SetFocus

110       cmdSaveDemographics.Caption = "Save && &Hold"
120       cmdSaveDemographics.Enabled = False
130       cmdSaveInc.Enabled = False

End Sub


Private Sub cmdSaveInc_Click()

10        pBar = 0

20        If Not EntriesOK(txtSampleID, txtName, txtSex, cmbWard.Text, cmbGP.Text) Then
30            Exit Sub
40        End If

          '50    If lblChartNumber.BackColor = vbRed Then
          '60      If iMsg("Confirm this Patient has" & vbCrLf & _
           '                lblChartNumber.Caption, vbYesNo + vbQuestion, , vbRed, 12) = vbNo Then
          '70        Exit Sub
          '80      End If
          '90    End If

50        If Not CheckTimes() Then Exit Sub

60        cmdSaveDemographics.Caption = "Saving"

70        GetSampleIDWithOffset

80        SaveDemographics
90        UPDATEMRU txtSampleID, cMRU

100       cmdSaveDemographics.Caption = "Save && &Hold"
110       cmdSaveDemographics.Enabled = False
120       cmdSaveInc.Enabled = False

130       txtSampleID = Format$(Val(txtSampleID) + 1)

140       GetSampleIDWithOffset
150       LoadAllDetails

160       txtSampleID.SelStart = 0
170       txtSampleID.SelLength = Len(txtSampleID)
180       txtSampleID.SetFocus

190       cmdSaveMicro.Enabled = False
200       cmdSaveHold.Enabled = False

End Sub

Private Sub bsearch_Click()


10        pBar = 0

20        With frmPatHistoryNew
30            .oHD(1) = True
40            .oFor(0) = True
50            .txtName = Trim$(txtName)
60            .FromEdit = True
70            .EditScreen = Me
80            .bsearch = True
90            If Not .NoPreviousDetails Then
100               .Show 1
110           Else
120               FlashNoPrevious lNoPrevious
130           End If
140       End With

End Sub

Private Sub cmdValidateMicro_Click()

      Dim tb As Recordset
      Dim sql As String
      Dim Section As String
      Dim Validate As Boolean
      Dim SaveSSTab As Integer

10    On Error GoTo cmdValidateMicro_Click_Error

20    GetSampleIDWithOffset

30    Validate = cmdValidateMicro.Caption = "&Validate"

40    Select Case SSTab1.Tab
      Case 0: Section = "DEMOGRAPHICS"
50    Case 1: Section = "URINE"
60    Case 2: Section = "IDENTIFICATION"
70    Case 3: Section = "FAECES"
80    Case 4: Section = "CANDS"
90    Case 5: Section = "FOB"
100   Case 6: Section = "ROTAADENO"
110   Case 7: Section = "REDSUB"
120   Case 8: Section = "RSV"
130   Case 9: Section = "CSF"
140   Case 10: Section = "CDIFF"
150   Case 11: Section = "OP"
160   Case 12: Section = "BLOODCULTURE"
170   Case 13: Section = "HPYLORI"
180   Case Else: Section = ""
190   End Select

200   If Validate Then
210       If cmdDemoVal.Caption = "&Validate" Then
220           If iMsg("Do you wish to validate demographics !", vbYesNo) = vbYes Then
230               ValidateDemographics True
240           End If
250       End If

260       If Section = "CANDS" Then
270           If QueryGent() Then Exit Sub
              'If QueryCEF() Then Exit Sub
280       End If
290       cmdValidateMicro.Caption = "Un&Validate"
300       UpdatePrintValidLog CStr(SampleIDWithOffset), Section, 1, 2

310       If cmdReleaseReport.BackColor = vbCyan Then
320           If iMsg("Do you want to re-release report to Consultant", vbQuestion + vbYesNo, , vbRed) = vbYes Then
                  'UpdateConsultantList CStr(SampleIDWithOffset), "Micro", ReleasedToConsultant, 0, 0
                  
330               If SampleAddedtoConsultantList(CStr(SampleIDWithOffset), "Micro") Then
340                   RemoveReport 0, CStr(SampleIDWithOffset), "N", 0
350                   Call PrintThis("SaveTemp")
360                   cmdReleaseReport.BackColor = vbCyan
                      
370               End If
380           Else
390               UpdateConsultantList SampleIDWithOffset, "Micro", RevertToLab, 0, 0
                  
400               If SampleAddedtoConsultantList(CStr(SampleIDWithOffset), "Micro") Then
410                   RemoveReport 0, CStr(SampleIDWithOffset), "N", 0
                      
420                   cmdReleaseReport.BackColor = vbButtonFace
                      
430               End If
440           End If
450       ElseIf cmdReleasetoWard.BackColor = vbGreen Then
460           If iMsg("Do you want to re-release report to ward?", vbYesNo) = vbYes Then
470               PrintThis ("SaveFinal")
480           Else
490               UpdateConsultantList SampleIDWithOffset, "Micro", RevertToLab, 0, 0
500               PrintThis "SaveBlank"
510           End If
520       End If
530       If ReleaseToHealthlinkOnValidate Then
540           If lblInterim.BackColor = vbGreen Then
550               ReleaseMicro SampleIDWithOffset, 1
560           Else
570               ReleaseMicro SampleIDWithOffset, 2
580           End If

590       End If

600       LockFraCS 1
610   Else
620       sql = "Select Password from Users where " & _
                "Name = '" & AddTicks(UserName) & "'"
630       Set tb = New Recordset
640       RecOpenServer 0, tb, sql
650       If Not tb.EOF Then
660           If UCase$(iBOX("Password Required", , , True)) = UCase$(tb!PassWord & "") Then
670               cmdValidateMicro.Caption = "&Validate"
680               UpdatePrintValidLog SampleIDWithOffset, Section, 0, 2
690               If ReleaseToHealthlinkOnValidate Then
700                   If lblInterim.BackColor = vbGreen Then
710                       ReleaseMicro SampleIDWithOffset, 1
720                   Else
730                       ReleaseMicro SampleIDWithOffset, 2
740                   End If
750               End If
760               LockFraCS 0
770           Else
780               Exit Sub
790           End If
800       Else
810           Exit Sub
820       End If
830   End If

840   Select Case SSTab1.Tab
      Case 1: SaveUrine
850   Case 2: SaveIdent
860   Case 3: SaveFaeces
870   Case 4: SaveIsolates
880       SaveSensitivities gYES
890   Case 5: SaveFOB
900   Case 6: SaveRotaAdeno
910   Case 7: SaveRedSub
920   Case 8: SaveRSV
930   Case 9: SaveFluids
940   Case 10: SaveCdiff
950   Case 11: SaveOP
960   Case 13: SaveHPylori
970   End Select
980   SaveComments
990   cmdSaveHold.Enabled = False
1000  cmdSaveMicro.Enabled = False
1010  UPDATEMRU txtSampleID, cMRU

1020  SaveSSTab = SSTab1.Tab

1030  If Validate And Section <> "CANDS" And Section <> "CSF" And Section <> "RSV" Then
1040      txtSampleID = Format$(Val(txtSampleID) + 1)
1050  End If


1060  LoadAllDetails

1070  If SaveSSTab > 0 Then
1080      If SSTab1.TabVisible(SaveSSTab) = True Then
1090          SSTab1.Tab = SaveSSTab
1100      End If
1110  End If

1120  cmdSaveDemographics.Enabled = False
1130  cmdSaveInc.Enabled = False
1140  cmdSaveMicro.Enabled = False
1150  cmdSaveHold.Enabled = False

1160  Exit Sub

cmdValidateMicro_Click_Error:

      Dim strES As String
      Dim intEL As Integer

1170  intEL = Erl
1180  strES = Err.Description
1190  LogError "frmEditMicrobiologyNew", "cmdValidateMicro_Click", intEL, strES, sql

End Sub




Private Sub LoadAllDetails()

      Dim WasTab     As Integer
      Dim v          As String
      Dim P          As String

10    On Error GoTo LoadAllDetails_Error

20    ForceSaveability = False

30    WasTab = SSTab1.Tab

40    SSTab1.TabCaption(1) = "Urine"
50    SSTab1.TabCaption(2) = "Identification"
60    SSTab1.TabCaption(3) = "Faeces"
70    SSTab1.TabCaption(4) = "C && S"
80    SSTab1.TabCaption(5) = "FOB"
90    SSTab1.TabCaption(6) = "Rota/Adeno"
100   SSTab1.TabCaption(7) = "Red/Sub"
110   SSTab1.TabCaption(8) = "RSV"
120   SSTab1.TabCaption(9) = "Fluids"
130   SSTab1.TabCaption(10) = "C.diff"
140   SSTab1.TabCaption(11) = "OP"
150   SSTab1.TabCaption(12) = "Blood Culture"
160   SSTab1.TabCaption(13) = "H.Pylori"
    
    
      'Zyam 12-12-23 Making Tab Visible
'      SSTab1.TabVisible(1) = True
'      SSTab1.TabVisible(2) = True
'      SSTab1.TabVisible(3) = True
'      SSTab1.TabVisible(4) = True
'      SSTab1.TabVisible(5) = True
'      SSTab1.TabVisible(6) = True
'      SSTab1.TabVisible(7) = True
'      SSTab1.TabVisible(8) = True
'      SSTab1.TabVisible(9) = True
'      SSTab1.TabVisible(10) = True
'      SSTab1.TabVisible(11) = True
'      SSTab1.TabVisible(12) = True
'      SSTab1.TabVisible(13) = True
      'Zyam 12-12-23 Making Tab Visible



170   ClearDemographics

180   LoadDemographics
190   CheckPatientNotePad (Trim$(txtSampleID))
200   CheckObserva

210   GetTabsFromSetUp

220   FaecesLoaded = False
230   UrineLoaded = False
240   IdentLoaded = False
250   CSLoaded = False
260   FOBLoaded = False
270   RotaAdenoLoaded = False
280   CdiffLoaded = False
290   OPLoaded = False
300   IdentificationLoaded = False
310   HPyloriLoaded = False
320   FluidsLoaded = False

330   ClearIndividualFaeces
340   ClearUrine
350   ClearIQ200
360   ClearFluid
370   ClearFaeces
380   ClearIdent
390   ClearCS

400   LoadFaecalOrders

410   If TabExistsForSite(cmbSite, UrineTab) Or SSTab1.TabVisible(UrineTab) = True Then
420       If LoadUrine() Then
430           LoadPrintValid "U", v, P
440           SSTab1.TabCaption(1) = "<<Urine>>" & v & P
450           SSTab1.TabVisible(1) = True
460           cmbSite.Enabled = False
470       End If
480   End If

490   If TabExistsForSite(cmbSite, FobTab) Or SSTab1.TabVisible(FobTab) = True Then
500       If LoadFOB() Then
510           LoadPrintValid "F", v, P
520           SSTab1.TabCaption(5) = "<<FOB>>" & v & P
530           SSTab1.TabVisible(5) = True
540           FOBLoaded = True
550           cmbSite.Enabled = False
560       End If
570   End If

580   If TabExistsForSite(cmbSite, RotaTab) Or SSTab1.TabVisible(RotaTab) = True Then
590       If LoadRotaAdeno() Then
600           LoadPrintValid "A", v, P
610           SSTab1.TabCaption(6) = "<<Rota/Adeno>>" & v & P
620           SSTab1.TabVisible(6) = True
630           RotaAdenoLoaded = True
640           cmbSite.Enabled = False
650       End If
660   End If

670   If TabExistsForSite(cmbSite, CDiffTab) Or SSTab1.TabVisible(CDiffTab) = True Then
680       If LoadCDiff() Then
690           LoadPrintValid "G", v, P
700           SSTab1.TabCaption(10) = "<<C.diff>>" & v & P
710           SSTab1.TabVisible(10) = True
720           CdiffLoaded = True
730           cmbSite.Enabled = False
740       End If
750   End If

760   If TabExistsForSite(cmbSite, OpTab) Or SSTab1.TabVisible(OpTab) = True Then
770       If LoadOP() Then
780           LoadPrintValid "O", v, P
790           SSTab1.TabCaption(11) = "<<OP>>" & v & P
800           SSTab1.TabVisible(11) = True
810           OPLoaded = True
820           cmbSite.Enabled = False
830       End If
840   End If

850   If TabExistsForSite(cmbSite, FaecesTab) Or SSTab1.TabVisible(FaecesTab) = True Then
860       If LoadFaeces() Then
870           SSTab1.TabCaption(3) = "<<Faeces>>"
880           SSTab1.TabVisible(3) = True
890           FaecesLoaded = True
900           cmbSite.Enabled = False
910       End If
920   End If

930   If TabExistsForSite(cmbSite, HPyloriTab) Or SSTab1.TabVisible(HPyloriTab) = True Then
940       If LoadHPylori() Then
950           LoadPrintValid "Y", v, P
960           SSTab1.TabCaption(13) = "<<H.Pylori>>" & v & P
970           SSTab1.TabVisible(13) = True
980           HPyloriLoaded = True
990           cmbSite.Enabled = False
1000      End If
1010  End If

1020  If TabExistsForSite(cmbSite, FluidsTab) Or SSTab1.TabVisible(FluidsTab) = True Then
1030      If LoadFluids() Then
1040          FluidsLoaded = True
1050          SSTab1.TabCaption(9) = "<<Fluids>>" & v & P
1060          cmbSite.Enabled = False
1070      End If
1080  End If

1090  If TabExistsForSite(cmbSite, BcTab) Or SSTab1.TabVisible(BcTab) = True Then
1100      If UCase$(HospName(0)) <> "PORTLAOISE" Then
1110          cmdBloodCulture(1).Visible = False
1120      End If
1130      If LoadBloodCulture() Then
1140          LoadPrintValid "B", v, P
1150          SSTab1.TabCaption(12) = "<<Blood Culture>>" & v & P
1160          SSTab1.TabVisible(12) = True
1170          cmbSite.Enabled = False
1180      End If

1190  End If

1200  IdentLoaded = False

1210  If TabExistsForSite(cmbSite, UrIdentTab) Or SSTab1.TabVisible(UrIdentTab) = True Then
1220      If LoadIdent() > 0 Then
1230          IdentLoaded = True
1240      End If
1250  End If

1260  If TabExistsForSite(cmbSite, CsTab) Or SSTab1.TabVisible(CsTab) = True Then

1270      If LoadIsolates() Then
1280          LoadPrintValid "D", v, P
1290          SSTab1.TabCaption(4) = "<<C && S>>" & v & P
1300          SSTab1.TabVisible(4) = True
1310          SetComboWidths
1320          cmbSite.Enabled = False

1330      End If
1340      If LoadSensitivities() = 0 Then
1350          CSLoaded = False
1360      Else
1370          CSLoaded = True
1380          cmbSite.Enabled = False
1390      End If
1400  End If

1410  LoadComments

1420  If TabExistsForSite(cmbSite, RsvTab) Or SSTab1.TabVisible(RsvTab) = True Then
1430      If LoadRSV() Then
1440          LoadPrintValid "V", v, P
1450          SSTab1.TabCaption(8) = "<<RSV>>" & v & P
1460          SSTab1.TabVisible(8) = True
1470          cmbSite.Enabled = False
1480      End If
1490  End If

1500  If TabExistsForSite(cmbSite, RsTab) Or SSTab1.TabVisible(RsTab) = True Then
1510      If LoadRedSub() Then
1520          LoadPrintValid "R", v, P
1530          SSTab1.TabCaption(7) = "<<R/S>>" & v & P
1540          SSTab1.TabVisible(7) = True
1550          cmbSite.Enabled = False
1560      End If
1570  End If

1580  FillHistoricalFaeces

      'FillForConsultantValidation

      'EnableCopyFrom

1590  CheckIfPhoned
      '
      '1140  SID = Val(txtSampleID) + SysOptMicroOffset(0)
      '1150  cmdCopySensitivities.Visible = False
      '1160  If IsAnyRecordPresent("Isolates", SID - 1) Then
      '1170    If IsAnyRecordPresent("Sensitivities", SID - 1) Then
      '1180      If Not IsAnyRecordPresent("Isolates", SID) Then
      '1190        If Not IsAnyRecordPresent("Sensitivities", SID) Then
      '1200          cmdCopySensitivities.Caption = "Copy from " & Val(txtSampleID) - 1
      '1210          cmdCopySensitivities.Visible = True
      '1220        End If
      '1230      End If
      '1240    End If
      '1250  End If

1600  If SSTab1.TabVisible(WasTab) Then
1610      SSTab1.Tab = WasTab 'shakeel
1620  End If

1630  CheckCC
1640  CheckExternals

1650  cmdArchive.Visible = True

1660  If ForceSaveability Then
1670      cmdSaveMicro.Enabled = True
1680      cmdSaveHold.Enabled = True
1690  End If

1700  ShowPrintValidFlags
1710  CheckValidStatus SSTab1.Tab

1720  ShowWhoSaved SSTab1.Tab

1730  txtSampleID.SelStart = 0
1740  txtSampleID.SelLength = 99

1750  cmdObserva(0).Visible = False
1760  cmdObserva(1).Visible = False
1770  cmdObserva(2).Visible = False
1780  If ObservaInUse Then
1790      If UCase$(cmbSite) = "URINE" Then
1800          cmdObserva(1).Visible = True
1810      ElseIf UCase$(cmbSite) <> "URINE" Then
1820          cmdObserva(0).Visible = True
1830          cmdObserva(2).Visible = True
1840      End If
1850  End If

1860  If InStr(1, UCase(cmbSite), "MRSA") > 0 Then
1870      txtNoCopies = GetOptionSetting("MRSADefaultCopies", 1)
1880  Else
1890      txtNoCopies = 1
1900  End If


1910  With lblFinal
1920      .BackColor = vbGreen
1930      .FontBold = True

1940  End With

1950  With lblInterim
1960      .BackColor = &H8000000F
1970      .FontBold = False
1980  End With

1990  CheckReportReleasetoWard (txtSampleID)
      'If IsMicroReleased(SampleIDWithOffset) Then
      '    Set cmdHealthLink.Picture = imgHGreen.Picture
      'Else
      '    Set cmdHealthLink.Picture = imgHRed.Picture
      'End If

2000  Exit Sub

LoadAllDetails_Error:

      Dim strES      As String
      Dim intEL      As Integer

2010  intEL = Erl
2020  strES = Err.Description
2030  LogError "frmEditMicrobiologyNew", "LoadAllDetails", intEL, strES

End Sub
Private Function LoadBloodCulture() As Boolean

      Dim tb As Recordset
      Dim sql As String
      Dim s As String
      Dim n As Integer
      Dim ANF As String
      Dim RetVal As Boolean

10    On Error GoTo LoadBloodCulture_Error

20    For n = 0 To 2
30        cmdVitek(n).Enabled = False
40        cmdVitek(n).Visible = False
50    Next

60    gBC.Rows = 2
70    gBC.AddItem ""
80    gBC.RemoveItem 1

90    If Val(txtSampleID) = 0 Then
100       LoadBloodCulture = False
110       Exit Function
120   End If

130   RetVal = False

140   gBC.Visible = False

150   sql = "SELECT * FROM BloodCultureResults WHERE " & _
            "SampleID = '" & txtSampleID + SysOptMicroOffset(0) & "' " & _
            "AND COALESCE(Result, '') <> '**' " & _
            "ORDER BY RunDateTime DESC"
160   Set tb = New Recordset
170   RecOpenServer 0, tb, sql
180   If Not tb.EOF Then
190       RetVal = True
200       Do While Not tb.EOF
210           s = Format$(tb!RunDateTime, "dd/MM/yy HH:nn") & vbTab & _
                  tb!BottleNumber & vbTab & _
                  tb!TypeOfTest & vbTab
220           Select Case tb!Result & ""
              Case "+": s = s & "Positive"
230           Case "-": s = s & "Negative"
240           Case "*": s = s & "Negative to date. Still under Test"
250           Case Else: s = s & "Unknown"
260           End Select
270           s = s & vbTab & _
                  tb!TTD & ""
280           gBC.AddItem s
290           gBC.Row = gBC.Rows - 1

300           Select Case UCase$(tb!TypeOfTest & "")
              Case GetOptionSetting("BcAerobicBottle", "BSA"): ANF = "A"
310           Case GetOptionSetting("BcAnarobicBottle", "BSN"): ANF = "N"
320           Case GetOptionSetting("BcFanBottle", "BFA"): ANF = "F"
330           Case Else: ANF = ""
340           End Select

350           Select Case IsVitekOrdered(ANF)
              Case -1
360               If tb!Result & "" = "+" Then
370                   cmdVitek(gBC.Row - 2).Visible = True
380                   cmdVitek(gBC.Row - 2).Enabled = True
390                   cmdVitek(gBC.Row - 2).Caption = "Order on Vitek"
400               Else
410                   cmdVitek(gBC.Row - 2).Visible = False
420               End If
430           Case 0
440               cmdVitek(gBC.Row - 2).Visible = True
450               cmdVitek(gBC.Row - 2).Enabled = False
460               cmdVitek(gBC.Row - 2).Caption = "Requested"
470           Case 1
480               cmdVitek(gBC.Row - 2).Visible = True
490               cmdVitek(gBC.Row - 2).Enabled = False
500               cmdVitek(gBC.Row - 2).Caption = "Ordered on Vitek"
510           Case 2
520               cmdVitek(gBC.Row - 2).Visible = True
530               cmdVitek(gBC.Row - 2).Enabled = False
540               cmdVitek(gBC.Row - 2).Caption = "Resulted"
550           End Select

560           gBC.RowHeight(gBC.Row) = 600
570           gBC.Col = 5
580           If tb!Valid Then
590               Set gBC.CellPicture = imgSquareTick.Picture
600           Else
610               Set gBC.CellPicture = imgSquareCross.Picture
620           End If
630           gBC.CellPictureAlignment = flexAlignCenterCenter
640           tb.MoveNext
650       Loop
660   End If

670   If Not RetVal Then
      '          sql = "IF EXISTS ( SELECT * FROM BloodCultureRequests WHERE " & _
      '                "            SampleID = '" & txtSampleID & "') " & _
      '                "  BEGIN " & _
      '                "    IF EXISTS ( SELECT * FROM BloodCultureResults WHERE " & _
      '                "                SampleID = '" & txtSampleID + SysOptMicroOffset(0) & "') " & _
      '                "      UPDATE BloodCultureResults " & _
      '                "      SET Result = '**', RunDateTime = getdate() " & _
      '                "      WHERE SampleID = '" & txtSampleID + SysOptMicroOffset(0) & "' " & _
      '                "    ELSE " & _
      '                "      INSERT INTO BloodCultureResults " & _
      '                "      (SampleID, RunDateTime, Result, Valid) VALUES " & _
      '                "      ('" & txtSampleID + SysOptMicroOffset(0) & "', getdate(), '**', 0) " & _
      '                "  END "
      '          Cnxn(0).Execute sql

680       sql = "SELECT DATEDIFF(HOUR,RequestedDateTime,GETDATE()) D " & _
                "FROM BloodCultureRequests " & _
                "WHERE SampleID = '" & txtSampleID & "'"

690       Set tb = New Recordset
700       Set tb = Cnxn(0).Execute(sql)

710       If Not tb.EOF Then
      '        s = Format$(Now, "dd/MM/yy HH:nn") & vbTab & _
      '            vbTab & vbTab & "Negative to date. Still under Test" & vbTab & _
      '            tb!D & " Hours"
720           s = "No info available" & vbTab & "No info available" & vbTab & "No info available" & vbTab & "No info available" & vbTab & "No info available"
730           gBC.AddItem s
740       End If

750   End If

760   If gBC.Rows > 2 Then
770       gBC.RemoveItem 1
780   End If
790   gBC.Visible = True

800   fraBC.Enabled = Not LoadLockStatus(12)

810   LoadBloodCulture = RetVal

820   Exit Function

LoadBloodCulture_Error:

      Dim strES As String
      Dim intEL As Integer

830   intEL = Erl
840   strES = Err.Description
850   LogError "frmEditMicrobiologyNew", "LoadBloodCulture", intEL, strES, sql
860   gBC.Visible = True
870   LoadBloodCulture = False

End Function
Private Sub GetTabsFromSetUp()

          Dim sql As String
          Dim tb As Recordset
          Dim n As Integer

10        On Error GoTo GetTabsFromSetUp_Error

20        sql = "SELECT * FROM MicroSetup WHERE " & _
                "Site = '" & cmbSite & "'"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        If Not tb.EOF Then
60            SSTab1.TabVisible(1) = tb!Urine
70            SSTab1.TabVisible(2) = tb!UrIdent
80            SSTab1.TabVisible(3) = tb!Faeces
90            SSTab1.TabVisible(4) = tb!cS
100           SSTab1.TabVisible(5) = tb!FOB
110           SSTab1.TabVisible(6) = tb!Rota
120           SSTab1.TabVisible(7) = tb!rs
130           SSTab1.TabVisible(8) = tb!RSV
140           SSTab1.TabVisible(9) = tb!Fluids
150           SSTab1.TabVisible(10) = tb!CDiff
160           SSTab1.TabVisible(11) = tb!OP
170           SSTab1.TabVisible(12) = tb!BC
180           SSTab1.TabVisible(13) = tb!HPylori
190       Else
200           For n = 1 To 13
210               SSTab1.TabVisible(n) = False 'shakeel
220           Next
                SSTab1.TabVisible(4) = True 'CS
230       End If

240       Exit Sub

GetTabsFromSetUp_Error:

          Dim strES As String
          Dim intEL As Integer

250       intEL = Erl
260       strES = Err.Description
270       LogError "frmEditMicrobiologyNew", "GetTabsFromSetUp", intEL, strES, sql

End Sub


Private Sub bcancel_Click()

10        pBar = 0

20        Unload Me

End Sub


Private Sub ClearDemographics()

          Dim TimeNow As String

10        On Error GoTo ClearDemographics_Error

20        dtRunDate = Format$(Now, "dd/mm/yyyy")
30        dtSampleDate = Format$(Now, "dd/mm/yyyy")
40        dtRecDate = dtSampleDate

50        TimeNow = Format$(Now, "HH:nn")
60        tRecTime.Mask = ""
70        tRecTime.Text = TimeNow
80        tRecTime.Mask = "##:##"
90        tSampleTime.Mask = ""
          'tSampleTime.Text = TimeNow
100       tSampleTime.Mask = "##:##"

110       txtChart = ""
120       txtAandE = ""
130       txtName = ""
140       taddress(0) = ""
150       taddress(1) = ""
160       txtSex = ""
170       txtDoB = ""
180       txtAge = ""
190       cmbWard = "GP"
200       cmbClinician = ""
210       cmbGP = ""
220       cmbHospital = HospName(0)
230       txtClinDetails = ""
240       txtDemographicComment = ""
250       lblChartNumber.Caption = HospName(0) & " Chart #"
260       lblChartNumber.BackColor = &H8000000F
270       lblChartNumber.ForeColor = vbBlack
280       chkPregnant = 0
290       StatusBar1.Panels(4).Text = ""
300       cmdObserva(0).Caption = "Request on Observa"
310       cmdObserva(1).Caption = "Request on Observa"
320       cmdObserva(2).Caption = "Request on Observa"

330       lblAddWardGP = ""

340       EnableDemographicEntry True

350       Exit Sub

ClearDemographics_Error:

          Dim strES As String
          Dim intEL As Integer

360       intEL = Erl
370       strES = Err.Description
380       LogError "frmEditMicrobiologyNew", "ClearDemographics", intEL, strES

End Sub
Private Sub cmbClinDetails_Click()

10        txtClinDetails = txtClinDetails & cmbClinDetails & " "
20        cmbClinDetails.ListIndex = -1

30        cmdSaveDemographics.Enabled = True
40        cmdSaveInc.Enabled = True

End Sub


Private Sub cmbClinDetails_LostFocus()

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo cmbClinDetails_LostFocus_Error

20        pBar = 0

30        If Trim$(cmbClinDetails) = "" Then Exit Sub

40        sql = "Select * from Lists where " & _
                "ListType = 'CD' " & _
                "and Code = '" & cmbClinDetails & "'"
50        Set tb = New Recordset
60        RecOpenServer 0, tb, sql
70        If Not tb.EOF Then
80            cmbClinDetails = tb!Text & ""
90        End If
      '100       cmbClinDetails.Text = QueryCombo(cmbClinDetails)

100       Exit Sub

cmbClinDetails_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

110       intEL = Erl
120       strES = Err.Description
130       LogError "frmEditMicrobiologyNew", "cmbClinDetails_LostFocus", intEL, strES, sql


End Sub


Private Sub cmbClinician_Click()

10        cmdSaveDemographics.Enabled = True
20        cmdSaveInc.Enabled = True

End Sub

Private Sub cmbClinician_KeyPress(KeyAscii As Integer)

      '10        KeyAscii = AutoComplete(cmbClinician, KeyAscii, False)

10        cmdSaveDemographics.Enabled = True
20        cmdSaveInc.Enabled = True

End Sub


Private Sub cmbClinician_LostFocus()

10        pBar = 0
20        cmbClinician = QueryKnown("Clin", cmbClinician, cmbHospital)

End Sub

Private Sub cmbGP_Change()

10        SetWardClinGP

20        cmbWard = "GP"

End Sub

Private Sub SetWardClinGP()

          Dim GPAddr As String

10        On Error GoTo SetWardClinGP_Error

20        GPAddr = AddressOfGP(cmbGP)

30        lblAddWardGP = Trim$(taddress(0)) & " " & Trim$(taddress(1)) & " : " & cmbWard & " : " & cmbGP & ":" & GPAddr & " " & cmbClinician

40        Exit Sub

SetWardClinGP_Error:

          Dim strES As String
          Dim intEL As Integer



50        intEL = Erl
60        strES = Err.Description
70        LogError "frmEditMicrobiologyNew", "SetWardClinGP", intEL, strES

End Sub

Private Sub cmbGP_Click()

10        pBar = 0

20        SetWardClinGP

30        cmbWard = "GP"
40        cmdSaveDemographics.Enabled = True
50        cmdSaveInc.Enabled = True

End Sub


Private Sub cmbGP_KeyPress(KeyAscii As Integer)

      '10        KeyAscii = AutoComplete(cmbGP, KeyAscii, False)

10        cmdSaveDemographics.Enabled = True
20        cmdSaveInc.Enabled = True

End Sub


Private Sub cmbGP_LostFocus()

10        cmbGP = QueryKnown("GP", cmbGP, cmbHospital)

End Sub



Private Sub cmdSetPrinter_Click()

10        On Error GoTo cmdSetPrinter_Click_Error

20        Set frmForcePrinter.f = frmEditMicrobiologyNew
30        frmForcePrinter.Show 1

40        If pPrintToPrinter = "Automatic Selection" Then
50            pPrintToPrinter = ""
60        End If

70        If pPrintToPrinter <> "" Then
80            cmdSetPrinter.BackColor = vbRed
90            cmdSetPrinter.ToolTipText = "Print Forced to " & pPrintToPrinter
100       Else
110           cmdSetPrinter.BackColor = vbButtonFace
120           pPrintToPrinter = ""
130           cmdSetPrinter.ToolTipText = "Printer Selected Automatically"
140       End If

150       Exit Sub

cmdSetPrinter_Click_Error:

          Dim strES As String
          Dim intEL As Integer



160       intEL = Erl
170       strES = Err.Description
180       LogError "frmEditMicrobiologyNew", "cmdSetPrinter_Click", intEL, strES


End Sub

Private Sub cMRU_Click()

10        txtSampleID = cMRU

20        GetSampleIDWithOffset

30        LoadAllDetails

40        cmdSaveDemographics.Enabled = False
50        cmdSaveInc.Enabled = False
60        cmdSaveMicro.Enabled = False
70        cmdSaveHold.Enabled = False

End Sub


Private Sub cMRU_KeyPress(KeyAscii As Integer)

10        KeyAscii = 0

End Sub


Private Sub cmbWard_Change()

10        SetWardClinGP

End Sub

Private Sub cmbWard_Click()

10        SetWardClinGP

20        cmdSaveDemographics.Enabled = True
30        cmdSaveInc.Enabled = True

End Sub


Private Sub cmbWard_KeyPress(KeyAscii As Integer)

      '10        KeyAscii = AutoComplete(cmbWard, KeyAscii, False)

10        cmdSaveDemographics.Enabled = True
20        cmdSaveInc.Enabled = True

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

70        sql = "SELECT * from wards WHERE " & _
                "(text = '" & AddTicks(cmbWard) & "' " & _
                "or code = '" & AddTicks(cmbWard) & "') " & _
                "and hospitalcode = '" & ListCodeFor("HO", cmbHospital) & "' And InUse = '1'"
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

180       intEL = Erl
190       strES = Err.Description
200       LogError "frmEditMicrobiologyNew", "cmbWard_LostFocus", intEL, strES, sql

End Sub



Private Sub dtRunDate_CloseUp()

10        pBar = 0

20        cmdSaveDemographics.Enabled = True
30        cmdSaveInc.Enabled = True

End Sub


Private Sub dtSampleDate_CloseUp()

10        pBar = 0

20        cmdSaveDemographics.Enabled = True
30        cmdSaveInc.Enabled = True

End Sub


Private Sub Form_Activate()

10        TimerBar.Enabled = True
20        pBar = 0

End Sub

Private Sub gBC_Click()

          Dim sql As String
          Dim v As Integer

10        On Error GoTo gBC_Click_Error

20        If gBC.MouseRow = 0 Then Exit Sub
30        If gBC.Col = 5 Then

40            If gBC.CellPicture = imgSquareCross.Picture Then
50                Set gBC.CellPicture = imgSquareTick.Picture
60                v = 1
70            Else
80                Set gBC.CellPicture = imgSquareCross.Picture
90                v = 0
100           End If

110           sql = "UPDATE BloodCultureResults " & _
                    "SET Valid = " & v & " " & _
                    "WHERE SampleID = '" & txtSampleID + SysOptMicroOffset(0) & "' " & _
                    "AND BottleNumber = '" & gBC.TextMatrix(gBC.Row, 1) & "'"
120           Cnxn(0).Execute sql

130       End If

140       Exit Sub

gBC_Click_Error:

          Dim strES As String
          Dim intEL As Integer

150       intEL = Erl
160       strES = Err.Description
170       LogError "frmEditMicrobiologyNew", "gBC_Click", intEL, strES, sql

End Sub

 Sub ValidateDemographics(ByVal Validate As Boolean)

          Dim sql As String
          Dim tb As New Recordset

10        On Error GoTo ValidateDemographics_Error

20        If Validate Then
30            If cmdSaveDemographics.Enabled Then
40                SaveDemographics
50            End If
60            sql = "SELECT * FROM Demographics WHERE " & _
                    "SampleID = '" & Val(txtSampleID) + SysOptMicroOffset(0) & "'"
70            Set tb = New Recordset
80            RecOpenServer 0, tb, sql
90            If Not tb.EOF Then
100               sql = "UPDATE Demographics SET Valid = 1, " & _
                        "UserName = '" & AddTicks(UserName) & "' WHERE " & _
                        "SampleID = '" & Val(txtSampleID) + SysOptMicroOffset(0) & "'"
110               Cnxn(0).Execute sql
120               EnableDemographicEntry False
130               cmdDemoVal.Caption = "VALID"
140               cmdSaveDemographics.Enabled = False
150               cmdSaveInc.Enabled = False
160           End If
170       Else
180           If UCase(iBOX("Enter password to unValidate ?", , , True)) = UserPass Then
190               sql = "SELECT * FROM Demographics WHERE " & _
                        "SampleID = '" & Val(txtSampleID) + SysOptMicroOffset(0) & "'"

200               Set tb = New Recordset
210               RecOpenServer 0, tb, sql
220               If Not tb.EOF Then
230                   sql = "UPDATE Demographics SET valid = 0, " & _
                            "UserName = '" & AddTicks(UserName) & "' WHERE " & _
                            "SampleID = '" & Val(txtSampleID) + SysOptMicroOffset(0) & "'"
240                   Cnxn(0).Execute sql
250                   EnableDemographicEntry True
260                   cmdDemoVal.Caption = "&Validate"
270               End If
280           End If
290       End If

300       Exit Sub

ValidateDemographics_Error:

          Dim strES As String
          Dim intEL As Integer

310       intEL = Erl
320       strES = Err.Description
330       LogError "frmEditMicrobiologyNew", "ValidateDemographics", intEL, strES, sql

End Sub

Private Sub EnableDemographicEntry(ByVal Enable As Boolean)

10        On Error GoTo EnableDemographicEntry_Error

20        cmdAddToConsultantList.Enabled = True    'Enable
30        cmbConsultantVal.Enabled = True    'Enable
40        lblChartNumber.Enabled = Enable
50        txtChart.Enabled = Enable
60        txtAandE.Enabled = Enable
70        txtName.Enabled = Enable
80        txtDoB.Enabled = Enable
90        txtAge.Enabled = Enable
100       txtSex.Enabled = Enable
110       lblABsInUse.Enabled = Enable
120       lblAddWardGP.Enabled = Enable
130       lblSiteDetails.Enabled = Enable
140       Label44.Enabled = Enable
150       bsearch.Enabled = Enable
160       bDoB.Enabled = Enable

170       Frame4.Enabled = Enable    'Hosp/Ward/Clinician
180       fraDate.Enabled = Enable    'Dates
190       Frame5.Enabled = Enable    'Routine/Out of Hours
200       Frame12.Enabled = Enable    'Site Details
210       Frame13.Enabled = Enable    'Current Antibodics
220       Frame14.Enabled = Enable    'Clinical Details
          'cmdCopyFromPrevious.Enabled = Enable

230       If Not Enable Then
240           StatusBar1.Panels(3).Text = "Demographics Validated"
250           StatusBar1.Panels(3).Bevel = sbrInset
260       Else
270           StatusBar1.Panels(3).Text = "Check Demographics"
280           StatusBar1.Panels(3).Bevel = sbrRaised
290       End If

300       Exit Sub

EnableDemographicEntry_Error:

          Dim strES As String
          Dim intEL As Integer

310       intEL = Erl
320       strES = Err.Description
330       LogError "frmEditMicrobiologyNew", "EnableDemographicEntry", intEL, strES

End Sub

Private Sub FillOrganismGroups()

          Dim n As Integer
          Dim tb As Recordset
          Dim sql As String
          Dim temp As String

10        On Error GoTo FillOrganisms_Error

20        sql = "Select * from Lists where " & _
                "ListType = 'OR' " & _
                "order by ListOrder"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql

50        For n = 1 To 4
60            cmbOrgGroup(n).Clear
70            cmbOrgName(n).Clear
80        Next

90        Do While Not tb.EOF
100           temp = tb!Text & ""
110           For n = 1 To 4
120               cmbOrgGroup(n).AddItem temp
130           Next
140           tb.MoveNext
150       Loop

160       SetComboWidths

170       Exit Sub

FillOrganisms_Error:

          Dim strES As String
          Dim intEL As Integer

180       intEL = Erl
190       strES = Err.Description
200       LogError "frmEditMicrobiologyNew", "FillOrganisms", intEL, strES, sql


End Sub

Private Sub ClearFluid()

          Dim n As Integer

10        On Error GoTo ClearFluid_Error

20        cmdLock(9).Visible = False
30        fraCSF.Enabled = True
40        cmbFluidAppearance(0) = ""
50        cmbFluidAppearance(1) = ""
60        cmbFluidGram(0) = ""
70        cmbFluidGram(1) = ""
80        cmbFluidLeishmans = ""
90        cmbFluidWetPrep = ""
100       cmbFluidCrystals = ""
110       cmbZN = ""

120       For n = 0 To 11
130           txtHaem(n) = ""
140       Next
          'txtFluidComment = ""
150       txtInHouseSID = ""
160       For n = 0 To 7
170           chkBio(n).Value = 0
180           txtBioResult(n) = ""
190       Next

200       lblPneuAT = ""
210       lblLegionellaAT = ""
220       chkFungal(0).Value = 0
230       chkFungal(1).Value = 0
240       lblBATResult.Caption = ""
250       txtBATComments = ""

260       Exit Sub

ClearFluid_Error:

          Dim strES As String
          Dim intEL As Integer

270       intEL = Erl
280       strES = Err.Description
290       LogError "frmEditMicrobiologyNew", "ClearFluid", intEL, strES

End Sub

Private Sub ClearFaeces()

          Dim x As Integer
          Dim Y As Integer

10        On Error GoTo ClearFaeces_Error

20        For x = 1 To 3
30            For Y = 1 To 4
40                cmbDay1(Y * 10 + x) = ""
50                cmbDay2(Y * 10 + x) = ""
60            Next
70            cmbDay3(x) = ""
80        Next

90        cmbDay2(51) = ""
100       cmbDay2(52) = ""
110       cmbDay2(53) = ""

120       cmbDay3(4) = ""
130       cmbDay3(5) = ""
140       cmbDay3(6) = ""

150       Exit Sub

ClearFaeces_Error:

          Dim strES As String
          Dim intEL As Integer

160       intEL = Erl
170       strES = Err.Description
180       LogError "frmEditMicrobiologyNew", "ClearFaeces", intEL, strES


End Sub

Private Sub ClearCS()

          Dim intIsolate As Integer

10        On Error GoTo ClearCS_Error
20        lblCells.Caption = ""
30        For intIsolate = 1 To 4
40            cmbOrgGroup(intIsolate) = ""
50            cmbOrgName(intIsolate) = ""
60            cmbQualifier(intIsolate) = ""
70            chkNonReportable(intIsolate - 1).Value = 0
80            cmbABSelect(intIsolate) = ""
90            grdAB(intIsolate).Clear
100           grdAB(intIsolate).FormatString = "<Antibiotic             |^S/R|>Rprt|<Result  |<Date/Time            |<Operator"
110       Next
120       txtMSC = ""
130       txtConC = ""

140       Exit Sub

ClearCS_Error:

          Dim strES As String
          Dim intEL As Integer

150       intEL = Erl
160       strES = Err.Description
170       LogError "frmEditMicrobiologyNew", "ClearCS", intEL, strES

End Sub


Private Sub FillCastsCrystalsMiscSite()

          Dim n As Integer
          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo FillCastsCrystalsMiscSite_Error

20        cmbCasts.Clear
30        cmbCrystals.Clear
40        cmbMisc(0).Clear
50        cmbMisc(1).Clear
60        cmbMisc(2).Clear
70        cmbSite.Clear
80        cmbSiteSearch.Clear
90        cmbClinDetails.Clear

100       sql = "SELECT Text FROM Lists WHERE " & _
                "ListType = 'CA' ORDER BY ListOrder"
110       Set tb = New Recordset
120       RecOpenServer 0, tb, sql
130       Do While Not tb.EOF
140           cmbCasts.AddItem tb!Text & ""
150           tb.MoveNext
160       Loop

170       sql = "SELECT Text FROM Lists WHERE " & _
                "ListType = 'CR' ORDER BY ListOrder"
180       Set tb = New Recordset
190       RecOpenServer 0, tb, sql
200       Do While Not tb.EOF
210           cmbCrystals.AddItem tb!Text & ""
220           tb.MoveNext
230       Loop

240       sql = "SELECT Text FROM Lists WHERE " & _
                "ListType = 'MI' ORDER BY ListOrder"
250       Set tb = New Recordset
260       RecOpenServer 0, tb, sql
270       Do While Not tb.EOF
280           cmbMisc(0).AddItem tb!Text & ""
290           cmbMisc(1).AddItem tb!Text & ""
300           cmbMisc(2).AddItem tb!Text & ""
310           tb.MoveNext
320       Loop

330       FillSites

340       sql = "SELECT Text FROM Lists WHERE " & _
                "ListType = 'OV' ORDER BY ListOrder"
350       Set tb = New Recordset
360       RecOpenServer 0, tb, sql
370       Do While Not tb.EOF
380           cmbOva(0).AddItem tb!Text & ""
390           cmbOva(1).AddItem tb!Text & ""
400           cmbOva(2).AddItem tb!Text & ""
410           tb.MoveNext
420       Loop

430       sql = "Select Text FROM Lists WHERE " & _
                "ListType = 'CD' ORDER BY ListOrder"
440       Set tb = New Recordset
450       RecOpenServer 0, tb, sql
460       Do While Not tb.EOF
470           cmbClinDetails.AddItem tb!Text & ""
480           tb.MoveNext
490       Loop

500       sql = "SELECT Text FROM Lists WHERE " & _
                "ListType = 'HO' ORDER BY ListOrder"
510       Set tb = New Recordset
520       RecOpenServer 0, tb, sql
530       Do While Not tb.EOF
540           cmbHospital.AddItem tb!Text & ""
550           tb.MoveNext
560       Loop

570       sql = "SELECT Text FROM Lists WHERE " & _
                "ListType = 'GS' ORDER BY ListOrder"
580       Set tb = New Recordset
590       RecOpenServer 0, tb, sql
600       Do While Not tb.EOF
610           For n = 1 To 4
620               cmbGram(n).AddItem tb!Text & ""
630           Next
640           tb.MoveNext
650       Loop

660       sql = "SELECT Text FROM Lists WHERE " & _
                "ListType = 'WP' ORDER BY ListOrder"
670       Set tb = New Recordset
680       RecOpenServer 0, tb, sql
690       Do While Not tb.EOF
700           For n = 1 To 4
710               cmbWetPrep(n).AddItem tb!Text & ""
720           Next
730           tb.MoveNext
740       Loop


750       FixComboWidth cmbCasts
760       FixComboWidth cmbCrystals
770       FixComboWidth cmbMisc(0)
780       FixComboWidth cmbMisc(1)
790       FixComboWidth cmbMisc(2)
800       FixComboWidth cmbSite
810       FixComboWidth cmbSiteSearch
820       FixComboWidth cmbClinDetails
830       For n = 1 To 4
840           FixComboWidth cmbGram(n)
850           FixComboWidth cmbWetPrep(n)
860       Next n

870       Exit Sub

FillCastsCrystalsMiscSite_Error:

          Dim strES As String
          Dim intEL As Integer

880       intEL = Erl
890       strES = Err.Description
900       LogError "frmEditMicrobiologyNew", "FillCastsCrystalsMiscSite", intEL, strES, sql

End Sub
Private Sub LoadListQualifier()

          Dim n As Integer
          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo LoadListQualifier_Error

20        For n = 1 To 4
30            cmbQualifier(n).Clear
40        Next

50        sql = "SELECT Text FROM Lists WHERE " & _
                "ListType = 'MQ' ORDER BY ListOrder"
60        Set tb = New Recordset
70        RecOpenServer 0, tb, sql
80        Do While Not tb.EOF
90            For n = 1 To 4
100               cmbQualifier(n).AddItem tb!Text & ""
110           Next
120           tb.MoveNext
130       Loop

140       For n = 1 To 4
150           FixComboWidth cmbQualifier(n)
160       Next

170       Exit Sub

LoadListQualifier_Error:

          Dim strES As String
          Dim intEL As Integer

180       intEL = Erl
190       strES = Err.Description
200       LogError "frmEditMicrobiologyNew", "LoadListQualifier", intEL, strES, sql

End Sub

Private Sub FillFaecesLists()

10        LoadListXLD
20        LoadListDCA
30        LoadListSMAC
40        LoadListCROMO
50        LoadListCAMP
60        LoadListSTEC

End Sub

Private Sub LoadListCAMP()

          Dim tb As Recordset
          Dim sql As String
          Dim x As Integer

10        On Error GoTo LoadListCAMP_Error

20        cmbDay2(31) = ""
30        cmbDay2(32) = ""
40        cmbDay2(33) = ""

50        cmbDay3(1) = ""
60        cmbDay3(2) = ""
70        cmbDay3(3) = ""

80        sql = "SELECT Text FROM Lists WHERE " & _
                "ListType = 'FaecesCAMP' " & _
                "ORDER BY ListOrder"
90        Set tb = New Recordset
100       RecOpenServer 0, tb, sql
110       Do While Not tb.EOF
120           For x = 1 To 3
130               cmbDay2(30 + x).AddItem tb!Text & ""
140               cmbDay3(x).AddItem tb!Text & ""
150           Next
160           tb.MoveNext
170       Loop

180       Exit Sub

LoadListCAMP_Error:

          Dim strES As String
          Dim intEL As Integer

190       intEL = Erl
200       strES = Err.Description
210       LogError "frmEditMicrobiologyNew", "LoadListCAMP", intEL, strES, sql

End Sub

Private Sub LoadListCROMO()

          Dim tb As Recordset
          Dim sql As String
          Dim x As Integer

10        On Error GoTo LoadListCROMO_Error

20        cmbDay2(21) = ""
30        cmbDay2(22) = ""
40        cmbDay2(23) = ""

50        cmbDay3(4) = ""
60        cmbDay3(5) = ""
70        cmbDay3(6) = ""

80        sql = "SELECT Text FROM Lists WHERE " & _
                "ListType = 'FaecesCROMO' " & _
                "ORDER BY ListOrder"
90        Set tb = New Recordset
100       RecOpenServer 0, tb, sql
110       Do While Not tb.EOF
120           For x = 1 To 3
130               cmbDay2(20 + x).AddItem tb!Text & ""
140               cmbDay3(3 + x).AddItem tb!Text & ""
150           Next
160           tb.MoveNext
170       Loop

180       Exit Sub

LoadListCROMO_Error:

          Dim strES As String
          Dim intEL As Integer

190       intEL = Erl
200       strES = Err.Description
210       LogError "frmEditMicrobiologyNew", "LoadListCROMO", intEL, strES, sql

End Sub

Private Sub LoadListSTEC()

          Dim tb As Recordset
          Dim sql As String
          Dim x As Integer

10        On Error GoTo LoadListSTEC_Error


20        For x = 1 To 3
30            cmbDay1(40 + x).Clear
40        Next

50        sql = "SELECT Text FROM Lists WHERE " & _
                "ListType = 'FaecesSTEC1' " & _
                "ORDER BY ListOrder"
60        Set tb = New Recordset
70        RecOpenServer 0, tb, sql
80        Do While Not tb.EOF
90            For x = 1 To 3
100               cmbDay1(40 + x).AddItem tb!Text & ""
110           Next
120           tb.MoveNext
130       Loop

140       For x = 1 To 3
150           cmbDay2(50 + x).Clear
160       Next

170       sql = "SELECT Text FROM Lists WHERE " & _
                "ListType = 'FaecesSTEC2' " & _
                "ORDER BY ListOrder"
180       Set tb = New Recordset
190       RecOpenServer 0, tb, sql
200       Do While Not tb.EOF
210           For x = 1 To 3
220               cmbDay2(50 + x).AddItem tb!Text & ""
230           Next
240           tb.MoveNext
250       Loop


260       Exit Sub

LoadListSTEC_Error:

          Dim strES As String
          Dim intEL As Integer

270       intEL = Erl
280       strES = Err.Description
290       LogError "frmEditMicrobiologyNew", "LoadListSTEC", intEL, strES, sql

End Sub

Private Sub LoadListSMAC()

          Dim tb As Recordset
          Dim sql As String
          Dim x As Integer

10        On Error GoTo LoadListSMAC_Error

20        For x = 1 To 3
30            cmbDay1(30 + x).Clear
40        Next

50        sql = "SELECT Text FROM Lists WHERE " & _
                "ListType = 'FaecesSMAC' " & _
                "ORDER BY ListOrder"
60        Set tb = New Recordset
70        RecOpenServer 0, tb, sql
80        Do While Not tb.EOF
90            For x = 1 To 3
100               cmbDay1(30 + x).AddItem tb!Text & ""
110           Next
120           tb.MoveNext
130       Loop

140       Exit Sub

LoadListSMAC_Error:

          Dim strES As String
          Dim intEL As Integer

150       intEL = Erl
160       strES = Err.Description
170       LogError "frmEditMicrobiologyNew", "LoadListSMAC", intEL, strES, sql

End Sub


Private Sub LoadListDCA()

          Dim tb As Recordset
          Dim sql As String
          Dim x As Integer

10        On Error GoTo LoadListDCA_Error

20        cmbDay1(21) = ""
30        cmbDay1(22) = ""
40        cmbDay1(23) = ""

50        cmbDay2(41) = ""
60        cmbDay2(42) = ""
70        cmbDay2(43) = ""

80        sql = "SELECT Text FROM Lists WHERE " & _
                "ListType = 'FaecesDCA' " & _
                "ORDER BY ListOrder"
90        Set tb = New Recordset
100       RecOpenServer 0, tb, sql
110       Do While Not tb.EOF
120           For x = 1 To 3
130               cmbDay1(20 + x).AddItem tb!Text & ""
140               cmbDay2(40 + x).AddItem tb!Text & ""
150           Next
160           tb.MoveNext
170       Loop

180       Exit Sub

LoadListDCA_Error:

          Dim strES As String
          Dim intEL As Integer

190       intEL = Erl
200       strES = Err.Description
210       LogError "frmEditMicrobiologyNew", "LoadListDCA", intEL, strES, sql

End Sub


Private Sub LoadListXLD()

          Dim tb As Recordset
          Dim sql As String
          Dim x As Integer

10        On Error GoTo LoadListXLD_Error

20        For x = 1 To 3
30            cmbDay1(10 + x).Clear
40            cmbDay2(10 + x).Clear
50        Next

60        sql = "SELECT Text FROM Lists WHERE " & _
                "ListType = 'FaecesXLD' " & _
                "ORDER BY ListOrder"
70        Set tb = New Recordset
80        RecOpenServer 0, tb, sql
90        Do While Not tb.EOF
100           For x = 1 To 3
110               cmbDay1(10 + x).AddItem tb!Text & ""
120               cmbDay2(10 + x).AddItem tb!Text & ""
130           Next
140           tb.MoveNext
150       Loop

160       Exit Sub

LoadListXLD_Error:

          Dim strES As String
          Dim intEL As Integer

170       intEL = Erl
180       strES = Err.Description
190       LogError "frmEditMicrobiologyNew", "LoadListXLD", intEL, strES, sql

End Sub


Private Sub FillAbGrid(ByVal Index As Integer)

          Dim tb As Recordset
          Dim sql As String
          Dim ReportCounter As Integer
          Dim n As Integer
          Dim Y As Integer
          Dim Found As Boolean
          Dim s As String

10        On Error GoTo FillAbGrid_Error

20        With grdAB(Index)
30            .Visible = False
40            .Rows = 2
50            .AddItem ""
60            .RemoveItem 1
70        End With

80        ReportCounter = 0

90        sql = "Select Distinct A.AntibioticName, D.ListOrder, " & _
                "COALESCE(D.AutoReport, 0) AutoReport, COALESCE(D.AutoReportIf,'') AutoReportIf, " & _
                "COALESCE(AutoPriority,0) AutoPriority, D.ListOrder, " & _
                "A.AllowIfPregnant, A.AllowIfOutPatient, A.AllowIfChild " & _
                "from ABDefinitions as D, Antibiotics as A where " & _
                "D.OrganismGroup = '" & cmbOrgGroup(Index) & "' " & _
                "and D.Site = '" & cmbSite & "' " & _
                "and D.PriSec = 'P' " & _
                "and D.AntibioticName = A.AntibioticName " & _
                "order by D.ListOrder"


100       Set tb = New Recordset
110       RecOpenClient 0, tb, sql
120       If tb.EOF Then
130           sql = "Select Distinct A.AntibioticName, D.ListOrder, " & _
                    "COALESCE(D.AutoReport, 0) AutoReport, COALESCE(D.AutoReportIf,'') AutoReportIf, " & _
                    "COALESCE(AutoPriority,0) AutoPriority, D.ListOrder, " & _
                    "A.AllowIfPregnant, A.AllowIfOutPatient, A.AllowIfChild " & _
                    "from ABDefinitions as D, Antibiotics as A where " & _
                    "D.OrganismGroup = '" & cmbOrgGroup(Index) & "' " & _
                    "and Site = 'Generic' " & _
                    "and D.PriSec = 'P' " & _
                    "and D.AntibioticName = A.AntibioticName " & _
                    "order by D.ListOrder"
140           Set tb = New Recordset
150           RecOpenClient 0, tb, sql
160           If tb.EOF Then
                  ' iMsg "Site/Organism not defined.", vbCritical
170               grdAB(Index).Visible = True
180               Exit Sub
190           End If
200       End If

210       Do While Not tb.EOF
220           s = Trim$(tb!AntibioticName) & vbTab & _
                  vbTab & _
                  vbTab & _
                  vbTab & _
                  vbTab & _
                  vbTab & _
                  tb!AutoReport & vbTab & _
                  tb!AutoReportIf & vbTab & _
                  tb!AutoPriority & vbTab & _
                  tb!ListOrder

230           grdAB(Index).AddItem s
240           grdAB(Index).Row = grdAB(Index).Rows - 1
250           grdAB(Index).Col = 2

260           If IsChild() And Not tb!AllowIfChild = 1 Then
270               Set grdAB(Index).CellPicture = imgSquareCross.Picture
280               grdAB(Index).TextMatrix(grdAB(Index).Row, 2) = "C"
290           ElseIf IsPregnant() And Not tb!AllowIfPregnant = 1 Then
300               Set grdAB(Index).CellPicture = imgSquareCross.Picture
310               grdAB(Index).TextMatrix(grdAB(Index).Row, 2) = "P"
320           ElseIf IsOutPatient() And Not tb!AllowIfOutPatient = 1 Then
330               Set grdAB(Index).CellPicture = imgSquareCross.Picture
340               grdAB(Index).TextMatrix(grdAB(Index).Row, 2) = "O"
350           Else
360               Set grdAB(Index).CellPicture = Me.Picture
370           End If
380           tb.MoveNext

390       Loop

400       For n = 0 To lstABsInUse.ListCount - 1
410           If lstABsInUse.List(n) <> "Antibiotic Not Stated" And lstABsInUse.List(n) <> "None" Then
420               Found = False
430               For Y = 1 To grdAB(Index).Rows - 1
440                   If grdAB(Index).TextMatrix(Y, 0) = lstABsInUse.List(n) Then
450                       Found = True
460                       Exit For
470                   End If
480               Next
490               If Not Found Then
500                   grdAB(Index).AddItem lstABsInUse.List(n)
510                   grdAB(Index).Row = grdAB(Index).Rows - 1
520               Else
530                   grdAB(Index).Row = Y
540               End If
550               grdAB(Index).Col = 2
560               Set grdAB(Index).CellPicture = imgSquareTick.Picture
570           End If
580       Next




590       If grdAB(Index).Rows > 2 Then
600           grdAB(Index).RemoveItem 1
610       End If
620       grdAB(Index).Visible = True

630       Exit Sub

FillAbGrid_Error:

          Dim strES As String
          Dim intEL As Integer

640       intEL = Erl
650       strES = Err.Description
660       LogError "frmEditMicrobiology", "FillAbGrid", intEL, strES, sql

End Sub

Function CheckConflict() As Boolean

          Dim s As String
          Dim Conflict As Boolean
          Dim sn As Recordset
          Dim tb As Recordset
          Dim sql As String
          Dim Organism(0 To 1) As String
          Dim n As Integer
          Dim OrgWas As String
          Dim OrgIs As String
          Dim ConflictList As String
          Dim grid As Integer
          Dim Org As String
          Dim ThisRunNumber As String
          Dim SampleDate As String

10        On Error GoTo CheckConflict_Error

20        If Trim(txtChart) = "" Then
30            CheckConflict = False
40            Exit Function
50        End If

60        sql = "select top 1 * from demographics where " & _
                "ForUrine = 1 and " & _
                "chart = '" & txtChart & "' and " & _
                "sampledatetime < '" & Format(dtSampleDate, "dd/mmm/yyyy") & "' and " & _
                "sampledatetime > '" & Format(DateAdd("d", -15, dtSampleDate), "dd/mmm/yyyy") & "' " & _
                "order by sampledatetime desc"

70        Set sn = New Recordset
80        RecOpenServer 0, sn, sql
90        If sn.EOF Then
100           CheckConflict = False
110           Exit Function
120       End If

130       ThisRunNumber = sn!RunNumber
140       SampleDate = Format(sn!GlobalSampleDateTime, "dd/mmm/yyyy")

150       sql = "Select * from urine where " & _
                "runnumber = '" & ThisRunNumber & "'"
160       Set tb = New Recordset
170       RecOpenServer 0, tb, sql
180       If tb.EOF Then
190           CheckConflict = False
200           Exit Function
210       End If
220       Organism(0) = tb!cult0 & ""
230       Organism(1) = tb!cult1 & ""
240       If Trim(Organism(0) & Organism(1) = "") Then
250           CheckConflict = False
260           Exit Function
270       End If
280       If cmbOrgGroup(0) <> Organism(0) And _
             cmbOrgGroup(0) <> Organism(1) And _
             cmbOrgGroup(1) <> Organism(0) And _
             cmbOrgGroup(1) <> Organism(1) Then
290           CheckConflict = False
300           Exit Function
310       End If

320       Conflict = False

330       If cmbOrgGroup(0) = Organism(0) Then
340           grid = 0: Org = cmbOrgGroup(0): GoSub SensCheck
350       End If
360       If cmbOrgGroup(1) = Organism(0) Then
370           grid = 1: Org = cmbOrgGroup(0): GoSub SensCheck
380       End If
390       If cmbOrgGroup(0) = Organism(1) Then
400           grid = 0: Org = cmbOrgGroup(1): GoSub SensCheck
410       End If
420       If cmbOrgGroup(1) = Organism(1) Then
430           grid = 1: Org = cmbOrgGroup(1): GoSub SensCheck
440       End If

450       If Conflict Then
460           s = "Sensitivity Conflict" & vbCrLf & _
                  "Sample Number " & ThisRunNumber & _
                  " (" & Format(SampleDate, "dd/mm/yyyy") & ")" & vbCrLf & _
                  ConflictList & _
                  "Do you wish to procede?"
470           If iMsg(s, vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
480               CheckConflict = False
490           Else
500               CheckConflict = True
510           End If
520       End If

530       Exit Function

SensCheck:
540       With grdAB(grid)
550           For n = 0 To .Rows - 1
560               .Col = 1
570               .Row = n
580               If .Text <> "" Then
590                   OrgIs = .Text
600                   .Col = 0
610                   sql = "Select * from sensitivities where " & _
                            "Samplenumber = '" & ThisRunNumber & "' " & _
                            "and Antibiotic = '" & .Text & "' " & _
                            "and Organism = '" & Org & "'"
620                   Set tb = New Recordset
630                   RecOpenServer 0, tb, sql

640                   If Not tb.EOF Then
650                       OrgWas = tb!Result & ""
660                       If OrgWas <> "" And (OrgWas <> OrgIs) Then
670                           Conflict = True
680                           ConflictList = ConflictList & cmbOrgGroup(grid) & " " & .Text & " was " & _
                                             Switch(OrgWas = "S", "Sensitive", _
                                                    OrgWas = "R", "Resistant", _
                                                    OrgWas = "I", "Indeterminate") & vbCrLf
690                       End If
700                   End If
710               End If
720           Next
730       End With

740       Return

750       Exit Function

CheckConflict_Error:

          Dim strES As String
          Dim intEL As Integer

760       intEL = Erl
770       strES = Err.Description
780       LogError "frmEditMicrobiology", "CheckConflict", intEL, strES, sql

End Function


Private Sub Form_Deactivate()

10        pBar = 0
20        TimerBar.Enabled = False

End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)

10        pBar = 0

End Sub

Private Sub Form_Load()
          Dim i As Integer

10        ObservaInUse = IIf(GetOptionSetting("ObservaInUse", "0") = "0", False, True)
20        ReleaseToHealthlinkOnValidate = GetOptionSetting("ReleaseToHealthlinkOnValidate", "0")
30        cmdObserva(0).Visible = ObservaInUse
40        cmdObserva(1).Visible = ObservaInUse
50        cmdObserva(2).Visible = ObservaInUse
          

60        BacTek3DInUse = IIf(GetOptionSetting("Bactek3DInUse", "0") = "0", False, True)

70        With lblChartNumber
80            .BackColor = &H8000000F
90            .ForeColor = vbBlack
100           Select Case UCase$(HospName(0))
              Case "PORTLAOISE", "DEMONSTRATION"
110               .Caption = initial2upper(HospName(0)) & " Chart #"
120               lblAandE.Visible = False
130               txtAandE.Visible = False
140               lblNameTitle.Left = txtAandE.Left
150               txtName.Left = txtAandE.Left         '4220
160               txtName.Width = txtName.Width + txtAandE.Width  '6025
170           Case "MULLINGAR", "TULLAMORE"
180               .Caption = initial2upper(HospName(0)) & " Chart #"
190               lblAandE.Visible = True
200               txtAandE.Visible = True
                  'txtAandE.Width = 1635
210               lblNameTitle.Left = txtName.Left
220               txtName.Left = 6510
230               txtName.Width = 4395
240           End Select
250       End With

260       cmdViewReports.Visible = SysOptRTFView(0)
270       fraIQ200.Visible = SysOptShowIQ200(0)

280       For i = 1 To 4
290           grdAB(i).ColWidth(6) = 0
300           grdAB(i).ColWidth(7) = 0
310           grdAB(i).ColWidth(8) = 0
320           grdAB(i).ColWidth(9) = 0
330       Next i
340       FillLists
350       FillOrganismGroups
360       FillCurrentABs
370       FillMSandConsultantComment
380       FillMRU cMRU

390       With lblChartNumber
400           .BackColor = &H8000000F
410           .ForeColor = vbBlack
420       End With

430       dtRunDate = Format$(Now, "dd/mm/yyyy")
440       dtSampleDate = Format$(Now, "dd/mm/yyyy")

450       UpDown1.Max = 99999999

460       If pForcedSID <> 0 Then
470           txtSampleID = pForcedSID
480       Else
490           txtSampleID = GetSetting("NetAcquire", "StartUp", "LastUsedMicro", "1")
500       End If

510       GetSampleIDWithOffset
520       LoadAllDetails

530       cmdSaveDemographics.Enabled = False
540       cmdSaveInc.Enabled = False
550       cmdSaveMicro.Enabled = False
560       cmdSaveHold.Enabled = False

570       cmdValidateMicro.Enabled = True

580       LoadListBacteria
590       LoadListPregnancy
600       LoadListGenericColour ListRCC(), "RR"
610       LoadListGenericColour ListWCC(), "WW"
620       LoadListGenericColour ListFOB(), "OccultBlood"
630       LoadListGenericColour ListHPylori(), "HPylori"
640       LoadListGenericColour ListAdeno(), "Adeno"
650       LoadListGenericColour ListRota(), "Rota"
660       LoadListGenericColour ListCDiffCulture(), "CDiffCulture"
670       LoadListGenericColour ListCDiffToxinAB(), "CDiffToxinAB"
680       LoadListGenericColour ListGDH(), "CDiffGDH"
690       LoadListGenericColour ListPCR(), "CDiffPCR"
700       LoadListGenericColour ListGDHDetail(), "CDiffGDHDetail"
710       LoadListGenericColour ListPCRDetail(), "CDiffPCRDetail"
720       LoadListGenericColour ListCrypto(), "Crypto"
730       LoadListGenericColour ListGiardia(), "Giardia"
740       LoadListGenericColour ListRSV(), "RSV"

750       Activated = False
          
760       cmdValidateMicro.Visible = (SSTab1.Tab > 0)
770       cmdHealthLink.Visible = False ' GetOptionSetting("DeptMicro", "0")
780       cmdHealthLink.Visible = False 'Not ReleaseToHealthlinkOnValidate
'          SSTab1.TabVisible(1) = True
'          SSTab1.TabVisible(2) = True
'          SSTab1.TabVisible(3) = True
'          SSTab1.TabVisible(4) = True
'          SSTab1.TabVisible(5) = True
'          SSTab1.TabVisible(6) = True
'          SSTab1.TabVisible(7) = True
'          SSTab1.TabVisible(8) = True
'          SSTab1.TabVisible(9) = True
'          SSTab1.TabVisible(10) = True
'          SSTab1.TabVisible(11) = True
'          SSTab1.TabVisible(12) = True
'          SSTab1.TabVisible(13) = True

End Sub
Private Sub LoadDemographics()

          Dim sql As String
          Dim tb As Recordset
          Dim SampleDate As String
          Dim n As Integer
          Dim RooH As Boolean

10        On Error GoTo LoadDemographics_Error

20        RooH = IsRoutine()
30        cRooH(0) = RooH
40        cRooH(1) = Not RooH

50        lstABsInUse.Clear
60        cmbSite = ""
70        txtSiteDetails = ""
80        lstABsInUse.Clear
90        lblABsInUse = ""
100       cmbClinDetails = ""
110       cmdDemoVal.Caption = "&Validate"
120       EnableDemographicEntry True
130       cmbSite.Enabled = True

140       If Trim$(txtSampleID) = "" Then Exit Sub

150       GetSampleIDWithOffset

160       sql = "SELECT * FROM MicroSiteDetails WHERE " & _
                "SampleID = '" & SampleIDWithOffset & "'"
170       Set tb = New Recordset
180       RecOpenClient 0, tb, sql
190       If Not tb.EOF Then
200           cmbSite = Trim$(tb!Site & "")


210           txtSiteDetails = tb!SiteDetails & ""
220           If tb!PCA0 & "" <> "" Then lstABsInUse.AddItem tb!PCA0 & ""
230           If tb!PCA1 & "" <> "" Then lstABsInUse.AddItem tb!PCA1 & ""
240           If tb!PCA2 & "" <> "" Then lstABsInUse.AddItem tb!PCA2 & ""
250           If tb!PCA3 & "" <> "" Then lstABsInUse.AddItem tb!PCA3 & ""
260       End If
270       lblABsInUse = ""
280       For n = 0 To lstABsInUse.ListCount - 1
290           lblABsInUse = lblABsInUse & lstABsInUse.List(n) & " "
300       Next

310       sql = "SELECT * FROM Demographics WHERE " & _
                "SampleID = '" & SampleIDWithOffset & "'"

320       Set tb = New Recordset
330       RecOpenClient 0, tb, sql
340       If tb.EOF Then
350           mNewRecord = True
360           dtRunDate = Format$(Now, "dd/mm/yyyy")
370           dtSampleDate = Format$(Now, "dd/mm/yyyy")
380           dtRecDate = dtSampleDate
390           tSampleTime.Mask = ""
400           tSampleTime.Text = ""     'Format$(Now, "HH:nn")
410           tSampleTime.Mask = "##:##"
420           tRecTime.Mask = ""
430           tRecTime.Text = Format$(Now, "HH:nn")
440           tRecTime.Mask = "##:##"
450           txtChart = ""
460           txtName = ""
470           taddress(0) = ""
480           taddress(1) = ""
490           txtSex = ""
500           txtDoB = ""
510           txtAge = ""
520           cmbWard = "GP"
530           cmbClinician = ""
540           cmbGP = ""
              '  cmbHospital = HospName(0)
550           txtClinDetails = ""
560           txtDemographicComment = ""
              '  lblChartNumber.Caption = HospName(0) & " Chart #"
570           lblChartNumber.BackColor = &H8000000F
580           lblChartNumber.ForeColor = vbBlack
590           chkPregnant = 0
600           chkPenicillin = 0
610       Else
620           If Trim$(tb!Hospital & "") <> "" Then
630               lblChartNumber = Trim$(tb!Hospital) & " Chart #"
640               cmbHospital = Trim$(tb!Hospital)
650               If UCase$(tb!Hospital) = UCase$(HospName(0)) Then
660                   lblChartNumber.BackColor = &H8000000F
670                   lblChartNumber.ForeColor = vbBlack
680               Else
690                   lblChartNumber.BackColor = vbRed
700                   lblChartNumber.ForeColor = vbYellow
710               End If
720           Else
730               cmbHospital = HospName(0)
740               lblChartNumber.Caption = HospName(0) & " Chart #"
750               lblChartNumber.BackColor = &H8000000F
760               lblChartNumber.ForeColor = vbBlack
770           End If
780           If IsDate(tb!SampleDate) Then
790               dtSampleDate = Format$(tb!SampleDate, "dd/mm/yyyy")
800           Else
810               dtSampleDate = Format$(Now, "dd/mm/yyyy")
820           End If
830           If IsDate(tb!Rundate) Then
840               dtRunDate = Format$(tb!Rundate, "dd/mm/yyyy")
850           Else
860               dtRunDate = Format$(Now, "dd/mm/yyyy")
870           End If
880           If IsDate(tb!RecDate) Then
890               dtRecDate = Format$(tb!RecDate, "dd/mm/yyyy")
900           Else
910               dtRecDate = dtRunDate
920           End If
930           mNewRecord = False
940           StatusBar1.Panels(4).Text = dtRunDate
950           If Trim$(tb!RooH & "") <> "" Then cRooH(0) = tb!RooH
960           If Trim$(tb!RooH & "") <> "" Then cRooH(1) = Not tb!RooH
970           txtChart = tb!Chart & ""
980           txtAandE = Trim(tb!AandE & "")
990           txtName = Trim(tb!PatName & "")
1000          taddress(0) = tb!Addr0 & ""
1010          taddress(1) = tb!Addr1 & ""
1020          Select Case Left$(Trim$(UCase$(tb!sex & "")), 1)
              Case "M": txtSex = "Male"
1030          Case "F": txtSex = "Female"
1040          Case Else: txtSex = ""
1050          End Select
1060          txtDoB = Format$(tb!Dob, "dd/mm/yyyy")
1070          txtAge = tb!Age & ""

1080          cmbClinician = tb!Clinician & ""
1090          cmbGP = tb!GP & ""
1100          cmbWard = tb!Ward & ""

1110          txtClinDetails = tb!ClDetails & ""
1120          If IsDate(tb!SampleDate) Then
1130              dtSampleDate = Format$(tb!SampleDate, "dd/mm/yyyy")
1140              If Format$(tb!SampleDate, "hh:mm") <> "00:00" Then
1150                  tSampleTime = Format$(tb!SampleDate, "hh:mm")
1160              Else
1170                  tSampleTime.Mask = ""
1180                  tSampleTime.Text = ""
1190                  tSampleTime.Mask = "##:##"
1200              End If
1210          Else
1220              dtSampleDate = Format$(Now, "dd/mm/yyyy")
1230              tSampleTime.Mask = ""
1240              tSampleTime.Text = ""
1250              tSampleTime.Mask = "##:##"
1260          End If
1270          If IsDate(tb!RecDate & "") Then
1280              dtRecDate = Format$(tb!RecDate, "dd/mm/yyyy")
1290              If Format$(tb!RecDate, "hh:mm") <> "00:00" Then
1300                  tRecTime = Format$(tb!RecDate, "hh:mm")
1310              Else
1320                  tRecTime.Mask = ""
1330                  tRecTime.Text = ""
1340                  tRecTime.Mask = "##:##"
1350              End If
1360          Else
1370              dtRecDate = dtSampleDate
1380              tRecTime.Mask = ""
1390              tRecTime.Text = ""
1400              tRecTime.Mask = "##:##"
1410          End If
1420          If IsNull(tb!Pregnant) Then
1430              chkPregnant = 0
1440          Else
1450              chkPregnant = IIf(tb!Pregnant, 1, 0)
1460          End If

1470          If IsNull(tb!PenicillinAllergy) Then
1480              chkPenicillin.Value = 0
1490          Else
1500              chkPenicillin.Value = IIf(tb!PenicillinAllergy, 1, 0)
1510          End If

1520          If tb!Valid = True Then
1530              cmdDemoVal.Caption = "VALID"
1540              EnableDemographicEntry False
1550          End If
1560      End If
1570      cmdSaveDemographics.Enabled = False
1580      cmdSaveInc.Enabled = False

1590      Exit Sub

LoadDemographics_Error:

          Dim strES As String
          Dim intEL As Integer

1600      intEL = Erl
1610      strES = Err.Description
1620      LogError "frmEditMicrobiology", "LoadDemographics", intEL, strES, sql

End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

10        pBar = 0

End Sub

Private Sub Form_Unload(Cancel As Integer)

      'If Val(txtSampleID) > Val(GetSetting("NetAcquire", "StartUp", "LastUsedMicro", "1")) Then
10        SaveSetting "NetAcquire", "StartUp", "LastUsedMicro", txtSampleID
          'End If

20        pPrintToPrinter = ""

30        Activated = False

40        pForcedSID = 0

End Sub

Private Sub Frame1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)

10        pBar = 0

End Sub


Private Sub Frame2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)

10        pBar = 0

End Sub

Private Sub fraMicroscopy_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

10        pBar = 0

End Sub

Private Sub Frame4_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

10        pBar = 0

End Sub

Private Sub Frame5_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

10        pBar = 0

End Sub

Private Sub fraSampleID_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

10        pBar = 0

20        If cmdSaveMicro.Enabled Then
30            MoveCursorToSaveButton
40        End If

End Sub


Private Sub irelevant_Click(Index As Integer)

          Dim sql As String
          Dim tb As Recordset
          Dim strDirection As String

10        On Error GoTo irelevant_Click_Error

20        If cmdSaveDemographics.Enabled Or cmdSaveInc.Enabled Then
30            If iMsg("Save Details", vbQuestion + vbYesNo) = vbYes Then
40                GetSampleIDWithOffset
50                SaveDemographics
60                cmdSaveDemographics.Enabled = False
70                cmdSaveInc.Enabled = False
80            End If
90        End If

100       strDirection = IIf(Index = 0, "<", ">")
110       GetSampleIDWithOffset

120       sql = "SELECT TOP 1 SampleID FROM MicroSiteDetails WHERE " & _
                "SampleID " & strDirection & " " & SampleIDWithOffset & " " & _
                "AND Site like '" & cmbSiteSearch & "' " & _
                "ORDER BY SampleID " & IIf(strDirection = "<", "Desc", "Asc")

130       Set tb = New Recordset
140       RecOpenClient 0, tb, sql
150       If Not tb.EOF Then
160           txtSampleID = Val(tb!SampleID & "") - SysOptMicroOffset(0)
170       End If

180       GetSampleIDWithOffset
190       LoadAllDetails

200       cmdSaveDemographics.Enabled = False
210       cmdSaveInc.Enabled = False
220       cmdSaveMicro.Enabled = False
230       cmdSaveHold.Enabled = False

240       Exit Sub

irelevant_Click_Error:

          Dim strES As String
          Dim intEL As Integer

250       intEL = Erl
260       strES = Err.Description
270       LogError "frmEditMicrobiology", "irelevant_Click", intEL, strES, sql

End Sub

Private Sub iRunDate_Click(Index As Integer)

10        If Index = 0 Then
20            dtRunDate = DateAdd("d", -1, dtRunDate)
30        Else
40            If DateDiff("d", dtRunDate, Now) > 0 Then
50                dtRunDate = DateAdd("d", 1, dtRunDate)
60            End If
70        End If

80        SetDatesColour Me

90        cmdSaveInc.Enabled = True
100       cmdSaveDemographics.Enabled = True

End Sub

Private Sub iSampleDate_Click(Index As Integer)

10        If Index = 0 Then
20            dtSampleDate = DateAdd("d", -1, dtSampleDate)
30        Else
40            If DateDiff("d", dtSampleDate, Now) > 0 Then
50                dtSampleDate = DateAdd("d", 1, dtSampleDate)
60            End If
70        End If

80        SetDatesColour Me

90        cmdSaveInc.Enabled = True
100       cmdSaveDemographics.Enabled = True

End Sub


Private Sub iToday_Click(Index As Integer)

10        If Index = 0 Then
20            dtRunDate = Format$(Now, "dd/mm/yyyy")
30        ElseIf Index = 1 Then
40            If DateDiff("d", dtRunDate, Now) > 0 Then
50                dtSampleDate = dtRunDate
60            Else
70                dtSampleDate = Format$(Now, "dd/mm/yyyy")
80            End If
90        Else
100           dtRecDate = Format$(Now, "dd/mm/yyyy")
110       End If

120       SetDatesColour Me

130       cmdSaveInc.Enabled = True
140       cmdSaveDemographics.Enabled = True

End Sub


Private Sub lblChartNumber_Click()

10        With lblChartNumber
20            .BackColor = &H8000000F
30            .ForeColor = vbBlack
40        End With

50        If Trim$(txtChart) <> "" Then
60            LoadPatientFromChart Me, mNewRecord
70            cmdSaveDemographics.Enabled = True
80            cmdSaveInc.Enabled = True
90        End If

End Sub


Private Sub lstABsInUse_Click()

          Dim n As Integer

10        lstABsInUse.RemoveItem lstABsInUse.ListIndex

20        lblABsInUse = ""
30        For n = 0 To lstABsInUse.ListCount - 1
40            lblABsInUse = lblABsInUse & lstABsInUse.List(n) & " "
50        Next

60        cmdSaveDemographics.Enabled = True
70        cmdSaveInc.Enabled = True

End Sub



Private Sub mnuQualifier_Click()

10        With frmListsGeneric
20            .ListType = "MQ"
30            .ListTypeName = "Qualifier"
40            .ListTypeNames = "Qualifiers"
50            .Show 1
60        End With

70        LoadListQualifier

End Sub


Private Sub mnuCandSSub_Click(Index As Integer)

10        On Error GoTo mnuCandSSub_Click_Error

20        Select Case Index
          Case 0:
30            frmMicroSites.Show 1
              'FillOrganismGroups
40        Case 1:
50            With frmListsGeneric
60                .ListType = "OR"
70                .ListTypeName = "Organism Group"
80                .ListTypeNames = "Organism Groups"
90                .Show 1
100           End With
              'FillOrganismGroups
110       Case 2:
120           frmOrganisms.Show 1
130       Case 3:
140           frmNewAntibiotics.Show 1
              'FillAvailable
150       Case 4:
160           frmAntibioticLists.Show 1
170       Case 5:
180           frmMicroSetUp.Show 1
190       End Select

200       Exit Sub

mnuCandSSub_Click_Error:

          Dim strES As String
          Dim intEL As Integer

210       intEL = Erl
220       strES = Err.Description
230       LogError "frmEditMicrobiologyNew", "mnuCandSSub_Click", intEL, strES

End Sub

Private Sub mnuConsultantList_Click()

10        On Error GoTo mnuConsultantList_Click_Error

20        frmLabConsultantList.Show 1

30        Exit Sub

mnuConsultantList_Click_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmEditMicrobiologyNew", "mnuConsultantList_Click", intEL, strES

End Sub

Private Sub mnuExit_Click()

10        Unload Me

End Sub



Private Sub mnuNegativeResults_Click()

10        frmNegativeResults.Show 1

End Sub

Private Sub mnuOrganismGroups_Click()

10        With frmListsGeneric
20            .ListType = "OR"
30            .ListTypeName = "Organism Group"
40            .ListTypeNames = "Organism Groups"
50            .Show 1
60        End With

70        FillOrganismGroups

End Sub

'Private Sub mnuSites_Click()
'
'Dim SaveSite As String
'
'SaveSite = cmbSite
'
'frmMicroSites.Show 1
'
'FillSites
'
'cmbSite = SaveSite
'
'End Sub

Private Sub FillSites()

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo FillSites_Error

20        cmbSite.Clear
30        cmbSiteSearch.Clear

40        sql = "SELECT Text FROM Lists WHERE " & _
                "ListType = 'SI' " & _
                "ORDER BY ListOrder"
50        Set tb = New Recordset
60        RecOpenServer 0, tb, sql
70        Do While Not tb.EOF
80            cmbSite.AddItem tb!Text & ""
90            cmbSiteSearch.AddItem tb!Text & ""
100           tb.MoveNext
110       Loop
120       FixComboWidth cmbSite
130       FixComboWidth cmbSiteSearch
140       Exit Sub

FillSites_Error:

          Dim strES As String
          Dim intEL As Integer

150       intEL = Erl
160       strES = Err.Description
170       LogError "frmEditMicrobiologyNew", "FillSites", intEL, strES, sql

End Sub


Private Sub mnuOrganisms_Click()

10        With frmListsGeneric
20            .ListType = "IN"
30            .ListTypeName = "Organism"
40            .ListTypeNames = "Organisms"
50            .Show 1
60        End With

70        LoadListOrganism

End Sub

Private Sub LoadListOrganism()

          Dim sql As String
          Dim tb As Recordset

10        On Error GoTo LoadListOrganism_Error

20        ReDim ListOrganism(0 To 0) As String
30        ListOrganism(0) = ""

40        sql = "SELECT * FROM Lists WHERE " & _
                "ListType = 'IN' " & _
                "ORDER BY ListOrder"
50        Set tb = New Recordset
60        RecOpenServer 0, tb, sql
70        Do While Not tb.EOF
80            ReDim Preserve ListOrganism(0 To UBound(ListOrganism) + 1)
90            ListOrganism(UBound(ListOrganism)) = tb!Text & ""
100           tb.MoveNext
110       Loop

120       Exit Sub

LoadListOrganism_Error:

          Dim strES As String
          Dim intEL As Integer

130       intEL = Erl
140       strES = Err.Description
150       LogError "frmEditMicrobiologyNew", "LoadListOrganism", intEL, strES, sql


End Sub

Private Sub mnuListsFaecesSub_Click(Index As Integer)

10        On Error GoTo mnuListsFaecesSub_Click_Error

20        Select Case Index
          Case 0:    'XLD
30            With frmListsGeneric
40                .ListType = "FaecesXLD"
50                .ListTypeName = "XLD Entry"
60                .ListTypeNames = "XLD Entries"
70                .Show 1
80            End With

90            LoadListXLD
100       Case 1:    'DCA
110           With frmListsGeneric
120               .ListType = "FaecesDCA"
130               .ListTypeName = "DCA Entry"
140               .ListTypeNames = "DCA Entries"
150               .Show 1
160           End With

170           LoadListDCA
180       Case 2:    'SMAC
190           With frmListsGeneric
200               .ListType = "FaecesSMAC"
210               .ListTypeName = "SMAC Entry"
220               .ListTypeNames = "SMAC Entries"
230               .Show 1
240           End With

250           LoadListSMAC
260       Case 3:    'CROMO
270           With frmListsGeneric
280               .ListType = "FaecesCROMO"
290               .ListTypeName = "CROMO Entry"
300               .ListTypeNames = "CROMO Entries"
310               .Show 1
320           End With

330           LoadListCROMO
340       Case 4:    'cAMP
350           With frmListsGeneric
360               .ListType = "FaecesCAMP"
370               .ListTypeName = "CAMP Entry"
380               .ListTypeNames = "CAMP Entries"
390               .Show 1
400           End With

410           LoadListCAMP
420       Case 5:    'STEC
430           With frmListsGeneric
440               .ListType = "FaecesSTEC1"
450               .ListTypeName = "Day1 STEC Entry"
460               .ListTypeNames = "Day1 STEC Entries"
470               .Show 1
480           End With

490           LoadListSTEC
500       Case 6:    'STEC
510           With frmListsGeneric
520               .ListType = "FaecesSTEC2"
530               .ListTypeName = "Day2 STEC Entry"
540               .ListTypeNames = "Day2 STEC Entries"
550               .Show 1
560           End With

570           LoadListSTEC

580       End Select

590       Exit Sub

mnuListsFaecesSub_Click_Error:

          Dim strES As String
          Dim intEL As Integer

600       intEL = Erl
610       strES = Err.Description
620       LogError "frmEditMicrobiologyNew", "mnuListsFaecesSub_Click", intEL, strES

End Sub

Private Sub mnuListsFluidsSub_Click(Index As Integer)

10        On Error GoTo mnuListsFluidsSub_Click_Error

20        Select Case Index
          Case 0:    'Appearance
30            With frmListsGeneric
40                .ListType = "FA"
50                .ListTypeName = "Fluid Appearance"
60                .ListTypeNames = "Fluid Appearances"
70                .Show 1
80            End With

90            LoadListFluidAppearance
100       Case 1:    'Cell Count
110           With frmListsGeneric
120               .ListType = "CC"
130               .ListTypeName = "Cell Count"
140               .ListTypeNames = "Cell Counts"
150               .Show 1
160           End With

170           LoadListFluidCellCount
180       Case 2:    'Gram Stain
190           With frmListsGeneric
200               .ListType = "CG"
210               .ListTypeName = "Gram Stain Result"
220               .ListTypeNames = "Gram Stain Results"
230               .Show 1
240           End With

250           LoadListFluidGram
260       Case 3:    'ZN Stains
270           With frmListsGeneric
280               .ListType = "FluidZN"
290               .ListTypeName = "ZN Stain Result"
300               .ListTypeNames = "ZN Stain Results"
310               .Show 1
320           End With

330           LoadListFluidZN
340       Case 4:    'Leishman Stain
350           With frmListsGeneric
360               .ListType = "CL"
370               .ListTypeName = "Leishman's Stain Result"
380               .ListTypeNames = "Leishman's Stain Results"
390               .Show 1
400           End With

410           LoadListFluidLeishman
420       Case 5:    'Wet Prep
430           With frmListsGeneric
440               .ListType = "FW"
450               .ListTypeName = "Wet Prep"
460               .ListTypeNames = "Wet Preps"
470               .Show 1
480           End With

490           LoadListFluidWetPrep
500       Case 6:    'Crystals
510           With frmListsGeneric
520               .ListType = "FC"
530               .ListTypeName = "Crystal"
540               .ListTypeNames = "Crystals"
550               .Show 1
560           End With

570           LoadListFluidCrystals
580       Case 7:    'Sites
590           frmMicroFluidSites.Show 1

600       End Select

610       Exit Sub

mnuListsFluidsSub_Click_Error:

          Dim strES As String
          Dim intEL As Integer

620       intEL = Erl
630       strES = Err.Description
640       LogError "frmEditMicrobiologyNew", "mnuListsFluidsSub_Click", intEL, strES

End Sub

Private Sub mnuListsIdentificationSub_Click(Index As Integer)

10        On Error GoTo mnuListsIdentificationSub_Click_Error


20        With frmMicroLists
30            Select Case Index
              Case 0:
40                .optList(3).Value = True
50            Case 1:
60                .optList(4).Value = True
70            End Select
80            .Show 1
90        End With

100       Exit Sub

mnuListsIdentificationSub_Click_Error:

          Dim strES As String
          Dim intEL As Integer

110       intEL = Erl
120       strES = Err.Description
130       LogError "frmEditMicrobiologyNew", "mnuListsIdentificationSub_Click", intEL, strES

End Sub

Private Sub mnuListsMessages_Click()

10        On Error GoTo mnuListsMessages_Click_Error

20        frmConfirmMessages.Show 1

30        Exit Sub

mnuListsMessages_Click_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmEditMicrobiologyNew", "mnuListsMessages_Click", intEL, strES

End Sub



Private Sub mnuListsTabSetup_Click()

10        On Error GoTo mnuListsTabSetup_Click_Error

20        frmMicroSetUp.Show 1

30        Exit Sub

mnuListsTabSetup_Click_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmEditMicrobiologyNew", "mnuListsTabSetup_Click", intEL, strES

End Sub

Private Sub mnuListsTitlesSub_Click(Index As Integer)

10        On Error GoTo mnuListsTitlesSub_Click_Error
          'this menu is not in use
20        Select Case Index
          Case 0:    'FOB
30            With frmListsGenericColour
40                .ListType = "OccultBlood"
50                .ListTypeName = "Occult Blood Entry"
60                .ListTypeNames = "Occult Blood Entries"
70                .Show 1
80            End With

90            LoadListGenericColour ListFOB(), "OccultBlood"
100       Case 1:    'H. Pylori
110           With frmListsGenericColour
120               .ListType = "HPylori"
130               .ListTypeName = "H. Pylori Entry"
140               .ListTypeNames = "H. Pylori Entries"
150               .Show 1
160           End With

170           LoadListGenericColour ListHPylori(), "HPylori"
180       Case 2:    'C.Diff Culture
190           With frmListsGenericColour
200               .ListType = "CDiffCulture"
210               .ListTypeName = "C. Diff Culture Entry"
220               .ListTypeNames = "C. Diff Culture Entries"
230               .Show 1
240           End With

250           LoadListGenericColour ListCDiffCulture(), "CDiffCulture"
260       Case 3:    'C.Diff Toxin/AB
270           With frmListsGenericColour
280               .ListType = "CDiffToxinAB"
290               .ListTypeName = "C. Diff Toxin A/B Entry"
300               .ListTypeNames = "C. Diff Toxin A/B Entries"
310               .Show 1
320           End With

330           LoadListGenericColour ListCDiffToxinAB(), "CDiffToxinAB"
340       Case 4:    'Rota
350           With frmListsGenericColour
360               .ListType = "Rota"
370               .ListTypeName = "Rota Virus Entry"
380               .ListTypeNames = "Rota Virus Entries"
390               .Show 1
400           End With

410           LoadListGenericColour ListRota(), "Rota"
420       Case 5:    'Adeno
430           With frmListsGenericColour
440               .ListType = "Adeno"
450               .ListTypeName = "Adeno Virus Entry"
460               .ListTypeNames = "Adeno Virus Entries"
470               .Show 1
480           End With

490           LoadListGenericColour ListAdeno(), "Adeno"
500       Case 6:    'RSV
510           With frmListsGenericColour
520               .ListType = "RSV"
530               .ListTypeName = "RSV Entry"
540               .ListTypeNames = "RSV Entries"
550               .Show 1
560           End With

570           LoadListGenericColour ListRSV(), "RSV"
580       Case 7:    'Cryptosporidium
590           With frmListsGenericColour
600               .ListType = "Crypto"
610               .ListTypeName = "Cryptosporidium Entry"
620               .ListTypeNames = "Cryptosporidium Entries"
630               .Show 1
640           End With

650           LoadListGenericColour ListCrypto(), "Crypto"
660       Case 8:    'OP Comments

670       End Select

680       Exit Sub

mnuListsTitlesSub_Click_Error:

          Dim strES As String
          Dim intEL As Integer

690       intEL = Erl
700       strES = Err.Description
710       LogError "frmEditMicrobiologyNew", "mnuListsTitlesSub_Click", intEL, strES
End Sub

Private Sub mnuListsUrineSub_Click(Index As Integer)

10        On Error GoTo mnuListsUrineSub_Click_Error

20        Select Case Index
          Case 0:    'Bacteria
30            With frmListsGeneric
40                .ListType = "BB"
50                .ListTypeName = "Bacteria Entry"
60                .ListTypeNames = "Bacteria Entries"
70                .Show 1
80            End With

90            LoadListBacteria
100       Case 1:    'WCC
110           With frmListsGenericColour
120               .ListType = "WW"
130               .ListTypeName = "WCC Entry"
140               .ListTypeNames = "WCC Entries"
150               .Show 1
160           End With

170           LoadListGenericColour ListWCC(), "WW"
180       Case 2:    'RCC
190           With frmListsGenericColour
200               .ListType = "RR"
210               .ListTypeName = "RCC Entry"
220               .ListTypeNames = "RCC Entries"
230               .Show 1
240           End With

250           LoadListGenericColour ListRCC(), "RR"
260       Case 3:    'Crystals
270           With frmMicroLists
280               .optList(6).Value = True
290               .Show 1
300           End With
310       Case 4:    'Casts
320           With frmMicroLists
330               .optList(5).Value = True
340               .Show 1
350           End With
360       Case 5:    'Misc
370           With frmMicroLists
380               .optList(7).Value = True
390               .Show 1
400           End With
410       Case 7:    'Pregnancy
420           With frmListsGeneric
430               .ListType = "PG"
440               .ListTypeName = "Pregnancy Entry"
450               .ListTypeNames = "Pregnancy Entries"
460               .Show 1
470           End With

480           LoadListPregnancy



490       End Select

500       Exit Sub

mnuListsUrineSub_Click_Error:

          Dim strES As String
          Dim intEL As Integer

510       intEL = Erl
520       strES = Err.Description
530       LogError "frmEditMicrobiologyNew", "mnuListsUrineSub_Click", intEL, strES

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)

          Dim v As String
          Dim P As String

10        If PreviousTab = 5 Or _
             PreviousTab = 6 Or _
             PreviousTab = 10 Or _
             PreviousTab = 11 Or _
             PreviousTab = 8 Or _
             PreviousTab = 13 Then

              '  UpdateLockStatus Val(txtSampleID) + SysOptMicroOffset(0), True, PreviousTab
20            If cmdSaveMicro.Enabled Or cmdSaveHold.Enabled Then
30                SaveFaecalTabs
40            End If
50        End If

60        GetSampleIDWithOffset

70        If PreviousTab = 0 And (cmdSaveDemographics.Enabled Or cmdSaveInc.Enabled) Then
              '80      If Not EntriesOK(txtSampleID, txtName, txtSex, cmbWard.Text, cmbGP.Text) Then
              '90            SSTab1.Tab = 0
              '100           Exit Sub
              '110     End If
80            SaveDemographics
90        End If

100       cmdValidateMicro.Visible = False
110       Select Case SSTab1.Tab
          Case 0:    'Demographics

120       Case 1:    'Urine
130           cmdValidateMicro.Visible = True
140           If Not UrineLoaded Then
150               If LoadUrine() Then
160                   LoadPrintValid "U", v, P
170                   SSTab1.TabCaption(1) = "<<Urine>>" & v & P
180                   If txtWCC.Visible And txtWCC.Enabled Then
190                       txtWCC.SetFocus
200                   End If
210               End If
220               UrineLoaded = True
230           End If

240       Case 2:    'Identification
250           If Not IdentLoaded Then
260               LoadIdent
270               IdentLoaded = True
280           End If

290       Case 3:    'Faeces
300           If Not FaecesLoaded Then
310               LoadFaeces
320               FaecesLoaded = True
330           End If

340       Case 4:    'Sensitivities
350           cmdValidateMicro.Visible = True
360           If Not CSLoaded Then
370               LoadSensitivities
380               CSLoaded = True
390           End If

400       Case 5:
410           cmdValidateMicro.Visible = True
420           If Not FOBLoaded Then
430               LoadFOB
440               FOBLoaded = True
450           End If

460       Case 6:
470           cmdValidateMicro.Visible = True
480           If Not RotaAdenoLoaded Then
490               LoadRotaAdeno
500               RotaAdenoLoaded = True
510           End If

520       Case 7:
530           cmdValidateMicro.Visible = True
540           LoadRedSub
550       Case 8:
560           cmdValidateMicro.Visible = True
570           LoadRSV
580       Case 9:
590           cmdValidateMicro.Visible = True
600           LoadFluids
610       Case 10:
620           cmdValidateMicro.Visible = True
630           If Not CdiffLoaded Then
640               LoadCDiff
650               CdiffLoaded = True
660           End If
670       Case 11:
680           cmdValidateMicro.Visible = True
690           If Not OPLoaded Then
700               LoadOP
710               OPLoaded = True
720           End If
730       Case 13:
740           cmdValidateMicro.Visible = True
750           If Not HPyloriLoaded Then
760               LoadHPylori
770               HPyloriLoaded = True
780           End If
790       End Select

800       ShowPrintValidFlags

810       LoadLockStatus SSTab1.Tab

820       CheckValidStatus SSTab1.Tab

830       ShowWhoSaved SSTab1.Tab

840       cmdSaveMicro.Enabled = ForceSaveability
850       cmdSaveHold.Enabled = ForceSaveability

End Sub

Private Function CheckValidStatus(ByVal Index As Integer) As Boolean
      'Pass the tab number: Returns true if Valid

          Dim tb As Recordset
          Dim sql As String
          Dim Dept As String

10        On Error GoTo CheckValidStatus_Error

20        CheckValidStatus = False
30        cmdValidateMicro.Caption = "&Validate"

40        If Index = 0 Then Exit Function

50        Dept = Choose(Index, "U", "", "", "D", "F", "A", "R", "V", "C", "G", "O", "", "Y")
60        If Dept = "" Then Exit Function

70        fraValid(Index).Enabled = True
80        lblValid(Index).Visible = False
90        lblPrinted(Index).Visible = False

100       sql = "SELECT Valid, Printed FROM PrintValidLog WHERE " & _
                "SampleID = '" & Val(txtSampleID) + SysOptMicroOffset(0) & "' " & _
                "AND Department = '" & Dept & "'"
110       Set tb = New Recordset
120       RecOpenServer 0, tb, sql
130       If Not tb.EOF Then
140           If tb!Valid = 1 Then
150               fraValid(Index).Enabled = False
160               CheckValidStatus = True
170               lblValid(Index).Visible = True
180               cmdValidateMicro.Caption = "Un&Validate"
190           End If
200           If tb!Printed = 1 Then
210               lblPrinted(Index).Visible = True
220           End If
230       End If

240       Exit Function

CheckValidStatus_Error:

          Dim strES As String
          Dim intEL As Integer

250       intEL = Erl
260       strES = Err.Description
270       LogError "frmEditMicrobiologyNew", "CheckValidStatus", intEL, strES, sql

End Function
Private Sub SSTab1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

10        pBar = 0

End Sub


Private Sub tRecTime_GotFocus()

10        tRecTime.SelStart = 0
20        tRecTime.SelLength = 0

End Sub


Private Sub tRecTime_KeyPress(KeyAscii As Integer)

10        pBar = 0

20        cmdSaveDemographics.Enabled = True
30        cmdSaveInc.Enabled = True

End Sub


Private Sub taddress_Change(Index As Integer)

10        SetWardClinGP

End Sub

Private Sub taddress_KeyPress(Index As Integer, KeyAscii As Integer)

10        cmdSaveDemographics.Enabled = True
20        cmdSaveInc.Enabled = True

End Sub


Private Sub taddress_LostFocus(Index As Integer)

10        taddress(Index) = initial2upper(taddress(Index))

End Sub


Private Sub LoadDemo(ByVal IDNumber As String)

          Dim tb As New Recordset
          Dim sql As String
          Dim IDType As String
          Dim n As Long

10        On Error GoTo LoadDemo_Error

20        IDType = CheckDemographics(IDNumber)
30        If IDType = "" Then
              'clearpatient
40            Exit Sub
50        End If

          'Rem Code Change 16/01/2006
60        sql = "SELECT * from patientifs WHERE " & _
                IDType & " = '" & AddTicks(IDNumber) & "' "

70        Set tb = New Recordset
80        RecOpenServer 0, tb, sql
90        If tb.EOF = True Then
              '   clearpatient
100       Else
110           If Trim(tb!Chart & "") = "" Then txtChart = tb!Mrn & "" Else txtChart = tb!Chart & ""
120           n = InStr(tb!PatName & "", "''")
130           If n <> 0 Then
140               tb!PatName = Left$(tb!PatName, n) & Mid$(tb!PatName, n + 2)
150               tb.Update
160           End If
170           txtName = initial2upper(tb!PatName & "")
180           If Not IsNull(tb!Dob) Then
190               txtDoB = Format(tb!Dob, "DD/MM/YYYY")
200           Else
210               txtDoB = ""
220           End If
230           lblAge = CalcAge(tb!Dob & "", dtSampleDate)
240           txtAge = lblAge
250           Select Case tb!sex & ""
              Case "M": lblSex = "Male"
260           Case "F": lblSex = "Female"
270           Case Else: lblSex = ""
280           End Select
290           txtSex = lblSex
300           n = InStr(tb!Address0 & "", "''")
310           If n <> 0 Then
320               tb!Address0 = Left$(tb!Address0, n) & Mid$(tb!Address0, n + 2)
330               tb.Update
340           End If

350           taddress(0) = initial2upper(Trim(tb!Address0 & ""))
360           taddress(1) = initial2upper(Trim(tb!Address1 & ""))
370           cmbWard.Text = initial2upper(tb!Ward & "")
380           cmbClinician.Text = initial2upper(tb!Clinician & "")
390       End If
400       tb.Close

410       Exit Sub

LoadDemo_Error:

          Dim strES As String
          Dim intEL As Integer

420       intEL = Erl
430       strES = Err.Description
440       LogError "frmEditMicrobiologyNew", "LoadDemo", intEL, strES, sql

End Sub

Private Sub tRecTime_LostFocus()

10        SetDatesColour Me

End Sub

Private Sub tSampleTime_LostFocus()

10        SetDatesColour Me

End Sub

Private Sub txtAandE_LostFocus()

10        On Error GoTo txtAandE_LostFocus_Error

20        If cmdDemoVal.Caption = "VALID" Then Exit Sub

30        txtAandE = Trim$(UCase$(txtAandE))

40        If UCase(HospName(0)) = "MULLINGAR" Then
50            LoadPatientFromAandE Me, True
60        End If


70        If Trim(txtName) = "" Then
80            LoadDemo txtAandE
90        End If

100       txtAandE = UCase(txtAandE)

110       cmdSaveDemographics.Enabled = True
120       cmdSaveInc.Enabled = True

130       Exit Sub

txtAandE_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

140       intEL = Erl
150       strES = Err.Description
160       LogError "frmEditMicrobiologyNew", "txtAandE_LostFocus", intEL, strES

End Sub


Private Sub txtAdeno_KeyPress(KeyAscii As Integer)

10        KeyAscii = 0

End Sub

Private Sub txtAdeno_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

10        CycleTextBox ListAdeno(), txtAdeno

20        cmdSaveHold.Enabled = True
30        cmdSaveMicro.Enabled = True

40        ShowUnlock 6

End Sub


Private Sub txtage_Change()

10        lblAge = txtAge

End Sub

Private Sub txtAge_KeyPress(KeyAscii As Integer)

10        cmdSaveDemographics.Enabled = True
20        cmdSaveInc.Enabled = True

End Sub


Private Sub txtBacteria_Click()

          Dim n As Integer
          Dim x As Integer

10        For n = 0 To UBound(ListBacteria)
20            If txtBacteria = ListBacteria(n) Then
30                If n = UBound(ListBacteria) Then
40                    x = 0
50                Else
60                    x = n + 1
70                End If
80                txtBacteria = ListBacteria(x)
90                Exit For
100           End If
110       Next

120       ShowUnlock 1

End Sub

Private Sub txtBacteria_KeyUp(KeyCode As Integer, Shift As Integer)

10        ShowUnlock 1

End Sub

Private Sub txtBacteria_LostFocus()

          Dim sql As String
          Dim tb As Recordset

10        On Error GoTo txtBacteria_LostFocus_Error

20        sql = "SELECT Text FROM Lists WHERE " & _
                "ListType = 'BB' " & _
                "AND Code = '" & AddTicks(txtBacteria) & "'"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        If Not tb.EOF Then
60            txtBacteria = tb!Text & ""
70        End If

80        Exit Sub

txtBacteria_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

90        intEL = Erl
100       strES = Err.Description
110       LogError "frmEditMicrobiologyNew", "txtBacteria_LostFocus", intEL, strES, sql


End Sub


Private Sub txtBioResult_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

10        cmdSaveHold.Enabled = True
20        cmdSaveMicro.Enabled = True

30        ShowUnlock 9

End Sub


Private Sub txtCDiffMSC_KeyUp(KeyCode As Integer, Shift As Integer)
10        cmdSaveMicro.Enabled = True
20        cmdSaveHold.Enabled = True
End Sub

Private Sub txtchart_Change()

10        lblChart = txtChart

End Sub

Private Sub txtChart_KeyPress(KeyAscii As Integer)

10        cmdSaveDemographics.Enabled = True
20        cmdSaveInc.Enabled = True

End Sub


Private Sub txtchart_LostFocus()

10        txtChart = Trim$(UCase$(txtChart))

20        If txtChart = "" Then Exit Sub
30        If Trim$(txtName) <> "" Then Exit Sub

40        LoadPatientFromChart Me, mNewRecord

End Sub


Private Sub txtClinDetails_KeyPress(KeyAscii As Integer)

10        cmdSaveDemographics.Enabled = True
20        cmdSaveInc.Enabled = True

End Sub





Private Sub txtConC_GotFocus()

10        If txtConC.Text = "Consultant Comments" Then
20            txtConC.Text = ""
30        End If

End Sub


Private Sub txtConC_KeyUp(KeyCode As Integer, Shift As Integer)

10        cmdSaveMicro.Enabled = True
20        cmdSaveHold.Enabled = True

End Sub


Private Sub txtConC_LostFocus()

10        If Trim$(txtConC) = "" Then
20            txtConC = "Consultant Comments"
30        End If

End Sub


Private Sub txtDemographicComment_KeyUp(KeyCode As Integer, Shift As Integer)

10        cmdSaveDemographics.Enabled = True
20        cmdSaveInc.Enabled = True

End Sub

Private Sub txtDoB_Change()

10        lblDoB = txtDoB

End Sub

Private Sub txtDoB_KeyPress(KeyAscii As Integer)

10        cmdSaveDemographics.Enabled = True
20        cmdSaveInc.Enabled = True

End Sub

Private Sub txtDoB_LostFocus()

10        If txtDoB.Locked Then Exit Sub

20        txtDoB = Convert62Date(txtDoB, BACKWARD)

30        If Not IsDate(txtDoB) Then
40            txtDoB = ""
50            Exit Sub
60        End If

70        txtAge = CalcAge(txtDoB, dtSampleDate)

80        If txtAge = "" Then
90            txtDoB.BackColor = vbRed
100       Else
110           txtDoB.BackColor = vbButtonFace
120       End If

End Sub


Private Sub TimerBar_Timer()

10        pBar = pBar + 1

20        If pBar = pBar.Max Then
30            Unload Me
40            Exit Sub
50        End If

End Sub


Private Sub txtFluidComment_KeyUp(KeyCode As Integer, Shift As Integer)

10        cmdSaveHold.Enabled = True
20        cmdSaveMicro.Enabled = True

30        ShowUnlock 9

End Sub


Private Sub txtHaem_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

10        cmdSaveHold.Enabled = True
20        cmdSaveMicro.Enabled = True

30        ShowUnlock 9

End Sub


Private Sub txtHCGLevel_KeyPress(KeyAscii As Integer)

10        ShowUnlock 1

End Sub


Private Sub txtIndole_Click(Index As Integer)

10        Select Case txtIndole(Index)
          Case "": txtIndole(Index) = "Pending"
20        Case "Pending": txtIndole(Index) = "Positive"
30        Case "Positive": txtIndole(Index) = "Negative"
40        Case "Negative": txtIndole(Index) = ""
50        End Select

60        cmdSaveMicro.Enabled = True
70        cmdSaveHold.Enabled = True

End Sub


Private Sub txtIndole_KeyPress(Index As Integer, KeyAscii As Integer)

10        KeyAscii = 0

20        Select Case txtIndole(Index)
          Case "": txtIndole(Index) = "Pending"
30        Case "Pending": txtIndole(Index) = "Positive"
40        Case "Positive": txtIndole(Index) = "Negative"
50        Case "Negative": txtIndole(Index) = ""
60        End Select

70        cmdSaveMicro.Enabled = True
80        cmdSaveHold.Enabled = True

End Sub


Private Sub txtInHouseSID_KeyUp(KeyCode As Integer, Shift As Integer)

10        cmdSaveHold.Enabled = True
20        cmdSaveMicro.Enabled = True

30        ShowUnlock 9

End Sub




Private Sub txtMSC_GotFocus()

10        If txtMSC.Text = "Medical Scientist Comments" Then
20            txtMSC.Text = ""
30        End If

End Sub


Private Sub txtMSC_KeyUp(KeyCode As Integer, Shift As Integer)

10        cmdSaveMicro.Enabled = True
20        cmdSaveHold.Enabled = True

End Sub

Private Sub txtMSC_LostFocus()

10        If Trim$(txtMSC) = "" Then
20            txtMSC = "Medical Scientist Comments"
30        End If

End Sub
Private Sub txtName_Change()

10        lblName = txtName

End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)

10        On Error GoTo txtName_KeyPress_Error

20        If txtName.Locked Then Exit Sub

30        cmdSaveDemographics.Enabled = True
40        cmdSaveInc.Enabled = True

50        Exit Sub

txtName_KeyPress_Error:

          Dim strES As String
          Dim intEL As Integer



60        intEL = Erl
70        strES = Err.Description
80        LogError "frmEditMicrobiology", "txtName_KeyPress", intEL, strES


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



80        intEL = Erl
90        strES = Err.Description
100       LogError "frmEditMicrobiology", "txtname_LostFocus", intEL, strES


End Sub

Private Sub txtCatalase_Click(Index As Integer)

10        ClickMe txtCatalase(Index)

End Sub


Private Sub ClickMe(c As Control)

10        On Error GoTo ClickMe_Error

20        With c
30            Select Case Trim(UCase(.Text))
              Case "": .Text = "Pending"
40            Case "PENDING":
50                Select Case .Tag
                  Case "Coa": .Text = "Negative"
60                Case "Cat": .Text = "Negative"
70                Case "Oxi": .Text = "Negative"
80                Case "Ure": .Text = "Negative"
90                Case "Rei": .Text = "Done"
100               Case "Uri": .Text = "Done"
110               Case "Ext": .Text = "Done"
120               End Select
130           Case "DONE": .Text = ""
140           Case "NEGATIVE": .Text = "Positive"
150           Case "POSITIVE": .Text = ""
160           End Select
170       End With

180       cmdSaveMicro.Enabled = True
190       cmdSaveHold.Enabled = True

200       Exit Sub

ClickMe_Error:

          Dim strES As String
          Dim intEL As Integer

210       intEL = Erl
220       strES = Err.Description
230       LogError "frmEditMicrobiologyNew", "ClickMe", intEL, strES


End Sub

Private Sub txtCoagulase_Click(Index As Integer)

10        ClickMe txtCoagulase(Index)

End Sub


Private Sub txtNotes_KeyPress(Index As Integer, KeyAscii As Integer)

10        cmdSaveMicro.Enabled = True
20        cmdSaveHold.Enabled = True

End Sub


Private Sub txtOxidase_Click(Index As Integer)

10        ClickMe txtOxidase(Index)

End Sub


Private Sub txtPregnancy_Click()

10        On Error GoTo txtPregnancy_Click_Error


20        txtHCGLevel = ""

30        If UCase$(txtPregnancy) = "EQUIVOCAL" Then
40            If InStr(1, txtUrineComment, "Please repeat specimen in 24-48 hours.") > 0 Then
50                txtUrineComment = Replace(txtUrineComment, "Please repeat specimen in 24-48 hours.", "")
60            End If
70        ElseIf UCase$(txtPregnancy) = "INCONCLUSIVE" Then
80            If InStr(1, txtUrineComment, "Inconclusive - Please repeat.") > 0 Then
90                txtUrineComment = Replace(txtUrineComment, "Inconclusive - Please repeat.", "")
100           End If
110       ElseIf UCase$(txtPregnancy) = "SPECIMEN UNSUITABLE" Then
120           If InStr(1, txtUrineComment, "Specimen Unsuitable - Please repeat.") Then
130               txtUrineComment = Replace(txtUrineComment, "Specimen Unsuitable - Please repeat.", "")
140           End If
150       End If



160       CycleControlValue ListPregnancy, txtPregnancy

170       If UCase$(txtPregnancy) = "NEGATIVE" Then
180           txtHCGLevel = GetOptionSetting("HCGLevelLOW", "")
190       ElseIf UCase$(txtPregnancy) = "POSITIVE" Then
200           txtHCGLevel = GetOptionSetting("HCGLevelHIGH", "")
210       ElseIf UCase$(txtPregnancy) = "EQUIVOCAL" Then
220           If txtUrineComment <> "" And Right(txtUrineComment, 1) <> " " Then txtUrineComment = txtUrineComment & " "
230           txtUrineComment = txtUrineComment & "Please repeat specimen in 24-48 hours."
240       ElseIf UCase$(txtPregnancy) = "INCONCLUSIVE" Then
250           If txtUrineComment <> "" And Right(txtUrineComment, 1) <> " " Then txtUrineComment = txtUrineComment & " "
260           txtUrineComment = txtUrineComment & "Inconclusive - Please repeat."
270       ElseIf UCase$(txtPregnancy) = "SPECIMEN UNSUITABLE" Then
280           If txtUrineComment <> "" And Right(txtUrineComment, 1) <> " " Then txtUrineComment = txtUrineComment & " "
290           txtUrineComment = txtUrineComment & "Specimen Unsuitable - Please repeat."
300       End If

          'Select Case Left$(txtPregnancy & " ", 1)
          '    Case " ":
          '        txtPregnancy = "Negative"
          '        txtHCGLevel = "<25"
          '        txtUrineComment = ""
          '    Case "N":
          '        txtPregnancy = "Positive"
          '        txtHCGLevel = ">=25"
          '        txtUrineComment = ""
          '    Case "P":
          '        If HospName(0) = "SIVH" Then
          '            txtPregnancy = ""
          '        Else
          '            txtPregnancy = "Equivocal"
          '            txtHCGLevel = ""
          '            txtUrineComment = "Please repeat specimen in 24-48 hours."
          '        End If
          '    Case "E":
          '        txtPregnancy = "Inconclusive"
          '        txtHCGLevel = ""
          '        txtUrineComment = "Inconclusive - Please repeat."
          '    Case "I"
          '        txtPregnancy = "Specimen Unsuitable"
          '        txtHCGLevel = ""
          '        txtUrineComment = "Specimen Unsuitable - Please repeat."
          '    Case "S":
          '        txtPregnancy = ""
          '        txtHCGLevel = ""
          '        txtUrineComment = ""
          'End Select

310       ShowUnlock 1

320       Exit Sub

txtPregnancy_Click_Error:

          Dim strES As String
          Dim intEL As Integer

330       intEL = Erl
340       strES = Err.Description
350       LogError "frmEditMicrobiologyNew", "txtPregnancy_Click", intEL, strES


End Sub


Private Sub txtRCC_Change()

10        If Trim$(txtWCC) <> "" And Trim$(txtRCC) <> "" Then
20            lblCells.Caption = "White Cells " & txtWCC & "         " & _
                                 "Red Cells " & txtRCC
30        ElseIf Trim$(txtWCC) <> "" Then
40            lblCells.Caption = "White Cells " & txtWCC
50        ElseIf Trim$(txtRCC) <> "" Then
60            lblCells.Caption = "Red Cells " & txtRCC
70        Else
80            lblCells = ""
90        End If

End Sub

Private Sub txtRCC_Click()

10        CycleTextBox ListRCC(), txtRCC
          '
          '      Dim n As Integer
          '      Dim X As Integer
          '
          '10    X = 1
          '20    For n = 0 To UBound(ListRCC)
          '30      If Left(txtRCC & Space(10), 10) = Left(ListRCC(n) & Space(10), 10) Then
          '40        If n = UBound(ListRCC) Then
          '50          X = 0
          '60        Else
          '70          X = n + 1
          '80        End If
          '90        txtRCC = ListRCC(X)
          '100       Exit For
          '110     End If
          '120   Next

20        ShowUnlock 1

End Sub

Private Sub txtRCC_KeyUp(KeyCode As Integer, Shift As Integer)

10        ShowUnlock 1

End Sub


Private Sub txtRCC_LostFocus()

          Dim sql As String
          Dim tb As Recordset

10        On Error GoTo txtRCC_LostFocus_Error

20        sql = "SELECT Text FROM Lists WHERE " & _
                "ListType = 'RR' " & _
                "AND Code = '" & AddTicks(txtRCC) & "'"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        If Not tb.EOF Then
60            txtRCC = tb!Text & ""
70        End If

80        Exit Sub

txtRCC_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

90        intEL = Erl
100       strES = Err.Description
110       LogError "frmEditMicrobiologyNew", "txtRCC_LostFocus", intEL, strES, sql


End Sub


Private Sub txtReincubation_Click(Index As Integer)

10        ClickMe txtReincubation(Index)

End Sub


Private Sub txtRota_KeyPress(KeyAscii As Integer)

10        KeyAscii = 0

End Sub

Private Sub txtRota_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

10        CycleTextBox ListRota(), txtRota

20        cmdSaveHold.Enabled = True
30        cmdSaveMicro.Enabled = True

40        ShowUnlock 6

End Sub


Private Sub txtSampleID_GotFocus()

10        On Error GoTo txtSampleID_GotFocus_Error

20        If cmdSaveDemographics.Enabled Or cmdSaveInc.Enabled Then
30            If iMsg("Save Details", vbQuestion + vbYesNo) = vbYes Then
40                GetSampleIDWithOffset
50                SaveDemographics
60                cmdSaveDemographics.Enabled = False
70                cmdSaveInc.Enabled = False
80            End If
90        End If

100       Exit Sub

txtSampleID_GotFocus_Error:

          Dim strES As String
          Dim intEL As Integer

110       intEL = Erl
120       strES = Err.Description
130       LogError "frmEditMicrobiologyNew", "txtSampleID_GotFocus", intEL, strES

End Sub

Private Sub txtSampleID_KeyPress(KeyAscii As Integer)

10        KeyAscii = VI(KeyAscii, Numeric_Only)

End Sub


Private Sub txtSampleID_LostFocus()

10        On Error GoTo txtSampleID_LostFocus_Error

20        txtSampleID = Format$(Val(txtSampleID))
30        If txtSampleID = 0 Then Exit Sub

40        GetSampleIDWithOffset

50        LoadAllDetails

60        cmdSaveDemographics.Enabled = False
70        cmdSaveInc.Enabled = False
80        cmdSaveMicro.Enabled = ForceSaveability
90        cmdSaveHold.Enabled = ForceSaveability

100       Exit Sub

txtSampleID_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

110       intEL = Erl
120       strES = Err.Description
130       LogError "frmEditMicrobiologyNew", "txtSampleID_LostFocus", intEL, strES

End Sub

Private Sub tSampleTime_KeyPress(KeyAscii As Integer)

10        cmdSaveDemographics.Enabled = True
20        cmdSaveInc.Enabled = True

End Sub


Private Sub txtSampleID_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

10        On Error GoTo txtSampleID_MouseMove_Error

20        If cmdSaveMicro.Enabled Then
30            MoveCursorToSaveButton
40        End If

50        Exit Sub

txtSampleID_MouseMove_Error:

          Dim strES As String
          Dim intEL As Integer

60        intEL = Erl
70        strES = Err.Description
80        LogError "frmEditMicrobiologyNew", "txtSampleID_MouseMove", intEL, strES

End Sub

Private Sub txtSex_Change()

10        On Error GoTo txtSex_Change_Error

20        lblSex = txtSex

30        Exit Sub

txtSex_Change_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmEditMicrobiologyNew", "txtSex_Change", intEL, strES


End Sub

Private Sub txtsex_Click()

10        On Error GoTo txtsex_Click_Error

20        Select Case Trim$(txtSex)
          Case "": txtSex = "Male"
30        Case "Male": txtSex = "Female"
40        Case "Female": txtSex = ""
50        Case Else: txtSex = ""
60        End Select

70        cmdSaveDemographics.Enabled = True
80        cmdSaveInc.Enabled = True

90        Exit Sub

txtsex_Click_Error:

          Dim strES As String
          Dim intEL As Integer

100       intEL = Erl
110       strES = Err.Description
120       LogError "frmEditMicrobiologyNew", "txtsex_Click", intEL, strES


End Sub


Private Sub txtsex_KeyPress(KeyAscii As Integer)

10        On Error GoTo txtsex_KeyPress_Error

20        KeyAscii = 0
30        txtsex_Click

40        Exit Sub

txtsex_KeyPress_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmEditMicrobiologyNew", "txtsex_KeyPress", intEL, strES


End Sub


Private Sub txtSex_LostFocus()

10        On Error GoTo txtSex_LostFocus_Error

20        SexLostFocus txtSex, txtName

30        Exit Sub

txtSex_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmEditMicrobiologyNew", "txtSex_LostFocus", intEL, strES


End Sub


Private Sub txtUrineComment_KeyPress(KeyAscii As Integer)

10        On Error GoTo txtUrineComment_KeyPress_Error

20        ShowUnlock 1

30        Exit Sub

txtUrineComment_KeyPress_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmEditMicrobiologyNew", "txtUrineComment_KeyPress", intEL, strES


End Sub

Private Sub txtSiteDetails_Change()

10        On Error GoTo txtSiteDetails_Change_Error

20        lblSiteDetails = cmbSite & " " & txtSiteDetails

30        Exit Sub

txtSiteDetails_Change_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmEditMicrobiologyNew", "txtSiteDetails_Change", intEL, strES


End Sub

Private Sub txtSiteDetails_KeyPress(KeyAscii As Integer)

10        On Error GoTo txtSiteDetails_KeyPress_Error

20        cmdSaveDemographics.Enabled = True
30        cmdSaveInc.Enabled = True

40        Exit Sub

txtSiteDetails_KeyPress_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmEditMicrobiologyNew", "txtSiteDetails_KeyPress", intEL, strES


End Sub


Private Sub txtWCC_Change()

10        If Trim$(txtWCC) <> "" And Trim$(txtRCC) <> "" Then
20            lblCells.Caption = "White Cells " & txtWCC & "         " & _
                                 "Red Cells " & txtRCC
30        ElseIf Trim$(txtWCC) <> "" Then
40            lblCells.Caption = "White Cells " & txtWCC
50        ElseIf Trim$(txtRCC) <> "" Then
60            lblCells.Caption = "Red Cells " & txtRCC
70        Else
80            lblCells = ""
90        End If

End Sub

Private Sub txtWCC_Click()

10        CycleTextBox ListWCC(), txtWCC
          '
          '      Dim n As Integer
          '      Dim X As Integer
          '
          '10    On Error GoTo txtWCC_Click_Error
          '
          '20    For n = 0 To UBound(ListWCC)
          '30      If txtWCC = ListWCC(n) Then
          '40        If n = UBound(ListWCC) Then
          '50          X = 0
          '60        Else
          '70          X = n + 1
          '80        End If
          '90        txtWCC = ListWCC(X)
          '100       Exit For
          '110     End If
          '120   Next

20        ShowUnlock 1

30        Exit Sub

txtWCC_Click_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmEditMicrobiologyNew", "txtWCC_Click", intEL, strES

End Sub

Private Sub txtWCC_KeyPress(KeyAscii As Integer)

10        ShowUnlock 1

End Sub


Private Sub txtWCC_LostFocus()
'
'      Dim sql As String
'      Dim tb As Recordset
'
'10    On Error GoTo txtWCC_LostFocus_Error
'
'20    sql = "SELECT Text FROM Lists WHERE " & _
  '            "ListType = 'WW' " & _
  '            "AND Code = '" & AddTicks(txtWCC) & "'"
'30    Set tb = New Recordset
'40    RecOpenServer 0, tb, sql
'50    If Not tb.EOF Then
'60      txtWCC = tb!Text & ""
'70    End If
'
'80    Exit Sub
'
'txtWCC_LostFocus_Error:
'
'      Dim strES As String
'      Dim intEL As Integer
'
'90    intEL = Erl
'100   strES = Err.Description
'110   LogError "frmEditMicrobiologyNew", "txtWCC_LostFocus", intEL, strES, sql
'
'
End Sub

Private Sub txtZN_Click(Index As Integer)

10        Select Case txtZN(Index)
          Case "": txtZN(Index) = "No acid fast bacilli seen"
20        Case "No acid fast bacilli seen": txtZN(Index) = "Acid fast bacilli seen"
30        Case "Acid fast bacilli seen": txtZN(Index) = ""
40        Case Else: txtZN(Index) = ""
50        End Select

60        cmdSaveMicro.Enabled = True
70        cmdSaveHold.Enabled = True

End Sub


Private Sub txtZN_KeyPress(Index As Integer, KeyAscii As Integer)

10        KeyAscii = 0

20        Select Case txtZN(Index)
          Case "": txtZN(Index) = "No acid fast bacilli seen"
30        Case "No acid fast bacilli seen": txtZN(Index) = "Acid fast bacilli seen"
40        Case "Acid fast bacilli seen": txtZN(Index) = ""
50        Case Else: txtZN(Index) = ""
60        End Select

70        cmdSaveMicro.Enabled = True
80        cmdSaveHold.Enabled = True

End Sub


Private Sub UpDown1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

10        On Error GoTo UpDown1_MouseMove_Error

20        If cmdSaveMicro.Enabled Then
30            MoveCursorToSaveButton
40        End If

50        Exit Sub

UpDown1_MouseMove_Error:

          Dim strES As String
          Dim intEL As Integer

60        intEL = Erl
70        strES = Err.Description
80        LogError "frmEditMicrobiologyNew", "UpDown1_MouseMove", intEL, strES


End Sub

Private Sub UpDown1_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

10        On Error GoTo UpDown1_MouseUp_Error

20        pBar = 0

30        GetSampleIDWithOffset

40        LoadAllDetails

          'SetTabVisibility

50        cmdSaveDemographics.Enabled = False
60        cmdSaveInc.Enabled = False
70        cmdSaveMicro.Enabled = False
80        cmdSaveHold.Enabled = False

90        Exit Sub

UpDown1_MouseUp_Error:

          Dim strES As String
          Dim intEL As Integer

100       intEL = Erl
110       strES = Err.Description
120       LogError "frmEditMicrobiologyNew", "UpDown1_MouseUp", intEL, strES

End Sub



Public Property Let PrintToPrinter(ByVal strNewValue As String)

10        On Error GoTo PrintToPrinter_Error

20        pPrintToPrinter = strNewValue

30        Exit Property

PrintToPrinter_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmEditMicrobiologyNew", "PrintToPrinter", intEL, strES


End Property
Public Property Get PrintToPrinter() As String

10        On Error GoTo PrintToPrinter_Error

20        PrintToPrinter = pPrintToPrinter

30        Exit Property

PrintToPrinter_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmEditMicrobiologyNew", "PrintToPrinter", intEL, strES


End Property

Private Sub udHistoricalFaecesView_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

10        On Error GoTo udHistoricalFaecesView_MouseUp_Error

20        FillHistoricalFaeces

30        Exit Sub

udHistoricalFaecesView_MouseUp_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmEditMicrobiologyNew", "udHistoricalFaecesView_MouseUp", intEL, strES


End Sub


Public Property Let ForcedSID(ByVal NewValue As Double)

10        pForcedSID = NewValue

End Property
Private Sub AdjustOrganism()

          Dim sql As String
          Dim tb As Recordset


10        On Error GoTo AdjustOrganism_Error

          'QMS Ref #818120

20        sql = "SELECT I.SampleID, I.IsolateNumber FROM Isolates I " & _
                "Inner Join Sensitivities S " & _
                "On I.SampleID = S.SampleID " & _
                "And I.IsolateNumber = S.IsolateNumber " & _
                "WHERE I.SampleID = '" & SampleIDWithOffset & "' " & _
                "AND I.OrganismName = 'Staphylococcus aureus' " & _
                "AND S.AntibioticCode = 'OXA' And S.RSI = 'R'"

30        Set tb = New Recordset
40        RecOpenClient 0, tb, sql
50        If Not tb.EOF Then
60            While Not tb.EOF
70                sql = "Update Isolates Set " & _
                        "OrganismName = 'Staphylococcus aureus (MRSA)' " & _
                        "Where SampleID = '" & tb!SampleID & "' " & _
                        "AND IsolateNumber = " & tb!IsolateNumber
80                Cnxn(0).Execute sql
90                cmbOrgName(tb!IsolateNumber) = "Staphylococcus aureus (MRSA)"
100               tb.MoveNext
110           Wend
120       End If



130       Exit Sub

AdjustOrganism_Error:

          Dim strES As String
          Dim intEL As Integer

140       intEL = Erl
150       strES = Err.Description
160       LogError "frmEditMicrobiologyNew", "AdjustOrganism", intEL, strES, sql

End Sub


Private Function AutoReportAB(GridIndex As Integer) As Boolean

10        On Error GoTo AutoReportAB_Error

20        AutoReportAB = False
30        With grdAB(GridIndex)
40            .Col = 2
50            If .CellPicture = 0 And _
                 .TextMatrix(.Row, 6) = "1" And _
                 .TextMatrix(.Row, 1) = .TextMatrix(.Row, 7) Then
60                AutoReportAB = True
70            End If
80        End With

90        Exit Function

AutoReportAB_Error:

          Dim strES As String
          Dim intEL As Integer

100       intEL = Erl
110       strES = Err.Description
120       LogError "frmEditMicrobiologyNew", "AutoReportAB", intEL, strES

End Function




Private Sub ApplyExclusionABRule()

          Dim Y As Integer
          Dim i As Integer
          Dim J As Integer
          Dim SelectedABCount As Integer
          Dim ABPerPage As Integer
          Dim ExclusionABCount As Integer
          Dim InclusionABCount As Integer

10        On Error GoTo ApplyExclusionABRule_Error

20        ABPerPage = 23


30        For Y = 1 To 4
40            ExclusionABCount = 0
50            SelectedABCount = 0
60            With grdAB(Y)
70                If .Rows > 2 Then
80                    For i = 1 To .Rows - 1
90                        .Row = i
100                       .Col = 2
110                       If .CellPicture = imgSquareTick Then
120                           SelectedABCount = SelectedABCount + 1
130                       End If
140                   Next i
150                   If ABPerPage > SelectedABCount Then
160                       InclusionABCount = ABPerPage - SelectedABCount
170                       .Col = 8
180                       .Sort = flexSortNumericAscending
190                       J = 1
200                       While J <= InclusionABCount And J < .Rows
210                           .Col = 2
220                           .Row = J
230                           If (.CellPicture = 0 Or .CellPicture = imgSquareCross.Picture) And _
                                 .TextMatrix(.Row, 6) = "1" And _
                                 .TextMatrix(.Row, 1) = .TextMatrix(.Row, 7) Then

240                               Set .CellPicture = imgSquareTick.Picture
250                           Else
260                               InclusionABCount = InclusionABCount + 1
270                           End If
280                           J = J + 1
290                       Wend
300                       .Col = 9
310                       .Sort = flexSortNumericAscending

320                   ElseIf SelectedABCount > ABPerPage Then
330                       iMsg "You have selected more than 23 antibiotics to report. " & _
                               "Extra low priority antibiotics will automatically be deselected"
340                       ExclusionABCount = SelectedABCount - ABPerPage
350                       .Col = 8
360                       .Sort = flexSortNumericDescending
370                       J = 1
380                       While J <= ExclusionABCount
390                           .Row = J
400                           .Col = 2
410                           If .CellPicture = imgSquareTick Then
420                               Set .CellPicture = imgSquareCross.Picture
430                           Else
440                               ExclusionABCount = ExclusionABCount + 1
450                           End If
460                           J = J + 1
470                       Wend
480                       .Col = 9
490                       .Sort = flexSortNumericAscending
500                   End If
510               End If
520           End With
530       Next Y
540       Exit Sub

ApplyExclusionABRule_Error:

          Dim strES As String
          Dim intEL As Integer

550       intEL = Erl
560       strES = Err.Description
570       LogError "frmEditMicrobiologyNew", "ApplyExclusionABRule", intEL, strES

End Sub


Private Sub ClearIQ200()
10        On Error GoTo ClearIQ200_Error

20        With grdIQ200
30            .Rows = 2
40            .AddItem ""
50            .RemoveItem 1
              
60            .ColWidth(0) = 0
70        End With
          

80        Exit Sub

ClearIQ200_Error:

          Dim strES As String
          Dim intEL As Integer

90        intEL = Erl
100       strES = Err.Description
110       LogError "frmEditMicrobiologyNew", "ClearIQ200", intEL, strES

End Sub

'---------------------------------------------------------------------------------------
' Procedure : CheckReportReleasetoWard
' Author    : XPMUser
' Date      : 3/13/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Sub CheckReportReleasetoWard(SampleID As String)

    On Error GoTo CheckReportReleasetoWard_Error



    Dim sql As String
    Dim tb As ADODB.Recordset
    Dim tbPVL As ADODB.Recordset
    Set tb = New Recordset
    Set tbPVL = New Recordset
    cmdReleaseReport.BackColor = &H8000000F
    cmdReleaseReport.Caption = "Release to Consultant"
    cmdReleasetoWard.BackColor = &H8000000F
    'cmdReleasetoWard.Caption = "Release To Ward"
    sql = "select Top 1 * from ConsultantList WHERE sampleid = " & SampleIDWithOffset
    RecOpenServer 0, tb, sql
    If Not tb.EOF Then
        Select Case tb!Status
            Case 0:
                'Released to consultant
                cmdReleaseReport.BackColor = vbCyan
                cmdReleaseReport.Caption = "Release To Consultant"
            Case 1:
                'Releasesd to Ward
                'cmdReleaseReport.BackColor = vbCyan
                cmdReleasetoWard.BackColor = vbGreen
            Case 2:
                'In lab for review
            Case 3:
                cmdReleasetoWard.BackColor = vbGreen
        End Select
        
        'Exit Sub
    End If
'    tb.Close
'    sql = "select Top 1 * from Reports where sampleid ='" & SampleIDWithOffset & "' AND Dept = 'N'"
'    RecOpenClient 0, tb, sql
'    If Not tb.EOF Then
'        sql = "SELECT TOP 1 * FROM PrintValidLog pvl WHERE sampleid =  '" & SampleIDWithOffset & "' ORDER BY validateddatetime desc"
'        RecOpenClient 0, tbPVL, sql
'        If Not tbPVL.EOF Then
'            If tbPVL!ValidatedDateTime > tb!printTime Then
'                cmdReleasetoWard.BackColor = vbYellow
'            Else
'                cmdReleasetoWard.BackColor = vbGreen
'            End If
'        Else
'            cmdReleasetoWard.BackColor = vbGreen
'        End If
'    End If


    Exit Sub


CheckReportReleasetoWard_Error:

    Dim strES As String
    Dim intEL As Integer

    intEL = Erl
    strES = Err.Description
    LogError "frmEditMicrobiologyNew", "CheckReportReleasetoWard", intEL, strES, sql
End Sub

Private Sub CheckPatientNotePad(SampleID As String)

          Dim tb As New Recordset
          Dim sql As String


10        On Error GoTo CheckPatientNotePad_Error

20        sql = "SELECT * from PatientNotePad WHERE " & _
                "SampleID = '" & Val(txtSampleID) + SysOptMicroOffset(0) & "'"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        If Not tb.EOF Then
60            cmdPatientNotePad.BackColor = &HFF00&
70        Else
80            cmdPatientNotePad.BackColor = &H8000000F
90        End If

100       On Error GoTo 0
110       Exit Sub

CheckPatientNotePad_Error:

          Dim strES As String
          Dim intEL As Integer

120       intEL = Erl
130       strES = Err.Description
140       LogError "frmEditMicrobiologyNew", "CheckPatientNotePad", intEL, strES, sql

End Sub

