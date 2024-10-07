VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frmEditAllNew 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - General Chemistry"
   ClientHeight    =   10845
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   14820
   Icon            =   "frmEditAllNew.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10845
   ScaleWidth      =   14820
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.CommandButton cmdPhone 
      Caption         =   "Phone Results"
      Height          =   750
      Left            =   13500
      Picture         =   "frmEditAllNew.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   243
      Top             =   7020
      Width           =   1275
   End
   Begin VB.Frame Frame6 
      Height          =   1800
      Left            =   135
      TabIndex        =   178
      Top             =   270
      Width           =   2475
      Begin VB.ComboBox cMRU 
         Height          =   315
         Left            =   180
         TabIndex        =   179
         Text            =   "cMRU"
         Top             =   1260
         Width           =   1980
      End
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
         Left            =   120
         MaxLength       =   12
         TabIndex        =   0
         Tag             =   "Sample Id "
         Top             =   510
         Width           =   1785
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   480
         Left            =   1936
         TabIndex        =   180
         Top             =   510
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   847
         _Version        =   327681
         Value           =   1
         BuddyControl    =   "txtSampleID"
         BuddyDispid     =   196612
         OrigLeft        =   1920
         OrigTop         =   540
         OrigRight       =   2160
         OrigBottom      =   1020
         Max             =   99999999
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label lblResultOrRequest 
         Alignment       =   2  'Center
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Results"
         Height          =   285
         Left            =   600
         TabIndex        =   183
         Top             =   210
         Width           =   885
      End
      Begin VB.Image imgLast 
         Height          =   300
         Left            =   2070
         Picture         =   "frmEditAllNew.frx":074C
         Stretch         =   -1  'True
         ToolTipText     =   "Find Last Record"
         Top             =   180
         Width           =   300
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "MRU"
         Height          =   195
         Left            =   855
         TabIndex        =   182
         Top             =   1035
         Width           =   375
      End
      Begin VB.Image iRelevant 
         Height          =   480
         Index           =   1
         Left            =   1485
         Picture         =   "frmEditAllNew.frx":0B8E
         Top             =   135
         Width           =   480
      End
      Begin VB.Image iRelevant 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   0
         Left            =   120
         Picture         =   "frmEditAllNew.frx":0E98
         Top             =   150
         Width           =   480
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         Caption         =   "Sample ID"
         Height          =   195
         Left            =   720
         TabIndex        =   181
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Patient Details"
      Height          =   1800
      Left            =   2565
      TabIndex        =   153
      Top             =   270
      Width           =   12150
      Begin VB.CommandButton bDoB 
         Appearance      =   0  'Flat
         Caption         =   "S&earch"
         Height          =   285
         Left            =   9795
         TabIndex        =   157
         Top             =   285
         Width           =   705
      End
      Begin VB.CommandButton bsearch 
         Appearance      =   0  'Flat
         Caption         =   "Se&arch"
         Height          =   345
         Left            =   6885
         TabIndex        =   156
         Top             =   180
         Width           =   675
      End
      Begin VB.TextBox txtSex 
         Height          =   285
         Left            =   8175
         MaxLength       =   6
         TabIndex        =   155
         Tag             =   "Sex"
         Top             =   1035
         Width           =   1545
      End
      Begin VB.TextBox txtAge 
         Height          =   285
         Left            =   8175
         MaxLength       =   4
         TabIndex        =   154
         Tag             =   "Age"
         Top             =   675
         Width           =   1545
      End
      Begin VB.TextBox txtDoB 
         Height          =   285
         Left            =   8175
         MaxLength       =   10
         TabIndex        =   5
         Tag             =   "Date of Birth"
         Top             =   315
         Width           =   1545
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
         Left            =   4050
         MaxLength       =   30
         TabIndex        =   4
         Tag             =   "Patient Name"
         ToolTipText     =   "Patients Name"
         Top             =   585
         Width           =   3495
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
         Left            =   90
         MaxLength       =   8
         TabIndex        =   1
         Tag             =   "Chart Number"
         ToolTipText     =   "Chart/Mrn Number"
         Top             =   570
         Width           =   1425
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
         Height          =   480
         Left            =   2790
         TabIndex        =   3
         Tag             =   "Pas Number(nopas)"
         ToolTipText     =   "Pas Number"
         Top             =   570
         Width           =   1245
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
         Height          =   480
         Left            =   1530
         TabIndex        =   2
         Tag             =   "A and E Number"
         ToolTipText     =   "A && E Number"
         Top             =   570
         Width           =   1245
      End
      Begin VB.Label lblUrgent 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "URGENT"
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
         Left            =   9945
         TabIndex        =   187
         Top             =   900
         Width           =   2085
      End
      Begin VB.Label lblDemographicComment 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   90
         TabIndex        =   184
         ToolTipText     =   "Demographic Comment"
         Top             =   1350
         Width           =   11985
      End
      Begin VB.Label lblSampledate 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   10530
         TabIndex        =   170
         Top             =   450
         Width           =   1575
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         Caption         =   "Sex"
         Height          =   195
         Left            =   7800
         TabIndex        =   167
         Top             =   1065
         Width           =   270
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "Age"
         Height          =   195
         Left            =   7770
         TabIndex        =   166
         Top             =   705
         Width           =   285
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         Caption         =   "D.o.B"
         Height          =   195
         Left            =   7680
         TabIndex        =   165
         Top             =   345
         Width           =   405
      End
      Begin VB.Label lblNameTitle 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         Height          =   195
         Left            =   4050
         TabIndex        =   164
         Top             =   360
         Width           =   420
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
         Left            =   5040
         TabIndex        =   163
         Top             =   270
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.Label lAddWardGP 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   90
         TabIndex        =   162
         Top             =   1095
         Width           =   7455
      End
      Begin VB.Label lblChartNumber 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Monaghan Chart #"
         Height          =   285
         Left            =   90
         TabIndex        =   161
         ToolTipText     =   "Click to change Location"
         Top             =   315
         Width           =   1425
      End
      Begin VB.Label lblNOPAS 
         AutoSize        =   -1  'True
         Caption         =   "NOPAS"
         Height          =   195
         Index           =   0
         Left            =   2970
         TabIndex        =   160
         Top             =   360
         Width           =   555
      End
      Begin VB.Label lblAandE 
         Caption         =   "A and E"
         Height          =   225
         Left            =   1755
         TabIndex        =   159
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblRundate 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4185
         TabIndex        =   158
         Top             =   630
         Width           =   1515
      End
      Begin VB.Label Label6 
         Caption         =   "Sample Date"
         Height          =   255
         Left            =   10575
         TabIndex        =   171
         Top             =   270
         Width           =   1035
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   144
      Top             =   10560
      Width           =   14820
      _ExtentX        =   26141
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "User Name"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "08/03/2006"
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
   Begin VB.CommandButton bOrderTests 
      Caption         =   "Order Tests"
      Height          =   780
      Left            =   13500
      Picture         =   "frmEditAllNew.frx":11A2
      Style           =   1  'Graphical
      TabIndex        =   142
      Tag             =   "bOrder"
      Top             =   3015
      Width           =   1290
   End
   Begin VB.CommandButton cmdSetPrinter 
      Height          =   705
      Left            =   13500
      Picture         =   "frmEditAllNew.frx":14AC
      Style           =   1  'Graphical
      TabIndex        =   108
      ToolTipText     =   "Back Colour Normal:Automatic Printer Selection.Back Colour Red:-Forced"
      Top             =   3870
      Width           =   1275
   End
   Begin VB.Timer TimerBar 
      Interval        =   1000
      Left            =   13770
      Top             =   540
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   135
      Left            =   180
      TabIndex        =   97
      Top             =   45
      Width           =   14520
      _ExtentX        =   25612
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   1
      Max             =   600
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdPrintHold 
      Caption         =   "Print && Hold"
      Height          =   705
      Left            =   13500
      Picture         =   "frmEditAllNew.frx":17B6
      Style           =   1  'Graphical
      TabIndex        =   89
      Top             =   4635
      Width           =   1275
   End
   Begin VB.CommandButton bViewBB 
      Caption         =   "Transfusion Details"
      Enabled         =   0   'False
      Height          =   840
      Left            =   13485
      Picture         =   "frmEditAllNew.frx":1AC0
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   2115
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.CommandButton bHistory 
      Caption         =   "&History"
      Height          =   705
      Left            =   13500
      Picture         =   "frmEditAllNew.frx":238A
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   7830
      Width           =   1275
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   705
      Left            =   13500
      Picture         =   "frmEditAllNew.frx":27CC
      Style           =   1  'Graphical
      TabIndex        =   25
      Tag             =   "bprint"
      Top             =   5400
      Width           =   1275
   End
   Begin VB.CommandButton bFAX 
      Caption         =   "&Fax"
      Height          =   795
      Left            =   13500
      Picture         =   "frmEditAllNew.frx":2AD6
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   6165
      Width           =   1275
   End
   Begin VB.CommandButton bCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   645
      Left            =   13500
      Picture         =   "frmEditAllNew.frx":2DE0
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   8595
      Width           =   1275
   End
   Begin TabDlg.SSTab sstabAll 
      Height          =   8280
      Left            =   180
      TabIndex        =   23
      Top             =   2160
      Width           =   13275
      _ExtentX        =   23416
      _ExtentY        =   14605
      _Version        =   393216
      Tabs            =   8
      TabsPerRow      =   8
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Demographics"
      TabPicture(0)   =   "frmEditAllNew.frx":30EA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame7"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdSaveDemographics"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdSaveInc"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame10(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "ssPanPgP"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdDemoVal"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Haematology"
      TabPicture(1)   =   "frmEditAllNew.frx":3106
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Command1"
      Tab(1).Control(1)=   "List1"
      Tab(1).Control(2)=   "List2"
      Tab(1).Control(3)=   "Combo3"
      Tab(1).Control(4)=   "Combo2"
      Tab(1).Control(5)=   "Text1"
      Tab(1).Control(6)=   "Combo1"
      Tab(1).Control(7)=   "cmdViewHaemRep"
      Tab(1).Control(8)=   "cmdHSaveH"
      Tab(1).Control(9)=   "txtCondition"
      Tab(1).Control(10)=   "bFilm"
      Tab(1).Control(11)=   "cFilm"
      Tab(1).Control(12)=   "bHaemGraphs"
      Tab(1).Control(13)=   "txtHaemComment"
      Tab(1).Control(14)=   "bValidateHaem"
      Tab(1).Control(15)=   "cmdSaveHaem"
      Tab(1).Control(16)=   "bViewHaemRepeat"
      Tab(1).Control(17)=   "Panel3D8"
      Tab(1).Control(18)=   "MSFlexGrid1"
      Tab(1).Control(19)=   "Image1"
      Tab(1).Control(20)=   "Label1(11)"
      Tab(1).Control(21)=   "Label7"
      Tab(1).Control(22)=   "lblAnalyser"
      Tab(1).Control(23)=   "Rundate(1)"
      Tab(1).Control(24)=   "lHDate"
      Tab(1).Control(25)=   "lblHaemValid"
      Tab(1).Control(26)=   "lHaemErrors"
      Tab(1).Control(27)=   "lblHaemPrinted"
      Tab(1).Control(28)=   "Label1(10)"
      Tab(1).ControlCount=   29
      TabCaption(2)   =   "Biochemistry"
      TabPicture(2)   =   "frmEditAllNew.frx":3122
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdViewBioReps"
      Tab(2).Control(1)=   "bReprint"
      Tab(2).Control(2)=   "bViewBioRepeat"
      Tab(2).Control(3)=   "Frame8"
      Tab(2).Control(4)=   "grdOutstanding"
      Tab(2).Control(5)=   "bValidateBio"
      Tab(2).Control(6)=   "cmdSaveBio"
      Tab(2).Control(7)=   "Frame3"
      Tab(2).Control(8)=   "bAddBio"
      Tab(2).Control(9)=   "bremoveduplicates"
      Tab(2).Control(10)=   "cAdd"
      Tab(2).Control(11)=   "tnewvalue"
      Tab(2).Control(12)=   "cUnits"
      Tab(2).Control(13)=   "cSampleType"
      Tab(2).Control(14)=   "Frame2"
      Tab(2).Control(15)=   "gBio"
      Tab(2).Control(16)=   "An2"
      Tab(2).Control(17)=   "An1"
      Tab(2).Control(18)=   "lblAss"
      Tab(2).Control(19)=   "Rundate(2)"
      Tab(2).Control(20)=   "lBDate"
      Tab(2).Control(21)=   "lRandom"
      Tab(2).Control(22)=   "lblViewSplit"
      Tab(2).ControlCount=   23
      TabCaption(3)   =   "Coagulation"
      TabPicture(3)   =   "frmEditAllNew.frx":313E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txtCoagComment"
      Tab(3).Control(1)=   "cmdViewCoagRep"
      Tab(3).Control(2)=   "cCunits"
      Tab(3).Control(3)=   "cmdPrintAll"
      Tab(3).Control(4)=   "Frame9"
      Tab(3).Control(5)=   "cmdValidateCoag"
      Tab(3).Control(6)=   "cmdSaveCoag"
      Tab(3).Control(7)=   "bViewCoagRepeat"
      Tab(3).Control(8)=   "bAddCoag"
      Tab(3).Control(9)=   "tResult"
      Tab(3).Control(10)=   "cParameter"
      Tab(3).Control(11)=   "grdCoag"
      Tab(3).Control(12)=   "grdOutstandingCoag"
      Tab(3).Control(13)=   "grdPrev"
      Tab(3).Control(14)=   "Rundate(3)"
      Tab(3).Control(15)=   "lCDate"
      Tab(3).Control(16)=   "lblPrevCoag"
      Tab(3).Control(17)=   "Label20"
      Tab(3).ControlCount=   18
      TabCaption(4)   =   "Endocrinology"
      TabPicture(4)   =   "frmEditAllNew.frx":315A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "cmdViewReports"
      Tab(4).Control(1)=   "Frame12(0)"
      Tab(4).Control(2)=   "cISampleType(0)"
      Tab(4).Control(3)=   "cIUnits(0)"
      Tab(4).Control(4)=   "tINewValue(0)"
      Tab(4).Control(5)=   "cIAdd(0)"
      Tab(4).Control(6)=   "cmdIremoveduplicates(0)"
      Tab(4).Control(7)=   "cmdIAdd(0)"
      Tab(4).Control(8)=   "cmdSaveImm(0)"
      Tab(4).Control(9)=   "bValidateImm(0)"
      Tab(4).Control(10)=   "Frame81(0)"
      Tab(4).Control(11)=   "bViewImmRepeat(0)"
      Tab(4).Control(12)=   "bImmRePrint(0)"
      Tab(4).Control(13)=   "grdOutstandings(0)"
      Tab(4).Control(14)=   "gImm(0)"
      Tab(4).Control(15)=   "Frame10(1)"
      Tab(4).Control(16)=   "Frame11(0)"
      Tab(4).Control(17)=   "lblEDate"
      Tab(4).Control(18)=   "Rundate(0)"
      Tab(4).Control(19)=   "lImmRan(0)"
      Tab(4).Control(20)=   "lblImmViewSplit(0)"
      Tab(4).ControlCount=   21
      TabCaption(5)   =   "Blood Gas"
      TabPicture(5)   =   "frmEditAllNew.frx":3176
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "cISampleType(2)"
      Tab(5).Control(1)=   "cmdIAdd(2)"
      Tab(5).Control(2)=   "cIAdd(2)"
      Tab(5).Control(3)=   "tINewValue(2)"
      Tab(5).Control(4)=   "cIUnits(2)"
      Tab(5).Control(5)=   "Frame15"
      Tab(5).Control(6)=   "bRePrintBga"
      Tab(5).Control(7)=   "bViewBgaRepeat"
      Tab(5).Control(8)=   "cmdValBG"
      Tab(5).Control(9)=   "cmdSaveBGa"
      Tab(5).Control(10)=   "Frame14"
      Tab(5).Control(11)=   "gBga"
      Tab(5).Control(12)=   "lblBgaDate"
      Tab(5).Control(13)=   "Rundate(5)"
      Tab(5).ControlCount=   14
      TabCaption(6)   =   "Immunology"
      TabPicture(6)   =   "frmEditAllNew.frx":3192
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "cmdViewImmRep"
      Tab(6).Control(1)=   "Frame12(1)"
      Tab(6).Control(2)=   "Frame11(1)"
      Tab(6).Control(3)=   "Frame81(1)"
      Tab(6).Control(4)=   "cmdIremoveduplicates(1)"
      Tab(6).Control(5)=   "cmdIAdd(1)"
      Tab(6).Control(6)=   "cmdSaveImm(1)"
      Tab(6).Control(7)=   "bValidateImm(1)"
      Tab(6).Control(8)=   "bViewImmRepeat(1)"
      Tab(6).Control(9)=   "bImmRePrint(1)"
      Tab(6).Control(10)=   "cISampleType(1)"
      Tab(6).Control(11)=   "cIUnits(1)"
      Tab(6).Control(12)=   "tINewValue(1)"
      Tab(6).Control(13)=   "cIAdd(1)"
      Tab(6).Control(14)=   "cmdGetBio"
      Tab(6).Control(15)=   "grdOutstandings(1)"
      Tab(6).Control(16)=   "gImm(1)"
      Tab(6).Control(17)=   "lImmRan(1)"
      Tab(6).Control(18)=   "lblImmViewSplit(1)"
      Tab(6).Control(19)=   "lblIRundate"
      Tab(6).Control(20)=   "Rundate(4)"
      Tab(6).ControlCount=   21
      TabCaption(7)   =   "Externals"
      TabPicture(7)   =   "frmEditAllNew.frx":31AE
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "cmdViewExtReport"
      Tab(7).Control(1)=   "baddtotests"
      Tab(7).Control(2)=   "txtEtc(8)"
      Tab(7).Control(3)=   "txtEtc(7)"
      Tab(7).Control(4)=   "txtEtc(6)"
      Tab(7).Control(5)=   "txtEtc(5)"
      Tab(7).Control(6)=   "txtEtc(1)"
      Tab(7).Control(7)=   "txtEtc(2)"
      Tab(7).Control(8)=   "txtEtc(3)"
      Tab(7).Control(9)=   "txtEtc(4)"
      Tab(7).Control(10)=   "txtEtc(0)"
      Tab(7).Control(11)=   "cmdSaveExt"
      Tab(7).Control(12)=   "cmdDel"
      Tab(7).Control(13)=   "grdExt"
      Tab(7).ControlCount=   14
      Begin VB.CommandButton Command1 
         Caption         =   "Film Report"
         Height          =   825
         Left            =   -68040
         Picture         =   "frmEditAllNew.frx":31CA
         Style           =   1  'Graphical
         TabIndex        =   263
         Top             =   7380
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ListBox List1 
         Enabled         =   0   'False
         Height          =   1230
         Left            =   -64800
         TabIndex        =   262
         Top             =   2700
         Width           =   2895
      End
      Begin VB.ListBox List2 
         Enabled         =   0   'False
         Height          =   1035
         Left            =   -64800
         TabIndex        =   261
         Top             =   3960
         Width           =   2865
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   -70320
         TabIndex        =   260
         Text            =   "cSampleType"
         Top             =   7800
         Width           =   1515
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   -71640
         TabIndex        =   259
         Text            =   "cUnits"
         Top             =   7800
         Width           =   1305
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   -73200
         MaxLength       =   15
         TabIndex        =   258
         Top             =   7800
         Width           =   1575
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   -74880
         Sorted          =   -1  'True
         TabIndex        =   257
         Text            =   "cAdd"
         Top             =   7800
         Width           =   1575
      End
      Begin VB.CommandButton cmdViewExtReport 
         Caption         =   "Reports"
         Height          =   915
         Left            =   -65055
         Picture         =   "frmEditAllNew.frx":34D4
         Style           =   1  'Graphical
         TabIndex        =   254
         Top             =   5895
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtCoagComment 
         BackColor       =   &H80000018&
         Height          =   1545
         Left            =   -66135
         MaxLength       =   320
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   91
         ToolTipText     =   "Only 360 Characters"
         Top             =   3780
         Width           =   2865
      End
      Begin VB.CommandButton cmdViewImmRep 
         Caption         =   "Reports"
         Height          =   915
         Left            =   -67305
         Picture         =   "frmEditAllNew.frx":37DE
         Style           =   1  'Graphical
         TabIndex        =   251
         Top             =   5880
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.CommandButton cmdViewHaemRep 
         Caption         =   "Reports"
         Height          =   825
         Left            =   -67245
         Picture         =   "frmEditAllNew.frx":3AE8
         Style           =   1  'Graphical
         TabIndex        =   250
         Top             =   7365
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdViewCoagRep 
         Caption         =   "Reports"
         Height          =   915
         Left            =   -68610
         Picture         =   "frmEditAllNew.frx":3DF2
         Style           =   1  'Graphical
         TabIndex        =   249
         Top             =   5895
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdViewBioReps 
         Caption         =   "Reports"
         Height          =   915
         Left            =   -68880
         Picture         =   "frmEditAllNew.frx":40FC
         Style           =   1  'Graphical
         TabIndex        =   248
         Top             =   5895
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdViewReports 
         Caption         =   "Reports"
         Height          =   915
         Left            =   -68070
         Picture         =   "frmEditAllNew.frx":4406
         Style           =   1  'Graphical
         TabIndex        =   247
         Top             =   5895
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.CommandButton baddtotests 
         Appearance      =   0  'Flat
         Caption         =   "Order External Test"
         Height          =   1095
         Left            =   -64440
         Picture         =   "frmEditAllNew.frx":4710
         Style           =   1  'Graphical
         TabIndex        =   239
         Top             =   4155
         Width           =   1425
      End
      Begin VB.TextBox txtEtc 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   8
         Left            =   -74790
         TabIndex        =   238
         Top             =   6225
         Width           =   9405
      End
      Begin VB.TextBox txtEtc 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   7
         Left            =   -74790
         TabIndex        =   237
         Top             =   5955
         Width           =   9405
      End
      Begin VB.TextBox txtEtc 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   6
         Left            =   -74790
         TabIndex        =   236
         Top             =   5715
         Width           =   9405
      End
      Begin VB.TextBox txtEtc 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   5
         Left            =   -74790
         TabIndex        =   235
         Top             =   5475
         Width           =   9405
      End
      Begin VB.TextBox txtEtc 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   -74790
         TabIndex        =   234
         Top             =   4425
         Width           =   9405
      End
      Begin VB.TextBox txtEtc 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   -74790
         TabIndex        =   233
         Top             =   4665
         Width           =   9405
      End
      Begin VB.TextBox txtEtc 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   -74790
         TabIndex        =   232
         Top             =   4935
         Width           =   9405
      End
      Begin VB.TextBox txtEtc 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   4
         Left            =   -74790
         TabIndex        =   231
         Top             =   5205
         Width           =   9405
      End
      Begin VB.TextBox txtEtc 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   -74790
         TabIndex        =   230
         Top             =   4155
         Width           =   9405
      End
      Begin VB.CommandButton cmdSaveExt 
         Caption         =   "Save Details"
         Enabled         =   0   'False
         Height          =   915
         Left            =   -64260
         Picture         =   "frmEditAllNew.frx":4A1A
         Style           =   1  'Graphical
         TabIndex        =   229
         Top             =   5880
         Width           =   765
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "Delete"
         Height          =   990
         Left            =   -63180
         Picture         =   "frmEditAllNew.frx":4D24
         Style           =   1  'Graphical
         TabIndex        =   228
         Top             =   510
         Width           =   960
      End
      Begin VB.Frame Frame12 
         Caption         =   "Specimen Condition"
         Height          =   1035
         Index           =   1
         Left            =   -67845
         TabIndex        =   217
         Top             =   3840
         Width           =   3285
         Begin VB.CheckBox Ih 
            Caption         =   "Haemolysed"
            Height          =   225
            Index           =   1
            Left            =   1350
            TabIndex        =   223
            Top             =   450
            Width           =   1245
         End
         Begin VB.CheckBox Iis 
            Caption         =   "Slightly Haemolysed"
            Height          =   225
            Index           =   1
            Left            =   1350
            TabIndex        =   222
            Top             =   210
            Width           =   1755
         End
         Begin VB.CheckBox Il 
            Alignment       =   1  'Right Justify
            Caption         =   "Lipaemic"
            Height          =   225
            Index           =   1
            Left            =   300
            TabIndex        =   221
            Top             =   210
            Width           =   975
         End
         Begin VB.CheckBox Io 
            Alignment       =   1  'Right Justify
            Caption         =   "Old Sample"
            Height          =   225
            Index           =   1
            Left            =   120
            TabIndex        =   220
            Top             =   450
            Width           =   1155
         End
         Begin VB.CheckBox Ig 
            Caption         =   "Grossly Haemolysed"
            Height          =   225
            Index           =   1
            Left            =   1350
            TabIndex        =   219
            Top             =   690
            Width           =   1755
         End
         Begin VB.CheckBox Ij 
            Alignment       =   1  'Right Justify
            Caption         =   "Icteric"
            Height          =   225
            Index           =   1
            Left            =   510
            TabIndex        =   218
            Top             =   690
            Width           =   765
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Delta Check"
         Height          =   1905
         Index           =   1
         Left            =   -71175
         TabIndex        =   213
         Top             =   3840
         Width           =   3240
         Begin VB.Label lIDelta 
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H000000FF&
            Height          =   1515
            Index           =   1
            Left            =   135
            TabIndex        =   214
            Top             =   225
            Width           =   3030
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame Frame81 
         Caption         =   "Immunology Comments"
         Height          =   1905
         Index           =   1
         Left            =   -74685
         TabIndex        =   211
         Top             =   3795
         Width           =   3330
         Begin VB.TextBox txtImmComment 
            BackColor       =   &H80000018&
            Height          =   1545
            Index           =   1
            Left            =   90
            MaxLength       =   480
            MultiLine       =   -1  'True
            TabIndex        =   212
            Top             =   270
            Width           =   3135
         End
      End
      Begin VB.CommandButton cmdIremoveduplicates 
         Caption         =   "Remove Duplicates"
         Height          =   915
         Index           =   1
         Left            =   -65700
         Picture         =   "frmEditAllNew.frx":502E
         Style           =   1  'Graphical
         TabIndex        =   210
         Top             =   5880
         Width           =   885
      End
      Begin VB.CommandButton cmdIAdd 
         Caption         =   "Add Result"
         Height          =   915
         Index           =   1
         Left            =   -66495
         Picture         =   "frmEditAllNew.frx":5338
         Style           =   1  'Graphical
         TabIndex        =   209
         Tag             =   "bAdd"
         Top             =   5880
         Width           =   765
      End
      Begin VB.CommandButton cmdSaveImm 
         Caption         =   "Save Details"
         Enabled         =   0   'False
         Height          =   915
         Index           =   1
         Left            =   -64800
         Picture         =   "frmEditAllNew.frx":5642
         Style           =   1  'Graphical
         TabIndex        =   208
         Top             =   5880
         Width           =   765
      End
      Begin VB.CommandButton bValidateImm 
         Caption         =   "Validate"
         Height          =   915
         Index           =   1
         Left            =   -64020
         Picture         =   "frmEditAllNew.frx":594C
         Style           =   1  'Graphical
         TabIndex        =   207
         Top             =   5880
         Width           =   705
      End
      Begin VB.CommandButton bViewImmRepeat 
         BackColor       =   &H0000FFFF&
         Caption         =   "View Repeat"
         Height          =   915
         Index           =   1
         Left            =   -62580
         Picture         =   "frmEditAllNew.frx":5C56
         Style           =   1  'Graphical
         TabIndex        =   206
         Top             =   5880
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.CommandButton bImmRePrint 
         Caption         =   "Re-Print"
         Height          =   915
         Index           =   1
         Left            =   -63300
         Picture         =   "frmEditAllNew.frx":5DE0
         Style           =   1  'Graphical
         TabIndex        =   205
         Top             =   5880
         Width           =   720
      End
      Begin VB.ComboBox cISampleType 
         Height          =   315
         Index           =   1
         Left            =   -68910
         TabIndex        =   204
         Text            =   "cSampleType"
         Top             =   6225
         Width           =   1515
      End
      Begin VB.ComboBox cIUnits 
         Height          =   315
         Index           =   1
         Left            =   -70200
         TabIndex        =   203
         Text            =   "cUnits"
         Top             =   6225
         Width           =   1305
      End
      Begin VB.TextBox tINewValue 
         Height          =   315
         Index           =   1
         Left            =   -73200
         MaxLength       =   300
         TabIndex        =   202
         Top             =   6225
         Width           =   2970
      End
      Begin VB.ComboBox cIAdd 
         Height          =   315
         Index           =   1
         Left            =   -74820
         TabIndex        =   201
         Text            =   "cAdd"
         Top             =   6225
         Width           =   1575
      End
      Begin VB.CommandButton cmdGetBio 
         Caption         =   "Get Bio Tests"
         Height          =   870
         Left            =   -64290
         Picture         =   "frmEditAllNew.frx":60EA
         Style           =   1  'Graphical
         TabIndex        =   200
         Top             =   4455
         Width           =   1095
      End
      Begin VB.ComboBox cISampleType 
         Height          =   315
         Index           =   2
         Left            =   -70095
         TabIndex        =   199
         Text            =   "cSampleType"
         Top             =   4950
         Width           =   1515
      End
      Begin VB.CommandButton cmdIAdd 
         Caption         =   "Add Result"
         Height          =   960
         Index           =   2
         Left            =   -67800
         Picture         =   "frmEditAllNew.frx":63F4
         Style           =   1  'Graphical
         TabIndex        =   198
         Tag             =   "bAdd"
         Top             =   5880
         Width           =   765
      End
      Begin VB.ComboBox cIAdd 
         Height          =   315
         Index           =   2
         Left            =   -74640
         TabIndex        =   197
         Text            =   "cAdd"
         Top             =   4950
         Width           =   1575
      End
      Begin VB.TextBox tINewValue 
         Height          =   315
         Index           =   2
         Left            =   -72975
         MaxLength       =   15
         TabIndex        =   196
         Top             =   4950
         Width           =   1485
      End
      Begin VB.ComboBox cIUnits 
         Height          =   315
         Index           =   2
         Left            =   -71415
         TabIndex        =   195
         Text            =   "cUnits"
         Top             =   4950
         Width           =   1305
      End
      Begin VB.Frame Frame15 
         Caption         =   "Delta Check"
         Height          =   1905
         Left            =   -68430
         TabIndex        =   193
         Top             =   630
         Width           =   4785
         Begin VB.Label lBgaDelta 
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H000000FF&
            Height          =   1515
            Left            =   120
            TabIndex        =   194
            ToolTipText     =   "Delta Check"
            Top             =   270
            Width           =   4560
            WordWrap        =   -1  'True
         End
      End
      Begin VB.CommandButton bRePrintBga 
         Caption         =   "Re-Print"
         Height          =   960
         Left            =   -65235
         Picture         =   "frmEditAllNew.frx":66FE
         Style           =   1  'Graphical
         TabIndex        =   192
         Top             =   5880
         Width           =   765
      End
      Begin VB.CommandButton bViewBgaRepeat 
         BackColor       =   &H0000FFFF&
         Caption         =   "View Repeat"
         Height          =   960
         Left            =   -64380
         Picture         =   "frmEditAllNew.frx":6A08
         Style           =   1  'Graphical
         TabIndex        =   191
         Top             =   5880
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.CommandButton cmdHSaveH 
         Caption         =   "Save && Hold"
         Enabled         =   0   'False
         Height          =   825
         Left            =   -63735
         Picture         =   "frmEditAllNew.frx":6B92
         Style           =   1  'Graphical
         TabIndex        =   189
         Top             =   7365
         Width           =   915
      End
      Begin VB.CommandButton cmdDemoVal 
         Caption         =   "&Validate"
         Height          =   735
         Left            =   6255
         Picture         =   "frmEditAllNew.frx":6E9C
         Style           =   1  'Graphical
         TabIndex        =   177
         Top             =   5670
         Width           =   945
      End
      Begin Threed.SSPanel ssPanPgP 
         Height          =   735
         Left            =   6030
         TabIndex        =   176
         Top             =   4005
         Visible         =   0   'False
         Width           =   1770
         _Version        =   65536
         _ExtentX        =   3122
         _ExtentY        =   1296
         _StockProps     =   15
         Caption         =   "PGP"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelInner      =   1
         Alignment       =   0
         Begin VB.CheckBox chkPgp 
            Caption         =   "Endocrinology"
            Height          =   240
            Left            =   135
            TabIndex        =   13
            Top             =   315
            Width           =   1365
         End
      End
      Begin VB.TextBox txtCondition 
         Height          =   855
         Left            =   -64740
         MaxLength       =   500
         MultiLine       =   -1  'True
         TabIndex        =   151
         Top             =   1740
         Width           =   2850
      End
      Begin VB.CommandButton bFilm 
         Caption         =   "Film"
         Height          =   375
         Left            =   -65520
         Style           =   1  'Graphical
         TabIndex        =   143
         Top             =   6720
         Width           =   1155
      End
      Begin VB.CommandButton cmdValBG 
         Caption         =   "Validate"
         Height          =   960
         Left            =   -66090
         Picture         =   "frmEditAllNew.frx":71A6
         Style           =   1  'Graphical
         TabIndex        =   139
         Top             =   5880
         Width           =   765
      End
      Begin VB.CommandButton cmdSaveBGa 
         Caption         =   "Save Details"
         Enabled         =   0   'False
         Height          =   960
         Left            =   -66930
         Picture         =   "frmEditAllNew.frx":74B0
         Style           =   1  'Graphical
         TabIndex        =   138
         Top             =   5880
         Width           =   765
      End
      Begin VB.Frame Frame14 
         Caption         =   "Blood Gas Comments"
         Height          =   1905
         Left            =   -68475
         TabIndex        =   136
         Top             =   2565
         Width           =   4845
         Begin VB.TextBox txtBGaComment 
            BackColor       =   &H80000018&
            Height          =   1545
            Left            =   135
            MaxLength       =   320
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   137
            Top             =   225
            Width           =   4605
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Specimen Condition"
         Height          =   1035
         Index           =   0
         Left            =   -65190
         TabIndex        =   128
         Top             =   4365
         Width           =   3285
         Begin VB.CheckBox Ij 
            Alignment       =   1  'Right Justify
            Caption         =   "Icteric"
            Height          =   225
            Index           =   0
            Left            =   510
            TabIndex        =   134
            Top             =   690
            Width           =   765
         End
         Begin VB.CheckBox Ig 
            Caption         =   "Grossly Haemolysed"
            Height          =   225
            Index           =   0
            Left            =   1350
            TabIndex        =   133
            Top             =   690
            Width           =   1755
         End
         Begin VB.CheckBox Io 
            Alignment       =   1  'Right Justify
            Caption         =   "Old Sample"
            Height          =   225
            Index           =   0
            Left            =   120
            TabIndex        =   132
            Top             =   450
            Width           =   1155
         End
         Begin VB.CheckBox Il 
            Alignment       =   1  'Right Justify
            Caption         =   "Lipaemic"
            Height          =   225
            Index           =   0
            Left            =   300
            TabIndex        =   131
            Top             =   210
            Width           =   975
         End
         Begin VB.CheckBox Iis 
            Caption         =   "Slightly Haemolysed"
            Height          =   225
            Index           =   0
            Left            =   1350
            TabIndex        =   130
            Top             =   210
            Width           =   1755
         End
         Begin VB.CheckBox Ih 
            Caption         =   "Haemolysed"
            Height          =   225
            Index           =   0
            Left            =   1350
            TabIndex        =   129
            Top             =   450
            Width           =   1245
         End
      End
      Begin VB.ComboBox cISampleType 
         Height          =   315
         Index           =   0
         Left            =   -70455
         TabIndex        =   127
         Text            =   "cSampleType"
         Top             =   6150
         Width           =   1515
      End
      Begin VB.ComboBox cIUnits 
         Height          =   315
         Index           =   0
         Left            =   -71745
         TabIndex        =   126
         Text            =   "cUnits"
         Top             =   6150
         Width           =   1305
      End
      Begin VB.TextBox tINewValue 
         Height          =   315
         Index           =   0
         Left            =   -73290
         MaxLength       =   15
         TabIndex        =   125
         Top             =   6165
         Width           =   1485
      End
      Begin VB.ComboBox cIAdd 
         Height          =   315
         Index           =   0
         Left            =   -74880
         TabIndex        =   124
         Text            =   "cAdd"
         Top             =   6150
         Width           =   1575
      End
      Begin VB.CommandButton cmdIremoveduplicates 
         Caption         =   "Remove Duplicates"
         Height          =   915
         Index           =   0
         Left            =   -66270
         Picture         =   "frmEditAllNew.frx":77BA
         Style           =   1  'Graphical
         TabIndex        =   123
         Top             =   5880
         Width           =   885
      End
      Begin VB.CommandButton cmdIAdd 
         Caption         =   "Add Result"
         Height          =   915
         Index           =   0
         Left            =   -67080
         Picture         =   "frmEditAllNew.frx":7AC4
         Style           =   1  'Graphical
         TabIndex        =   122
         Tag             =   "bAdd"
         Top             =   5880
         Width           =   765
      End
      Begin VB.CommandButton cmdSaveImm 
         Caption         =   "Save Details"
         Enabled         =   0   'False
         Height          =   915
         Index           =   0
         Left            =   -65310
         Picture         =   "frmEditAllNew.frx":7DCE
         Style           =   1  'Graphical
         TabIndex        =   119
         Top             =   5880
         Width           =   765
      End
      Begin VB.CommandButton bValidateImm 
         Caption         =   "Validate"
         Height          =   915
         Index           =   0
         Left            =   -64440
         Picture         =   "frmEditAllNew.frx":80D8
         Style           =   1  'Graphical
         TabIndex        =   118
         Top             =   5880
         Width           =   765
      End
      Begin VB.Frame Frame81 
         Caption         =   "Endocrinology Comments"
         Height          =   1905
         Index           =   0
         Left            =   -65595
         TabIndex        =   115
         Top             =   2520
         Width           =   3690
         Begin VB.TextBox txtImmComment 
            BackColor       =   &H80000018&
            Height          =   1545
            Index           =   0
            Left            =   90
            MaxLength       =   320
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   116
            Top             =   270
            Width           =   3450
         End
      End
      Begin VB.CommandButton bViewImmRepeat 
         BackColor       =   &H0000FFFF&
         Caption         =   "View Repeat"
         Height          =   915
         Index           =   0
         Left            =   -62640
         Picture         =   "frmEditAllNew.frx":83E2
         Style           =   1  'Graphical
         TabIndex        =   114
         Top             =   5880
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.CommandButton bImmRePrint 
         Caption         =   "Re-Print"
         Height          =   915
         Index           =   0
         Left            =   -63525
         Picture         =   "frmEditAllNew.frx":856C
         Style           =   1  'Graphical
         TabIndex        =   113
         Top             =   5880
         Width           =   765
      End
      Begin VB.ComboBox cCunits 
         Height          =   315
         Left            =   -71670
         Style           =   2  'Dropdown List
         TabIndex        =   112
         Top             =   4890
         Width           =   1005
      End
      Begin VB.Frame Frame10 
         Caption         =   "Category"
         Height          =   825
         Index           =   0
         Left            =   6000
         TabIndex        =   111
         Top             =   3120
         Width           =   2385
         Begin VB.ComboBox cCat 
            Height          =   315
            Index           =   0
            Left            =   90
            Sorted          =   -1  'True
            TabIndex        =   21
            Top             =   270
            Width           =   2145
         End
      End
      Begin VB.CommandButton cmdPrintAll 
         Caption         =   "Print All"
         Height          =   915
         Left            =   -65130
         Picture         =   "frmEditAllNew.frx":8876
         Style           =   1  'Graphical
         TabIndex        =   105
         Top             =   5880
         Width           =   765
      End
      Begin VB.CommandButton bReprint 
         Caption         =   "Re-Print"
         Height          =   915
         Left            =   -64500
         Picture         =   "frmEditAllNew.frx":8B80
         Style           =   1  'Graphical
         TabIndex        =   104
         Top             =   5880
         Width           =   765
      End
      Begin VB.CheckBox cFilm 
         Caption         =   "Film"
         Height          =   195
         Left            =   -64260
         TabIndex        =   103
         Top             =   6780
         Width           =   645
      End
      Begin VB.CommandButton bHaemGraphs 
         Caption         =   "Graph"
         Height          =   825
         Left            =   -66480
         Picture         =   "frmEditAllNew.frx":8E8A
         Style           =   1  'Graphical
         TabIndex        =   100
         Top             =   7365
         Width           =   795
      End
      Begin VB.CommandButton cmdSaveInc 
         Caption         =   "&Save"
         Height          =   735
         Left            =   8640
         Picture         =   "frmEditAllNew.frx":92CC
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   5670
         Width           =   1155
      End
      Begin VB.Frame Frame9 
         Height          =   1035
         Left            =   -66045
         TabIndex        =   92
         Top             =   3000
         Visible         =   0   'False
         Width           =   2775
         Begin VB.CommandButton bPrintINR 
            Caption         =   "Print INR"
            Height          =   285
            Left            =   1290
            TabIndex        =   94
            Top             =   360
            Width           =   855
         End
         Begin VB.TextBox tWarfarin 
            Height          =   285
            Left            =   270
            MaxLength       =   5
            TabIndex        =   93
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Warfarin"
            Height          =   195
            Index           =   14
            Left            =   330
            TabIndex        =   95
            Top             =   150
            Width           =   600
         End
      End
      Begin VB.CommandButton bViewBioRepeat 
         BackColor       =   &H0000FFFF&
         Caption         =   "View Repeat"
         Height          =   915
         Left            =   -63630
         Picture         =   "frmEditAllNew.frx":95D6
         Style           =   1  'Graphical
         TabIndex        =   90
         Top             =   5880
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.TextBox txtHaemComment 
         Height          =   1185
         Left            =   -64740
         MaxLength       =   320
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   87
         ToolTipText     =   "Only 320 Characters"
         Top             =   360
         Width           =   2910
      End
      Begin VB.Frame Frame4 
         Height          =   5565
         Left            =   450
         TabIndex        =   69
         Top             =   1230
         Width           =   5445
         Begin VB.TextBox txtGpId 
            Height          =   285
            Left            =   4410
            TabIndex        =   253
            Top             =   315
            Width           =   915
         End
         Begin VB.CommandButton cmdCopyTo 
            Caption         =   "++ cc ++"
            Height          =   960
            Left            =   4995
            TabIndex        =   246
            Top             =   2925
            Width           =   375
         End
         Begin VB.ComboBox cmbHospital 
            Height          =   315
            Left            =   1050
            TabIndex        =   8
            Text            =   "cmbHospital"
            ToolTipText     =   "Hospital"
            Top             =   2565
            Width           =   3915
         End
         Begin VB.ComboBox cmbGP 
            Height          =   315
            Left            =   1050
            TabIndex        =   11
            ToolTipText     =   "Gp"
            Top             =   3600
            Width           =   3915
         End
         Begin VB.ComboBox cmbClinician 
            Height          =   315
            Left            =   1050
            TabIndex        =   10
            ToolTipText     =   "Clinician"
            Top             =   3285
            Width           =   3915
         End
         Begin VB.TextBox taddress 
            Height          =   285
            Index           =   1
            Left            =   750
            MaxLength       =   30
            TabIndex        =   7
            Tag             =   "Address Line 2"
            ToolTipText     =   "Address Line 2"
            Top             =   1875
            Width           =   4215
         End
         Begin VB.TextBox taddress 
            Height          =   285
            Index           =   0
            Left            =   750
            MaxLength       =   30
            TabIndex        =   6
            Tag             =   "Address Line 1"
            ToolTipText     =   "Address Line 1"
            Top             =   1605
            Width           =   4215
         End
         Begin VB.ComboBox cmbWard 
            Height          =   315
            Left            =   1050
            TabIndex        =   9
            ToolTipText     =   "Ward"
            Top             =   2940
            Width           =   3915
         End
         Begin VB.ComboBox cClDetails 
            Height          =   315
            Left            =   1050
            Sorted          =   -1  'True
            TabIndex        =   17
            ToolTipText     =   "Clinical Details"
            Top             =   4980
            Width           =   3915
         End
         Begin VB.TextBox txtDemographicComment 
            Height          =   990
            Left            =   1050
            MaxLength       =   160
            MultiLine       =   -1  'True
            TabIndex        =   15
            Tag             =   "Demographic Comment"
            ToolTipText     =   "Demographic Comment"
            Top             =   3930
            Width           =   3885
         End
         Begin VB.Label Label10 
            Caption         =   "GpID"
            Height          =   285
            Left            =   3960
            TabIndex        =   252
            Top             =   315
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Hospital"
            Height          =   195
            Left            =   420
            TabIndex        =   168
            Top             =   2640
            Width           =   570
         End
         Begin VB.Label lblNOPAS 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Index           =   1
            Left            =   2790
            TabIndex        =   141
            Top             =   315
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.Label Label35 
            Caption         =   "Nopas"
            Height          =   285
            Left            =   2295
            TabIndex        =   140
            Top             =   330
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "GP"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   8
            Left            =   765
            TabIndex        =   85
            Top             =   3630
            Width           =   225
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Clinician"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   7
            Left            =   405
            TabIndex        =   84
            Top             =   3330
            Width           =   585
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Comments"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   9
            Left            =   255
            TabIndex        =   83
            Top             =   3960
            Width           =   735
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Address"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   5
            Left            =   135
            TabIndex        =   82
            Top             =   1620
            Width           =   570
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Ward"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   6
            Left            =   600
            TabIndex        =   81
            Top             =   3000
            Width           =   390
         End
         Begin VB.Label Label23 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Sex"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   3660
            TabIndex        =   80
            Top             =   1200
            Width           =   270
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Age"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   3
            Left            =   2490
            TabIndex        =   79
            Top             =   1200
            Width           =   285
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "D.o.B"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   4
            Left            =   210
            TabIndex        =   78
            Top             =   1230
            Width           =   405
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Name"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   2
            Left            =   210
            TabIndex        =   77
            Top             =   810
            Width           =   420
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Chart #"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   90
            TabIndex        =   76
            Top             =   330
            Width           =   525
         End
         Begin VB.Label Label36 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Cl Details"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   330
            TabIndex        =   75
            Top             =   5040
            Width           =   660
         End
         Begin VB.Label lChart 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   720
            TabIndex        =   74
            Top             =   300
            Width           =   1485
         End
         Begin VB.Label lName 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   750
            TabIndex        =   73
            Top             =   780
            Width           =   4215
         End
         Begin VB.Label lDoB 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   750
            TabIndex        =   72
            Top             =   1230
            Width           =   1515
         End
         Begin VB.Label lAge 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2910
            TabIndex        =   71
            Top             =   1200
            Width           =   585
         End
         Begin VB.Label lSex 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   3990
            TabIndex        =   70
            Top             =   1200
            Width           =   705
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Biochemistry Comments"
         Height          =   1905
         Left            =   -65820
         TabIndex        =   28
         Top             =   2475
         Width           =   3165
         Begin VB.TextBox txtBioComment 
            BackColor       =   &H80000018&
            Height          =   1545
            Left            =   150
            MaxLength       =   320
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   29
            Tag             =   "Biochemistry Comment"
            ToolTipText     =   "Only 360 Characters"
            Top             =   240
            Width           =   2940
         End
      End
      Begin MSFlexGridLib.MSFlexGrid grdOutstanding 
         Height          =   5265
         Left            =   -67260
         TabIndex        =   30
         ToolTipText     =   "Outstanding Tests"
         Top             =   495
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   9287
         _Version        =   393216
         Cols            =   1
         FixedCols       =   0
         BackColor       =   -2147483624
         ForeColor       =   -2147483635
         BackColorFixed  =   -2147483647
         ForeColorFixed  =   -2147483624
         GridLines       =   0
         GridLinesFixed  =   0
         ScrollBars      =   2
         FormatString    =   "<Outstanding  "
      End
      Begin VB.CommandButton cmdSaveDemographics 
         Caption         =   "Save && &Hold"
         Enabled         =   0   'False
         Height          =   735
         Left            =   7335
         Picture         =   "frmEditAllNew.frx":9760
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   5670
         Width           =   1155
      End
      Begin VB.CommandButton cmdValidateCoag 
         Caption         =   "Validate"
         Height          =   915
         Left            =   -66060
         Picture         =   "frmEditAllNew.frx":9A6A
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   5880
         Width           =   765
      End
      Begin VB.CommandButton cmdSaveCoag 
         Caption         =   "Save Details"
         Enabled         =   0   'False
         Height          =   915
         Left            =   -66930
         Picture         =   "frmEditAllNew.frx":9D74
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   5880
         Width           =   765
      End
      Begin VB.CommandButton bValidateBio 
         Caption         =   "Validate"
         Height          =   915
         Left            =   -65415
         Picture         =   "frmEditAllNew.frx":A07E
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   5880
         Width           =   765
      End
      Begin VB.CommandButton cmdSaveBio 
         Caption         =   "Save Details"
         Enabled         =   0   'False
         Height          =   915
         Left            =   -66300
         Picture         =   "frmEditAllNew.frx":A388
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   5880
         Width           =   765
      End
      Begin VB.CommandButton bValidateHaem 
         Caption         =   "Validate"
         Height          =   825
         Left            =   -62790
         Picture         =   "frmEditAllNew.frx":A692
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   7365
         Width           =   765
      End
      Begin VB.CommandButton cmdSaveHaem 
         Caption         =   "Save"
         Enabled         =   0   'False
         Height          =   825
         Left            =   -64680
         Picture         =   "frmEditAllNew.frx":A99C
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   7365
         Width           =   825
      End
      Begin VB.CommandButton bViewHaemRepeat 
         BackColor       =   &H0000FFFF&
         Caption         =   "View Repeat"
         Height          =   825
         Left            =   -65670
         Picture         =   "frmEditAllNew.frx":ACA6
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   7365
         Width           =   960
      End
      Begin VB.CommandButton bViewCoagRepeat 
         BackColor       =   &H0000FFFF&
         Caption         =   "View Repeat"
         Height          =   915
         Left            =   -64260
         Picture         =   "frmEditAllNew.frx":AE30
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   5880
         Width           =   795
      End
      Begin VB.CommandButton bAddCoag 
         Caption         =   "Add Result"
         Height          =   915
         Left            =   -67800
         Picture         =   "frmEditAllNew.frx":AFBA
         Style           =   1  'Graphical
         TabIndex        =   40
         Tag             =   "bAdd"
         Top             =   5880
         Width           =   765
      End
      Begin VB.PictureBox Panel3D8 
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00FFFF80&
         Height          =   1425
         Left            =   -68760
         ScaleHeight     =   1365
         ScaleWidth      =   6870
         TabIndex        =   41
         Top             =   5100
         Width           =   6930
         Begin VB.VScrollBar VScroll1 
            Height          =   1215
            LargeChange     =   500
            Left            =   6540
            Max             =   2500
            SmallChange     =   100
            TabIndex        =   42
            Top             =   120
            Width           =   270
         End
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H80000008&
            Height          =   1215
            Left            =   120
            ScaleHeight     =   1185
            ScaleWidth      =   6195
            TabIndex        =   43
            Top             =   90
            Width           =   6225
            Begin VB.PictureBox pdelta 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H00C0C0C0&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Courier"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   4320
               Left            =   -540
               ScaleHeight     =   4320
               ScaleWidth      =   6945
               TabIndex        =   44
               Top             =   -480
               Width           =   6945
            End
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Delta Check"
         Height          =   1905
         Left            =   -65820
         TabIndex        =   45
         Top             =   540
         Width           =   3210
         Begin VB.Label ldelta 
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H000000FF&
            Height          =   1560
            Left            =   135
            TabIndex        =   46
            ToolTipText     =   "Delta Check"
            Top             =   270
            Width           =   2895
            WordWrap        =   -1  'True
         End
      End
      Begin VB.TextBox tResult 
         Height          =   315
         Left            =   -73200
         TabIndex        =   47
         Top             =   4890
         Width           =   1485
      End
      Begin VB.ComboBox cParameter 
         Height          =   315
         Left            =   -74760
         TabIndex        =   48
         Text            =   "cParameter"
         Top             =   4890
         Width           =   1545
      End
      Begin VB.CommandButton bAddBio 
         Caption         =   "Add Result"
         Height          =   915
         Left            =   -68070
         Picture         =   "frmEditAllNew.frx":B2C4
         Style           =   1  'Graphical
         TabIndex        =   49
         Tag             =   "bAdd"
         Top             =   5880
         Width           =   765
      End
      Begin VB.Frame Frame7 
         Caption         =   "Date"
         Height          =   1815
         Left            =   5985
         TabIndex        =   50
         Top             =   1260
         Width           =   5805
         Begin MSComCtl2.DTPicker dtRunDate 
            Height          =   315
            Left            =   2370
            TabIndex        =   20
            Top             =   1050
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            _Version        =   393216
            Format          =   61931521
            CurrentDate     =   36942
         End
         Begin MSComCtl2.DTPicker dtSampleDate 
            Height          =   315
            Left            =   690
            TabIndex        =   12
            Top             =   315
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            _Version        =   393216
            Format          =   61931521
            CurrentDate     =   36942
         End
         Begin MSMask.MaskEdBox tSampleTime 
            Height          =   315
            Left            =   2070
            TabIndex        =   18
            ToolTipText     =   "Time of Sample"
            Top             =   300
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
            Left            =   3540
            TabIndex        =   16
            Top             =   270
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            _Version        =   393216
            Format          =   61931521
            CurrentDate     =   36942
         End
         Begin MSMask.MaskEdBox tRecTime 
            Height          =   315
            Left            =   4920
            TabIndex        =   19
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
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Received"
            Height          =   195
            Index           =   0
            Left            =   2820
            TabIndex        =   152
            Top             =   330
            Width           =   690
         End
         Begin VB.Image iRecDate 
            Height          =   330
            Index           =   0
            Left            =   3540
            Picture         =   "frmEditAllNew.frx":B5CE
            Stretch         =   -1  'True
            ToolTipText     =   "Previous Day"
            Top             =   600
            Width           =   480
         End
         Begin VB.Image iRecDate 
            Height          =   330
            Index           =   1
            Left            =   4410
            Picture         =   "frmEditAllNew.frx":BA10
            Stretch         =   -1  'True
            ToolTipText     =   "Next Day"
            Top             =   600
            Width           =   480
         End
         Begin VB.Image iToday 
            Height          =   330
            Index           =   2
            Left            =   4020
            Picture         =   "frmEditAllNew.frx":BE52
            Stretch         =   -1  'True
            ToolTipText     =   "Set to Today"
            Top             =   600
            Width           =   360
         End
         Begin VB.Image iToday 
            Height          =   330
            Index           =   1
            Left            =   1170
            Picture         =   "frmEditAllNew.frx":C294
            Stretch         =   -1  'True
            ToolTipText     =   "Set to Today"
            Top             =   630
            Width           =   360
         End
         Begin VB.Image iToday 
            Height          =   330
            Index           =   0
            Left            =   2850
            Picture         =   "frmEditAllNew.frx":C6D6
            Stretch         =   -1  'True
            ToolTipText     =   "Set to Today"
            Top             =   1380
            Width           =   360
         End
         Begin VB.Image iSampleDate 
            Height          =   330
            Index           =   1
            Left            =   1560
            Picture         =   "frmEditAllNew.frx":CB18
            Stretch         =   -1  'True
            ToolTipText     =   "Next Day"
            Top             =   630
            Width           =   480
         End
         Begin VB.Image iSampleDate 
            Height          =   330
            Index           =   0
            Left            =   720
            Picture         =   "frmEditAllNew.frx":CF5A
            Stretch         =   -1  'True
            ToolTipText     =   "Previous Day"
            Top             =   630
            Width           =   480
         End
         Begin VB.Image iRunDate 
            Height          =   330
            Index           =   1
            Left            =   3240
            Picture         =   "frmEditAllNew.frx":D39C
            Stretch         =   -1  'True
            ToolTipText     =   "Next Day"
            Top             =   1380
            Width           =   480
         End
         Begin VB.Image iRunDate 
            Height          =   330
            Index           =   0
            Left            =   2340
            Picture         =   "frmEditAllNew.frx":D7DE
            Stretch         =   -1  'True
            ToolTipText     =   "Previous Day"
            Top             =   1380
            Width           =   480
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Run"
            Height          =   195
            Index           =   2
            Left            =   1980
            TabIndex        =   51
            Top             =   1110
            Width           =   300
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Sample"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   52
            Top             =   360
            Width           =   525
         End
      End
      Begin VB.Frame Frame5 
         Height          =   915
         Left            =   8490
         TabIndex        =   53
         Top             =   3045
         Width           =   1455
         Begin VB.CheckBox chkUrgent 
            Alignment       =   1  'Right Justify
            Caption         =   "Urgent"
            Height          =   195
            Left            =   135
            TabIndex        =   188
            Top             =   630
            Width           =   1185
         End
         Begin VB.OptionButton cRooH 
            Alignment       =   1  'Right Justify
            Caption         =   "Out of Hours"
            Height          =   195
            Index           =   1
            Left            =   90
            TabIndex        =   54
            Top             =   420
            Width           =   1215
         End
         Begin VB.OptionButton cRooH 
            Alignment       =   1  'Right Justify
            Caption         =   "Routine"
            Height          =   195
            Index           =   0
            Left            =   420
            TabIndex        =   55
            Top             =   180
            Width           =   885
         End
      End
      Begin VB.CommandButton bremoveduplicates 
         Caption         =   "Remove Duplicates"
         Height          =   915
         Left            =   -67260
         Picture         =   "frmEditAllNew.frx":DC20
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   5880
         Width           =   885
      End
      Begin VB.ComboBox cAdd 
         Height          =   315
         Left            =   -74910
         Sorted          =   -1  'True
         TabIndex        =   57
         Text            =   "cAdd"
         Top             =   6090
         Width           =   1575
      End
      Begin VB.TextBox tnewvalue 
         Height          =   315
         Left            =   -73320
         MaxLength       =   15
         TabIndex        =   58
         Top             =   6090
         Width           =   1575
      End
      Begin VB.ComboBox cUnits 
         Height          =   315
         Left            =   -71730
         TabIndex        =   59
         Text            =   "cUnits"
         Top             =   6090
         Width           =   1305
      End
      Begin VB.ComboBox cSampleType 
         Height          =   315
         Left            =   -70440
         TabIndex        =   60
         Text            =   "cSampleType"
         Top             =   6090
         Width           =   1515
      End
      Begin VB.Frame Frame2 
         Caption         =   "Specimen Condition"
         Height          =   1035
         Left            =   -65820
         TabIndex        =   61
         Top             =   4455
         Width           =   3240
         Begin VB.CheckBox oH 
            Caption         =   "Haemolysed"
            Height          =   225
            Left            =   1350
            TabIndex        =   62
            Top             =   450
            Width           =   1245
         End
         Begin VB.CheckBox oS 
            Caption         =   "Slightly Haemolysed"
            Height          =   225
            Left            =   1350
            TabIndex        =   63
            Top             =   210
            Width           =   1755
         End
         Begin VB.CheckBox oL 
            Alignment       =   1  'Right Justify
            Caption         =   "Lipaemic"
            Height          =   225
            Left            =   300
            TabIndex        =   64
            Top             =   210
            Width           =   975
         End
         Begin VB.CheckBox oO 
            Alignment       =   1  'Right Justify
            Caption         =   "Old Sample"
            Height          =   225
            Left            =   120
            TabIndex        =   65
            Top             =   450
            Width           =   1155
         End
         Begin VB.CheckBox oG 
            Caption         =   "Grossly Haemolysed"
            Height          =   225
            Left            =   1350
            TabIndex        =   66
            Top             =   675
            Width           =   1755
         End
         Begin VB.CheckBox oJ 
            Alignment       =   1  'Right Justify
            Caption         =   "Icteric"
            Height          =   225
            Left            =   510
            TabIndex        =   67
            Top             =   690
            Width           =   765
         End
      End
      Begin MSFlexGridLib.MSFlexGrid gBio 
         Height          =   5265
         Left            =   -74910
         TabIndex        =   68
         ToolTipText     =   "Biochemistry Results"
         Top             =   495
         Width           =   7635
         _ExtentX        =   13467
         _ExtentY        =   9287
         _Version        =   393216
         Cols            =   10
         BackColor       =   -2147483628
         ForeColor       =   -2147483635
         BackColorFixed  =   -2147483647
         ForeColorFixed  =   -2147483624
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         GridLines       =   3
         GridLinesFixed  =   3
         AllowUserResizing=   1
         FormatString    =   "<Test                  |<Result  |<Units    |^Ref Range  |^H/L|^   |^VP |^CP|^AL    |^Comment     "
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
      Begin MSFlexGridLib.MSFlexGrid grdCoag 
         Height          =   4275
         Left            =   -74775
         TabIndex        =   86
         Top             =   540
         Width           =   6825
         _ExtentX        =   12039
         _ExtentY        =   7541
         _Version        =   393216
         Cols            =   7
         FixedCols       =   0
         BackColor       =   -2147483624
         ForeColor       =   -2147483635
         BackColorFixed  =   -2147483647
         ForeColorFixed  =   -2147483624
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   0
         GridLines       =   3
         GridLinesFixed  =   3
         AllowUserResizing=   1
         FormatString    =   "<Parameter            |<Result    |<Units       |^Ref Range    |<Flag|^V |^P "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid grdOutstandingCoag 
         Height          =   4275
         Left            =   -67935
         TabIndex        =   99
         Top             =   540
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   7541
         _Version        =   393216
         Cols            =   1
         FixedCols       =   0
         BackColor       =   -2147483624
         ForeColor       =   -2147483635
         BackColorFixed  =   -2147483647
         ForeColorFixed  =   -2147483624
         GridLines       =   0
         GridLinesFixed  =   0
         ScrollBars      =   2
         FormatString    =   "<Outstanding  "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid grdPrev 
         Height          =   2175
         Left            =   -66000
         TabIndex        =   110
         ToolTipText     =   "Do Not give Out Results!!"
         Top             =   810
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   3836
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         BackColor       =   -2147483624
         BackColorFixed  =   -2147483647
         ForeColorFixed  =   -2147483624
         GridLines       =   3
         GridLinesFixed  =   3
         ScrollBars      =   2
         FormatString    =   "<Parameter            |<Result       |<Units          "
      End
      Begin MSFlexGridLib.MSFlexGrid grdOutstandings 
         Height          =   5265
         Index           =   0
         Left            =   -67080
         TabIndex        =   117
         Top             =   495
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   9287
         _Version        =   393216
         Cols            =   1
         FixedCols       =   0
         BackColor       =   -2147483624
         ForeColor       =   -2147483635
         BackColorFixed  =   -2147483647
         ForeColorFixed  =   -2147483624
         GridLines       =   0
         GridLinesFixed  =   0
         ScrollBars      =   2
         FormatString    =   "<Outstanding  "
      End
      Begin MSFlexGridLib.MSFlexGrid gImm 
         Height          =   5265
         Index           =   0
         Left            =   -74775
         TabIndex        =   135
         Top             =   495
         Width           =   7650
         _ExtentX        =   13494
         _ExtentY        =   9287
         _Version        =   393216
         Cols            =   8
         BackColor       =   -2147483624
         ForeColor       =   -2147483635
         BackColorFixed  =   -2147483647
         ForeColorFixed  =   -2147483624
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         GridLines       =   3
         GridLinesFixed  =   3
         ScrollBars      =   2
         AllowUserResizing=   1
         FormatString    =   "<Test                  |<Result              |<Units    |<Ref Range         |^H/L|^   |^VP |Comment  "
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
      Begin MSFlexGridLib.MSFlexGrid gBga 
         Height          =   4275
         Left            =   -74685
         TabIndex        =   190
         Top             =   630
         Width           =   6060
         _ExtentX        =   10689
         _ExtentY        =   7541
         _Version        =   393216
         Cols            =   7
         FixedCols       =   0
         BackColor       =   -2147483624
         ForeColor       =   -2147483635
         BackColorFixed  =   -2147483647
         ForeColorFixed  =   -2147483624
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   0
         GridLines       =   3
         GridLinesFixed  =   3
         AllowUserResizing=   1
         FormatString    =   "<Parameter            |<Result    |<Units       |^Ref Range    |<Flag|^V |^P "
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
      Begin MSFlexGridLib.MSFlexGrid grdOutstandings 
         Height          =   3330
         Index           =   1
         Left            =   -63300
         TabIndex        =   215
         Top             =   510
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   5874
         _Version        =   393216
         Cols            =   1
         FixedCols       =   0
         BackColor       =   -2147483624
         ForeColor       =   -2147483635
         BackColorFixed  =   -2147483647
         ForeColorFixed  =   -2147483624
         GridLines       =   0
         GridLinesFixed  =   0
         ScrollBars      =   2
         FormatString    =   "<Outstanding  "
      End
      Begin MSFlexGridLib.MSFlexGrid gImm 
         Height          =   3285
         Index           =   1
         Left            =   -74820
         TabIndex        =   216
         Top             =   495
         Width           =   11160
         _ExtentX        =   19685
         _ExtentY        =   5794
         _Version        =   393216
         Cols            =   9
         BackColor       =   -2147483624
         ForeColor       =   -2147483635
         BackColorFixed  =   -2147483647
         ForeColorFixed  =   -2147483624
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         GridLines       =   3
         GridLinesFixed  =   3
         ScrollBars      =   2
         AllowUserResizing=   1
         FormatString    =   $"frmEditAllNew.frx":DF2A
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
      Begin MSFlexGridLib.MSFlexGrid grdExt 
         Height          =   3585
         Left            =   -74820
         TabIndex        =   240
         Top             =   525
         Width           =   11490
         _ExtentX        =   20267
         _ExtentY        =   6324
         _Version        =   393216
         Cols            =   9
         FixedCols       =   2
         RowHeightMin    =   400
         BackColor       =   -2147483624
         ForeColor       =   -2147483635
         BackColorFixed  =   -2147483647
         ForeColorFixed  =   -2147483624
         WordWrap        =   -1  'True
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   0
         GridLines       =   3
         GridLinesFixed  =   3
         ScrollBars      =   2
         AllowUserResizing=   2
         FormatString    =   $"frmEditAllNew.frx":DFCB
      End
      Begin VB.Frame Frame10 
         Caption         =   "Category"
         Height          =   825
         Index           =   1
         Left            =   -64290
         TabIndex        =   174
         Top             =   1800
         Width           =   2385
         Begin VB.ComboBox cCat 
            Height          =   315
            Index           =   1
            Left            =   90
            Sorted          =   -1  'True
            TabIndex        =   175
            Top             =   270
            Width           =   2145
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Delta Check"
         Height          =   1425
         Index           =   0
         Left            =   -65640
         TabIndex        =   120
         Top             =   540
         Width           =   3735
         Begin VB.Label lIDelta 
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H000000FF&
            Height          =   1005
            Index           =   0
            Left            =   90
            TabIndex        =   121
            Top             =   270
            Width           =   3570
            WordWrap        =   -1  'True
         End
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   7425
         Left            =   -74880
         TabIndex        =   256
         ToolTipText     =   "Biochemistry Results"
         Top             =   360
         Width           =   6075
         _ExtentX        =   10716
         _ExtentY        =   13097
         _Version        =   393216
         Rows            =   30
         Cols            =   7
         BackColor       =   -2147483628
         ForeColor       =   -2147483635
         BackColorFixed  =   -2147483647
         ForeColorFixed  =   -2147483624
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         GridLines       =   3
         GridLinesFixed  =   3
         AllowUserResizing=   1
         FormatString    =   "<Test                  |<Result  |<Units    |^Ref Range  |^H/L|^VP |^AL     "
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
      Begin VB.Image Image1 
         Height          =   4515
         Left            =   -68760
         Top             =   420
         Width           =   3795
      End
      Begin VB.Label Label1 
         Caption         =   "Condition"
         Height          =   255
         Index           =   11
         Left            =   -64740
         TabIndex        =   255
         Top             =   1560
         Width           =   1065
      End
      Begin VB.Label Label7 
         Caption         =   "Analyser :"
         Height          =   195
         Left            =   -68640
         TabIndex        =   245
         Top             =   6600
         Width           =   690
      End
      Begin VB.Label lblAnalyser 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   -67920
         TabIndex        =   244
         Top             =   6600
         Width           =   2130
      End
      Begin VB.Label lblBgaDate 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   -73050
         TabIndex        =   241
         Top             =   5400
         Width           =   1605
      End
      Begin VB.Label Rundate 
         Caption         =   "Rundate"
         Height          =   255
         Index           =   5
         Left            =   -73695
         TabIndex        =   242
         Top             =   5430
         Width           =   675
      End
      Begin VB.Label lImmRan 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Random Sample"
         Height          =   465
         Index           =   1
         Left            =   -64290
         TabIndex        =   227
         ToolTipText     =   "Click to Toggle"
         Top             =   3960
         Width           =   1080
      End
      Begin VB.Label lblImmViewSplit 
         Alignment       =   2  'Center
         BackColor       =   &H00800080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Viewing Secondary Split"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Index           =   1
         Left            =   -74820
         TabIndex        =   226
         Top             =   5910
         Width           =   5955
      End
      Begin VB.Label lblIRundate 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   -71790
         TabIndex        =   225
         Top             =   6585
         Width           =   1515
      End
      Begin VB.Label Rundate 
         Caption         =   "Rundate"
         Height          =   255
         Index           =   4
         Left            =   -72480
         TabIndex        =   224
         Top             =   6615
         Width           =   675
      End
      Begin VB.Image An2 
         Height          =   645
         Left            =   -62490
         Top             =   1215
         Width           =   645
      End
      Begin VB.Image An1 
         Height          =   645
         Left            =   -62490
         Top             =   495
         Width           =   645
      End
      Begin VB.Label lblEDate 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   -72510
         TabIndex        =   186
         Top             =   6525
         Width           =   1515
      End
      Begin VB.Label Rundate 
         Caption         =   "Rundate"
         Height          =   255
         Index           =   0
         Left            =   -73155
         TabIndex        =   185
         Top             =   6555
         Width           =   675
      End
      Begin VB.Label lImmRan 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Random Sample"
         Height          =   375
         Index           =   0
         Left            =   -63255
         TabIndex        =   173
         ToolTipText     =   "Click to Toggle"
         Top             =   5445
         Width           =   1305
      End
      Begin VB.Label lblImmViewSplit 
         Alignment       =   2  'Center
         BackColor       =   &H00800080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Viewing Secondary Split"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Index           =   0
         Left            =   -74880
         TabIndex        =   172
         Top             =   5850
         Width           =   5955
      End
      Begin VB.Label lblAss 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Associated  Glucose 1"
         Height          =   705
         Left            =   -62760
         TabIndex        =   169
         Top             =   6075
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label Rundate 
         Caption         =   "Rundate"
         Height          =   255
         Index           =   3
         Left            =   -74610
         TabIndex        =   150
         Top             =   5310
         Width           =   675
      End
      Begin VB.Label lCDate 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   -73830
         TabIndex        =   149
         Top             =   5280
         Width           =   2295
      End
      Begin VB.Label Rundate 
         Caption         =   "Rundate"
         Height          =   255
         Index           =   2
         Left            =   -73710
         TabIndex        =   148
         Top             =   6510
         Width           =   675
      End
      Begin VB.Label lBDate 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   -73020
         TabIndex        =   147
         Top             =   6480
         Width           =   1515
      End
      Begin VB.Label Rundate 
         Caption         =   "Rundate"
         Height          =   255
         Index           =   1
         Left            =   -68640
         TabIndex        =   146
         Top             =   6960
         Width           =   675
      End
      Begin VB.Label lHDate 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   -67920
         TabIndex        =   145
         Top             =   6960
         Width           =   2115
      End
      Begin VB.Label lblPrevCoag 
         Alignment       =   2  'Center
         BackColor       =   &H80000001&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "No Chart # for Previous Details"
         ForeColor       =   &H80000018&
         Height          =   255
         Left            =   -66030
         TabIndex        =   109
         Top             =   540
         Width           =   2715
      End
      Begin VB.Label lblHaemValid 
         AutoSize        =   -1  'True
         Caption         =   "Already Validated"
         Height          =   195
         Left            =   -63120
         TabIndex        =   106
         Top             =   7080
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.Label lHaemErrors 
         AutoSize        =   -1  'True
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FLAGS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   -64260
         TabIndex        =   102
         ToolTipText     =   "Click Here to Show Flags"
         Top             =   2940
         Width           =   1065
      End
      Begin VB.Label lRandom 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Random Sample"
         Height          =   465
         Left            =   -62715
         TabIndex        =   101
         ToolTipText     =   "Click to Toggle"
         Top             =   5535
         Width           =   870
      End
      Begin VB.Label lblHaemPrinted 
         AutoSize        =   -1  'True
         Caption         =   "Already Printed"
         Height          =   195
         Left            =   -63060
         TabIndex        =   98
         Top             =   6840
         Width           =   1065
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Coagulation Comments"
         Height          =   195
         Left            =   -65940
         TabIndex        =   96
         Top             =   4050
         Width           =   1620
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Haematology Comment"
         Height          =   195
         Index           =   10
         Left            =   -72795
         TabIndex        =   88
         Top             =   4380
         Width           =   1635
      End
      Begin VB.Label lblViewSplit 
         Alignment       =   2  'Center
         BackColor       =   &H00800080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Viewing Secondary Split"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   -74910
         TabIndex        =   107
         Top             =   5790
         Width           =   6015
      End
   End
End
Attribute VB_Name = "frmEditAllNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mNewRecord As Boolean

Private PreviousImm As Boolean
Private PreviousBio As Boolean
Private PreviousHaem As Boolean
Private PreviousCoag As Boolean
Private PreviousBga As Boolean
Private PreviousEnd As Boolean
Private PreviousExt As Boolean

Private HistImm As Boolean
Private HistBio As Boolean
Private HistHaem As Boolean
Private HistCoag As Boolean
Private HistBga As Boolean
Private HistEnd As Boolean
Private HistExt As Boolean

Private BioChanged As Boolean
Private ImmChanged As Boolean
Private EndChanged As Boolean

Private HaemLoaded As Boolean
Private BioLoaded As Boolean
Private CoagLoaded As Boolean
Private ImmLoaded As Boolean
Private EndLoaded As Boolean
Private BgaLoaded As Boolean
Private ExtLoaded As Boolean
Private Activated As Boolean
Private UrgentTest As Boolean
Private HaemAnalyser As String
Private pPrintToPrinter As String

Private Sub bAddBio_Click()

Dim tb As New Recordset
Dim SQL As String
Dim n As Long
Dim s As String

On Error GoTo bAddBio_Click_Error

pBar = 0

txtSampleID = Format(Val(txtSampleID))


For n = 1 To gBio.Rows - 1
  If cAdd = gBio.TextMatrix(n, 0) Then
    iMsg "Test already Exists. Please delete before adding!"
    Exit Sub
  End If
Next

s = Check_Bio(cAdd.Text, cUnits, cSampleType)
If s <> "" Then
  iMsg s & " is incorrect!"
  Exit Sub
End If

If cAdd.Text = "" Then Exit Sub
If Val(txtSampleID) = 0 Then Exit Sub
If Len(cUnits) = 0 Then
  If iMsg("SELECT Units?", vbYesNo) = vbYes Then
    Exit Sub
  End If
End If
  
SQL = "INSERT into BioResults " & _
      "(RunDate, SampleID, Code, Result, RunTime, Units, SampleType, Valid, Printed) VALUES " & _
      "('" & Format$(dtRunDate, "dd/mmm/yyyy") & "', " & _
      "'" & txtSampleID & "', " & _
      "'" & CodeForShortName(cAdd.Text) & "', " & _
      "'" & tnewvalue & "', " & _
      "'" & Format$(Now, "dd/mmm/yyyy hh:mm") & "', " & _
      "'" & cUnits & "', " & _
      "'" & ListCodeFor("ST", cSampleType) & "', 0, 0);"
      
Set tb = New Recordset
RecOpenServer 0, tb, SQL



'Code added 22/08/05
'This allows the user delete
'oustanding requests where sample is bad
'it also marks bad samples printed and valid
If SysOptBioCodeForBad(0) = CodeForShortName(cAdd.Text) Then
    SQL = "update bioresults set valid = 1, printed = 1 " & _
          "where code = '" & SysOptBioCodeForBad(0) & "' " & _
          "and sampleID = '" & txtSampleID & "'"
    Cnxn(0).Execute SQL
  If iMsg("Do you wish all outstanding requests Deleted!", vbYesNo) = vbYes Then
    SQL = "DELETE from biorequests WHERE sampleID = '" & txtSampleID & "'"
    Cnxn(0).Execute SQL
  End If
  txtBioComment = iBOX("Enter Bad Comment")
End If


LoadBiochemistry

cAdd = ""
tnewvalue = ""
cUnits = ""


Exit Sub

bAddBio_Click_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /bAddBio_Click ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub bAddCoag_Click()

Dim Code As String
Dim s As String
Dim SQL As String
Dim Num As Long

On Error GoTo bAddCoag_Click_Error

pBar = 0

If cParameter = "" Then Exit Sub
If Trim$(tResult) = "" Then Exit Sub

For Num = 1 To grdCoag.Rows - 1
  If grdCoag.TextMatrix(Num, 0) = cParameter Then
    iMsg "Result already exists!"
    Exit Sub
  End If
Next

Code = CoagCodeFor(cParameter)
s = cParameter & vbTab & _
    tResult & vbTab & _
    cCunits & vbTab & _
    vbTab & _
    ""
grdCoag.AddItem s

If grdCoag.TextMatrix(1, 0) = "" Then
  grdCoag.RemoveItem 1
End If

SQL = "INSERT into CoagResults " & _
      "(RunDate, SampleID, Code, Result, RunTime, Units, Valid, Printed) VALUES " & _
      "('" & Format$(dtRunDate, "dd/mmm/yyyy") & "', " & _
      "'" & txtSampleID & "', " & _
      "'" & Trim(CoagCodeFor(cParameter.Text)) & "', " & _
      "'" & tResult & "', " & _
      "'" & Format$(Now, "dd/mmm/yyyy hh:mm") & "', " & _
      "'" & cCunits & "', " & _
      "0, 0);"

Cnxn(0).Execute SQL

'Code added 22/08/05
'Remove Coag requests if required
'set bad result to valid and printed

If SysOptCBad(0) = CoagCodeFor(cParameter) Then
    SQL = "update coagresults set valid = 1, printed = 1 " & _
          "where code = '" & SysOptCBad(0) & "' " & _
          "and sampleID = '" & txtSampleID & "'"
    Cnxn(0).Execute SQL
  If iMsg("Do you wish all outstanding requests Deleted!", vbYesNo) = vbYes Then
    SQL = "DELETE from coagrequests WHERE sampleID = '" & txtSampleID & "'"
    Cnxn(0).Execute SQL
  End If
End If

For Num = 1 To grdOutstandingCoag.Rows - 1
  If grdOutstandingCoag.TextMatrix(Num, 0) = cParameter Then
      SQL = "DELETE from coagRequests WHERE " & _
            "SampleID = '" & txtSampleID & "' " & _
            "and code = '" & CoagCodeFor(cParameter) & "'"
      Cnxn(0).Execute SQL
      LoadOutstandingrdCoag
      Exit For
  End If
Next
    
LoadCoagulation

cParameter = ""
tResult = ""
cCunits.ListIndex = -1
'cmdSaveCoag.Enabled = True
cmdValidateCoag.Enabled = True
cmdValidateCoag.Caption = "&Validate"

Exit Sub

bAddCoag_Click_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /bAddCoag_Click ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub baddtotests_Click()

frmAddToTests.Show 1

End Sub

Private Sub bCancel_Click()

pBar = 0

Unload Me

End Sub

Private Sub bcleardiff_click()
Dim n As Long
Dim A As Long

On Error GoTo bcleardiff_click_Error

pBar = 0

'If SysOptHaemAn1(0) <> "ADVIA" Then
  lWIC = ""
  lWOC = ""
'End If

txtMPXI = ""
txtLI = ""

grdH.Visible = False
For n = 1 To 6
  For A = 0 To 3 Step 3
    grdH.Row = n
    grdH.Col = A
    grdH.CellBackColor = &HFFFFFF
    grdH.CellForeColor = 1
    grdH = ""
  Next
Next

grdH.Visible = True

cmdSaveHaem.Enabled = True: cmdHSaveH.Enabled = True
bValidateHaem.Enabled = True

Exit Sub

bcleardiff_click_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /bcleardiff_click ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub bDoB_Click()

On Error GoTo bDoB_Click_Error

pBar = 0

With frmPatHistoryNew
  If Hospname(0) = "Monaghan" And sstabAll.Tab = 0 Then
    .oHD(0) = True
  Else
    .oHD(1) = True
  End If
  .oFor(2) = True
  .txtName = txtDoB
  If cmdDemoVal.Caption = "VALID" Then .mDemoVal = True Else .mDemoVal = False
  .FromEdit = True
  .EditScreen = Me
  .bsearch = True
  If .g.TextMatrix(1, 13) <> "" Then
    .Show 1
  Else
    FlashNoPrevious
  End If
End With

Exit Sub

bDoB_Click_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /bDoB_Click ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub bFAX_Click()
Dim tb As New Recordset
Dim SQL As String
Dim FaxNumber As String

On Error GoTo bFAX_Click_Error

pBar = 0


If sstabAll.Tab = 1 And lblHaemValid.Visible = False Then
  iMsg "Haematology not Validated"
  Exit Sub
End If
  

pBar = 0

If Trim$(txtSex) = "" Then
  If iMsg("Sex not entered." & vbCrLf & "Do you want to enter sex now?", vbQuestion + vbYesNo) = vbYes Then
    Exit Sub
  End If
End If

If Trim$(txtSampleID) = "" Then
  iMsg "Must have Lab Number.", vbCritical
  Exit Sub
End If

If Len(cmbWard) = 0 Then
  iMsg "Must have Ward entry.", vbCritical
  Exit Sub
End If

If UCase(Trim$(cmbWard)) = "GP" Then
  If Len(cmbGP) = 0 Then
    iMsg "Must have Ward or GP entry.", vbCritical
    Exit Sub
  End If
End If


If UCase(cmbWard) = "GP" Then
  SQL = "SELECT * from GPS WHERE text = '" & cmbGP & "'"
  Set tb = New Recordset
  RecOpenServer 0, tb, SQL
    If Not tb.EOF Then
      FaxNumber = tb!FAX
    End If
Else
  SQL = "SELECT * from wards WHERE text = '" & cmbWard & "'"
  Set tb = New Recordset
  RecOpenServer 0, tb, SQL
    If Not tb.EOF Then
      FaxNumber = tb!FAX
    End If
 End If


FaxNumber = iBOX("Faxnumber ", , FaxNumber)

FaxNumber = Trim(FaxNumber)

If Trim(FaxNumber) = "" Then
  iMsg "No Fax Number Entered!" & vbCrLf & "Fax Cancelled!"
  Exit Sub
End If


If Not IsNumeric(FaxNumber) Then
  iMsg "Incorrect Fax Number Entered!" & vbCrLf & "Fax Cancelled!"
  Exit Sub
End If


If Len(FaxNumber) < 4 Then
  iMsg "Incorrect Fax Number Entered!" & vbCrLf & "Fax Cancelled!"
  Exit Sub
End If

SaveDemographics

If SysOptFaxCom(0) Then
  If sstabAll.Tab <> 0 Then
    SQL = "SELECT * from PrintPending WHERE " & _
          "Department = 'M' " & _
          "and SampleID = '" & txtSampleID & "'"
    Set tb = New Recordset
    RecOpenServer 0, tb, SQL
    If tb.EOF Then
      tb.AddNew
    End If
    tb!Ward = cmbWard
    tb!Clinician = cmbClinician
    tb!GP = cmbGP
    tb!SampleID = txtSampleID
    tb!Department = "M"
    tb!Initiator = UserName
    tb!UsePrinter = pPrintToPrinter
    tb!FaxNumber = FaxNumber
    tb.Update
  End If
Else
  If sstabAll.Tab <> 0 Then
    LogTimeOfPrinting txtSampleID, Choose(sstabAll.Tab, "H", "B", "C", "E", "Q", "I", "")
    SQL = "SELECT * from PrintPending WHERE " & _
          "Department = '" & Choose(sstabAll.Tab, "H", "B", "C", "E", "Q", "I", "") & "' " & _
          "and SampleID = '" & txtSampleID & "'"
    Set tb = New Recordset
    RecOpenClient 0, tb, SQL
    If tb.EOF Then
      tb.AddNew
    End If
    tb!SampleID = txtSampleID
    tb!Department = Choose(sstabAll.Tab, "H", "B", "C", "E", "Q", "I", "")
    If SysOptRealImm(0) And tb!Department = "I" Then tb!Department = "J"
    tb!Initiator = UserName
    tb!Ward = cmbWard
    tb!Clinician = cmbClinician
    tb!GP = cmbGP
    tb!UsePrinter = pPrintToPrinter
    tb!FaxNumber = FaxNumber
    tb!ptime = Now
    tb.Update
  End If
End If

Exit Sub

bFAX_Click_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /bFAX_Click ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub bFilm_Click()

On Error GoTo bFilm_Click_Error

With frmDifferentials
  If bFilm.BackColor = vbBlue Then
    .LoadDiff = True
  End If
  .SampleID = txtSampleID
  .lWBC = tWBC
  .Show 1
    .LoadDiff = False
End With

Exit Sub

bFilm_Click_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /bFilm_Click ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub bHaemGraphs_Click()

frmHaemGraphs.SampleID = txtSampleID
frmHaemGraphs.Show 1

End Sub

Private Sub bHistory_Click()

On Error GoTo bHistory_Click_Error

pBar = 0

Select Case sstabAll.Tab
  Case 1:
    With frmFullHaem
      .lblChart = txtChart
      .lblName = txtName
      .lblDoB = txtDoB
      .lblNOPAS = txtNOPAS
      .Tn = "0"
      .Show 1
    End With
  Case 2:
    With frmFullBio
      .lblChart = txtChart
      .lblName = txtName
      .lblDoB = txtDoB
      .lblNOPAS = txtNOPAS
      .lblAandE = txtAandE
      .Tn = "0"
      .Show 1
    End With
  Case 3:
    With frmFullCoag
      .lblChart = txtChart
      .lblName = txtName
      .lblDoB = txtDoB
      .lblNOPAS = txtNOPAS
      .Tn = "0"
      .Show 1
    End With
  Case 4:
    With frmFullEnd
      .lblChart = txtChart
      .lblName = txtName
      .lblDoB = txtDoB
      .lblNOPAS = txtNOPAS
      .Tn = "0"
      .Show 1
    End With
  Case 5:
    With frmFullBga
      .lblChart = txtChart
      .lblName = txtName
      .lblDoB = txtDoB
      .lblNOPAS = txtNOPAS
      .Tn = "0"
      .Show 1
    End With
  Case 6:
    With frmFullImm
      .lblChart = txtChart
      .lblName = txtName
      .lblDoB = txtDoB
      .lblNOPAS = txtNOPAS
      .Tn = "0"
      .Show 1
    End With
  Case 7:
    With frmFullExt
      .lblChart = txtChart
      .lblName = txtName
      .lblDoB = txtDoB
      .lblNOPAS = txtNOPAS
      .Tn = "0"
      .Show 1
    End With
End Select


'With ffull
'  .lChart = txtchart
'  .lName = tName
'  .lDoB = txtdob
'  .Show 1
'End With

Exit Sub

bHistory_Click_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /bHistory_Click ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub bImmRePrint_Click(Index As Integer)
Dim tb As New Recordset
Dim SQL As String

On Error GoTo bImmRePrint_Click_Error

  pBar = 0

  txtSampleID = Format(Val(txtSampleID))
  If Val(txtSampleID) = 0 Then Exit Sub

  If Trim$(txtSex) = "" Then
    If iMsg("Sex not entered." & vbCrLf & "Do you want to enter sex now?", vbQuestion + vbYesNo) = vbYes Then
      Exit Sub
    End If
  End If
  
  If Trim$(txtSampleID) = "" Then
    iMsg "Must have Lab Number.", vbCritical
    Exit Sub
  End If
  
  If Trim$(cmbWard) = "" Then
    iMsg "Must have Ward entry.", vbCritical
    Exit Sub
  End If
  
  If Trim$(cmbWard) = "GP" Then
    If Trim$(cmbGP) = "" Then
      iMsg "Must have Ward or GP entry.", vbCritical
      Exit Sub
    End If
  End If

  If SysOptDemoVal(0) And cmdDemoVal.Caption = "&Validate" Then
    If iMsg("Do you wish to validate demographics !", vbYesNo) = vbNo Then
      Exit Sub
    Else
      cmdDemoVal_Click
    End If
  End If

  SaveDemographics

If Index = 0 Then
  LogTimeOfPrinting txtSampleID, "E"
  SQL = "UPDATE EndResults " & _
        "Set Printed = '0', Valid = 1 WHERE " & _
        "SampleID = '" & txtSampleID & "'"
  Cnxn(0).Execute SQL
  SQL = "SELECT * from PrintPending WHERE " & _
        "Department = 'E' " & _
        "and SampleID = '" & txtSampleID & "'"
  Set tb = New Recordset
  RecOpenClient 0, tb, SQL
  If tb.EOF Then
    tb.AddNew
  End If
  tb!Ward = cmbWard
  tb!Clinician = cmbClinician
  tb!GP = cmbGP
  tb!SampleID = txtSampleID
  tb!Department = "E"
  tb!Initiator = UserName
  tb!UsePrinter = pPrintToPrinter
  tb.Update
Else
  LogTimeOfPrinting txtSampleID, "I"
  SQL = "UPDATE ImmResults " & _
        "Set Printed = '0', Valid = 1 WHERE " & _
        "SampleID = '" & txtSampleID & "'"
  Cnxn(0).Execute SQL
  
  If SysOptRealImm(0) Then
    SQL = "SELECT * from PrintPending WHERE " & _
          "Department = 'J' " & _
          "and SampleID = '" & txtSampleID & "'"
  Else
    SQL = "SELECT * from PrintPending WHERE " & _
          "Department = 'I' " & _
          "and SampleID = '" & txtSampleID & "'"
  End If
  Set tb = New Recordset
  RecOpenClient 0, tb, SQL
  If tb.EOF Then
    tb.AddNew
  End If
  tb!SampleID = txtSampleID
  tb!Ward = cmbWard
  tb!Clinician = cmbClinician
  tb!GP = cmbGP
  If SysOptRealImm(0) Then tb!Department = "J" Else tb!Department = "I"
  tb!Initiator = UserName
  tb!UsePrinter = pPrintToPrinter
  tb.Update
End If

Exit Sub

bImmRePrint_Click_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /bImmRePrint_Click ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select

End Sub

Private Sub bOrderTests_Click()

On Error GoTo bOrderTests_Click_Error

pBar = 0

If cmdSaveDemographics.Enabled = True Or cmdSaveInc.Enabled = True Then
  If iMsg("Save Demographics!", vbYesNo) = vbYes Then
    cmdSaveDemographics_Click
  End If
End If

With frmNewOrder
  .FromEdit = True
  .SampleID = Format(Val(txtSampleID))
  .Show 1
End With

If SysOptDeptEnd(0) Then LoadOutstandingEnd
If SysOptDeptImm(0) Then LoadOutstandingImm
If SysOptDeptBio(0) Then LoadOutstandingBio
If SysOptDeptCoag(0) Then LoadOutstandingrdCoag
'If SysOptDeptHaem Then loadoutstandingHaem

LoadDemographics

Exit Sub

bOrderTests_Click_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /bOrderTests_Click ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub bremoveduplicates_Click()

Dim tb As New Recordset
Dim SQL As String
Dim Y As Long
Dim Code As String
Dim Result As String
On Error GoTo bremoveduplicates_Click_Error

pBar = 0

If gBio.Rows < 3 Then Exit Sub

Screen.MousePointer = 11

For Y = 1 To gBio.Rows - 1
  Code = CodeForShortName(gBio.TextMatrix(Y, 0))
  Result = gBio.TextMatrix(Y, 1)
  SQL = "SELECT * from bioresults WHERE " & _
        "sampleid = '" & txtSampleID & "' " & _
        "and code = '" & Code & "'  order by runtime asc"
  Set tb = New Recordset
  RecOpenClient 0, tb, SQL
  If tb.RecordCount > 1 Then
    SQL = "DELETE from bioresults WHERE sampleid = '" & txtSampleID & "' and code = '" & Code & "' and runtime = '" & Format(tb!RunTime, "dd/MMM/yyyy hh:mm:ss") & "'"
    Cnxn(0).Execute SQL
  End If
Next

LoadBiochemistry

Screen.MousePointer = 0

Exit Sub

bremoveduplicates_Click_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /bremoveduplicates_Click ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub bReprint_Click()
7
Dim tb As New Recordset
Dim SQL As String

On Error GoTo bReprint_Click_Error

txtSampleID = Format(Val(txtSampleID))
If Val(txtSampleID) = 0 Then Exit Sub

pBar = 0

If Trim$(txtSex) = "" Then
  If iMsg("Sex not entered." & vbCrLf & "Do you want to enter sex now?", vbQuestion + vbYesNo) = vbYes Then
    Exit Sub
  End If
End If

If Trim$(txtSampleID) = "" Then
  iMsg "Must have Lab Number.", vbCritical
  Exit Sub
End If

If Trim$(cmbWard) = "" Then
  iMsg "Must have Ward entry.", vbCritical
  Exit Sub
End If

If Trim$(cmbWard) = "GP" Then
  If Trim$(cmbGP) = "" Then
    iMsg "Must have Ward or GP entry.", vbCritical
    Exit Sub
  End If
End If

SaveDemographics

LogTimeOfPrinting txtSampleID, "B"

SQL = "UPDATE BioResults " & _
      "Set Printed = '0', Valid = 1 WHERE " & _
      "SampleID = '" & txtSampleID & "' AND code <> '" & SysOptBioCodeForBad(0) & "'"
Cnxn(0).Execute SQL

SQL = "SELECT * from PrintPending WHERE " & _
      "Department = 'B' " & _
      "and SampleID = '" & txtSampleID & "'"
Set tb = New Recordset
RecOpenClient 0, tb, SQL
If tb.EOF Then
  tb.AddNew
End If
tb!SampleID = txtSampleID
tb!Ward = cmbWard
tb!Clinician = cmbClinician
tb!GP = cmbGP
tb!Department = "B"
tb!Initiator = UserName
tb!UsePrinter = pPrintToPrinter
tb!ptime = Now
tb.Update

LoadBiochemistry

Exit Sub

bReprint_Click_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /bReprint_Click ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub bRePrintBga_Click()
Dim tb As New Recordset
Dim SQL As String

On Error GoTo bReprint_Click_Error

txtSampleID = Format(Val(txtSampleID))
If Val(txtSampleID) = 0 Then Exit Sub

pBar = 0

If Trim$(txtSex) = "" Then
  If iMsg("Sex not entered." & vbCrLf & "Do you want to enter sex now?", vbQuestion + vbYesNo) = vbYes Then
    Exit Sub
  End If
End If

If Trim$(txtSampleID) = "" Then
  iMsg "Must have Lab Number.", vbCritical
  Exit Sub
End If

If Trim$(cmbWard) = "" Then
  iMsg "Must have Ward entry.", vbCritical
  Exit Sub
End If

If Trim$(cmbWard) = "GP" Then
  If Trim$(cmbGP) = "" Then
    iMsg "Must have Ward or GP entry.", vbCritical
    Exit Sub
  End If
End If

SaveDemographics

LogTimeOfPrinting txtSampleID, "G"

SQL = "UPDATE BgaResults " & _
      "Set Printed = '0', Valid = 1 WHERE " & _
      "SampleID = '" & txtSampleID & "'"
Cnxn(0).Execute SQL

SQL = "SELECT * from PrintPending WHERE " & _
      "Department = 'Q' " & _
      "and SampleID = '" & txtSampleID & "'"
Set tb = New Recordset
RecOpenClient 0, tb, SQL
If tb.EOF Then
  tb.AddNew
End If
tb!SampleID = txtSampleID
tb!Ward = cmbWard
tb!Clinician = cmbClinician
tb!GP = cmbGP
tb!Department = "Q"
tb!Initiator = UserName
tb!UsePrinter = pPrintToPrinter
tb!ptime = Now
tb.Update

LoadBloodGas

Exit Sub

bReprint_Click_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /bReprint_Click ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select

End Sub

Private Sub bsearch_Click()

On Error GoTo bsearch_Click_Error

pBar = 0

With frmPatHistoryNew
  If Hospname(0) = "Monaghan" And sstabAll.Tab = 0 Then
    .oHD(0) = True
'  ElseIf UCase(Hospname(0)) = "PORTLAOISE" And sstabAll.Tab = 0 Then
'    .oHD(0) = True
  Else
    .oHD(1) = True
  End If
  .oFor(0) = True
  .txtName = txtName
  If cmdDemoVal.Caption = "VALID" Then .mDemoVal = True Else .mDemoVal = False
  .FromEdit = True
  .EditScreen = Me
  .bsearch = True
  If .g.TextMatrix(1, 13) <> "" Then
    .Show 1
  Else
      FlashNoPrevious
    End If
  
End With

Exit Sub

bsearch_Click_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /bsearch_Click ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub bValidateBio_Click()

On Error GoTo bValidateBio_Click_Error

pBar = 0

txtSampleID = Format(Val(txtSampleID))
If Val(txtSampleID) = 0 Then Exit Sub

If bValidateBio.Caption = "VALID" Then
  If UCase(iBOX("Unvalidate ! Enter Password" & vbCrLf & "You get only 1 Chance!", , , True)) = UCase(UserPass) Then
    SaveBiochemistry False, True
    SaveComments
    Me.Refresh
  End If
Else
  If SysOptDemoVal(0) And cmdDemoVal.Caption = "&Validate" Then
    If iMsg("Do you wish to validate demographics !", vbYesNo) = vbNo Then
      Exit Sub
    Else
      cmdDemoVal_Click
    End If
  End If
  If txtDoB = "" Then iMsg "No Dob. Adult Age 25 used for Normal Ranges!"
  SaveBiochemistry True
  SaveComments
  UPDATEMRU
  Frame2.Enabled = False
  lRandom.Enabled = False
  txtBioComment.Locked = True
  Me.Refresh
  txtSampleID = Format$(Val(txtSampleID) + 1)
End If
LoadAllDetails

Exit Sub

bValidateBio_Click_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /bValidateBio_Click ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub bValidateHaem_Click()

On Error GoTo bValidateHaem_Click_Error

pBar = 0

txtSampleID = Format(Val(txtSampleID))
If Val(txtSampleID) = 0 Then Exit Sub

If bValidateHaem.Caption = "VALID" Then
  If UCase(iBOX("Unvalidate ! Enter Password" & vbCrLf & "You get only 1 Chance!", , , True)) = UserPass Then
    SaveHaematology False
    SaveComments
    Panel3D4.Enabled = True
    Panel3D5.Enabled = True
    Panel3D6.Enabled = True
'    Panel3D7.Enabled = True
'    txtHaemComment.Enabled = True
    bValidateHaem.Caption = "&Validate"
    lblHaemValid.Visible = False
    txtHaemComment.Locked = False
    LoadHaematology
    Me.Refresh
  Else
    Exit Sub
  End If
Else
  If SysOptDemoVal(0) And cmdDemoVal.Caption = "&Validate" Then
    If iMsg("Do you wish to validate demographics !", vbYesNo) = vbNo Then
      Exit Sub
    Else
      cmdDemoVal_Click
    End If
  End If
  
  If Trim(txtDoB) = "" Then iMsg "No Dob. Adult Age 25 used for Normal Ranges!"

  SaveHaematology 1
  SaveComments
  UPDATEMRU
  Panel3D4.Enabled = False
  Panel3D5.Enabled = False
  Panel3D6.Enabled = False
'  Panel3D7.Enabled = False
'  If SysOptCommVal(0) Then txtHaemComment.Enabled = False
  txtSampleID = Format$(Val(txtSampleID) + 1)
  LoadAllDetails
  Me.Refresh
End If

Exit Sub

bValidateHaem_Click_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /bValidateHaem_Click ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub bValidateImm_Click(Index As Integer)
On Error GoTo bValidateImm_Click_Error

pBar = 0

txtSampleID = Format(Val(txtSampleID))
If Val(txtSampleID) = 0 Then Exit Sub

If Index = 0 Then
  
  If bValidateImm(0).Caption = "VALID" Then
    If UCase(iBOX("Unvalidate ! Enter Password" & vbCrLf & "You get only 1 Chance!", , , True)) = UCase(UserPass) Then
      SaveEndocrinology False, True
      SaveComments
      Me.Refresh
    End If
  Else
    If SysOptDemoVal(0) And cmdDemoVal.Caption = "&Validate" Then
      If iMsg("Do you wish to validate demographics !", vbYesNo) = vbNo Then
        Exit Sub
      Else
        cmdDemoVal_Click
      End If
    End If
    If Trim(txtDoB) = "" Then iMsg "No Dob. Adult Age 25 used for Normal Ranges!"
    SaveEndocrinology True
    SaveComments
    UPDATEMRU
    Frame12(0).Enabled = False
    lImmRan(0).Enabled = False
    'txtImmComment(0).Locked = True
    Me.Refresh
    txtSampleID = Format$(Val(txtSampleID) + 1)
  End If
  

Else

  If bValidateImm(1).Caption = "VALID" Then
    If UCase(iBOX("Unvalidate ! Enter Password" & vbCrLf & "You get only 1 Chance!", , , True)) = UCase(UserPass) Then
      SaveImmunology False, True
          SaveComments
    End If
  Else
    If SysOptDemoVal(0) And cmdDemoVal.Caption = "&Validate" Then
      If iMsg("Do you wish to validate demographics !", vbYesNo) = vbNo Then
        Exit Sub
      Else
        cmdDemoVal_Click
      End If
    End If
    If Trim(txtDoB) = "" Then iMsg "No Dob. Adult Age 25 used for Normal Ranges!"
    SaveImmunology True
    SaveComments
    UPDATEMRU
    Frame12(1).Enabled = False
    lImmRan(1).Enabled = False
    'txtImmComment(1).Locked = True
    txtSampleID = Format$(Val(txtSampleID) + 1)
  End If

End If



    LoadAllDetails


Exit Sub

bValidateImm_Click_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /bValidateImm_Click ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub bViewBB_Click()

pBar = 0

If Trim$(txtChart) <> "" Then
  frmViewBB.lChart = txtChart
  frmViewBB.Show 1
End If

End Sub

Private Sub bViewBgaRepeat_Click()

pBar = 0

frmViewBgaRepeat.Show 1

End Sub

Private Sub bViewBioRepeat_Click()

pBar = 0

frmViewBioRepeat.Show 1

End Sub

Private Sub bViewCoagRepeat_Click()

pBar = 0

With frmCoagRepeats
  .EditForm = Me
  .SampleID = txtSampleID
  .Show 1
End With

End Sub

Private Sub bViewHaemRepeat_Click()

pBar = 0

With frmViewHaemRep
  .EditForm = Me
  .lSampleID = txtSampleID
  .lName = txtName
  .Show 1
End With

LoadHaematology

End Sub

Private Sub bViewImmRepeat_Click(Index As Integer)

pBar = 0

If Index = 0 Then
  frmViewEndRepeat.Show 1
Else
  frmViewImmRepeat.Show 1
End If

End Sub

Private Sub cAdd_Click()

On Error GoTo cAdd_Click_Error

pBar = 0

Dim SampleType As String
Dim tb As New Recordset
Dim SQL As String

cUnits.Enabled = True

SampleType = ListCodeFor("ST", cSampleType)

SQL = "SELECT * from biotestdefinitions WHERE code = '" & CodeForShortName(cAdd) & "'"
Set tb = New Recordset
RecOpenServer 0, tb, SQL
If Not tb.EOF Then
  cUnits = tb!Units
Else
  cUnits = ""
End If

cUnits.Enabled = False

Exit Sub

cAdd_Click_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /cAdd_Click ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub cAdd_KeyPress(KeyAscii As Integer)

KeyAscii = 0

End Sub

Private Sub cASot_Click()

If cASot = 0 Then
  If Trim$(tASOt) = "?" Then
    tASOt = ""
  ElseIf Trim$(tASOt) <> "" Then
    cASot = 1
  End If
Else
  If Trim$(tASOt) = "" Then
    tASOt = "?"
  End If
End If

cmdSaveHaem.Enabled = True: cmdHSaveH.Enabled = True

End Sub

Private Sub cCat_Change(Index As Integer)

If Index = 0 Then
    cmdSaveDemographics.Enabled = True
    cmdSaveInc.Enabled = True
    cCat(1) = cCat(0)
Else
    cCat(0) = cCat(1)
End If


End Sub

Private Sub cCat_Click(Index As Integer)
Dim SQL As String

On Error GoTo cCat_Click_Error

If Index = 0 Then
    cmdSaveDemographics.Enabled = True
    cmdSaveInc.Enabled = True
    cCat(1) = cCat(0)
Else
    If EndLoaded = True Then
        SQL = "UPDATE demographics set category = '" & cCat(1) & "' WHERE sampleid = " & txtSampleID & ""
        Cnxn(0).Execute SQL
        cCat(0) = cCat(1)

    End If
End If

Exit Sub

cCat_Click_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /cCat_Click ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub cClDetails_Click()

cmdSaveDemographics.Enabled = True
cmdSaveInc.Enabled = True

End Sub

Private Sub cClDetails_LostFocus()


On Error GoTo cClDetails_LostFocus_Error

pBar = 0

If Trim$(cClDetails) = "" Then Exit Sub

On Error Resume Next
If ListText("CD", cClDetails) <> "" Then
  cClDetails = ListText("CD", cClDetails)
End If

Exit Sub

cClDetails_LostFocus_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /cClDetails_LostFocus ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub cESR_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error GoTo cESR_MouseUp_Error

pBar = 0

If cESR = 0 Then
  If Trim$(tESR) = "?" Then
    tESR = ""
  ElseIf Trim$(tESR) <> "" Then
    cESR = 1
  End If
Else
  If Trim$(tESR) = "" Then
    tESR = "?"
  End If
End If

cmdSaveHaem.Enabled = True: cmdHSaveH.Enabled = True

Exit Sub

cESR_MouseUp_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /cESR_MouseUp ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub cFilm_Click()

cmdSaveHaem.Enabled = True: cmdHSaveH.Enabled = True

End Sub

Private Sub CheckAssGlucose(ByVal CurrentBRs As BIEResults)

Dim tb As New Recordset
Dim SQL As String

On Error GoTo CheckAssGlucose_Error

If CurrentBRs.Count = 1 Then
  If CurrentBRs(1).Code = SysOptBioCodeForGlucose(0) Then
    'check prev or next for general
    SQL = "SELECT distinct D.SampleID " & _
          "from Demographics as D " & _
          "WHERE D.sampleid in " & _
          "  (  SELECT SampleID from BioResults WHERE " & _
          "     (SampleID = '" & Val(txtSampleID) - 1 & "' or SampleID = '" & Val(txtSampleID) + 1 & "') " & _
          "     and Code <> '" & SysOptBioCodeForGlucose(0) & "'  ) " & _
          "and D.PatName = '" & AddTicks(txtName) & "' " & _
          "and (D.SampleID = '" & Val(txtSampleID) - 1 & "' or SampleID = '" & Val(txtSampleID) + 1 & "')"
    Set tb = New Recordset
    RecOpenServer 0, tb, SQL
    If Not tb.EOF Then
      lblAss = "Associated Results " & tb!SampleID
      lblAss.Visible = True
    End If
  Else
    SQL = "SELECT distinct D.SampleID " & _
          "from Demographics as D " & _
          "WHERE D.sampleid in " & _
          "  (  SELECT SampleID from BioResults WHERE " & _
          "     (SampleID = '" & Val(txtSampleID) - 1 & "' or SampleID = '" & Val(txtSampleID) + 1 & "') " & _
          "     and Code = '" & SysOptBioCodeForGlucose(0) & "'  ) " & _
          "and D.PatName = '" & AddTicks(txtName) & "' " & _
          "and (D.SampleID = '" & Val(txtSampleID) - 1 & "' or SampleID = '" & Val(txtSampleID) + 1 & "')"
    Set tb = New Recordset
    RecOpenServer 0, tb, SQL
    If Not tb.EOF Then
      lblAss = "Associated Glucose " & tb!SampleID
      lblAss.Visible = True
    End If
  End If
Else
  SQL = "SELECT distinct D.SampleID " & _
        "from Demographics as D " & _
        "WHERE D.sampleid in " & _
        "  (  SELECT SampleID from BioResults WHERE " & _
        "     (SampleID = '" & Val(txtSampleID) - 1 & "' or SampleID = '" & Val(txtSampleID) + 1 & "') " & _
        "     and Code = '" & SysOptBioCodeForGlucose(0) & "'  ) " & _
        "and D.PatName = '" & AddTicks(txtName) & "' " & _
        "and (D.SampleID = '" & Val(txtSampleID) - 1 & "' or SampleID = '" & Val(txtSampleID) + 1 & "')"
  Set tb = New Recordset
  RecOpenServer 0, tb, SQL
  If Not tb.EOF Then
    lblAss = "Associated Glucose " & tb!SampleID
    lblAss.Visible = True
  End If
End If

Exit Sub

CheckAssGlucose_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /CheckAssGlucose ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub CheckCalcEPSA(ByVal Ims As BIEResults)

Dim Im As BIEResult
Dim FPS As Single
Dim FPSTime As String
Dim FPSDate As String
Dim PSA As Single
Dim Ratio As Single
Dim Code As String

On Error GoTo CheckCalcEPSA_Error

If Ims Is Nothing Then Exit Sub

FPS = 0
PSA = 0
Ratio = 0

For Each Im In Ims
  Code = UCase$(Trim$(Im.Code))
  If Code = "FPS" Then
    FPS = Val(Im.Result)
    FPSDate = Im.Rundate
    FPSTime = Im.RunTime
  ElseIf Code = "PSA" Then
    PSA = Val(Im.Result)
  ElseIf Code = "FPR" Then
    Ratio = Val(Im.Result)
  End If
Next

If (FPS * PSA) <> 0 And Ratio = 0 Then
  Ratio = FPS / PSA
  Set Im = New BIEResult
  Im.SampleID = txtSampleID
  Im.Code = "FPR"
  Im.Rundate = FPSDate
  Im.RunTime = FPSTime
  Im.Result = Format$(Ratio, "#0.00")
  Im.Units = ""
  Im.Printed = 0
  Im.Valid = 0
  Ims.Add Im
  Ims.Save "End", Ims
End If

Exit Sub

CheckCalcEPSA_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /CheckCalcEPSA ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub CheckCalcIPSA(ByVal Ims As BIEResults)

Dim Im As BIEResult
Dim FPS As Single
Dim FPSTime As String
Dim FPSDate As String
Dim PSA As Single
Dim Ratio As Single
Dim Code As String

On Error GoTo CheckCalcIPSA_Error

If Ims Is Nothing Then Exit Sub

FPS = 0
PSA = 0
Ratio = 0

For Each Im In Ims
  Code = UCase$(Trim$(Im.Code))
  If Code = "FPS" Then
    FPS = Val(Im.Result)
    FPSDate = Im.Rundate
    FPSTime = Im.RunTime
  ElseIf Code = "PSA" Then
    PSA = Val(Im.Result)
  ElseIf Code = "FPR" Then
    Ratio = Val(Im.Result)
  End If
Next

If (FPS * PSA) <> 0 And Ratio = 0 Then
  Ratio = FPS / PSA
  Set Im = New BIEResult
  Im.SampleID = txtSampleID
  Im.Code = "FPR"
  Im.Rundate = FPSDate
  Im.RunTime = FPSTime
  Im.Result = Format$(Ratio, "#0.00")
  Im.Units = ""
  Im.Printed = 0
  Im.Valid = 0
  Ims.Add Im
  Ims.Save "Imm", Ims
End If

Exit Sub

CheckCalcIPSA_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /CheckCalcIPSA ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

'Added 15/Jul/2004
'cmdSaveHaem_Click change
'cmdSaveBio_Click change
'cmdSaveCoag_click change
'cmdSaveImm_Click


Private Sub CheckCalcPSA(ByVal BRs As BIEResults)

Dim br As BIEResult
Dim FPS As Single
Dim FPSTime As String
Dim FPSDate As String
Dim PSA As Single
Dim Ratio As Single
Dim Code As String

On Error GoTo CheckCalcPSA_Error

If BRs Is Nothing Then Exit Sub

FPS = 0
PSA = 0
Ratio = 0

For Each br In BRs
  Code = UCase$(Trim$(br.Code))
  If Code = "FPS" Then
    FPS = Val(br.Result)
    FPSDate = br.Rundate
    FPSTime = br.RunTime
  ElseIf Code = "PSA" Then
    PSA = Val(br.Result)
  ElseIf Code = "FPR" Then
    Ratio = Val(br.Result)
  End If
Next

If (FPS * PSA) <> 0 And Ratio = 0 Then
  Ratio = FPS / PSA
  Set br = New BIEResult
  br.SampleID = txtSampleID
  br.Code = "FPR"
  br.Rundate = FPSDate
  br.RunTime = FPSTime
  br.Result = Format$(Ratio, "#0.00")
  br.Units = ""
  br.Printed = 0
  br.Valid = 0
  BRs.Add br
  BRs.Save "Bio", BRs
End If

Exit Sub

CheckCalcPSA_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /CheckCalcPSA ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub CheckCC()

Dim SQL As String
Dim tb As Recordset

On Error GoTo ehCCC

cmdCopyTo.Caption = "cc"
cmdCopyTo.Font.Bold = False
cmdCopyTo.BackColor = &H8000000F

If Trim$(txtSampleID) = "" Then Exit Sub
  
SQL = "Select * from SendCopyTo where " & _
      "SampleID = '" & Val(txtSampleID) & "'"
Set tb = New Recordset
RecOpenServer 0, tb, SQL
If Not tb.EOF Then
  cmdCopyTo.Caption = "++ cc ++"
  cmdCopyTo.Font.Bold = True
  cmdCopyTo.BackColor = &H8080FF
End If

Exit Sub

ehCCC:
Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Screen.MousePointer = 0
LogError "frmeditall/CheckCC:" & Str(er) & ":" & ers
Exit Sub

End Sub

Private Sub CheckCholHDL(ByVal BRs As BIEResults)

Dim br As BIEResult
Dim Chol As Single
Dim HDL As Single
Dim CholTime As String
Dim CholDate As String
Dim Ratio As Single
Dim Code As String
Dim BRResNew As New BIEResults

On Error GoTo CheckCholHDL_Error

If BRs Is Nothing Then Exit Sub

Chol = 0
HDL = 0
Ratio = 0

For Each br In BRs
  Code = UCase$(Trim$(br.Code))
  If Code = SysOptBioCodeForChol(0) Then
    Chol = Val(br.Result)
    CholDate = br.Rundate
    CholTime = br.RunTime
  ElseIf Code = SysOptBioCodeForHDL(0) Then
    HDL = Val(br.Result)
  ElseIf Code = SysOptBioCodeForCholHDLRatio(0) Then
    Ratio = Val(br.Result)
  End If
Next

If (Chol * HDL) <> 0 And Ratio = 0 Then
  Ratio = Chol / HDL
  Set br = New BIEResult
  br.SampleID = txtSampleID
  br.Code = SysOptBioCodeForCholHDLRatio(0)
  br.ShortName = "C/H R"
  br.Rundate = CholDate
  br.RunTime = CholTime
  br.Result = Format$(Ratio, "#0.00")
  br.SampleType = "S"
  br.Units = "Ratio"
  br.Valid = 0
  br.Printed = 0
'  BR.Authorised = 0
  br.Printformat = 1

  BRs.Add br
  BRResNew.Add br
  BRResNew.Save "bio", BRResNew
End If

Exit Sub

CheckCholHDL_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /CheckCholHDL ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Function CheckDemographics(ByVal TrialID As String) _
   As String

Dim sn As New Recordset
Dim SQL As String
Dim n As Long
Dim pName(1 To 4) As String
Dim pAddress(1 To 4) As String
Dim pDoB(1 To 4) As String
Dim IDFound(1 To 4) As Boolean
Dim Found As Long
Dim F As Form

On Error GoTo CheckDemographics_Error

If TrialID = "" Then Exit Function

Set sn = New Recordset
With sn
   Found = 0
   For n = 1 To 4
      IDFound(n) = False
      SQL = "SELECT * from patientifs WHERE " & _
         Choose(n, "CHART", "NOPAS", "MRN", "AandE") & " = '" & TrialID & "'"
      RecOpenServer 0, sn, SQL
      If Not .EOF Then
        Do While Not sn.EOF
         IDFound(n) = True
         Found = Found + 1
         pName(n) = initial2upper(!PatName)
         If Not IsNull(!DoB) Then pDoB(n) = Format(!DoB, "dd/MM/yyyy")
         pAddress(n) = initial2upper(!Address0 & "") & " " & initial2upper(!Address1 & "")
         sn.MoveNext
           Loop
      End If
      .Close
   Next
End With

If Found = 0 Then
   CheckDemographics = ""
ElseIf Found = 1 Then
   For n = 1 To 4
      If IDFound(n) Then
         CheckDemographics = Choose(n, "CHART", "NOPAS", "MRN", "AandE")
         Exit For
      End If
   Next
Else
   Set F = New frmDemogCheck
   With F
      For n = 1 To 4
         If IDFound(n) Then
            .bSelect(n).Visible = True
            .lName(n) = initial2upper(pName(n))
            .lAddress(n) = initial2upper(pAddress(n))
            .lDoB(n) = pDoB(n)
         End If
      Next
      .Show 1
      CheckDemographics = .IDType
   End With
   Unload F
   Set F = Nothing
End If


Exit Function

CheckDemographics_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /CheckDemographics ")
  Case 1:    End
  Case 2:    Exit Function
  Case 3:    Resume Next
End Select


End Function

Private Sub CheckDepartments()

On Error GoTo CheckDepartments_Error

If SysOptDeptHaem(0) = True Then
  If AreHaemResultsPresent(txtSampleID) = 1 Then
    sstabAll.TabCaption(1) = "<<Haematology>>"
  End If
End If

If SysOptDeptBio(0) = True Then
  If AreBioResultsPresent(txtSampleID) = 1 Then
    sstabAll.TabCaption(2) = "<<Biochemistry>>"
  End If
End If

If SysOptDeptCoag(0) = True Then
  If AreCoagResultsPresent(txtSampleID) = 1 Then
    sstabAll.TabCaption(3) = "<<Coagulation>>"
  End If
End If

If SysOptDeptEnd(0) = True Then
  If AreEndResultsPresent(txtSampleID) = 1 Then
    sstabAll.TabCaption(4) = "<<Endocrinology>>"
  End If
End If

If SysOptDeptBga(0) = True Then
  If AreBgaResultsPresent(txtSampleID) = 1 Then
    sstabAll.TabCaption(5) = "<<Blood Gas>>"
  End If
End If

If SysOptDeptImm(0) = True Then
  If AreImmResultsPresent(txtSampleID) = 1 Then
    sstabAll.TabCaption(6) = "<<Immunology>>"
  End If
End If

If SysOptDeptExt(0) = True Then
  If AreExtResultsPresent(txtSampleID) = 1 Then
    sstabAll.TabCaption(7) = "<<Externals>>"
  End If
End If


Exit Sub

CheckDepartments_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /CheckDepartments ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub CheckIfPhoned()

If CheckPhoneLog(txtSampleID) Then
  cmdPhone.BackColor = vbYellow
  cmdPhone.Caption = "Results Phoned"
  cmdPhone.ToolTipText = "Results Phoned"
Else
  cmdPhone.BackColor = &H8000000F
  cmdPhone.Caption = "Phone Results"
  cmdPhone.ToolTipText = "Phone Results"
End If

End Sub

Private Sub chkBad_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim SQL As String

pBar = 0

If chkBad.Value = 1 Then
  'Code added 22/08/05
  If iMsg("Do you wish all outstanding requests Deleted!", vbYesNo) = vbYes Then
    SQL = "DELETE from haemrequests WHERE sampleID = '" & txtSampleID & "'"
    Cnxn(0).Execute SQL
  End If
End If

cmdSaveHaem.Enabled = True: cmdHSaveH.Enabled = True

End Sub

Private Sub chkMalaria_Click()

cmdSaveHaem.Enabled = True: cmdHSaveH.Enabled = True

End Sub

Private Sub chkPgp_Click()

cmdSaveDemographics.Enabled = True
cmdSaveInc.Enabled = True

End Sub

Private Sub chkSickledex_Click()

cmdSaveHaem.Enabled = True: cmdHSaveH.Enabled = True

End Sub

Private Sub cIAdd_Click(Index As Integer)
Dim SampleType As String
Dim tb As New Recordset
Dim SQL As String



On Error GoTo cIAdd_Click_Error

If Index = 0 Then

  cIUnits(0).Enabled = True
  
  SampleType = ListCodeFor("ST", cISampleType(Index))
  
  SQL = "SELECT * from endtestdefinitions WHERE code = '" & eCodeForShortName(cIAdd(0)) & "'"
  Set tb = New Recordset
  RecOpenServer 0, tb, SQL
  If Not tb.EOF Then
    cIUnits(0) = Trim(tb!Units & "")
  Else
    cIUnits(0) = ""
  End If
  
  cIUnits(0).Enabled = False
ElseIf Index = 1 Then
  cIUnits(1).Enabled = True
  
  
  SQL = "SELECT * from Immtestdefinitions WHERE code = '" & ICodeForShortName(cIAdd(1)) & "'"
  Set tb = New Recordset
  RecOpenServer 0, tb, SQL
  If Not tb.EOF Then
    cIUnits(1) = Trim(tb!Units) & ""
  Else
    cIUnits(1) = ""
  End If
  
  
  cIUnits(1).Enabled = False
ElseIf Index = 2 Then
  cIUnits(2).Enabled = True
  
  
  SQL = "SELECT * from bgatestdefinitions WHERE code = '" & BgaCodeForShortName(cIAdd(2)) & "'"
  Set tb = New Recordset
  RecOpenServer 0, tb, SQL
  If Not tb.EOF Then
    cIUnits(2) = Trim(tb!Units) & ""
  Else
    cIUnits(2) = ""
  End If
  
  
  cIUnits(2).Enabled = False
End If

Exit Sub

cIAdd_Click_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /cIAdd_Click ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub cIAdd_KeyPress(Index As Integer, KeyAscii As Integer)

KeyAscii = 0

End Sub

Private Sub cISampleType_Change(Index As Integer)

If Index = 0 Then
  FillcEAdd
ElseIf Index = 1 Then
  FillcIAdd
ElseIf Index = 2 Then
  FillcbAdd
End If

End Sub

Private Sub cISampleType_Click(Index As Integer)

If Index = 0 Then
  FillcEAdd
ElseIf Index = 1 Then
  FillcIAdd
ElseIf Index = 2 Then
  FillcbAdd
End If

End Sub

Private Sub ClearCoagulation()

On Error GoTo ClearCoagulation_Error

cParameter = ""
cCunits.ListIndex = -1
tResult = ""
'txtCoagComment = ""
tWarfarin = ""
bViewCoagRepeat.Visible = False
lCDate = ""

Exit Sub

ClearCoagulation_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /ClearCoagulation ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Public Sub ClearDemographics()
Dim n As Long
Dim Temp As String

  lblUrgent.Visible = False
  mNewRecord = True
  dtRunDate = Format$(Now, "dd/mm/yyyy")
  lblRundate = dtRunDate
  dtSampleDate = Format$(Now, "dd/mm/yyyy")
  lblSampledate = dtSampleDate
  dtRecDate = Format$(Now, "dd/mm/yyyy")
  If SysOptDemoVal(0) Then cmdDemoVal.Caption = "&Validate"
  txtChart = ""
  txtName = ""
  taddress(0) = ""
  taddress(1) = ""
  txtNOPAS = ""
  txtAandE = ""
  lblNOPAS(1) = ""
  StatusBar1.Panels(4).Text = ""
  txtSex = ""
  txtDoB = ""
  txtAge = ""
  lDoB = ""
  lAge = ""
  lSex = ""
  cmbWard = "GP"
  cmbClinician = ""
  cmbGP = ""
  cClDetails = ""
  txtDemographicComment = ""
  tSampleTime.Mask = ""
  tSampleTime.Text = ""
  tSampleTime.Mask = "##:##"
  tRecTime.Mask = ""
  tRecTime.Text = ""
  tRecTime.Mask = "##:##"
  lblChartNumber.Caption = Hospname(0) & " Chart #"
  lblChartNumber.BackColor = &H8000000F
  lblChartNumber.ForeColor = vbBlack
  cCat(0).ListIndex = -1
  cCat(1).ListIndex = -1
    
  If cmbHospital = "" Then
    For n = 0 To cmbHospital.ListCount - 1
      If UCase(cmbHospital.List(n)) = Hospname(0) Then
        cmbHospital.ListIndex = n
      End If
    Next
  End If
  Set_Demo True
  chkPgp.Value = 0

End Sub

Private Sub ClearEndFlags()

Ih(0) = 0
Iis(0) = 0
Il(0) = 0
Io(0) = 0
Ig(0) = 0
Ij(0) = 0

End Sub

Private Sub ClearExt()

ClearFGrid grdExt
grdExt.Visible = True

End Sub

Private Sub ClearHaematologyResults()

Dim n As Long

On Error GoTo ClearHaematologyResults_Error

'ClearRbcGrid
'ClearHaemDiffGrid
'
''HGB_Click
'
'lWIC = ""
'lWOC = ""
'
'tWBC = ""
'tWBC.BackColor = &HFFFFFF
'tWBC.ForeColor = &H0&
'
'tPlt = ""
'tPlt.BackColor = &HFFFFFF
'tPlt.ForeColor = &H0&
'
'tMPV = ""
'tMPV.BackColor = &HFFFFFF
'tMPV.ForeColor = &H0&
'
'lblAnalyser = "Analyser : "
'txtLI = ""
'txtMPXI = ""
'
'pdelta.Cls
'lHDate = ""
'cESR = 0
'cRetics = 0
'cMonospot = 0
'cRA = 0
'cASot = 0
'chkMalaria = 0
'chkSickledex = 0
'chkBad = 0
'
'tESR = ""
'tESR.BackColor = &HFFFFFF
'tESR.ForeColor = &H0&
'
'txtEsr1 = ""
'txtEsr1.BackColor = &HFFFFFF
'txtEsr1.ForeColor = &H0&
'
'tRetA = ""
'tRetP = ""
'tRetA.BackColor = vbWhite
'tRetA.ForeColor = 1
'tMonospot = ""
'cFilm = 0
'tRa = ""
'tASOt = ""
'lblMalaria = ""
'lblSickledex = ""
'
''cCoag = 0
'
'tWarfarin = ""
'
'For n = 0 To 5
'  ipflag(n).Visible = False
'Next

Exit Sub

ClearHaematologyResults_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /ClearHaematologyResults ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub ClearHaemDiffGrid()

On Error GoTo ClearHaemDiffGridError

'With grdH
'
'  .Rows = 2
'  .AddItem ""
'  .RemoveItem 1
'
'  .AddItem vbTab & vbTab & "Neut"
'  .AddItem vbTab & vbTab & "Lymph"
'  .AddItem vbTab & vbTab & "Mono"
'  .AddItem vbTab & vbTab & "Eos"
'  .AddItem vbTab & vbTab & "Bas"
'  .AddItem vbTab & vbTab & "Luc"
'
'  .RemoveItem 1
'
'End With

Exit Sub

ClearHaemDiffGridError:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll/ClearHaemDiffGrid ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub ClearHgb()
Dim n As Long

On Error GoTo HGB_Click_Error

pBar = 0

bcleardiff_click

  gRbc.TextMatrix(1, 1) = ""
  gRbc.Row = 1
  gRbc.Col = 1
  gRbc.CellBackColor = vbWhite
  gRbc.CellForeColor = 1
  gRbc.Col = 2
  gRbc.TextMatrix(1, 2) = ""
  gRbc.CellBackColor = &H8000000F
  gRbc.CellForeColor = 1

For n = 3 To gRbc.Rows - 1
  gRbc.Row = n
  gRbc.Col = 1
  gRbc = ""
  gRbc.CellBackColor = vbWhite
  gRbc.CellForeColor = 1
  gRbc.Col = 2
  gRbc = ""
  gRbc.CellBackColor = &H8000000F
  gRbc.CellForeColor = 1
Next


tWBC = ""
tWBC.BackColor = &HFFFFFF
tWBC.ForeColor = 1
tPlt = ""
tPlt.BackColor = &HFFFFFF
tPlt.ForeColor = 1
txtMPXI = ""
txtMPXI.BackColor = &HFFFFFF
txtMPXI.ForeColor = 1
lWIC = ""
lWOC = ""
tMPV = ""
tMPV.BackColor = &HFFFFFF
tMPV.ForeColor = 1
txtLI = ""
txtLI.BackColor = &HFFFFFF
txtLI.ForeColor = 1

cmdSaveHaem.Enabled = True: cmdHSaveH.Enabled = True
bValidateHaem.Enabled = True

Exit Sub

HGB_Click_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /HGB_Click ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub ClearImmFlags()

Ih(1) = 0
Iis(1) = 0
Il(1) = 0
Io(1) = 0
Ig(1) = 0
Ij(1) = 0

End Sub

Private Sub ClearOutstandingBio()

On Error GoTo ClearOutstandingBio_Error

With grdOutstanding
  .Rows = 2
  .AddItem ""
  .RemoveItem 1
End With

Exit Sub

ClearOutstandingBio_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /ClearOutstandingBio ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub ClearOutstandingEnd()

With grdOutstandings(0)
  .Rows = 2
  .AddItem ""
  .RemoveItem 1
End With

End Sub

Private Sub ClearOutstandingImm()

With grdOutstandings(1)
  .Rows = 2
  .AddItem ""
  .RemoveItem 1
End With

End Sub

Private Sub ClearRbcGrid()

Dim n As Long

On Error GoTo ClearRbcGridError

'With gRbc
'
'  .Rows = 2
'  .AddItem ""
'  .RemoveItem 1
'
'  .AddItem "RBC"
'  .AddItem "Hgb"
'  .AddItem "HCT"
'  .AddItem "MCV"
'  .AddItem "HDW"
'  .AddItem "MCH"
'  .AddItem "MCHC"
'  .AddItem "CHCM"
'  .AddItem "RDW"
'  .AddItem "NRBC%"
'  .AddItem "HYPO%"
'
'  .RemoveItem 1
'
'  For n = 1 To .Rows - 1
'    .Row = n
'    .Col = 0
'    .CellFontBold = True
'    .CellBackColor = &H8000000F
'    .CellForeColor = &HC0&
'    .Col = 1
'    .CellFontBold = True
'    .CellBackColor = &H80000005
'    .CellForeColor = vbBlack
'    .Col = 2
'    .CellFontBold = True
'    .CellBackColor = &H80000005
'    .CellForeColor = vbBlack
'  Next
'End With

Exit Sub

ClearRbcGridError:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /ClearRbc ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub cmbClinician_Change()

SetWardClinGp

End Sub

Private Sub cmbClinician_Click()

cmdSaveDemographics.Enabled = True
cmdSaveInc.Enabled = True

End Sub

Private Sub cmbClinician_KeyPress(KeyAscii As Integer)

cmdSaveDemographics.Enabled = True
cmdSaveInc.Enabled = True

End Sub

Private Sub cmbClinician_LostFocus()

pBar = 0
cmbClinician = QueryKnown("Clin", cmbClinician, UCase(cmbHospital))

End Sub

Private Sub cmbGP_Change()

SetWardClinGp

cmbWard = "GP"

End Sub

Private Sub cmbGP_Click()

On Error GoTo cmbGP_Click_Error

pBar = 0

lAddWardGP = Trim$(taddress(0)) & " : GP : " & cmbGP
cmbWard = "GP"
cmdSaveDemographics.Enabled = True
cmdSaveInc.Enabled = True

Exit Sub

cmbGP_Click_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /cmbGP_Click ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub cmbGP_KeyPress(KeyAscii As Integer)

cmdSaveDemographics.Enabled = True
cmdSaveInc.Enabled = True

End Sub

Private Sub cmbGP_LostFocus()

cmbGP = QueryKnown("GP", cmbGP, cmbHospital)

End Sub

Private Sub cmbHospital_Click()


On Error GoTo cmbHospital_Click_Error





FillGPsClinWard Me, cmbHospital

cmdSaveDemographics.Enabled = True
cmdSaveInc.Enabled = True

Exit Sub

cmbHospital_Click_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /cmbHospital_Click ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub cmbHospital_LostFocus()
Dim n As Long

For n = 0 To cmbHospital.ListCount
  If UCase(cmbHospital) = UCase(Left(cmbHospital.List(n), Len(cmbHospital))) Then
    cmbHospital.ListIndex = n
  End If
Next



End Sub

Private Sub cmbWard_Change()

SetWardClinGp

End Sub

Private Sub cmbWard_Click()

lAddWardGP = Trim$(taddress(0)) & " : " & cmbWard & " : " & cmbGP
If Hospname(0) = "STJOHNS" And cmbClinician <> "" Then
  lAddWardGP = Trim$(taddress(0)) & " : " & cmbWard & " : " & cmbClinician
End If

cmdSaveDemographics.Enabled = True
cmdSaveInc.Enabled = True

End Sub

Private Sub cmbWard_KeyPress(KeyAscii As Integer)

cmdSaveDemographics.Enabled = True
cmdSaveInc.Enabled = True

End Sub

Private Sub cmbWard_LostFocus()

Dim Found As Boolean
Dim tb As New Recordset
Dim SQL As String

On Error GoTo cmbWard_LostFocus_Error

If Trim$(cmbWard) = "" Then
  cmbWard = "GP"
  Exit Sub
End If

Found = False

SQL = "SELECT * from wards WHERE text = '" & AddTicks(cmbWard) & "' or code = '" & AddTicks(cmbWard) & "' and hospitalcode = '" & ListCodeFor("HO", cmbHospital) & "'"
Set tb = New Recordset
RecOpenServer 0, tb, SQL
  If Not tb.EOF Then
    cmbWard = Trim(tb!Text)
    Found = True
  End If

If Not Found Then
  cmbWard = "GP"
End If

Exit Sub

cmbWard_LostFocus_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /cmbWard_LostFocus ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub cmdCopyTo_Click()
Dim s As String

s = cmbWard & " " & cmbClinician & " " & cmbGP
s = Trim$(s)


frmCopyTo.lblOriginal = s
frmCopyTo.lblSampleID = txtSampleID
frmCopyTo.Show 1

CheckCC


End Sub

Private Sub cmdDel_Click()

Dim Str As String


On Error GoTo cmdDel_Click_Error

If grdExt.TextMatrix(grdExt.Row, 0) = "Test Number" Then Exit Sub

If grdExt.TextMatrix(grdExt.RowSel, 0) = "" Then Exit Sub


Str = "  Test Name : " & grdExt.TextMatrix(grdExt.Row, 1) & vbCrLf & _
    "Test Number : " & grdExt.TextMatrix(grdExt.Row, 0) & vbCrLf & _
    "DELETE this test?"
If iMsg(Str, vbQuestion + vbYesNo, "Confirm Deletion") = vbYes Then
  Str = "DELETE from extresults WHERE " & _
      "sampleid = '" & txtSampleID & "' " & _
      "and Analyte = " & grdExt.TextMatrix(grdExt.Row, 0)
  Cnxn(0).Execute Str
  LoadExt
End If


Exit Sub

cmdDel_Click_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /cmdDel_Click ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub cmdDemoVal_Click()
Dim SQL As String
Dim tb As New Recordset

If cmdDemoVal.Caption = "&Validate" Then
  If cmdSaveDemographics.Enabled Then SaveDemographics
  SQL = "SELECT * from demographics WHERE sampleid = '" & txtSampleID & "'"
  Set tb = New Recordset
  RecOpenServer 0, tb, SQL
  If Not tb.EOF Then
    Archive 0, tb, "arcdemographics", txtSampleID
    SQL = "UPDATE demographics set valid = 1, " & _
          "username = '" & UserName & "' WHERE " & _
          "sampleid = '" & txtSampleID & "'"
    Cnxn(0).Execute SQL
    Set_Demo False
    cmdDemoVal.Caption = "VALID"
    cmdSaveDemographics.Enabled = False
    cmdSaveInc.Enabled = False
  End If
Else
  If UCase(iBOX("Enter password to unValidate ?", , , True)) = UserPass Then
    SQL = "SELECT * from demographics WHERE sampleid = '" & txtSampleID & "'"
    Set tb = New Recordset
    RecOpenServer 0, tb, SQL
    If Not tb.EOF Then
      Archive 0, tb, "arcdemographics", txtSampleID
      SQL = "UPDATE demographics set valid = 0, " & _
            "username = '" & UserName & "' WHERE " & _
            "sampleid = '" & txtSampleID & "'"
      Cnxn(0).Execute SQL
      Set_Demo True
      cmdDemoVal.Caption = "&Validate"
    End If
  End If
End If
End Sub

Private Sub cmdGetBio_Click()

LoadBiochemistry
frmBio2Imm.Show 1
LoadImmunology

End Sub

Private Sub cmdHSaveH_Click()

On Error GoTo cmdSaveHaem_Click_Error

pBar = 0

txtSampleID = Format(Val(txtSampleID))
If Val(txtSampleID) = 0 Then Exit Sub

'Added 15/Jul/2004

If bValidateHaem.Caption = "&Validate" Then
  SaveHaematology 0
Else
  SaveHaematology 1
End If

SaveComments
UPDATEMRU

LoadAllDetails

cmdSaveHaem.Enabled = False
cmdHSaveH.Enabled = False
Exit Sub

cmdSaveHaem_Click_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /cmdHSaveH_Click ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select

End Sub

Private Sub cmdIAdd_Click(Index As Integer)
Dim SQL As String
Dim s As String

Dim n As Long

On Error GoTo cmdIAdd_Click_Error

pBar = 0

If Index = 0 Then

  If cIAdd(0).Text = "" Then Exit Sub
  If Trim$(tINewValue(0)) = "" Then Exit Sub
  If Trim$(txtSampleID) = "" Then Exit Sub
    
  For n = 1 To gImm(0).Rows - 1
    If cIAdd(0) = gImm(0).TextMatrix(n, 0) Then
      iMsg "Test already Exists. Please delete before adding!"
      Exit Sub
    End If
  Next
  s = Check_End(cIAdd(0).Text, cIUnits(0), cISampleType(0))
  If s <> "" Then
    iMsg s & " is incorrect!"
    Exit Sub
  End If
    
  SQL = "INSERT into endResults " & _
        "(RunDate, SampleID, Code, Result, RunTime, Units, SampleType, Valid, Printed) VALUES " & _
        "('" & Format$(dtRunDate, "dd/mmm/yyyy") & "', " & _
        "'" & txtSampleID & "', " & _
        "'" & eCodeForShortName(cIAdd(0).Text) & "', " & _
        "'" & tINewValue(0) & "', " & _
        "'" & Format$(Now, "dd/mmm/yyyy hh:mm") & "', " & _
        "'" & cIUnits(0) & "', " & _
        "'" & ListCodeFor("ST", cISampleType(0)) & "', 0, 0);"
        
  Cnxn(0).Execute SQL
  
  
 
  LoadEndocrinology
  
  cIAdd(0) = ""
  tINewValue(0) = ""
  cIUnits(0) = ""
  

ElseIf Index = 1 Then
      

  If cIAdd(1).Text = "" Then Exit Sub
  If Trim$(tINewValue(1)) = "" Then Exit Sub
  If Trim$(txtSampleID) = "" Then Exit Sub
    
  For n = 1 To gImm(1).Rows - 1
    If gImm(1).TextMatrix(n, 0) = cIAdd(1) Then
        iMsg "Test already in List!"
        Exit Sub
    End If
  Next
    
  s = Check_Imm(cIAdd(1).Text, cIUnits(1), cISampleType(1))
  If s <> "" Then
    iMsg s & " is incorrect!"
    Exit Sub
  End If
  
  SQL = "INSERT into ImmResults " & _
        "(RunDate, SampleID, Code, Result, RunTime, Units, SampleType, Valid, Printed) VALUES " & _
        "('" & Format$(dtRunDate, "dd/mmm/yyyy") & "', " & _
        "'" & txtSampleID & "', " & _
        "'" & ICodeForShortName(cIAdd(1).Text) & "', " & _
        "'" & tINewValue(1) & "', " & _
        "'" & Format$(Now, "dd/mmm/yyyy hh:mm") & "', " & _
        "'" & cIUnits(1) & "', " & _
        "'" & ListCodeFor("ST", cISampleType(1)) & "', 0, 0);"
        
  Cnxn(0).Execute SQL
  
  LoadImmunology
  
  cIAdd(1) = ""
  tINewValue(1) = ""
  cIUnits(1) = ""

ElseIf Index = 2 Then
  
  If cIAdd(2).Text = "" Then Exit Sub
  If Trim$(tINewValue(2)) = "" Then Exit Sub
  If Trim$(txtSampleID) = "" Then Exit Sub
    
  For n = 1 To gBga.Rows - 1
    If gBga.TextMatrix(n, 0) = cIAdd(2) Then
        iMsg "Test already in List!"
        Exit Sub
    End If
  Next
    
  SQL = "INSERT into bgaResults " & _
        "(RunDate, SampleID, Code, Result, RunTime, Units, SampleType, Valid, Printed) VALUES " & _
        "('" & Format$(dtRunDate, "dd/mmm/yyyy") & "', " & _
        "'" & txtSampleID & "', " & _
        "'" & BgaCodeForShortName(cIAdd(2).Text) & "', " & _
        "'" & tINewValue(2) & "', " & _
        "'" & Format$(Now, "dd/mmm/yyyy hh:mm") & "', " & _
        "'" & cIUnits(2) & "', " & _
        "'S', 0, 0);"
        
  Cnxn(0).Execute SQL
  
  
  SQL = "UPDATE Demographics set Forbga = 1 WHERE sampleID = '" & txtSampleID & "'"
  Cnxn(0).Execute SQL
  
  LoadBloodGas
  
  cIAdd(2) = ""
  tINewValue(2) = ""
  cIUnits(2) = ""

End If


Exit Sub

cmdIAdd_Click_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /cmdIAdd_Click ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub cmdIremoveduplicates_Click(Index As Integer)
Dim tb As New Recordset
Dim SQL As String
Dim Y As Long
Dim Code As String

On Error GoTo cmdIremoveduplicates_Click_Error

pBar = 0

If Index = 0 Then

  If gImm(0).Rows < 3 Then Exit Sub
  
  Screen.MousePointer = 11
  
  For Y = 1 To gImm(0).Rows - 1
    Code = eCodeForShortName(gImm(0).TextMatrix(Y, 0))
    SQL = "SELECT * from Endresults WHERE " & _
          "sampleid = '" & txtSampleID & "' " & _
          "and code = '" & Code & "' order by runtime asc"
    Set tb = New Recordset
    RecOpenClient 0, tb, SQL
    Do While tb.RecordCount > 1
      tb.DELETE
      tb.MoveNext
    Loop
  Next
  
  LoadEndocrinology
Else

  If gImm(1).Rows < 3 Then Exit Sub
  
  Screen.MousePointer = 11
  
  For Y = 1 To gImm(1).Rows - 1
    Code = ICodeForShortName(gImm(1).TextMatrix(Y, 0))
    SQL = "SELECT * from Immresults WHERE " & _
          "sampleid = '" & txtSampleID & "' " & _
          "and code = '" & Code & "' order by runtime asc"
    Set tb = New Recordset
    RecOpenClient 0, tb, SQL
    Do While tb.RecordCount > 1
      tb.DELETE
      tb.MoveNext
    Loop
  Next
  
  LoadImmunology
End If
Screen.MousePointer = 0

Exit Sub

cmdIremoveduplicates_Click_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /cmdIremoveduplicates_Click ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select

End Sub

Private Sub cmdPhone_Click()

With frmPhoneLog
  .SampleID = txtSampleID
  If cmbGP <> "" Then
    .GP = cmbGP
    .WardOrGP = "GP"
  Else
    .GP = cmbWard
    .WardOrGP = "Ward"
  End If
  .Show 1
End With

CheckIfPhoned

End Sub

Private Sub cmdPrint_Click()

Dim tb As New Recordset
Dim SQL As String

On Error GoTo cmdPrint_Click_Error

pBar = 0



txtSampleID = Format(Val(txtSampleID))
If Val(txtSampleID) = 0 Then Exit Sub

If Trim$(txtSex) = "" Then
  If iMsg("Sex not entered." & vbCrLf & "Do you want to enter sex now?", vbQuestion + vbYesNo) = vbYes Then
    Exit Sub
  End If
End If

If Trim$(txtSampleID) = "" Then
  iMsg "Must have Lab Number.", vbCritical
  Exit Sub
End If


If Len(cmbWard) = 0 Then
  iMsg "Must have Ward entry.", vbCritical
  Exit Sub
End If

If Trim$(cmbWard) = "GP" Then
  If Len(cmbGP) = 0 Then
    iMsg "Must have Ward or GP entry.", vbCritical
    Exit Sub
  End If
End If


  If SysOptDemoVal(0) And cmdDemoVal.Caption = "&Validate" Then
    If iMsg("Do you wish to validate demographics !", vbYesNo) = vbNo Then
      Exit Sub
    Else
      cmdDemoVal_Click
    End If
  End If

SaveDemographics

If sstabAll.Tab <> 0 Then
  LogTimeOfPrinting txtSampleID, Choose(sstabAll.Tab, "H", "B", "C", "E", "Q", "I", "")
  SQL = "SELECT * from PrintPending WHERE " & _
        "Department = '" & Choose(sstabAll.Tab, "H", "B", "C", "E", "Q", "I", "") & "' " & _
        "and SampleID = '" & txtSampleID & "'"
  Set tb = New Recordset
  RecOpenClient 0, tb, SQL
  If tb.EOF Then
    tb.AddNew
  End If
  tb!SampleID = txtSampleID
  tb!Department = Choose(sstabAll.Tab, "H", "B", "C", "E", "Q", "I", "")
  If SysOptRealImm(0) And tb!Department = "I" Then tb!Department = "J"
  tb!Initiator = UserName
  tb!Ward = cmbWard
  tb!Clinician = cmbClinician
  tb!GP = cmbGP
  tb!UsePrinter = pPrintToPrinter
  tb!ptime = Now
  tb.Update
End If

If sstabAll.Tab = 1 Then
  If lblHaemValid.Visible = False Then
    SaveHaematology 1
  End If
  SQL = "SELECT * from HaemResults WHERE " & _
        "SampleID = '" & txtSampleID & "'"
  Set tb = New Recordset
  RecOpenClient 0, tb, SQL
  If Not tb.EOF Then
    tb!Printed = 1
    tb.Update
  End If
ElseIf sstabAll.Tab = 2 And bValidateBio.Caption = "&Validate" Then
  SQL = "UPDATE bioResults " & _
        "Set valid = 1 , operator = '" & UserCode & "' WHERE " & _
        "SampleID = '" & txtSampleID & "' and valid <> 1"
  Cnxn(0).Execute SQL
ElseIf sstabAll.Tab = 3 Then
  If cmdValidateCoag.Caption = "&Validate" Then
    SQL = "UPDATE CoagResults " & _
          "Set valid = 1 , username = '" & UserCode & "' WHERE " & _
          "SampleID = " & txtSampleID & " and valid <> 1"
    Cnxn(0).Execute SQL
  End If
  SQL = "UPDATE CoagResults " & _
        "Set Printed = 0 WHERE " & _
        "SampleID = '" & txtSampleID & "'"
  Cnxn(0).Execute SQL
ElseIf sstabAll.Tab = 4 And bValidateImm(0).Caption = "&Validate" Then
  SQL = "UPDATE endResults " & _
        "Set valid = 1 , operator = '" & UserCode & "' WHERE " & _
        "SampleID = '" & txtSampleID & "' and valid <> 1"
  Cnxn(0).Execute SQL
ElseIf sstabAll.Tab = 5 And cmdValBG.Caption = "&Validate" Then
  SQL = "UPDATE bgaResults " & _
        "Set valid = 1 , operator = '" & UserCode & "' WHERE " & _
        "SampleID = '" & txtSampleID & "' and valid <> 1"
  Cnxn(0).Execute SQL
ElseIf sstabAll.Tab = 6 And bValidateImm(1).Caption = "&Validate" Then
  SQL = "UPDATE immResults " & _
        "Set valid = 1 , operator = '" & UserCode & "' WHERE " & _
        "SampleID = '" & txtSampleID & "' and (valid <> 1 or valid is null)"
  Cnxn(0).Execute SQL
End If

txtSampleID = Format$(Val(txtSampleID) + 1)
LoadAllDetails

Exit Sub

cmdPrint_Click_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /cmdPrint_Click ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub cmdPrintAll_Click()

Dim tb As New Recordset
Dim SQL As String

On Error GoTo cmdPrintAll_Click_Error

pBar = 0

txtSampleID = Format(Val(txtSampleID))
If Val(txtSampleID) = 0 Then Exit Sub

If Trim$(txtSex) = "" Then
  If iMsg("Sex not entered." & vbCrLf & "Do you want to enter sex now?", vbQuestion + vbYesNo) = vbYes Then
    Exit Sub
  End If
End If

If Trim$(txtSampleID) = "" Then
  iMsg "Must have Lab Number.", vbCritical
  Exit Sub
End If

If Trim$(cmbWard) = "" Then
  iMsg "Must have Ward entry.", vbCritical
  Exit Sub
End If

If Trim$(cmbWard) = "GP" Then
  If Trim$(cmbGP) = "" Then
    iMsg "Must have Ward or GP entry.", vbCritical
    Exit Sub
  End If
End If

SaveDemographics

If sstabAll.Tab <> 0 Then
  SQL = "SELECT * from PrintPending WHERE " & _
        "Department = 'D' " & _
        "and SampleID = '" & txtSampleID & "'"
  Set tb = New Recordset
  RecOpenClient 0, tb, SQL
  If tb.EOF Then
    tb.AddNew
  End If
  tb!SampleID = txtSampleID
  tb!Ward = cmbWard
  tb!Clinician = cmbClinician
  tb!GP = cmbGP
  tb!Department = "D"
  tb!Initiator = UserName
  tb!UsePrinter = pPrintToPrinter
  tb.Update
End If

SaveCoag 1
SQL = "UPDATE CoagResults " & _
      "Set Valid = 1, Printed = 1 WHERE " & _
      "SampleID = '" & txtSampleID & "'"
Cnxn(0).Execute SQL

txtSampleID = Format$(Val(txtSampleID) + 1)
LoadAllDetails

Exit Sub

cmdPrintAll_Click_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /cmdPrintAll_Click ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub cmdPrintesr_Click()

Dim SQL As String
Dim tb As New Recordset



On Error GoTo cmdPrintesr_Click_Error

txtSampleID = Format(Val(txtSampleID))
If Val(txtSampleID) = 0 Then Exit Sub

pBar = 0

If Trim$(txtSex) = "" Then
  If iMsg("Sex not entered." & vbCrLf & "Do you want to enter sex now?", vbQuestion + vbYesNo) = vbYes Then
    Exit Sub
  End If
End If

If Trim$(txtSampleID) = "" Then
  iMsg "Must have Lab Number.", vbCritical
  Exit Sub
End If

If Trim$(cmbWard) = "" Then
  iMsg "Must have Ward entry.", vbCritical
  Exit Sub
End If

If Trim$(cmbWard) = "GP" Then
  If Trim$(cmbGP) = "" Then
    iMsg "Must have GP entry.", vbCritical
    Exit Sub
  End If
End If

SaveDemographics
SaveHaematology 1

PrintResultESRWin txtSampleID

SQL = "SELECT * from HaemResults WHERE " & _
      "SampleID = '" & txtSampleID & "'"
Set tb = New Recordset
RecOpenClient 0, tb, SQL

If Not tb.EOF Then
  tb!Printed = 1
  tb!Valid = 1
  tb.Update
End If


Exit Sub

cmdPrintesr_Click_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /cmdPrintesr_Click ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub cmdPrintHold_Click()

Dim tb As New Recordset
Dim SQL As String

On Error GoTo cmdPrintHold_Click_Error

pBar = 0

txtSampleID = Format(Val(txtSampleID))
If Val(txtSampleID) = 0 Then Exit Sub

If Trim$(txtSex) = "" Then
  If iMsg("Sex not entered." & vbCrLf & "Do you want to enter sex now?", vbQuestion + vbYesNo) = vbYes Then
    Exit Sub
  End If
End If

If Trim$(txtSampleID) = "" Then
  iMsg "Must have Lab Number.", vbCritical
  Exit Sub
End If

If Len(cmbWard) = 0 Then
  iMsg "Must have Ward entry.", vbCritical
  Exit Sub
End If

If Trim$(cmbWard) = "GP" Then
  If Len(cmbGP) = 0 Then
    iMsg "Must have Ward or GP entry.", vbCritical
    Exit Sub
  End If
End If

If SysOptDemoVal(0) And cmdDemoVal.Caption = "&Validate" Then
  If iMsg("Do you wish to validate demographics !", vbYesNo) = vbNo Then
    Exit Sub
  Else
    cmdDemoVal_Click
  End If
End If

SaveDemographics

If sstabAll.Tab <> 0 Then
  LogTimeOfPrinting txtSampleID, Choose(sstabAll.Tab, "H", "B", "C", "E", "Q", "I", "X")
  SQL = "SELECT * from PrintPending WHERE " & _
        "Department = '" & Choose(sstabAll.Tab, "H", "B", "C", "E", "Q", "I", "X") & "' " & _
        "and SampleID = '" & txtSampleID & "'"
  Set tb = New Recordset
  RecOpenServer 0, tb, SQL
  If tb.EOF Then
    tb.AddNew
  End If
  tb!SampleID = txtSampleID
  tb!Department = Choose(sstabAll.Tab, "H", "B", "C", "E", "Q", "I", "X")
  If SysOptRealImm(0) And tb!Department = "I" Then tb!Department = "J"
  tb!Initiator = UserName
  tb!Ward = cmbWard
  tb!Clinician = cmbClinician
  tb!GP = cmbGP
  tb!UsePrinter = pPrintToPrinter
  tb!ptime = Now
  tb.Update
End If

If sstabAll.Tab = 1 Then
  If lblHaemValid.Visible = False Then
    SaveHaematology 1
  End If
  SQL = "SELECT * from HaemResults WHERE " & _
        "SampleID = '" & txtSampleID & "'"
  Set tb = New Recordset
  RecOpenClient 0, tb, SQL
  If Not tb.EOF Then
    tb!Printed = 1
    tb.Update
  End If
  LoadHaematology
ElseIf sstabAll.Tab = 2 Then
  SQL = "UPDATE bioResults " & _
        "Set valid = 1 , operator = '" & UserCode & "' WHERE " & _
        "SampleID = '" & txtSampleID & "' and valid <> 1"
  Cnxn(0).Execute SQL
ElseIf sstabAll.Tab = 3 Then
  If cmdValidateCoag.Caption = "&Validate" Then
    SQL = "UPDATE CoagResults " & _
          "Set valid = 1 , username = '" & UserCode & "' WHERE " & _
          "SampleID = " & txtSampleID & " and valid <> 1"
    Cnxn(0).Execute SQL
  End If
  SQL = "UPDATE CoagResults " & _
        "Set Printed = 0 WHERE " & _
        "SampleID = '" & txtSampleID & "'"
  Cnxn(0).Execute SQL
ElseIf sstabAll.Tab = 4 Then
  SQL = "UPDATE EndResults " & _
        "Set valid = 1 WHERE " & _
        "SampleID = '" & txtSampleID & "'"
  Cnxn(0).Execute SQL
ElseIf sstabAll.Tab = 5 Then
  SQL = "UPDATE BgaResults " & _
        "Set valid = 1 WHERE " & _
        "SampleID = '" & txtSampleID & "'"
  Cnxn(0).Execute SQL
ElseIf sstabAll.Tab = 6 Then
  SQL = "UPDATE immResults " & _
        "Set valid = 1 WHERE " & _
        "SampleID = '" & txtSampleID & "'"
  Cnxn(0).Execute SQL
End If

Exit Sub

cmdPrintHold_Click_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /cmdPrintHold_Click ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub cmdPrintINR_Click()

On Error GoTo cmdPrintINR_Click_Error

If SysNopas(0) = True Then
  If Trim$(txtNOPAS) = "" Then
    iMsg "Enter Nopas Number", vbCritical
    Exit Sub
  End If
Else
If Trim$(txtChart) = "" Then
  iMsg "Enter Chart Number", vbExclamation
  Exit Sub
End If
End If

txtSampleID = Format(Val(txtSampleID))
If Val(txtSampleID) = 0 Then Exit Sub

pBar = 0

SaveDemographics
SaveCoag 1

With frmINR
  .tnopas = txtNOPAS
  .tChart = txtChart
  .Ward = cmbWard
  .LoadDetails
  .Show 1
End With

Exit Sub

cmdPrintINR_Click_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /cmdPrintINR_Click ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub cmdSaveBGa_Click()

On Error GoTo cmdSaveBGa_Click_Error

pBar = 0

txtSampleID = Format(Val(txtSampleID))
If Val(txtSampleID) = 0 Then Exit Sub

'Added 15/Jul/2004

If cmdValBG.Caption = "&Validate" Then
  SaveBloodGas False
Else
  SaveBloodGas True
End If
SaveComments
UPDATEMRU

cmdSaveBGa.Enabled = False

Exit Sub

cmdSaveBGa_Click_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /cmdSaveBio_Click ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub cmdSaveBio_Click()

On Error GoTo cmdSaveBio_Click_Error

pBar = 0

txtSampleID = Format(Val(txtSampleID))
If Val(txtSampleID) = 0 Then Exit Sub

'Added 15/Jul/2004

If bValidateBio.Caption = "&Validate" Then
  SaveBiochemistry False
Else
  SaveBiochemistry True
End If
SaveComments
UPDATEMRU

cmdSaveBio.Enabled = False

Exit Sub

cmdSaveBio_Click_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /cmdSaveBio_Click ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub cmdSaveCoag_Click()

On Error GoTo cmdSaveCoag_Click_Error

pBar = 0

txtSampleID = Format(Val(txtSampleID))
If Val(txtSampleID) = 0 Then Exit Sub

'added 15/Jul/2004
If cmdValidateCoag.Caption = "&Validate" Then
  SaveCoag 0
Else
  SaveCoag 1
End If

SaveComments
UPDATEMRU

cmdSaveCoag.Enabled = False

Exit Sub

cmdSaveCoag_Click_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /cmdSaveCoag_Click ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub cmdSaveComm_Click()
Dim SQL As String

SQL = "Update haemresults set healthlink = 0 where sampleid = '" & txtSampleID & "'"
Cnxn(0).Execute SQL

SaveComments

End Sub

Private Sub cmdSaveDemographics_Click()

On Error GoTo cmdSaveDemographics_Click_Error

pBar = 0

txtSampleID = Format(Val(txtSampleID))
If Val(txtSampleID) = 0 Then Exit Sub

If Trim$(txtSex) = "" Then
  If iMsg("Sex not entered." & vbCrLf & "Do you want to enter sex now?", vbQuestion + vbYesNo) = vbYes Then
    Exit Sub
  End If
End If

If Trim$(txtSampleID) = "" Then
  iMsg "Must have Lab Number.", vbCritical
  Exit Sub
End If

If Trim$(txtName) <> "" Then
  If Trim$(cmbWard) = "" Then
    iMsg "Must have Ward entry.", vbCritical
    Exit Sub
  End If
  
  If Trim$(cmbWard) = "GP" Then
    If Trim$(cmbGP) = "" Then
      iMsg "Must have GP entry.", vbCritical
      Exit Sub
    End If
  End If
End If


If dtRunDate < dtSampleDate Then
    iMsg "Sample Date After Run Date. Please Amend!"
    Exit Sub
End If

If dtRunDate < dtRecDate Then
    iMsg "Rec. Date After Run Date. Please Amend!"
    Exit Sub
End If

If dtRecDate < dtSampleDate Then
    iMsg "Sample Date After Rec. Date. Please Amend!"
    Exit Sub
End If

If Format(dtRunDate, "dd/MM/yyyy") <> Format(Now, "dd/MM/yyyy") Then
  If iMsg("Rundate not today. Proceed ?", vbYesNo) = vbNo Then
    Exit Sub
  End If
End If
    
cmdSaveDemographics.Caption = "Saving"

SaveDemographics
UPDATEMRU
LoadDemographics
cmdSaveDemographics.Caption = "Save && &Hold"
cmdSaveDemographics.Enabled = False
cmdSaveInc.Enabled = False

txtSampleID.SetFocus

Exit Sub

cmdSaveDemographics_Click_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /cmdSaveDemographics_Click ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub cmdSaveExt_Click()

On Error GoTo bSaveExt_Click_Error

txtSampleID = Format(Val(txtSampleID))
If Val(txtSampleID) = 0 Then Exit Sub

SaveExtern
UPDATEMRU
cmdSaveExt.Enabled = False


LoadExt

Exit Sub

bSaveExt_Click_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /bSaveExt_Click ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub cmdSaveHaem_Click()

On Error GoTo cmdSaveHaem_Click_Error

pBar = 0

txtSampleID = Format(Val(txtSampleID))
If Val(txtSampleID) = 0 Then Exit Sub

'Added 15/Jul/2004

If bValidateHaem.Caption = "&Validate" Then
  SaveHaematology 0
Else
  SaveHaematology 1
End If

SaveComments
UPDATEMRU

txtSampleID = Format$(Val(txtSampleID) + 1)
LoadAllDetails

cmdSaveHaem.Enabled = False
cmdHSaveH.Enabled = False

Exit Sub

cmdSaveHaem_Click_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /cmdSaveHaem_Click ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub cmdSaveImm_Click(Index As Integer)

On Error GoTo cmdSaveImm_Click_Error

pBar = 0


If Index = 0 Then
  pBar = 0
  
  txtSampleID = Format(Val(txtSampleID))
  If Val(txtSampleID) = 0 Then Exit Sub
  
  'added 15/Jul/2004
  If bValidateImm(0).Caption = "&Validate" Then
    SaveEndocrinology False
  Else
    SaveEndocrinology True
  End If
  SaveComments
  UPDATEMRU
  
  cmdSaveImm(0).Enabled = False
Else
  txtSampleID = Format(Val(txtSampleID))
  If Val(txtSampleID) = 0 Then Exit Sub
  
  'added 15/Jul/2004
  If bValidateImm(1).Caption = "&Validate" Then
    SaveImmunology False
  Else
    SaveImmunology True
  End If
  SaveComments
  UPDATEMRU
  
  cmdSaveImm(1).Enabled = False
End If

Exit Sub

cmdSaveImm_Click_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /cmdSaveImm_Click ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub cmdSaveInc_Click()

On Error GoTo cmdSaveInc_Click_Error

pBar = 0

txtSampleID = Format(Val(txtSampleID))
If Val(txtSampleID) = 0 Then Exit Sub

If Trim$(txtSex) = "" Then
  If iMsg("Sex not entered." & vbCrLf & "Do you want to enter sex now?", vbQuestion + vbYesNo) = vbYes Then
    Exit Sub
  End If
End If

If Trim$(txtSampleID) = "" Then
  iMsg "Must have Lab Number.", vbCritical
  Exit Sub
End If

If Trim$(txtName) <> "" Then
  If Trim$(cmbWard) = "" Then
    iMsg "Must have Ward entry.", vbCritical
    Exit Sub
  End If
  
  If Trim$(cmbWard) = "GP" Then
    If Trim$(cmbGP) = "" Then
      iMsg "Must have GP entry.", vbCritical
      Exit Sub
    End If
  End If
End If

If lblChartNumber.BackColor = vbRed And Trim(txtChart) <> "" Then
  If iMsg("Confirm this Patient has" & vbCrLf & _
          lblChartNumber.Caption, vbYesNo + vbQuestion, , vbRed, 12) = vbNo Then
    Exit Sub
  End If
End If

If dtRunDate < dtSampleDate Then
    iMsg "Sample Date After Run Date. Please Amend!"
    Exit Sub
End If

If dtRunDate < dtRecDate Then
    iMsg "Rec. Date After Run Date. Please Amend!"
    Exit Sub
End If

If dtRecDate < dtSampleDate Then
    iMsg "Sample Date After Rec. Date. Please Amend!"
    Exit Sub
End If
          
If Format(dtRunDate, "dd/MMM/yyyy") <> Format(Now, "dd/MMM/yyyy") Then
  If iMsg("Rundate not today. Proceed ?", vbYesNo) = vbNo Then
    Exit Sub
  End If
End If
    
  
cmdSaveDemographics.Caption = "Saving"

SaveDemographics
UPDATEMRU

cmdSaveDemographics.Caption = "Save && &Hold"
cmdSaveDemographics.Enabled = False
cmdSaveInc.Enabled = False

If SysOptBlankSid(0) Then
  txtSampleID = ""
Else
  txtSampleID = Format$(Val(txtSampleID) + 1)
End If

LoadAllDetails

cmdSaveHaem.Enabled = False
cmdHSaveH.Enabled = False
cmdSaveBio.Enabled = False
cmdSaveCoag.Enabled = False
cmdSaveImm(0).Enabled = False
cmdSaveImm(1).Enabled = False
cmdSaveBGa.Enabled = False

txtSampleID.SetFocus

Exit Sub

cmdSaveInc_Click_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /cmdSaveInc_Click ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub cmdSetPrinter_Click()

Set frmForcePrinter.F = frmEditAll
frmForcePrinter.Show 1
  
If pPrintToPrinter = "Automatic SELECTion" Then
  pPrintToPrinter = ""
End If

If pPrintToPrinter <> "" Then
  cmdSetPrinter.BackColor = vbRed
  cmdSetPrinter.ToolTipText = "Print Forced to " & pPrintToPrinter
Else
  cmdSetPrinter.BackColor = vbButtonFace
  pPrintToPrinter = ""
  cmdSetPrinter.ToolTipText = "Printer SELECTed Automatically"
End If
  
End Sub

Private Sub cmdValBG_Click()
Dim SQL As String

On Error GoTo cmdValBG_Click_Error

If SysOptDemoVal(0) And cmdDemoVal.Caption = "&Validate" Then
  If iMsg("Do you wish to validate demographics !", vbYesNo) = vbNo Then
    Exit Sub
  Else
    cmdDemoVal_Click
  End If
End If

If Trim(txtDoB) = "" Then iMsg "No Dob. Adult Age 25 used for Normal Ranges!"

SQL = "UPDATE BgaResults set valid = '1' WHERE sampleid = '" & txtSampleID & "'"
Cnxn(0).Execute SQL

SQL = "UPDATE demographics set forbga = '1' WHERE sampleid = '" & txtSampleID & "'"
Cnxn(0).Execute SQL

cmdValBG.Caption = "VALID"

Exit Sub

cmdValBG_Click_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /cmdValBG_Click ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub cmdValidateCoag_Click()


On Error GoTo cmdValidateCoag_Click_Error

pBar = 0

If cmdSaveCoag.Enabled = True Then
  iMsg "You must first save !"
  Exit Sub
End If


txtSampleID = Format(Val(txtSampleID))
If Val(txtSampleID) = 0 Then Exit Sub

If cmdValidateCoag.Caption = "VALID" Then
  If UCase(iBOX("Unvalidate ! Enter Password" & vbCrLf & "You get only 1 Chance!", , , True)) = UserPass Then
    SaveCoag False
    SaveComments
    'txtCoagComment.Locked = False
    cmdValidateCoag.Caption = "&Validate"
    Me.Refresh
  End If
Else
  If SysOptDemoVal(0) And cmdDemoVal.Caption = "&Validate" Then
    If iMsg("Do you wish to validate demographics !", vbYesNo) = vbNo Then
      Exit Sub
    Else
      cmdDemoVal_Click
    End If
  End If
  If Trim(txtDoB) = "" Then iMsg "No Dob. Adult Age 25 used for Normal Ranges!"
  SaveCoag True
  SaveComments
  UPDATEMRU
  'txtCoagComment.Locked = True
  If SysOptHaemAn1(0) = "ADVIA" Then txtSampleID = Format(Val(txtSampleID)) + 1
  Me.Refresh
End If

LoadAllDetails

Exit Sub

cmdValidateCoag_Click_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /cmdValidateCoag_Click ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub cmdViewBioReps_Click()

frmRFT.SampleID = txtSampleID
frmRFT.Dept = "B"
frmRFT.Show 1

End Sub

Private Sub cmdViewCoagRep_Click()

frmRFT.SampleID = txtSampleID
frmRFT.Dept = "C"
frmRFT.Show 1

End Sub

Private Sub cmdViewExtReport_Click()

frmRFT.SampleID = txtSampleID
frmRFT.Dept = "X"
frmRFT.Show 1

End Sub

Private Sub cmdViewHaemRep_Click()

frmRFT.SampleID = txtSampleID
frmRFT.Dept = "H"
frmRFT.Show 1
  
End Sub

Private Sub cmdViewImmRep_Click()

frmRFT.SampleID = txtSampleID
frmRFT.Dept = "I"
frmRFT.Show 1

End Sub

Private Sub cmdViewReports_Click()

frmRFT.SampleID = txtSampleID
frmRFT.Dept = "E"
frmRFT.Show 1
  
End Sub

Private Sub cMonospot_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If cMonospot = 0 Then
  If Trim$(tMonospot) = "?" Then
    tMonospot = ""
  ElseIf Trim$(tMonospot) <> "" Then
    cMonospot = 1
  End If
Else
  If Trim$(tMonospot) = "" Then
    tMonospot = "?"
  End If
End If

cmdSaveHaem.Enabled = True: cmdHSaveH.Enabled = True

End Sub

Private Sub cMRU_Click()

txtSampleID = cMRU

LoadAllDetails

cmdSaveDemographics.Enabled = False
cmdSaveInc.Enabled = False
cmdSaveHaem.Enabled = False
cmdSaveComm.Enabled = False
cmdHSaveH.Enabled = False
cmdSaveBio.Enabled = False
cmdSaveCoag.Enabled = False
cmdSaveImm(0).Enabled = False
cmdSaveImm(1).Enabled = False
cmdSaveBGa.Enabled = False

End Sub

Private Sub cMRU_KeyPress(KeyAscii As Integer)

KeyAscii = 0

End Sub

Private Sub Colourise(ByVal Analyte As String, _
                      ByVal Destination As TextBox, _
                      ByVal strValue As String, _
                      ByVal sex As String, _
                      ByVal DoB As String)
    
Dim Value As Single
Dim SQL As String
Dim tb As Recordset

On Error GoTo Colourise_Error

Value = Val(strValue)


SQL = "SELECT * from haemtestdefinitions WHERE analytename = '" & Analyte & "'"
Set tb = New Recordset
RecOpenServer 0, tb, SQL
If Not tb.EOF Then
    Select Case Val(tb!Printformat & "")
    Case 0:
        Destination = strValue
    Case 1:
        Destination = Format(strValue, "##0.0")
    Case 2:
        Destination = Format(strValue, "##0.00")
    Case 3:
        Destination = Format(strValue, "##0.000")
    End Select
Else
    Destination = strValue
End If


If Trim$(strValue) = "" Then
  Destination.BackColor = &HFFFFFF
  Destination.ForeColor = &H0&
  Exit Sub
End If

Select Case InterpH(Value, Analyte, sex, DoB, 0, dtRunDate)
  Case "X":
    Destination.BackColor = SysOptPlasBack(0)
    Destination.ForeColor = SysOptPlasFore(0)
  Case "H":
    Destination.BackColor = SysOptHighBack(0)
    Destination.ForeColor = SysOptHighFore(0)
  Case "L"
    Destination.BackColor = SysOptLowBack(0)
    Destination.ForeColor = SysOptLowFore(0)
  Case Else
    Destination.BackColor = &HFFFFFF
    Destination.ForeColor = &H0&
End Select

Exit Sub

Colourise_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /Colourise ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select

  
End Sub

Private Sub ColouriseG(ByVal Analyte As String, _
                      ByVal Destination As MSFlexGrid, _
                      ByVal X As Long, _
                      ByVal Y As Long, _
                      ByVal strValue As String, _
                      ByVal sex As String, _
                      ByVal DoB As String)
    
Dim Value As Single
Dim SQL As String
Dim tb As Recordset

On Error GoTo ColouriseG_Error

Value = Trim(Val(strValue))

SQL = "SELECT * from haemtestdefinitions WHERE analytename = '" & Analyte & "'"
Set tb = New Recordset
RecOpenServer 0, tb, SQL
If Not tb.EOF Then
    Select Case Val(tb!Printformat & "")
    Case 0:
        Destination.TextMatrix(X, Y) = strValue
    Case 1:
        Destination.TextMatrix(X, Y) = Format(strValue, "##0.0")
    Case 2:
        Destination.TextMatrix(X, Y) = Format(strValue, "##0.00")
    Case 3:
        Destination.TextMatrix(X, Y) = Format(strValue, "##0.000")
    End Select
Else
    Destination.TextMatrix(X, Y) = strValue
End If

Destination.Col = Y
Destination.Row = X

If Trim$(strValue) = "" Then
  
  
  Destination.CellBackColor = &HFFFFFF
  Destination.CellForeColor = 1
  Exit Sub
End If

Select Case InterpH(Value, Analyte, sex, DoB, 0)
  Case "X":
    Destination.CellBackColor = SysOptPlasBack(0)
    Destination.CellForeColor = SysOptPlasFore(0)
  Case "H":
    Destination.CellBackColor = SysOptHighBack(0)
    Destination.CellForeColor = SysOptHighFore(0)
  Case "L"
    Destination.CellBackColor = SysOptLowBack(0)
    Destination.CellForeColor = SysOptLowFore(0)
  Case Else
    Destination.CellBackColor = &HFFFFFF
    Destination.CellForeColor = 1
End Select

Exit Sub

ColouriseG_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /ColouriseG ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select

  
End Sub

Private Sub cParameter_Click()
On Error GoTo cParameter_Click_Error

pBar = 0
Dim n As Long
Dim Unit As String

cCunits.Enabled = True

Dim SampleType As String
Dim Code As String

SampleType = ListCodeFor("ST", cSampleType)
Code = CoagCodeFor(cParameter)
Unit = ACoagUnitsFor(cParameter)


If Unit <> "" Then
  For n = 1 To cCunits.ListCount
    If cCunits.List(n) = Trim(Unit) Then
        cCunits.ListIndex = n
        Exit For
    End If
  Next
End If

If cParameter = "PT" Then cCunits.Enabled = True Else cCunits.Enabled = False

Exit Sub

cParameter_Click_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /cParameter_Click ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub cRA_Click()

If cRA = 0 Then
  If Trim$(tRa) = "?" Then
    tRa = ""
  ElseIf Trim$(tRa) <> "" Then
    cRA = 1
  End If
Else
  If Trim$(tRa) = "" Then
    tRa = "?"
  End If
End If

cmdSaveHaem.Enabled = True: cmdHSaveH.Enabled = True
End Sub

Private Function CreateHist(ByVal Dept As String) As String
Dim SQL As String
Dim Asql As String
Dim nSql As String
Dim Csql As String
Dim tsql As String
On Error GoTo CreateHist_Error

If Trim(txtSampleID) = "" Then Exit Function

SQL = "SELECT top 1 Demographics.SampleID, Demographics.RunDate from Demographics, " & Dept & "results WHERE ("

If Trim(txtChart) <> "" Then Csql = "Demographics.Chart = '" & EncryptA(txtChart) & "' "
If Trim(txtNOPAS) <> "" Then nSql = "Demographics.nopas = '" & Trim(txtNOPAS) & "' "
If Trim(txtAandE) <> "" Then Asql = "Demographics.aande = '" & Trim(txtAandE) & "' "

  If Csql <> "" And nSql <> "" And Asql <> "" Then
    tsql = "((" & Csql & ") or (" & nSql & ") or (" & Asql & ")) and "
  ElseIf Csql <> "" And nSql <> "" And Asql = "" Then
    tsql = "((" & Csql & ") or (" & nSql & "))  and "
  ElseIf Csql <> "" And nSql = "" And Asql <> "" Then
    tsql = "((" & Csql & ") or (" & Asql & "))  and "
  ElseIf Csql = "" And nSql <> "" And Asql <> "" Then
    tsql = "((" & nSql & ") or (" & Asql & "))  and "
  ElseIf Csql <> "" And nSql = "" And Asql = "" Then
    tsql = "(" & Csql & ") and "
  ElseIf Csql = "" And nSql <> "" And Asql = "" Then
    tsql = "(" & nSql & ") and "
  ElseIf Csql = "" And nSql = "" And Asql <> "" Then
    tsql = "(" & Asql & ") and "
  End If
  
  
  If tsql = "" Then
    tsql = " Demographics.patname = '" & AddTicks(txtName) & "' and Demographics.dob = '" & Format(txtDoB, "dd/MMM/yyyy") & "' and "
  End If
  
  SQL = SQL & tsql & "Demographics.SampleID <> '" & EncryptN(txtSampleID) & "') and " & Dept & "results.sampleid = Demographics.sampleid " & _
        "order by Demographics.SampleID desc"




CreateHist = SQL

Exit Function

CreateHist_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /CreateHist ")
  Case 1:    End
  Case 2:    Exit Function
  Case 3:    Resume Next
End Select


End Function

Private Function CreateSql(ByVal Dept As String) As String
Dim SQL As String
Dim Asql As String
Dim nSql As String
Dim Csql As String
Dim tsql As String
On Error GoTo CreateSql_Error

SQL = "SELECT top 1 Demographics.SampleID, Demographics.RunDate from Demographics, " & Dept & "results WHERE ("

If Trim(txtChart) <> "" Then Csql = "Demographics.Chart = '" & EncryptA(txtChart) & "' "
If Trim(txtNOPAS) <> "" Then nSql = "Demographics.nopas = '" & Trim(txtNOPAS) & "' "
If Trim(txtAandE) <> "" Then Asql = "Demographics.aande = '" & Trim(txtAandE) & "' "

If Csql <> "" And nSql <> "" And Asql <> "" Then
  tsql = "((" & Csql & ") or (" & nSql & ") or (" & Asql & ")) and "
ElseIf Csql <> "" And nSql <> "" And Asql = "" Then
  tsql = "((" & Csql & ") or (" & nSql & "))  and "
ElseIf Csql <> "" And nSql = "" And Asql <> "" Then
  tsql = "((" & Csql & ") or (" & Asql & "))  and "
ElseIf Csql = "" And nSql <> "" And Asql <> "" Then
  tsql = "((" & nSql & ") or (" & Asql & "))  and "
ElseIf Csql <> "" And nSql = "" And Asql = "" Then
  tsql = "(" & Csql & ") and "
ElseIf Csql = "" And nSql <> "" And Asql = "" Then
  tsql = "(" & nSql & ") and "
ElseIf Csql = "" And nSql = "" And Asql <> "" Then
  tsql = "(" & Asql & ") and "
End If

If tsql = "" Then
  tsql = " Demographics.patname = '" & AddTicks(txtName) & "' and Demographics.dob = '" & Format(txtDoB, "dd/MMM/yyyy") & "' and "
End If

SQL = SQL & tsql & "Demographics.SampleID <> '" & EncryptN(txtSampleID) & "' and Demographics.rundate < '" & Format(lblRundate, "dd/MMM/yyyy") & "') and " & Dept & "results.sampleid = Demographics.sampleid " & _
      "order by Demographics.rundate desc"

CreateSql = SQL

Exit Function

CreateSql_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /CreateSql ")
  Case 1:    End
  Case 2:    Exit Function
  Case 3:    Resume Next
End Select


End Function

Private Sub cRetics_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error GoTo cRetics_MouseUp_Error

If cRetics = 0 Then
  If Trim$(tRetA) = "?" Then
    tRetA = ""
    tRetP = ""
  ElseIf Trim$(tRetA) <> "" Then
    cRetics = 1
  End If
Else
  If Trim$(tRetA) = "" Then
    tRetA = "?"
    tRetP = "?"
  End If
End If

cmdSaveHaem.Enabled = True: cmdHSaveH.Enabled = True

Exit Sub

cRetics_MouseUp_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /cRetics_MouseUp ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub cRooH_Click(Index As Integer)

cmdSaveDemographics.Enabled = True
cmdSaveInc.Enabled = True

End Sub

Private Sub cSampleType_Change()

FillcAdd

End Sub

Private Sub cSampleType_Click()

FillcAdd

End Sub

Private Sub cSampleType_KeyPress(KeyAscii As Integer)

KeyAscii = 0

End Sub

Private Sub cUnits_KeyPress(KeyAscii As Integer)

KeyAscii = 0

End Sub

Private Sub DeltaCheck(ByVal Analyte As String, _
                       ByVal Value As String, _
                       ByVal PreviousValue As String, _
                       ByVal PreviousDate As String, _
                       ByVal PreviousID As String)
                              
Dim HD As HaemTestDefinition

'Set HD = colHaemTestDefinitions(Analyte, 0, MaxAgeToDays, Hospname(0))
'If HD Is Nothing Then Exit Sub

On Error GoTo DeltaCheck_Error

For Each HD In colHaemTestDefinitions
  If HD.AnalyteName = Analyte Then
    If HD.DoDelta And PreviousValue <> 0 Then
      If HD.AgeFromDays > 0 And HD.AgeToDays >= MaxAgeToDays Then
      If Abs(Val(PreviousValue) - Val(Value)) > HD.DeltaValue Then
        pdelta.ForeColor = vbBlue
        pdelta.Print Left$(Format$(PreviousDate, "dd/mm/yyyy") & _
                     "(" & PreviousID & ") " & _
                     Analyte & ":" & Space(25), 25); PreviousValue
        Exit For
      End If
      End If
    End If
End If
Next

Exit Sub

DeltaCheck_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /DeltaCheck ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub dtRecDate_CloseUp()
pBar = 0

cmdSaveDemographics.Enabled = True
cmdSaveInc.Enabled = True
End Sub

Private Sub dtRunDate_CloseUp()

pBar = 0

cmdSaveDemographics.Enabled = True
cmdSaveInc.Enabled = True

End Sub

Private Sub dtSampleDate_CloseUp()

pBar = 0

lblSampledate = dtSampleDate

cmdSaveDemographics.Enabled = True
cmdSaveInc.Enabled = True

End Sub

Private Sub FillcAdd()

Dim tb As New Recordset
Dim SQL As String

On Error GoTo FillcAdd_Error



SQL = "SELECT distinct B.ShortName, B.PrintPriority " & _
      "from BioTestDefinitions as B, Lists as L " & _
      "WHERE B.SampleType = L.Code " & _
      "and L.ListType = 'ST' " & _
      "and L.Text like '" & cSampleType & "%' and b.inuse = '1' " & _
      "order by B.PrintPriority"
Set tb = New Recordset
RecOpenServer 0, tb, SQL

cAdd.Clear

Do While Not tb.EOF
  cAdd.AddItem tb!ShortName
  tb.MoveNext
Loop

Exit Sub

FillcAdd_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /FillcAdd ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub FillCats()

Dim SQL As String
Dim tb As New Recordset

On Error GoTo FillCats_Error

cCat(0).Clear
cCat(1).Clear

SQL = "SELECT * from categorys"
Set tb = New Recordset
RecOpenServer 0, tb, SQL
Do While Not tb.EOF
  cCat(0).AddItem Trim(tb!Cat)
  cCat(1).AddItem Trim(tb!Cat)
  tb.MoveNext
Loop

If cCat(0).ListCount > 0 Then
  cCat(0).ListIndex = 0
  cCat(1).ListIndex = 0
End If

Exit Sub

FillCats_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /FillCats ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub FillcbAdd()

Dim tb As New Recordset
Dim SQL As String
Dim Found As Boolean
Dim n As Long

On Error GoTo FillcbAdd_Error

SQL = "SELECT distinct ShortName, PrintPriority " & _
      "from bgaTestDefinitions " & _
      "WHERE InUse = '1' and sampletype = '" & Left(cISampleType(2), 1) & "' " & _
      "order by shortname"
Set tb = New Recordset
RecOpenServer 0, tb, SQL

cIAdd(2).Clear
Do While Not tb.EOF
  Found = False
  For n = 0 To cIAdd(2).ListCount - 1
    If cIAdd(2).List(n) = tb!ShortName Then
      Found = True
    End If
  Next
  If Not Found Then
    cIAdd(2).AddItem tb!ShortName
  End If
  Found = False
  tb.MoveNext
Loop

Exit Sub

FillcbAdd_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /FillcbAdd ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub FillcEAdd()

Dim tb As New Recordset
Dim SQL As String
Dim Found As Boolean
Dim n As Long

On Error GoTo FillcEAdd_Error

SQL = "SELECT distinct ShortName, PrintPriority " & _
      "from EndTestDefinitions " & _
      "WHERE InUse = '1' and sampletype = '" & ListCodeFor("ST", cISampleType(0)) & "' " & _
      "order by PrintPriority"
Set tb = New Recordset
RecOpenServer 0, tb, SQL

cIAdd(0).Clear
Do While Not tb.EOF
  Found = False
  For n = 0 To cIAdd(0).ListCount - 1
    If cIAdd(0).List(n) = tb!ShortName Then
      Found = True
    End If
  Next
  If Not Found Then
    cIAdd(0).AddItem tb!ShortName
  End If
  tb.MoveNext
Loop

Exit Sub

FillcEAdd_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /FillcEAdd ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub FillcIAdd()

Dim tb As New Recordset
Dim SQL As String
Dim Found As Boolean
Dim n As Long

On Error GoTo FillcIAdd_Error

SQL = "SELECT distinct ShortName, PrintPriority " & _
      "from ImmTestDefinitions " & _
      "WHERE InUse = '1' and sampletype = '" & ListCodeFor("ST", cISampleType(1)) & "' " & _
      "order by shortname"
Set tb = New Recordset
RecOpenServer 0, tb, SQL

cIAdd(1).Clear
Do While Not tb.EOF
  Found = False
  For n = 0 To cIAdd(1).ListCount - 1
    If cIAdd(1).List(n) = tb!ShortName Then
      Found = True
    End If
  Next
  If Not Found Then
    cIAdd(1).AddItem tb!ShortName
  End If
  Found = False
  tb.MoveNext
Loop

Exit Sub

FillcIAdd_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /FillcIAdd ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub FillcParameter()

Dim tb As New Recordset
Dim n As Long
Dim InList As Boolean
Dim InUList As Boolean
Dim SQL As String

On Error GoTo FillcParameter_Error

cParameter.Clear


SQL = "SELECT * from coagtestdefinitions"
Set tb = New Recordset
RecOpenServer 0, tb, SQL

Do While Not tb.EOF
  InList = False
  For n = 0 To cParameter.ListCount - 1
    If cParameter.List(n) = Trim(tb!TestName) Then
      InList = True
    End If
  Next
  If Not InList Then
    cParameter.AddItem Trim(tb!TestName)
  End If
  InUList = False
  For n = 0 To cCunits.ListCount - 1
    If cCunits.List(n) = Trim(tb!Units) Then
      InUList = True
    End If
  Next
  If Not InUList Then
    cCunits.AddItem Trim(tb!Units)
  End If
  tb.MoveNext
Loop


Exit Sub

FillcParameter_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /FillcParameter ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub FillcSampleType()

Dim SQL As String
Dim tb As New Recordset
Dim n As Integer


On Error GoTo FillcSampleType_Error

cSampleType.Clear

SQL = "SELECT * from lists WHERE listtype = 'ST'"
Set tb = New Recordset
RecOpenServer 0, tb, SQL
Do While Not tb.EOF
  cSampleType.AddItem Trim(tb!Text)
  cISampleType(0).AddItem Trim(tb!Text)
  cISampleType(1).AddItem Trim(tb!Text)
  cISampleType(2).AddItem Trim(tb!Text)
  tb.MoveNext
Loop

If cSampleType.ListCount > 0 Then
  For n = 1 To cSampleType.ListCount - 1
    If InStr(UCase(cSampleType.List(n)), "SERUM") > 0 Then
        cSampleType.ListIndex = n
        cISampleType(0).ListIndex = n
        cISampleType(1).ListIndex = n
    End If
  Next
  FillcAdd
  FillcEAdd
  FillcIAdd
  cISampleType(2).ListIndex = 0
End If

Exit Sub

FillcSampleType_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /FillcSampleType ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub FillLists()
Dim tb As New Recordset
Dim SQL As String


On Error GoTo FillLists_Error

FillGPsClinWard Me, Hospname(0)

FillUnits

cClDetails.Clear
cmbHospital.Clear

SQL = "SELECT * from lists WHERE listtype = 'UN' or listtype = 'CD' or listtype = 'HO' order by listorder"
Set tb = New Recordset
RecOpenServer 0, tb, SQL
Do While Not tb.EOF
    If Trim(tb!ListType) = "CD" Then
      cClDetails.AddItem Trim(tb!Text)
    ElseIf Trim(tb!ListType) = "HO" Then
      cmbHospital.AddItem Trim(tb!Text)
    End If
  tb.MoveNext
Loop

cClDetails.ListIndex = -1
cmbHospital.ListIndex = -1

Exit Sub

FillLists_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /FillLists ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select

End Sub

Private Sub FillMRU()

Dim SQL As String
Dim tb As New Recordset

On Error GoTo FillMRU_Error

SQL = "SELECT top 10 * from MRU WHERE " & _
      "UserCode = '" & UserCode & "' " & _
      "Order by DateTime desc"
Set tb = New Recordset
RecOpenServer 0, tb, SQL
      
cMRU.Clear
Do While Not tb.EOF
  cMRU.AddItem Trim$(tb!SampleID & "")
  tb.MoveNext
Loop
If cMRU.ListCount > 0 Then
  cMRU = ""
End If

Exit Sub

FillMRU_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /FillMRU ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub FillUnits()
Dim tb As New Recordset
Dim SQL As String

On Error GoTo FillUnits_Error

cUnits.Clear
cCunits.Clear
cIUnits(0).Clear
cIUnits(1).Clear

SQL = "SELECT * from lists WHERE listtype = 'UN'"
Set tb = New Recordset
RecOpenServer 0, tb, SQL
Do While Not tb.EOF
      cUnits.AddItem Trim(tb!Text)
      cIUnits(0).AddItem Trim(tb!Text)
      cIUnits(1).AddItem Trim(tb!Text)
      cCunits.AddItem Trim(tb!Text)
  tb.MoveNext
  Loop
cUnits.ListIndex = -1
cCunits.ListIndex = -1
cIUnits(0).ListIndex = -1
cIUnits(1).ListIndex = -1

Exit Sub

FillUnits_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /FillUnits ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub FlashNoPrevious()

Dim T As Single
Dim n As Long

On Error GoTo FlashNoPrevious_Error

For n = 1 To 5
  lNoPrevious.Visible = True
  lNoPrevious.Refresh
  T = Timer
  Do While Timer - T < 0.1: DoEvents: Loop
  lNoPrevious.Visible = False
  lNoPrevious.Refresh
  T = Timer
  Do While Timer - T < 0.1: DoEvents: Loop
Next

Exit Sub

FlashNoPrevious_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /FlashNoPrevious ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub Form_Activate()

On Error GoTo Form_Activate_Error

TimerBar.Enabled = True
pBar = 0

Set_Font Me

UpDown1.max = 9999999

Exit Sub

Form_Activate_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /Form_Activate ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub Form_Deactivate()

Me.Refresh
pBar = 0
TimerBar.Enabled = False

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

pBar = 0

End Sub

Private Sub Form_Load()
Dim n As Long
Dim SQL As String
Dim tb As New Recordset
Dim ax As Control

On Error GoTo Form_Load_Error

SQL = "SELECT * from options WHERE " & _
      "username = '" & UserName & "' " & _
      "and description like 'frmEditAll.%' order by contents desc"
Set tb = New Recordset
RecOpenServer 0, tb, SQL
Do While Not tb.EOF
  For Each ax In Me
    If UCase("frmEditAll" & ax.Name) = UCase(Trim(tb!Description)) Then
     ax.TabIndex = tb!Contents
    End If
  Next
  tb.MoveNext
Loop

n = n + 1

UpDown1.max = 9999999 ''(2 ^ 31) - 1

If SysOptDontShowPrevCoag(0) = True Then
   grdPrev.Visible = False
   lblPrevCoag.Visible = False
End If

EndLoaded = False
ImmLoaded = False

StatusBar1.Panels(1).Text = UserName

If SysOptDemoVal(0) = False Then cmdDemoVal.Visible = False
If SysOptDeptBio(0) = False Then sstabAll.TabVisible(2) = False Else n = n + 1
If SysOptDeptHaem(0) = False Then sstabAll.TabVisible(1) = False Else n = n + 1
If SysOptDeptCoag(0) = False Then sstabAll.TabVisible(3) = False Else n = n + 1
If SysOptDeptEnd(0) = False Then sstabAll.TabVisible(4) = False Else n = n + 1
If SysOptDeptBga(0) = False Then sstabAll.TabVisible(5) = False Else n = n + 1
If SysOptDeptImm(0) = False Then sstabAll.TabVisible(6) = False Else n = n + 1
If SysOptDeptExt(0) = False Then sstabAll.TabVisible(7) = False Else n = n + 1
If PrnAll(0) = False Then cmdPrintAll.Visible = False
If SysOptPhone(0) = False Then cmdPhone.Visible = False

If SysOptPgp(0) = True Then
  ssPanPgP.Visible = True
Else
  ssPanPgP.Visible = False
End If

If SysOptHaemAn1(0) = "ADVIA" Then
'  Label3 = "WBCP"
'  Label18 = "WBCB"
'  cmdPrintEsr.Visible = False
End If

'lblNopas(1).Visible = SysNopas(0)
'txtNoPas.Visible = SysNopas(0)
'txtAE.Visible = SysAandE(0)
'lblNopas(1).Visible = SysNopas(0)
'lblNopas(0).Visible = SysNopas(0)
'lblAE.Visible = SysAandE(0)


With lblChartNumber
  .BackColor = &H8000000F
  .ForeColor = vbBlack
  Select Case Hospname(0)
    Case "Mallow", "SIVH", "Bantry", "STJOHNS"
      .Caption = "Chart #"
      lblAandE.Visible = True
      lblNOPAS(0).Visible = True
      lblNOPAS(1).Visible = True
      txtAandE.Visible = True
      txtNOPAS.Visible = True
      lblNameTitle.Left = 4050
      txtName.Left = 4050
      txtName.Width = 3495
    Case "Cavan"
      .Caption = "Cavan Chart #"
      lblAandE.Visible = False
      lblNOPAS(0).Visible = False
      lblNOPAS(1).Visible = False
      txtAandE.Visible = False
      txtNOPAS.Visible = False
      lblNameTitle.Left = 1530
      txtName.Left = 1530
      txtName.Width = 6015
    Case "Monaghan"
      .Caption = "Monaghan Chart #"
      lblAandE.Visible = False
      lblNOPAS(0).Visible = False
      lblNOPAS(1).Visible = False
      txtAandE.Visible = False
      txtNOPAS.Visible = False
      lblNameTitle.Left = 1530
      txtName.Left = 1530
      txtName.Width = 6015
    Case "PORTLAOISE", "TULLAMORE"
      .Caption = initial2upper(Hospname(0)) & " Chart #"
      lblAandE.Visible = False
      lblNOPAS(0).Visible = False
      lblNOPAS(1).Visible = False
      txtAandE.Visible = False
      txtNOPAS.Visible = False
      lblNameTitle.Left = 1530
      txtName.Left = 1530
      txtName.Width = 6015
    Case "MULLINGAR"
      .Caption = initial2upper(Hospname(0)) & " Chart #"
      lblAandE.Visible = True
      lblNOPAS(0).Visible = False
      lblNOPAS(1).Visible = False
      txtAandE.Visible = True
      txtNOPAS.Visible = False
      txtAandE.Width = 2000
      lblNameTitle.Left = 3550
      txtName.Left = 3550
      txtName.Width = 4000
  End Select
End With

sstabAll.TabsPerRow = n

With lblViewSplit
  Select Case GetSetting("NetAcquire", "StartUp", "Split", "All")
    Case "All":
      .Caption = "Viewing All"
      .BackColor = &H8000000F
      .ForeColor = vbBlack
    Case "Pri":
      .Caption = "Viewing Primary Split"
      .BackColor = &H800080
      .ForeColor = &HFF00&
    Case "Viewing Sec":
      .Caption = "Viewing Secondary Split"
      .BackColor = &H800080
      .ForeColor = &HFF00&
  End Select
End With


With lblImmViewSplit(0)
  Select Case GetSetting("NetAcquire", "StartUp", "EndSplit", "All")
    Case "All":
      .Caption = "Viewing All"
      .BackColor = &H8000000F
      .ForeColor = vbBlack
    Case "Pri":
      .Caption = "Viewing Primary Split"
      .BackColor = &H800080
      .ForeColor = &HFF00&
    Case "Viewing Sec":
      .Caption = "Viewing Secondary Split"
      .BackColor = &H800080
      .ForeColor = &HFF00&
  End Select
End With
With lblImmViewSplit(1)
  Select Case GetSetting("NetAcquire", "StartUp", "ImmSplit", "All")
    Case "All":
      .Caption = "Viewing All"
      .BackColor = &H8000000F
      .ForeColor = vbBlack
    Case "Pri":
      .Caption = "Viewing Primary Split"
      .BackColor = &H800080
      .ForeColor = &HFF00&
    Case "Viewing Sec":
      .Caption = "Viewing Secondary Split"
      .BackColor = &H800080
      .ForeColor = &HFF00&
  End Select
End With

cmdViewBioReps.Visible = SysOptRTFView(0)
cmdViewReports.Visible = SysOptRTFView(0)
cmdViewCoagRep.Visible = SysOptRTFView(0)
cmdViewHaemRep.Visible = SysOptRTFView(0)
cmdViewHaemRep.Visible = SysOptRTFView(0)
cmdViewImmRep.Visible = SysOptRTFView(0)
cmdViewExtReport.Visible = SysOptRTFView(0)

FillcSampleType
FillcParameter
FillLists
FillCats
ClearHaemDiffGrid

FillMRU
ClearRbcGrid

With lblChartNumber
  .BackColor = &H8000000F
  .ForeColor = vbBlack
  Select Case Entity(0)
    Case "01"
      .Caption = "Cavan Chart #"
    Case "31"
      .Caption = "Monaghan Chart #"
    Case "03"
      .Caption = "Portlaoise Chart #"
    Case "04"
      .Caption = "Tullamore Chart #"
    Case "21"
      .Caption = "St Johns Chart #"
  End Select
End With

dtRunDate = Format$(Now, "dd/mm/yyyy")
lblRundate = dtRunDate
dtSampleDate = Format$(Now, "dd/mm/yyyy")

UpDown1.max = 999999

txtSampleID = GetSetting("NetAcquire", "StartUp", "LastUsed", "1")


LoadAllDetails

pBar.max = LogOffDelaySecs

If UserMemberOf = "Secretarys" Then
  For n = 1 To 6
    sstabAll.TabVisible(n) = False
  Next
Else
'  cmdSaveDemographics.Enabled = False
'  cmdSaveInc.Enabled = False
'  cmdSaveHaem.Enabled = False
'  cmdSaveComm.Enabled = False
'  cmdHSaveH.Enabled = False
'  cmdSaveBio.Enabled = False
'  cmdSaveCoag.Enabled = False
'  cmdSaveImm(0).Enabled = False
'  cmdSaveImm(1).Enabled = False
End If



Activated = False

Exit Sub

Form_Load_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /Form_Load ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

pBar = 0

End Sub

Private Sub Form_Paint()

Dim TabNumber As Long
Dim ax As Control

On Error GoTo Form_Paint_Error

If Activated Then Exit Sub

Activated = True

If SysOptDefaultTab(0) <> "" Then
  TabNumber = Val(SysOptDefaultTab(0))
Else
  TabNumber = Val(GetSetting("NetAcquire", "StartUp", "LastDepartment", "0"))
End If

If SysOptDontShowPrevCoag(0) = True Then
   grdPrev.Visible = False
   lblPrevCoag.Visible = False
End If


If sstabAll.TabVisible(TabNumber) = False Then TabNumber = 1

If UserMemberOf = "Secretarys" Then TabNumber = 0

sstabAll.Tab = TabNumber

For Each ax In Me
  If ax.Name = SysSetFoc(0) Then
    ax.SetFocus
  End If
Next

Exit Sub

Form_Paint_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /Form_Paint ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub Form_Unload(Cancel As Integer)

Dim strSplitStatus As String

On Error GoTo Form_Unload_Error

If Val(txtSampleID) <> Val(GetSetting("NetAcquire", "StartUp", "LastUsed", "1")) Then
  SaveSetting "NetAcquire", "StartUp", "LastUsed", txtSampleID
End If

SaveSetting "NetAcquire", "StartUp", "LastDepartment", CStr(sstabAll.Tab)

With lblViewSplit
  If InStr(.Caption, "All") Then
    strSplitStatus = "All"
  ElseIf InStr(.Caption, "Pri") Then
    strSplitStatus = "Pri"
  ElseIf InStr(.Caption, "Sec") Then
    strSplitStatus = "Sec"
  End If
  SaveSetting "NetAcquire", "StartUp", "Split", strSplitStatus
End With


With lblImmViewSplit(0)
  If InStr(.Caption, "All") Then
    strSplitStatus = "All"
  ElseIf InStr(.Caption, "Pri") Then
    strSplitStatus = "Pri"
  ElseIf InStr(.Caption, "Sec") Then
    strSplitStatus = "Sec"
  End If
  SaveSetting "NetAcquire", "StartUp", "EndSplit", strSplitStatus
End With


With lblImmViewSplit(1)
  If InStr(.Caption, "All") Then
    strSplitStatus = "All"
  ElseIf InStr(.Caption, "Pri") Then
    strSplitStatus = "Pri"
  ElseIf InStr(.Caption, "Sec") Then
    strSplitStatus = "Sec"
  End If
  SaveSetting "NetAcquire", "StartUp", "ImmSplit", strSplitStatus
End With


pPrintToPrinter = ""

Activated = False

Exit Sub

Form_Unload_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /Form_Unload ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

pBar = 0

End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

pBar = 0

End Sub

Private Sub Frame3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

pBar = 0

End Sub

Private Sub Frame4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

pBar = 0

End Sub

Private Sub Frame5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

pBar = 0

End Sub

Private Sub Frame6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

pBar = 0

End Sub

Private Sub Frame7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

pBar = 0

End Sub

Private Sub Frame8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

pBar = 0

End Sub

Private Sub Frame9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

pBar = 0

End Sub

Private Sub gBga_Click()
Dim SQL As String

  If gBga.MouseRow = 0 Then Exit Sub
  
  If iMsg("DELETE " & gBga.TextMatrix(gBga.Row, 0) & " !", vbYesNo) <> vbYes Then
    Exit Sub
  End If
  
  If InStr(gBga.TextMatrix(gBga.Row, 6), "V") > 0 Then
    If UCase(iBOX("Valid Result! Enter Password", , , True)) <> UserPass Then
      Exit Sub
    End If
  End If
    
     
  tINewValue(2) = gBga.TextMatrix(gBga.Row, 1)
  cIUnits(2) = gBga.TextMatrix(gBga.Row, 2)
  
  gBga.Col = 0
  cIAdd(2) = gBga
  
  SQL = "DELETE from bgaresults WHERE " & _
        "sampleid = '" & txtSampleID & "' " & _
        "and code = '" & BgaCodeForShortName(gBga) & "'"
  Cnxn(0).Execute SQL
  
  LoadBloodGas
  
  tINewValue(2).SetFocus

End Sub

Private Sub gBio_Click()

Dim tb As New Recordset
Dim SQL As String
Dim s As String
Dim strShortname As String
Dim strCode As String
Dim RoSEl As Long

On Error GoTo gBio_Click_Error

If gBio.MouseRow = 0 Then Exit Sub
If gBio.RowSel = 0 Then Exit Sub

If gBio.TextMatrix(gBio.Row, 0) = "HbA1c" Then
  frmViewFullDataHBA.SampleID = txtSampleID
  frmViewFullDataHBA.Show 1
  Exit Sub
End If
   
If gBio.Col = 5 Then
  Select Case gBio
    Case "": Exit Sub
    Case "AE": s = "ADC Error"
    Case "AH": s = "Initial Absorbance High"
    Case "BH": s = "Blank Absorbance High"
    Case "BL": s = "Blank Absorbance Low"
    Case "BN": s = "Blank Mean Deviation"
    Case "BO": s = "Blank Maximum Deviation"
    Case "DH": s = "Dynamic Range High"
    Case "DL": s = "Dynamic range Low"
    Case "DR": s = "Reference Drift (ISE)"
    Case "EA": s = "Erratic ADC (ISE)"
    Case "HR": s = "Reaction Absorbance High"
    Case "IR": s = "Initial Absorbance High"
    Case "IT": s = "Iteration Tolerance"
    Case "LR": s = "Reaction Absorbance Low"
    Case "NT": s = "Noise Threshold"
    Case "OH": s = "ORDAC High"
    Case "OL": s = "ORDAC Low"
    Case "OT": s = "Outliers Threshold"
    Case "RH": s = "Reaction Rate High"
    Case "RL": s = "Reaction Rate Low"
    Case "RN": s = "Reaction Mean Deviation"
    Case "RO": s = "Reaction Maximum Deviation"
    Case "SD": s = "Substrate Depleted"
    Case "SH": s = "Blank Rate High"
    Case "SL": s = "Blank Rate Low"
    Case "TM": s = "Temperature"
    Case Else: s = "Unknown Error"
  End Select
  iMsg s, vbInformation
  Exit Sub
End If

If InStr(gBio.TextMatrix(gBio.Row, 6), "V") > 0 Then
  If UCase(iBOX("Valid Result! Enter Password", , , True)) <> UserPass Then
    Exit Sub
  End If
End If

If gBio.MouseCol = 7 Then
  RoSEl = gBio.RowSel
  If gBio = "" Then
      gBio = "P"
  ElseIf gBio = "P" Then
      gBio = "C"
  ElseIf gBio = "C" Then
      gBio = "PC"
  ElseIf gBio = "PC" Then
      gBio = ""
  End If
  SQL = "UPDATE Bioresults set PC = '" & gBio & "' WHERE sampleid = '" & txtSampleID & "' " & _
         " and code = '" & CodeForShortName(gBio.TextMatrix(RoSEl, 0)) & "'"
  Cnxn(0).Execute SQL
  Exit Sub
End If

If gBio.MouseCol = 9 Then
  RoSEl = gBio.RowSel
  With frmComment
     .Discipline = "BIO"
      .X = gBio.MouseRow
      .txtComment = gBio
      .Show 1
  End With
  SQL = "UPDATE Bioresults set comment = '" & gBio & "' WHERE sampleid = '" & txtSampleID & "' " & _
         " and code = '" & CodeForShortName(gBio.TextMatrix(RoSEl, 0)) & "'"
  Cnxn(0).Execute SQL
  Exit Sub
End If

If gBio.ColSel = 1 Then
  If iMsg("DELETE " & gBio.TextMatrix(gBio.Row, 0) & " !", vbYesNo) = vbYes Then
    tnewvalue = gBio.TextMatrix(gBio.Row, 1)
    cUnits = gBio.TextMatrix(gBio.Row, 2)
    cAdd = gBio.TextMatrix(gBio.Row, 0)
    gBio.Col = 0
    SQL = "SELECT shortname, code from biotestdefinitions WHERE shortname = '" & gBio.TextMatrix(gBio.Row, 0) & "' and inuse = 1"
    Set tb = New Recordset
    RecOpenServer 0, tb, SQL
    If Not tb.EOF Then
      strShortname = Trim(tb!ShortName & "")
      strCode = Trim(tb!Code & "")
    Else
      strShortname = ""
      strCode = ""
    End If
    SQL = "SELECT * from bioresults WHERE " & _
          "sampleid = '" & txtSampleID & "' " & _
          "and code = '" & strCode & "'"
    Set tb = New Recordset
    RecOpenServer 0, tb, SQL
    If Not tb.EOF Then
      Archive 0, tb, "ArcBioResults", txtSampleID
    End If
    Frame2.Enabled = True
    SQL = "DELETE from bioresults WHERE " & _
          "sampleid = '" & txtSampleID & "' " & _
          "and code = '" & strCode & "'"
    Cnxn(0).Execute SQL
    LoadBiochemistry

    tnewvalue.SetFocus
  End If
End If


Exit Sub

gBio_Click_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /gBio_Click ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub gBio_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Y = gBio.MouseCol
X = gBio.MouseRow
gBio.ToolTipText = "Biochemistry Results"

If gBio.MouseCol = 9 Then
    If Trim(gBio.TextMatrix(X, Y)) <> "" Then gBio.ToolTipText = gBio.TextMatrix(X, Y)
ElseIf gBio.MouseCol = 7 Then
    If gBio.TextMatrix(X, Y) = "P" Then
        gBio.ToolTipText = "Phoned"
    ElseIf gBio.TextMatrix(X, Y) = "C" Then
        gBio.ToolTipText = "Checked"
    ElseIf gBio.TextMatrix(X, Y) = "PC" Then
        gBio.ToolTipText = "Checked & Phoned"
    Else
        gBio.ToolTipText = "Biochemistry Results"
    End If
End If

pBar = 0

End Sub

Private Sub gImm_Click(Index As Integer)

Dim SQL As String
Dim X As Long

On Error GoTo gImm_Click_Error



If Index = 0 Then
  
  
  
  If gImm(0).MouseRow = 0 Then Exit Sub
  
'    If UCase$(gImm(0).TextMatrix(gImm(0).Row, 0)) = "HBA1C" Then
'      frmViewFullDataHBA.HbA1c = gImm(0).TextMatrix(gImm(0).Row, 1)
'      frmViewFullDataHBA.Sampleid = txtSampleID
'      frmViewFullDataHBA.Show 1
'      Exit Sub
'    End If
  
  If gImm(0).MouseCol = 7 Then
    X = gImm(0).MouseRow
    With frmComment
        .X = X
        .Discipline = "END"
        .txtComment = gImm(0)
        .Show 1
    End With
    SQL = "UPDATE endresults set comment = " & _
          " '" & gImm(0).TextMatrix(X, 7) & "' " & _
          "WHERE sampleid = '" & txtSampleID & "' " & _
          "and code = '" & eCodeForShortName(gImm(0).TextMatrix(X, 0)) & "'"
    Cnxn(0).Execute SQL
    Exit Sub
  End If
  
  If iMsg("DELETE " & gImm(0).TextMatrix(gImm(0).Row, 0) & " !", vbYesNo) <> vbYes Then
    Exit Sub
  End If
  
  If InStr(gImm(0).TextMatrix(gImm(0).Row, 6), "V") > 0 Then
    If UCase(iBOX("Valid Result! Enter Password", , , True)) <> UserPass Then
      Exit Sub
    End If
  End If
    
     
  tINewValue(0) = gImm(0).TextMatrix(gImm(0).Row, 1)
  cIUnits(0) = gImm(0).TextMatrix(gImm(0).Row, 2)
  
  gImm(0).Col = 0
  cIAdd(0) = gImm(0)
  
  SQL = "DELETE from endresults WHERE " & _
        "sampleid = '" & txtSampleID & "' " & _
        "and code = '" & eCodeForShortName(gImm(0)) & "'"
  Cnxn(0).Execute SQL
  
  LoadEndocrinology
  
  tINewValue(0).SetFocus
Else
  If gImm(1).MouseRow = 0 Then Exit Sub
  
  If gImm(1).TextMatrix(gImm(1).MouseRow, 0) = "" Then Exit Sub
  
  If gImm(1).MouseCol = 7 Then
    X = gImm(1).MouseRow
    If gImm(1) = "" Then
        gImm(1) = "P"
    ElseIf gImm(1) = "P" Then
        gImm(1) = "C"
    ElseIf gImm(1) = "C" Then
        gImm(1) = "PC"
    ElseIf gImm(1) = "PC" Then
        gImm(1) = ""
    End If
    SQL = "UPDATE immresults set pc = " & _
          " '" & gImm(1).TextMatrix(X, 7) & "' " & _
          "WHERE sampleid = '" & txtSampleID & "' " & _
          "and code = '" & gImm(1).TextMatrix(X, 0) & "'"
    Cnxn(0).Execute SQL
    Exit Sub
  End If
  
  If gImm(1).MouseCol = 8 Then
    X = gImm(1).MouseRow
    With frmComment
        .X = X
        .Discipline = "IMM"
        .txtComment = gImm(1)
        .Show 1
    End With
    SQL = "UPDATE immresults set comment = " & _
          " '" & gImm(1).TextMatrix(X, 8) & "' " & _
          "WHERE sampleid = '" & txtSampleID & "' " & _
          "and code = '" & gImm(1).TextMatrix(X, 0) & "'"
    Cnxn(0).Execute SQL
    Exit Sub
  End If
  
  If iMsg("DELETE " & gImm(1).TextMatrix(gImm(1).Row, 0) & " !", vbYesNo) <> vbYes Then
    Exit Sub
  End If
  
  If InStr(gImm(1).TextMatrix(gImm(1).Row, 6), "V") > 0 Then
    If UCase(iBOX("Valid Result! Enter Password", , , True)) <> UserPass Then
      Exit Sub
    End If
  End If
    
     
  tINewValue(1) = gImm(1).TextMatrix(gImm(1).Row, 1)
  cIUnits(1) = gImm(1).TextMatrix(gImm(1).Row, 2)
  
  gImm(1).Col = 0
  cIAdd(1) = gImm(1)
  
  SQL = "DELETE from immresults WHERE " & _
        "sampleid = '" & txtSampleID & "' " & _
        "and code = '" & ICodeForShortName(gImm(1)) & "'"
  Cnxn(0).Execute SQL
  
  LoadImmunology
  
  tINewValue(1).SetFocus
End If

Exit Sub

gImm_Click_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /gImm_Click ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select

End Sub

Private Sub gImm_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If Index = 1 Then
    Y = gImm(1).MouseCol
    X = gImm(1).MouseRow
    gImm(1).ToolTipText = "Immunology Results"
    
    If gImm(1).MouseCol = 8 Then
        If Trim(gImm(1).TextMatrix(X, Y)) <> "" Then gImm(1).ToolTipText = gImm(1).TextMatrix(X, Y)
    ElseIf gImm(1).MouseCol = 1 Then
        If Trim(gImm(1).TextMatrix(X, Y)) <> "" Then gImm(1).ToolTipText = gImm(1).TextMatrix(X, Y)
    ElseIf gImm(1).MouseCol = 7 Then
        If gImm(1).TextMatrix(X, Y) = "P" Then
            gImm(1).ToolTipText = "Phoned"
        ElseIf gImm(1).TextMatrix(X, Y) = "C" Then
            gImm(1).ToolTipText = "Checked"
        ElseIf gImm(1).TextMatrix(X, Y) = "PC" Then
            gImm(1).ToolTipText = "Checked & Phoned"
        Else
            gImm(1).ToolTipText = "Immunology Results"
        End If
    End If
Else
    Y = gImm(0).MouseCol
    X = gImm(0).MouseRow
    gImm(0).ToolTipText = "Immunology Results"
    
    If gImm(0).MouseCol = 7 Then
        If Trim(gImm(0).TextMatrix(X, Y)) <> "" Then gImm(0).ToolTipText = gImm(0).TextMatrix(X, Y)
    ElseIf gImm(0).MouseCol = 1 Then
        If Trim(gImm(0).TextMatrix(X, Y)) <> "" Then gImm(0).ToolTipText = gImm(0).TextMatrix(X, Y)
    ElseIf gImm(0).MouseCol = 5 Then
        If gImm(0).TextMatrix(X, Y) = "P" Then
            gImm(0).ToolTipText = "Phoned"
        ElseIf gImm(0).TextMatrix(X, Y) = "C" Then
            gImm(0).ToolTipText = "Checked"
        ElseIf gImm(0).TextMatrix(X, Y) = "PC" Then
            gImm(0).ToolTipText = "Checked & Phoned"
        Else
            gImm(0).ToolTipText = "Endcrinology Results"
        End If
    End If
End If

End Sub

Private Sub gRBC_Click()


On Error GoTo gRBC_Click_Error

If gRbc.ColSel = 0 And gRbc.RowSel = 2 Then
  ClearHgb
End If


If gRbc.ColSel = 1 Then
    txtInput.Text = gRbc.TextMatrix(gRbc.RowSel, 1)
    txtInput.SetFocus
    Exit Sub
End If




If SysOptHaemAn1(0) = "ADVIA" Then
If Trim(gRbc.TextMatrix(11, 1)) = "" Then Exit Sub
'  n = 100 - Val(gRbc.TextMatrix(12, 1))
'
'  tWBC = (tWBC / 100) * n
End If



Exit Sub

gRBC_Click_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /gRBC_Click ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub grdCoag_Click()
Dim tb As New Recordset
Dim SQL As String
Dim Code As String



  
On Error GoTo grdCoag_Click_Error

  If grdCoag.MouseRow = 0 Then Exit Sub
  
  If grdCoag.TextMatrix(grdCoag.Row, 0) = "" Then Exit Sub
  
  
  If InStr(grdCoag.TextMatrix(grdCoag.Row, 5), "V") > 0 Then
    If UCase(iBOX("Valid Result! Enter Password", , , True)) <> UserPass Then
      Exit Sub
    End If
  End If
  
  
With grdCoag
  
  Select Case .Col
    
    Case 0:
      If iMsg("DELETE " & .Text & "?", vbQuestion + vbYesNo) = vbYes Then
        cParameter = .Text
        tResult = .TextMatrix(.Row, 1)
        cCunits = .TextMatrix(.Row, 2)
        If .Rows = 2 Then
          .AddItem ""
          .RemoveItem 1
        Else
          .RemoveItem .Row
        End If
        Code = CoagCodeFor(cParameter)
        SQL = "SELECT * from coagresults WHERE " & _
              "sampleid = '" & txtSampleID & "' " & _
              "and Code = '" & Code & "'"
        Set tb = New Recordset
        RecOpenServer 0, tb, SQL
        If Not tb.EOF Then
          Archive 0, tb, "ArcCoagResults", txtSampleID
        End If
        SQL = "DELETE from CoagResults WHERE " & _
              "SampleID = '" & txtSampleID & "' " & _
              "and Code = '" & Code & "'"
              cmdSaveCoag.Enabled = True
        Cnxn(0).Execute SQL
      End If
    
    Case 1:
      .Text = iBOX("Enter new Value for " & .TextMatrix(.Row, 0), , .Text)
      cmdSaveCoag.Enabled = True
    
    Case 5:
      .Text = IIf(.Text = "", "V", "")
      cmdSaveCoag.Enabled = True
  
  End Select

End With

Exit Sub

grdCoag_Click_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /grdCoag_Click ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub grdCoag_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

pBar = 0

End Sub

Private Sub grdExt_Click()
Dim Str As String
Dim Prompt As String

On Error GoTo g_Click_Error

If grdExt.MouseRow = 0 Then Exit Sub
If grdExt.Col = 2 Then
  Prompt = "Enter result for " & grdExt.TextMatrix(grdExt.Row, 1)
  Str = iBOX(Prompt, , grdExt.TextMatrix(grdExt.Row, 2))
  If Str <> "" Then
    grdExt.TextMatrix(grdExt.Row, 2) = Str
    grdExt.TextMatrix(grdExt.Row, 7) = Format(Now, "dd/mmm/yyyy")
  End If
ElseIf grdExt.Col = 8 Then
  Prompt = "Enter Sap Code for " & grdExt.TextMatrix(grdExt.Row, 1)
  Str = iBOX(Prompt, , grdExt.TextMatrix(grdExt.Row, 2))
  If Str <> "" Then
    grdExt.TextMatrix(grdExt.Row, 8) = Str
  End If
End If
cmdSaveExt.Enabled = True

Exit Sub

g_Click_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /g_Click ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub grdH_Click()

If grdH.Height = 360 Then
  grdH.Height = 2000
End If

End Sub

Private Sub grdOutstanding_Click()

Dim tb As New Recordset
Dim SQL As String

On Error GoTo grdOutstanding_Click_Error

With grdOutstanding
  If .MouseRow = 0 Then Exit Sub
  If .Text = "" Then Exit Sub
  If iMsg("Remove " & .Text & " from Requests?", vbQuestion + vbYesNo) = vbYes Then
    SQL = "DELETE from BioRequests WHERE " & _
          "SampleID = '" & txtSampleID & "' " & _
          "and code = '" & CodeForShortName(.Text) & "'"
    Set tb = New Recordset
    RecOpenServer 0, tb, SQL
    If .Rows > 2 Then
      .RemoveItem .Row
    Else
      .AddItem ""
      .RemoveItem 1
    End If
  End If
End With

Exit Sub

grdOutstanding_Click_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /grdOutstanding_Click ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub grdOutstandingCoag_Click()
Dim tb As New Recordset
Dim SQL As String

On Error GoTo grdOutstandings_Click_Error

  With grdOutstandingCoag
    If .MouseRow = 0 Then Exit Sub
    If .Text = "" Then Exit Sub
    If iMsg("Remove " & .Text & " from Requests?", vbQuestion + vbYesNo) = vbYes Then
      SQL = "DELETE from coagRequests WHERE " & _
            "SampleID = '" & txtSampleID & "' " & _
            "and code = '" & CoagCodeFor(.Text) & "'"
      Set tb = New Recordset
      RecOpenClient 0, tb, SQL
      If .Rows > 2 Then
        .RemoveItem .Row
      Else
        .AddItem ""
        .RemoveItem 1
      End If
    End If
  End With

Exit Sub

grdOutstandings_Click_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /grdOutstandings_Click ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select



End Sub

Private Sub grdOutstandings_Click(Index As Integer)


Dim tb As New Recordset
Dim SQL As String

On Error GoTo grdOutstandings_Click_Error

If Index = 0 Then
  With grdOutstandings(0)
    If .MouseRow = 0 Then Exit Sub
    If .Text = "" Then Exit Sub
    If iMsg("Remove " & .Text & " from Requests?", vbQuestion + vbYesNo) = vbYes Then
      SQL = "DELETE from EndRequests WHERE " & _
            "SampleID = '" & txtSampleID & "' " & _
            "and code = '" & eCodeForShortName(.Text) & "'"
      Set tb = New Recordset
      RecOpenClient 0, tb, SQL
      If .Rows > 2 Then
        .RemoveItem .Row
      Else
        .AddItem ""
        .RemoveItem 1
      End If
    End If
  End With
Else
  With grdOutstandings(1)
    If .MouseRow = 0 Then Exit Sub
    If .Text = "" Then Exit Sub
    If iMsg("Remove " & .Text & " from Requests?", vbQuestion + vbYesNo) = vbYes Then
      SQL = "DELETE from ImmRequests WHERE " & _
            "SampleID = '" & txtSampleID & "' " & _
            "and code = '" & ICodeForShortName(.Text) & "'"
      Set tb = New Recordset
      RecOpenClient 0, tb, SQL
      If .Rows > 2 Then
        .RemoveItem .Row
      Else
        .AddItem ""
        .RemoveItem 1
      End If
    End If
  End With
End If

Exit Sub

grdOutstandings_Click_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /grdOutstandings_Click ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub grdOutstandings_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

pBar = 0

End Sub

Private Sub Ig_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = 0 Then
  EndChanged = True
  cmdSaveImm(0).Enabled = True
Else
  ImmChanged = True
  cmdSaveImm(1).Enabled = True
End If

If Ig(Index).Value = 1 Then txtImmComment(Index) = Trim(txtImmComment(Index) & " " & Ig(Index).Caption)

End Sub

Private Sub Ih_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If Index = 0 Then
  EndChanged = True
  cmdSaveImm(0).Enabled = True

Else
  ImmChanged = True
  cmdSaveImm(1).Enabled = True
End If

If Ih(Index).Value = 1 Then txtImmComment(Index) = Trim(txtImmComment(Index) & " " & Ih(Index).Caption)

End Sub

Private Sub Iis_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If Index = 0 Then
  EndChanged = True
  cmdSaveImm(0).Enabled = True
Else
  ImmChanged = True
  cmdSaveImm(1).Enabled = True
End If

If Iis(Index).Value = 1 Then txtImmComment(Index) = Trim(txtImmComment(Index) & " " & Iis(Index).Caption)

End Sub

Private Sub Ij_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If Index = 0 Then
  EndChanged = True
  cmdSaveImm(0).Enabled = True
Else
  ImmChanged = True
  cmdSaveImm(1).Enabled = True
End If

If Ij(Index).Value = 1 Then txtImmComment(Index) = Trim(txtImmComment(Index) & " " & Ij(Index).Caption)

End Sub

Private Sub Il_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = 0 Then
  EndChanged = True
  cmdSaveImm(0).Enabled = True
Else
  ImmChanged = True
  cmdSaveImm(1).Enabled = True
End If

If Il(Index).Value = 1 Then txtImmComment(Index) = Trim(txtImmComment(Index) & " " & Il(Index).Caption)

End Sub

Private Sub imgLast_Click()

Dim SQL As String
Dim tb As New Recordset
Dim strDept As String
Dim strSplitSELECT As String



On Error GoTo imgLast_Click_Error

Select Case sstabAll.Tab
  Case 0:
    txtSampleID = Format$(Val(txtSampleID) + 1)
    LoadAllDetails

    cmdSaveDemographics.Enabled = False
    cmdSaveInc.Enabled = False
    cmdSaveHaem.Enabled = False
    cmdSaveComm.Enabled = False
    cmdHSaveH.Enabled = False
    cmdSaveBio.Enabled = False
    cmdSaveCoag.Enabled = False
    cmdSaveImm(0).Enabled = False
    cmdSaveImm(1).Enabled = False
    cmdSaveBGa.Enabled = False
    Exit Sub
  
  Case 1: strDept = "Haem"
  Case 2: strDept = "Bio"
  Case 3: strDept = "Coag"
  Case 4: strDept = "End"
  Case 5: strDept = "Bga"
  Case 6: strDept = "Imm"
  Case 7: strDept = "Ext"
End Select

SQL = "SELECT top 1 SampleID from " & strDept & "Results "
If Hospname(0) = "PORTLAOISE" Then
  SQL = SQL & "WHERE sampleid < 9000000 "
End If
  SQL = SQL & "Order by SampleID desc"

If strDept = "Bio" Then
  If InStr(lblViewSplit, "Pri") Then
    strSplitSELECT = LoadSplitList(1)
  ElseIf InStr(lblViewSplit, "Sec") Then
    strSplitSELECT = LoadSplitList(2)
  End If
  If strSplitSELECT <> "" Then
    SQL = "SELECT top 1 SampleID from " & strDept & "Results WHERE " & _
          "(" & strSplitSELECT & ") "
          If Hospname(0) = "PORTLAOISE" Then
            SQL = SQL & "and sampleid < 9000000 "
          End If
          SQL = SQL & "Order by SampleID desc"
  End If
ElseIf strDept = "Imm" Then
  If InStr(lblImmViewSplit(1), "Pri") Then
    strSplitSELECT = LoadImmSplitList(1)
  ElseIf InStr(lblImmViewSplit(1), "Sec") Then
    strSplitSELECT = LoadImmSplitList(2)
  End If
  If strSplitSELECT <> "" Then
    SQL = "SELECT top 1 SampleID from " & strDept & "Results WHERE " & _
          "(" & strSplitSELECT & ") "
          If Hospname(0) = "PORTLAOISE" Then
            SQL = SQL & "and sampleid < 9000000 "
          End If
          SQL = SQL & "Order by SampleID desc"
  End If

End If

Set tb = New Recordset
RecOpenServer 0, tb, SQL
If Not tb.EOF Then
  txtSampleID = tb!SampleID & ""
End If

LoadAllDetails

cmdSaveDemographics.Enabled = False
cmdSaveInc.Enabled = False
cmdSaveHaem.Enabled = False
cmdSaveComm.Enabled = False
cmdHSaveH.Enabled = False
cmdSaveBio.Enabled = False
cmdSaveCoag.Enabled = False
cmdSaveImm(0).Enabled = False
cmdSaveImm(1).Enabled = False
cmdSaveBGa.Enabled = False



Exit Sub

imgLast_Click_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /imgLast_Click ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub Io_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If Index = 0 Then
  EndChanged = True
  cmdSaveImm(0).Enabled = True
Else
  ImmChanged = True
  cmdSaveImm(1).Enabled = True
End If

If Io(Index).Value = 1 Then txtImmComment(Index) = Trim(txtImmComment(Index) & " " & Io(Index).Caption)

End Sub

Private Sub iRecDate_Click(Index As Integer)

If Index = 0 Then
  dtRecDate = DateAdd("d", -1, dtRecDate)
Else
  If DateDiff("d", dtRecDate, Now) > 0 Then
    dtRecDate = DateAdd("d", 1, dtRecDate)
  End If
End If

cmdSaveInc.Enabled = True
cmdSaveDemographics.Enabled = True
End Sub

Private Sub irelevant_Click(Index As Integer)

Dim SQL As String
Dim tb As New Recordset
Dim strDept As String
Dim strDirection As String
Dim strSplitSELECT As String
Dim strArrow As String



On Error GoTo irelevant_Click_Error

If txtSampleID = "" Then Exit Sub

Select Case sstabAll.Tab
  Case 0:
    If Index = 0 Then
      txtSampleID = Format$(Val(txtSampleID) - 1)
    Else
      txtSampleID = Format$(Val(txtSampleID) + 1)
    End If
    
    If SysOptNumLen(0) > 0 Then
      If Len(txtSampleID) > SysOptNumLen(0) Then
        iMsg "Sample Id longer then recommended!"
      End If
    End If
    
    LoadAllDetails

    cmdSaveDemographics.Enabled = False
    cmdSaveInc.Enabled = False
    cmdSaveHaem.Enabled = False
    cmdSaveComm.Enabled = False
    cmdHSaveH.Enabled = False
    cmdSaveBio.Enabled = False
    cmdSaveCoag.Enabled = False
    cmdSaveImm(0).Enabled = False
    cmdSaveImm(1).Enabled = False
    cmdSaveBGa.Enabled = False
    Exit Sub
  
  Case 1: strDept = "Haem"
  Case 2: strDept = "Bio"
  Case 3: strDept = "Coag"
  Case 4: strDept = "End"
  Case 5: strDept = "Bga"
  Case 6: strDept = "Imm"
  Case 7: strDept = "Ext"
End Select

strDirection = IIf(Index = 0, "Desc", "Asc")
strArrow = IIf(Index = 0, "<", ">")

If lblResultOrRequest = "Results" Then
  SQL = "SELECT top 1 SampleID from " & strDept & "Results WHERE " & _
        "SampleID " & strArrow & " " & txtSampleID & " " & _
        "Order by SampleID " & strDirection
ElseIf sstabAll.Tab = 7 Then 'ext
  SQL = "SELECT top 1 SampleID from " & strDept & "Results WHERE " & _
        "SampleID " & strArrow & " " & txtSampleID & " " & _
        "Order by SampleID " & strDirection
ElseIf lblResultOrRequest = "Requests" Then
  SQL = "SELECT top 1 SampleID from " & strDept & "Requests WHERE " & _
        "SampleID " & strArrow & " " & txtSampleID & " " & _
        "Order by SampleID " & strDirection
Else
  SQL = "SELECT top 1 SampleID from " & strDept & "Results WHERE " & _
        "SampleID " & strArrow & " " & txtSampleID & " and valid <> 1 " & _
        "Order by SampleID " & strDirection
End If

If strDept = "Bio" Then
  If InStr(lblViewSplit, "Pri") Then
    strSplitSELECT = LoadSplitList(1)
  ElseIf InStr(lblViewSplit, "Sec") Then
    strSplitSELECT = LoadSplitList(2)
  End If
  If lblResultOrRequest = "Results" Then
    If strSplitSELECT <> "" Then
      SQL = "SELECT top 1 SampleID from BioResults WHERE " & _
            "SampleID " & strArrow & " " & txtSampleID & " " & _
            "and (" & strSplitSELECT & ") " & _
            "Order by SampleID " & strDirection
    End If
  Else
    If strSplitSELECT <> "" Then
      SQL = "SELECT top 1 SampleID from BioRequests WHERE " & _
            "SampleID " & strArrow & " " & txtSampleID & " " & _
            "and (" & strSplitSELECT & ") " & _
            "Order by SampleID " & strDirection
    End If
  End If
ElseIf strDept = "Imm" Then
  If InStr(lblImmViewSplit(1), "Pri") Then
    strSplitSELECT = LoadImmSplitList(1)
  ElseIf InStr(lblImmViewSplit(1), "Sec") Then
    strSplitSELECT = LoadImmSplitList(2)
  End If
  If lblResultOrRequest = "Results" Then
    If strSplitSELECT <> "" Then
      SQL = "SELECT top 1 SampleID from ImmResults WHERE " & _
            "SampleID " & strArrow & " " & txtSampleID & " " & _
            "and (" & strSplitSELECT & ") " & _
            "Order by SampleID " & strDirection
    End If
  Else
    If strSplitSELECT <> "" Then
      SQL = "SELECT top 1 SampleID from ImmRequests WHERE " & _
            "SampleID " & strArrow & " " & txtSampleID & " " & _
            "and (" & strSplitSELECT & ") " & _
            "Order by SampleID " & strDirection
    End If
  End If
ElseIf strDept = "End" Then
  If InStr(lblImmViewSplit(0), "Pri") Then
    strSplitSELECT = LoadImmSplitList(1)
  ElseIf InStr(lblImmViewSplit(0), "Sec") Then
    strSplitSELECT = LoadImmSplitList(0)
  End If
  If lblResultOrRequest = "Results" Then
    If strSplitSELECT <> "" Then
      SQL = "SELECT top 1 SampleID from endResults WHERE " & _
            "SampleID " & strArrow & " " & txtSampleID & " " & _
            "and (" & strSplitSELECT & ") " & _
            "Order by SampleID " & strDirection
    End If
  Else
    If strSplitSELECT <> "" Then
      SQL = "SELECT top 1 SampleID from endRequests WHERE " & _
            "SampleID " & strArrow & " " & txtSampleID & " " & _
            "and (" & strSplitSELECT & ") " & _
            "Order by SampleID " & strDirection
    End If
  End If

End If

Set tb = New Recordset
RecOpenServer 0, tb, SQL
If Not tb.EOF Then
  txtSampleID = tb!SampleID & ""
End If

LoadAllDetails

cmdSaveDemographics.Enabled = False
cmdSaveInc.Enabled = False
cmdSaveHaem.Enabled = False
cmdSaveComm.Enabled = False
cmdHSaveH.Enabled = False
cmdSaveBio.Enabled = False
cmdSaveCoag.Enabled = False
cmdSaveImm(0).Enabled = False
cmdSaveImm(1).Enabled = False
cmdSaveBGa.Enabled = False

Exit Sub

irelevant_Click_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /irelevant_Click ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub iRunDate_Click(Index As Integer)

On Error GoTo iRunDate_Click_Error

If Index = 0 Then
  dtRunDate = DateAdd("d", -1, dtRunDate)
Else
  If DateDiff("d", dtRunDate, Now) > 0 Then
    dtRunDate = DateAdd("d", 1, dtRunDate)
  End If
End If

cmdSaveInc.Enabled = True
cmdSaveDemographics.Enabled = True

Exit Sub

iRunDate_Click_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /iRunDate_Click ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub iSampleDate_Click(Index As Integer)

If Index = 0 Then
  dtSampleDate = DateAdd("d", -1, dtSampleDate)
Else
  If DateDiff("d", dtSampleDate, Now) > 0 Then
    dtSampleDate = DateAdd("d", 1, dtSampleDate)
  End If
End If

cmdSaveInc.Enabled = True
cmdSaveDemographics.Enabled = True

End Sub

Private Function IsControl(ByVal Chart As String) As Boolean

Dim n As Long

IsControl = False

If Trim(Chart) <> "" Then
For n = 0 To UBound(ControlName)
  If Trim(UCase(Chart)) = UCase(ControlName(n)) Then
    IsControl = True
    Exit For
  End If
Next
End If

End Function

Private Sub iToday_Click(Index As Integer)

On Error GoTo iToday_Click_Error

If Index = 0 Then
  dtRunDate = Format$(Now, "dd/mm/yyyy")
ElseIf Index = 1 Then
  If DateDiff("d", dtRunDate, Now) > 0 Then
    dtSampleDate = dtRunDate
  Else
    dtSampleDate = Format$(Now, "dd/mm/yyyy")
  End If
ElseIf Index = 2 Then
  If DateDiff("d", dtRunDate, Now) > 0 Then
    dtRecDate = dtRunDate
  Else
    dtRecDate = Format$(Now, "dd/mm/yyyy")
  End If
End If

cmdSaveInc.Enabled = True
cmdSaveDemographics.Enabled = True

Exit Sub

iToday_Click_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /iToday_Click ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub lblAss_Click()
Dim Num As Long
Dim Numx As Long

For Num = Len(lblAss) To 1 Step -1
    If Mid(lblAss, Num, 1) = " " Then
      Numx = Num
      Exit For
    End If
Next

txtSampleID = Trim(Mid(lblAss, Numx))
txtSampleID_LostFocus

End Sub

Private Sub lblChartNumber_Click()

With lblChartNumber
If InStr(.Caption, Hospname(0)) = 0 Then
      .BackColor = vbRed
      .ForeColor = vbYellow
Else
  .BackColor = &H8000000F
  .ForeColor = vbBlack
End If

End With

If Trim$(txtChart) <> "" Then
  LoadPatientFromChart frmEditAll, True
  cmdSaveDemographics.Enabled = True
  cmdSaveInc.Enabled = True
End If

End Sub

Private Sub lblImmViewSplit_Click(Index As Integer)
On Error GoTo lblImmViewSplit_Click_Error

If Index = 0 Then
  With lblImmViewSplit(0)
    Select Case .Caption
      Case "Viewing All":
        .Caption = "Viewing Primary Split"
        .BackColor = &H800080
        .ForeColor = &HFF00&
      Case "Viewing Primary Split":
        .Caption = "Viewing Secondary Split"
        .BackColor = &H800080
        .ForeColor = &HFF00&
      Case "Viewing Secondary Split":
        .Caption = "Viewing All"
        .BackColor = &H8000000F
        .ForeColor = vbBlack
    End Select
  End With
Else
  With lblImmViewSplit(1)
    Select Case .Caption
      Case "Viewing All":
        .Caption = "Viewing Primary Split"
        .BackColor = &H800080
        .ForeColor = &HFF00&
      Case "Viewing Primary Split":
        .Caption = "Viewing Secondary Split"
        .BackColor = &H800080
        .ForeColor = &HFF00&
      Case "Viewing Secondary Split":
        .Caption = "Viewing All"
        .BackColor = &H8000000F
        .ForeColor = vbBlack
    End Select
  End With
End If

Exit Sub

lblImmViewSplit_Click_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /lblImmViewSplit_Click ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub lblMalaria_Change()
cmdSaveHaem.Enabled = True: cmdHSaveH.Enabled = True

If Trim$(lblMalaria) <> "" Then
  chkMalaria = 1
Else
  chkMalaria = 0
End If

End Sub

Private Sub lblMalaria_Click()

If lblMalaria = "" Then
    lblMalaria = "Positive"
ElseIf lblMalaria = "Positive" Then
    lblMalaria = "Negative"
ElseIf lblMalaria = "Negative" Then
    lblMalaria = "Inconclusive"
ElseIf lblMalaria = "Inconclusive" Then
    lblMalaria = ""
End If

End Sub

Private Sub lblResultOrRequest_Click()

If sstabAll.Tab <> 0 Then
  If lblResultOrRequest = "Results" Then
    lblResultOrRequest = "Request"
  ElseIf lblResultOrRequest = "Request" Then
    lblResultOrRequest = "UnValid"
  Else
    lblResultOrRequest = "Results"
  End If
End If

End Sub

Private Sub lblSickledex_Change()

cmdSaveHaem.Enabled = True: cmdHSaveH.Enabled = True

If Trim$(lblSickledex) <> "" Then
  chkSickledex = 1
Else
  chkSickledex = 0
End If


End Sub

Private Sub lblSickledex_Click()

If lblSickledex = "" Then
    lblSickledex = "Positive"
ElseIf lblSickledex = "Positive" Then
    lblSickledex = "Negative"
ElseIf lblSickledex = "Negative" Then
    lblSickledex = "Inconclusive"
ElseIf lblSickledex = "Inconclusive" Then
    lblSickledex = ""
End If

End Sub

Private Sub lblUrgent_Click()

lblUrgent.Visible = False

Cnxn(0).Execute "UPDATE demographics set urgent = 0 WHERE sampleid = '" & txtSampleID & "'"

End Sub

Private Sub lblViewSplit_Click()

On Error GoTo lblViewSplit_Click_Error

With lblViewSplit
  Select Case .Caption
    Case "Viewing All":
      .Caption = "Viewing Primary Split"
      .BackColor = &H800080
      .ForeColor = &HFF00&
    Case "Viewing Primary Split":
      .Caption = "Viewing Secondary Split"
      .BackColor = &H800080
      .ForeColor = &HFF00&
    Case "Viewing Secondary Split":
      .Caption = "Viewing All"
      .BackColor = &H8000000F
      .ForeColor = vbBlack
  End Select
End With

Exit Sub

lblViewSplit_Click_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /lblViewSplit_Click ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub lHaemErrors_Click()

Unload frmHaemErrors

With frmHaemErrors
  .Analyser = HaemAnalyser
  .ErrorNumber = lHaemErrors.Tag
  .Show 1
End With

End Sub

Private Sub lImmRan_Click(Index As Integer)

If Index = 0 Then
  If lImmRan(0) = "Random Sample" Then
    lImmRan(0) = "Fasting Sample"
  Else
    lImmRan(0) = "Random Sample"
  End If
  
  LoadEndocrinology
  
  cmdSaveImm(0).Enabled = True
Else
  If lImmRan(1) = "Random Sample" Then
    lImmRan(1) = "Fasting Sample"
  Else
    lImmRan(1) = "Random Sample"
  End If
  
  LoadImmunology
  
  cmdSaveImm(1).Enabled = True
End If

End Sub

Private Sub LoadAllDetails()

On Error GoTo LoadAllDetails_Error

HaemLoaded = False
BioLoaded = False
CoagLoaded = False
ImmLoaded = False
BgaLoaded = False
ExtLoaded = False
EndLoaded = False

cAdd = ""
cUnits = ""
tnewvalue = ""

cIAdd(0) = ""
cIUnits(0) = ""
tINewValue(0) = ""

cIAdd(1) = ""
cIUnits(1) = ""
tINewValue(1) = ""

ClearDemographics
ClearHaematologyResults
ClearCoagulation
ClearOutstandingImm
ClearOutstandingBio
ClearImmFlags
ClearEndFlags
'ClearBga
ClearExt

sstabAll.TabCaption(1) = "Haematology"
sstabAll.TabCaption(2) = "Biochemistry"
sstabAll.TabCaption(3) = "Coagulation"
sstabAll.TabCaption(4) = "Endocrinology"
sstabAll.TabCaption(5) = "Blood Gas"
sstabAll.TabCaption(6) = "Immunology"
sstabAll.TabCaption(7) = "Externals"

LoadDemographics
CheckDepartments
LoadComments
CheckIfPhoned

Select Case sstabAll.Tab
  Case 0:
  Case 1: LoadHaematology
          HaemLoaded = True
  Case 2: LoadBiochemistry
          BioLoaded = True
  Case 3: LoadCoagulation
          CoagLoaded = True
  Case 4: LoadEndocrinology
          EndLoaded = True
  Case 5: LoadBloodGas
          BgaLoaded = True
  Case 6: LoadImmunology
          ImmLoaded = True
  Case 7: LoadExt
          ExtLoaded = True
End Select


SetViewHistory

'cmdSaveHaem.Enabled = False
'cmdSaveComm.Enabled = False
'cmdHSaveH.Enabled = False
'cmdSaveBio.Enabled = False
'cmdSaveImm(0).Enabled = False
'cmdSaveImm(1).Enabled = False

Exit Sub

LoadAllDetails_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /LoadAllDetails ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Public Sub LoadBiochemistry()

Dim DeltaSn As Recordset
Dim Deltatb As Recordset
Dim tb As New Recordset
Dim SQL As String
Dim s As String
Dim Value As Single
Dim OldValue As Single
Dim valu As String
Dim PreviousDate As String
Dim PreviousRec As String
Dim Res As String
Dim n As Long
Dim e As String
Dim DeltaLimit As Single
Dim SampleType As String
Dim BRs As New BIEResults
Dim BRres As BIEResults
Dim br As BIEResult
Dim Fasting As Boolean
Dim Flag As String
Dim T As String
Dim Code As String
Dim CodeTb As Recordset
Dim sn As New Recordset


On Error GoTo LoadBiochemistry_Error

If txtSampleID = "" Then Exit Sub

Frame2.Enabled = True
lRandom.Enabled = True
'txtBioComment.Locked = False

Fasting = lRandom = "Fasting Sample"
lblAss.Visible = False

ClearFGrid gBio

oH = 0
oS = 0
oL = 0
oO = 0
oG = 0
oJ = 0
lBDate = ""
ldelta = ""
bViewBioRepeat.Visible = False

sstabAll.TabCaption(2) = "Biochemistry"

'get date & run number of previous record
PreviousBio = False
HistBio = False


If txtName <> "" And txtDoB <> "" Then
  SQL = CreateHist("bio")
    Set sn = New Recordset
    RecOpenServer 0, sn, SQL
    If Not sn.EOF Then
        HistBio = True
    End If
  
          
    SQL = CreateSql("Bio")
    Set Deltatb = New Recordset
    RecOpenServer 0, Deltatb, SQL
    If Not Deltatb.EOF Then
      PreviousDate = Deltatb!Rundate & ""
      PreviousRec = DeencryptN(Deltatb!SampleID & "")
      PreviousBio = True
  End If
End If

gBio.Visible = False

Set BRres = BRs.Load("Bio", txtSampleID, "Results", gDONTCARE, gDONTCARE, 0, cCat(0), dtRunDate)

CheckCalcPSA BRres

If Not BRres Is Nothing Then
  If SysOptDoAssGlucose(0) Then
    CheckAssGlucose BRres
  End If
  CheckCalcPSA BRres
  If SysOptCheckCholHDLRatio(0) Then CheckCholHDL BRres
End If


If IsControl(txtChart) Then
  ldelta = ""
  
  gBio.Rows = 2
  gBio.AddItem ""
  gBio.RemoveItem 1
If Not BRres Is Nothing Then
  For Each br In BRres
    Code = br.Code
    
    If SysOptBioAn1(0) = "ROCHE" Then
      If Code = 91 Then
        If Val(br.Result) > 200 Then
          oG.Value = True
        ElseIf Val(br.Result) > 80 Then
          oH.Value = True
        ElseIf Val(br.Result) > 30 Then
          oS.Value = True
        End If
      End If
    End If
    s = br.LongName & vbTab
    If Not IsNull(br.Result) Then
      Value = br.Result
    Else
      Value = 0
    End If
    If Value <= 1 Then
      valu = Format(Value, "0.00")
    ElseIf Value > 1 And Value <= 10 Then
      valu = Format(Value, "0.0")
    Else
      valu = Format(Value)
    End If
    s = s & valu & vbTab
  SQL = "SELECT * from controls WHERE controlname = '" & txtChart & "' and parameter = '" & Code & "'"
  Set CodeTb = New Recordset
  RecOpenServer 0, CodeTb, SQL
    If Not CodeTb.EOF Then
      If Not IsNull(CodeTb!mean) And Not IsNull(CodeTb("1sd")) Then
        s = s & InterC(Value, CodeTb!mean - CodeTb("1sd") * 2, CodeTb!mean + CodeTb("1sd") * 2) & vbTab
        s = s & (CodeTb!mean - CodeTb("1sd") * 2) & "  -  " & (CodeTb!mean + CodeTb("1sd") * 2) & vbTab & vbTab & vbTab & vbTab
      End If
    End If
    s = s & br.Pc & vbTab
    Select Case Trim(br.Analyser)
    Case "4": s = s & "Immuno"
    Case "A": s = s & "Bio (A)"
    Case "B": s = s & "Bio (B)"
    Case "P1": s = s & SysOptBioN1(0)
    Case "P2": s = s & SysOptBioN2(0)
    Case Else: s = s & "General"
  End Select
    s = s & vbTab & br.Comment
    gBio.AddItem s
  gBio.Refresh
  Next
  End If
  gBio.Visible = True
  If gBio.Rows > 2 Then
    gBio.RemoveItem 1
  End If
  Exit Sub
End If




If Not BRres Is Nothing Then
  sstabAll.TabCaption(2) = ">>Biochemistry<<"
  For Each br In BRres
    Flag = ""
    SampleType = br.SampleType
    cSampleType = ListText("ST", br.SampleType)
    If Len(SampleType) = 0 Then SampleType = "S"
    s = br.ShortName & vbTab
    lBDate = br.RunTime
    If IsNumeric(br.Result) Then
      Value = Val(br.Result)
      Select Case br.Printformat
        Case 0: valu = Format$(Value, "0")
        Case 1: valu = Format$(Value, "0.0")
        Case 2: valu = Format$(Value, "0.00")
        Case 3: valu = Format$(Value, "0.000")
        Case Else: valu = Format$(Value, "0.000")
      End Select
    Else
      valu = br.Result
    End If
 '   If UserMemberOf = "The World" And Not BR.Valid Then
 '     s = s & "" & vbTab
 '   Else
    s = s & valu & vbTab
    If ListText("UN", br.Units) <> "" Then
      s = s & ListText("UN", br.Units)
    Else
      s = s & br.Units
    End If
    s = s & vbTab
    T = ""
    If IsNumeric(br.Result) Then
      If Value > Val(br.PlausibleHigh) Then
        Flag = "X"
        s = s & br.Low & " - " & br.High & vbTab
        s = s & "X"
      ElseIf Value < Val(br.PlausibleLow) Then
        Flag = "X"
        s = s & br.Low & " - " & br.High & vbTab
        s = s & "X"
      ElseIf br.Code = SysOptBioCodeForGlucose(0) Or _
             br.Code = SysOptBioCodeForChol(0) Or _
             br.Code = SysOptBioCodeForTrig(0) Then
        If Fasting Then
          If br.Code = SysOptBioCodeForGlucose(0) Then
            SQL = "SELECT * from fastings WHERE testname = '" & "GLU" & "'"
          ElseIf br.Code = SysOptBioCodeForChol(0) Then
            SQL = "SELECT * from fastings WHERE testname = '" & "CHO" & "'"
          ElseIf br.Code = SysOptBioCodeForTrig(0) Then
            SQL = "SELECT * from fastings WHERE testname = '" & "TRI" & "'"
          End If
          Set tb = New Recordset
          RecOpenServer 0, tb, SQL
          
          If Not tb.EOF Then
            s = s & tb!FastingText & vbTab
            If Value > tb!FastingHigh Then
                Flag = "H"
              s = s & "H"
            ElseIf Value < tb!FastingLow Then
              Flag = "L"
              s = s & "L"
            End If
          Else
          End If
        Else
          s = s & br.Low & " - " & br.High & vbTab
          If Value < Val(br.FlagLow) Then
            Flag = "FL"
            T = "FL"
          ElseIf Value > Val(br.FlagHigh) Then
            Flag = "FH"
            T = "FH"
          End If
          If Value < Val(br.Low) Then
            Flag = "L"
            T = "L"
          ElseIf Value > Val(br.High) Then
            Flag = "H"
            T = "H"
          End If
        End If
      Else
        s = s & br.Low & " - " & br.High & vbTab
        If Value < Val(br.FlagLow) Then
          Flag = "FL"
          T = "FL"
        ElseIf Value > Val(br.FlagHigh) Then
          Flag = "FH"
          T = "FH"
        End If
        If Value < Val(br.Low) Then
          Flag = "L"
          T = "L"
        ElseIf Value > Val(br.High) Then
          Flag = "H"
          T = "H"
        End If
      End If
    Else
        s = s & br.Low & " - " & br.High & vbTab
    End If
    e = ""
    e = Trim(br.Flags & "")
    s = s & T
    s = s & vbTab & _
            IIf(e <> "", e, "") & vbTab & _
            IIf(br.Valid, "V", " ") & _
            IIf(br.Printed, "P", " ") & vbTab
    If br.Valid = True Then
      Frame2.Enabled = False
      lRandom.Enabled = False
 '     txtBioComment.Locked = True
    End If
    s = s & br.Pc & vbTab
    Select Case Trim(br.Analyser)
    Case "4": s = s & "Immuno"
    Case "A": s = s & "Bio (A)"
    Case "B": s = s & "Bio (B)"
    Case "P1": s = s & SysOptBioN1(0)
    Case "P2": s = s & SysOptBioN2(0)
    Case Else: s = s & "General"
  End Select
    s = s & vbTab & br.Comment
    gBio.AddItem s
    
    If Flag <> "" Then
      gBio.Row = gBio.Rows - 1
      gBio.Col = 1
      Select Case Flag
        Case "H":
        If Hospname(0) = "STJOHNS" Then
            For n = 1 To 2
            gBio.Col = n
            gBio.CellBackColor = SysOptHighBack(0)
            gBio.CellForeColor = SysOptHighFore(0)
          Next
        Else
            For n = 0 To 9
            gBio.Col = n
            gBio.CellBackColor = SysOptHighBack(0)
            gBio.CellForeColor = SysOptHighFore(0)
          Next
        End If
        Case "L":
        If Hospname(0) = "STJOHNS" Then
            For n = 1 To 2
            gBio.Col = n
            gBio.CellBackColor = SysOptLowBack(0)
            gBio.CellForeColor = SysOptLowFore(0)
          Next
        Else
            For n = 0 To 9
            gBio.Col = n
            gBio.CellBackColor = SysOptLowBack(0)
            gBio.CellForeColor = SysOptLowFore(0)
          Next
        End If
        Case "X":
        If Hospname(0) = "STJOHNS" Then
            For n = 1 To 2
            gBio.Col = n
            gBio.CellBackColor = SysOptPlasBack(0)
            gBio.CellForeColor = SysOptPlasFore(0)
          Next
        Else
            For n = 0 To 9
            gBio.Col = n
            gBio.CellBackColor = SysOptPlasBack(0)
            gBio.CellForeColor = SysOptPlasFore(0)
          Next
        End If
      End Select
    End If

    If br.DoDelta And PreviousBio Then
      SQL = "SELECT * from bioresults WHERE " & _
            "sampleid = '" & PreviousRec & "' " & _
            "and code = '" & br.Code & "'"
      Set DeltaSn = New Recordset
      RecOpenServer 0, DeltaSn, SQL
      If Not DeltaSn.EOF Then
        OldValue = Val(DeltaSn!Result)
        If OldValue <> 0 Then
          DeltaLimit = br.DeltaLimit
          If Abs(OldValue - Value) > DeltaLimit Then
            Res = Format$(PreviousDate, "dd/mm/yyyy") & " (" & PreviousRec & ") " & _
                  br.ShortName & " " & _
                  OldValue & vbCr
            ldelta = ldelta & Res
          End If
        End If
      End If
    End If
  Next
End If
  
FixG gBio
  
With gBio
  bValidateBio.Caption = "VALID"
  lblUrgent.Visible = False
  For n = 1 To .Rows - 1
    If .TextMatrix(n, 3) = "X" Then
      .Row = n
      .Col = 1
      .CellForeColor = vbWhite
      .CellBackColor = vbBlack
    End If
    If InStr(.TextMatrix(n, 6), "V") = "0" Then
        bValidateBio.Caption = "&Validate"
        lblUrgent.Visible = UrgentTest
    End If
  Next
End With


LoadOutstandingBio

SQL = "SELECT * from BioRepeats WHERE " & _
      "SampleID = '" & txtSampleID & "'"
Set tb = New Recordset
RecOpenServer 0, tb, SQL
bViewBioRepeat.Visible = False
If Not tb.EOF Then
  bViewBioRepeat.Visible = True
End If

SQL = "SELECT * from Masks WHERE " & _
      "SampleID = '" & txtSampleID & "'"
Set tb = New Recordset
RecOpenServer 0, tb, SQL
If Not tb.EOF Then
  oH = IIf(tb!h, 1, 0)
  oS = IIf(tb!s, 1, 0)
  oL = IIf(tb!l, 1, 0)
  oO = IIf(tb!o, 1, 0)
  oG = IIf(tb!g, 1, 0)
  oJ = IIf(tb!J, 1, 0)
End If



Exit Sub

LoadBiochemistry_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /LoadBiochemistry ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Public Sub LoadBloodGas()



Dim DeltaSn As Recordset
Dim Deltatb As Recordset
Dim tb As New Recordset
Dim SQL As String
Dim s As String
Dim Value As Single
Dim OldValue As Single
Dim valu As String
Dim PreviousDate As String
Dim PreviousRec As String
Dim Res As String
Dim n As Long
Dim e As String
Dim DeltaLimit As Single
Dim SampleType As String
Dim BRs As New BIEResults
Dim BRres As BIEResults
Dim br As BIEResult
Dim Flag As String
Dim T As String
Dim Code As String
Dim sn As New Recordset



On Error GoTo LoadBloodGas_Error

If txtSampleID = "" Then Exit Sub

lblBgaDate = ""
lBgaDelta = ""

ClearFGrid gBga

bViewBgaRepeat.Visible = False

sstabAll.TabCaption(5) = "Blood Gas"

'get date & run number of previous record
PreviousBga = False
HistBga = False


If txtName <> "" And txtDoB <> "" Then
  SQL = CreateHist("bga")
    Set sn = New Recordset
    RecOpenServer 0, sn, SQL
    If Not sn.EOF Then
        HistBga = True
    End If
  
          
    SQL = CreateSql("Bga")
    Set Deltatb = New Recordset
    RecOpenServer 0, Deltatb, SQL
    If Not Deltatb.EOF Then
      PreviousDate = Deltatb!Rundate & ""
      PreviousRec = DeencryptN(Deltatb!SampleID & "")
      PreviousBga = True
  End If
End If

Set BRres = BRs.Load("Bga", txtSampleID, "Results", gDONTCARE, gDONTCARE, 0, cCat(0), dtRunDate)
  
With gBga
  .Rows = 2
  .AddItem ""
  .RemoveItem 1
End With

If Not BRres Is Nothing Then
  sstabAll.TabCaption(5) = ">>Blood Gas<<"
  For Each br In BRres
    lblBgaDate = Format(br.RunTime, "dd/MMM/yyyy hh:mm")
    Flag = ""
    SampleType = br.SampleType
    If Len(SampleType) = 0 Then SampleType = "S"
    s = br.ShortName & vbTab
    lBDate = br.RunTime
    If IsNumeric(br.Result) Then
      Value = Val(br.Result)
      Select Case br.Printformat
        Case 0: valu = Format$(Value, "0")
        Case 1: valu = Format$(Value, "0.0")
        Case 2: valu = Format$(Value, "0.00")
        Case 3: valu = Format$(Value, "0.000")
        Case Else: valu = Format$(Value, "0.000")
      End Select
    Else
      valu = br.Result
    End If
    s = s & valu & vbTab
    If ListText("UN", br.Units) <> "" Then
      s = s & ListText("UN", br.Units)
    Else
      s = s & br.Units
    End If
    s = s & vbTab
    s = s & br.Low & " - " & br.High & vbTab
    T = ""
    If IsNumeric(br.Result) Then
      If Value > Val(br.PlausibleHigh) Then
        Flag = "X"
        s = s & "X"
      ElseIf Value < Val(br.PlausibleLow) Then
        Flag = "X"
        s = s & "X"
      Else
        If Value < Val(br.FlagLow) Then
          Flag = "FL"
          T = "FL"
        ElseIf Value > Val(br.FlagHigh) Then
          Flag = "FH"
          T = "FH"
        End If
        If Value < Val(br.Low) Then
          Flag = "L"
          T = "L"
        ElseIf Value > Val(br.High) Then
          Flag = "H"
          T = "H"
        End If
      End If
    End If
    e = ""
    e = Trim(br.Flags & "")
    s = s & T
    s = s & vbTab & _
            IIf(br.Valid, "V", " ") & vbTab & _
            IIf(br.Printed, "P", " ") & vbTab
    gBga.AddItem s
    
    If Flag <> "" Then
      gBga.Row = gBga.Rows - 1
      gBga.Col = 1
      Select Case Flag
        Case "H":
        If Hospname(0) = "STJOHNS" Then
            For n = 1 To 2
            gBga.Col = n
            gBga.CellBackColor = SysOptHighBack(0)
            gBga.CellForeColor = SysOptHighFore(0)
          Next
        Else
            For n = 0 To 7
            gBga.Col = n
            gBga.CellBackColor = SysOptHighBack(0)
            gBga.CellForeColor = SysOptHighFore(0)
          Next
        End If
        Case "L":
        If Hospname(0) = "STJOHNS" Then
            For n = 1 To 2
            gBga.Col = n
            gBga.CellBackColor = SysOptLowBack(0)
            gBga.CellForeColor = SysOptLowFore(0)
          Next
        Else
            For n = 0 To 7
            gBga.Col = n
            gBga.CellBackColor = SysOptLowBack(0)
            gBga.CellForeColor = SysOptLowFore(0)
          Next
        End If
        Case "X":
        If Hospname(0) = "STJOHNS" Then
            For n = 1 To 2
            gBga.Col = n
            gBga.CellBackColor = SysOptPlasBack(0)
            gBga.CellForeColor = SysOptPlasFore(0)
          Next
        Else
            For n = 0 To 7
            gBga.Col = n
            gBga.CellBackColor = SysOptPlasBack(0)
            gBga.CellForeColor = SysOptPlasFore(0)
          Next
        End If
      End Select
    End If

    If br.DoDelta And PreviousBga Then
      SQL = "SELECT * from bgaresults WHERE " & _
            "sampleid = '" & PreviousRec & "' " & _
            "and code = '" & br.Code & "'"
      Set DeltaSn = New Recordset
      RecOpenClient 0, DeltaSn, SQL
      If Not DeltaSn.EOF Then
        OldValue = Val(DeltaSn!Result)
        If OldValue <> 0 Then
          DeltaLimit = br.DeltaLimit
          If Abs(OldValue - Value) > DeltaLimit Then
            Res = Format$(PreviousDate, "dd/mm/yyyy") & " (" & PreviousRec & ") " & _
                  br.ShortName & " " & _
                  OldValue & vbCr
            lBgaDelta = lBgaDelta & Res
          End If
        End If
      End If
    End If
  Next
End If
  
FixG gBga
  
With gBga
  cmdValBG.Caption = "VALID"
  lblUrgent.Visible = False
  For n = 1 To .Rows - 1
    If .TextMatrix(n, 3) = "X" Then
      .Row = n
      .Col = 1
      .CellForeColor = vbWhite
      .CellBackColor = vbBlack
    End If
    If InStr(.TextMatrix(n, 5), "V") = "0" Then
        cmdValBG.Caption = "&Validate"
        lblUrgent.Visible = UrgentTest
    End If
  Next
End With



SQL = "SELECT * from BgaRepeats WHERE " & _
      "SampleID = '" & txtSampleID & "'"
Set tb = New Recordset
RecOpenServer 0, tb, SQL
bViewBgaRepeat.Visible = False
If Not tb.EOF Then
  bViewBgaRepeat.Visible = True
End If



Exit Sub

LoadBloodGas_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /LoadBloodGas ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select



End Sub

Public Sub LoadCoagulation()
Dim CodeTb As Recordset
Dim CRs As New CoagResults
Dim cRR As New CoagResults
Dim CR As CoagResult
Dim s As String
Dim n As Long
Dim X As Long
Dim SQL As String
Dim tb As New Recordset
Dim g As String
Dim sex As String
Dim sn As New Recordset

On Error GoTo LoadCoagulation_Error

If txtSampleID = "" Then Exit Sub

ClearFGrid grdCoag

cmdValidateCoag.Caption = "&Validate"
'txtCoagComment.Locked = False

HistCoag = False

SQL = CreateHist("coag")
  Set sn = New Recordset
  RecOpenServer 0, sn, SQL
  If Not sn.EOF Then
      HistCoag = True
  End If



Set CRs = CRs.Load(txtSampleID, gDONTCARE, gDONTCARE, Trim(SysOptExp(0)), 0)
Set cRR = cRR.LoadRepeats(txtSampleID, gDONTCARE, gDONTCARE, Trim(SysOptExp(0)))

ClearCoagulation

sstabAll.TabCaption(3) = "Coagulation"

SQL = "SELECT * from demographics WHERE sampleid = '" & txtSampleID & "'"
Set tb = New Recordset
RecOpenServer 0, tb, SQL
If Not tb.EOF Then
  sex = tb!sex & ""
End If

SQL = "SELECT * from coagresults WHERE sampleid = '" & txtSampleID & "'"
Set tb = New Recordset
RecOpenServer 0, tb, SQL
If Not tb.EOF Then
  sstabAll.TabCaption(3) = ">>Coagulation<<"
  If tb!Valid = True Then cmdValidateCoag.Caption = "VALID"
End If

'LoadComments


For Each CR In CRs
  lCDate = CR.Rundate & " " & Format(CR.RunTime, "hh:mm")
  SQL = "SELECT * from coagtestdefinitions WHERE code = '" & Trim(CR.Code) & "' " & _
       "and agefromdays = '0' and agetodays > '43819'"
  Set tb = New Recordset
  RecOpenServer 0, tb, SQL
  If Not tb.EOF Then
    If tb!InUse = True Then
      If Trim(CR.Units) = "INR" Then
        s = "INR" & vbTab
      Else
        s = CoagNameFor(CR.Code) & vbTab
      End If
      If UserMemberOf = "The World" And Not CR.Valid Then
        s = s & vbTab
      Else
      Select Case CoagPrintFormat(Trim(CR.Code) & "", Trim(CR.Units) & "")
        Case 0: g = Format$(CR.Result, "0")
        Case 1: g = Format$(CR.Result, "0.0")
        Case 2: g = Format$(CR.Result, "0.00")
      End Select
        If g = "0" Or g = "0.0" Or g = "0.00" Then g = "Check"
        s = s & g & vbTab & _
        UnitConv(CR.Units) & vbTab
          SQL = "SELECT * from coagcontrols WHERE controlname = '" & txtChart & "' and parameter = '" & CR.Code & "'"
          Set CodeTb = New Recordset
          RecOpenServer 0, CodeTb, SQL
            If Not CodeTb.EOF Then
              If Not IsNull(CodeTb!mean) And Not IsNull(CodeTb("1sd")) Then
                s = s & (CodeTb!mean - CodeTb("1sd") * 2) & " - " & (CodeTb!mean + CodeTb("1sd") * 2) & vbTab
                s = s & InterC(CR.Result, CodeTb!mean - CodeTb("1sd") * 2, CodeTb!mean + CodeTb("1sd") * 2) & vbTab
              Else
                s = s & vbTab & vbTab
              End If
          s = s & IIf(CR.Valid, "V", "") & vbTab & _
          IIf(CR.Printed, "P", "")
        Else
          If Trim(UCase(CR.Units)) = "INR" Then
            s = s & vbTab
          Else
          If sex = "M" Then
            s = s & Trim(tb!MaleLow) & " - " & Trim(tb!MaleHigh) & vbTab
          ElseIf sex = "F" Then
            s = s & Trim(tb!FemaleLow) & " - " & Trim(tb!FemaleHigh) & vbTab
          Else
            s = s & Trim(tb!FemaleLow) & " - " & Trim(tb!MaleHigh) & vbTab
          End If
          End If
          If Trim(UCase(CR.Units)) = "INR" Then
            s = s & vbTab
          Else
             s = s & InterpCoag(sex, CR.Code, CR.Result, CR.Units) & vbTab
          End If
          s = s & IIf(CR.Valid, "V", "") & vbTab & _
          IIf(CR.Printed, "P", "")
        End If
        grdCoag.AddItem s
      End If
    End If
  ElseIf Trim(CR.Units) = "INR" Then
      s = CoagNameFor(CR.Code, CR.Units) & vbTab
      If UserMemberOf = "The World" And Not CR.Valid Then
        s = s & vbTab
      Else
      Select Case CoagPrintFormat(Trim(CR.Code) & "", Trim(CR.Units) & "")
        Case 0: g = Format$(CR.Result, "0")
        Case 1: g = Format$(CR.Result, "0.0")
        Case 2: g = Format$(CR.Result, "0.00")
      End Select
      End If
        s = s & g & vbTab & _
        UnitConv(CR.Units) & vbTab & vbTab
        s = s & vbTab
        s = s & IIf(CR.Valid, "V", "") & vbTab & _
        IIf(CR.Printed, "P", "")
        grdCoag.AddItem s
  End If
Next
  
  
FixG grdCoag

If grdCoag.Rows > 2 Then
  sstabAll.TabCaption(3) = ">>Coagulation<<"
End If
 
If txtCoagComment <> "" Then
  sstabAll.TabCaption(3) = ">>Coagulation<<"
End If
 
With grdCoag
  If grdCoag.TextMatrix(1, 0) <> "" Then
    cmdValidateCoag.Caption = "VALID"
    'txtCoagComment.Locked = True
    lblUrgent.Visible = False
    For n = 1 To .Rows - 1
      If .TextMatrix(n, 0) <> "" Then
        If InStr(.TextMatrix(n, 5), "V") = "0" Then
            cmdValidateCoag.Caption = "&Validate"
            lblUrgent.Visible = UrgentTest
 '           txtCoagComment.Locked = False
        End If
      End If
    Next
  End If
End With

  
For n = 1 To grdCoag.Rows - 1
  Select Case Left(grdCoag.TextMatrix(n, 4), 1)
    Case "H":
      grdCoag.Row = n
      For X = 0 To grdCoag.Cols - 1
      grdCoag.Col = X
      grdCoag.CellBackColor = vbRed
      grdCoag.CellForeColor = vbYellow
      Next
    Case "L":
      grdCoag.Row = n
      For X = 0 To grdCoag.Cols - 1
      grdCoag.Col = X
      grdCoag.CellBackColor = vbBlue
      grdCoag.CellForeColor = vbYellow
      Next
    Case Else
      grdCoag.Row = n
      For X = 0 To grdCoag.Cols - 1
      grdCoag.Col = X
      grdCoag.CellBackColor = vbWhite
      grdCoag.CellForeColor = vbBlack
      Next
  End Select
  If grdCoag.TextMatrix(n, 1) = "Check" Then
      grdCoag.Row = n
      For X = 0 To grdCoag.Cols - 1
      grdCoag.Col = X
      grdCoag.CellBackColor = vbBlue
      grdCoag.CellForeColor = vbYellow
      Next
   End If
Next
  
  
  
bViewCoagRepeat.Visible = cRR.Count <> 0


LoadOutstandingrdCoag

If SysOptDontShowPrevCoag(0) = True Then
   grdPrev.Visible = False
   lblPrevCoag.Visible = False
Else
  LoadPreviousCoag
End If

Exit Sub

LoadCoagulation_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /LoadCoagulation ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub LoadComments()

Dim Cx As New Comment
Dim Cxs As New Comments
  
'On Error Resume Next

On Error GoTo LoadComments_Error

txtBioComment = ""
txtHaemComment = ""
txtDemographicComment = ""
lblDemographicComment = ""
txtCoagComment = ""
txtImmComment(0) = ""
txtImmComment(1) = ""
txtBGaComment = ""
If Trim$(txtSampleID) = "" Then Exit Sub

Set Cx = Cxs.Load(0, txtSampleID)
If Not Cx Is Nothing Then
  txtBioComment = Split_Comm(Cx.Biochemistry)
  txtHaemComment = Split_Comm(Cx.Haematology)
  txtDemographicComment = Split_Comm(Cx.Demographics)
  lblDemographicComment = txtDemographicComment
  txtCoagComment = Split_Comm(Cx.Coagulation)
  txtImmComment(1) = Split_Comm(Cx.Immunology)
  txtImmComment(0) = Split_Comm(Cx.Endocrinology)
  txtBGaComment = Split_Comm(Cx.BloodGas)
End If

Exit Sub

LoadComments_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /LoadComments ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub LoadDemo(ByVal IDNumber As String)

Dim tb As New Recordset
Dim SQL As String
Dim IDType As String
Dim n As Long


On Error GoTo LoadDemo_Error

IDType = CheckDemographics(IDNumber)
If IDType = "" Then
   'clearpatient
   Exit Sub
End If

'Rem Code Change 16/01/2006
SQL = "SELECT * from patientifs WHERE " & _
   IDType & " = '" & AddTicks(IDNumber) & "' "

Set tb = New Recordset
RecOpenServer 0, tb, SQL
If tb.EOF = True Then
'   clearpatient
Else
   If Trim(tb!Chart & "") = "" Then txtChart = tb!mrn & "" Else txtChart = tb!Chart & ""
   txtAandE = tb!aande & ""
   txtNOPAS = tb!NOPAS & ""
   n = InStr(tb!PatName & "", "''")
   If n <> 0 Then
     tb!PatName = Left$(tb!PatName, n) & Mid$(tb!PatName, n + 2)
     tb.Update
   End If
   txtName = initial2upper(tb!PatName & "")
   If Not IsNull(tb!DoB) Then
      lDoB = Format(tb!DoB, "DD/MM/YYYY")
      txtDoB = Format(tb!DoB, "DD/MM/YYYY")
   Else
      lDoB = ""
      txtDoB = ""
   End If
   lAge = CalcAge(tb!DoB & "")
   txtAge = lAge
   Select Case tb!sex & ""
      Case "M": lSex = "Male"
      Case "F": lSex = "Female"
      Case Else: lSex = ""
   End Select
   txtSex = lSex
   n = InStr(tb!Address0 & "", "''")
   If n <> 0 Then
     tb!Address0 = Left$(tb!Address0, n) & Mid$(tb!Address0, n + 2)
     tb.Update
   End If
   
   taddress(0) = initial2upper(Trim(tb!Address0 & ""))
   taddress(1) = initial2upper(Trim(tb!Address1 & ""))
   cmbWard.Text = initial2upper(tb!Ward & "")
   cmbClinician.Text = initial2upper(tb!Clinician & "")
End If
tb.Close





Exit Sub

LoadDemo_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /LoadDemo ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Public Sub LoadDemographics()

Dim SQL As String
Dim tb As New Recordset
Dim SampleDate As String
Dim RooH As Boolean

On Error GoTo LoadDemographics_Error

UrgentTest = False
RooH = IsRoutine()
cRooH(0) = RooH
cRooH(1) = Not RooH
bViewBB.Enabled = False
txtAge = ""
lAge = ""
If Trim$(txtSampleID) = "" Then Exit Sub
  
lRandom = "Random Sample"

Screen.MousePointer = 11

SQL = "SELECT * from Demographics WHERE " & _
      "SampleID = '" & EncryptN(txtSampleID) & "'"
      
Set tb = New Recordset
RecOpenServer 0, tb, SQL
If Not tb.EOF Then
  If Trim$(tb!Hospital & "") <> "" Then
     lblChartNumber = Trim$(UCase(tb!Hospital)) & " Chart #"
    If UCase(tb!Hospital) = Hospname(0) Then
      lblChartNumber.BackColor = &H8000000F
      lblChartNumber.ForeColor = vbBlack
    Else
      lblChartNumber.BackColor = vbRed
      lblChartNumber.ForeColor = vbYellow
    End If
  Else
    lblChartNumber.Caption = Hospname(0) & " Chart #"
    lblChartNumber.BackColor = &H8000000F
    lblChartNumber.ForeColor = vbBlack
  End If
  If IsDate(tb!SampleDate) Then
    dtSampleDate = Format$(tb!SampleDate, "dd/mm/yyyy")
    lblSampledate = dtSampleDate
  Else
    dtSampleDate = Format$(Now, "dd/mm/yyyy")
    lblSampledate = dtSampleDate
  End If
  If IsDate(tb!Rundate) Then
    dtRunDate = Format$(tb!Rundate, "dd/mm/yyyy")
  Else
    dtRunDate = Format$(Now, "dd/mm/yyyy")
  End If
  StatusBar1.Panels(4).Text = dtRunDate
  mNewRecord = False
  If Trim$(tb!RooH & "") <> "" Then cRooH(0) = tb!RooH
  If Trim$(tb!RooH & "") <> "" Then cRooH(1) = Not tb!RooH
  txtChart = DeencryptA(Trim(tb!Chart & ""))
  txtName = Trim(initial2upper(DeencryptA(tb!PatName & "")))
  txtNOPAS = Trim(tb!NOPAS & "")
  txtAandE = Trim(tb!aande & "")
  lblNOPAS(1) = Trim(tb!NOPAS & "")
  taddress(0) = DeencryptA(tb!Addr0 & "")
  taddress(1) = DeencryptA(tb!Addr1 & "")
  Select Case Left$(Trim$(UCase$(tb!sex & "")), 1)
    Case "M": txtSex = "Male"
    Case "F": txtSex = "Female"
    Case Else: txtSex = ""
  End Select
  If Trim(tb!DoB & "") <> "" Then txtDoB = Format$(tb!DoB, "dd/mm/yyyy") Else txtDoB = ""
  If tb!Age & "" <> "" Then
    txtAge = Trim(tb!Age)
  Else
    If Trim(tb!DoB & "") <> "" Then txtAge = CalcOldAge(tb!DoB, dtRunDate)
  End If
  lAge = txtAge & ""
  lDoB = txtDoB
  If Trim(tb!Hospital) & "" <> "" Then cmbHospital = tb!Hospital Else cmbHospital = Hospname(0)
  cmbClinician = Trim(tb!Clinician & "")
  cmbGP = Trim(tb!GP & "")
  cmbWard = Trim(tb!Ward & "")
  cClDetails = Trim(tb!cldetails & "")
  If Trim$(tb!Category & "") <> "" Then
    cCat(0) = Trim(tb!Category & "")
    cCat(1) = Trim(tb!Category & "")
  Else
    cCat(0) = "Default"
    cCat(1) = "Default"
 End If
'  If SysOptPgp(0) Then
'   If Trim(tb!forend & "") = True Then chkPgp.Value = 1 Else chkPgp.Value = 0
'  End If

'16/01/2006
Rem
'This code was removed and replaced by the
'code below the rem below

'  If IsDate(tb!SampleDate) Then
'    dtSampleDate = Format$(tb!SampleDate, "dd/mm/yyyy")
'    If Format$(tb!SampleDate, "hh:mm") <> "00:00" Then
'      tSampleTime = Format$(tb!SampleDate, "hh:mm")
'    Else
'      tSampleTime.Mask = ""
'      tSampleTime.Text = ""
'      tSampleTime.Mask = "##:##"
'    End If
'  Else
'      dtSampleDate = Format$(Now, "dd/mm/yyyy")
'      tSampleTime.Mask = ""
'      tSampleTime.Text = ""
'      tSampleTime.Mask = "##:##"
'  End If

Rem  Code Change
  If IsDate(tb!SampleDate) Then
    dtSampleDate = Format$(tb!SampleDate, "dd/mm/yyyy")
    If Format$(tb!SampleDate, "hh:mm") <> "00:00" Then
      tSampleTime = Format$(tb!SampleDate, "hh:mm")
    Else
      tSampleTime.Mask = ""
      tSampleTime.Text = ""
      tSampleTime.Mask = "##:##"
    End If
  ElseIf IsDate(tb!RecDate) Then
    dtSampleDate = Format$(tb!RecDate, "dd/mm/yyyy")
    tSampleTime.Mask = ""
    tSampleTime.Text = ""
    tSampleTime.Mask = "##:##"
  ElseIf IsDate(tb!Rundate & "") Then
    dtSampleDate = Format$(tb!Rundate, "dd/mm/yyyy")
    tSampleTime.Mask = ""
    tSampleTime.Text = ""
    tSampleTime.Mask = "##:##"
  End If
  If SysOptDemoVal(0) = True Then
    If tb!Valid = True Then
      cmdDemoVal.Caption = "VALID"
      Set_Demo False
    Else
      cmdDemoVal.Caption = "&Validate"
      Set_Demo True
    End If
  End If
  If IsDate(tb!RecDate & "") Then
    dtRecDate = Format$(tb!RecDate, "dd/mm/yyyy")
    If Format$(tb!RecDate, "hh:mm") <> "00:00" Then
      tRecTime = Format$(tb!RecDate, "hh:mm")
    Else
      tRecTime.Mask = ""
      tRecTime.Text = ""
      tRecTime.Mask = "##:##"
    End If
  Else
    If Trim(tb!Rundate & "") <> "" Then dtRecDate = Format$(tb!Rundate, "dd/mm/yyyy")
    tRecTime.Mask = ""
    tRecTime.Text = ""
    tRecTime.Mask = "##:##"
  End If
  If Trim$(tb!Fasting & "") <> "" Then
    If tb!Fasting Then
      lRandom = "Fasting Sample"
    End If
  End If
  If SysOptUrgent(0) Then
    If tb!urgent = 1 Then
      lblUrgent.Visible = True
      chkUrgent.Value = 1
      UrgentTest = True
    Else
      chkUrgent.Value = 0
      UrgentTest = False
    End If
  End If
End If



cmdSaveDemographics.Enabled = False
cmdSaveInc.Enabled = False

If SysOptViewTrans(0) = True Then
  bViewBB.Visible = True
  If CnxnBB(0) Is Nothing Then
  Else
  If Trim$(txtChart) <> "" And Right(CnxnBB(0), 2) <> "=;" Then
    SQL = "SELECT  * from PatientDetails WHERE " & _
          "PatNum = '" & txtChart & "'"
    Set tb = New Recordset
    RecOpenClientBB tb, SQL
    bViewBB.Enabled = Not tb.EOF
  End If
  End If
End If

CheckCC

Screen.MousePointer = 0


Exit Sub

LoadDemographics_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /LoadDemographics ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Public Sub LoadEndocrinology()

Dim DeltaSn As Recordset
Dim Deltatb As Recordset
Dim tb As New Recordset
Dim SQL As String
Dim s As String
Dim Value As Single
Dim OldValue As Single
Dim valu As String
Dim PreviousDate As String
Dim PreviousRec As Long
Dim Res As String
Dim n As Long
Dim e As String
Dim DeltaLimit As Single
Dim SampleType As String
Dim Ims As New BIEResults
Dim IMres As BIEResults
Dim Im As BIEResult
Dim Fasting As Boolean
Dim Flag As String
Dim Cat As String
Dim sn As New Recordset

On Error GoTo LoadEndocrinology_Error

If txtSampleID = "" Then Exit Sub

PreviousEnd = False
HistEnd = False

Fasting = lImmRan(0) = "Fasting Sample"

lblEDate = ""
ldelta = ""
bViewImmRepeat(0).Visible = False

sstabAll.TabCaption(4) = "Endocrinology"

ClearFGrid gImm(0)

'get date & run number of previous record


SQL = CreateHist("end")
  Set sn = New Recordset
  RecOpenServer 0, sn, SQL
  If Not sn.EOF Then
      HistEnd = True
  End If

SQL = CreateSql("End")

        
    
  Set Deltatb = New Recordset
  RecOpenServer 0, Deltatb, SQL
  If Not Deltatb.EOF Then
    PreviousDate = Deltatb!Rundate & ""
    PreviousRec = DeencryptN(Deltatb!SampleID & "")
    PreviousEnd = True
  End If

  

If cCat(0) = "" Then Cat = "Default" Else Cat = cCat(0)

Set IMres = Ims.Load("End", txtSampleID, "Results", gDONTCARE, gDONTCARE, 0, Cat, dtRunDate)

CheckCalcEPSA IMres

If Not IMres Is Nothing Then
  sstabAll.TabCaption(4) = ">>Endocrinology<<"
  For Each Im In IMres
    SampleType = Im.SampleType
    If Len(SampleType) = 0 Then SampleType = "S"
    s = Im.ShortName & vbTab
    lblEDate = Im.RunTime
    If IsNumeric(Im.Result) Then
      Value = Val(Im.Result)
      Select Case Im.Printformat
        Case 0: valu = Format$(Value, "0")
        Case 1: valu = Format$(Value, "0.0")
        Case 2: valu = Format$(Value, "0.00")
        Case 3: valu = Format$(Value, "0.000")
        Case Else: valu = Format$(Value, "0.000")
      End Select
    Else
      valu = Im.Result
    End If
 '   If UserMemberOf = "The World" And Not BR.Valid Then
 '     s = s & "" & vbTab
 '   Else
    s = s & valu & vbTab
    If ListText("UN", Im.Units) <> "" Then
      s = s & ListText("UN", Im.Units)
    Else
      s = s & Im.Units
    End If
    s = s & vbTab
    s = s & Im.Low & " - " & Im.High & vbTab
    
    If IsNumeric(Im.Result) Then
      If Value > Val(Im.PlausibleHigh) Then
        Flag = "X"
        s = s & "X"
      ElseIf Value < Val(Im.PlausibleLow) Then
        Flag = "X"
        s = s & "X"
      Else
        If Value < Val(Im.Low) Then
          Flag = "L"
          s = s & "L"
        ElseIf Value > Val(Im.High) Then
          Flag = "H"
          s = s & "H"
        End If
      End If
    Else
        If Left(Im.Result, 1) = "<" Then
          Flag = "L"
          s = s & "L"
        ElseIf Left(Im.Result, 1) = ">" Then
          Flag = "H"
          s = s & "H"
       End If
    End If
    If Im.Flags = "1" Then e = "C" Else e = ""
    s = s & vbTab & _
            IIf(e <> "", e, "") & vbTab & _
            IIf(Im.Valid, "V", " ") & _
            IIf(Im.Printed, "P", " ") & vbTab & Trim(Im.Comment)
    gImm(0).AddItem s
    If Flag <> "" Then
      gImm(0).Row = gImm(0).Rows - 1
      gImm(0).Col = 1
      Select Case Flag
        Case "H":
          For n = 0 To 7
            gImm(0).Col = n
            gImm(0).CellBackColor = vbRed
            gImm(0).CellForeColor = vbYellow
          Next
        Case "L":
          For n = 0 To 7
            gImm(0).Col = n
            gImm(0).CellBackColor = vbBlue
            gImm(0).CellForeColor = vbYellow
          Next
        Case "X":
          For n = 0 To 7
            gImm(0).Col = n
            gImm(0).CellBackColor = vbGreen
            gImm(0).CellForeColor = vbWhite
          Next
      End Select
    End If
    Flag = ""
    If Im.DoDelta And PreviousEnd Then
      SQL = "SELECT * from endresults WHERE " & _
            "sampleid = '" & PreviousRec & "' " & _
            "and code = '" & Im.Code & "'"
      Set DeltaSn = New Recordset
      RecOpenClient 0, DeltaSn, SQL
      If Not DeltaSn.EOF Then
        OldValue = Val(DeltaSn!Result)
        If OldValue <> 0 Then
          DeltaLimit = Im.DeltaLimit
          If Abs(OldValue - Value) > DeltaLimit Then
            Res = Format$(PreviousDate, "dd/mm/yyyy") & " (" & PreviousRec & ") " & _
                  Im.ShortName & " " & _
                  OldValue & vbCr
            ldelta = ldelta & Res
          End If
        End If
      End If
    End If
  Next
End If
  
FixG gImm(0)
  
With gImm(0)
  bValidateImm(0).Caption = "VALID"
  lblUrgent.Visible = False
  For n = 1 To .Rows - 1
    If .TextMatrix(n, 3) = "X" Then
      .Row = n
      .Col = 1
      .CellForeColor = vbWhite
      .CellBackColor = vbBlack
    End If
    If InStr(.TextMatrix(n, 6), "V") = "0" Then
        bValidateImm(0).Caption = "&Validate"
        lblUrgent.Visible = UrgentTest
    End If
  Next
End With

LoadOutstandingEnd

SQL = "SELECT * from endRepeats WHERE " & _
      "SampleID = '" & txtSampleID & "'"
Set tb = New Recordset
RecOpenServer 0, tb, SQL
bViewImmRepeat(0).Visible = False
If Not tb.EOF Then
  bViewImmRepeat(0).Visible = True
End If

SQL = "SELECT * from EndMasks WHERE " & _
      "SampleID = '" & txtSampleID & "'"
Set tb = New Recordset
RecOpenServer 0, tb, SQL
If Not tb.EOF Then
  Ih(0) = IIf(tb!h, 1, 0)
  Iis(0) = IIf(tb!s, 1, 0)
  Il(0) = IIf(tb!l, 1, 0)
  Io(0) = IIf(tb!o, 1, 0)
  Ig(0) = IIf(tb!g, 1, 0)
  Ij(0) = IIf(tb!J, 1, 0)
End If

Exit Sub

LoadEndocrinology_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /LoadEndocrinology ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub LoadExt()
Dim SQL As String
Dim tb As New Recordset
Dim Deltatb As Recordset
Dim Str As String
Dim TestName As String
Dim PreviousDate As String
Dim PreviousRec As Long
Dim sn As New Recordset

On Error GoTo LoadExt_Error

If txtSampleID = "" Then Exit Sub

ClearFGrid grdExt

sstabAll.TabCaption(7) = "Externals"

PreviousExt = False
HistExt = False


SQL = CreateHist("Ext")
  Set sn = New Recordset
  RecOpenServer 0, sn, SQL
  If Not sn.EOF Then
      HistExt = True
  End If

SQL = CreateSql("Ext")
        
  Set Deltatb = New Recordset
  RecOpenServer 0, Deltatb, SQL
  If Not Deltatb.EOF Then
    PreviousDate = Deltatb!Rundate & ""
    PreviousRec = DeencryptN(Deltatb!SampleID & "")
    PreviousExt = True
End If

SQL = "SELECT * FROM Extresults WHERE sampleid = " & txtSampleID & ""
Set tb = New Recordset
RecOpenServer 0, tb, SQL
Do While Not tb.EOF
  sstabAll.TabCaption(7) = ">>Externals<<"
  If IsNumeric(tb!Analyte) Then TestName = eNumber2Name(Trim(tb!Analyte & ""))
  Str = Trim(tb!Analyte) & vbTab
  Str = Str & TestName & vbTab
  Str = Str & Trim(tb!Result) & vbTab
  Str = Str & eName2Normal(TestName) & vbTab
  Str = Str & eName2Units(TestName) & vbTab
  Str = Str & eName2SendTo(TestName) & vbTab
  If Not IsNull(tb!SENTDate) Then
    Str = Str & Format(tb!SENTDate, "dd/mmm/yyyy")
  End If
  Str = Str & vbTab
  If Not IsNull(tb!RetDate) Then
    Str = Str & Format(tb!RetDate, "dd/mmm/yyyy")
  End If
  Str = Str & vbTab & Trim(tb!SapCode & "")
  grdExt.AddItem Str
  tb.MoveNext
Loop

FixG grdExt

SQL = "SELECT * from etc WHERE sampleid = '" & txtSampleID & "'"
Set tb = New Recordset
RecOpenServer 0, tb, SQL
If Not tb.EOF Then
  txtEtc(0) = tb!etc0 & ""
  txtEtc(1) = tb!etc1 & ""
  txtEtc(2) = tb!etc2 & ""
  txtEtc(3) = tb!etc3 & ""
  txtEtc(4) = tb!etc4 & ""
  txtEtc(5) = tb!etc5 & ""
  txtEtc(6) = tb!etc6 & ""
  txtEtc(7) = tb!etc7 & ""
  txtEtc(8) = tb!etc8 & ""
Else
  txtEtc(0) = ""
  txtEtc(1) = ""
  txtEtc(2) = ""
  txtEtc(3) = ""
  txtEtc(4) = ""
  txtEtc(5) = ""
  txtEtc(6) = ""
  txtEtc(7) = ""
  txtEtc(8) = ""
End If

cmdSaveExt.Enabled = False

Exit Sub

LoadExt_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /LoadExt ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Public Sub LoadHaematology()

Dim tb As New Recordset
Dim sn As New Recordset
Dim n As Long
Dim ip As String
Dim e As String
Dim PrevDate As String
Dim PrevID As String
Dim SQL As String
   

On Error GoTo LoadHaematology_Error

ReDim i(0 To 6) As String
Dim HD As HaemTestDefinition
Dim PrevChcm As Single
Dim PrevRBC As Single
Dim PrevHgb As Single
Dim PrevMCV As Single
Dim PrevLucp As Single
Dim PrevLucA As Single
Dim PrevHct As Single
Dim PrevRDWCV As Single
Dim PrevRDWSD As Single
Dim PrevMCH As Single
Dim PrevMCHC As Single
Dim Prevplt As Single
Dim PrevMPV As Single
Dim PrevPLCR As Single
Dim PrevPdw As Single
Dim PrevWBC As Single
Dim PrevLymA As Single
Dim PrevLymP As Single
Dim PrevMonoA As Single
Dim PrevMonoP As Single
Dim PrevNeutA As Single
Dim PrevNeutP As Single
Dim PrevEosA As Single
Dim PrevEosP As Single
Dim PrevBasA As Single
Dim PrevBasP As Single
Dim PrevHDW As Single
Dim DoB As String
Dim ThisValid As Long
Dim GR As Long
Dim Asql As String
Dim Csql As String
Dim nSql As String
Dim DaysOld As Long
Dim SA As Long


'Panel3D4.Enabled = True
'Panel3D5.Enabled = True
'Panel3D6.Enabled = True
''txtHaemComment.Locked = False
''Panel3D7.Enabled = True
'
'bHaemGraphs.Visible = False
'bViewHaemRepeat.Visible = False
'PreviousHaem = False
'HistHaem = False
'sstabAll.TabCaption(1) = "Haematology"
'lHaemErrors.Visible = False
'bValidateHaem.Caption = "&Validate"
'SA = 0
'lblHaemPrinted = ""
'txtCondition = ""
'lblHaemValid.Visible = True
'txtEsr1 = ""
'txtEsr1.Visible = False
'
'grdH.Height = 2000
'
'If Trim$(txtSampleID) = "" Then Exit Sub
'
''grdH.Visible = False
''gRbc.Visible = False
'
'DoB = txtDoB
'If DoB <> "" Then DaysOld = Abs(DateDiff("d", Now, txtDoB)) Else DaysOld = 12783
'
'
'If txtName <> "" And txtDoB <> "" Then
'SQL = CreateHist("haem")
'  Set sn = New Recordset
'  RecOpenServer 0, sn, SQL
'  If Not sn.EOF Then
'      HistHaem = True
'  End If
'
'SQL = CreateSql("haem")
'
'  Set sn = New Recordset
'  RecOpenServer 0, sn, SQL
'  If Not sn.EOF Then
'    PrevDate = sn!Rundate
'    PrevID = DeencryptN(sn!SampleID)
'    SQL = "SELECT * from HaemResults WHERE " & _
'          "SampleID = '" & PrevID & "'"
'    Set tb = New Recordset
'    RecOpenServer 0, tb, SQL
'    If Not tb.EOF Then
'      PreviousHaem = True
'      PrevRBC = Val(tb!rbc & "")
'      PrevHgb = Val(tb!Hgb & "")
'      PrevMCV = Val(tb!MCV & "")
'      PrevHct = Val(tb!Hct & "")
'      PrevRDWCV = Val(tb!RDWCV & "")
'      PrevRDWSD = Val(tb!rdwsd & "")
'      PrevMCH = Val(tb!mch & "")
'      PrevMCHC = Val(tb!mchc & "")
'      Prevplt = Val(tb!Plt & "")
'      PrevMPV = Val(tb!mpv & "")
'      PrevPLCR = Val(tb!plcr & "")
'      PrevPdw = Val(tb!pdw & "")
'      PrevWBC = Val(tb!wbc & "")
'      PrevLymA = Val(tb!LymA & "")
'      PrevLymP = Val(tb!LymP & "")
'      PrevMonoA = Val(tb!MonoA & "")
'      PrevMonoP = Val(tb!MonoP & "")
'      PrevNeutA = Val(tb!NeutA & "")
'      PrevNeutP = Val(tb!NeutP & "")
'      PrevEosA = Val(tb!EosA & "")
'      PrevEosP = Val(tb!EosP & "")
'      PrevBasA = Val(tb!BasA & "")
'      PrevBasP = Val(tb!BasP & "")
'      PrevLucA = Val(tb!luca & "")
'      PrevLucp = Val(tb!lucp & "")
'      PrevChcm = Val(tb!cH & "")
'      PrevHDW = Val(tb!hdw & "")
'    End If
'  End If
'End If
'SQL = "SELECT * from HaemResults WHERE " & _
'      "SampleID = '" & txtSampleID & "'"
'
'Set tb = New Recordset
'RecOpenServer 0, tb, SQL
'
'If tb.EOF Then
'  bValidateHaem.Enabled = False
'  lblAgeSex.Visible = False
'  lblHaemValid.Visible = False
'Else
'  If txtDoB = "" Then
'    SA = 2
'  End If
'
'  If txtSex = "" Then
'    SA = SA + 1
'  End If
'  Select Case SA
'  Case 3: lblAgeSex = "Ref Range Not Age/Sex Related"
'          lblAgeSex.Visible = True
'  Case 2: lblAgeSex = "Ref Range Not Age Related"
'          lblAgeSex.Visible = True
'  Case 1: lblAgeSex = "Ref Range Not Sex Related"
'          lblAgeSex.Visible = True
'  Case Else
'    lblAgeSex.Visible = False
'  End Select
'  bValidateHaem.Enabled = True
'  If SysOptHaemAn1(0) = "ADVIA" And Trim(tb!Analyser) & "" = "1" Or SysOptHaemAn2(0) = "ADVIA" Or SysOptHaemAn2(0) = "ADVIA60" Then
'    SQL = "SELECT * from HaemFlags WHERE " & _
'      "sampleid = '" & txtSampleID & "'"
'      Set sn = New Recordset
'      RecOpenServer 0, sn, SQL
'      If Not sn.EOF Then
'          If Trim(sn!Flags) = "" Or IsNull(sn!Flags) Then
'              lHaemErrors.Visible = False
'          Else
'              lHaemErrors.Visible = True
'          End If
'      Else
'        lHaemErrors.Visible = False
'      End If
'  Else
'    If Not IsNull(tb!LongError) Then
'      If Val(tb!LongError) > 1 Then
'        lHaemErrors.Visible = True
'        lHaemErrors.Tag = Format$(tb!LongError)
'      End If
'    End If
'  End If
'  If Not IsNull(tb!gwb1) Or Not IsNull(tb!gwb2) Or Not IsNull(tb!gRbc) Or Not IsNull(tb!gplt) Then
'    bHaemGraphs.Visible = True
'  Else
'    bHaemGraphs.Visible = True
'
'  End If
'
'  pdelta.Cls
'  lHDate = Format(tb!RunDateTime, "dd/MM/yyyy hh:mm:ss")
'  If tb!wic & "" <> "" Then
'    lWIC = Trim(tb!wic & "")
'    lWOC = Trim(tb!woc & "")
'    Label3 = "WIC"
'    Label18 = "WOC"
'  Else
'    Label3 = "WCBC"
'    Label18 = "WCBP"
'  End If
'  If lWIC = "" Then lWIC = Trim(tb!wb & "")
'  If lWOC = "" Then lWOC = Trim(tb!wp & "")
'
'  cFilm = 0
'  If Not IsNull(tb!cFilm) Then
'    cFilm = IIf(tb!cFilm, 1, 0)
'  End If
'
'  If Not IsNull(tb!rbc) Then
'    ColouriseG "RBC", gRbc, 1, 1, Trim(tb!rbc), txtSex, DoB
'    If PreviousHaem Then DeltaCheck "RBC", tb!rbc, PrevRBC, PrevDate, PrevID
'    For Each HD In colHaemTestDefinitions
'      If HD.AnalyteName = "RBC" Then
'        If DaysOld >= HD.AgeFromDays And DaysOld <= HD.AgeToDays Then
'          If Left(txtSex, 1) = "M" Then
'            gRbc.TextMatrix(1, 2) = HD.MaleLow & " - " & HD.MaleHigh
'          ElseIf Left(txtSex, 1) = "F" Then
'            gRbc.TextMatrix(1, 2) = HD.FemaleLow & " - " & HD.FemaleHigh
'          Else
'            gRbc.TextMatrix(1, 2) = HD.FemaleLow & " - " & HD.MaleHigh
'          End If
'        End If
'      End If
'    Next
'  End If
'
'
'
'
'  If Not IsNull(tb!Hgb) Then
'    ColouriseG "Hgb", gRbc, 2, 1, Trim(tb!Hgb), txtSex, DoB
'    gRbc.Row = 2
'
'      gRbc.Col = 1
'      gRbc.CellFontSize = 12
'
'    If PreviousHaem Then DeltaCheck "Hgb", tb!Hgb, PrevHgb, PrevDate, PrevID
'    For Each HD In colHaemTestDefinitions
'      If HD.AnalyteName = "Hgb" Then
'        If DaysOld >= HD.AgeFromDays And DaysOld <= HD.AgeToDays Then
'          If Left(txtSex, 1) = "M" Then
'            gRbc.TextMatrix(2, 2) = HD.MaleLow & " - " & HD.MaleHigh
'          ElseIf Left(txtSex, 1) = "F" Then
'            gRbc.TextMatrix(2, 2) = HD.FemaleLow & " - " & HD.FemaleHigh
'          Else
'            gRbc.TextMatrix(2, 2) = HD.FemaleLow & " - " & HD.MaleHigh
'          End If
'        End If
'      End If
'    Next
'  End If
'
'
'  If Not IsNull(tb!Hct) Then
'    ColouriseG "Hct", gRbc, 3, 1, Trim(tb!Hct), txtSex, DoB
'    If PreviousHaem Then DeltaCheck "Hct", tb!Hct, PrevHct, PrevDate, PrevID
'    For Each HD In colHaemTestDefinitions
'      If HD.AnalyteName = "Hct" Then
'        If DaysOld >= HD.AgeFromDays And DaysOld <= HD.AgeToDays Then
'          If Left(txtSex, 1) = "M" Then
'            gRbc.TextMatrix(3, 2) = HD.MaleLow & " - " & HD.MaleHigh
'          ElseIf Left(txtSex, 1) = "F" Then
'            gRbc.TextMatrix(3, 2) = HD.FemaleLow & " - " & HD.FemaleHigh
'          Else
'            gRbc.TextMatrix(3, 2) = HD.FemaleLow & " - " & HD.MaleHigh
'          End If
'        End If
'      End If
'    Next
'  End If
'
'
'  If Not IsNull(tb!MCV) Then
'    ColouriseG "MCV", gRbc, 4, 1, Trim(tb!MCV), txtSex, DoB
'    If PreviousHaem Then DeltaCheck "MCV", tb!MCV, PrevMCV, PrevDate, PrevID
'    For Each HD In colHaemTestDefinitions
'      If HD.AnalyteName = "MCV" Then
'          If DaysOld >= HD.AgeFromDays And DaysOld <= HD.AgeToDays Then
'            If Left(txtSex, 1) = "M" Then
'              gRbc.TextMatrix(4, 2) = HD.MaleLow & " - " & HD.MaleHigh
'            ElseIf Left(txtSex, 1) = "F" Then
'              gRbc.TextMatrix(4, 2) = HD.FemaleLow & " - " & HD.FemaleHigh
'            Else
'              gRbc.TextMatrix(4, 2) = HD.FemaleLow & " - " & HD.MaleHigh
'            End If
'          End If
'      End If
'    Next
'  End If
'
'
'  If SysOptHaemAn1(0) <> "" Then
'  If Not IsNull(tb!hdw) Then
'    ColouriseG "HDW", gRbc, 5, 1, Trim(tb!hdw), txtSex, DoB
'    If PreviousHaem Then DeltaCheck "HDW", tb!hdw, PrevHDW, PrevDate, PrevID
'    For Each HD In colHaemTestDefinitions
'      If HD.AnalyteName = "HDW" Then
'          If DaysOld >= HD.AgeFromDays And DaysOld <= HD.AgeToDays Then
'            If Left(txtSex, 1) = "M" Then
'              gRbc.TextMatrix(5, 2) = HD.MaleLow & " - " & HD.MaleHigh
'            ElseIf Left(txtSex, 1) = "F" Then
'              gRbc.TextMatrix(5, 2) = HD.FemaleLow & " - " & HD.FemaleHigh
'            Else
'              gRbc.TextMatrix(5, 2) = HD.FemaleLow & " - " & HD.MaleHigh
'            End If
'          End If
'      End If
'    Next
'  End If
'  End If
'
'  If Not IsNull(tb!mch) Then
'    ColouriseG "MCH", gRbc, 6, 1, Trim(tb!mch), txtSex, DoB
'    If PreviousHaem Then DeltaCheck "MCH", tb!mch, PrevMCH, PrevDate, PrevID
'    For Each HD In colHaemTestDefinitions
'      If HD.AnalyteName = "MCH" Then
'          If DaysOld >= HD.AgeFromDays And DaysOld <= HD.AgeToDays Then
'            If Left(txtSex, 1) = "M" Then
'              gRbc.TextMatrix(6, 2) = HD.MaleLow & " - " & HD.MaleHigh
'            ElseIf Left(txtSex, 1) = "F" Then
'              gRbc.TextMatrix(6, 2) = HD.FemaleLow & " - " & HD.FemaleHigh
'            Else
'              gRbc.TextMatrix(6, 2) = HD.FemaleLow & " - " & HD.MaleHigh
'            End If
'          End If
'      End If
'    Next
'  End If
'
'
'  If Not IsNull(tb!mchc) Then
'    ColouriseG "MCHC", gRbc, 7, 1, Trim(tb!mchc), txtSex, DoB
'    If PreviousHaem Then DeltaCheck "MCHC", tb!mchc, PrevMCHC, PrevDate, PrevID
'    For Each HD In colHaemTestDefinitions
'      If HD.AnalyteName = "MCHC" Then
'          If DaysOld >= HD.AgeFromDays And DaysOld <= HD.AgeToDays Then
'            If Left(txtSex, 1) = "M" Then
'              gRbc.TextMatrix(7, 2) = HD.MaleLow & " - " & HD.MaleHigh
'            ElseIf Left(txtSex, 1) = "F" Then
'              gRbc.TextMatrix(7, 2) = HD.FemaleLow & " - " & HD.FemaleHigh
'            Else
'              gRbc.TextMatrix(7, 2) = HD.FemaleLow & " - " & HD.MaleHigh
'            End If
'          End If
'         End If
'    Next
'  End If
'
'  If SysOptHaemAn1(0) <> "" Then
'  If Not IsNull(tb!cH) Then
'    ColouriseG "ChcM", gRbc, 8, 1, Trim(tb!cH), txtSex, DoB
'    If PreviousHaem Then DeltaCheck "CHCM", tb!cH, PrevLucp, PrevDate, PrevID
'    For Each HD In colHaemTestDefinitions
'      If HD.AnalyteName = "CHCM" Then
'          If DaysOld >= HD.AgeFromDays And DaysOld <= HD.AgeToDays Then
'            If Left(txtSex, 1) = "M" Then
'              gRbc.TextMatrix(8, 2) = HD.MaleLow & " - " & HD.MaleHigh
'            ElseIf Left(txtSex, 1) = "F" Then
'              gRbc.TextMatrix(8, 2) = HD.FemaleLow & " - " & HD.FemaleHigh
'            Else
'              gRbc.TextMatrix(8, 2) = HD.FemaleLow & " - " & HD.MaleHigh
'            End If
'          End If
'      End If
'    Next
'  End If
'  End If
'
'  If Not IsNull(tb!RDWCV) And Val(tb!RDWCV & "") <> 0 Then
'    ColouriseG "RDW", gRbc, 9, 1, Trim(tb!RDWCV), txtSex, DoB
'    If PreviousHaem Then DeltaCheck "RDW", tb!RDWCV, PrevRDWCV, PrevDate, PrevID
'    For Each HD In colHaemTestDefinitions
'      If HD.AnalyteName = "RDW" Then
'        If DaysOld >= HD.AgeFromDays And DaysOld <= HD.AgeToDays Then
'          If Left(txtSex, 1) = "M" Then
'            gRbc.TextMatrix(9, 2) = HD.MaleLow & " - " & HD.MaleHigh
'          ElseIf Left(txtSex, 1) = "F" Then
'            gRbc.TextMatrix(9, 2) = HD.FemaleLow & " - " & HD.FemaleHigh
'          Else
'            gRbc.TextMatrix(9, 2) = HD.FemaleLow & " - " & HD.MaleHigh
'          End If
'        End If
'      End If
'    Next
'  End If
'
'
'  If Not IsNull(tb!Plt) Then
'    Colourise "Plt", tPlt, Trim(tb!Plt), txtSex, DoB
'    If PreviousHaem Then DeltaCheck "plt", tb!Plt, Prevplt, PrevDate, PrevID
'  End If
'
'  If Not IsNull(tb!mpv) Then
'    Colourise "MPV", tMPV, Trim(tb!mpv), txtSex, DoB
'    If PreviousHaem Then DeltaCheck "MPV", tb!mpv, PrevMPV, PrevDate, PrevID
'  End If
'
'  If Not IsNull(tb!wbc) Then
'    Colourise "WBC", tWBC, tb!wbc, txtSex, DoB
'    If PreviousHaem Then DeltaCheck "WBC", tb!wbc, PrevWBC, PrevDate, PrevID
'    If SysOptWBCDC(0) = True Then
'        tWBC = Format(tb!wbc, "##0.0")
'    Else
'        tWBC = Format(tb!wbc, "##0.00")
'    End If
'  End If
'
'  'Diff
'
'  If Not IsNull(tb!NeutA) Then
'    ColouriseG "NeutA", grdH, 1, 0, Trim(tb!NeutA & ""), txtSex, DoB
'    If PreviousHaem Then DeltaCheck "NeutA", tb!NeutA, PrevNeutA, PrevDate, PrevID
'    For Each HD In colHaemTestDefinitions
'      If HD.AnalyteName = "NeutA" Then
'        If DaysOld >= HD.AgeFromDays And DaysOld <= HD.AgeToDays Then
'          If Left(txtSex, 1) = "M" Then
'            grdH.TextMatrix(1, 1) = HD.MaleLow & " - " & HD.MaleHigh
'          ElseIf Left(txtSex, 1) = "F" Then
'            grdH.TextMatrix(1, 1) = HD.FemaleLow & " - " & HD.FemaleHigh
'          Else
'            grdH.TextMatrix(1, 1) = HD.FemaleLow & " - " & HD.MaleHigh
'          End If
'        End If
'      End If
'    Next
'  End If
'
'  If Not IsNull(tb!NeutP) Then
'    ColouriseG "NeutP", grdH, 1, 3, Trim(tb!NeutP & ""), txtSex, DoB
'    If PreviousHaem Then DeltaCheck "NeutP", tb!NeutP, PrevNeutP, PrevDate, PrevID
'  End If
'
'  If Not IsNull(tb!LymA) Then
'    ColouriseG "LymA", grdH, 2, 0, Trim(tb!LymA & ""), txtSex, DoB
'    If PreviousHaem Then DeltaCheck "LymA", tb!LymA, PrevLymA, PrevDate, PrevID
'    For Each HD In colHaemTestDefinitions
'      If HD.AnalyteName = "LymA" Then
'        If DaysOld >= HD.AgeFromDays And DaysOld <= HD.AgeToDays Then
'          If Left(txtSex, 1) = "M" Then
'            grdH.TextMatrix(2, 1) = HD.MaleLow & " - " & HD.MaleHigh
'          ElseIf Left(txtSex, 1) = "F" Then
'            grdH.TextMatrix(2, 1) = HD.FemaleLow & " - " & HD.FemaleHigh
'          Else
'            grdH.TextMatrix(2, 1) = HD.FemaleLow & " - " & HD.MaleHigh
'          End If
'        End If
'      End If
'    Next
'  End If
'
'  If Not IsNull(tb!LymP) Then
'    ColouriseG "LymP", grdH, 2, 3, Trim(tb!LymP & ""), txtSex, DoB
'    If PreviousHaem Then DeltaCheck "LymP", tb!LymP, PrevLymP, PrevDate, PrevID
'  End If
'
'  If Not IsNull(tb!MonoA) Then
'    ColouriseG "MonoA", grdH, 3, 0, Trim(tb!MonoA & ""), txtSex, DoB
'    If PreviousHaem Then DeltaCheck "MonoA", tb!MonoA, PrevMonoA, PrevDate, PrevID
'    For Each HD In colHaemTestDefinitions
'      If HD.AnalyteName = "MonoA" Then
'        If DaysOld >= HD.AgeFromDays And DaysOld <= HD.AgeToDays Then
'          If Left(txtSex, 1) = "M" Then
'            grdH.TextMatrix(3, 1) = HD.MaleLow & " - " & HD.MaleHigh
'          ElseIf Left(txtSex, 1) = "F" Then
'            grdH.TextMatrix(3, 1) = HD.FemaleLow & " - " & HD.FemaleHigh
'          Else
'            grdH.TextMatrix(3, 1) = HD.FemaleLow & " - " & HD.MaleHigh
'          End If
'        End If
'      End If
'    Next
'  End If
'
'  If Not IsNull(tb!MonoP) Then
'    ColouriseG "MonoP", grdH, 3, 3, Trim(tb!MonoP & ""), txtSex, DoB
'    If PreviousHaem Then DeltaCheck "MonoP", tb!MonoP, PrevMonoP, PrevDate, PrevID
'  End If
'
'
'  If Not IsNull(tb!EosA) Then
'    ColouriseG "EosA", grdH, 4, 0, Trim(tb!EosA & ""), txtSex, DoB
'    If PreviousHaem Then DeltaCheck "EosA", tb!EosA, PrevEosA, PrevDate, PrevID
'    For Each HD In colHaemTestDefinitions
'      If HD.AnalyteName = "EosA" Then
'        If DaysOld >= HD.AgeFromDays And DaysOld <= HD.AgeToDays Then
'          If Left(txtSex, 1) = "M" Then
'            grdH.TextMatrix(4, 1) = HD.MaleLow & " - " & HD.MaleHigh
'          ElseIf Left(txtSex, 1) = "F" Then
'            grdH.TextMatrix(4, 1) = HD.FemaleLow & " - " & HD.FemaleHigh
'          Else
'            grdH.TextMatrix(4, 1) = HD.FemaleLow & " - " & HD.MaleHigh
'          End If
'        End If
'      End If
'    Next
'  End If
'
'  If Not IsNull(tb!EosP) Then
'    ColouriseG "EosP", grdH, 4, 3, Trim(tb!EosP & ""), txtSex, DoB
'    If PreviousHaem Then DeltaCheck "EosP", tb!EosP, PrevEosP, PrevDate, PrevID
'  End If
'
'  If Not IsNull(tb!BasA) Then
'    ColouriseG "BasA", grdH, 5, 0, Trim(tb!BasA & ""), txtSex, DoB
'    If PreviousHaem Then DeltaCheck "BasA", tb!BasA, PrevBasA, PrevDate, PrevID
'    For Each HD In colHaemTestDefinitions
'      If HD.AnalyteName = "BasA" Then
'        If DaysOld >= HD.AgeFromDays And DaysOld <= HD.AgeToDays Then
'          If Left(txtSex, 1) = "M" Then
'            grdH.TextMatrix(5, 1) = HD.MaleLow & " - " & HD.MaleHigh
'          ElseIf Left(txtSex, 1) = "F" Then
'            grdH.TextMatrix(5, 1) = HD.FemaleLow & " - " & HD.FemaleHigh
'          Else
'            grdH.TextMatrix(5, 1) = HD.FemaleLow & " - " & HD.MaleHigh
'          End If
'        End If
'      End If
'    Next
'  End If
'
'  If Not IsNull(tb!BasP) Then
'    ColouriseG "BasP", grdH, 5, 3, Trim(tb!BasP & ""), txtSex, DoB
'    If PreviousHaem Then DeltaCheck "BasP", tb!BasP, PrevBasP, PrevDate, PrevID
'  End If
'
'  If SysOptHaemAn1(0) <> "" Then
'    If Not IsNull(tb!luca) Then
'      ColouriseG "LucA", grdH, 6, 0, Trim(tb!luca & ""), txtSex, DoB
'      If PreviousHaem Then DeltaCheck "LucA", tb!luca, PrevLucA, PrevDate, PrevID
'      For Each HD In colHaemTestDefinitions
'        If UCase(HD.AnalyteName) = UCase("LUCA") Then
'          If DaysOld >= HD.AgeFromDays And DaysOld <= HD.AgeToDays Then
'            If Left(txtSex, 1) = "M" Then
'              grdH.TextMatrix(6, 1) = HD.MaleLow & " - " & HD.MaleHigh
'            ElseIf Left(txtSex, 1) = "F" Then
'              grdH.TextMatrix(6, 1) = HD.FemaleLow & " - " & HD.FemaleHigh
'            Else
'              grdH.TextMatrix(6, 1) = HD.FemaleLow & " - " & HD.MaleHigh
'            End If
'          End If
'        End If
'      Next
'    End If
'
'    If Not IsNull(tb!lucp) Then
'      ColouriseG "LucP", grdH, 6, 3, Trim(tb!lucp & ""), txtSex, DoB
'      If PreviousHaem Then DeltaCheck "LucP", tb!lucp, PrevLucp, PrevDate, PrevID
'    End If
'    tASOt = Trim(tb!tASOt & "")
'    tRa = Trim(tb!tRa & "")
'    If Not IsNull(tb!cASot) Then
'      cASot = IIf(tb!cASot, 1, 0)
'    Else
'      cASot = 0
'    End If
'    If Trim(tb!Analyser) = "1" Then
'      lblAnalyser = lblAnalyser & " " & SysOptHaemN1(0)
'      HaemAnalyser = "1"
'    ElseIf Trim(tb!Analyser) = "2" Then
'      HaemAnalyser = "2"
'      lblAnalyser = lblAnalyser & " " & SysOptHaemN2(0)
'    End If
'  End If
'
'  If SysOptHaemAn1(0) = "ADVIA" Then
'    If Trim(tb!LS & "") <> "" Or Trim(tb!va & "") <> "" _
'    Or Trim(tb!At & "") <> "" Or Trim(tb!bl & "") <> "" _
'    Or Trim(tb!An & "") <> "" Or Trim(tb!mi & "") <> "" _
'    Or Trim(tb!ca & "") <> "" Or Trim(tb!ho & "") <> "" _
'    Or Trim(tb!he & "") <> "" Or Trim(tb!Ig & "") <> "" _
'    Or Trim(tb!mpo & "") <> "" Or Trim(tb!lplt & "") <> "" _
'    Or Trim(tb!pclm & "") <> "" Or Trim(tb!rbcf & "") <> "" _
'    Or Trim(tb!rbcg & "") <> "" Or Trim(tb!nrbca & "") <> "" Then
'                lHaemErrors.Visible = True
'    End If
'    txtLI = Trim(tb!Li & "")
'    txtMPXI = Trim(tb!mpxi & "")
'    gRbc.TextMatrix(11, 1) = Trim(tb!Hyp & "")
'  End If
'
'
'  If SysOptBadRes(0) Then
'    If Not IsNull(tb!cbad) Then
'      chkBad = IIf(tb!cbad, 1, 0)
'    Else
'      chkBad = 0
'    End If
'  End If
'  '
'  gRbc.TextMatrix(10, 1) = Trim(tb!nrbcp & "")
'
''  tESR = Trim(tb!esr & "")
'
'  If SysOptESR1(0) Then
'    If Trim(tb!esr & "") <> "" Then Colourise "ESR", tESR, Trim(tb!esr), txtSex, DoB
'    If Trim(tb!esr1 & "") <> "" Then
'      txtEsr1.Visible = True
'      Colourise "ESR1", txtEsr1, Trim(tb!esr1), txtSex, DoB
'    End If
'  Else
'    If Trim(tb!esr & "") <> "" Then Colourise "ESR", tESR, Trim(tb!esr), txtSex, DoB
'  End If
'  If Trim(tb!reta & "") <> "" Then Colourise "RETA", tRetA, Trim(tb!reta), txtSex, DoB
'
'  tRetP = Trim(tb!retp) & ""
'
'  Select Case Trim$(tb!Monospot & "")
'    Case "P": tMonospot = "Positive"
'    Case "N": tMonospot = "Negative"
'    Case Else: tMonospot = ""
'  End Select
'
'  If Not IsNull(tb!cESR) Then
'    cESR = IIf(tb!cESR, 1, 0)
'  Else
'    cESR = 0
'  End If
'
'  If Not IsNull(tb!cRA) Then
'    cRA = IIf(tb!cRA, 1, 0)
'  Else
'    cRA = 0
'  End If
'
'  If SysOptBadRes(0) Then
'    If Not IsNull(tb!cbad) Then
'      chkBad = IIf(tb!cbad, 1, 0)
'    Else
'      chkBad = 0
'    End If
'  End If
'
'  If Not IsNull(tb!cRetics) Then
'    cRetics = IIf(tb!cRetics, 1, 0)
'  Else
'    cRetics = 0
'  End If
'
'  If Not IsNull(tb!cMonospot) Then
'    cMonospot = IIf(tb!cMonospot, 1, 0)
'  Else
'    cMonospot = 0
'  End If
'
'  If Not IsNull(tb!cMalaria) Then
'    chkMalaria = IIf(tb!cMalaria, 1, 0)
'  Else
'    chkMalaria = 0
'  End If
'  lblMalaria = Trim(tb!malaria & "")
'
'  If Not IsNull(tb!csickledex) Then
'    chkSickledex = IIf(tb!csickledex, 1, 0)
'  Else
'    chkSickledex = 0
'  End If
'  lblSickledex = Trim(tb!sickledex & "")
'
'  tWarfarin = Trim(tb!Warfarin & "")
'
'  ip = Left$(tb!ipmessage & "000000", 6)
'  For n = 0 To 5
'    ipflag(n).Enabled = Mid$(ip, n + 1, 1) = "1"
'  Next
'
'  e = tb!negposerror & ""
'
'  buildinterp tb, i()
'  If i(0) <> "" Then pdelta.Print
'  For n = 0 To 6
'    pdelta.ForeColor = vbRed
'    pdelta.Print i(n)
'  Next
'
'  ThisValid = False
'  If Not IsNull(tb!Valid) Then
'    ThisValid = IIf(tb!Valid, 1, 0)
'  End If
'  If ThisValid = 1 Then
'    bValidateHaem.Caption = "VALID"
'    lblUrgent.Visible = False
'  Else
'    bValidateHaem.Caption = "&Validate"
'    lblUrgent.Visible = UrgentTest
'  End If
'  If ThisValid = 1 Then lblHaemValid.Visible = True Else lblHaemValid.Visible = False
'  If ThisValid = 1 Then Panel3D4.Enabled = False Else Panel3D4.Enabled = True
'  If ThisValid = 1 Then Panel3D5.Enabled = False Else Panel3D5.Enabled = True
'  If ThisValid = 1 Then Panel3D6.Enabled = False Else Panel3D6.Enabled = True
''  If ThisValid = 1 Then txtHaemComment.Locked = True Else txtHaemComment.Locked = False
''  If ThisValid = 1 Then Panel3D7.Enabled = False Else Panel3D7.Enabled = True
'
'  If Not IsNull(tb!Printed) Then
'    If tb!Printed = 1 Then
'      lblHaemPrinted = "Already Printed"
'    Else
'      lblHaemPrinted = "Not Printed"
'    End If
'  Else
'    lblHaemPrinted = "Not Printed"
'  End If
'  If ThisValid = 0 Then
'    SQL = "SELECT * from HaemRepeats WHERE " & _
'          "SampleID = '" & txtSampleID & "'"
'    Set tb = New Recordset
'    RecOpenServer 0, tb, SQL
'    If Not tb.EOF Then
'      If tb!wbc & "" <> "" Or tb!reta & "" <> "" Then bViewHaemRepeat.Visible = True
'    End If
'  End If
'
'  If SysOptView(0) = True Then
'    SQL = "SELECT * from HaemRepeats WHERE " & _
'          "SampleID = '" & txtSampleID & "'"
'    Set tb = New Recordset
'    RecOpenServer 0, tb, SQL
'    If Not tb.EOF Then
'      If tb!wbc & "" <> "" Then bViewHaemRepeat.Visible = True
'    End If
'  End If
'
'  If Trim(txtChart) <> "" Then
'    SQL = "SELECT * from HaemCondition WHERE " & _
'          "chart = '" & txtChart & "'"
'    Set tb = New Recordset
'    RecOpenServer 0, tb, SQL
'    If Not tb.EOF Then
'      txtCondition = Trim(tb!condition)
'    End If
'  End If
'
'  sstabAll.TabCaption(1) = ">>Haematology<<"
'
'End If
'
'grdH.Visible = True
'gRbc.Visible = True
'
'SQL = "SELECT * from Differentials WHERE " & _
'      "runnumber = '" & txtSampleID & "'"
'Set tb = New Recordset
'RecOpenServer 0, tb, SQL
'If tb.EOF Then
'  bFilm.BackColor = &H8000000F
'Else
'  bFilm.BackColor = vbBlue
'  If tb!prndiff = True Then
'    grdH.Height = 360
'  End If
'End If
'
'FixG gRbc
'
'cmdSaveHaem.Enabled = False
'cmdSaveComm.Enabled = False
'cmdHSaveH.Enabled = False
'Screen.MousePointer = 0

Exit Sub

LoadHaematology_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /LoadHaematology ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select

End Sub

Private Function LoadImmSplitList(ByVal Index As Integer) As String

Dim tb As New Recordset
Dim SQL As String
Dim strIndex As String
Dim strReturn As String

On Error GoTo LoadImmSplitList_Error

strIndex = Index

SQL = "SELECT distinct Code, PrintPriority, SplitList " & _
      "from ImmTestDefinitions " & _
      "WHERE SplitList = " & strIndex & " " & _
      "order by PrintPriority"
      
Set tb = New Recordset
RecOpenServer 0, tb, SQL

strReturn = ""
Do While Not tb.EOF
  strReturn = strReturn & "Code = '" & tb!Code & "' or "
  tb.MoveNext
Loop
If strReturn <> "" Then
  strReturn = Left$(strReturn, Len(strReturn) - 3)
End If

LoadImmSplitList = strReturn

Exit Function

LoadImmSplitList_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /LoadImmSplitList ")
  Case 1:    End
  Case 2:    Exit Function
  Case 3:    Resume Next
End Select


End Function

Public Sub LoadImmunology()

Dim DeltaSn As Recordset
Dim Deltatb As Recordset
Dim tb As New Recordset
Dim SQL As String
Dim s As String
Dim Value As Single
Dim OldValue As Single
Dim valu As String
Dim PreviousDate As String
Dim PreviousRec As Long
Dim Res As String
Dim n As Long
Dim e As String
Dim DeltaLimit As Single
Dim SampleType As String
Dim Ims As New BIEResults
Dim IMres As BIEResults
Dim Im As BIEResult
Dim Fasting As Boolean
Dim Flag As String
Dim Cat As String
Dim sn As New Recordset


On Error GoTo LoadImmunology_Error

If txtSampleID = "" Then Exit Sub

PreviousImm = False
HistImm = False
lblIRundate = ""
cmdGetBio.Visible = True
Frame12(1).Enabled = True
'txtImmComment(1).Locked = False
'
Fasting = lImmRan(1) = "Fasting Sample"

ClearFGrid gImm(1)

lIDelta(1) = ""
bViewImmRepeat(1).Visible = False

sstabAll.TabCaption(6) = "Immunology"

'get date & run number of previous record
PreviousImm = False

If txtName <> "" And txtDoB <> "" Then
  
  SQL = CreateHist("imm")
    Set sn = New Recordset
    RecOpenServer 0, sn, SQL
    If Not sn.EOF Then
        HistImm = True
    End If
    
    SQL = CreateSql("Imm")
    Set Deltatb = New Recordset
    RecOpenServer 0, Deltatb, SQL
    If Not Deltatb.EOF Then
      PreviousDate = Deltatb!Rundate & ""
      PreviousRec = DeencryptN(Deltatb!SampleID & "")
      PreviousImm = True
    End If
End If


If cCat(0) = "" Then Cat = "Default" Else Cat = cCat(0)

Set IMres = Ims.Load("Imm", txtSampleID, "Results", gDONTCARE, gDONTCARE, 0, Cat, dtRunDate)

CheckCalcIPSA IMres

If Not IMres Is Nothing Then
  sstabAll.TabCaption(6) = ">>Immunology<<"
  For Each Im In IMres
    lblIRundate = Im.RunTime
    SampleType = Im.SampleType
    If Len(SampleType) = 0 Then SampleType = "S"
    s = Im.ShortName & vbTab
    
    If IsNumeric(Im.Result) Then
      Value = Val(Im.Result)
      Select Case Im.Printformat
        Case 0: valu = Format$(Value, "0")
        Case 1: valu = Format$(Value, "0.0")
        Case 2: valu = Format$(Value, "0.00")
        Case 3: valu = Format$(Value, "0.000")
        Case Else: valu = Format$(Value, "0.000")
      End Select
    Else
      valu = Im.Result
      If InStr(valu, "Pos") Then Flag = "N"
    End If
 '   If UserMemberOf = "The World" And Not BR.Valid Then
 '     s = s & "" & vbTab
 '   Else
    s = s & valu & vbTab
    If ListText("UN", Im.Units) <> "" Then
      s = s & ListText("UN", Im.Units)
    Else
      s = s & Im.Units
    End If
    s = s & vbTab
    If Im.PrnRR Then
        s = s & Im.Low & " - " & Im.High & vbTab
    Else
        s = s & vbTab
    End If
    If SysOptRealImm(0) Then
        If IsNumeric(Im.Result) Then
          If Value > Val(Im.PlausibleHigh) Then
            Flag = "X"
            s = s & "X"
          ElseIf Value < Val(Im.PlausibleLow) Then
            Flag = "X"
            s = s & "X"
          Else
            If Value < Val(Im.Low) Then
              Flag = "L"
              s = s & "L"
            ElseIf Value > Val(Im.High) Then
              Flag = "H"
              s = s & "H"
            Else
              Flag = ""
              s = s & ""
            End If
          End If
        Else
            s = s & ""
        End If
    Else
        If IsNumeric(Im.Result) Then
          If Value > Val(Im.PlausibleHigh) Then
            Flag = "X"
            s = s & "X"
          ElseIf Value < Val(Im.PlausibleLow) Then
            Flag = "X"
            s = s & "X"
          Else
            If Value < Val(Im.Low) Then
              Flag = "N"
              s = s & "N"
            ElseIf Value > Val(Im.High) Then
              Flag = "I"
              s = s & "I"
            Else
              Flag = "E"
              s = s & "E"
            End If
          End If
    Else
        Flag = ""
    End If
    End If
    e = ""
    s = s & vbTab & _
            IIf(e <> "", e, "") & vbTab & _
            IIf(Im.Valid, "V", " ") & _
            IIf(Im.Printed, "P", " ")
    s = s & vbTab & Im.Pc & vbTab & Im.Comment
    gImm(1).AddItem s
    If Flag <> "" Then
      gImm(1).Row = gImm(1).Rows - 1
      gImm(1).Col = 1
      Select Case Flag
        Case "N", "H":
          For n = 0 To 8
            gImm(1).Col = n
            gImm(1).CellBackColor = vbRed
            gImm(1).CellForeColor = vbYellow
          Next
        Case "E", "L":
          For n = 0 To 8
            gImm(1).Col = n
            gImm(1).CellBackColor = vbBlue
            gImm(1).CellForeColor = vbYellow
          Next
        Case "X":
          For n = 0 To 8
            gImm(1).Col = n
            gImm(1).CellBackColor = vbGreen
            gImm(1).CellForeColor = vbWhite
          Next
      End Select
    End If
    Flag = ""
    If Im.DoDelta And PreviousImm Then
      SQL = "SELECT * from immresults WHERE " & _
            "sampleid = '" & PreviousRec & "' " & _
            "and code = '" & Im.Code & "'"
      Set DeltaSn = New Recordset
      RecOpenClient 0, DeltaSn, SQL
      If Not DeltaSn.EOF Then
        OldValue = Val(DeltaSn!Result)
        If OldValue <> 0 Then
          DeltaLimit = Im.DeltaLimit
          If Abs(OldValue - Value) > DeltaLimit Then
            Res = Format$(PreviousDate, "dd/mm/yyyy") & " (" & PreviousRec & ") " & _
                  Im.ShortName & " " & _
                  OldValue & vbCr
            lIDelta(1) = lIDelta(1) & Res
          End If
        End If
      End If
    End If
  Next
End If
  
FixG gImm(1)
  
With gImm(1)
  bValidateImm(1).Caption = "VALID"
  lblUrgent.Visible = False
  'txtImmComment(1).Locked = True
  For n = 1 To .Rows - 1
    Frame2.Enabled = False
    If .TextMatrix(n, 3) = "X" Then
      .Row = n
      .Col = 1
      .CellForeColor = vbWhite
      .CellBackColor = vbBlack
    End If
    If InStr(.TextMatrix(n, 6), "V") = "0" Then
        bValidateImm(1).Caption = "&Validate"
        lblUrgent.Visible = UrgentTest
'        txtImmComment(1).Locked = False
    End If
  Next
End With

LoadOutstandingImm

SQL = "SELECT * from ImmRepeats WHERE " & _
      "SampleID = '" & Val(txtSampleID) & "'"
Set tb = New Recordset
RecOpenServer 0, tb, SQL
bViewImmRepeat(1).Visible = False
If Not tb.EOF Then
  bViewImmRepeat(1).Visible = True
End If

SQL = "SELECT * from ImmMasks WHERE " & _
      "SampleID = '" & txtSampleID & "'"
Set tb = New Recordset
RecOpenServer 0, tb, SQL
If Not tb.EOF Then
  Ih(1) = IIf(tb!h, 1, 0)
  Iis(1) = IIf(tb!s, 1, 0)
  Il(1) = IIf(tb!l, 1, 0)
  Io(1) = IIf(tb!o, 1, 0)
  Ig(1) = IIf(tb!g, 1, 0)
  Ij(1) = IIf(tb!J, 1, 0)
End If

If gImm(1).Rows > 2 And gImm(1).TextMatrix(1, 0) <> "" Then
  SQL = "SELECT * from bioresults WHERE sampleid = '" & txtSampleID & "' and " & _
        " (code = '" & SysOptBioCodeForAlb(0) & "' or " & _
        " code = '" & SysOptBioCodeForUProt(0) & "' or " & _
        " code = '" & SysOptBioCodeForGlob(0) & "' or " & _
        " code = '" & SysOptBioCodeFor24UProt(0) & "' or " & _
        " code = '" & SysOptBioCodeForTProt(0) & "' or " & _
        " code = '" & SysOptBioCodeFor24Vol(0) & "') "
  
  Set tb = New Recordset
  RecOpenServer 0, tb, SQL
  If Not tb.EOF Then cmdGetBio.Visible = True Else cmdGetBio.Visible = False
End If

Exit Sub

LoadImmunology_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /LoadImmunology ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub LoadOutstandingBio()

Dim tb As New Recordset
Dim SQL As String


On Error GoTo LoadOutstandingBio_Error

ClearOutstandingBio

SQL = "SELECT distinct(shortname) from biorequests, biotestdefinitions WHERE " & _
      "biorequests.sampleid = '" & txtSampleID & "' and biotestdefinitions.code = biorequests.code and biotestdefinitions.sampletype = biorequests.sampletype"
Set tb = New Recordset
RecOpenServer 0, tb, SQL
Do While Not tb.EOF
  grdOutstanding.AddItem Trim(tb!ShortName & "")
  tb.MoveNext
Loop

If grdOutstanding.Rows > 2 Then
  grdOutstanding.RemoveItem 1
End If


Exit Sub

LoadOutstandingBio_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /LoadOutstandingBio ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub LoadOutstandingEnd()

Dim tb As New Recordset
Dim SQL As String


On Error GoTo LoadOutstandingEnd_Error

ClearOutstandingEnd

SQL = "SELECT * from Endrequests WHERE " & _
      "sampleid = '" & txtSampleID & "'"
Set tb = New Recordset
RecOpenServer 0, tb, SQL
Do While Not tb.EOF
  grdOutstandings(0).AddItem EndShortNameFor(tb!Code & "")
  tb.MoveNext
Loop

If grdOutstandings(0).Rows > 2 Then
  grdOutstandings(0).RemoveItem 1
End If


Exit Sub

LoadOutstandingEnd_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /LoadOutstandingEnd ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub LoadOutstandingImm()

Dim tb As New Recordset
Dim SQL As String



On Error GoTo LoadOutstandingImm_Error

ClearOutstandingImm

If txtSampleID = "" Then Exit Sub

SQL = "SELECT * from Immrequests WHERE " & _
      "sampleid = '" & txtSampleID & "'"
Set tb = New Recordset
RecOpenServer 0, tb, SQL
Do While Not tb.EOF
  grdOutstandings(1).AddItem ImmShortNameFor(tb!Code & "")
  tb.MoveNext
Loop

If grdOutstandings(1).Rows > 2 Then
  grdOutstandings(1).RemoveItem 1
End If


Exit Sub

LoadOutstandingImm_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /LoadOutstandingImm ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub LoadOutstandingrdCoag()

Dim tb As New Recordset
Dim SQL As String



On Error GoTo LoadOutstandingrdCoag_Error

With grdOutstandingCoag
  .Rows = 2
  .AddItem ""
  .RemoveItem 1
End With

SQL = "SELECT * from CoagRequests WHERE " & _
      "sampleid = '" & txtSampleID & "'"
Set tb = New Recordset
RecOpenServer 0, tb, SQL
Do While Not tb.EOF
  grdOutstandingCoag.AddItem CoagNameFor(tb!Code & "", "")
  tb.MoveNext
Loop

If grdOutstandingCoag.Rows > 2 Then
  grdOutstandingCoag.RemoveItem 1
End If

Exit Sub

LoadOutstandingrdCoag_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /LoadOutstandingrdCoag ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub LoadPreviousCoag()

Dim tb As New Recordset
Dim SQL As String
Dim CRs As CoagResults
Dim CR As CoagResult
Dim PrevDate As String
Dim PrevID As String
Dim s As String
Dim g As String

On Error GoTo LoadPreviousCoag_Error

PreviousCoag = False

ClearFGrid grdPrev

  SQL = CreateSql("Coag")
  Set tb = New Recordset
  RecOpenServer 0, tb, SQL
  If Not tb.EOF Then
    
    PreviousCoag = True
    
    PrevDate = Format$(tb!Rundate, "dd/mm/yy")
    PrevID = DeencryptN(tb!SampleID)

    Set CRs = New CoagResults
    Set CRs = CRs.Load(PrevID, gDONTCARE, gDONTCARE, Trim(SysOptExp(0)), 0)

    If Not CRs Is Nothing Then
      For Each CR In CRs
          If Trim(CR.Units) = "INR" Then
            s = "INR" & vbTab
          Else
            s = CoagNameFor(CR.Code, CR.Units) & vbTab
          End If
          Select Case CoagPrintFormat(Trim(CR.Code) & "", Trim(CR.Units) & "")
              Case 0: g = Format$(CR.Result, "0")
              Case 1: g = Format$(CR.Result, "0.0")
              Case 2: g = Format$(CR.Result, "0.00")
          End Select
          s = s & g & vbTab & UnitConv(CR.Units)
          grdPrev.AddItem s
      Next
      lblPrevCoag = PrevDate & " Result for " & txtChart
    Else
      lblPrevCoag = "No Previous Coag Details"
    End If
  Else
    lblPrevCoag = "No Previous Coag Details"
  End If

FixG grdPrev

Exit Sub

LoadPreviousCoag_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /LoadPreviousCoag ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Function LoadSplitList(ByVal Index As Integer) As String

Dim tb As New Recordset
Dim SQL As String
Dim strIndex As String
Dim strReturn As String

On Error GoTo LoadSplitList_Error

strIndex = Index

SQL = "SELECT distinct Code, PrintPriority, SplitList " & _
      "from BioTestDefinitions " & _
      "WHERE SplitList = " & strIndex & " " & _
      "order by PrintPriority"
      
Set tb = New Recordset
RecOpenClient 0, tb, SQL

strReturn = ""
Do While Not tb.EOF
  strReturn = strReturn & "Code = '" & tb!Code & "' or "
  tb.MoveNext
Loop
If strReturn <> "" Then
  strReturn = Left$(strReturn, Len(strReturn) - 3)
End If

LoadSplitList = strReturn

Exit Function

LoadSplitList_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /LoadSplitList ")
  Case 1:    End
  Case 2:    Exit Function
  Case 3:    Resume Next
End Select


End Function

Private Sub lRandom_Click()

On Error GoTo lRandom_Click_Error

If lRandom = "Random Sample" Then
  lRandom = "Fasting Sample"
Else
  lRandom = "Random Sample"
End If

LoadBiochemistry

cmdSaveBio.Enabled = True

Exit Sub

lRandom_Click_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /lRandom_Click ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub oG_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If oG.Value = 1 Then txtBioComment = Trim(txtBioComment & " " & oG.Caption)
BioChanged = True
cmdSaveBio.Enabled = True

End Sub

Private Sub oH_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

BioChanged = True
cmdSaveBio.Enabled = True
If oH.Value = 1 Then txtBioComment = Trim(txtBioComment & " " & oH.Caption)

End Sub

Private Sub oJ_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

BioChanged = True
cmdSaveBio.Enabled = True
If oJ.Value = 1 Then txtBioComment = Trim(txtBioComment & " " & oJ.Caption)

End Sub

Private Sub oL_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

BioChanged = True
cmdSaveBio.Enabled = True
If oL.Value = 1 Then txtBioComment = Trim(txtBioComment & " " & oL.Caption)

End Sub

Private Sub oO_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

BioChanged = True
cmdSaveBio.Enabled = True
If oO.Value = 1 Then txtBioComment = Trim(txtBioComment & " " & oO.Caption)

End Sub

Private Sub oS_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

BioChanged = True
cmdSaveBio.Enabled = True
If oS.Value = 1 Then txtBioComment = Trim(txtBioComment & " " & oS.Caption)

End Sub

Public Property Let PrintToPrinter(ByVal strNewValue As String)
Attribute PrintToPrinter.VB_HelpID = 182

pPrintToPrinter = strNewValue

End Property

Public Property Get PrintToPrinter() As String

PrintToPrinter = pPrintToPrinter

End Property

Private Sub SaveBiochemistry(ByVal Validate As Boolean, Optional ByVal Unval As Boolean)

Dim SQL As String
Dim tb As New Recordset



On Error GoTo SaveBiochemistry_Error

txtSampleID = Format(Val(txtSampleID))
If Val(txtSampleID) = 0 Then Exit Sub

If Validate Then
  SQL = "UPDATE BioResults set valid = 1 WHERE " & _
        "sampleid = '" & txtSampleID & "'"
  Cnxn(0).Execute SQL
ElseIf Unval = True Then
  SQL = "UPDATE BioResults set valid = 0, healthlink = 0 WHERE " & _
        "sampleid = '" & txtSampleID & "'"
  Cnxn(0).Execute SQL
End If
If Validate Then
  SQL = "UPDATE BioResults " & _
        "set operator = '" & UserCode & "' WHERE " & _
        "SampleID = '" & txtSampleID & "' "
  Cnxn(0).Execute SQL
End If

If oH Or oS Or oL Or oO Or oG Or oJ Then
  SQL = "SELECT * from Masks WHERE " & _
        "SampleID = '" & txtSampleID & "'"
  Set tb = New Recordset
  RecOpenClient 0, tb, SQL
  If tb.EOF Then tb.AddNew
  tb!SampleID = txtSampleID
  If Trim(tb!Rundate) & "" = "" Then tb!Rundate = Format$(dtRunDate, "dd/mmm/yyyy")
  tb!h = oH
  tb!s = oS
  tb!l = oL
  tb!o = oO
  tb!g = oG
  tb!J = oJ
  tb.Update
Else
  SQL = "DELETE from Masks WHERE " & _
        "SampleID = '" & txtSampleID & "'"
  Cnxn(0).Execute SQL
End If

SQL = "SELECT * from Demographics WHERE " & _
      "SampleID = '" & txtSampleID & "'"
Set tb = New Recordset
RecOpenClient 0, tb, SQL
If tb.EOF Then
  tb.AddNew
Else
  Archive 0, tb, "ArcDemographics", txtSampleID
End If
If lRandom = "Fasting Sample" Then
  tb!Fasting = 1
Else
  tb!Fasting = 0
End If
tb!Faxed = 0
tb!RooH = cRooH(0)
If Trim(tb!Rundate) & "" = "" Then tb!Rundate = Format$(dtRunDate, "dd/mmm/yyyy")
If IsDate(tSampleTime) Then
  tb!SampleDate = Format$(dtSampleDate, "dd/mmm/yyyy") & " " & Format$(tSampleTime, "hh:mm")
Else
  tb!SampleDate = Format$(dtSampleDate, "dd/mmm/yyyy")
End If
tb!SampleID = txtSampleID
tb.Update


Exit Sub

SaveBiochemistry_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /SaveBiochemistry ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub SaveBloodGas(ByVal Validate As Boolean, Optional ByVal Unval As Boolean)
Dim SQL As String
Dim tb As New Recordset



On Error GoTo SaveBloodGas_Error

txtSampleID = Format(Val(txtSampleID))
If Val(txtSampleID) = 0 Then Exit Sub

If Validate Then
  SQL = "UPDATE BgaResults set valid = 1 WHERE " & _
        "sampleid = '" & txtSampleID & "'"
  Cnxn(0).Execute SQL
ElseIf Unval = True Then
  SQL = "UPDATE BgaResults set valid = 0 WHERE " & _
        "sampleid = '" & txtSampleID & "'"
  Cnxn(0).Execute SQL
End If
If Validate Then
  SQL = "UPDATE BgaResults " & _
        "set operator = '" & UserCode & "' WHERE " & _
        "SampleID = '" & txtSampleID & "' "
  Cnxn(0).Execute SQL
End If



SQL = "SELECT * from Demographics WHERE " & _
      "SampleID = '" & txtSampleID & "'"
Set tb = New Recordset
RecOpenClient 0, tb, SQL
If tb.EOF Then
  tb.AddNew
  tb!ForESR = 0
Else
  Archive 0, tb, "ArcDemographics", txtSampleID
End If
If lRandom = "Fasting Sample" Then
  tb!Fasting = 1
Else
  tb!Fasting = 0
End If
tb!Faxed = 0
tb!RooH = cRooH(0)
If Trim(tb!Rundate) & "" = "" Then tb!Rundate = Format$(dtRunDate, "dd/mmm/yyyy")
If IsDate(tSampleTime) Then
  tb!SampleDate = Format$(dtSampleDate, "dd/mmm/yyyy") & " " & Format$(tSampleTime, "hh:mm")
Else
  tb!SampleDate = Format$(dtSampleDate, "dd/mmm/yyyy")
End If
tb!SampleID = txtSampleID
tb.Update


Exit Sub

SaveBloodGas_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /SaveBiochemistry ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select



End Sub

Private Sub SaveCoag(ByVal Validate As Boolean)

Dim SQL As String
Dim tb As New Recordset
Dim n As Long
Dim Code As String
Dim Unit As String


On Error GoTo SaveCoag_Error

txtSampleID = Format(Val(txtSampleID))
If Val(txtSampleID) = 0 Then Exit Sub

If grdCoag.Rows = 2 And grdCoag.TextMatrix(1, 0) = "" And txtCoagComment = "" Then Exit Sub

If grdCoag.Rows > 1 And grdCoag.TextMatrix(1, 0) <> "" Then
  For n = 1 To grdCoag.Rows - 1
    Code = CoagCodeFor(grdCoag.TextMatrix(n, 0))
    If grdCoag.TextMatrix(n, 0) = "INR" Then Unit = "INR" Else Unit = grdCoag.TextMatrix(n, 2)
    SQL = "SELECT * from CoagResults WHERE " & _
          "SampleID = '" & txtSampleID & "' " & _
          "and Code = '" & Trim(Code) & "' and units = '" & Unit & "'"
    Set tb = New Recordset
    RecOpenClient 0, tb, SQL
      If tb.EOF And SysOptExp(0) = False Then
        SQL = "SELECT * from CoagResults WHERE " & _
          "SampleID = '" & txtSampleID & "' " & _
          "and Code = '" & Code & "'"
          Set tb = New Recordset
          RecOpenClient 0, tb, SQL
      End If
    If tb.EOF Then
      tb.AddNew
      tb!Rundate = Format$(dtRunDate, "dd/mmm/yyyy")
      tb!RunTime = Format$(Now, "dd/mmm/yyyy hh:mm")
    Else
      Archive 0, tb, "arccoagresults", txtSampleID
    End If
    tb!Code = Trim(Code)
    tb!Result = Left(grdCoag.TextMatrix(n, 1), 6)
    tb!SampleID = txtSampleID
    tb!Units = Unit
    If Validate Then
      tb!Valid = 1
      tb!UserName = UserCode
      tb!Printed = IIf(grdCoag.TextMatrix(n, 5) = "P", 1, 0)
    ElseIf Validate = False Then
      tb!Valid = 0
      tb!healthlink = 0
      tb!UserName = UserCode
      tb!Printed = IIf(grdCoag.TextMatrix(n, 5) = "P", 1, 0)
    Else
      tb!Valid = IIf(grdCoag.TextMatrix(n, 4) = "V", 1, 0)
      tb!Printed = IIf(grdCoag.TextMatrix(n, 5) = "P", 1, 0)
    End If
    tb.Update
  Next
  tb.Close
End If

If Trim(tWarfarin) <> "" Then
  SQL = "SELECT * from HaemResults WHERE " & _
        "SampleID = '" & txtSampleID & "'"
  Set tb = New Recordset
  RecOpenClient 0, tb, SQL
  If tb.EOF Then
    tb.AddNew
    tb!SampleID = txtSampleID
  End If
  tb!Warfarin = Trim$(tWarfarin)
  tb.Update
End If

  


Exit Sub

SaveCoag_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /SaveCoag ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub SaveComments()

Dim Cx As New Comment
Dim Cxs As New Comments

On Error GoTo SaveComments_Error

txtSampleID = Format(Val(txtSampleID))
If Val(txtSampleID) = 0 Then Exit Sub

With Cx
  .lngSampleID = txtSampleID
  .Biochemistry = Trim$(txtBioComment)
  .Demographics = Trim$(txtDemographicComment)
  .Haematology = Trim$(txtHaemComment)
  .Coagulation = Trim$(txtCoagComment)
  .Immunology = Trim$(txtImmComment(1))
  .Endocrinology = Trim$(txtImmComment(0))
  .BloodGas = Trim$(txtBGaComment)
  Cxs.Save Cx
  
End With

Exit Sub

SaveComments_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /SaveComments ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub SaveDemographics()

Dim SQL As String
Dim tb As New Recordset
Dim Hosp As String

On Error GoTo SaveDemographics_Error

txtSampleID = Format(Val(txtSampleID))
If Val(txtSampleID) = 0 Then Exit Sub

SaveComments

If Trim$(tSampleTime) <> "__:__" Then
  If Not IsDate(tSampleTime) Then
    iMsg "Invalid Time", vbExclamation
    Exit Sub
  End If
End If

If InStr(lblChartNumber, "Cavan") Then
  Hosp = "Cavan"
ElseIf InStr(lblChartNumber, "Monaghan") Then
  Hosp = "Monaghan"
Else
  Hosp = Hospname(0)
End If

SQL = "SELECT * from Demographics WHERE " & _
      "SampleID = '" & EncryptN(txtSampleID) & "'"

Set tb = New Recordset
RecOpenServer 0, tb, SQL
If tb.EOF Then
  tb.AddNew
  tb!DateTimeDemographics = Format(Now, "dd/MMM/yyyy hh:mm:ss")
  If lRandom = "Fasting Sample" Then
    tb!Fasting = 1
  Else
    tb!Fasting = 0
  End If
  tb!Faxed = 0
Else
 Archive 0, tb, "arcdemographics", txtSampleID
End If

If SysOptPgp(0) Then
  If chkPgp Then
    tb!forpgp = 1
  End If
End If

tb!RooH = cRooH(0)

tb!Rundate = Format$(dtRunDate, "dd/mmm/yyyy")

If IsDate(tSampleTime) Then
  tb!SampleDate = Format$(dtSampleDate, "dd/mmm/yyyy") & " " & Format$(tSampleTime, "hh:mm")
ElseIf SysOptSampleTime(0) Then
  tb!SampleDate = Format$(dtSampleDate, "dd/mmm/yyyy") & " " & Format$(Now, "hh:mm")
Else
  tb!SampleDate = Format$(dtSampleDate, "dd/mmm/yyyy")
End If

If IsDate(tRecTime) Then
'  If Format$(dtRecDate, "yyyy/mmm/dd") <= Format$(dtSampleDate, "yyyy/mmm/dd") Then
'    tb!RecDate = Format$(dtSampleDate, "dd/mmm/yyyy") & " " & Format$(tRecTime, "hh:mm")
'  Else
    tb!RecDate = Format$(dtRecDate, "dd/mmm/yyyy") & " " & Format$(tRecTime, "hh:mm")
'  End If
Else
    tb!RecDate = Format$(dtRecDate, "dd/mmm/yyyy") & " " & Format$(Now, "hh:mm")
End If
tb!SampleID = EncryptN(txtSampleID)
tb!Chart = EncryptA(txtChart)
tb!NOPAS = Trim(txtNOPAS)
tb!aande = Trim(txtAandE)
tb!PatName = EncryptA(Trim$(txtName))
If IsDate(lDoB) Then
  tb!DoB = Format$(lDoB, "dd/mmm/yyyy")
Else
  tb!DoB = Null
End If
If cCat(0) = "Default" Then tb!Category = "" Else tb!Category = cCat(0)
If cCat(1) = "Default" Then tb!Category = "" Else tb!Category = cCat(1)
If Len(lAge) > 5 Then tb!Age = lAge
tb!sex = Left$(lSex, 1)
tb!Addr0 = EncryptA(taddress(0))
tb!Addr1 = EncryptA(taddress(1))
tb!Ward = StrConv(Left$(cmbWard, 50), vbProperCase)
tb!Clinician = Left$(cmbClinician, 50)
tb!GP = Left$(cmbGP, 50)
tb!cldetails = Left$(cClDetails, 30)
tb!Hospital = cmbHospital
tb!UserName = UserName
If SysOptUrgent(0) Then
  If chkUrgent.Value = 1 Then tb!urgent = 1 Else tb!urgent = 0
End If
tb.Update


If SysOptSpy(0) And UserMemberOf = "Secretarys" Then
  SQL = "INSERT into demospy (sampleid, patname, username, Saved, UserNom) values ('" & txtSampleID & "', '" & txtName & "', '" & UserName & "', '1', '" & UserName & "')"
  Cnxn(0).Execute SQL
End If

LogTimeOfPrinting txtSampleID, "D"

Screen.MousePointer = 0

Exit Sub

SaveDemographics_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /SaveDemographics ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub SaveEndocrinology(ByVal Validate As Boolean, Optional ByVal Unval As Boolean)

Dim SQL As String
Dim tb As New Recordset


On Error GoTo SaveEndocrinology_Error

txtSampleID = Format(Val(txtSampleID))
If Val(txtSampleID) = 0 Then Exit Sub
If Validate Then
  SQL = "UPDATE endResults set valid = 1 WHERE " & _
        "sampleid = '" & txtSampleID & "'"
  Cnxn(0).Execute SQL
ElseIf Unval = True Then
  SQL = "UPDATE endResults set valid = 0, healthlink = 0 WHERE " & _
        "sampleid = '" & txtSampleID & "'"
  Cnxn(0).Execute SQL
End If

If Validate Then
  SQL = "select * from  EndResults where " & _
        "SampleID = " & txtSampleID & " "
  Set tb = New Recordset
  RecOpenServer 0, tb, SQL
  If Not tb.EOF Then
    Do While Not tb.EOF
      Archive 0, tb, "ARCendresults", txtSampleID
      tb!Operator = UserCode
      tb.Update
      tb.MoveNext
    Loop
  End If
End If

If Ih(0) Or Iis(0) Or Il(0) Or Io(0) Or Ig(0) Or Ij(0) Then
  SQL = "SELECT * from EndMasks WHERE " & _
        "SampleID = '" & txtSampleID & "'"
  Set tb = New Recordset
  RecOpenClient 0, tb, SQL
  If tb.EOF Then tb.AddNew
  tb!SampleID = txtSampleID
  tb!Rundate = Format$(dtRunDate, "dd/mmm/yyyy")
  tb!h = Ih(0)
  tb!s = Iis(0)
  tb!l = Il(0)
  tb!o = Io(0)
  tb!g = Ig(0)
  tb!J = Ij(0)
  tb.Update
Else
  SQL = "DELETE from EndMasks WHERE " & _
        "SampleID = '" & txtSampleID & "'"
  Cnxn(0).Execute SQL
End If

SQL = "SELECT * from Demographics WHERE " & _
      "SampleID = '" & txtSampleID & "'"
Set tb = New Recordset
RecOpenClient 0, tb, SQL
If tb.EOF Then
  tb.AddNew
Else
  Archive 0, tb, "ArcDemographics", txtSampleID
End If
If lImmRan(0) = "Fasting Sample" Then
  tb!Fasting = 1
Else
  tb!Fasting = 0
End If
tb!Faxed = 0
tb!RooH = cRooH(0)
tb!Rundate = Format$(dtRunDate, "dd/mmm/yyyy")
If IsDate(tSampleTime) Then
  tb!SampleDate = Format$(dtSampleDate, "dd/mmm/yyyy") & " " & Format$(tSampleTime, "hh:mm")
Else
  tb!SampleDate = Format$(dtSampleDate, "dd/mmm/yyyy")
End If
tb!SampleID = txtSampleID
tb.Update


Exit Sub

SaveEndocrinology_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /SaveEndocrinology ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub SaveExtern()

Dim tb As New Recordset
Dim Num As Long
Dim TestNumber As Long
Dim SQL As String

On Error GoTo SaveExtern_Error

txtSampleID = Format(Val(txtSampleID))
If Val(txtSampleID) = 0 Then Exit Sub

For Num = 1 To grdExt.Rows - 1
  grdExt.Row = Num
  grdExt.Col = 0
  If Trim(grdExt) <> "" Then
    TestNumber = Val(grdExt)
    SQL = "SELECT * from extresults WHERE " & _
          "sampleid = '" & txtSampleID & "' " & _
          "and analyte = '" & TestNumber & "'"
    Set tb = New Recordset
    RecOpenServer 0, tb, SQL
    If tb.EOF Then
      tb.AddNew
    End If
    tb!SampleID = txtSampleID
    tb!Analyte = TestNumber
    grdExt.Col = 2
    tb!Result = grdExt
    grdExt.Col = 4
    tb!Units = grdExt
    grdExt.Col = 5
    tb!SendTo = grdExt
    grdExt.Col = 6
    If IsDate(grdExt) Then
      grdExt = Format(grdExt, "dd/mmm/yyyy")
      tb!SENTDate = grdExt
    Else
      tb!SENTDate = Format(Now, "dd/mmm/yyyy")
    End If
    grdExt.Col = 7
    If IsDate(grdExt) Then
      grdExt = Format(grdExt, "dd/mmm/yyyy")
      tb!RetDate = grdExt
    Else
      tb!RetDate = Null
    End If
    grdExt.Col = 8
    tb!SapCode = grdExt
    tb.Update
  End If
Next

SQL = "SELECT * from etc WHERE " & _
      "sampleid = '" & txtSampleID & "'"
Set tb = New Recordset
RecOpenServer 0, tb, SQL
If tb.EOF Then
  tb.AddNew
  tb!SampleID = txtSampleID
End If
tb!etc0 = txtEtc(0)
tb!etc1 = txtEtc(1)
tb!etc2 = txtEtc(2)
tb!etc3 = txtEtc(3)
tb!etc4 = txtEtc(4)
tb!etc5 = txtEtc(5)
tb!etc6 = txtEtc(6)
tb!etc7 = txtEtc(7)
tb!etc8 = txtEtc(8)
tb.Update

Exit Sub

SaveExtern_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /SaveExtern ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub SaveHaematology(ByVal Validate As Boolean)

Dim tb As New Recordset
Dim SQL As String



On Error GoTo SaveHaematology_Error

txtSampleID = Format(Val(txtSampleID))
If Val(txtSampleID) = 0 Then Exit Sub

If Trim$(txtSampleID) = "" Then
  iMsg "Must have Sample ID Number.", vbCritical
  Exit Sub
End If
  


SQL = "SELECT * from HaemResults WHERE " & _
      "SampleID = '" & txtSampleID & "'"
Set tb = New Recordset
RecOpenServer 0, tb, SQL
If tb.EOF Then
  tb.AddNew
  tb!Rundate = Format$(dtRunDate, "dd/mmm/yyyy")
  tb!RunDateTime = Format$(Now, "dd/mmm/yyyy hh:mm")
  tb!SampleID = txtSampleID
  tb!Faxed = 0
  tb!Printed = 0
Else
  Archive 0, tb, "archaemresults", txtSampleID
End If

tb!rbc = gRbc.TextMatrix(1, 1)
tb!Hgb = gRbc.TextMatrix(2, 1)
tb!Hct = gRbc.TextMatrix(3, 1)
tb!MCV = gRbc.TextMatrix(4, 1)
tb!hdw = gRbc.TextMatrix(5, 1)
tb!mch = gRbc.TextMatrix(6, 1)
tb!mchc = gRbc.TextMatrix(7, 1)
tb!cH = gRbc.TextMatrix(8, 1)
tb!RDWCV = gRbc.TextMatrix(9, 1)
tb!nrbcp = gRbc.TextMatrix(10, 1)
tb!Hyp = gRbc.TextMatrix(11, 1)
tb!Plt = tPlt
tb!mpv = tMPV
tb!wbc = tWBC
tb!LymA = grdH.TextMatrix(2, 0)
tb!LymP = grdH.TextMatrix(2, 3)
tb!MonoA = grdH.TextMatrix(3, 0)
tb!MonoP = grdH.TextMatrix(3, 3)
tb!NeutA = grdH.TextMatrix(1, 0)
tb!NeutP = grdH.TextMatrix(1, 3)
tb!EosA = grdH.TextMatrix(4, 0)
tb!EosP = grdH.TextMatrix(4, 3)
tb!BasA = grdH.TextMatrix(5, 0)
tb!BasP = grdH.TextMatrix(5, 3)
tb!luca = grdH.TextMatrix(6, 0)
tb!lucp = grdH.TextMatrix(6, 3)


tb!esr = tESR

If SysOptESR1(0) Then
  tb!esr1 = txtEsr1
End If

tb!reta = Format(tRetA, "###.0")
tb!retp = Trim(tRetP)
tb!Monospot = Left$(tMonospot, 1)
tb!tASOt = tASOt
tb!tRa = tRa

tb!cESR = cESR = 1
tb!cRetics = cRetics = 1
tb!cMonospot = cMonospot = 1
tb!cRA = cRA = 1
tb!cASot = cASot = 1
tb!cMalaria = chkMalaria = 1
tb!csickledex = chkSickledex = 1

tb!malaria = lblMalaria
tb!sickledex = lblSickledex

If SysOptBadRes(0) Then
  tb!cbad = chkBad = 1
End If

tb!ccoag = 0
tb!cFilm = cFilm = 1

tb!Warfarin = tWarfarin

If Validate Then
  tb!Valid = 1
Else
  tb!Valid = 0
  tb!healthlink = 0
End If
tb!Operator = UserCode

tb.Update


  If Trim(txtCondition) <> "" Then
    SQL = "SELECT * from HaemCondition WHERE " & _
          "chart = '" & txtChart & "'"
    Set tb = New Recordset
    RecOpenClient 0, tb, SQL
    If tb.EOF Then tb.AddNew
      tb!Chart = txtChart
      tb!condition = Trim(txtCondition)
      tb.Update
  End If

tb.Close
Set tb = Nothing

Screen.MousePointer = 0


Exit Sub

SaveHaematology_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /SaveHaematology ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub SaveImmunology(ByVal Validate As Boolean, Optional ByVal Unval As Boolean)

Dim SQL As String
Dim tb As New Recordset


On Error GoTo SaveImmunology_Error

txtSampleID = Format(Val(txtSampleID))
If Val(txtSampleID) = 0 Then Exit Sub

If Validate Then
  SQL = "UPDATE ImmResults set valid = 1, operator = '" & UserCode & "' WHERE " & _
        "sampleid = '" & txtSampleID & "' "
  Cnxn(0).Execute SQL
ElseIf Unval = True Then
  SQL = "UPDATE ImmResults set valid = 0, healthlink = 0, operator = '" & UserCode & "' WHERE " & _
        "sampleid = '" & txtSampleID & "'"
  Cnxn(0).Execute SQL
End If

If Validate Then
  SQL = "UPDATE ImmResults " & _
        "set operator = '" & UserCode & "' WHERE " & _
        "SampleID = '" & txtSampleID & "' " & _
        "and operator is null "
  Cnxn(0).Execute SQL
End If

If Ih(1) Or Iis(1) Or Il(1) Or Io(1) Or Ig(1) Or Ij(1) Then
  SQL = "SELECT * from ImmMasks WHERE " & _
        "SampleID = '" & txtSampleID & "'"
  Set tb = New Recordset
  RecOpenClient 0, tb, SQL
  If tb.EOF Then tb.AddNew
  tb!SampleID = txtSampleID
  tb!Rundate = Format$(dtRunDate, "dd/mmm/yyyy")
  tb!h = Ih(1)
  tb!s = Iis(1)
  tb!l = Il(1)
  tb!o = Io(1)
  tb!g = Ig(1)
  tb!J = Ij(1)
  tb.Update
Else
  SQL = "DELETE from ImmMasks WHERE " & _
        "SampleID = '" & txtSampleID & "'"
  Cnxn(0).Execute SQL
End If

SQL = "SELECT * from Demographics WHERE " & _
      "SampleID = '" & txtSampleID & "'"
Set tb = New Recordset
RecOpenClient 0, tb, SQL
If tb.EOF Then
  tb.AddNew
Else
  Archive 0, tb, "ArcDemographics", txtSampleID
End If
If lImmRan(1) = "Fasting Sample" Then
  tb!Fasting = 1
Else
  tb!Fasting = 0
End If
tb!Faxed = 0
tb!RooH = cRooH(0)
tb!Rundate = Format$(dtRunDate, "dd/mmm/yyyy")
If IsDate(tSampleTime) Then
  tb!SampleDate = Format$(dtSampleDate, "dd/mmm/yyyy") & " " & Format$(tSampleTime, "hh:mm")
Else
  tb!SampleDate = Format$(dtSampleDate, "dd/mmm/yyyy")
End If
tb!SampleID = txtSampleID
tb.Update


Exit Sub

SaveImmunology_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /SaveImmunology ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub Set_Demo(ByVal Demo As Boolean)

Frame4.Enabled = Demo
Frame5.Enabled = Demo
Frame7.Enabled = Demo
ssPanPgP.Enabled = Demo
Frame10(0).Enabled = Demo
txtChart.Locked = Not Demo
txtAandE.Locked = Not Demo
txtNOPAS.Locked = Not Demo
txtName.Locked = Not Demo
txtDoB.Locked = Not Demo
txtAge.Locked = Not Demo
txtSex.Locked = Not Demo

If Demo = False Then
  StatusBar1.Panels(3).Text = "Demographics Validated"
  StatusBar1.Panels(3).Bevel = sbrInset
Else
  StatusBar1.Panels(3).Text = "Check Demographics"
  StatusBar1.Panels(3).Bevel = sbrRaised
End If
End Sub

Private Sub SetViewHistory()

On Error GoTo SetViewHistory_Error

Select Case sstabAll.Tab
  Case 0: bHistory.Visible = False
  Case 1: bHistory.Visible = HistHaem
  Case 2: bHistory.Visible = HistBio
  Case 3: bHistory.Visible = HistCoag
  Case 4: bHistory.Visible = HistEnd
  Case 5: bHistory.Visible = HistBga
  Case 6: bHistory.Visible = HistImm
  Case 7: bHistory.Visible = HistExt
End Select


Exit Sub

SetViewHistory_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /SetViewHistory ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub SetWardClinGp()

lAddWardGP = Trim$(taddress(0)) & " " & Trim$(taddress(1)) & " : " & cmbWard & " : " & cmbGP & " " & cmbClinician
If Hospname(0) = "STJOHNS" And cmbClinician <> "" Then
  lAddWardGP = Trim$(taddress(0)) & " : " & cmbWard & " : " & cmbClinician
End If

End Sub

Private Sub sstabAll_Click(PreviousTab As Integer)

On Error GoTo sstabAll_Click_Error

Select Case PreviousTab
  Case 0
    If cmdSaveDemographics.Enabled Then
      If iMsg("Demographic Details have changed!" & vbCrLf & "Save?", vbQuestion + vbYesNo) = vbYes Then
        cmdSaveDemographics_Click
      End If
    End If
  Case 1
    If cmdSaveHaem.Enabled Then
      If iMsg("Haematology Details have changed!" & vbCrLf & "Save?", vbQuestion + vbYesNo) = vbYes Then
        cmdSaveHaem_Click
      End If
    End If
'    If cmdSaveComm.Enabled Then
'      If iMsg("Haematology Comments have changed!" & vbCrLf & "Save?", vbQuestion + vbYesNo) = vbYes Then
'        cmdSaveComm_Click
'      End If
'    End If
  Case 2
    If cmdSaveBio.Enabled Then
      If iMsg("Biochemistry Details have changed!" & vbCrLf & "Save?", vbQuestion + vbYesNo) = vbYes Then
        cmdSaveBio_Click
      End If
    End If
  Case 3
    If cmdSaveCoag.Enabled Then
      If iMsg("Coagulation Details have changed!" & vbCrLf & "Save?", vbQuestion + vbYesNo) = vbYes Then
        cmdSaveCoag_Click
      End If
    End If
  Case 4
    If cmdSaveImm(0).Enabled Then
      If iMsg("Endocrinology Details have changed!" & vbCrLf & "Save?", vbQuestion + vbYesNo) = vbYes Then
        cmdSaveImm_Click (0)
      End If
    End If
  Case 5
    If cmdSaveBGa.Enabled Then
      If iMsg("Blood Gas Details have changed!" & vbCrLf & "Save?", vbQuestion + vbYesNo) = vbYes Then
        cmdSaveBGa_Click
      End If
    End If
  Case 6
    If cmdSaveImm(1).Enabled Then
      If iMsg("Immunology Details have changed!" & vbCrLf & "Save?", vbQuestion + vbYesNo) = vbYes Then
        cmdSaveImm_Click (1)
      End If
    End If
  Case 7
    If cmdSaveExt.Enabled Then
      If iMsg("External Details have changed!" & vbCrLf & "Save?", vbQuestion + vbYesNo) = vbYes Then
        cmdSaveExt_Click
      End If
    End If
End Select


cmdPrint.Visible = True
cmdPrintHold.Visible = True
bFAX.Visible = True
cmdSetPrinter.Visible = True
Select Case sstabAll.Tab
  Case 0: 'Demographics
   cmdPrint.Visible = False
   If SysOptAllowDemoPrint(0) = False Then cmdPrintHold.Visible = False
   bFAX.Visible = False
   cmdSetPrinter.Visible = False
  Case 1: 'Haematology
    If Not HaemLoaded Then
      LoadHaematology
      HaemLoaded = True
    ElseIf bValidateHaem.Caption = "VALID" Then
      lblUrgent.Visible = False
    ElseIf bValidateHaem.Caption <> "VALID" Then
      lblUrgent.Visible = UrgentTest
    End If
    
  Case 2: 'Biochemistry
    If Not BioLoaded Then
      LoadBiochemistry
      BioLoaded = True
    ElseIf bValidateBio.Caption = "VALID" Then
      lblUrgent.Visible = False
    ElseIf bValidateBio.Caption <> "VALID" Then
      lblUrgent.Visible = UrgentTest
    End If
    
  Case 3: 'Coagulation
    If Not CoagLoaded Then
      LoadCoagulation
      CoagLoaded = True
    ElseIf cmdValidateCoag.Caption = "VALID" Then
      lblUrgent.Visible = False
    ElseIf cmdValidateCoag.Caption <> "VALID" Then
      lblUrgent.Visible = UrgentTest
    End If
    
  Case 4: 'Endocrinology
    If Not EndLoaded Then
      LoadEndocrinology
      EndLoaded = True
    ElseIf bValidateImm(0).Caption = "VALID" Then
      lblUrgent.Visible = False
    ElseIf bValidateImm(0).Caption <> "VALID" Then
      lblUrgent.Visible = UrgentTest
    End If

  Case 5: 'Biochemistry
    If Not BgaLoaded Then
      LoadBloodGas
      BgaLoaded = True
    ElseIf cmdValBG.Caption = "VALID" Then
      lblUrgent.Visible = False
    ElseIf cmdValBG.Caption <> "VALID" Then
      lblUrgent.Visible = UrgentTest
    End If
   
 Case 6: 'Immunology
    If Not ImmLoaded Then
      LoadImmunology
      ImmLoaded = True
    ElseIf bValidateImm(1).Caption = "VALID" Then
      lblUrgent.Visible = False
    ElseIf bValidateImm(1).Caption <> "VALID" Then
      lblUrgent.Visible = UrgentTest
    End If

 Case 7:
   bFAX.Visible = False
   If Not ExtLoaded Then
    LoadExt
    ExtLoaded = True
   End If
   

End Select

SetViewHistory

Exit Sub

sstabAll_Click_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /sstabAll_Click ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub sstabAll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

pBar = 0

End Sub

Private Sub StatusBar1_PanelClick(ByVal Panel As MSComctlLib.Panel)

If Panel <> "" Then
  If Panel.Index = 3 And Panel.Bevel = sbrRaised Then
    sstabAll.Tab = 0
  End If
End If

End Sub

Private Sub taddress_Change(Index As Integer)

SetWardClinGp

End Sub

Private Sub taddress_KeyPress(Index As Integer, KeyAscii As Integer)

cmdSaveDemographics.Enabled = True
cmdSaveInc.Enabled = True

End Sub

Private Sub taddress_LostFocus(Index As Integer)

taddress(Index) = StrConv(taddress(Index), vbProperCase)

End Sub

Private Sub tASOt_Change()
If Trim$(tASOt) <> "" Then
  cASot = 1
Else
  cASot = 0
End If

End Sub

Private Sub tasot_Click()

On Error GoTo tasot_Click_Error

cmdSaveHaem.Enabled = True: cmdHSaveH.Enabled = True

If Trim$(tASOt) = "" Or tASOt = "?" Then
  tASOt = "Negative"
ElseIf tASOt = "Negative" Then
  tASOt = "Positive"
Else
  tASOt = ""
End If

Exit Sub

tasot_Click_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /tasot_Click ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub tasot_KeyPress(KeyAscii As Integer)

On Error GoTo tasot_KeyPress_Error

cmdSaveHaem.Enabled = True: cmdHSaveH.Enabled = True

If Trim$(tASOt) = "" Then
  tASOt = "Negative"
ElseIf tASOt = "Negative" Then
  tASOt = "Positive"
Else
  tASOt = ""
End If

Exit Sub

tasot_KeyPress_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /tasot_KeyPress ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select

End Sub

Private Sub tESR_Change()

If Trim$(tESR) <> "" Then
  cESR = 1
Else
  cESR = 0
End If

cmdSaveHaem.Enabled = True: cmdHSaveH.Enabled = True: cmdSaveComm.Enabled = True

End Sub

Private Sub tESR_KeyDown(KeyCode As Integer, Shift As Integer)

If SysOptESR1(0) Then
  If KeyCode = vbKeyF2 Then
    txtEsr1.Visible = True
  End If
End If

End Sub

Private Sub tESR_KeyPress(KeyAscii As Integer)

cmdSaveHaem.Enabled = True: cmdHSaveH.Enabled = True



End Sub

Private Sub tESR_LostFocus()

If tESR = "" Then Exit Sub

If tESR <> "?" Then
  If Not IsNumeric(tESR) Then
    iMsg "Result must be numeric"
    tESR = "?"
    Exit Sub
  End If
End If

End Sub

Private Sub TimerBar_Timer()

pBar = pBar + 1

'code added 22/08/2005
'not live
'If pBar = pBar.max / 2 Then
'  txtSampleID_LostFocus
'  Exit Sub
'End If
  
If pBar = pBar.max Then
  Unload Me
  Exit Sub
End If

End Sub

Private Sub tINewValue_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim SQL As String
Dim tb As Recordset
Dim s As String

If KeyCode = 113 And Index = 1 Then
    SQL = "SELECT * from lists WHERE listtype = 'IR' and code = '" & tINewValue(1) & "'"
    Set tb = New Recordset
    RecOpenServer 0, tb, SQL
    If Not tb.EOF Then
        tINewValue(1) = Trim(tb!Text)
        tINewValue(1).SelStart = Len(tINewValue(1)) + 1
    End If
ElseIf KeyCode = 114 And Index = 1 Then
    SQL = "SELECT * from lists WHERE listtype = 'IR'"
    Set tb = New Recordset
    RecOpenServer 0, tb, SQL
    Do While Not tb.EOF
        s = Trim(tb!Text)
       frmMessages.lstComm.AddItem s
          tb.MoveNext
    Loop
    
    Set frmMessages.F = Me
    Set frmMessages.T = tINewValue(1)
    frmMessages.Show 1
    tINewValue(1).SelStart = Len(tINewValue(1)) + 1
End If


End Sub

Private Sub tINewValue_LostFocus(Index As Integer)

If Not IsNumeric(tINewValue(Index)) Then
  tINewValue(Index) = Trim(tINewValue(Index))
End If

End Sub

Private Sub tMonospot_Change()

If Trim$(tMonospot) <> "" Then
  cMonospot = 1
Else
  cMonospot = 0
End If

End Sub

Private Sub tMonospot_Click()

cmdSaveHaem.Enabled = True: cmdHSaveH.Enabled = True

If Trim$(tMonospot) = "" Or tMonospot = "?" Then
  tMonospot = "Negative"
ElseIf tMonospot = "Negative" Then
  tMonospot = "Positive"
ElseIf tMonospot = "Positive" Then
  tMonospot = "Inconclusive"
Else
  tMonospot = ""
End If

End Sub

Private Sub tMonospot_KeyPress(KeyAscii As Integer)

cmdSaveHaem.Enabled = True: cmdHSaveH.Enabled = True

If Trim$(tMonospot) = "" Then
  tMonospot = "Negative"
ElseIf tMonospot = "Negative" Then
  tMonospot = "Positive"
Else
  tMonospot = ""
End If

End Sub

Private Sub tMPV_KeyPress(KeyAscii As Integer)

cmdSaveHaem.Enabled = True: cmdHSaveH.Enabled = True

End Sub

Private Sub tnewvalue_Click()

If InStr(UCase(cAdd), "PREG") > 0 Then
  If tnewvalue = "" Then
    tnewvalue = "Neg"
  ElseIf tnewvalue = "Neg" Then
    tnewvalue = "Pos"
  ElseIf tnewvalue = "Pos" Then
    tnewvalue = "WKPos"
  ElseIf tnewvalue = "WKPos" Then
    tnewvalue = "STPos"
  ElseIf tnewvalue = "STPos" Then
    tnewvalue = "Equiv"
  ElseIf tnewvalue = "Equiv" Then
    tnewvalue = ""
  End If
End If

End Sub

Private Sub tPlt_KeyPress(KeyAscii As Integer)

cmdSaveHaem.Enabled = True: cmdHSaveH.Enabled = True

End Sub

Private Sub tRa_Change()

If Trim$(tRa) <> "" Then
  cRA = 1
Else
  cRA = 0
End If

End Sub

Private Sub tRa_Click()

cmdSaveHaem.Enabled = True: cmdHSaveH.Enabled = True

If Trim$(tRa) = "" Or tRa = "?" Then
  tRa = "Negative"
ElseIf tRa = "Negative" Then
  tRa = "Positive"
Else
  tRa = ""
End If

End Sub

Private Sub tRa_KeyPress(KeyAscii As Integer)

cmdSaveHaem.Enabled = True: cmdHSaveH.Enabled = True

If Trim$(tRa) = "" Then
  tRa = "Negative"
ElseIf tRa = "Negative" Then
  tRa = "Positive"
Else
  tRa = ""
End If

End Sub

Private Sub tRecTime_Change()

cmdSaveDemographics.Enabled = True
cmdSaveInc.Enabled = True

End Sub

Private Sub tRetA_Change()

If Trim$(tRetA) <> "" Then
  cRetics = 1
Else
  cRetics = 0
End If

End Sub

Private Sub tRetA_KeyPress(KeyAscii As Integer)

cmdSaveHaem.Enabled = True: cmdHSaveH.Enabled = True

End Sub

Private Sub tRetA_LostFocus()

If tRetA = "" Then Exit Sub

If tRetA <> "?" Then
  If Not IsNumeric(tRetA) Then
    iMsg "Result must be numeric"
    tRetA = "?"
    Exit Sub
  End If
End If

End Sub

Private Sub tRetP_Change()

If Trim$(tRetP) <> "" Then
  cRetics = 1
Else
  cRetics = 0
End If

End Sub

Private Sub tRetP_KeyPress(KeyAscii As Integer)

cmdSaveHaem.Enabled = True: cmdHSaveH.Enabled = True

End Sub

Private Sub tRetP_LostFocus()

If tRetP = "" Then Exit Sub

If tRetP <> "?" Then
  If Not IsNumeric(tRetP) Then
    iMsg "Result must be numeric"
    tRetP = "?"
    Exit Sub
  End If
End If

End Sub

Private Sub tSampleTime_KeyPress(KeyAscii As Integer)

cmdSaveDemographics.Enabled = True
cmdSaveInc.Enabled = True

End Sub

Private Sub tWBC_KeyPress(KeyAscii As Integer)

cmdSaveHaem.Enabled = True: cmdHSaveH.Enabled = True

End Sub

Private Sub txtAandE_Lostfocus()


LoadPatientFromAandE Me, True

If Trim(txtName) = "" Then
  LoadDemo txtAandE
End If

txtAandE = UCase(txtAandE)

End Sub

Private Sub txtage_Change()

lAge = txtAge

End Sub

Private Sub txtAge_KeyPress(KeyAscii As Integer)

If txtAge.Locked Then Exit Sub

cmdSaveDemographics.Enabled = True
cmdSaveInc.Enabled = True

End Sub

Private Sub txtBGaComment_KeyDown(KeyCode As Integer, Shift As Integer)
Dim s As Variant
Dim n As Long
Dim tb As New Recordset
Dim SQL As String

If KeyCode = 113 Then

If Len(txtBGaComment) < 2 Then Exit Sub

n = txtBGaComment.SelStart

s = UCase(Mid(txtBGaComment, n - 1, 2))

'For n = 0 To UBound(s)
  If ListText("BG", s) <> "" Then
    s = ListText("BG", s)
  End If
'Next

txtBGaComment = Left(txtBGaComment, n - 2)
txtBGaComment = txtBGaComment & s

txtBGaComment.SelStart = Len(txtBGaComment)

ElseIf KeyCode = 114 Then
  SQL = "SELECT * from lists WHERE listtype = 'BG'"
  Set tb = New Recordset
  RecOpenServer 0, tb, SQL
  Do While Not tb.EOF
      s = Trim(tb!Text)
     frmMessages.lstComm.AddItem s
    tb.MoveNext
  Loop
  Set frmMessages.F = Me
  Set frmMessages.T = txtBGaComment
  frmMessages.Show 1

End If

cmdSaveBGa.Enabled = True

End Sub

Private Sub txtBioComment_Change()

'If bValidateBio.Caption = "VALID" Then Exit Sub

cmdSaveBio.Enabled = True

End Sub

Private Sub txtBioComment_KeyDown(KeyCode As Integer, Shift As Integer)
Dim SQL As String
Dim tb As New Recordset
Dim s As Variant
Dim n As Long

'If bValidateBio.Caption = "VALID" Then Exit Sub


If KeyCode = 113 Then

If Len(txtBioComment) < 2 Then Exit Sub

n = txtBioComment.SelStart

s = UCase(Mid(txtBioComment, n - 1, 2))

If ListText("BI", s) <> "" Then
  s = ListText("BI", s)
End If

txtBioComment = Left(txtBioComment, n - 2)
txtBioComment = txtBioComment & s

txtBioComment.SelStart = Len(txtBioComment)

ElseIf KeyCode = 114 Then
  
  SQL = "SELECT * from lists WHERE listtype = 'BI'"
  Set tb = New Recordset
  RecOpenServer 0, tb, SQL
  Do While Not tb.EOF
      s = Trim(tb!Text)
     frmMessages.lstComm.AddItem s
        tb.MoveNext
  Loop
  
  Set frmMessages.F = Me
  Set frmMessages.T = txtBioComment
  frmMessages.Show 1

End If

cmdSaveBio.Enabled = True

End Sub

Private Sub txtBioComment_KeyPress(KeyAscii As Integer)

'If bValidateBio.Caption = "VALID" Then Exit Sub

cmdSaveBio.Enabled = True

End Sub

Private Sub txtchart_Change()


lChart = txtChart

End Sub

Private Sub txtChart_KeyPress(KeyAscii As Integer)

If txtChart.Locked Then Exit Sub

cmdSaveDemographics.Enabled = True
cmdSaveInc.Enabled = True

End Sub

Private Sub txtchart_LostFocus()

If txtChart.Locked Then Exit Sub

If Trim$(txtChart) = "" Then Exit Sub
If Trim$(txtName) <> "" Then Exit Sub

txtChart = UCase(txtChart)

LoadPatientFromChart Me, True

If txtName <> "" Then
  LoadDemo txtChart
End If

End Sub

Private Sub txtCoagComment_Change()

'If cmdValidateCoag.Caption = "VALID" Then Exit Sub

cmdSaveCoag.Enabled = True

End Sub

Private Sub txtCoagComment_KeyDown(KeyCode As Integer, Shift As Integer)
Dim SQL As String
Dim tb As New Recordset
Dim s As Variant
Dim n As Long
Dim T As String
Dim z As Integer

'allow 3 chars

On Error GoTo txtCoagComment_KeyDown_Error

'If cmdValidateCoag.Caption = "VALID" Then Exit Sub

If KeyCode = 113 Then

  n = txtCoagComment.SelStart
  
      z = 2
      s = Mid(txtCoagComment, (n - z), z + 1)
      z = 3
  If ListText("CO", s) <> "" Then
    s = ListText("CO", s)
  Else
    s = ""
  End If
  
  If s = "" Then
      z = 1
      s = Mid(txtCoagComment, (n - z), z + 1)
      z = 2
    If ListText("CO", s) <> "" Then
        s = ListText("CO", s)
    Else
        s = ""
    End If
  End If
  
  If s = "" Then
    z = 1
    s = Mid(txtCoagComment, n, z + 1)
    
    If ListText("CO", s) <> "" Then
      s = ListText("CO", s)
    End If
  End If
    
  txtCoagComment = Left(txtCoagComment, (n - (z)))
  txtCoagComment = txtCoagComment & s
  
  txtCoagComment.SelStart = Len(txtCoagComment)

ElseIf KeyCode = 114 Then
  
  SQL = "SELECT * from lists WHERE listtype = 'CO'"
  Set tb = New Recordset
  RecOpenServer 0, tb, SQL
  Do While Not tb.EOF
      s = Trim(tb!Text)
     frmMessages.lstComm.AddItem s
    tb.MoveNext
  Loop
  
  Set frmMessages.F = Me
  Set frmMessages.T = txtCoagComment
  frmMessages.Show 1
  cmdSaveCoag.Enabled = True
  
End If

Exit Sub

txtCoagComment_KeyDown_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /txtCoagComment_KeyDown ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub txtCoagComment_KeyPress(KeyAscii As Integer)

'If cmdValidateCoag.Caption = "VALID" Then Exit Sub

cmdSaveCoag.Enabled = True

End Sub

Private Sub txtCondition_KeyPress(KeyAscii As Integer)

cmdSaveHaem.Enabled = True: cmdHSaveH.Enabled = True

End Sub

Private Sub txtDemographicComment_KeyDown(KeyCode As Integer, Shift As Integer)
Dim s As Variant
Dim n As Long
Dim SQL As String
Dim tb As New Recordset
Dim z As Long



On Error GoTo txtDemographicComment_KeyDown_Error

If KeyCode = 113 Then

If txtDemographicComment = "" Then Exit Sub

If Len(txtDemographicComment) < 2 Then Exit Sub

n = txtDemographicComment.SelStart

     z = 2
      s = Mid(txtDemographicComment, (n - z) + 1, z + 1)
      z = 3
  If ListText("DE", s) <> "" Then
    s = ListText("DE", s)
  Else
    s = ""
  End If
  
  If s = "" Then
      z = 1
      s = Mid(txtDemographicComment, n - z, z + 1)
      z = 2
    If ListText("DE", s) <> "" Then
        s = ListText("DE", s)
    Else
        s = ""
    End If
  End If
  
  If s = "" Then
    z = 1
    s = Mid(txtDemographicComment, n, z)
    
    If ListText("DE", s) <> "" Then
      s = ListText("DE", s)
    End If
  End If
    
  txtDemographicComment = Left(txtDemographicComment, (n - (z)) + 1)
  txtDemographicComment = txtDemographicComment & s
  
  txtDemographicComment.SelStart = Len(txtDemographicComment)

    cmdSaveDemographics.Enabled = True
    cmdSaveInc.Enabled = True

ElseIf KeyCode = 114 Then
  
  SQL = "SELECT * from lists WHERE listtype = 'DE'"
  Set tb = New Recordset
  RecOpenServer 0, tb, SQL
  Do While Not tb.EOF
      s = Trim(tb!Text)
     frmMessages.lstComm.AddItem s
    tb.MoveNext
  Loop
  
  Set frmMessages.F = frmEditAll
  Set frmMessages.T = txtDemographicComment
  frmMessages.Show 1
  

    cmdSaveDemographics.Enabled = True
    cmdSaveInc.Enabled = True

End If

Exit Sub

txtDemographicComment_KeyDown_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /txtDemographicComment_KeyDown ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub txtDemographicComment_KeyPress(KeyAscii As Integer)

cmdSaveDemographics.Enabled = True
cmdSaveInc.Enabled = True

End Sub

Private Sub txtDemographicComment_LostFocus()

txtDemographicComment = initial2upper(txtDemographicComment)
lblDemographicComment = txtDemographicComment

End Sub

Private Sub txtDoB_Change()

lDoB = txtDoB

End Sub

Private Sub txtDoB_KeyPress(KeyAscii As Integer)
 
If txtDoB.Locked Then Exit Sub

cmdSaveDemographics.Enabled = True
cmdSaveInc.Enabled = True

End Sub

Private Sub txtDoB_LostFocus()


On Error GoTo txtDoB_LostFocus_Error

If txtDoB.Locked Then Exit Sub

txtDoB = Convert62Date(txtDoB, BACKWARD)

If Not IsDate(txtDoB) Then
  txtDoB = ""
  Exit Sub
End If

txtAge = CalcAge(txtDoB)

If txtAge = "" And txtDoB <> "" Then
  GoTo txtDoB_LostFocus_Error
End If



Exit Sub

txtDoB_LostFocus_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /txtDoB_LostFocus ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub txtEsr1_Change()

cmdSaveHaem.Enabled = True: cmdHSaveH.Enabled = True: cmdSaveComm.Enabled = True

End Sub

Private Sub txtEtc_Change(Index As Integer)

cmdSaveExt.Enabled = True

End Sub

Private Sub txtHaemComment_Change()

pBar = 0
cmdSaveHaem.Enabled = True: cmdHSaveH.Enabled = True: cmdSaveComm.Enabled = True

End Sub

Private Sub txtHaemComment_KeyDown(KeyCode As Integer, Shift As Integer)
Dim s As Variant
Dim n As Long
Dim z As Long
Dim tb As New Recordset
Dim SQL As String


'If bValidateHaem.Caption = "VALID" Then Exit Sub


On Error GoTo txtHaemComment_KeyDown_Error

If KeyCode = 113 Then

  If txtHaemComment = "" Then Exit Sub '
  
  n = txtHaemComment.SelStart
  
      z = 2
      s = Mid(txtHaemComment, (n - z), z + 1)
      z = 3
  If ListText("HA", s) <> "" Then
    s = ListText("HA", s)
  Else
    s = ""
  End If
  
  If s = "" Then
      z = 1
      s = Mid(txtHaemComment, (n - z), z + 1)
      z = 2
    If ListText("HA", s) <> "" Then
        s = ListText("HA", s)
    Else
        s = ""
    End If
  End If
  
  If s = "" Then
    z = 1
    s = Mid(txtHaemComment, n, z + 1)
    
    If ListText("HA", s) <> "" Then
      s = ListText("HA", s)
    End If
  End If
    
  txtHaemComment = Left(txtHaemComment, (n - (z)))
  txtHaemComment = txtHaemComment & s
  
  txtHaemComment.SelStart = Len(txtHaemComment)
  
ElseIf KeyCode = 114 Then
  
  SQL = "SELECT * from lists WHERE listtype = 'HA'"
  Set tb = New Recordset
  RecOpenServer 0, tb, SQL
  Do While Not tb.EOF
      s = Trim(tb!Text)
     frmMessages.lstComm.AddItem s
    tb.MoveNext
  Loop
  
  Set frmMessages.F = Me
  Set frmMessages.T = txtHaemComment
  frmMessages.Show 1

End If

Exit Sub

txtHaemComment_KeyDown_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /txtHaemComment_KeyDown ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub txtHaemComment_KeyPress(KeyAscii As Integer)

'If bValidateHaem.Caption = "VALID" Then Exit Sub

cmdSaveHaem.Enabled = True: cmdHSaveH.Enabled = True

End Sub

Private Sub txtImmComment_Change(Index As Integer)

'If bValidateImm(Index).Caption = "VALID" Then Exit Sub

If Index = 0 Then
  cmdSaveImm(0).Enabled = True
Else
  cmdSaveImm(1).Enabled = True
End If
End Sub

Private Sub txtImmComment_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim SQL As String
Dim tb As New Recordset
Dim s As Variant
Dim n As Long
Dim z As Long

On Error GoTo txtImmComment_KeyDown_Error

'If bValidateImm(Index).Caption = "VALID" Then Exit Sub


If Index = 0 Then
If KeyCode = 113 Then
'If txtImmComment(0) = "" Then Exit Sub

If Len(txtImmComment(0)) < 2 Then Exit Sub

n = txtImmComment(0).SelStart

     z = 3
      s = Mid(txtImmComment(0), (n - z) + 1, z + 1)
      z = 3
  If ListText("EN", s) <> "" Then
    s = ListText("EN", s)
  Else
    s = ""
  End If
  
  If s = "" Then
      z = 1
      s = Mid(txtImmComment(0), n - z, z + 1)
      z = 2
    If ListText("EN", s) <> "" Then
        s = ListText("EN", s)
    Else
        s = ""
    End If
  End If
  
  If s = "" Then
    z = 1
    s = Mid(txtImmComment(0), n, z)
    
    If ListText("EN", s) <> "" Then
      s = ListText("EN", s)
    End If
  End If
    
  txtImmComment(0) = Left(txtImmComment(0), (n - (z)))
  txtImmComment(0) = txtImmComment(0) & s
  
  txtImmComment(0).SelStart = Len(txtImmComment(0))

ElseIf KeyCode = 114 Then
  
  SQL = "SELECT * from lists WHERE listtype = 'EN'"
  Set tb = New Recordset
  RecOpenServer 0, tb, SQL
  Do While Not tb.EOF
      s = Trim(tb!Text)
     frmMessages.lstComm.AddItem s
        tb.MoveNext
  Loop
  
  Set frmMessages.F = Me
  Set frmMessages.T = txtImmComment(0)
  frmMessages.Show 1

End If

'  If KeyCode = 113 Then
'
'  If Len(txtImmComment(0)) < 2 Then Exit Sub
'
'  n = txtImmComment(0).SelStart
'
'  s = UCase(Mid(txtImmComment(0), n - 1, 2))
'
'  If ListText("EN", s) <> "" Then
'    s = ListText("EN", s)
'  End If
'
'  txtImmComment(0) = Left(txtImmComment(0), n - 2)
'  txtImmComment(0) = txtImmComment(0) & s
'
'  txtImmComment(0).SelStart = Len(txtImmComment(0))
'
'  ElseIf KeyCode = 114 Then
'
'    SQL = "SELECT * from lists WHERE listtype = 'EN'"
'    Set tb = New Recordset
'    RecOpenServer 0, tb, SQL
'    Do While Not tb.EOF
'        s = Trim(tb!Text)
'       frmMessages.lstComm.AddItem s
'          tb.MoveNext
'    Loop
'
'    Set frmMessages.F = Me
'    Set frmMessages.T = txtImmComment(0)
'    frmMessages.Show 1
'
'  End If
  
  cmdSaveImm(0).Enabled = True
Else

  If KeyCode = 113 Then
  
  If Len(txtImmComment(1)) < 2 Then Exit Sub
  
  n = txtImmComment(1).SelStart
  
  s = UCase(Mid(txtImmComment(1), (n - 2), 3))
  
  If ListText("IM", s) <> "" Then
    s = ListText("IM", s)
  End If
  
  txtImmComment(1) = Left(txtImmComment(1), (n) - 3)
  txtImmComment(1) = txtImmComment(1) & s
  
  txtImmComment(1).SelStart = Len(txtImmComment(1))
  
  ElseIf KeyCode = 114 Then
    
    SQL = "SELECT * from lists WHERE listtype = 'IM'"
    Set tb = New Recordset
    RecOpenServer 0, tb, SQL
    Do While Not tb.EOF
        s = Trim(tb!Text)
       frmMessages.lstComm.AddItem s
          tb.MoveNext
    Loop
    
    Set frmMessages.F = Me
    Set frmMessages.T = txtImmComment(1)
    frmMessages.Show 1

  End If

  cmdSaveImm(1).Enabled = True
End If


Exit Sub

txtImmComment_KeyDown_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /txtImmComment_KeyDown ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub txtInput_Change()

txtInput.SelStart = Len(txtInput)

gRbc.TextMatrix(gRbc.RowSel, 1) = Trim(txtInput)

If gRbc.TextMatrix(gRbc.RowSel, 1) = "" Then
  gRbc.TextMatrix(gRbc.RowSel, 2) = ""
End If

cmdSaveHaem.Enabled = True: cmdHSaveH.Enabled = True

End Sub

Private Sub txtName_Change()

lName = txtName

End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)

If txtName.Locked Then Exit Sub

cmdSaveDemographics.Enabled = True
cmdSaveInc.Enabled = True

End Sub

Private Sub txtname_LostFocus()

Dim strName As String
Dim strSex As String

strName = txtName
strSex = txtSex

NameLostFocus strName, strSex

txtName = strName
txtSex = strSex

End Sub

Private Sub txtNoPas_LostFocus()

If Trim(txtName) = "" Then
  LoadDemo txtNOPAS
End If

End Sub

Public Sub txtSampleID_LostFocus()

On Error GoTo txtSampleID_LostFocus_Error


'If Not Me.ActiveControl = "" Or Me.ActiveControl Is Nothing Then
'    Exit Sub
'End If
'if <1 or >2147483647 or = ""


If Val(txtSampleID) < 1 Or Trim$(txtSampleID) = "" Or Val(txtSampleID) > (2 ^ 31) - 1 Then
  txtSampleID = ""
  txtSampleID.SetFocus
  Exit Sub
End If

txtSampleID = Val(txtSampleID)


LoadAllDetails

cmdSaveDemographics.Enabled = False
cmdSaveInc.Enabled = False
cmdSaveHaem.Enabled = False
cmdHSaveH.Enabled = False
cmdSaveBio.Enabled = False
cmdSaveCoag.Enabled = False
cmdSaveImm(0).Enabled = False
cmdSaveImm(1).Enabled = False
cmdSaveBGa.Enabled = False

Exit Sub

txtSampleID_LostFocus_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /txtSampleID_LostFocus ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub txtSex_Change()

lSex = txtSex

End Sub

Private Sub txtsex_Click()

On Error GoTo txtsex_Click_Error

If txtSex.Locked Then Exit Sub

Select Case Trim$(txtSex)
  Case "": txtSex = "Male"
  Case "Male": txtSex = "Female"
  Case "Female": txtSex = ""
  Case Else: txtSex = ""
End Select

cmdSaveDemographics.Enabled = True
cmdSaveInc.Enabled = True

Exit Sub

txtsex_Click_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /txtsex_Click ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub txtsex_KeyPress(KeyAscii As Integer)

KeyAscii = 0
txtsex_Click

End Sub

Private Sub txtSex_LostFocus()

If txtSex.Locked = True Then Exit Sub

SexLostFocus txtSex, txtName

If sstabAll.Tab = 0 Then
  taddress(0).SetFocus
End If

End Sub

Private Sub UPDATEMRU()

Dim SQL As String
Dim tb As New Recordset
Dim n As Long
Dim Found As Boolean
Dim NewMRU(0 To 9, 0 To 1) As String
'(x,0) SampleID
'(x,1) DateTime

On Error GoTo UPDATEMRU_Error

SQL = "SELECT top 10 * from MRU WHERE " & _
      "UserCode = '" & UserCode & "' " & _
      "Order by DateTime desc"
Set tb = New Recordset
RecOpenServer 0, tb, SQL
      
n = -1
Do While Not tb.EOF
  n = n + 1
  NewMRU(n, 0) = Trim$(tb!SampleID)
  NewMRU(n, 1) = tb!DateTime
  tb.MoveNext
Loop

Found = False
For n = 0 To 9
  If txtSampleID = NewMRU(n, 0) Then
    SQL = "UPDATE MRU " & _
          "Set DateTime = '" & Format$(Now, "dd/mmm/yyyy hh:mm:ss") & "' " & _
          "WHERE SampleID = '" & txtSampleID & "' " & _
          "and UserCode = '" & UserCode & "'"
    Cnxn(0).Execute SQL
    Found = True
    Exit For
  End If
Next
    
If Not Found Then
  SQL = "DELETE from MRU WHERE " & _
        "UserCode = '" & UserCode & "'"
  Cnxn(0).Execute SQL
  For n = 0 To 8
    If NewMRU(n, 0) <> "" Then
      SQL = "INSERT into MRU " & _
            "(SampleID, DateTime, UserCode ) VALUES " & _
            "('" & NewMRU(n, 0) & "', " & _
            "'" & Format$(NewMRU(n, 1), "dd/mmm/yyyy hh:mm:ss") & "', " & _
            "'" & UserCode & "')"
      Cnxn(0).Execute SQL
    End If
  Next
  SQL = "INSERT into MRU " & _
        "(SampleID, DateTime, UserCode ) VALUES " & _
        "('" & txtSampleID & "', " & _
        "'" & Format$(Now, "dd/mmm/yyyy hh:mm:ss") & "', " & _
        "'" & UserCode & "')"
  Cnxn(0).Execute SQL
End If

FillMRU

Exit Sub

UPDATEMRU_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /UPDATEMRU ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub UpDown1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error GoTo UpDown1_MouseUp_Error

pBar = 0


UpDown1.Enabled = False

If SysOptNumLen(0) > 0 Then
  If Len(txtSampleID) > SysOptNumLen(0) Then
    iMsg "Sample Id longer then recommended!"
  End If
End If

UpDown1.Enabled = True

LoadAllDetails

cmdSaveDemographics.Enabled = False
cmdSaveInc.Enabled = False
cmdSaveHaem.Enabled = False
cmdHSaveH.Enabled = False
cmdSaveBio.Enabled = False
cmdSaveCoag.Enabled = False
cmdSaveImm(1).Enabled = False
cmdSaveImm(0).Enabled = False
cmdSaveBGa.Enabled = False
cmdSaveExt.Enabled = False

Exit Sub

UpDown1_MouseUp_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " frmEditAll /UpDown1_MouseUp ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Private Sub VScroll1_Change()

pdelta.Top = -VScroll1

End Sub
