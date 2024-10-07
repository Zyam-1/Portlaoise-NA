VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "comct232.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.ocx"
Begin VB.Form frmEditAll 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - General Chemistry"
   ClientHeight    =   9780
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   14925
   Icon            =   "frmEditAll.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9780
   ScaleWidth      =   14925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin ComctlLib.ProgressBar pBar 
      Height          =   135
      Left            =   0
      TabIndex        =   335
      Top             =   0
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   238
      _Version        =   327682
      Appearance      =   1
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   334
      Top             =   9450
      Width           =   14925
      _ExtentX        =   26326
      _ExtentY        =   582
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   5
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdResend 
      Caption         =   "Resend Results"
      Height          =   360
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   330
      Tag             =   "bOrder"
      ToolTipText     =   "Order Tests for Sample"
      Top             =   2520
      Width           =   1590
   End
   Begin VB.TextBox txtText 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   15000
      TabIndex        =   166
      Top             =   600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdAudit 
      Caption         =   "Audit Trail"
      Height          =   360
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   323
      Tag             =   "bOrder"
      ToolTipText     =   "Order Tests for Sample"
      Top             =   2520
      Width           =   2190
   End
   Begin VB.ComboBox cmbEndResults 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   14160
      TabIndex        =   321
      Top             =   60
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.CommandButton cmdPhone 
      Caption         =   "Phone Results"
      Height          =   750
      Left            =   13500
      Picture         =   "frmEditAll.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   271
      ToolTipText     =   "View Phone Log"
      Top             =   7020
      Width           =   1275
   End
   Begin VB.Frame Frame6 
      Height          =   1800
      Left            =   135
      TabIndex        =   205
      Top             =   270
      Width           =   3195
      Begin VB.CommandButton cmdPatientNotePad 
         Height          =   530
         Index           =   1
         Left            =   2520
         Picture         =   "frmEditAll.frx":074C
         Style           =   1  'Graphical
         TabIndex        =   333
         Top             =   1080
         Visible         =   0   'False
         Width           =   530
      End
      Begin VB.CommandButton cmdPatientNotePad 
         Height          =   530
         Index           =   0
         Left            =   2550
         Picture         =   "frmEditAll.frx":0B2C
         Style           =   1  'Graphical
         TabIndex        =   331
         Tag             =   "bprint"
         Top             =   495
         Width           =   530
      End
      Begin VB.ComboBox cMRU 
         Height          =   315
         Left            =   180
         TabIndex        =   206
         Text            =   "cMRU"
         ToolTipText     =   "Most Recently Used Numbers"
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
         Left            =   105
         MaxLength       =   8
         TabIndex        =   0
         Tag             =   "Sample Id "
         ToolTipText     =   "Sample Identification Number"
         Top             =   510
         Width           =   1785
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   480
         Left            =   1936
         TabIndex        =   207
         Top             =   510
         Width           =   240
         _ExtentX        =   450
         _ExtentY        =   847
         _Version        =   327681
         Value           =   1
         BuddyControl    =   "txtSampleID"
         BuddyDispid     =   196617
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
      Begin VB.Label lblResultOrRequest 
         Alignment       =   2  'Center
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Results"
         Height          =   285
         Left            =   600
         TabIndex        =   210
         ToolTipText     =   "Click to Toggle"
         Top             =   210
         Width           =   885
      End
      Begin VB.Image imgLast 
         Height          =   300
         Left            =   2070
         Picture         =   "frmEditAll.frx":13F6
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
         TabIndex        =   209
         Top             =   1035
         Width           =   375
      End
      Begin VB.Image iRelevant 
         Height          =   480
         Index           =   1
         Left            =   1485
         Picture         =   "frmEditAll.frx":1838
         ToolTipText     =   "Find Next Relevant Sample "
         Top             =   135
         Width           =   480
      End
      Begin VB.Image iRelevant 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   0
         Left            =   120
         Picture         =   "frmEditAll.frx":1B42
         ToolTipText     =   "Find Previous Relevant Sample"
         Top             =   150
         Width           =   480
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         Caption         =   "Sample ID"
         Height          =   195
         Left            =   720
         TabIndex        =   208
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Patient Details"
      Height          =   1800
      Left            =   3330
      TabIndex        =   181
      Top             =   270
      Width           =   11430
      Begin VB.CommandButton cmdDartViewer 
         Height          =   390
         Left            =   4635
         Picture         =   "frmEditAll.frx":1E4C
         Style           =   1  'Graphical
         TabIndex        =   287
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton bDoB 
         Appearance      =   0  'Flat
         Caption         =   "S&earch"
         Height          =   285
         Left            =   9030
         TabIndex        =   184
         ToolTipText     =   "Search using Date of Birth"
         Top             =   285
         Width           =   705
      End
      Begin VB.CommandButton bsearch 
         Appearance      =   0  'Flat
         Caption         =   "Se&arch"
         Height          =   345
         Left            =   6120
         TabIndex        =   183
         ToolTipText     =   "Search using Name"
         Top             =   180
         Width           =   675
      End
      Begin VB.TextBox txtSex 
         Height          =   285
         Left            =   7410
         MaxLength       =   6
         TabIndex        =   5
         Tag             =   "Sex"
         Top             =   1035
         Width           =   1545
      End
      Begin VB.TextBox txtAge 
         Height          =   285
         Left            =   7410
         MaxLength       =   4
         TabIndex        =   182
         Tag             =   "Age"
         Top             =   675
         Width           =   1545
      End
      Begin VB.TextBox txtDoB 
         Height          =   285
         Left            =   7410
         MaxLength       =   10
         TabIndex        =   4
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
         Left            =   2595
         MaxLength       =   30
         TabIndex        =   3
         Tag             =   "Patient Name"
         ToolTipText     =   "Patients Name"
         Top             =   600
         Width           =   4335
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
         Top             =   600
         Width           =   1225
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
         Left            =   1350
         TabIndex        =   2
         Tag             =   "A and E Number"
         ToolTipText     =   "A & E Number"
         Top             =   600
         Width           =   1235
      End
      Begin VB.Label lblUrgent 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "URGENT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9180
         TabIndex        =   215
         ToolTipText     =   "Results Needed Urgently"
         Top             =   900
         Width           =   2085
      End
      Begin VB.Label lblDemographicComment 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   90
         TabIndex        =   212
         ToolTipText     =   "Demographic Comment"
         Top             =   1350
         Width           =   11190
      End
      Begin VB.Label lblSampledate 
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   9765
         TabIndex        =   199
         ToolTipText     =   "Sample Date"
         Top             =   450
         Width           =   1575
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         Caption         =   "Sex"
         Height          =   195
         Left            =   7035
         TabIndex        =   192
         Top             =   1065
         Width           =   270
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "Age"
         Height          =   195
         Left            =   7005
         TabIndex        =   191
         Top             =   705
         Width           =   285
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         Caption         =   "D.o.B"
         Height          =   195
         Left            =   6915
         TabIndex        =   190
         Top             =   345
         Width           =   405
      End
      Begin VB.Label lNoPrevious 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "No Previous Details"
         ForeColor       =   &H0000FFFF&
         Height          =   450
         Left            =   5085
         TabIndex        =   189
         Top             =   120
         Visible         =   0   'False
         Width           =   960
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblAddWardGP 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   90
         TabIndex        =   188
         ToolTipText     =   "Patient Location Information"
         Top             =   1095
         Width           =   6825
      End
      Begin VB.Label lblChartNumber 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Monaghan Chart #"
         Height          =   285
         Left            =   90
         TabIndex        =   187
         ToolTipText     =   "Click to change Location"
         Top             =   315
         Width           =   1425
      End
      Begin VB.Label lblAandE 
         Caption         =   "A and E          Name"
         Height          =   225
         Left            =   1755
         TabIndex        =   186
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label lblRundate 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3420
         TabIndex        =   185
         Top             =   630
         Width           =   1515
      End
      Begin VB.Label Label6 
         Caption         =   "Sample Date"
         Height          =   255
         Left            =   9810
         TabIndex        =   200
         Top             =   270
         Width           =   1035
      End
   End
   Begin VB.CommandButton bOrderTests 
      Caption         =   "Order Tests"
      Height          =   780
      Left            =   13500
      Picture         =   "frmEditAll.frx":2716
      Style           =   1  'Graphical
      TabIndex        =   163
      Tag             =   "bOrder"
      ToolTipText     =   "Order Tests for Sample"
      Top             =   3015
      Width           =   1290
   End
   Begin VB.CommandButton cmdSetPrinter 
      Height          =   705
      Left            =   13500
      Picture         =   "frmEditAll.frx":2A20
      Style           =   1  'Graphical
      TabIndex        =   127
      ToolTipText     =   "Back Colour Normal:Automatic Printer Selection.Back Colour Red:-Forced"
      Top             =   3870
      Width           =   1275
   End
   Begin VB.Timer TimerBar 
      Interval        =   1000
      Left            =   13770
      Top             =   540
   End
   Begin VB.CommandButton cmdPrintHold 
      Caption         =   "Print && Hold"
      Height          =   705
      Left            =   13500
      Picture         =   "frmEditAll.frx":2D2A
      Style           =   1  'Graphical
      TabIndex        =   103
      ToolTipText     =   "Print Result && Stay at Sample"
      Top             =   4635
      Width           =   1275
   End
   Begin VB.CommandButton bViewBB 
      Caption         =   "Transfusion Details"
      Enabled         =   0   'False
      Height          =   840
      Left            =   13485
      Picture         =   "frmEditAll.frx":3034
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
      Picture         =   "frmEditAll.frx":38FE
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "View Patient History"
      Top             =   7830
      Width           =   1275
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   705
      Left            =   13500
      Picture         =   "frmEditAll.frx":3D40
      Style           =   1  'Graphical
      TabIndex        =   25
      Tag             =   "bprint"
      ToolTipText     =   "Print Result"
      Top             =   5400
      Width           =   1275
   End
   Begin VB.CommandButton bFAX 
      Caption         =   "&Fax"
      Height          =   795
      Left            =   13500
      Picture         =   "frmEditAll.frx":404A
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Fax Result"
      Top             =   6165
      Width           =   1275
   End
   Begin VB.CommandButton bCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   645
      Left            =   13500
      Picture         =   "frmEditAll.frx":4354
      Style           =   1  'Graphical
      TabIndex        =   27
      Tag             =   "101"
      ToolTipText     =   "Exit Screen"
      Top             =   8595
      Width           =   1275
   End
   Begin TabDlg.SSTab ssTabAll 
      Height          =   7350
      Left            =   120
      TabIndex        =   23
      Top             =   2115
      Width           =   13275
      _ExtentX        =   23416
      _ExtentY        =   12965
      _Version        =   393216
      Tabs            =   8
      Tab             =   6
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
      TabPicture(0)   =   "frmEditAll.frx":465E
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdDemoVal"
      Tab(0).Control(1)=   "Frame10(0)"
      Tab(0).Control(2)=   "cmdSaveInc"
      Tab(0).Control(3)=   "Frame4"
      Tab(0).Control(4)=   "cmdSaveDemographics"
      Tab(0).Control(5)=   "fraDate"
      Tab(0).Control(6)=   "Frame5"
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Haematology"
      TabPicture(1)   =   "frmEditAll.frx":467A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1(10)"
      Tab(1).Control(1)=   "lblHaemPrinted"
      Tab(1).Control(2)=   "lHaemErrors"
      Tab(1).Control(3)=   "lblHaemValid"
      Tab(1).Control(4)=   "lHDate"
      Tab(1).Control(5)=   "Rundate(1)"
      Tab(1).Control(6)=   "Label1(11)"
      Tab(1).Control(7)=   "lblAnalyser"
      Tab(1).Control(8)=   "Label1(20)"
      Tab(1).Control(9)=   "lblRepeats"
      Tab(1).Control(10)=   "Panel3D8"
      Tab(1).Control(11)=   "bViewHaemRepeat"
      Tab(1).Control(12)=   "cmdSaveHaem"
      Tab(1).Control(13)=   "bValidateHaem"
      Tab(1).Control(14)=   "Panel3D4"
      Tab(1).Control(15)=   "Panel3D5"
      Tab(1).Control(16)=   "Panel3D7"
      Tab(1).Control(17)=   "txtHaemComment"
      Tab(1).Control(18)=   "bHaemGraphs"
      Tab(1).Control(19)=   "cFilm"
      Tab(1).Control(20)=   "bFilm"
      Tab(1).Control(21)=   "txtCondition"
      Tab(1).Control(22)=   "cmdHSaveH"
      Tab(1).Control(23)=   "cmdSaveComm"
      Tab(1).Control(24)=   "cmdViewHaemRep"
      Tab(1).Control(25)=   "cmdUnvalPrint"
      Tab(1).Control(26)=   "Panel3D6"
      Tab(1).Control(27)=   "Picture2"
      Tab(1).ControlCount=   28
      TabCaption(2)   =   "Biochemistry"
      TabPicture(2)   =   "frmEditAll.frx":4696
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblViewSplit"
      Tab(2).Control(1)=   "lRandom"
      Tab(2).Control(2)=   "lBDate"
      Tab(2).Control(3)=   "Rundate(2)"
      Tab(2).Control(4)=   "lblAss"
      Tab(2).Control(5)=   "An1"
      Tab(2).Control(6)=   "An2"
      Tab(2).Control(7)=   "gBio"
      Tab(2).Control(8)=   "Frame2"
      Tab(2).Control(9)=   "cUnits"
      Tab(2).Control(10)=   "tnewvalue"
      Tab(2).Control(11)=   "cAdd"
      Tab(2).Control(12)=   "bremoveduplicates"
      Tab(2).Control(13)=   "bAddBio"
      Tab(2).Control(14)=   "Frame3"
      Tab(2).Control(15)=   "cmdSaveBio"
      Tab(2).Control(16)=   "bValidateBio"
      Tab(2).Control(17)=   "grdOutstanding"
      Tab(2).Control(18)=   "Frame8"
      Tab(2).Control(19)=   "bViewBioRepeat"
      Tab(2).Control(20)=   "bReprint"
      Tab(2).Control(21)=   "cmdViewBioReps"
      Tab(2).Control(22)=   "fraSelectPrint(1)"
      Tab(2).Control(23)=   "cISampleType(3)"
      Tab(2).ControlCount=   24
      TabCaption(3)   =   "Coagulation"
      TabPicture(3)   =   "frmEditAll.frx":46B2
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame11(2)"
      Tab(3).Control(1)=   "fraSelectPrint(2)"
      Tab(3).Control(2)=   "txtCoagComment"
      Tab(3).Control(3)=   "cmdViewCoagRep"
      Tab(3).Control(4)=   "cCunits"
      Tab(3).Control(5)=   "cmdPrintAll"
      Tab(3).Control(6)=   "Frame9"
      Tab(3).Control(7)=   "cmdValidateCoag"
      Tab(3).Control(8)=   "cmdSaveCoag"
      Tab(3).Control(9)=   "bViewCoagRepeat"
      Tab(3).Control(10)=   "bAddCoag"
      Tab(3).Control(11)=   "tResult"
      Tab(3).Control(12)=   "cParameter"
      Tab(3).Control(13)=   "grdCoag"
      Tab(3).Control(14)=   "grdOutstandingCoag"
      Tab(3).Control(15)=   "grdPrev"
      Tab(3).Control(16)=   "Rundate(3)"
      Tab(3).Control(17)=   "lCDate"
      Tab(3).Control(18)=   "lblPrevCoag"
      Tab(3).Control(19)=   "Label20"
      Tab(3).ControlCount=   20
      TabCaption(4)   =   "Endocrinology"
      TabPicture(4)   =   "frmEditAll.frx":46CE
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "lblImmViewSplit(0)"
      Tab(4).Control(1)=   "lImmRan(0)"
      Tab(4).Control(2)=   "Rundate(0)"
      Tab(4).Control(3)=   "lblEDate"
      Tab(4).Control(4)=   "Frame11(0)"
      Tab(4).Control(5)=   "Frame10(1)"
      Tab(4).Control(6)=   "gImm(0)"
      Tab(4).Control(7)=   "grdOutstandings(0)"
      Tab(4).Control(8)=   "bImmRePrint(0)"
      Tab(4).Control(9)=   "bViewImmRepeat(0)"
      Tab(4).Control(10)=   "Frame81(0)"
      Tab(4).Control(11)=   "bValidateImm(0)"
      Tab(4).Control(12)=   "cmdSaveImm(0)"
      Tab(4).Control(13)=   "cmdIAdd(0)"
      Tab(4).Control(14)=   "cmdIremoveduplicates(0)"
      Tab(4).Control(15)=   "cIAdd(0)"
      Tab(4).Control(16)=   "tINewValue(0)"
      Tab(4).Control(17)=   "cIUnits(0)"
      Tab(4).Control(18)=   "cISampleType(0)"
      Tab(4).Control(19)=   "Frame12(0)"
      Tab(4).Control(20)=   "cmdViewReports"
      Tab(4).Control(21)=   "fraSelectPrint(3)"
      Tab(4).Control(22)=   "cmdGetBioEnd"
      Tab(4).ControlCount=   23
      TabCaption(5)   =   "Blood Gas"
      TabPicture(5)   =   "frmEditAll.frx":46EA
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "bImmRePrint(2)"
      Tab(5).Control(1)=   "cISampleType(2)"
      Tab(5).Control(2)=   "cmdIAdd(2)"
      Tab(5).Control(3)=   "cIAdd(2)"
      Tab(5).Control(4)=   "tINewValue(2)"
      Tab(5).Control(5)=   "cIUnits(2)"
      Tab(5).Control(6)=   "Frame15"
      Tab(5).Control(7)=   "bViewBgaRepeat"
      Tab(5).Control(8)=   "cmdValBG"
      Tab(5).Control(9)=   "cmdSaveBGa"
      Tab(5).Control(10)=   "Frame14"
      Tab(5).Control(11)=   "gBga"
      Tab(5).Control(12)=   "lblBgaDate"
      Tab(5).Control(13)=   "Rundate(5)"
      Tab(5).ControlCount=   14
      TabCaption(6)   =   "Immunology"
      TabPicture(6)   =   "frmEditAll.frx":4706
      Tab(6).ControlEnabled=   -1  'True
      Tab(6).Control(0)=   "Rundate(4)"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).Control(1)=   "lblIRundate"
      Tab(6).Control(1).Enabled=   0   'False
      Tab(6).Control(2)=   "lblImmViewSplit(1)"
      Tab(6).Control(2).Enabled=   0   'False
      Tab(6).Control(3)=   "lImmRan(1)"
      Tab(6).Control(3).Enabled=   0   'False
      Tab(6).Control(4)=   "gImm(1)"
      Tab(6).Control(4).Enabled=   0   'False
      Tab(6).Control(5)=   "grdOutstandings(1)"
      Tab(6).Control(5).Enabled=   0   'False
      Tab(6).Control(6)=   "cmdGetBio"
      Tab(6).Control(6).Enabled=   0   'False
      Tab(6).Control(7)=   "cIAdd(1)"
      Tab(6).Control(7).Enabled=   0   'False
      Tab(6).Control(8)=   "tINewValue(1)"
      Tab(6).Control(8).Enabled=   0   'False
      Tab(6).Control(9)=   "cIUnits(1)"
      Tab(6).Control(9).Enabled=   0   'False
      Tab(6).Control(10)=   "cISampleType(1)"
      Tab(6).Control(10).Enabled=   0   'False
      Tab(6).Control(11)=   "bImmRePrint(1)"
      Tab(6).Control(11).Enabled=   0   'False
      Tab(6).Control(12)=   "bViewImmRepeat(1)"
      Tab(6).Control(12).Enabled=   0   'False
      Tab(6).Control(13)=   "bValidateImm(1)"
      Tab(6).Control(13).Enabled=   0   'False
      Tab(6).Control(14)=   "cmdSaveImm(1)"
      Tab(6).Control(14).Enabled=   0   'False
      Tab(6).Control(15)=   "cmdIAdd(1)"
      Tab(6).Control(15).Enabled=   0   'False
      Tab(6).Control(16)=   "cmdIremoveduplicates(1)"
      Tab(6).Control(16).Enabled=   0   'False
      Tab(6).Control(17)=   "Frame81(1)"
      Tab(6).Control(17).Enabled=   0   'False
      Tab(6).Control(18)=   "Frame11(1)"
      Tab(6).Control(18).Enabled=   0   'False
      Tab(6).Control(19)=   "Frame12(1)"
      Tab(6).Control(19).Enabled=   0   'False
      Tab(6).Control(20)=   "cmdViewImmRep"
      Tab(6).Control(20).Enabled=   0   'False
      Tab(6).Control(21)=   "cmdOrderPhoresis"
      Tab(6).Control(21).Enabled=   0   'False
      Tab(6).Control(22)=   "cmdPhoresisComments"
      Tab(6).Control(22).Enabled=   0   'False
      Tab(6).Control(23)=   "fraSelectPrint(0)"
      Tab(6).Control(23).Enabled=   0   'False
      Tab(6).ControlCount=   24
      TabCaption(7)   =   "Externals"
      TabPicture(7)   =   "frmEditAll.frx":4722
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "baddtotests(1)"
      Tab(7).Control(1)=   "cmdSaveImm(2)"
      Tab(7).Control(2)=   "bValidateImm(2)"
      Tab(7).Control(3)=   "cmdExcel"
      Tab(7).Control(4)=   "cmdViewExtReport"
      Tab(7).Control(5)=   "baddtotests(0)"
      Tab(7).Control(6)=   "txtEtc(8)"
      Tab(7).Control(7)=   "txtEtc(7)"
      Tab(7).Control(8)=   "txtEtc(6)"
      Tab(7).Control(9)=   "txtEtc(5)"
      Tab(7).Control(10)=   "txtEtc(1)"
      Tab(7).Control(11)=   "txtEtc(2)"
      Tab(7).Control(12)=   "txtEtc(3)"
      Tab(7).Control(13)=   "txtEtc(4)"
      Tab(7).Control(14)=   "txtEtc(0)"
      Tab(7).Control(15)=   "cmdDel"
      Tab(7).Control(16)=   "grdExt"
      Tab(7).Control(17)=   "lblExcelInfo"
      Tab(7).ControlCount=   18
      Begin VB.CommandButton bImmRePrint 
         Caption         =   "Re-Print"
         Height          =   960
         Index           =   2
         Left            =   -65220
         Picture         =   "frmEditAll.frx":473E
         Style           =   1  'Graphical
         TabIndex        =   332
         ToolTipText     =   "Re Print already Printed Results"
         Top             =   6180
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.CommandButton baddtotests 
         Appearance      =   0  'Flat
         Caption         =   "NVRL        St. James"
         Height          =   1100
         Index           =   1
         Left            =   -63335
         Picture         =   "frmEditAll.frx":4A48
         Style           =   1  'Graphical
         TabIndex        =   329
         ToolTipText     =   "Order External Tests"
         Top             =   4710
         Width           =   1100
      End
      Begin VB.CommandButton cmdSaveImm 
         Caption         =   "Save Details"
         Enabled         =   0   'False
         Height          =   915
         Index           =   2
         Left            =   -63780
         Picture         =   "frmEditAll.frx":5312
         Style           =   1  'Graphical
         TabIndex        =   328
         ToolTipText     =   "Save Changes"
         Top             =   6180
         Width           =   765
      End
      Begin VB.CommandButton cmdGetBioEnd 
         Caption         =   "Get Bio Tests"
         Height          =   915
         Left            =   -68820
         Picture         =   "frmEditAll.frx":561C
         Style           =   1  'Graphical
         TabIndex        =   327
         ToolTipText     =   "Retrieve Biochemistry Tests Relevant to Immunology"
         Top             =   6180
         Width           =   915
      End
      Begin VB.CommandButton bValidateImm 
         Caption         =   "Validate"
         Height          =   915
         Index           =   2
         Left            =   -62940
         Picture         =   "frmEditAll.frx":5926
         Style           =   1  'Graphical
         TabIndex        =   326
         ToolTipText     =   "Result Validation"
         Top             =   6180
         Width           =   705
      End
      Begin VB.Frame Frame11 
         Caption         =   "Delta Check"
         Height          =   1065
         Index           =   2
         Left            =   -74760
         TabIndex        =   324
         Top             =   5160
         Width           =   4080
         Begin VB.Label lIDelta 
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H000000FF&
            Height          =   735
            Index           =   2
            Left            =   135
            TabIndex        =   325
            ToolTipText     =   "Delta Check"
            Top             =   180
            Width           =   3870
            WordWrap        =   -1  'True
         End
      End
      Begin VB.ComboBox cISampleType 
         Height          =   315
         Index           =   3
         Left            =   -70395
         TabIndex        =   322
         Text            =   "cSampleType"
         ToolTipText     =   "Choose Sample Type"
         Top             =   6375
         Width           =   1500
      End
      Begin VB.PictureBox Picture2 
         Height          =   945
         Left            =   -66585
         ScaleHeight     =   885
         ScaleWidth      =   2070
         TabIndex        =   316
         Top             =   3660
         Width           =   2130
         Begin VB.TextBox txtReadingDateTime 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   90
            Locked          =   -1  'True
            TabIndex        =   319
            ToolTipText     =   "DateTime of reading"
            Top             =   540
            Width           =   1935
         End
         Begin VB.TextBox txtViscosity 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   90
            Locked          =   -1  'True
            TabIndex        =   317
            ToolTipText     =   "Viscosity at 37 degree celsius"
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Viscosity at 37 Celsius"
            Height          =   195
            Index           =   21
            Left            =   277
            TabIndex        =   318
            Top             =   0
            Width           =   1560
         End
      End
      Begin VB.Frame fraSelectPrint 
         Height          =   435
         Index           =   3
         Left            =   -69030
         TabIndex        =   312
         Top             =   360
         Width           =   2085
         Begin VB.CommandButton cmdGreenTick 
            Height          =   285
            Index           =   3
            Left            =   1740
            Picture         =   "frmEditAll.frx":5C30
            Style           =   1  'Graphical
            TabIndex        =   314
            Top             =   120
            Width           =   315
         End
         Begin VB.CommandButton cmdRedCross 
            Height          =   285
            Index           =   3
            Left            =   1410
            Picture         =   "frmEditAll.frx":5F06
            Style           =   1  'Graphical
            TabIndex        =   313
            Top             =   120
            Width           =   315
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Select for Printing"
            Height          =   195
            Index           =   3
            Left            =   45
            TabIndex        =   315
            Top             =   180
            Width           =   1275
         End
      End
      Begin VB.Frame fraSelectPrint 
         Height          =   435
         Index           =   2
         Left            =   -68700
         TabIndex        =   308
         Top             =   390
         Width           =   2085
         Begin VB.CommandButton cmdGreenTick 
            Height          =   285
            Index           =   2
            Left            =   1740
            Picture         =   "frmEditAll.frx":61DC
            Style           =   1  'Graphical
            TabIndex        =   310
            Top             =   120
            Width           =   315
         End
         Begin VB.CommandButton cmdRedCross 
            Height          =   285
            Index           =   2
            Left            =   1410
            Picture         =   "frmEditAll.frx":64B2
            Style           =   1  'Graphical
            TabIndex        =   309
            Top             =   120
            Width           =   315
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Select for Printing"
            Height          =   195
            Index           =   2
            Left            =   45
            TabIndex        =   311
            Top             =   180
            Width           =   1275
         End
      End
      Begin VB.Frame fraSelectPrint 
         Height          =   435
         Index           =   1
         Left            =   -68640
         TabIndex        =   304
         Top             =   360
         Width           =   2085
         Begin VB.CommandButton cmdGreenTick 
            Height          =   285
            Index           =   1
            Left            =   1740
            Picture         =   "frmEditAll.frx":6788
            Style           =   1  'Graphical
            TabIndex        =   306
            Top             =   120
            Width           =   315
         End
         Begin VB.CommandButton cmdRedCross 
            Height          =   285
            Index           =   1
            Left            =   1410
            Picture         =   "frmEditAll.frx":6A5E
            Style           =   1  'Graphical
            TabIndex        =   305
            Top             =   120
            Width           =   315
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Select for Printing"
            Height          =   195
            Index           =   1
            Left            =   45
            TabIndex        =   307
            Top             =   180
            Width           =   1275
         End
      End
      Begin VB.Frame fraSelectPrint 
         Height          =   435
         Index           =   0
         Left            =   9420
         TabIndex        =   300
         Top             =   360
         Width           =   2085
         Begin VB.CommandButton cmdRedCross 
            Height          =   285
            Index           =   0
            Left            =   1410
            Picture         =   "frmEditAll.frx":6D34
            Style           =   1  'Graphical
            TabIndex        =   303
            Top             =   120
            Width           =   315
         End
         Begin VB.CommandButton cmdGreenTick 
            Height          =   285
            Index           =   0
            Left            =   1740
            Picture         =   "frmEditAll.frx":700A
            Style           =   1  'Graphical
            TabIndex        =   302
            Top             =   120
            Width           =   315
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Select for Printing"
            Height          =   195
            Index           =   0
            Left            =   45
            TabIndex        =   301
            Top             =   180
            Width           =   1275
         End
      End
      Begin VB.CommandButton cmdPhoresisComments 
         Caption         =   "Phoresis Comments"
         Height          =   1155
         Left            =   10620
         Picture         =   "frmEditAll.frx":72E0
         Style           =   1  'Graphical
         TabIndex        =   299
         Top             =   4350
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.PictureBox Panel3D6 
         Height          =   1425
         Left            =   -66360
         ScaleHeight     =   1365
         ScaleWidth      =   1605
         TabIndex        =   292
         Top             =   720
         Width           =   1665
         Begin VB.TextBox tMPV 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   495
            TabIndex        =   294
            ToolTipText     =   "Mean Platelet Volume"
            Top             =   855
            Width           =   915
         End
         Begin VB.TextBox tPlt 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   495
            MaxLength       =   5
            TabIndex        =   293
            ToolTipText     =   "Platlets"
            Top             =   450
            Width           =   915
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Plt"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   90
            TabIndex        =   298
            Top             =   495
            Width           =   285
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "MPV"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   60
            TabIndex        =   297
            Top             =   930
            Width           =   405
         End
         Begin VB.Label ipflag 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H0000FFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Abnormal"
            ForeColor       =   &H00800080&
            Height          =   225
            Index           =   4
            Left            =   750
            TabIndex        =   296
            Top             =   30
            Visible         =   0   'False
            Width           =   825
         End
         Begin VB.Label ipflag 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H0000FFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Suspect"
            ForeColor       =   &H00800080&
            Height          =   225
            Index           =   5
            Left            =   30
            TabIndex        =   295
            Top             =   30
            Visible         =   0   'False
            Width           =   735
         End
      End
      Begin VB.CommandButton cmdOrderPhoresis 
         Caption         =   "Order Phoresis"
         Height          =   1155
         Left            =   11670
         Picture         =   "frmEditAll.frx":81AA
         Style           =   1  'Graphical
         TabIndex        =   291
         Top             =   4350
         Width           =   975
      End
      Begin VB.CommandButton cmdExcel 
         Caption         =   "Export to Excel"
         Height          =   990
         Left            =   -63165
         Picture         =   "frmEditAll.frx":9074
         Style           =   1  'Graphical
         TabIndex        =   286
         Top             =   2055
         Width           =   960
      End
      Begin VB.CommandButton cmdUnvalPrint 
         Caption         =   "Unvalidated Report"
         Height          =   825
         Left            =   -68205
         Picture         =   "frmEditAll.frx":937E
         Style           =   1  'Graphical
         TabIndex        =   284
         ToolTipText     =   "View Unvalidated Print Outs"
         Top             =   6465
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.CommandButton cmdViewExtReport 
         Caption         =   "Reports"
         Height          =   915
         Left            =   -64572
         Picture         =   "frmEditAll.frx":9688
         Style           =   1  'Graphical
         TabIndex        =   283
         ToolTipText     =   "View Printed && Faxed Reports"
         Top             =   6180
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtCoagComment 
         BackColor       =   &H80000018&
         Height          =   1545
         Left            =   -64965
         MaxLength       =   360
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   105
         ToolTipText     =   "Only 360 Characters"
         Top             =   4080
         Width           =   2865
      End
      Begin VB.CommandButton cmdViewImmRep 
         Caption         =   "Reports"
         Height          =   915
         Left            =   7695
         Picture         =   "frmEditAll.frx":9992
         Style           =   1  'Graphical
         TabIndex        =   280
         ToolTipText     =   "View Printed && Faxed Reports"
         Top             =   6180
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.CommandButton cmdViewHaemRep 
         Caption         =   "Reports"
         Height          =   825
         Left            =   -67125
         Picture         =   "frmEditAll.frx":9C9C
         Style           =   1  'Graphical
         TabIndex        =   279
         ToolTipText     =   "View Printed && Faxed Reports"
         Top             =   6465
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdViewCoagRep 
         Caption         =   "Reports"
         Height          =   915
         Left            =   -67665
         Picture         =   "frmEditAll.frx":9FA6
         Style           =   1  'Graphical
         TabIndex        =   278
         ToolTipText     =   "View Printed && Faxed Reports"
         Top             =   6180
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdViewBioReps 
         Caption         =   "Reports"
         Height          =   915
         Left            =   -68880
         Picture         =   "frmEditAll.frx":A2B0
         Style           =   1  'Graphical
         TabIndex        =   277
         ToolTipText     =   "View Printed && Faxed Reports"
         Top             =   6180
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdViewReports 
         Caption         =   "Reports"
         Height          =   915
         Left            =   -67845
         Picture         =   "frmEditAll.frx":A5BA
         Style           =   1  'Graphical
         TabIndex        =   276
         ToolTipText     =   "View Printed && Faxed Reports"
         Top             =   6180
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.CommandButton baddtotests 
         Appearance      =   0  'Flat
         Caption         =   "Order External Test"
         Height          =   1100
         Index           =   0
         Left            =   -64572
         Picture         =   "frmEditAll.frx":A8C4
         Style           =   1  'Graphical
         TabIndex        =   266
         ToolTipText     =   "Order External Tests"
         Top             =   4710
         Width           =   1100
      End
      Begin VB.TextBox txtEtc 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   8
         Left            =   -74790
         MaxLength       =   110
         TabIndex        =   265
         Top             =   6540
         Width           =   10068
      End
      Begin VB.TextBox txtEtc 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   7
         Left            =   -74790
         MaxLength       =   110
         TabIndex        =   264
         Top             =   6268
         Width           =   10068
      End
      Begin VB.TextBox txtEtc 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   6
         Left            =   -74790
         MaxLength       =   110
         TabIndex        =   263
         Top             =   6009
         Width           =   10068
      End
      Begin VB.TextBox txtEtc 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   5
         Left            =   -74790
         MaxLength       =   110
         TabIndex        =   262
         Top             =   5750
         Width           =   10068
      End
      Begin VB.TextBox txtEtc 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   -74790
         MaxLength       =   110
         TabIndex        =   258
         Top             =   4714
         Width           =   10068
      End
      Begin VB.TextBox txtEtc 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   -74790
         MaxLength       =   110
         TabIndex        =   259
         Top             =   4973
         Width           =   10068
      End
      Begin VB.TextBox txtEtc 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   -74790
         MaxLength       =   110
         TabIndex        =   260
         Top             =   5232
         Width           =   10068
      End
      Begin VB.TextBox txtEtc 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   4
         Left            =   -74790
         MaxLength       =   110
         TabIndex        =   261
         Top             =   5491
         Width           =   10068
      End
      Begin VB.TextBox txtEtc 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   -74790
         MaxLength       =   110
         TabIndex        =   257
         Top             =   4455
         Width           =   10068
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "Delete"
         Height          =   990
         Left            =   -63180
         Picture         =   "frmEditAll.frx":ABCE
         Style           =   1  'Graphical
         TabIndex        =   256
         ToolTipText     =   "Delete Test"
         Top             =   810
         Width           =   960
      End
      Begin VB.Frame Frame12 
         Caption         =   "Specimen Condition"
         Height          =   1035
         Index           =   1
         Left            =   7155
         TabIndex        =   245
         Top             =   4140
         Width           =   3285
         Begin VB.CheckBox Ih 
            Caption         =   "Haemolysed"
            Height          =   225
            Index           =   1
            Left            =   1350
            TabIndex        =   251
            Top             =   450
            Width           =   1245
         End
         Begin VB.CheckBox Iis 
            Caption         =   "Slightly Haemolysed"
            Height          =   225
            Index           =   1
            Left            =   1350
            TabIndex        =   250
            Top             =   210
            Width           =   1755
         End
         Begin VB.CheckBox Il 
            Alignment       =   1  'Right Justify
            Caption         =   "Lipaemic"
            Height          =   225
            Index           =   1
            Left            =   300
            TabIndex        =   249
            Top             =   210
            Width           =   975
         End
         Begin VB.CheckBox Io 
            Alignment       =   1  'Right Justify
            Caption         =   "Old Sample"
            Height          =   225
            Index           =   1
            Left            =   120
            TabIndex        =   248
            Top             =   450
            Width           =   1155
         End
         Begin VB.CheckBox Ig 
            Caption         =   "Grossly Haemolysed"
            Height          =   225
            Index           =   1
            Left            =   1350
            TabIndex        =   247
            Top             =   690
            Width           =   1755
         End
         Begin VB.CheckBox Ij 
            Alignment       =   1  'Right Justify
            Caption         =   "Icteric"
            Height          =   225
            Index           =   1
            Left            =   510
            TabIndex        =   246
            Top             =   690
            Width           =   765
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Delta Check"
         Height          =   1905
         Index           =   1
         Left            =   3825
         TabIndex        =   241
         Top             =   4140
         Width           =   3240
         Begin VB.Label lIDelta 
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H000000FF&
            Height          =   1515
            Index           =   1
            Left            =   135
            TabIndex        =   242
            ToolTipText     =   "Delta Check"
            Top             =   225
            Width           =   3030
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame Frame81 
         Caption         =   "Immunology Comments"
         Height          =   1905
         Index           =   1
         Left            =   315
         TabIndex        =   239
         Top             =   4095
         Width           =   3330
         Begin VB.TextBox txtImmComment 
            BackColor       =   &H80000018&
            Height          =   1545
            Index           =   1
            Left            =   90
            MaxLength       =   480
            MultiLine       =   -1  'True
            TabIndex        =   240
            ToolTipText     =   "Immunology Comment"
            Top             =   300
            Width           =   3135
         End
      End
      Begin VB.CommandButton cmdIremoveduplicates 
         Caption         =   "Remove Duplicates"
         Height          =   915
         Index           =   1
         Left            =   9300
         Picture         =   "frmEditAll.frx":AED8
         Style           =   1  'Graphical
         TabIndex        =   238
         ToolTipText     =   "Remove Result"
         Top             =   6180
         Width           =   885
      End
      Begin VB.CommandButton cmdIAdd 
         Caption         =   "Add Result"
         Height          =   915
         Index           =   1
         Left            =   8505
         Picture         =   "frmEditAll.frx":B1E2
         Style           =   1  'Graphical
         TabIndex        =   237
         Tag             =   "bAdd"
         ToolTipText     =   "Add Result Manually"
         Top             =   6180
         Width           =   765
      End
      Begin VB.CommandButton cmdSaveImm 
         Caption         =   "Save Details"
         Enabled         =   0   'False
         Height          =   915
         Index           =   1
         Left            =   11745
         Picture         =   "frmEditAll.frx":B4EC
         Style           =   1  'Graphical
         TabIndex        =   236
         ToolTipText     =   "Save Changes"
         Top             =   6180
         Width           =   765
      End
      Begin VB.CommandButton bValidateImm 
         Caption         =   "Validate"
         Height          =   915
         Index           =   1
         Left            =   12510
         Picture         =   "frmEditAll.frx":B7F6
         Style           =   1  'Graphical
         TabIndex        =   235
         ToolTipText     =   "Result Validation"
         Top             =   6180
         Width           =   705
      End
      Begin VB.CommandButton bViewImmRepeat 
         BackColor       =   &H0000FFFF&
         Caption         =   "View Repeat"
         Height          =   915
         Index           =   1
         Left            =   10215
         Picture         =   "frmEditAll.frx":BB00
         Style           =   1  'Graphical
         TabIndex        =   234
         ToolTipText     =   "View Repeated Tests"
         Top             =   6180
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.CommandButton bImmRePrint 
         Caption         =   "Re-Print"
         Height          =   915
         Index           =   1
         Left            =   10995
         Picture         =   "frmEditAll.frx":BC8A
         Style           =   1  'Graphical
         TabIndex        =   233
         ToolTipText     =   "Re Print already Printed Results"
         Top             =   6180
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.ComboBox cISampleType 
         Height          =   315
         Index           =   1
         Left            =   4590
         TabIndex        =   232
         Text            =   "cSampleType"
         ToolTipText     =   "Choose Sample Type"
         Top             =   6525
         Width           =   1560
      End
      Begin VB.ComboBox cIUnits 
         Height          =   315
         Index           =   1
         Left            =   3225
         TabIndex        =   231
         Text            =   "cUnits"
         ToolTipText     =   "Choose Units"
         Top             =   6525
         Width           =   1305
      End
      Begin VB.TextBox tINewValue 
         Height          =   315
         Index           =   1
         Left            =   1815
         MaxLength       =   300
         TabIndex        =   230
         ToolTipText     =   "Enter Result"
         Top             =   6525
         Width           =   1350
      End
      Begin VB.ComboBox cIAdd 
         Height          =   315
         Index           =   1
         Left            =   180
         TabIndex        =   229
         Text            =   "cAdd"
         ToolTipText     =   "Choose Test"
         Top             =   6525
         Width           =   1575
      End
      Begin VB.CommandButton cmdGetBio 
         Caption         =   "Get Bio Tests"
         Height          =   780
         Left            =   9150
         Picture         =   "frmEditAll.frx":BF94
         Style           =   1  'Graphical
         TabIndex        =   228
         ToolTipText     =   "Retrieve Biochemistry Tests Relevant to Immunology"
         Top             =   5280
         Width           =   1275
      End
      Begin VB.ComboBox cISampleType 
         Height          =   315
         Index           =   2
         Left            =   -70095
         TabIndex        =   227
         Text            =   "cSampleType"
         Top             =   5250
         Width           =   1440
      End
      Begin VB.CommandButton cmdIAdd 
         Caption         =   "Add Result"
         Height          =   960
         Index           =   2
         Left            =   -67800
         Picture         =   "frmEditAll.frx":C29E
         Style           =   1  'Graphical
         TabIndex        =   226
         Tag             =   "bAdd"
         Top             =   6180
         Width           =   765
      End
      Begin VB.ComboBox cIAdd 
         Height          =   315
         Index           =   2
         Left            =   -74640
         TabIndex        =   225
         Text            =   "cAdd"
         Top             =   5250
         Width           =   1575
      End
      Begin VB.TextBox tINewValue 
         Height          =   315
         Index           =   2
         Left            =   -73005
         MaxLength       =   15
         TabIndex        =   224
         Top             =   5250
         Width           =   1485
      End
      Begin VB.ComboBox cIUnits 
         Height          =   315
         Index           =   2
         Left            =   -71460
         TabIndex        =   223
         Text            =   "cUnits"
         Top             =   5250
         Width           =   1305
      End
      Begin VB.CommandButton cmdSaveComm 
         Caption         =   "Save Comment"
         Height          =   915
         Left            =   -65100
         Picture         =   "frmEditAll.frx":C5A8
         Style           =   1  'Graphical
         TabIndex        =   222
         ToolTipText     =   "Save Comment Changes"
         Top             =   4800
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Frame Frame15 
         Caption         =   "Delta Check"
         Height          =   1905
         Left            =   -68430
         TabIndex        =   220
         Top             =   930
         Width           =   4785
         Begin VB.Label lBgaDelta 
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H000000FF&
            Height          =   1515
            Left            =   120
            TabIndex        =   221
            ToolTipText     =   "Delta Check"
            Top             =   270
            Width           =   4560
            WordWrap        =   -1  'True
         End
      End
      Begin VB.CommandButton bViewBgaRepeat 
         BackColor       =   &H0000FFFF&
         Caption         =   "View Repeat"
         Height          =   960
         Left            =   -64380
         Picture         =   "frmEditAll.frx":C8B2
         Style           =   1  'Graphical
         TabIndex        =   219
         Top             =   6180
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.CommandButton cmdHSaveH 
         Caption         =   "Save && Hold"
         Enabled         =   0   'False
         Height          =   825
         Left            =   -63615
         Picture         =   "frmEditAll.frx":CA3C
         Style           =   1  'Graphical
         TabIndex        =   217
         ToolTipText     =   "Save Changes && Stay at Sample"
         Top             =   6465
         Width           =   915
      End
      Begin VB.CommandButton cmdDemoVal 
         Caption         =   "&Validate"
         Height          =   735
         Left            =   -68745
         Picture         =   "frmEditAll.frx":CD46
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   5970
         Width           =   945
      End
      Begin VB.TextBox txtCondition 
         Height          =   975
         Left            =   -66585
         MaxLength       =   500
         MultiLine       =   -1  'True
         TabIndex        =   179
         ToolTipText     =   "Patient Medical Condition"
         Top             =   2550
         Width           =   2130
      End
      Begin VB.CommandButton bFilm 
         Caption         =   "Film"
         Height          =   375
         Left            =   -74865
         Style           =   1  'Graphical
         TabIndex        =   171
         Top             =   4650
         Width           =   1155
      End
      Begin VB.CommandButton cmdValBG 
         Caption         =   "Validate"
         Height          =   960
         Left            =   -66090
         Picture         =   "frmEditAll.frx":D050
         Style           =   1  'Graphical
         TabIndex        =   158
         Top             =   6180
         Width           =   765
      End
      Begin VB.CommandButton cmdSaveBGa 
         Caption         =   "Save Details"
         Enabled         =   0   'False
         Height          =   960
         Left            =   -66930
         Picture         =   "frmEditAll.frx":D35A
         Style           =   1  'Graphical
         TabIndex        =   157
         Top             =   6180
         Width           =   765
      End
      Begin VB.Frame Frame14 
         Caption         =   "Blood Gas Comments"
         Height          =   1905
         Left            =   -68475
         TabIndex        =   155
         Top             =   2865
         Width           =   4845
         Begin VB.TextBox txtBGaComment 
            BackColor       =   &H80000018&
            Height          =   1545
            Left            =   135
            MaxLength       =   320
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   156
            Top             =   225
            Width           =   4605
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Specimen Condition"
         Height          =   1035
         Index           =   0
         Left            =   -65190
         TabIndex        =   147
         Top             =   4665
         Width           =   3285
         Begin VB.CheckBox Ij 
            Alignment       =   1  'Right Justify
            Caption         =   "Icteric"
            Height          =   225
            Index           =   0
            Left            =   510
            TabIndex        =   153
            Top             =   690
            Width           =   765
         End
         Begin VB.CheckBox Ig 
            Caption         =   "Grossly Haemolysed"
            Height          =   225
            Index           =   0
            Left            =   1350
            TabIndex        =   152
            Top             =   690
            Width           =   1755
         End
         Begin VB.CheckBox Io 
            Alignment       =   1  'Right Justify
            Caption         =   "Old Sample"
            Height          =   225
            Index           =   0
            Left            =   120
            TabIndex        =   151
            Top             =   450
            Width           =   1155
         End
         Begin VB.CheckBox Il 
            Alignment       =   1  'Right Justify
            Caption         =   "Lipaemic"
            Height          =   225
            Index           =   0
            Left            =   300
            TabIndex        =   150
            Top             =   210
            Width           =   975
         End
         Begin VB.CheckBox Iis 
            Caption         =   "Slightly Haemolysed"
            Height          =   225
            Index           =   0
            Left            =   1350
            TabIndex        =   149
            Top             =   210
            Width           =   1755
         End
         Begin VB.CheckBox Ih 
            Caption         =   "Haemolysed"
            Height          =   225
            Index           =   0
            Left            =   1350
            TabIndex        =   148
            Top             =   450
            Width           =   1245
         End
      End
      Begin VB.ComboBox cISampleType 
         Height          =   315
         Index           =   0
         Left            =   -70335
         TabIndex        =   146
         Text            =   "cSampleType"
         ToolTipText     =   "Choose Sample Type"
         Top             =   6450
         Width           =   1380
      End
      Begin VB.ComboBox cIUnits 
         Height          =   315
         Index           =   0
         Left            =   -71700
         TabIndex        =   145
         Text            =   "cUnits"
         ToolTipText     =   "Choose Units"
         Top             =   6450
         Width           =   1305
      End
      Begin VB.TextBox tINewValue 
         Height          =   315
         Index           =   0
         Left            =   -73005
         MaxLength       =   15
         TabIndex        =   144
         ToolTipText     =   "Enter Result"
         Top             =   6450
         Width           =   1245
      End
      Begin VB.ComboBox cIAdd 
         Height          =   315
         Index           =   0
         Left            =   -74880
         TabIndex        =   143
         Text            =   "cAdd"
         ToolTipText     =   "Choose Test"
         Top             =   6450
         Width           =   1815
      End
      Begin VB.CommandButton cmdIremoveduplicates 
         Caption         =   "Remove Duplicates"
         Height          =   915
         Index           =   0
         Left            =   -66045
         Picture         =   "frmEditAll.frx":D664
         Style           =   1  'Graphical
         TabIndex        =   142
         ToolTipText     =   "Remove Result"
         Top             =   6180
         Width           =   885
      End
      Begin VB.CommandButton cmdIAdd 
         Caption         =   "Add Result"
         Height          =   915
         Index           =   0
         Left            =   -66855
         Picture         =   "frmEditAll.frx":D96E
         Style           =   1  'Graphical
         TabIndex        =   141
         Tag             =   "bAdd"
         ToolTipText     =   "Add Result Manually"
         Top             =   6180
         Width           =   765
      End
      Begin VB.CommandButton cmdSaveImm 
         Caption         =   "Save Details"
         Enabled         =   0   'False
         Height          =   915
         Index           =   0
         Left            =   -63480
         Picture         =   "frmEditAll.frx":DC78
         Style           =   1  'Graphical
         TabIndex        =   138
         ToolTipText     =   "Save Changes"
         Top             =   6180
         Width           =   765
      End
      Begin VB.CommandButton bValidateImm 
         Caption         =   "Validate"
         Height          =   915
         Index           =   0
         Left            =   -62670
         Picture         =   "frmEditAll.frx":DF82
         Style           =   1  'Graphical
         TabIndex        =   137
         ToolTipText     =   "Result Validation"
         Top             =   6180
         Width           =   765
      End
      Begin VB.Frame Frame81 
         Caption         =   "Endocrinology Comments"
         Height          =   1905
         Index           =   0
         Left            =   -65565
         TabIndex        =   134
         Top             =   2730
         Width           =   3660
         Begin VB.TextBox txtImmComment 
            BackColor       =   &H80000018&
            Height          =   1545
            Index           =   0
            Left            =   90
            MaxLength       =   2000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   135
            Top             =   270
            Width           =   3450
         End
      End
      Begin VB.CommandButton bViewImmRepeat 
         BackColor       =   &H0000FFFF&
         Caption         =   "View Repeat"
         Height          =   915
         Index           =   0
         Left            =   -65100
         Picture         =   "frmEditAll.frx":E28C
         Style           =   1  'Graphical
         TabIndex        =   133
         ToolTipText     =   "View Repeated Tests"
         Top             =   6180
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.CommandButton bImmRePrint 
         Caption         =   "Re-Print"
         Height          =   915
         Index           =   0
         Left            =   -64245
         Picture         =   "frmEditAll.frx":E416
         Style           =   1  'Graphical
         TabIndex        =   132
         ToolTipText     =   "Re Print already Printed Results"
         Top             =   6180
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.ComboBox cCunits 
         Height          =   315
         Left            =   -71670
         TabIndex        =   131
         Text            =   "cCunits"
         ToolTipText     =   "Choose Units"
         Top             =   6330
         Width           =   1005
      End
      Begin VB.Frame Frame10 
         Caption         =   "Category"
         Height          =   825
         Index           =   0
         Left            =   -69000
         TabIndex        =   130
         Top             =   3420
         Width           =   2385
         Begin VB.ComboBox cCat 
            Height          =   315
            Index           =   0
            Left            =   90
            TabIndex        =   19
            Top             =   270
            Width           =   2145
         End
      End
      Begin VB.CommandButton cmdPrintAll 
         Caption         =   "Print All"
         Height          =   915
         Left            =   -65130
         Picture         =   "frmEditAll.frx":E720
         Style           =   1  'Graphical
         TabIndex        =   122
         ToolTipText     =   "Print Result"
         Top             =   6180
         Width           =   765
      End
      Begin VB.CommandButton bReprint 
         Caption         =   "Re-Print"
         Height          =   915
         Left            =   -65460
         Picture         =   "frmEditAll.frx":EA2A
         Style           =   1  'Graphical
         TabIndex        =   121
         ToolTipText     =   "Re Print already Printed Results"
         Top             =   6180
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.CheckBox cFilm 
         Caption         =   "Film"
         Height          =   195
         Left            =   -73650
         TabIndex        =   120
         Top             =   4710
         Width           =   645
      End
      Begin VB.CommandButton bHaemGraphs 
         Caption         =   "Graph"
         Height          =   825
         Left            =   -66360
         Picture         =   "frmEditAll.frx":ED34
         Style           =   1  'Graphical
         TabIndex        =   113
         ToolTipText     =   "view Graph"
         Top             =   6465
         Width           =   795
      End
      Begin VB.CommandButton cmdSaveInc 
         Caption         =   "&Save"
         Height          =   735
         Left            =   -66360
         Picture         =   "frmEditAll.frx":F176
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   5970
         Width           =   1155
      End
      Begin VB.Frame Frame9 
         Height          =   1035
         Left            =   -64965
         TabIndex        =   106
         Top             =   3300
         Visible         =   0   'False
         Width           =   2865
         Begin VB.CommandButton bPrintINR 
            Caption         =   "Print INR"
            Height          =   285
            Left            =   1290
            TabIndex        =   108
            Top             =   360
            Width           =   855
         End
         Begin VB.TextBox tWarfarin 
            Height          =   285
            Left            =   270
            MaxLength       =   5
            TabIndex        =   107
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Warfarin"
            Height          =   195
            Index           =   14
            Left            =   330
            TabIndex        =   109
            Top             =   150
            Width           =   600
         End
      End
      Begin VB.CommandButton bViewBioRepeat 
         BackColor       =   &H0000FFFF&
         Caption         =   "View Repeat"
         Height          =   915
         Left            =   -66315
         Picture         =   "frmEditAll.frx":F480
         Style           =   1  'Graphical
         TabIndex        =   104
         ToolTipText     =   "View Repeated Tests"
         Top             =   6180
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.TextBox txtHaemComment 
         Height          =   705
         Left            =   -72840
         MaxLength       =   2000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   101
         ToolTipText     =   "Only 320 Characters"
         Top             =   4890
         Width           =   7590
      End
      Begin VB.PictureBox Panel3D7 
         Height          =   4845
         Left            =   -64200
         ScaleHeight     =   4785
         ScaleWidth      =   2235
         TabIndex        =   93
         Top             =   705
         Width           =   2295
         Begin VB.CheckBox chkBad 
            Caption         =   "Bad Result"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   240
            TabIndex        =   211
            ToolTipText     =   "Allows Bad Samples to be Counted"
            Top             =   4410
            Width           =   1275
         End
         Begin VB.CheckBox chkMalaria 
            Caption         =   "Malaria Screen"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   240
            TabIndex        =   196
            Top             =   3180
            Width           =   1635
         End
         Begin VB.CheckBox chkSickledex 
            Caption         =   "Sickle Screen"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   240
            TabIndex        =   195
            Top             =   3825
            Width           =   1995
         End
         Begin VB.TextBox tASOt 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   240
            TabIndex        =   162
            ToolTipText     =   "Antistreptococcal Antibody Titres"
            Top             =   2790
            Width           =   1065
         End
         Begin VB.CheckBox cASot 
            Caption         =   "ASOT"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   240
            TabIndex        =   161
            Top             =   2565
            Width           =   1155
         End
         Begin VB.TextBox tRa 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   240
            TabIndex        =   160
            ToolTipText     =   "Rheumatoid Factor"
            Top             =   2205
            Width           =   1065
         End
         Begin VB.CheckBox cRA 
            Caption         =   "RF"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   240
            TabIndex        =   159
            Top             =   1980
            Width           =   615
         End
         Begin VB.TextBox tRetA 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   225
            TabIndex        =   125
            Top             =   990
            Width           =   705
         End
         Begin VB.CheckBox cMonospot 
            Caption         =   "Monospot/IM"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   240
            TabIndex        =   100
            Top             =   1410
            Width           =   1995
         End
         Begin VB.TextBox tMonospot 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   240
            TabIndex        =   99
            Top             =   1620
            Width           =   1065
         End
         Begin VB.CheckBox cRetics 
            Caption         =   "Retics"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   240
            TabIndex        =   98
            Top             =   780
            Width           =   885
         End
         Begin VB.CheckBox cESR 
            Caption         =   "ESR"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   240
            TabIndex        =   97
            Top             =   180
            Width           =   765
         End
         Begin VB.TextBox tESR 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   240
            MaxLength       =   3
            TabIndex        =   96
            ToolTipText     =   "Erythrocyte Sedimentation Rate"
            Top             =   390
            Width           =   1035
         End
         Begin VB.TextBox tRetP 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   960
            MaxLength       =   4
            TabIndex        =   95
            Top             =   990
            Width           =   675
         End
         Begin VB.CommandButton cmdPrintEsr 
            Appearance      =   0  'Flat
            Height          =   435
            Left            =   1890
            Picture         =   "frmEditAll.frx":F60A
            Style           =   1  'Graphical
            TabIndex        =   94
            Top             =   4590
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.TextBox txtEsr1 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   225
            MaxLength       =   3
            TabIndex        =   275
            ToolTipText     =   "ESR1  Result"
            Top             =   405
            Visible         =   0   'False
            Width           =   1350
         End
         Begin VB.Label lblMalaria 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   240
            TabIndex        =   198
            Top             =   3390
            Width           =   1095
         End
         Begin VB.Label lblSickledex 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   240
            TabIndex        =   197
            Top             =   4020
            Width           =   1095
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   240
            Left            =   1710
            TabIndex        =   124
            Top             =   1020
            Width           =   210
         End
      End
      Begin VB.PictureBox Panel3D5 
         ForeColor       =   &H80000005&
         Height          =   3885
         Left            =   -74820
         ScaleHeight     =   3825
         ScaleWidth      =   3765
         TabIndex        =   87
         Top             =   750
         Width           =   3825
         Begin VB.TextBox txtMPXI 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2430
            MaxLength       =   5
            TabIndex        =   168
            ToolTipText     =   "Mean Peroxidase Index"
            Top             =   3510
            Width           =   825
         End
         Begin VB.TextBox txtLI 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   750
            MaxLength       =   5
            TabIndex        =   167
            Top             =   3510
            Width           =   825
         End
         Begin MSFlexGridLib.MSFlexGrid grdH 
            Height          =   2205
            Left            =   60
            TabIndex        =   164
            ToolTipText     =   "Analyser Differential"
            Top             =   825
            Width           =   3675
            _ExtentX        =   6482
            _ExtentY        =   3889
            _Version        =   393216
            Rows            =   7
            Cols            =   4
            FixedCols       =   0
            ScrollBars      =   0
            FormatString    =   "^Abs      |^Ref - Range|^Diff        |^%        "
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
         Begin VB.TextBox tWBC 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   750
            MaxLength       =   6
            TabIndex        =   89
            ToolTipText     =   "White Blood Count"
            Top             =   405
            Width           =   1005
         End
         Begin VB.CommandButton bClearDiff 
            Caption         =   "Clear &Diff"
            Height          =   315
            Left            =   2160
            TabIndex        =   88
            Top             =   450
            Width           =   1095
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "MPXI"
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
            Height          =   195
            Index           =   19
            Left            =   1890
            TabIndex        =   170
            Top             =   3540
            Width           =   465
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "LI"
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
            Height          =   195
            Index           =   4
            Left            =   390
            TabIndex        =   169
            Top             =   3540
            Width           =   180
         End
         Begin VB.Label lWOC 
            Alignment       =   2  'Center
            BackColor       =   &H80000005&
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
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2430
            TabIndex        =   119
            Top             =   3180
            Width           =   810
         End
         Begin VB.Label lWIC 
            Alignment       =   2  'Center
            BackColor       =   &H80000005&
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
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   780
            TabIndex        =   118
            Top             =   3180
            Width           =   810
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   " WOC"
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
            Height          =   195
            Index           =   18
            Left            =   1830
            TabIndex        =   117
            Top             =   3210
            Width           =   510
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "WIC"
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
            Height          =   195
            Index           =   17
            Left            =   210
            TabIndex        =   116
            Top             =   3210
            Width           =   555
         End
         Begin VB.Label ipflag 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H0000FFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Suspect"
            ForeColor       =   &H00800080&
            Height          =   225
            Index           =   1
            Left            =   90
            TabIndex        =   92
            Top             =   90
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.Label ipflag 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H0000FFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Abnormal"
            ForeColor       =   &H00800080&
            Height          =   225
            Index           =   0
            Left            =   2970
            TabIndex        =   91
            Top             =   90
            Visible         =   0   'False
            Width           =   780
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "WBC"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   0
            Left            =   180
            TabIndex        =   90
            Top             =   495
            Width           =   525
         End
      End
      Begin VB.PictureBox Panel3D4 
         Height          =   3855
         Left            =   -70800
         ScaleHeight     =   3795
         ScaleWidth      =   4215
         TabIndex        =   84
         Top             =   840
         Width           =   4275
         Begin MSFlexGridLib.MSFlexGrid gRbc 
            Height          =   3405
            Left            =   120
            TabIndex        =   165
            Top             =   360
            Width           =   4035
            _ExtentX        =   7117
            _ExtentY        =   6006
            _Version        =   393216
            Cols            =   3
            FixedCols       =   0
            HighLight       =   2
            FormatString    =   "^FBC                 |^Result           |^Ref Range  "
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
         Begin VB.TextBox txtInput 
            Height          =   285
            Left            =   4590
            TabIndex        =   172
            Top             =   3105
            Width           =   645
         End
         Begin VB.Label lblAgeSex 
            Alignment       =   2  'Center
            BackColor       =   &H00C00000&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   810
            TabIndex        =   268
            Top             =   45
            Visible         =   0   'False
            Width           =   2625
         End
         Begin VB.Label ipflag 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H0000FFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Abnormal"
            ForeColor       =   &H000000FF&
            Height          =   225
            Index           =   2
            Left            =   3465
            TabIndex        =   86
            Top             =   75
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label ipflag 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H0000FFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Suspect"
            ForeColor       =   &H000000FF&
            Height          =   225
            Index           =   3
            Left            =   45
            TabIndex        =   85
            Top             =   75
            Visible         =   0   'False
            Width           =   735
         End
      End
      Begin VB.Frame Frame4 
         Height          =   5565
         Left            =   -74550
         TabIndex        =   66
         Top             =   1530
         Width           =   5445
         Begin VB.TextBox txtGpId 
            Height          =   285
            Left            =   4050
            TabIndex        =   282
            Top             =   300
            Width           =   915
         End
         Begin VB.CommandButton cmdCopyTo 
            Caption         =   "++ cc ++"
            Height          =   960
            Left            =   4995
            TabIndex        =   274
            Top             =   2925
            Width           =   375
         End
         Begin VB.ComboBox cmbHospital 
            Height          =   315
            Left            =   1050
            TabIndex        =   8
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
            TabIndex        =   15
            ToolTipText     =   "Clinical Details"
            Top             =   4980
            Width           =   3915
         End
         Begin VB.TextBox txtDemographicComment 
            Height          =   990
            Left            =   1050
            MaxLength       =   160
            MultiLine       =   -1  'True
            TabIndex        =   14
            Tag             =   "Demographic Comment"
            ToolTipText     =   "Demographic Comment"
            Top             =   3930
            Width           =   3885
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "GpID"
            Height          =   195
            Index           =   12
            Left            =   3570
            TabIndex        =   281
            Top             =   330
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Hospital"
            Height          =   195
            Index           =   15
            Left            =   420
            TabIndex        =   193
            Top             =   2640
            Width           =   570
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "GP"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   8
            Left            =   765
            TabIndex        =   82
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
            TabIndex        =   81
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
            TabIndex        =   80
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
            TabIndex        =   79
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
            TabIndex        =   78
            Top             =   3000
            Width           =   390
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Sex"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   13
            Left            =   3660
            TabIndex        =   77
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
            TabIndex        =   76
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
            TabIndex        =   75
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
            TabIndex        =   74
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
            TabIndex        =   73
            Top             =   330
            Width           =   525
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Cl Details"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   16
            Left            =   330
            TabIndex        =   72
            Top             =   5040
            Width           =   660
         End
         Begin VB.Label lChart 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   720
            TabIndex        =   71
            Top             =   300
            Width           =   1485
         End
         Begin VB.Label lName 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   750
            TabIndex        =   70
            Top             =   780
            Width           =   4215
         End
         Begin VB.Label lDoB 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   750
            TabIndex        =   69
            Top             =   1230
            Width           =   1515
         End
         Begin VB.Label lAge 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2910
            TabIndex        =   68
            Top             =   1200
            Width           =   585
         End
         Begin VB.Label lSex 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   3990
            TabIndex        =   67
            Top             =   1200
            Width           =   705
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Biochemistry Comments"
         Height          =   1845
         Left            =   -65010
         TabIndex        =   28
         Top             =   2940
         Width           =   3165
         Begin VB.TextBox txtBioComment 
            BackColor       =   &H80000018&
            Height          =   1545
            Left            =   150
            MaxLength       =   2000
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
         Left            =   -66480
         TabIndex        =   30
         ToolTipText     =   "Outstanding Tests"
         Top             =   795
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
         Left            =   -67665
         Picture         =   "frmEditAll.frx":FC74
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   5970
         Width           =   1155
      End
      Begin VB.CommandButton cmdValidateCoag 
         Caption         =   "&Validate"
         Height          =   915
         Left            =   -63435
         Picture         =   "frmEditAll.frx":FF7E
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Result Validation"
         Top             =   6180
         Width           =   765
      End
      Begin VB.CommandButton cmdSaveCoag 
         Caption         =   "Save Details"
         Enabled         =   0   'False
         Height          =   915
         Left            =   -64290
         Picture         =   "frmEditAll.frx":10288
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Save Changes"
         Top             =   6180
         Width           =   765
      End
      Begin VB.CommandButton bValidateBio 
         Caption         =   "Validate"
         Height          =   915
         Left            =   -63840
         Picture         =   "frmEditAll.frx":10592
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Result Validation"
         Top             =   6180
         Width           =   765
      End
      Begin VB.CommandButton cmdSaveBio 
         Caption         =   "Save Details"
         Enabled         =   0   'False
         Height          =   915
         Left            =   -64650
         Picture         =   "frmEditAll.frx":1089C
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Save Changes"
         Top             =   6180
         Width           =   765
      End
      Begin VB.CommandButton bValidateHaem 
         Caption         =   "Validate"
         Height          =   825
         Left            =   -62670
         Picture         =   "frmEditAll.frx":10BA6
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Result Validation"
         Top             =   6465
         Width           =   765
      End
      Begin VB.CommandButton cmdSaveHaem 
         Caption         =   "Save"
         Enabled         =   0   'False
         Height          =   825
         Left            =   -64560
         Picture         =   "frmEditAll.frx":10EB0
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Save Changes"
         Top             =   6465
         Width           =   825
      End
      Begin VB.CommandButton bViewHaemRepeat 
         BackColor       =   &H0000FFFF&
         Caption         =   "View Repeat"
         Height          =   825
         Left            =   -65550
         Picture         =   "frmEditAll.frx":111BA
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "View Repeated Tests"
         Top             =   6465
         Width           =   960
      End
      Begin VB.CommandButton bViewCoagRepeat 
         BackColor       =   &H0000FFFF&
         Caption         =   "View Repeat"
         Height          =   915
         Left            =   -66045
         Picture         =   "frmEditAll.frx":11344
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "View Repeated Tests"
         Top             =   6180
         Width           =   795
      End
      Begin VB.CommandButton bAddCoag 
         Caption         =   "Add Result"
         Height          =   915
         Left            =   -66855
         Picture         =   "frmEditAll.frx":114CE
         Style           =   1  'Graphical
         TabIndex        =   40
         Tag             =   "bAdd"
         ToolTipText     =   "Add Result Manually"
         Top             =   6180
         Width           =   765
      End
      Begin VB.PictureBox Panel3D8 
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00FFFF80&
         Height          =   1425
         Left            =   -74865
         ScaleHeight     =   1365
         ScaleWidth      =   6210
         TabIndex        =   41
         Top             =   5700
         Width           =   6270
         Begin VB.VScrollBar VScroll1 
            Height          =   1215
            LargeChange     =   500
            Left            =   5805
            Max             =   2500
            SmallChange     =   100
            TabIndex        =   42
            Top             =   45
            Width           =   270
         End
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H80000008&
            Height          =   1215
            Left            =   45
            ScaleHeight     =   1185
            ScaleWidth      =   5670
            TabIndex        =   43
            Top             =   90
            Width           =   5700
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
               Height          =   3015
               Left            =   45
               ScaleHeight     =   3015
               ScaleWidth      =   5595
               TabIndex        =   44
               Top             =   45
               Width           =   5595
            End
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Delta Check"
         Height          =   1785
         Left            =   -65010
         TabIndex        =   45
         Top             =   1140
         Width           =   3150
         Begin VB.Label ldelta 
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H000000FF&
            Height          =   1560
            Left            =   135
            TabIndex        =   46
            ToolTipText     =   "Delta Check"
            Top             =   180
            Width           =   2895
            WordWrap        =   -1  'True
         End
      End
      Begin VB.TextBox tResult 
         Height          =   315
         Left            =   -73200
         TabIndex        =   47
         ToolTipText     =   "Enter Result"
         Top             =   6330
         Width           =   1485
      End
      Begin VB.ComboBox cParameter 
         Height          =   315
         Left            =   -74760
         TabIndex        =   48
         Text            =   "cParameter"
         ToolTipText     =   "Choose Test"
         Top             =   6330
         Width           =   1545
      End
      Begin VB.CommandButton bAddBio 
         Caption         =   "Add Result"
         Height          =   915
         Left            =   -68070
         Picture         =   "frmEditAll.frx":117D8
         Style           =   1  'Graphical
         TabIndex        =   49
         Tag             =   "bAdd"
         ToolTipText     =   "Add Result Manually"
         Top             =   6180
         Width           =   765
      End
      Begin VB.Frame fraDate 
         Caption         =   "Sample Date"
         Height          =   1815
         Left            =   -69030
         TabIndex        =   50
         Top             =   1560
         Width           =   5805
         Begin MSComCtl2.DTPicker dtRunDate 
            Height          =   315
            Left            =   2370
            TabIndex        =   18
            Top             =   1050
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            _Version        =   393216
            Format          =   188743681
            CurrentDate     =   36942
         End
         Begin MSComCtl2.DTPicker dtSampleDate 
            Height          =   315
            Left            =   180
            TabIndex        =   12
            Top             =   315
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            _Version        =   393216
            Format          =   188743681
            CurrentDate     =   36942
         End
         Begin MSMask.MaskEdBox tSampleTime 
            Height          =   315
            Left            =   1560
            TabIndex        =   20
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
            Left            =   3540
            TabIndex        =   16
            Top             =   315
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            _Version        =   393216
            Format          =   188743681
            CurrentDate     =   36942
         End
         Begin MSMask.MaskEdBox tRecTime 
            Height          =   315
            Left            =   4920
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
         Begin VB.Label lblDateError 
            Alignment       =   2  'Center
            BackColor       =   &H000000FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Date Sequence Error"
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
            Height          =   675
            Left            =   4680
            TabIndex        =   290
            Top             =   1140
            Visible         =   0   'False
            Width           =   1125
         End
         Begin VB.Label lblDate 
            Alignment       =   1  'Right Justify
            Caption         =   "Run Date"
            Height          =   225
            Index           =   1
            Left            =   1350
            TabIndex        =   289
            Top             =   1110
            Width           =   930
         End
         Begin VB.Label lblDate 
            Alignment       =   2  'Center
            Caption         =   "Received in Lab"
            Height          =   255
            Index           =   0
            Left            =   3390
            TabIndex        =   288
            Top             =   0
            Width           =   1500
         End
         Begin VB.Image iRecDate 
            Height          =   330
            Index           =   0
            Left            =   3540
            Picture         =   "frmEditAll.frx":11AE2
            Stretch         =   -1  'True
            ToolTipText     =   "Previous Day"
            Top             =   630
            Width           =   480
         End
         Begin VB.Image iRecDate 
            Height          =   330
            Index           =   1
            Left            =   4410
            Picture         =   "frmEditAll.frx":11F24
            Stretch         =   -1  'True
            ToolTipText     =   "Next Day"
            Top             =   630
            Width           =   480
         End
         Begin VB.Image iToday 
            Height          =   330
            Index           =   2
            Left            =   4020
            Picture         =   "frmEditAll.frx":12366
            Stretch         =   -1  'True
            ToolTipText     =   "Set to Today"
            Top             =   630
            Width           =   360
         End
         Begin VB.Image iToday 
            Height          =   330
            Index           =   1
            Left            =   690
            Picture         =   "frmEditAll.frx":127A8
            Stretch         =   -1  'True
            ToolTipText     =   "Set to Today"
            Top             =   630
            Width           =   360
         End
         Begin VB.Image iToday 
            Height          =   330
            Index           =   0
            Left            =   2850
            Picture         =   "frmEditAll.frx":12BEA
            Stretch         =   -1  'True
            ToolTipText     =   "Set to Today"
            Top             =   1380
            Width           =   360
         End
         Begin VB.Image iSampleDate 
            Height          =   330
            Index           =   1
            Left            =   1050
            Picture         =   "frmEditAll.frx":1302C
            Stretch         =   -1  'True
            ToolTipText     =   "Next Day"
            Top             =   630
            Width           =   480
         End
         Begin VB.Image iSampleDate 
            Height          =   330
            Index           =   0
            Left            =   210
            Picture         =   "frmEditAll.frx":1346E
            Stretch         =   -1  'True
            ToolTipText     =   "Previous Day"
            Top             =   630
            Width           =   480
         End
         Begin VB.Image iRunDate 
            Height          =   330
            Index           =   1
            Left            =   3240
            Picture         =   "frmEditAll.frx":138B0
            Stretch         =   -1  'True
            ToolTipText     =   "Next Day"
            Top             =   1380
            Width           =   480
         End
         Begin VB.Image iRunDate 
            Height          =   330
            Index           =   0
            Left            =   2340
            Picture         =   "frmEditAll.frx":13CF2
            Stretch         =   -1  'True
            ToolTipText     =   "Previous Day"
            Top             =   1380
            Width           =   480
         End
      End
      Begin VB.Frame Frame5 
         Height          =   915
         Left            =   -66510
         TabIndex        =   51
         Top             =   3345
         Width           =   1455
         Begin VB.CheckBox chkUrgent 
            Alignment       =   1  'Right Justify
            Caption         =   "Urgent"
            Height          =   195
            Left            =   525
            TabIndex        =   216
            Top             =   660
            Width           =   795
         End
         Begin VB.OptionButton cRooH 
            Alignment       =   1  'Right Justify
            Caption         =   "Out of Hours"
            Height          =   195
            Index           =   1
            Left            =   90
            TabIndex        =   52
            Top             =   420
            Width           =   1215
         End
         Begin VB.OptionButton cRooH 
            Alignment       =   1  'Right Justify
            Caption         =   "Routine"
            Height          =   195
            Index           =   0
            Left            =   420
            TabIndex        =   53
            Top             =   180
            Width           =   885
         End
      End
      Begin VB.CommandButton bremoveduplicates 
         Caption         =   "Remove Duplicates"
         Height          =   915
         Left            =   -67260
         Picture         =   "frmEditAll.frx":14134
         Style           =   1  'Graphical
         TabIndex        =   54
         ToolTipText     =   "Remove Result"
         Top             =   6180
         Width           =   885
      End
      Begin VB.ComboBox cAdd 
         Height          =   315
         Left            =   -74925
         Sorted          =   -1  'True
         TabIndex        =   55
         Text            =   "cAdd"
         ToolTipText     =   "Choose Test"
         Top             =   6390
         Width           =   1695
      End
      Begin VB.TextBox tnewvalue 
         Height          =   315
         Left            =   -73155
         MaxLength       =   15
         TabIndex        =   56
         ToolTipText     =   "Enter Result"
         Top             =   6390
         Width           =   1335
      End
      Begin VB.ComboBox cUnits 
         Height          =   315
         Left            =   -71760
         TabIndex        =   57
         Text            =   "cUnits"
         ToolTipText     =   "Choose Units"
         Top             =   6390
         Width           =   1305
      End
      Begin VB.Frame Frame2 
         Caption         =   "Specimen Condition"
         Height          =   960
         Left            =   -65010
         TabIndex        =   58
         Top             =   4860
         Width           =   3165
         Begin VB.CheckBox oH 
            Caption         =   "Haemolysed"
            Height          =   225
            Left            =   1290
            TabIndex        =   59
            Top             =   450
            Width           =   1245
         End
         Begin VB.CheckBox oS 
            Caption         =   "Slightly Haemolysed"
            Height          =   225
            Left            =   1290
            TabIndex        =   60
            Top             =   210
            Width           =   1755
         End
         Begin VB.CheckBox oL 
            Alignment       =   1  'Right Justify
            Caption         =   "Lipaemic"
            Height          =   225
            Left            =   240
            TabIndex        =   61
            Top             =   210
            Width           =   975
         End
         Begin VB.CheckBox oO 
            Alignment       =   1  'Right Justify
            Caption         =   "Old Sample"
            Height          =   225
            Left            =   60
            TabIndex        =   62
            Top             =   450
            Width           =   1155
         End
         Begin VB.CheckBox oG 
            Caption         =   "Grossly Haemolysed"
            Height          =   225
            Left            =   1290
            TabIndex        =   63
            Top             =   675
            Width           =   1755
         End
         Begin VB.CheckBox oJ 
            Alignment       =   1  'Right Justify
            Caption         =   "Icteric"
            Height          =   225
            Left            =   450
            TabIndex        =   64
            Top             =   690
            Width           =   765
         End
      End
      Begin MSFlexGridLib.MSFlexGrid gBio 
         Height          =   5265
         Left            =   -74910
         TabIndex        =   65
         ToolTipText     =   "Biochemistry Results"
         Top             =   795
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   9287
         _Version        =   393216
         Cols            =   11
         BackColor       =   -2147483628
         ForeColor       =   -2147483635
         BackColorFixed  =   -2147483647
         ForeColorFixed  =   -2147483624
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         GridLines       =   3
         GridLinesFixed  =   3
         AllowUserResizing=   1
         FormatString    =   "<Test                  |<Result  |<Units    |^Ref Range  |^H/L|^   |^VP |^CP|^AL    |<Comment              |^P"
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
         TabIndex        =   83
         Top             =   840
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   7541
         _Version        =   393216
         Cols            =   9
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
         FormatString    =   "<Parameter            |<Result    |<Units       |^Ref Range    |<Flag|^V |^P |<Analyser |^P"
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
         Left            =   -66465
         TabIndex        =   112
         ToolTipText     =   "Tests still not run."
         Top             =   840
         Width           =   1455
         _ExtentX        =   2566
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid grdPrev 
         Height          =   2175
         Left            =   -64950
         TabIndex        =   129
         ToolTipText     =   "Do Not give Out Results!! Only Historical Results for comparision."
         Top             =   1110
         Width           =   2865
         _ExtentX        =   5054
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
         Left            =   -66960
         TabIndex        =   136
         ToolTipText     =   "Tests still not run."
         Top             =   795
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
         Left            =   -74865
         TabIndex        =   154
         Top             =   795
         Width           =   7950
         _ExtentX        =   14023
         _ExtentY        =   9287
         _Version        =   393216
         Cols            =   9
         RowHeightMin    =   315
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
         FormatString    =   "<Test                  |<Result              |<Units    |<Ref Range         |^H/L|^   |^VP |Comment  |^P"
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
         TabIndex        =   218
         Top             =   930
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
         Height          =   3270
         Index           =   1
         Left            =   11580
         TabIndex        =   243
         ToolTipText     =   "Tests still not run."
         Top             =   810
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   5768
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
         Left            =   180
         TabIndex        =   244
         Top             =   795
         Width           =   11400
         _ExtentX        =   20108
         _ExtentY        =   5794
         _Version        =   393216
         Cols            =   11
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
         FormatString    =   $"frmEditAll.frx":1443E
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
         TabIndex        =   267
         Top             =   825
         Width           =   11430
         _ExtentX        =   20161
         _ExtentY        =   6324
         _Version        =   393216
         Cols            =   9
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
         GridLines       =   2
         AllowUserResizing=   3
         FormatString    =   $"frmEditAll.frx":144EC
      End
      Begin VB.Frame Frame10 
         Caption         =   "Category"
         Height          =   615
         Index           =   1
         Left            =   -64290
         TabIndex        =   203
         Top             =   2010
         Width           =   2385
         Begin VB.ComboBox cCat 
            Height          =   315
            Index           =   1
            Left            =   90
            Sorted          =   -1  'True
            TabIndex        =   204
            Top             =   210
            Width           =   2205
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Delta Check"
         Height          =   1395
         Index           =   0
         Left            =   -65550
         TabIndex        =   139
         Top             =   720
         Width           =   3645
         Begin VB.Label lIDelta 
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H000000FF&
            Height          =   1005
            Index           =   0
            Left            =   60
            TabIndex        =   140
            ToolTipText     =   "Delta Check"
            Top             =   270
            Width           =   3540
            WordWrap        =   -1  'True
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
         Left            =   -63240
         TabIndex        =   320
         Top             =   3090
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label lblRepeats 
         AutoSize        =   -1  'True
         Caption         =   "Tests Repeated"
         Height          =   195
         Left            =   -63120
         TabIndex        =   285
         Top             =   6150
         Width           =   1140
      End
      Begin VB.Label Label1 
         Caption         =   "Analyser :"
         Height          =   195
         Index           =   20
         Left            =   -66990
         TabIndex        =   273
         Top             =   6150
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
         Left            =   -66270
         TabIndex        =   272
         Top             =   6105
         Width           =   2130
      End
      Begin VB.Label lblBgaDate 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   -73050
         TabIndex        =   269
         Top             =   5700
         Width           =   2865
      End
      Begin VB.Label Rundate 
         Caption         =   "Rundate"
         Height          =   255
         Index           =   5
         Left            =   -73695
         TabIndex        =   270
         Top             =   5730
         Width           =   675
      End
      Begin VB.Label lImmRan 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Random Sample"
         Height          =   465
         Index           =   1
         Left            =   7680
         TabIndex        =   255
         ToolTipText     =   "Click to Toggle"
         Top             =   5400
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
         Left            =   180
         TabIndex        =   254
         Top             =   6210
         Width           =   5955
      End
      Begin VB.Label lblIRundate 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3210
         TabIndex        =   253
         Top             =   6885
         Width           =   2910
      End
      Begin VB.Label Rundate 
         Caption         =   "Rundate"
         Height          =   255
         Index           =   4
         Left            =   2520
         TabIndex        =   252
         Top             =   6915
         Width           =   675
      End
      Begin VB.Image An2 
         Height          =   645
         Left            =   -62820
         Top             =   450
         Width           =   645
      End
      Begin VB.Image An1 
         Height          =   645
         Left            =   -63570
         Top             =   450
         Width           =   645
      End
      Begin VB.Label lblEDate 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   -73005
         TabIndex        =   214
         ToolTipText     =   "Run Date && Time"
         Top             =   6825
         Width           =   2595
      End
      Begin VB.Label Rundate 
         Caption         =   "Rundate"
         Height          =   255
         Index           =   0
         Left            =   -73740
         TabIndex        =   213
         Top             =   6855
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
         TabIndex        =   202
         ToolTipText     =   "Click to Toggle"
         Top             =   5745
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
         TabIndex        =   201
         Top             =   6150
         Width           =   5955
      End
      Begin VB.Label lblAss 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Associated  Glucose 1"
         Height          =   705
         Left            =   -62760
         TabIndex        =   194
         Top             =   6375
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "Condition"
         Height          =   255
         Index           =   11
         Left            =   -66585
         TabIndex        =   180
         Top             =   2235
         Width           =   1065
      End
      Begin VB.Label Rundate 
         Caption         =   "Rundate"
         Height          =   255
         Index           =   3
         Left            =   -74610
         TabIndex        =   178
         Top             =   6750
         Width           =   675
      End
      Begin VB.Label lCDate 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   -73830
         TabIndex        =   177
         Top             =   6720
         Width           =   2295
      End
      Begin VB.Label Rundate 
         Caption         =   "Rundate"
         Height          =   255
         Index           =   2
         Left            =   -73890
         TabIndex        =   176
         Top             =   6810
         Width           =   675
      End
      Begin VB.Label lBDate 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   -73155
         TabIndex        =   175
         ToolTipText     =   "Run Date && Time"
         Top             =   6780
         Width           =   2685
      End
      Begin VB.Label Rundate 
         Caption         =   "Rundate"
         Height          =   255
         Index           =   1
         Left            =   -66990
         TabIndex        =   174
         Top             =   5835
         Width           =   675
      End
      Begin VB.Label lHDate 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   -66270
         TabIndex        =   173
         Top             =   5790
         Width           =   2115
      End
      Begin VB.Label lblPrevCoag 
         Alignment       =   2  'Center
         BackColor       =   &H80000001&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "No Chart # for Previous Details"
         ForeColor       =   &H80000018&
         Height          =   255
         Left            =   -64920
         TabIndex        =   128
         Top             =   840
         Width           =   2805
      End
      Begin VB.Label lblHaemValid 
         AutoSize        =   -1  'True
         Caption         =   "Already Validated"
         Height          =   195
         Left            =   -63210
         TabIndex        =   123
         Top             =   5745
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
         Left            =   -74865
         TabIndex        =   115
         ToolTipText     =   "Click Here to Show Flags"
         Top             =   5070
         Width           =   1065
      End
      Begin VB.Label lRandom 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Random Sample"
         Height          =   465
         Left            =   -62760
         TabIndex        =   114
         ToolTipText     =   "Click to Toggle"
         Top             =   5835
         Width           =   915
      End
      Begin VB.Label lblHaemPrinted 
         AutoSize        =   -1  'True
         Caption         =   "Already Printed"
         Height          =   195
         Left            =   -63060
         TabIndex        =   111
         Top             =   5955
         Width           =   1065
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Coagulation Comments"
         Height          =   195
         Left            =   -64860
         TabIndex        =   110
         Top             =   4350
         Width           =   1620
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Haematology Comment"
         Height          =   195
         Index           =   10
         Left            =   -72795
         TabIndex        =   102
         Top             =   4680
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
         TabIndex        =   126
         Top             =   6090
         Width           =   6015
      End
   End
   Begin VB.Image imgGreenTick 
      Height          =   225
      Left            =   0
      Picture         =   "frmEditAll.frx":145AC
      Top             =   210
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image imgRedCross 
      Height          =   225
      Left            =   0
      Picture         =   "frmEditAll.frx":14882
      Top             =   0
      Visible         =   0   'False
      Width           =   210
   End
End
Attribute VB_Name = "frmEditAll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmEditAll
' Author    : Trevor Dunican
' Date      : 09/10/2015
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit

Private mNewRecord As Boolean

Private PreviousImm As Boolean
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
Private SampleType As String
Private grd As MSFlexGrid

Private Function SavePrintInhibit() As Boolean
'Returns True if there is something to print

    Dim sql As String
    Dim Y As Integer
    Dim Discipline As String
    Dim g As MSFlexGrid

10  On Error GoTo SavePrintInhibit_Error

20  Discipline = ""

30  Select Case ssTabAll.Tab
    Case 0: SavePrintInhibit = True
40  Case 1: SavePrintInhibit = True
50  Case 2: SavePrintInhibit = False: Discipline = "Bio": Set g = gBio
60  Case 3: SavePrintInhibit = False: Discipline = "Coa": Set g = grdCoag
70  Case 4: SavePrintInhibit = False: Discipline = "End": Set g = gImm(0)
80  Case 5: SavePrintInhibit = True
90  Case 6: SavePrintInhibit = False: Discipline = "Imm": Set g = gImm(1)
100 Case 7: SavePrintInhibit = True
110 End Select

120 If Discipline = "" Then Exit Function

130 sql = "DELETE FROM PrintInhibit WHERE " & _
          "SampleID = '" & txtSampleID & "' " & _
          "AND Discipline = '" & Discipline & "'"
140 Cnxn(0).Execute sql

150 If ssTabAll.Tab = 6 Then
160     g.Col = 9
170 Else
180     g.Col = g.Cols - 1
190 End If
200 For Y = 1 To g.Rows - 1
210     g.Row = Y
220     If g.CellPicture = imgRedCross.Picture Then
230         sql = "INSERT INTO PrintInhibit " & _
                  "(SampleID, Discipline, Parameter) VALUES " & _
                  "('" & txtSampleID & "', " & _
                " '" & Discipline & "', " & _
                " '" & g.TextMatrix(Y, 0) & "' )"
240         Cnxn(0).Execute sql
250     ElseIf g.CellPicture = imgGreenTick.Picture Then
260         SavePrintInhibit = True
270     End If
280 Next

290 Select Case ssTabAll.Tab
    Case 2: If Len(Trim$(txtBioComment)) > 0 Then SavePrintInhibit = True    'Bio
300 Case 3: If Len(Trim$(txtCoagComment)) > 0 Then SavePrintInhibit = True    'Coag
310 Case 4: If Len(Trim$(txtImmComment(0))) > 0 Then SavePrintInhibit = True    'Endo
320 Case 6: If Len(Trim$(txtImmComment(1))) > 0 Then SavePrintInhibit = True    'Imm
330 End Select
340 Exit Function

SavePrintInhibit_Error:

    Dim strES As String
    Dim intEL As Integer

350 intEL = Erl
360 strES = Err.Description
370 LogError "frmEditAll", "SavePrintInhibit", intEL, strES, sql

End Function
Private Function IsAllergy() As Boolean


    Dim tb As Recordset
    Dim sql As String
    Dim RetVal As Boolean
    Dim Content As String
    Dim i As Integer

10  On Error GoTo IsAllergy_Error


20  RetVal = False
30  Content = ""
40  With gImm(1)
50      For i = 1 To .Rows - 1
60          gImm(1).Row = i
70          gImm(1).Col = 9
80          If gImm(1).CellPicture = imgGreenTick Then
90              If UCase$(Mid$(.TextMatrix(i, 0), 2, 1)) = "X" Then
100                 RetVal = True
110                 Exit For
120             Else
130                 Content = Content & "LongName = '" & .TextMatrix(i, 10) & "' OR "
140             End If
150         End If
160     Next
170 End With
180 If Not RetVal And Content <> "" Then
190     Content = Left$(Content, Len(Content) - 3)
200     sql = "SELECT COUNT(*) AS Tot FROM ImmTestDefinitions WHERE " & _
              "IsAllergy = 1 " & _
              "AND (" & Content & ")"
210     Set tb = New Recordset
220     RecOpenServer 0, tb, sql
230     RetVal = tb!Tot > 0
240 End If

250 IsAllergy = RetVal


260 Exit Function

IsAllergy_Error:

    Dim strES As String
    Dim intEL As Integer

270 intEL = Erl
280 strES = Err.Description
290 LogError "frmEditAll", "IsAllergy", intEL, strES, sql


End Function

Private Sub SetPrintInhibit(ByVal Dept As String)

    Dim Y As Integer
    Dim FpIndex As Integer  'Food panel index
    Dim PcIndex As Integer  'Panel constituent index
    Dim tb As Recordset
    Dim sql As String

10  On Error GoTo SetPrintInhibit_Error

20  Select Case Dept
    Case "Imm"
30      With gImm(1)

40          If .TextMatrix(1, 0) <> "" Then
50              For FpIndex = 1 To .Rows - 1
60                  .Row = FpIndex: .Col = 9
70                  If .CellPicture = 0 Then
80                      If InStr(.TextMatrix(FpIndex, 6), "P") Then
90                          Set .CellPicture = imgRedCross.Picture
100                     Else
110                         Set .CellPicture = imgGreenTick.Picture
120                     End If
130                 End If
140                 If InStr(UCase(.TextMatrix(FpIndex, 0)), "X") And UCase(.TextMatrix(FpIndex, 1)) = "NEGATIVE" Then
                        'FOOD PANEL FOUND (get all panel constituents)
150                     sql = "SELECT Content FROM IPanels WHERE " & _
                              "PanelType = 'AL' " & _
                              "AND PanelName = '" & .TextMatrix(FpIndex, 10) & "' " & _
                              "AND Hospital = '" & HospName(0) & "'"
160                     Set tb = New Recordset
170                     RecOpenClient 0, tb, sql
180                     If Not tb.EOF Then
190                         For PcIndex = 1 To .Rows - 1
200                             If InStr(UCase(.TextMatrix(PcIndex, 0)), "X") = 0 Then
210                                 tb.MoveFirst
220                                 Do While Not tb.EOF
230                                     If UCase(.TextMatrix(PcIndex, 10)) = UCase(tb!Content) Then
240                                         .Row = PcIndex: .Col = 9
250                                         Set .CellPicture = imgRedCross.Picture
260                                     End If
270                                     tb.MoveNext
280                                 Loop
290                             End If
300                         Next PcIndex
310                     End If
320                 End If
330             Next FpIndex
340         End If
350     End With
360 Case "End"
370     gImm(0).Col = 8
380     If gImm(0).TextMatrix(1, 0) <> "" Then
390         For Y = 1 To gImm(0).Rows - 1
400             If InStr(gImm(0).TextMatrix(Y, 6), "P") Then
410                 gImm(0).Row = Y
420                 Set gImm(0).CellPicture = imgRedCross.Picture
430             Else
440                 gImm(0).Row = Y
450                 Set gImm(0).CellPicture = imgGreenTick.Picture
460             End If
470         Next
480     End If
490 Case "Bio"
500     gBio.Col = 10
510     If gBio.TextMatrix(1, 0) <> "" Then
520         For Y = 1 To gBio.Rows - 1
530             gBio.Row = Y
540             If InStr(gBio.TextMatrix(Y, 6), "P") Or gBio.CellBackColor = vbRed Then
550                 Set gBio.CellPicture = imgRedCross.Picture
560             Else
570                 gBio.Row = Y
580                 Set gBio.CellPicture = imgGreenTick.Picture
590             End If
600         Next
610     End If
620 Case "Coa"
630     grdCoag.Col = 8
640     If grdCoag.TextMatrix(1, 0) <> "" Then
650         For Y = 1 To grdCoag.Rows - 1
660             If InStr(grdCoag.TextMatrix(Y, 6), "P") Then
670                 grdCoag.Row = Y
680                 Set grdCoag.CellPicture = imgRedCross.Picture
690             Else
700                 grdCoag.Row = Y
710                 Set grdCoag.CellPicture = imgGreenTick.Picture
720             End If
730         Next
740     End If
750 End Select

760 Exit Sub

SetPrintInhibit_Error:

    Dim strES As String
    Dim intEL As Integer

770 intEL = Erl
780 strES = Err.Description
790 LogError "frmEditAll", "SetPrintInhibit", intEL, strES

End Sub

Private Sub bAddBio_Click()

    Dim tb As New Recordset
    Dim sql As String
    Dim n As Long
    Dim s As String

10  On Error GoTo bAddBio_Click_Error

20  pBar = 0

30  txtSampleID = Format(Val(txtSampleID))

40  For n = 1 To gBio.Rows - 1
50      If cAdd = gBio.TextMatrix(n, 0) Then
60          iMsg "Test already Exists. Please delete before adding!"
70          Exit Sub
80      End If
90  Next

100 s = Check_Bio(cAdd.Text, cUnits, cISampleType(3))
110 If s <> "" Then
120     iMsg s & " is incorrect!"
130     Exit Sub
140 End If

150 If cAdd.Text = "" Then Exit Sub
160 If Val(txtSampleID) = 0 Then Exit Sub
170 If Len(cUnits) = 0 Then
180     If iMsg("SELECT Units?", vbYesNo) = vbYes Then
190         Exit Sub
200     End If
210 End If

220 sql = "INSERT into BioResults " & _
          "(RunDate, SampleID, Code, Result, RunTime, Units, SampleType, Valid, Printed) VALUES " & _
          "('" & Format$(dtRunDate, "dd/mmm/yyyy") & "', " & _
          "'" & txtSampleID & "', " & _
          "'" & CodeForShortName(cAdd.Text) & "', " & _
          "'" & tnewvalue & "', " & _
          "'" & Format$(Now, "dd/mmm/yyyy hh:mm") & "', " & _
          "'" & cUnits & "', " & _
          "'" & ListCodeFor("ST", cISampleType(3)) & "', 0, 0);"
230 Cnxn(0).Execute sql

240 sql = "DELETE FROM BioRequests " & _
          "WHERE SampleID = '" & txtSampleID & "' " & _
          "AND Code = '" & CodeForShortName(cAdd.Text) & "'"
250 Cnxn(0).Execute sql

    'Code added 22/08/05
    'This allows the user delete
    'oustanding requests where sample is bad
    'it also marks bad samples printed and valid
260 If SysOptBioCodeForBad(0) = CodeForShortName(cAdd.Text) Then
270     sql = "update bioresults set valid = 1, printed = 1, Operator = '" & AddTicks(UserCode) & "' " & _
              "where code = '" & SysOptBioCodeForBad(0) & "' " & _
              "and sampleID = '" & txtSampleID & "'"
280     Cnxn(0).Execute sql
290     If iMsg("Do you wish all outstanding requests Deleted!", vbYesNo) = vbYes Then
300         sql = "DELETE from biorequests WHERE sampleID = '" & txtSampleID & "'"
310         Cnxn(0).Execute sql
320     End If
330     txtBioComment = Trim(txtBioComment & " " & iBOX("Enter Bad Comment"))
340     SaveComments
350 End If

360 LoadBiochemistry
370 LoadComments
380 cAdd = ""
390 tnewvalue = ""
400 cUnits = ""

410 Exit Sub

bAddBio_Click_Error:

    Dim strES As String
    Dim intEL As Integer

420 intEL = Erl
430 strES = Err.Description
440 LogError "frmEditAll", "bAddBio_Click", intEL, strES, sql

End Sub

Private Sub bAddCoag_Click()

    Dim Code As String
    Dim s As String
    Dim sql As String
    Dim Num As Long

10  On Error GoTo bAddCoag_Click_Error

20  pBar = 0

30  If cParameter = "" Then Exit Sub
40  If Trim$(tResult) = "" Then Exit Sub

50  For Num = 1 To grdCoag.Rows - 1
60      If grdCoag.TextMatrix(Num, 0) = cParameter Then
70          iMsg "Result already exists!"
80          Exit Sub
90      End If
100 Next

110 Code = CoagCodeFor(cParameter)
120 s = cParameter & vbTab & _
        tResult & vbTab & _
        cCunits & vbTab & _
        vbTab & vbTab & vbTab & vbTab & _
        "Manual"
130 grdCoag.AddItem s

140 If grdCoag.TextMatrix(1, 0) = "" Then
150     grdCoag.RemoveItem 1
160 End If

170 sql = "INSERT into CoagResults " & _
          "(RunDate, SampleID, Code, Result, RunTime, Units, Valid, Printed, Analyser) VALUES " & _
          "('" & Format$(dtRunDate, "dd/mmm/yyyy") & "', " & _
          "'" & txtSampleID & "', " & _
          "'" & Trim(CoagCodeFor(cParameter.Text)) & "', " & _
          "'" & tResult & "', " & _
          "'" & Format$(Now, "dd/mmm/yyyy hh:mm") & "', " & _
          "'" & cCunits & "', " & _
          "0, 0, 'Manual');"

180 Cnxn(0).Execute sql

    'Code added 22/08/05
    'Remove Coag requests if required
    'set bad result to valid and printed

190 If SysOptCBad(0) = CoagCodeFor(cParameter) Then
200     sql = "update coagresults set valid = 1, printed = 1, Username = '" & AddTicks(UserCode) & "'  " & _
              "where code = '" & SysOptCBad(0) & "' " & _
              "and sampleID = '" & txtSampleID & "'"
210     Cnxn(0).Execute sql
220     If iMsg("Do you wish all outstanding requests Deleted!", vbYesNo) = vbYes Then
230         sql = "DELETE from coagrequests WHERE sampleID = '" & txtSampleID & "'"
240         Cnxn(0).Execute sql
250     End If
260 End If

270 For Num = 1 To grdOutstandingCoag.Rows - 1
280     If grdOutstandingCoag.TextMatrix(Num, 0) = cParameter Then
290         sql = "DELETE from coagRequests WHERE " & _
                  "SampleID = '" & txtSampleID & "' " & _
                  "and code = '" & CoagCodeFor(cParameter) & "'"
300         Cnxn(0).Execute sql
310         LoadOutstandingrdCoag
320         Exit For
330     End If
340 Next

350 LoadCoagulation

360 cParameter = ""
370 tResult = ""
380 cCunits.ListIndex = -1
    'cmdSaveCoag.Enabled = True
390 cmdValidateCoag.Enabled = True
400 cmdValidateCoag.Caption = "&Validate"

410 Exit Sub

bAddCoag_Click_Error:

    Dim strES As String
    Dim intEL As Integer

420 intEL = Erl
430 strES = Err.Description
440 LogError "frmEditAll", "bAddCoag_Click", intEL, strES, sql

End Sub

Private Sub baddtotests_Click(Index As Integer)

    Dim MediBridgePathToViewer As String

10  On Error GoTo baddtotests_Click_Error

20  If Index = 0 Then
30      With frmAddToTests
40          .sex = txtSex
50          .SampleID = txtSampleID
60          .ClinDetails = cClDetails
70          .SampleDate = dtSampleDate
80          .SampleTime = tSampleTime
90          .Department = "General"
100         .Ward = cmbWard
110         .Clinician = cmbClinician
120         .GP = cmbGP
130         .Show 1
140     End With

150     LoadExt
160 ElseIf Index = 1 Then
170     If baddtotests(Index).BackColor <> vbYellow Then Exit Sub
        'view external results (Path changed to app.path because custom path was creating trouble
180     MediBridgePathToViewer = App.Path & "\MediBridgeViewer.exe "             ' GetOptionSetting("MedibridgePathToViewer", "")
190     If MediBridgePathToViewer <> "" Then
200         Shell MediBridgePathToViewer & " /SampleID=" & txtSampleID & _
                " /UserName=""" & UserName & """" & _
                " /Password=""" & UserPass & """" & _
                " /Department=""Medibridge""", vbNormalFocus
210     End If
220 End If

230 Exit Sub

baddtotests_Click_Error:

    Dim strES As String
    Dim intEL As Integer

240 intEL = Erl
250 strES = Err.Description
260 LogError "frmEditAll", "baddtotests_Click", intEL, strES

End Sub

Private Sub bcancel_Click()

10  On Error GoTo bCancel_Click_Error

20  pBar = 0

30  Unload Me

40  Exit Sub

bCancel_Click_Error:

    Dim strES As String
    Dim intEL As Integer

50  intEL = Erl
60  strES = Err.Description
70  LogError "frmEditAll", "bcancel_Click", intEL, strES


End Sub

Private Sub bcleardiff_click()
    Dim n As Long
    Dim A As Long


10  On Error GoTo bcleardiff_click_Error

20  pBar = 0

    'If SysOptHaemAn1(0) <> "ADVIA" Then
30  lWIC = ""
40  lWOC = ""
    'End If

50  txtMPXI = ""
60  txtLI = ""

70  grdH.Visible = False
80  For n = 1 To 6
90      For A = 0 To 3 Step 3
100         grdH.Row = n
110         grdH.Col = A
120         grdH.CellBackColor = &HFFFFFF
130         grdH.CellForeColor = 1
140         grdH = ""
150     Next
160 Next

170 grdH.Visible = True

180 cmdSaveHaem.Enabled = True: cmdHSaveH.Enabled = True
190 bValidateHaem.Enabled = True




200 Exit Sub

bcleardiff_click_Error:

    Dim strES As String
    Dim intEL As Integer

210 intEL = Erl
220 strES = Err.Description
230 LogError "frmEditAll", "bcleardiff_click", intEL, strES


End Sub

Private Sub bDoB_Click()


10  On Error GoTo bDoB_Click_Error

20  pBar = 0

30  With frmPatHistoryNew
40      .oHD(1) = True
50      .oFor(2) = True
60      .txtName = txtDoB
70      If cmdDemoVal.Caption = "VALID" Then .mDemoVal = True Else .mDemoVal = False
80      .FromEdit = True
90      .EditScreen = Me
100     .bsearch = True
110     If .g.TextMatrix(1, 13) <> "" Then
120         .Show 1
130     Else
140         FlashNoPrevious lNoPrevious
150     End If
160 End With




170 Exit Sub

bDoB_Click_Error:

    Dim strES As String
    Dim intEL As Integer

180 intEL = Erl
190 strES = Err.Description
200 LogError "frmEditAll", "bDoB_Click", intEL, strES


End Sub

Private Sub bFAX_Click()
    Dim tb As New Recordset
    Dim sql As String
    Dim FaxNumber As String
    Dim Disp As String
    Dim Department As String

10  On Error GoTo bFAX_Click_Error

20  pBar = 0


30  If ssTabAll.Tab = 1 And lblHaemValid.Visible = False Then
40      iMsg "Haematology not Validated"
50      Exit Sub
60  End If


70  pBar = 0

80  If Trim$(txtSex) = "" Then
90      If iMsg("Sex not entered." & vbCrLf & "Do you want to enter sex now?", vbQuestion + vbYesNo) = vbYes Then
100         Exit Sub
110     End If
120 End If

130 If Trim$(txtSampleID) = "" Then
140     iMsg "Must have Lab Number.", vbCritical
150     Exit Sub
160 End If

170 If Len(cmbWard) = 0 Then
180     iMsg "Must have Ward entry.", vbCritical
190     Exit Sub
200 End If

210 If UCase(Trim$(cmbWard)) = "GP" Then
220     If Len(cmbGP) = 0 Then
230         iMsg "Must have Ward or GP entry.", vbCritical
240         Exit Sub
250     End If
260 End If


270 If UCase(cmbWard) = "GP" Then
280     sql = "SELECT * from GPS WHERE text = '" & AddTicks(cmbGP) & "' and hospitalcode = '" & ListCodeFor("HO", cmbHospital) & "' and INUSE = 1"
290     Set tb = New Recordset
300     RecOpenServer 0, tb, sql
310     If Not tb.EOF Then
320         FaxNumber = Trim$(tb!FAX & "")
330     End If
340 Else
350     sql = "SELECT * from wards WHERE text = '" & AddTicks(cmbWard) & "' and hospitalcode = '" & ListCodeFor("HO", cmbHospital) & "'"
360     Set tb = New Recordset
370     RecOpenServer 0, tb, sql
380     If Not tb.EOF Then
390         FaxNumber = Trim$(tb!FAX & "")
400     End If
410 End If


420 FaxNumber = iBOX("Faxnumber ", , FaxNumber)

430 FaxNumber = Trim(FaxNumber)

440 If Trim(FaxNumber) = "" Then
450     iMsg "No Fax Number Entered!" & vbCrLf & "Fax Cancelled!"
460     Exit Sub
470 End If


480 If Not IsNumeric(FaxNumber) Then
490     iMsg "Incorrect Fax Number Entered!" & vbCrLf & "Fax Cancelled!"
500     Exit Sub
510 End If

520 Disp = Choose(ssTabAll.Tab, "H", "B", "C", "E", "Q", "I", "")

530 If Len(FaxNumber) < 4 Then
540     iMsg "Incorrect Fax Number Entered!" & vbCrLf & "Fax Cancelled!"
550     Exit Sub
560 End If

570 SaveDemographics

580 sql = "If Exists(Select 1 From PrintPending " & _
          "Where SampleID = @SampleID0 " & _
          "And Department = '@Department1' " & _
          "And FaxNumber = '@FaxNumber4' ) " & _
          "Begin " & _
          "Update PrintPending Set " & _
          "SampleID = @SampleID0, " & _
          "Department = '@Department1', " & _
          "Initiator = '@Initiator2', " & _
          "UsePrinter = '@UsePrinter3', " & _
          "FaxNumber = '@FaxNumber4', " & _
          "ptime = '@ptime5', " & _
          "Ward = '@Ward8', " & _
          "Clinician = '@Clinician9', " & _
          "GP = '@GP10' " & _
          "Where SampleID = @SampleID0 " & _
          "And Department = '@Department1' " & _
          "And FaxNumber = '@FaxNumber4'  " & _
              "End  "
590 sql = sql & "Else " & _
          "Begin  " & _
          "Insert Into PrintPending (SampleID, Department, Initiator, UsePrinter, FaxNumber, " & _
          "ptime, Ward, Clinician, GP) " & _
          "Values " & _
          "(@SampleID0, '@Department1', '@Initiator2', '@UsePrinter3', '@FaxNumber4', " & _
          "'@ptime5', '@Ward8', '@Clinician9', '@GP10') " & _
          "End"

600 If SysOptFaxCom(0) And (Disp = "H" Or Disp = "C" Or Disp = "B") Then
610     If ssTabAll.Tab <> 0 Then

620         sql = Replace(sql, "@SampleID0", txtSampleID)
630         sql = Replace(sql, "@Department1", "M")
640         sql = Replace(sql, "@Initiator2", UserName)
650         sql = Replace(sql, "@UsePrinter3", pPrintToPrinter)
660         sql = Replace(sql, "@FaxNumber4", FaxNumber)
670         sql = Replace(sql, "@ptime5", Format(Now, "dd/MMM/yyyy hh:mm:ss"))
680         sql = Replace(sql, "@Ward8", AddTicks(cmbWard))
690         sql = Replace(sql, "@Clinician9", AddTicks(cmbClinician))
700         sql = Replace(sql, "@GP10", AddTicks(cmbGP))

710         Cnxn(0).Execute sql



            '        sql = "SELECT * FROM PrintPending WHERE " & _
                     '              "Department = 'M' " & _
                     '              "AND SampleID = '" & txtSampleID & "' " & _
                     '              "AND COALESCE(FaxNumber,'') <> ''"
            '        Set tb = New Recordset
            '        RecOpenServer 0, tb, sql
            '        If tb.EOF Then
            '            tb.AddNew
            '        End If
            '        tb!Ward = cmbWard
            '        tb!Clinician = cmbClinician
            '        tb!GP = cmbGP
            '        tb!SampleID = txtSampleID
            '        tb!Department = "M"
            '        tb!pTime = Now
            '        tb!Initiator = Username
            '        tb!UsePrinter = pPrintToPrinter
            '        tb!FaxNumber = FaxNumber
            '        tb.Update
720     End If
730 Else
740     If ssTabAll.Tab <> 0 Then
750         LogTimeOfPrinting txtSampleID, Choose(ssTabAll.Tab, "H", "B", "C", "E", "Q", "I", "")


760         Department = Choose(ssTabAll.Tab, "H", "B", "C", "E", "Q", "I", "")
770         If SysOptRealImm(0) And Department = "I" Then Department = "J"

780         sql = Replace(sql, "@SampleID0", txtSampleID)
790         sql = Replace(sql, "@Department1", Department)
800         sql = Replace(sql, "@Initiator2", UserName)
810         sql = Replace(sql, "@UsePrinter3", "Fax")
820         sql = Replace(sql, "@FaxNumber4", FaxNumber)
830         sql = Replace(sql, "@ptime5", Format(Now, "yyyy-MM-dd hh:mm:ss"))
840         sql = Replace(sql, "@Ward8", AddTicks(cmbWard))
850         sql = Replace(sql, "@Clinician9", AddTicks(cmbClinician))
860         sql = Replace(sql, "@GP10", AddTicks(cmbGP))

870         Cnxn(0).Execute sql

            '        sql = "SELECT * FROM PrintPending WHERE " & _
                     '              "Department = '" & Choose(ssTabAll.Tab, "H", "B", "C", "E", "Q", "I", "") & "' " & _
                     '              "AND SampleID = '" & txtSampleID & "' " & _
                     '              "AND COALESCE(FaxNumber,'') <> ''"
            '        Set tb = New Recordset
            '        RecOpenClient 0, tb, sql
            '        If tb.EOF Then
            '            tb.AddNew
            '        End If
            '        tb!SampleID = txtSampleID
            '        tb!Department = Choose(ssTabAll.Tab, "H", "B", "C", "E", "Q", "I", "")
            '        If SysOptRealImm(0) And tb!Department = "I" Then tb!Department = "J"
            '        tb!Initiator = Username
            '        tb!Ward = cmbWard
            '        tb!Clinician = cmbClinician
            '        tb!GP = cmbGP
            '        tb!UsePrinter = "Fax"
            '        tb!FaxNumber = FaxNumber
            '        tb.Update
880         If ssTabAll.Tab = 2 Then
890             If cmdSaveBio.Enabled = True Then SaveBiochemistry True
                '            sql = "UPDATE BIORESULTS SET PRINTED = 0 WHERE SAMPLEID = " & txtSampleID & ""
                '            Cnxn(0).Execute sql
900         ElseIf ssTabAll.Tab = 3 Then
                '            sql = "UPDATE COAGRESULTS SET PRINTED = 0 WHERE SAMPLEID = " & txtSampleID & ""
                '            Cnxn(0).Execute sql
910         ElseIf ssTabAll.Tab = 4 Then
                '            sql = "UPDATE ENDRESULTS SET PRINTED = 0 WHERE SAMPLEID = " & txtSampleID & ""
                '            Cnxn(0).Execute sql
920         ElseIf ssTabAll.Tab = 6 Then
                '            sql = "UPDATE IMMRESULTS SET PRINTED = 0 WHERE SAMPLEID = " & txtSampleID & ""
                '            Cnxn(0).Execute sql
930         End If
940     End If
950 End If




960 Exit Sub

bFAX_Click_Error:

    Dim strES As String
    Dim intEL As Integer

970 intEL = Erl
980 strES = Err.Description
990 LogError "frmEditAll", "bFAX_Click", intEL, strES, sql


End Sub

Private Sub bFilm_Click()


10  On Error GoTo bFilm_Click_Error

20  With frmDifferentials
30      If bFilm.BackColor = vbBlue Then
40          .LoadDiff = True
50      End If
60      .SampleID = txtSampleID
70      .lWBC = tWBC
80      .Show 1
90      .LoadDiff = False
100 End With



110 Exit Sub

bFilm_Click_Error:

    Dim strES As String
    Dim intEL As Integer

120 intEL = Erl
130 strES = Err.Description
140 LogError "frmEditAll", "bFilm_Click", intEL, strES


End Sub

Private Sub bHaemGraphs_Click()

10  On Error GoTo bHaemGraphs_Click_Error

20  frmHaemGraphs.SampleID = txtSampleID
30  frmHaemGraphs.Show 1

40  Exit Sub

bHaemGraphs_Click_Error:

    Dim strES As String
    Dim intEL As Integer

50  intEL = Erl
60  strES = Err.Description
70  LogError "frmEditAll", "bHaemGraphs_Click", intEL, strES


End Sub

Private Sub bHistory_Click()


10  On Error GoTo bHistory_Click_Error

20  pBar = 0

30  Select Case ssTabAll.Tab
    Case 1:
40      With frmFullHaem
50          .lblChart = txtChart
60          .lblName = txtName
70          .lblDoB = txtDoB
80          .lblSex = txtSex
90          .Tn = "0"
100         .Show 1
110     End With
120 Case 2:
130     With frmFullBio
140         .lblChart = txtChart
150         .lblName = txtName
160         .lblDoB = txtDoB
170         .lblAandE = txtAandE
180         .lblSex = txtSex
190         .Show 1
200     End With
210 Case 3:
220     With frmFullCoag
230         .lblChart = txtChart
240         .lblName = txtName
250         .lblDoB = txtDoB
260         .Tn = "0"
270         .Show 1
280     End With
290 Case 4:
300     With frmFullEnd
310         .lblChart = txtChart
320         .lblName = txtName
330         .lblDoB = txtDoB
340         .Tn = "0"
350         .Show 1
360     End With
370 Case 5:
380     With frmFullBga
390         .lblChart = txtChart
400         .lblName = txtName
410         .lblDoB = txtDoB
420         .Tn = "0"
430         .Show 1
440     End With
450 Case 6:
460     With frmFullImm
470         .lblSex = txtSex
480         .lblChart = txtChart
490         .lblName = txtName
500         .lblDoB = txtDoB
510         .Tn = "0"
520         .Show 1
530     End With
540 Case 7:
550     With frmFullExt
560         .lblChart = txtChart
570         .lblName = txtName
580         .lblDoB = txtDoB
590         .Tn = "0"
600         .Show 1
610     End With
620 End Select

630 Exit Sub

bHistory_Click_Error:

    Dim strES As String
    Dim intEL As Integer

640 intEL = Erl
650 strES = Err.Description
660 LogError "frmEditAll", "bHistory_Click", intEL, strES

End Sub

Private Sub bImmRePrint_Click(Index As Integer)

Dim tb As New Recordset
Dim sql As String
Dim Validating As Boolean

On Error GoTo bImmRePrint_Click_Error

pBar = 0

txtSampleID = Format(Val(txtSampleID))

If Not EntriesOK(txtSampleID, txtName, txtSex, cmbWard.Text, cmbGP.Text) Then
    Exit Sub
End If

Validating = cmdDemoVal.Caption = "&Validate"

If Validating Then
    If iMsg("Do you wish to validate demographics !", vbYesNo) = vbNo Then
        Exit Sub
    Else
        ValidateDemographics True
    End If
End If

SaveDemographics

If Index = 0 Then
    LogTimeOfPrinting txtSampleID, "E"
    sql = "UPDATE EndResults " & _
          "Set Printed = '0', Valid = 1 WHERE " & _
          "SampleID = '" & txtSampleID & "' and valid = 1"
    Cnxn(0).Execute sql
    sql = "SELECT * FROM PrintPending WHERE " & _
          "Department = 'E' " & _
          "AND SampleID = '" & txtSampleID & "'"
    Set tb = New Recordset
    RecOpenClient 0, tb, sql
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
ElseIf Index = 1 Then
    LogTimeOfPrinting txtSampleID, "I"
    sql = "UPDATE ImmResults " & _
          "Set Printed = '0', Valid = 1 WHERE " & _
          "SampleID = '" & txtSampleID & "' and valid = 1"
    Cnxn(0).Execute sql

    If SysOptRealImm(0) Then
        sql = "SELECT * FROM PrintPending WHERE " & _
              "Department = 'J' " & _
              "AND SampleID = '" & txtSampleID & "'"
    Else
        sql = "SELECT * FROM PrintPending WHERE " & _
              "Department = 'I' " & _
              "AND SampleID = '" & txtSampleID & "'"
    End If
    Set tb = New Recordset
    RecOpenClient 0, tb, sql
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
ElseIf Index = 2 Then
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

    LogTimeOfPrinting txtSampleID, "G"

    sql = "UPDATE BgaResults " & _
          "Set Printed = '0', Valid = 1 WHERE " & _
          "SampleID = '" & txtSampleID & "'"
    Cnxn(0).Execute sql
    sql = "SELECT * FROM PrintPending WHERE " & _
          "Department = 'Q' " & _
          "AND SampleID = '" & txtSampleID & "'"
    Set tb = New Recordset
    RecOpenClient 0, tb, sql
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
    tb!pTime = Now
    tb.Update

    LoadBloodGas
    
End If

Exit Sub

bImmRePrint_Click_Error:

Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "frmEditAll", "bImmRePrint_Click", intEL, strES, sql

End Sub

Private Sub bOrderTests_Click()


10  On Error GoTo bOrderTests_Click_Error

20  pBar = 0

30  If cmdSaveDemographics.Enabled = True Or cmdSaveInc.Enabled = True Then
40      If iMsg("Save Demographics!", vbYesNo) = vbYes Then
50          cmdSaveDemographics_Click
60      End If
70  End If

80  With frmNewOrder
90      .FromEdit = True
100     .SampleID = Format(Val(txtSampleID))
110     .Show 1
120 End With

130 If SysOptDeptEnd(0) Then LoadOutstandingEnd
140 If SysOptDeptImm(0) Then LoadOutstandingImm
150 If SysOptDeptBio(0) Then LoadOutstandingBio
160 If SysOptDeptCoag(0) Then LoadOutstandingrdCoag
    'If SysOptDeptHaem Then loadoutstandingHaem

170 LoadDemographics




180 Exit Sub

bOrderTests_Click_Error:

    Dim strES As String
    Dim intEL As Integer

190 intEL = Erl
200 strES = Err.Description
210 LogError "frmEditAll", "bOrderTests_Click", intEL, strES


End Sub

Private Sub bremoveduplicates_Click()

    Dim tb As New Recordset
    Dim sql As String
    Dim Y As Long
    Dim Code As String
    Dim Result As String

10  On Error GoTo bremoveduplicates_Click_Error

20  pBar = 0

30  If gBio.Rows < 3 Then Exit Sub

40  For Y = 1 To gBio.Rows - 1
50      Code = CodeForShortName(gBio.TextMatrix(Y, 0))
60      Result = gBio.TextMatrix(Y, 1)
70      sql = "SELECT * from bioresults WHERE " & _
              "sampleid = '" & txtSampleID & "' " & _
              "and code = '" & Code & "'  order by runtime asc"
80      Set tb = New Recordset
90      RecOpenClient 0, tb, sql
100     If tb.recordCount > 1 Then
110         sql = "DELETE from bioresults WHERE sampleid = '" & txtSampleID & "' and code = '" & Code & "' and runtime = '" & Format(tb!RunTime, "dd/MMM/yyyy hh:mm:ss") & "'"
120         Cnxn(0).Execute sql
130     End If
140 Next

150 LoadBiochemistry

160 Exit Sub

bremoveduplicates_Click_Error:

    Dim strES As String
    Dim intEL As Integer

170 intEL = Erl
180 strES = Err.Description
190 LogError "frmEditAll", "bremoveduplicates_Click", intEL, strES, sql

End Sub

Private Sub bReprint_Click()

    Dim tb As New Recordset
    Dim sql As String

10  On Error GoTo bReprint_Click_Error

20  pBar = 0

30  txtSampleID = Format(Val(txtSampleID))

40  If Not EntriesOK(txtSampleID, txtName, txtSex, cmbWard.Text, cmbGP.Text) Then
50      Exit Sub
60  End If

70  If cmdDemoVal.Caption = "&Validate" Then
80      If iMsg("Do you wish to validate demographics !", vbYesNo) = vbNo Then
90          Exit Sub
100     Else
110         ValidateDemographics True
120     End If
130 End If

140 SaveDemographics

150 LogTimeOfPrinting txtSampleID, "B"

160 sql = "UPDATE BioResults SET Printed = '0', Valid = 1 WHERE " & _
          "SampleID = '" & txtSampleID & "' " & _
          "AND code <> '" & SysOptBioCodeForBad(0) & "' " & _
          "AND Valid = 1"
170 Cnxn(0).Execute sql

180 sql = "SELECT * FROM PrintPending WHERE " & _
          "Department = 'B' " & _
          "AND SampleID = '" & txtSampleID & "' " & _
          "AND (FaxNumber = '' OR FaxNumber IS NULL) "
190 Set tb = New Recordset
200 RecOpenClient 0, tb, sql
210 If tb.EOF Then
220     tb.AddNew
230 End If
240 tb!SampleID = txtSampleID
250 tb!Ward = cmbWard
260 tb!Clinician = cmbClinician
270 tb!GP = cmbGP
280 tb!Department = "B"
290 tb!Initiator = UserName
300 tb!UsePrinter = pPrintToPrinter
310 tb!pTime = Now
320 tb.Update

330 LoadBiochemistry


340 Exit Sub

bReprint_Click_Error:

    Dim strES As String
    Dim intEL As Integer

350 intEL = Erl
360 strES = Err.Description
370 LogError "frmEditAll", "bReprint_Click", intEL, strES, sql

End Sub

'Private Sub bRePrintBga_Click()
'
'    Dim tb As New Recordset
'    Dim sql As String
'
'10  On Error GoTo bRePrintBga_Click_Error
'
'20  txtSampleID = Format(Val(txtSampleID))
'30  If Val(txtSampleID) = 0 Then Exit Sub
'
'40  PBar = 0
'
'50  If Trim$(txtSex) = "" Then
'60      If iMsg("Sex not entered." & vbCrLf & "Do you want to enter sex now?", vbQuestion + vbYesNo) = vbYes Then
'70          Exit Sub
'80      End If
'90  End If
'
'100 If Trim$(txtSampleID) = "" Then
'110     iMsg "Must have Lab Number.", vbCritical
'120     Exit Sub
'130 End If
'
'140 If Trim$(cmbWard) = "" Then
'150     iMsg "Must have Ward entry.", vbCritical
'160     Exit Sub
'170 End If
'
'180 If Trim$(cmbWard) = "GP" Then
'190     If Trim$(cmbGP) = "" Then
'200         iMsg "Must have Ward or GP entry.", vbCritical
'210         Exit Sub
'220     End If
'230 End If
'
'240 SaveDemographics
'
'250 LogTimeOfPrinting txtSampleID, "G"
'
'260 sql = "UPDATE BgaResults " & _
'          "Set Printed = '0', Valid = 1 WHERE " & _
'          "SampleID = '" & txtSampleID & "'"
'270 Cnxn(0).Execute sql
'280 sql = "SELECT * FROM PrintPending WHERE " & _
'          "Department = 'Q' " & _
'          "AND SampleID = '" & txtSampleID & "'"
'290 Set tb = New Recordset
'300 RecOpenClient 0, tb, sql
'310 If tb.EOF Then
'320     tb.AddNew
'330 End If
'340 tb!SampleID = txtSampleID
'350 tb!Ward = cmbWard
'360 tb!Clinician = cmbClinician
'370 tb!GP = cmbGP
'380 tb!Department = "Q"
'390 tb!Initiator = UserName
'400 tb!UsePrinter = pPrintToPrinter
'410 tb!pTime = Now
'420 tb.Update
'
'430 LoadBloodGas
'
'440 Exit Sub
'
'bRePrintBga_Click_Error:
'
'    Dim strES As String
'    Dim intEL As Integer
'
'450 intEL = Erl
'460 strES = Err.Description
'470 LogError "frmEditAll", "bRePrintBga_Click", intEL, strES, sql
'
'End Sub

Private Sub bsearch_Click()

10  On Error GoTo bsearch_Click_Error

20  pBar = 0

30  With frmPatHistoryNew
40      .oHD(1) = True
50      .oFor(0) = True
60      .txtName = txtName
70      If cmdDemoVal.Caption = "VALID" Then .mDemoVal = True Else .mDemoVal = False
80      .FromEdit = True
90      .EditScreen = Me
100     .bsearch = True
110     If .g.TextMatrix(1, 13) <> "" Then
120         .Show 1
130     Else
140         FlashNoPrevious lNoPrevious
150     End If

160 End With

170 Exit Sub

bsearch_Click_Error:

    Dim strES As String
    Dim intEL As Integer

180 intEL = Erl
190 strES = Err.Description
200 LogError "frmEditAll", "bsearch_Click", intEL, strES

End Sub

Private Sub bValidateBio_Click()

10  On Error GoTo bValidateBio_Click_Error

20  pBar = 0

30  txtSampleID = Format(Val(txtSampleID))
40  If Val(txtSampleID) = 0 Then Exit Sub

50  If bValidateBio.Caption = "VALID" Then
60      If UCase(iBOX("Unvalidate ! Enter Password", , , True)) = UCase(UserPass) Then
70          SaveBiochemistry False, True
80          SaveComments
90          Me.Refresh
100     End If
110 Else
120     If cmdDemoVal.Caption = "&Validate" Then
130         If iMsg("Do you wish to validate demographics ?", vbYesNo) = vbNo Then
140             Exit Sub
150         Else
160             ValidateDemographics True
170         End If
180     End If
        'If txtDoB = "" Then iMsg "No Dob. Adult Age 25 used for Normal Ranges!"
190     SaveBiochemistry True
200     SaveComments
210     UPDATEMRU txtSampleID, cMRU
220     Frame2.Enabled = False
230     lRandom.Enabled = False
240     txtBioComment.Locked = True
250     Me.Refresh
260     If SysOptBioValFore(0) = True Then
270         txtSampleID = Format$(Val(txtSampleID) + 1)
280     End If
290 End If

300 LoadAllDetails

310 Exit Sub

bValidateBio_Click_Error:

    Dim strES As String
    Dim intEL As Integer

320 intEL = Erl
330 strES = Err.Description
340 LogError "frmEditAll", "bValidateBio_Click", intEL, strES

End Sub

Private Sub bValidateHaem_Click()

10  On Error GoTo bValidateHaem_Click_Error

20  pBar = 0

30  txtSampleID = Format(Val(txtSampleID))
40  If Val(txtSampleID) = 0 Then Exit Sub

50  If bValidateHaem.Caption = "VALID" Then
60      If UCase(iBOX("Unvalidate ! Enter Password", , , True)) = UserPass Then
70          SaveHaematology False
80          SaveComments
90          Panel3D4.Enabled = True
100         Panel3D5.Enabled = True
110         Panel3D6.Enabled = True
            'Haemlock
120         Panel3D7.Enabled = True
130         txtHaemComment.Enabled = True
140         txtHaemComment.Locked = False
150         bValidateHaem.Caption = "&Validate"
160         lblHaemValid.Visible = False
170         LoadHaematology
180         Me.Refresh
190     Else
200         Exit Sub
210     End If
220 Else
230     If cmdDemoVal.Caption = "&Validate" Then
240         If iMsg("Do you wish to validate demographics?", vbQuestion + vbYesNo) = vbNo Then
250             Exit Sub
260         Else
270             ValidateDemographics True
280         End If
290     End If

        'If Trim(txtDoB) = "" Then iMsg "No Dob. Adult Age 25 used for Normal Ranges!"

300     SaveHaematology 1
310     SaveComments
320     UPDATEMRU txtSampleID, cMRU
330     Panel3D4.Enabled = False
340     Panel3D5.Enabled = False
350     Panel3D6.Enabled = False
        'Haemlock
360     Panel3D7.Enabled = False
370     If SysOptCommVal(0) Then txtHaemComment.Enabled = False
380     txtSampleID = Format$(Val(txtSampleID) + 1)
390     LoadAllDetails
400     Me.Refresh
410 End If

420 Exit Sub

bValidateHaem_Click_Error:

    Dim strES As String
    Dim intEL As Integer

430 intEL = Erl
440 strES = Err.Description
450 LogError "frmEditAll", "bValidateHaem_Click", intEL, strES

End Sub

Private Sub bValidateImm_Click(Index As Integer)

10  On Error GoTo bValidateImm_Click_Error

20  pBar = 0

30  txtSampleID = Format(Val(txtSampleID))
40  If Val(txtSampleID) = 0 Then Exit Sub

50  If Index = 0 Then

60      If bValidateImm(0).Caption = "VALID" Then
70          If UCase(iBOX("Unvalidate ! Enter Password", , , True)) = UCase(UserPass) Then
80              SaveEndocrinology False, True
90              SaveComments
100             Me.Refresh
110         End If
120     Else
130         If SysOptDemoVal(0) And cmdDemoVal.Caption = "&Validate" Then
140             If iMsg("Do you wish to validate demographics !", vbYesNo) = vbNo Then
150                 Exit Sub
160             Else
170                 ValidateDemographics True
180             End If
190         End If
            'If Trim(txtDoB) = "" Then iMsg "No Dob. Adult Age 25 used for Normal Ranges!"
200         SaveEndocrinology True
210         SaveComments
220         UPDATEMRU txtSampleID, cMRU
230         Frame12(0).Enabled = False
240         lImmRan(0).Enabled = False
250         Me.Refresh
260         txtSampleID = Format$(Val(txtSampleID) + 1)
270     End If

280 ElseIf Index = 1 Then

290     If bValidateImm(1).Caption = "VALID" Then
300         If UCase(iBOX("Unvalidate ! Enter Password", , , True)) = UCase(UserPass) Then
310             SaveImmunology False, True
320             SaveComments
330         End If
340     Else
350         If SysOptDemoVal(0) And cmdDemoVal.Caption = "&Validate" Then
360             If iMsg("Do you wish to validate demographics?", vbQuestion + vbYesNo) = vbNo Then
370                 Exit Sub
380             Else
390                 ValidateDemographics True
400             End If
410         End If
            '        If Trim(txtDoB) = "" Then
            '            iMsg "No Date of Birth Specified." & vbCrLf & "Adult Age 25 used for Normal Ranges!", vbInformation
            '        End If
420         SaveImmunology True
430         SaveComments
440         UPDATEMRU txtSampleID, cMRU
450         Frame12(1).Enabled = False
460         lImmRan(1).Enabled = False
470     End If

480 ElseIf Index = 2 Then
490     If bValidateImm(2).Caption = "VALID" Then
500         If UCase(iBOX("Unvalidate ! Enter Password", , , True)) = UCase(UserPass) Then
510             SaveExtern False, True
520         End If
530     Else
540         If SysOptDemoVal(0) And cmdDemoVal.Caption = "&Validate" Then
550             If iMsg("Do you wish to validate demographics?", vbQuestion + vbYesNo) = vbNo Then
560                 Exit Sub
570             Else
580                 ValidateDemographics True
590             End If
600         End If
            '            If Trim(txtDoB) = "" Then
            '                iMsg "No Date of Birth Specified." & vbCrLf & "Adult Age 25 used for Normal Ranges!", vbInformation
            '            End If
610         SaveExtern True
620         UPDATEMRU txtSampleID, cMRU
630     End If

640 End If

650 LoadAllDetails

660 Exit Sub

bValidateImm_Click_Error:

    Dim strES As String
    Dim intEL As Integer

670 intEL = Erl
680 strES = Err.Description
690 LogError "frmEditAll", "bValidateImm_Click", intEL, strES

End Sub

Private Sub bViewBB_Click()

10  On Error GoTo bViewBB_Click_Error

20  pBar = 0

30  If Trim$(txtChart) <> "" Then
40      frmViewBB.lChart = txtChart
50      frmViewBB.Show 1
60  End If

70  Exit Sub

bViewBB_Click_Error:

    Dim strES As String
    Dim intEL As Integer

80  intEL = Erl
90  strES = Err.Description
100 LogError "frmEditAll", "bViewBB_Click", intEL, strES


End Sub

Private Sub bViewBgaRepeat_Click()

10  On Error GoTo bViewBgaRepeat_Click_Error

20  pBar = 0

30  frmViewBgaRepeat.Show 1

40  Exit Sub

bViewBgaRepeat_Click_Error:

    Dim strES As String
    Dim intEL As Integer

50  intEL = Erl
60  strES = Err.Description
70  LogError "frmEditAll", "bViewBgaRepeat_Click", intEL, strES


End Sub

Private Sub bViewBioRepeat_Click()

10        On Error GoTo bViewBioRepeat_Click_Error

20        pBar = 0

30        frmViewBioRepeat.Show 1
40        LoadBiochemistry

50        Exit Sub

bViewBioRepeat_Click_Error:

          Dim strES As String
          Dim intEL As Integer

60        intEL = Erl
70        strES = Err.Description
80        LogError "frmEditAll", "bViewBioRepeat_Click", intEL, strES


End Sub

Private Sub bViewCoagRepeat_Click()

10  On Error GoTo bViewCoagRepeat_Click_Error

20  pBar = 0

30  With frmCoagRepeats
40      .EditForm = Me
50      .SampleID = txtSampleID
60      .Show 1
70  End With

80  Exit Sub

bViewCoagRepeat_Click_Error:

    Dim strES As String
    Dim intEL As Integer

90  intEL = Erl
100 strES = Err.Description
110 LogError "frmEditAll", "bViewCoagRepeat_Click", intEL, strES


End Sub

Private Sub bViewHaemRepeat_Click()

10  On Error GoTo bViewHaemRepeat_Click_Error

20  pBar = 0

30  With frmViewHaemRep
40      .EditForm = Me
50      .lSampleID = txtSampleID
60      .lName = txtName
70      .Show 1
80  End With

90  LoadHaematology

100 Exit Sub

bViewHaemRepeat_Click_Error:

    Dim strES As String
    Dim intEL As Integer

110 intEL = Erl
120 strES = Err.Description
130 LogError "frmEditAll", "bViewHaemRepeat_Click", intEL, strES


End Sub

Private Sub bViewImmRepeat_Click(Index As Integer)

10  On Error GoTo bViewImmRepeat_Click_Error

20  pBar = 0

30  If Index = 0 Then
40      frmViewEndRepeat.Show 1
50  Else
60      frmViewImmRepeat.Show 1
70  End If

80  Exit Sub

bViewImmRepeat_Click_Error:

    Dim strES As String
    Dim intEL As Integer

90  intEL = Erl
100 strES = Err.Description
110 LogError "frmEditAll", "bViewImmRepeat_Click", intEL, strES


End Sub

Private Sub cAdd_Click()

    Dim n As Integer
    Dim Code As String
    Dim EGFRCode As String
    Dim EGFRCodeP As String
    Dim EGFROK As Boolean


    On Error GoTo cAdd_Click_Error

    pBar = 0

    Dim SampleType As String
    Dim tb As New Recordset
    Dim sql As String

    cUnits.Enabled = True

    'SampleType = ListCodeFor("ST", cISampleType(3))
    '
    'If SampleType = "T" Then
    '    cAdd = ""
    '    tnewvalue = ""
    '    cUnits = ""
    '    frmEditToxicology.Show 1
    '    LoadBiochemistry
    '    Exit Sub
    'End If

    sql = "SELECT * FROM BioTestDefinitions WHERE " & _
          "Code = '" & CodeForShortName(cAdd) & "'"
    Set tb = New Recordset
    RecOpenServer 0, tb, sql
    If Not tb.EOF Then
        cUnits = tb!Units
    Else
        cUnits = ""
    End If

    cUnits.Enabled = False



    EGFRCode = GetOptionSetting("BioCodeForEGFR", "EGFR")
    EGFRCodeP = GetOptionSetting("BioCodeForEGFRPlasma", "EGFRP")
    If EGFRCode = CodeForShortName(cAdd) Or EGFRCodeP = CodeForShortName(cAdd) Then
        EGFROK = False
        For n = 1 To gBio.Rows - 1
            Code = CodeForShortName(gBio.TextMatrix(n, 0))
            If Code = "235" Or Code = "235P" Then
                If IsNumeric(gBio.TextMatrix(n, 1)) Then
                    tnewvalue = calculateEGFR(gBio.TextMatrix(n, 1))

                Else
                    iMsg "Creatinine Result not numeric." & vbCrLf & "Can't add eGFR.", vbInformation
                    cAdd = ""
                    cUnits = ""
                    tnewvalue = ""
                End If
                EGFROK = True
                Exit For
            End If
        Next
        If Not EGFROK Then
            For n = 1 To gBio.Rows - 1
                Code = CodeForShortName(gBio.TextMatrix(n, 0))
                If Code = "234" Or Code = "741" Or Code = "234P" Or Code = "741P" Then
                    If IsNumeric(gBio.TextMatrix(n, 1)) Then
                        tnewvalue = calculateEGFR(gBio.TextMatrix(n, 1))
                    Else
                        iMsg "Creatinine Result not numeric." & vbCrLf & "Can't add eGFR.", vbInformation
                        cAdd = ""
                        cUnits = ""
                        tnewvalue = ""
                    End If
                    EGFROK = True
                    Exit For
                End If
            Next
        End If
        If Not EGFROK Then
            iMsg "No Creatinine Result." & vbCrLf & "Can't add eGFR.", vbInformation
            cAdd = ""
            cUnits = ""
            tnewvalue = ""
        End If
    End If

    Exit Sub

cAdd_Click_Error:

    Dim strES As String
    Dim intEL As Integer

    intEL = Erl
    strES = Err.Description
    LogError "frmEditAll", "cAdd_Click", intEL, strES, sql

End Sub

Private Sub cAdd_KeyPress(KeyAscii As Integer)

10  On Error GoTo cAdd_KeyPress_Error

20  KeyAscii = AutoComplete(cAdd, KeyAscii, False)
    'KeyAscii = 0

30  Exit Sub

cAdd_KeyPress_Error:

    Dim strES As String
    Dim intEL As Integer

40  intEL = Erl
50  strES = Err.Description
60  LogError "frmEditAll", "cAdd_KeyPress", intEL, strES


End Sub

Private Function CalculateGPCR(ByVal UrCreat As String, ByVal UrProptien As String) As String

10  On Error GoTo CalculateGPCR_Error

20  CalculateGPCR = Round((UrProptien * 1000) / UrCreat, 2)

30  Exit Function

CalculateGPCR_Error:

    Dim strES As String
    Dim intEL As Integer

40  intEL = Erl
50  strES = Err.Description
60  LogError "frmEditAll", "CalculateGPCR", intEL, strES

End Function

Private Function calculateEGFR(ByVal Creat As String) As String

          Dim s As Long
          Dim pAge As Long

10    On Error GoTo calculateEGFR_Error

20    calculateEGFR = Creat

30    If txtDoB = "" Then Exit Function

40    pAge = CalcpAge(txtDoB)
      ' -------------New Formula 19/11/2019------------------------------
50    s = (175 * ((Val(Creat) / 88.4) ^ (-1.154))) * Val(pAge) ^ (-0.203)

      '-----------Old Formual--------------------------------------
      's = (32788 * (Val(Creat) ^ (-1.154))) * Val(pAge) ^ (-0.203)

60    If Left(lSex, 1) = "F" Then s = s * 0.742

70    If s > 60 Then
80      calculateEGFR = ">60"
90    Else
100     calculateEGFR = s
110   End If

120   Exit Function

calculateEGFR_Error:

          Dim strES As String
          Dim intEL As Integer

130   intEL = Erl
140   strES = Err.Description
150   LogError "frmEditAll", "calculateEGFR", intEL, strES

End Function

Private Sub cAdd_LostFocus()
10  On Error GoTo cAdd_LostFocus_Error

20  cAdd.Text = QueryCombo(cAdd)

30  Exit Sub

cAdd_LostFocus_Error:

    Dim strES As String
    Dim intEL As Integer

40  intEL = Erl
50  strES = Err.Description
60  LogError "frmEditAll", "cAdd_LostFocus", intEL, strES

End Sub

Private Sub cASot_Click()

10  On Error GoTo cASot_Click_Error

20  If cASot = 0 Then
30      If Trim$(tASOt) = "?" Then
40          tASOt = ""
50      ElseIf Trim$(tASOt) <> "" Then
60          cASot = 1
70      End If
80  Else
90      If Trim$(tASOt) = "" Then
100         tASOt = "?"
110     End If
120 End If

130 cmdSaveHaem.Enabled = True: cmdHSaveH.Enabled = True

140 Exit Sub

cASot_Click_Error:

    Dim strES As String
    Dim intEL As Integer

150 intEL = Erl
160 strES = Err.Description
170 LogError "frmEditAll", "cASot_Click", intEL, strES


End Sub

Private Sub cCat_Change(Index As Integer)

10  On Error GoTo cCat_Change_Error

20  If Index = 0 Then
30      cmdSaveDemographics.Enabled = True
40      cmdSaveInc.Enabled = True
50      cCat(1) = cCat(0)
60  Else
70      cmdSaveImm(0).Enabled = True
80      cCat(0) = cCat(1)
90  End If

100 Exit Sub

cCat_Change_Error:

    Dim strES As String
    Dim intEL As Integer

110 intEL = Erl
120 strES = Err.Description
130 LogError "frmEditAll", "cCat_Change", intEL, strES

End Sub

Private Sub cCat_Click(Index As Integer)

    Dim sql As String

10  On Error GoTo cCat_Click_Error

20  If Index = 0 Then
30      cmdSaveDemographics.Enabled = True
40      cmdSaveInc.Enabled = True
50      cCat(1) = cCat(0)
60  Else
70      If EndLoaded = True Then
80          sql = "UPDATE demographics set category = '" & cCat(1) & "' WHERE sampleid = " & txtSampleID & ""
90          Cnxn(0).Execute sql
100         cCat(0) = cCat(1)

110     End If
120 End If

130 Exit Sub

cCat_Click_Error:

    Dim strES As String
    Dim intEL As Integer

140 intEL = Erl
150 strES = Err.Description
160 LogError "frmEditAll", "cCat_Click", intEL, strES

End Sub

Private Sub cClDetails_Click()

10  On Error GoTo cClDetails_Click_Error


20  cmdSaveDemographics.Enabled = True
30  cmdSaveInc.Enabled = True

40  Exit Sub

cClDetails_Click_Error:

    Dim strES As String
    Dim intEL As Integer

50  intEL = Erl
60  strES = Err.Description
70  LogError "frmEditAll", "cClDetails_Click", intEL, strES

End Sub

Private Sub cClDetails_KeyPress(KeyAscii As Integer)

10  On Error GoTo cClDetails_KeyPress_Error

    '20        KeyAscii = AutoComplete(cClDetails, KeyAscii, False)

20  Exit Sub

cClDetails_KeyPress_Error:

    Dim strES As String
    Dim intEL As Integer

30  intEL = Erl
40  strES = Err.Description
50  LogError "frmEditAll", "cClDetails_KeyPress", intEL, strES

End Sub

Private Sub cClDetails_LostFocus()

10  On Error GoTo cClDetails_LostFocus_Error

20  pBar = 0

30  If Trim$(cClDetails) = "" Then Exit Sub

40  If ListText("CD", cClDetails) <> "" Then
50      cClDetails = ListText("CD", cClDetails)
60  End If

    '70        cClDetails.Text = QueryCombo(cClDetails)

70  Exit Sub

cClDetails_LostFocus_Error:

    Dim strES As String
    Dim intEL As Integer

80  intEL = Erl
90  strES = Err.Description
100 LogError "frmEditAll", "cClDetails_LostFocus", intEL, strES

End Sub

Private Sub cESR_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

10  On Error GoTo cESR_MouseUp_Error

20  pBar = 0

30  If cESR = 0 Then
40      If Trim$(tESR) = "?" Then
50          tESR = ""
60      ElseIf Trim$(tESR) <> "" Then
70          cESR = 1
80      End If
90  Else
100     If Trim$(tESR) = "" Then
110         tESR = "?"
120     End If
130 End If

140 cmdSaveHaem.Enabled = True: cmdHSaveH.Enabled = True

150 Exit Sub

cESR_MouseUp_Error:

    Dim strES As String
    Dim intEL As Integer

160 intEL = Erl
170 strES = Err.Description
180 LogError "frmEditAll", "cESR_MouseUp", intEL, strES

End Sub

Private Sub cFilm_Click()

10  On Error GoTo cFilm_Click_Error

20  cmdSaveHaem.Enabled = True: cmdHSaveH.Enabled = True

30  Exit Sub

cFilm_Click_Error:

    Dim strES As String
    Dim intEL As Integer

40  intEL = Erl
50  strES = Err.Description
60  LogError "frmEditAll", "cFilm_Click", intEL, strES

End Sub

Private Sub CheckAssGlucose(ByVal CurrentBRs As BIEResults)

    Dim tb As New Recordset
    Dim sql As String

10  On Error GoTo CheckAssGlucose_Error

20  If CurrentBRs.Count = 1 Then
30      If CurrentBRs(1).Code = SysOptBioCodeForGlucose(0) Then
            'check prev or next for general
40          sql = "SELECT DISTINCT D.SampleID " & _
                  "FROM Demographics D " & _
                  "WHERE D.SampleID IN " & _
                "  (  SELECT SampleID FROM BioResults WHERE " & _
                "       (SampleID = '" & Val(txtSampleID) - 1 & "' " & _
                "        OR SampleID = '" & Val(txtSampleID) + 1 & "') " & _
                "     AND Code <> '" & SysOptBioCodeForGlucose(0) & "'  ) " & _
                  "AND D.PatName = '" & AddTicks(txtName) & "' " & _
                  "AND (D.SampleID = '" & Val(txtSampleID) - 1 & "' " & _
                  "OR D.SampleID = '" & Val(txtSampleID) + 1 & "')"
50          Set tb = New Recordset
60          RecOpenServer 0, tb, sql
70          If Not tb.EOF Then
80              lblAss = "Associated Results " & tb!SampleID
90              lblAss.Visible = True
100         End If
110     Else
120         sql = "SELECT distinct D.SampleID " & _
                  "from Demographics as D " & _
                  "WHERE D.sampleid in " & _
                "  (  SELECT SampleID from BioResults WHERE " & _
                "     (SampleID = '" & Val(txtSampleID) - 1 & "' or SampleID = '" & Val(txtSampleID) + 1 & "') " & _
                "     and Code = '" & SysOptBioCodeForGlucose(0) & "'  ) " & _
                  "and D.PatName = '" & AddTicks(txtName) & "' " & _
                  "and (D.SampleID = '" & Val(txtSampleID) - 1 & "' or SampleID = '" & Val(txtSampleID) + 1 & "')"
130         Set tb = New Recordset
140         RecOpenServer 0, tb, sql
150         If Not tb.EOF Then
160             lblAss = "Associated Glucose " & tb!SampleID
170             lblAss.Visible = True
180         End If
190     End If
200 Else
210     sql = "SELECT distinct D.SampleID " & _
              "from Demographics as D " & _
              "WHERE D.sampleid in " & _
            "  (  SELECT SampleID from BioResults WHERE " & _
            "     (SampleID = '" & Val(txtSampleID) - 1 & "' or SampleID = '" & Val(txtSampleID) + 1 & "') " & _
            "     and Code = '" & SysOptBioCodeForGlucose(0) & "'  ) " & _
              "and D.PatName = '" & AddTicks(txtName) & "' " & _
              "and (D.SampleID = '" & Val(txtSampleID) - 1 & "' or SampleID = '" & Val(txtSampleID) + 1 & "')"
220     Set tb = New Recordset
230     RecOpenServer 0, tb, sql
240     If Not tb.EOF Then
250         lblAss = "Associated Glucose " & tb!SampleID
260         lblAss.Visible = True
270     End If
280 End If

290 Exit Sub

CheckAssGlucose_Error:

    Dim strES As String
    Dim intEL As Integer

300 intEL = Erl
310 strES = Err.Description
320 LogError "frmEditAll", "CheckAssGlucose", intEL, strES, sql

End Sub

Private Sub CheckCalcEPSA(ByVal Ims As BIEResults)

    Dim Im As BIEResult
    Dim FPS As Single
    Dim FPSTime As String
    Dim FPSDate As String
    Dim PSA As Single
    Dim Ratio As Single
    Dim Code As String

10  On Error GoTo CheckCalcEPSA_Error

20  If Ims Is Nothing Then Exit Sub

30  FPS = 0
40  PSA = 0
50  Ratio = 0

60  For Each Im In Ims
70      Code = UCase$(Trim$(Im.Code))
80      If Code = "FPS" Then
90          FPS = Val(Im.Result)
100         FPSDate = Im.Rundate
110         FPSTime = Im.RunTime
120     ElseIf Code = "PSA" Then
130         PSA = Val(Im.Result)
140     ElseIf Code = "FPR" Then
150         Ratio = Val(Im.Result)
160     End If
170 Next

180 If (FPS * PSA) <> 0 And Ratio = 0 Then
190     Ratio = FPS / PSA
200     Set Im = New BIEResult
210     Im.SampleID = txtSampleID
220     Im.Code = "FPR"
230     Im.Rundate = FPSDate
240     Im.RunTime = FPSTime
250     Im.Result = Format$(Ratio, "#0.00")
260     Im.Units = ""
270     Im.Printed = 0
280     Im.Valid = 0
290     Ims.Add Im
300     Ims.Save "End", Ims
310 End If

320 Exit Sub

CheckCalcEPSA_Error:

    Dim strES As String
    Dim intEL As Integer

330 intEL = Erl
340 strES = Err.Description
350 LogError "frmEditAll", "CheckCalcEPSA", intEL, strES

End Sub

Private Sub CheckCalcIPSA(ByVal Ims As BIEResults)

    Dim Im As BIEResult
    Dim FPS As Single
    Dim FPSTime As String
    Dim FPSDate As String
    Dim PSA As Single
    Dim Ratio As Single
    Dim Code As String

10  On Error GoTo CheckCalcIPSA_Error

20  If Ims Is Nothing Then Exit Sub

30  FPS = 0
40  PSA = 0
50  Ratio = 0

60  For Each Im In Ims
70      Code = UCase$(Trim$(Im.Code))
80      If Code = "FPS" Then
90          FPS = Val(Im.Result)
100         FPSDate = Im.Rundate
110         FPSTime = Im.RunTime
120     ElseIf Code = "PSA" Then
130         PSA = Val(Im.Result)
140     ElseIf Code = "FPR" Then
150         Ratio = Val(Im.Result)
160     End If
170 Next

180 If (FPS * PSA) <> 0 And Ratio = 0 Then
190     Ratio = FPS / PSA
200     Set Im = New BIEResult
210     Im.SampleID = txtSampleID
220     Im.Code = "FPR"
230     Im.Rundate = FPSDate
240     Im.RunTime = FPSTime
250     Im.Result = Format$(Ratio, "#0.00")
260     Im.Units = ""
270     Im.Printed = 0
280     Im.Valid = 0
290     Ims.Add Im
300     Ims.Save "Imm", Ims
310 End If

320 Exit Sub

CheckCalcIPSA_Error:

    Dim strES As String
    Dim intEL As Integer

330 intEL = Erl
340 strES = Err.Description
350 LogError "frmEditAll", "CheckCalcIPSA", intEL, strES

End Sub

Private Sub CheckCalcPSA(ByVal BRs As BIEResults)

    Dim br As BIEResult
    Dim FPS As Single
    Dim FPSTime As String
    Dim FPSDate As String
    Dim PSA As Single
    Dim Ratio As Single
    Dim Code As String

10  On Error GoTo CheckCalcPSA_Error

20  If BRs Is Nothing Then Exit Sub

30  FPS = 0
40  PSA = 0
50  Ratio = 0

60  For Each br In BRs
70      Code = UCase$(Trim$(br.Code))
80      If Code = "FPS" Then
90          FPS = Val(br.Result)
100         FPSDate = br.Rundate
110         FPSTime = br.RunTime
120     ElseIf Code = "PSA" Then
130         PSA = Val(br.Result)
140     ElseIf Code = "FPR" Then
150         Ratio = Val(br.Result)
160     End If
170 Next

180 If (FPS * PSA) <> 0 And Ratio = 0 Then
190     Ratio = FPS / PSA
200     Set br = New BIEResult
210     br.SampleID = txtSampleID
220     br.Code = "FPR"
230     br.Rundate = FPSDate
240     br.RunTime = FPSTime
250     br.Result = Format$(Ratio, "#0.00")
260     br.Units = ""
270     br.Printed = 0
280     br.Valid = 0
290     BRs.Add br
300     BRs.Save "Bio", BRs
310 End If

320 Exit Sub

CheckCalcPSA_Error:

    Dim strES As String
    Dim intEL As Integer

330 intEL = Erl
340 strES = Err.Description
350 LogError "frmEditAll", "CheckCalcPSA", intEL, strES

End Sub

Private Function CheckGPCR(ByVal BRs As BIEResults) As Boolean
'returns True if GPCR added

    Dim br As BIEResult
    Dim Code As String
    Dim GPCR As String
    Dim UrCreat As String
    Dim UrProtein As String
    Dim Rundate As String
    Dim RunTime As String
    Dim bnew As BIEResult
    Dim sql As String
    Dim tb As Recordset
    Dim GPCRCode As String
    Dim SampleType As String

10  On Error GoTo CheckGPCR_Error

20  CheckGPCR = False

30  If BRs Is Nothing Then Exit Function

40  GPCRCode = GetOptionSetting("BioCodeForGPCR", "972")
50  For Each br In BRs

60      Code = UCase$(Trim$(br.Code))
70      If Code = GPCRCode Then    '"972"
80          CheckGPCR = False
90          Exit Function
100     End If
110     If Code = "123" Then
120         UrProtein = br.Result
130         If IsDate(br.Rundate) Then
140             Rundate = br.Rundate
150         Else
160             Rundate = Format$(br.RunTime, "dd/mmm/yyyy")
170         End If
180         RunTime = br.RunTime
190         SampleType = br.SampleType
200     End If
210     If Code = "691" Then
220         UrCreat = br.Result
230     End If
240 Next
250 If UrProtein = "" Or UrCreat = "" Then Exit Function

260 sql = "SELECT Count(*) AS Cnt FROM BioRequests WHERE Code = '" & GPCRCode & "' AND SampleID = " & BRs(1).SampleID
270 Set tb = New Recordset
280 RecOpenServer 0, tb, sql
290 If tb!Cnt = 0 Then
300     Exit Function
310 End If

320 GPCR = CalculateGPCR(UrCreat, UrProtein)
330 sql = "SELECT * FROM BioResults WHERE " & _
          "SampleID = '" & txtSampleID & "' " & _
          "AND Code = '" & GPCRCode & "'"
340 Set tb = New Recordset
350 RecOpenClient 0, tb, sql
360 If tb.EOF Then
370     tb.AddNew
380 End If
390 tb!SampleID = txtSampleID
400 tb!Rundate = Rundate
410 tb!RunTime = RunTime
420 tb!Code = GPCRCode    '5555
430 tb!Result = GPCR
440 tb!Units = "mg/mmol"
450 tb!Printed = 0
460 tb!Valid = 0
470 tb!Faxed = 0
480 tb!Analyser = ""
490 tb!SampleType = SampleType
500 tb.Update

510 Set bnew = New BIEResult
520 bnew.SampleID = txtSampleID
530 bnew.Code = GPCRCode    '"5555"
540 bnew.Rundate = Rundate
550 bnew.RunTime = RunTime
560 bnew.Result = GPCR
570 bnew.Units = "mg/mmol"
580 bnew.Printed = 0
590 bnew.Valid = 0
600 bnew.SampleType = SampleType
610 bnew.LongName = "Gest. PCR"
620 BRs.Add bnew

630 CheckGPCR = True


640 sql = "DELETE FROM BioRequests  WHERE Code = '" & GPCRCode & "' AND SampleID = " & BRs(1).SampleID
650 Cnxn(0).Execute sql

660 Exit Function

CheckGPCR_Error:

    Dim strES As String
    Dim intEL As Integer

670 intEL = Erl
680 strES = Err.Description
690 LogError "frmEditAll", "CheckGPCR", intEL, strES, sql


End Function

Private Function CheckEGFR(ByVal BRs As BIEResults) As Boolean
'returns True if eGFR added

    Dim br As BIEResult
    Dim Code As String
    Dim eGFR As String
    Dim Rundate As String
    Dim RunTime As String
    Dim bnew As BIEResult
    Dim sql As String
    Dim tb As Recordset
    Dim EGFRCode As String
    Dim EGFRUsername As String

10  On Error GoTo CheckEGFR_Error

20  CheckEGFR = False

30  If BRs Is Nothing Then Exit Function

40  EGFRCode = GetOptionSetting("BioCodeForEGFR", "EGFR")
50  EGFRUsername = GetOptionSetting("BioUsernameForEGFR", "")
60  For Each br In BRs

70      Code = UCase$(Trim$(br.Code))
80      If Code = EGFRCode Then    '"5555"
90          Exit Function
100     End If
110 Next
120 sql = "SELECT Count(*) AS Cnt FROM BioRequests WHERE Code = '" & EGFRCode & "' AND SampleID = " & BRs(1).SampleID
130 Set tb = New Recordset
140 RecOpenServer 0, tb, sql
150 If tb!Cnt = 0 Then
160     Exit Function
170 End If


180 For Each br In BRs
190     Code = UCase$(Trim$(br.Code))
200     If Code = "234" Or Code = "741" Then
210         eGFR = calculateEGFR(br.Result)
220         If Val(eGFR) <> Val(br.Result) Then
230             If IsDate(br.Rundate) Then
240                 Rundate = br.Rundate
250             Else
260                 Rundate = Format$(br.RunTime, "dd/mmm/yyyy")
270             End If
280             RunTime = br.RunTime
290             sql = "SELECT * FROM BioResults WHERE " & _
                      "SampleID = '" & txtSampleID & "' " & _
                      "AND Code = '" & EGFRCode & "'"
300             Set tb = New Recordset
310             RecOpenClient 0, tb, sql
320             If tb.EOF Then
330                 tb.AddNew
340             End If
350             tb!SampleID = txtSampleID
360             tb!Rundate = Rundate
370             tb!RunTime = RunTime
380             tb!Code = EGFRCode    '5555
390             tb!Result = eGFR
400             tb!Units = "ml/min/1.73m2"
410             tb!Printed = 0
420             tb!Valid = IIf(EGFRUsername = "", 0, 1)
430             tb!Operator = EGFRUsername
440             tb!Faxed = 0
450             tb!Analyser = ""
460             tb!SampleType = br.SampleType
470             tb.Update

480             Set bnew = New BIEResult
490             bnew.SampleID = txtSampleID
500             bnew.Code = EGFRCode    '"5555"
510             bnew.Rundate = Rundate
520             bnew.RunTime = RunTime
530             bnew.Result = eGFR
540             bnew.Units = "ml/min/1.73m2"
550             bnew.Printed = 0
560             tb!Operator = EGFRUsername
570             bnew.Valid = IIf(EGFRUsername = "", 0, 1)
580             bnew.SampleType = br.SampleType
590             bnew.LongName = "eGFR"
600             BRs.Add bnew

610             CheckEGFR = True
620             Exit For
630         End If
640     End If
650 Next

660 sql = "DELETE FROM BioRequests  WHERE Code = '" & EGFRCode & "' AND SampleID = " & BRs(1).SampleID
670 Cnxn(0).Execute sql

680 Exit Function

CheckEGFR_Error:

    Dim strES As String
    Dim intEL As Integer

690 intEL = Erl
700 strES = Err.Description
710 LogError "frmEditAll", "CheckEGFR", intEL, strES, sql


End Function


Private Sub CheckCC()

    Dim sql As String
    Dim tb As Recordset

10  On Error GoTo CheckCC_Error

20  cmdCopyTo.Caption = "cc"
30  cmdCopyTo.Font.Bold = False
40  cmdCopyTo.BackColor = &H8000000F

50  If Trim$(txtSampleID) = "" Then Exit Sub

60  sql = "Select * from SendCopyTo where " & _
          "SampleID = '" & Val(txtSampleID) & "'"
70  Set tb = New Recordset
80  RecOpenServer 0, tb, sql
90  If Not tb.EOF Then
100     cmdCopyTo.Caption = "++ cc ++"
110     cmdCopyTo.Font.Bold = True
120     cmdCopyTo.BackColor = &H8080FF
130 End If

140 Exit Sub

CheckCC_Error:

    Dim strES As String
    Dim intEL As Integer

150 intEL = Erl
160 strES = Err.Description
170 LogError "frmEditAll", "CheckCC", intEL, strES, sql

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

10  On Error GoTo CheckCholHDL_Error

20  If BRs Is Nothing Then Exit Sub

30  Chol = 0
40  HDL = 0
50  Ratio = 0

60  For Each br In BRs
70      Code = UCase$(Trim$(br.Code))
80      If Code = SysOptBioCodeForChol(0) Then
90          Chol = Val(br.Result)
100         CholDate = br.Rundate
110         CholTime = br.RunTime
120     ElseIf Code = SysOptBioCodeForHDL(0) Then
130         HDL = Val(br.Result)
140     ElseIf Code = SysOptBioCodeForCholHDLRatio(0) Then
150         Ratio = Val(br.Result)
160     End If
170 Next

180 If (Chol * HDL) <> 0 And Ratio = 0 Then
190     Ratio = Chol / HDL
200     Set br = New BIEResult
210     br.SampleID = txtSampleID
220     br.Code = SysOptBioCodeForCholHDLRatio(0)
230     br.ShortName = "C/H R"
240     br.Rundate = CholDate
250     br.RunTime = CholTime
260     br.Result = Format$(Ratio, "#0.00")
270     br.SampleType = "S"
280     br.Units = "Ratio"
290     br.Valid = 0
300     br.Printed = 0
        '  BR.Authorised = 0
310     br.Printformat = 1

320     BRs.Add br
330     BRResNew.Add br
340     BRResNew.Save "bio", BRResNew
350 End If

360 Exit Sub

CheckCholHDL_Error:

    Dim strES As String
    Dim intEL As Integer

370 intEL = Erl
380 strES = Err.Description
390 LogError "frmEditAll", "CheckCholHDL", intEL, strES

End Sub


Private Sub CheckDepartments()

10  On Error GoTo CheckDepartments_Error

20  If SysOptDeptHaem(0) = True Then
30      If AreHaemResultsPresent(txtSampleID) = 1 Then
40          ssTabAll.TabCaption(1) = "<<Haematology>>"
50      End If
60  End If

70  If SysOptDeptBio(0) = True Then
80      If AreBioResultsPresent(txtSampleID) = 1 Then
90          ssTabAll.TabCaption(2) = "<<Biochemistry>>"
100     End If
110 End If

120 If SysOptDeptCoag(0) = True Then
130     If AreCoagResultsPresent(txtSampleID) = 1 Then
140         ssTabAll.TabCaption(3) = "<<Coagulation>>"
150     End If
160 End If

170 If SysOptDeptEnd(0) = True Then
180     If AreEndResultsPresent(txtSampleID) = 1 Then
190         ssTabAll.TabCaption(4) = "<<Endocrinology>>"
200     End If
210 End If

220 If SysOptDeptBga(0) = True Then
230     If AreBgaResultsPresent(txtSampleID) = 1 Then
240         ssTabAll.TabCaption(5) = "<<Blood Gas>>"
250     End If
260 End If

270 If SysOptDeptImm(0) = True Then
280     If AreImmResultsPresent(txtSampleID) = 1 Then
290         ssTabAll.TabCaption(6) = "<<Immunology>>"
300     End If
310 End If

320 If SysOptDeptExt(0) = True Then
330     If AreExtResultsPresent(txtSampleID) = 1 Then
340         ssTabAll.TabCaption(7) = "<<Externals>>"
350     End If
360 End If

370 Exit Sub

CheckDepartments_Error:

    Dim strES As String
    Dim intEL As Integer

380 intEL = Erl
390 strES = Err.Description
400 LogError "frmEditAll", "CheckDepartments", intEL, strES

End Sub



Private Sub CheckIfPhoned()

10  On Error GoTo CheckIfPhoned_Error

20  If CheckPhoneLog(txtSampleID) Then
30      cmdPhone.BackColor = vbYellow
40      cmdPhone.Caption = "Results Phoned"
50      cmdPhone.ToolTipText = "Results Phoned"
60  Else
70      cmdPhone.BackColor = &H8000000F
80      cmdPhone.Caption = "Phone Results"
90      cmdPhone.ToolTipText = "Phone Results"
100 End If

110 Exit Sub

CheckIfPhoned_Error:

    Dim strES As String
    Dim intEL As Integer

120 intEL = Erl
130 strES = Err.Description
140 LogError "frmEditAll", "CheckIfPhoned", intEL, strES

End Sub

Private Sub chkBad_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

    Dim sql As String

10  On Error GoTo chkBad_MouseUp_Error

20  pBar = 0

30  If chkBad.Value = 1 Then
        'Code added 22/08/05
40      If iMsg("Do you wish all outstanding requests Deleted!", vbYesNo) = vbYes Then
50          sql = "DELETE from haemrequests WHERE sampleID = '" & txtSampleID & "'"
60          Cnxn(0).Execute sql
70      End If
80      If iMsg("Do you wish all to clear all results!", vbYesNo) = vbYes Then
90          ClearHgb
100         gRbc.TextMatrix(2, 1) = ""
110         gRbc.TextMatrix(2, 2) = ""
120         tESR = ""
130         tRetA = ""
140         tRetP = ""
150         tMonospot = ""
160         tRa = ""
170         tASOt = ""
180         lblMalaria = ""
190         lblSickledex = ""
200     End If
210 End If

220 cmdSaveHaem.Enabled = True: cmdHSaveH.Enabled = True

230 Exit Sub

chkBad_MouseUp_Error:

    Dim strES As String
    Dim intEL As Integer

240 intEL = Erl
250 strES = Err.Description
260 LogError "frmEditAll", "chkBad_MouseUp", intEL, strES

End Sub

Private Sub chkMalaria_Click()

10  On Error GoTo chkMalaria_Click_Error

20  cmdSaveHaem.Enabled = True: cmdHSaveH.Enabled = True

30  Exit Sub

chkMalaria_Click_Error:

    Dim strES As String
    Dim intEL As Integer

40  intEL = Erl
50  strES = Err.Description
60  LogError "frmEditAll", "chkMalaria_Click", intEL, strES

End Sub

Private Sub chkSickledex_Click()

10  On Error GoTo chkSickledex_Click_Error

20  cmdSaveHaem.Enabled = True: cmdHSaveH.Enabled = True

30  Exit Sub

chkSickledex_Click_Error:

    Dim strES As String
    Dim intEL As Integer

40  intEL = Erl
50  strES = Err.Description
60  LogError "frmEditAll", "chkSickledex_Click", intEL, strES

End Sub

Private Sub cIAdd_Click(Index As Integer)

    Dim SampleType As String
    Dim tb As New Recordset
    Dim sql As String

10  On Error GoTo cIAdd_Click_Error

20  If Index = 0 Then

30      cIUnits(0).Enabled = True

40      SampleType = ListCodeFor("ST", cISampleType(Index))

50      sql = "SELECT * from endtestdefinitions WHERE code = '" & eCodeForShortName(cIAdd(0)) & "'"
60      Set tb = New Recordset
70      RecOpenServer 0, tb, sql
80      If Not tb.EOF Then
90          cIUnits(0) = Trim(tb!Units & "")
100     Else
110         cIUnits(0) = ""
120     End If

130     cIUnits(0).Enabled = False
140 ElseIf Index = 1 Then
150     cIUnits(1).Enabled = True

160     sql = "SELECT * from Immtestdefinitions WHERE code = '" & ICodeForShortName(cIAdd(1)) & "'"
170     Set tb = New Recordset
180     RecOpenServer 0, tb, sql
190     If Not tb.EOF Then
200         cIUnits(1) = Trim(tb!Units) & ""
210     Else
220         cIUnits(1) = ""
230     End If

240     cIUnits(1).Enabled = False
250 ElseIf Index = 2 Then
260     cIUnits(2).Enabled = True

270     sql = "SELECT * from bgatestdefinitions WHERE code = '" & BgaCodeForShortName(cIAdd(2)) & "'"
280     Set tb = New Recordset
290     RecOpenServer 0, tb, sql
300     If Not tb.EOF Then
310         cIUnits(2) = Trim(tb!Units) & ""
320     Else
330         cIUnits(2) = ""
340     End If

350     cIUnits(2).Enabled = False
360 End If

370 Exit Sub

cIAdd_Click_Error:

    Dim strES As String
    Dim intEL As Integer

380 intEL = Erl
390 strES = Err.Description
400 LogError "frmEditAll", "cIAdd_Click", intEL, strES

End Sub

Private Sub cIAdd_KeyPress(Index As Integer, KeyAscii As Integer)

10  On Error GoTo cIAdd_KeyPress_Error

    '20        KeyAscii = AutoComplete(cIAdd(Index), KeyAscii, False)
    'KeyAscii = 0

20  Exit Sub

cIAdd_KeyPress_Error:

    Dim strES As String
    Dim intEL As Integer

30  intEL = Erl
40  strES = Err.Description
50  LogError "frmEditAll", "cIAdd_KeyPress", intEL, strES


End Sub

Private Sub cIAdd_LostFocus(Index As Integer)

10  On Error GoTo cIAdd_LostFocus_Error

    '20        cIAdd(Index).Text = QueryCombo(cIAdd(Index))

20  Exit Sub

cIAdd_LostFocus_Error:

    Dim strES As String
    Dim intEL As Integer

30  intEL = Erl
40  strES = Err.Description
50  LogError "frmEditAll", "cIAdd_LostFocus", intEL, strES

End Sub

Private Sub cISampleType_Change(Index As Integer)

'Dim Department As String

    On Error GoTo cISampleType_Change_Error

    'Department = ""
    'Department = Choose(Index, "End", "Imm", "Bio")
    'If SampleType <> "" And Department <> "" Then
    '    If ListCodeFor("ST", cISampleType(Index)) <> GetSampleType(Department, txtSampleID) Then
    '        If iMsg("Sample type for previous results is different. Prevoius sample type would be changed to " & _
             '                cISampleType(Index) & " Do you want to proceed?", vbYesNo) = vbNo Then
    '            Exit Sub
    '        Else
    '            'update sample type for previous results
    '
    '        End If
    '    End If
    'End If

    If Index = 0 Then
        FillcEAdd
    ElseIf Index = 1 Then
        FillcIAdd
    ElseIf Index = 2 Then
        FillcbAdd
    ElseIf Index = 3 Then
        FillcAdd
    End If




    Exit Sub

cISampleType_Change_Error:

    Dim strES As String
    Dim intEL As Integer

    intEL = Erl
    strES = Err.Description
140 LogError "frmEditAll", "cISampleType_Change", intEL, strES

End Sub

Private Sub cISampleType_Click(Index As Integer)

    On Error GoTo cISampleType_Click_Error

    If Index = 0 Then
        FillcEAdd
    ElseIf Index = 1 Then
        FillcIAdd
    ElseIf Index = 2 Then
        FillcbAdd
    ElseIf Index = 3 Then
        FillcAdd
        SampleType = ListCodeFor("ST", cISampleType(Index))
        If SampleType = "T" Then
            cAdd = ""
            tnewvalue = ""
            cUnits = ""
            frmEditToxicology.Show 1
            LoadBiochemistry
        End If

    End If

    Exit Sub

cISampleType_Click_Error:

    Dim strES As String
    Dim intEL As Integer

    intEL = Erl
    strES = Err.Description
    LogError "frmEditAll", "cISampleType_Click", intEL, strES

End Sub

Private Sub ClearCoagulation()

10  On Error GoTo ClearCoagulation_Error

20  cParameter = ""
30  cCunits.ListIndex = -1
40  tResult = ""
50  tWarfarin = ""
60  bViewCoagRepeat.Visible = False
70  lCDate = ""

80  Exit Sub

ClearCoagulation_Error:

    Dim strES As String
    Dim intEL As Integer

90  intEL = Erl
100 strES = Err.Description
110 LogError "frmEditAll", "ClearCoagulation", intEL, strES

End Sub



Public Sub ClearDemographics()

    Dim n As Long
    Dim TimeNow As String

10  On Error GoTo ClearDemographics_Error

20  lblUrgent.Visible = False
30  mNewRecord = True
40  dtRunDate = Format$(Now, "dd/mm/yyyy")
50  lblRundate = dtRunDate
60  dtSampleDate = Format$(Now, "dd/mm/yyyy")
70  lblSampleDate = dtSampleDate
80  dtRecDate = Format$(Now, "dd/mm/yyyy")
90  If SysOptDemoVal(0) Then cmdDemoVal.Caption = "&Validate"
100 txtChart = ""
110 txtName = ""
120 taddress(0) = ""
130 taddress(1) = ""
140 txtAandE = ""
150 StatusBar1.Panels(4).Text = ""
160 txtSex = ""
170 txtDoB = ""
180 txtAge = ""
190 lDoB = ""
200 lAge = ""
210 lSex = ""
220 cmbWard = "GP"
230 cmbClinician = ""
240 cmbGP = ""
250 cClDetails = ""
260 txtDemographicComment = ""

270 TimeNow = Format$(Now, "HH:nn")
280 tSampleTime.Mask = ""
290 tSampleTime.Text = ""   'TimeNow
300 tSampleTime.Mask = "##:##"
310 tRecTime.Mask = ""
320 tRecTime.Text = TimeNow
330 tRecTime.Mask = "##:##"

340 lblChartNumber.Caption = HospName(0) & " Chart #"
350 lblChartNumber.BackColor = &H8000000F
360 lblChartNumber.ForeColor = vbBlack
370 cCat(0).ListIndex = -1
380 cCat(1).ListIndex = -1

390 If cmbHospital = "" Then
400     For n = 0 To cmbHospital.ListCount - 1
410         If UCase(cmbHospital.List(n)) = HospName(0) Then
420             cmbHospital.ListIndex = n
430         End If
440     Next
450 End If
460 EnableDemographicEntry True

470 Exit Sub

ClearDemographics_Error:

    Dim strES As String
    Dim intEL As Integer

480 intEL = Erl
490 strES = Err.Description
500 LogError "frmEditAll", "ClearDemographics", intEL, strES

End Sub

Private Sub ClearEndFlags()

10  On Error GoTo ClearEndFlags_Error

20  Ih(0) = 0
30  Iis(0) = 0
40  Il(0) = 0
50  Io(0) = 0
60  Ig(0) = 0
70  Ij(0) = 0

80  Exit Sub

ClearEndFlags_Error:

    Dim strES As String
    Dim intEL As Integer

90  intEL = Erl
100 strES = Err.Description
110 LogError "frmEditAll", "ClearEndFlags", intEL, strES

End Sub

Private Sub ClearExt()

10  On Error GoTo ClearExt_Error

20  ClearFGrid grdExt
30  grdExt.Visible = True

40  Exit Sub

ClearExt_Error:

    Dim strES As String
    Dim intEL As Integer

50  intEL = Erl
60  strES = Err.Description
70  LogError "frmEditAll", "ClearExt", intEL, strES

End Sub

Private Sub ClearHaematologyResults()

    Dim n As Long

10  On Error GoTo ClearHaematologyResults_Error

20  ClearRbcGrid
30  ClearHaemDiffGrid

    'HGB_Click

40  lWIC = ""
50  lWOC = ""

60  tWBC = ""
70  tWBC.BackColor = &HFFFFFF
80  tWBC.ForeColor = &H0&

90  tPlt = ""
100 tPlt.BackColor = &HFFFFFF
110 tPlt.ForeColor = &H0&

120 tMPV = ""
130 tMPV.BackColor = &HFFFFFF
140 tMPV.ForeColor = &H0&

150 lblAnalyser = "Analyser : "
160 txtLI = ""
170 txtMPXI = ""

180 pdelta.Cls
190 lHDate = ""
200 cESR = 0
210 cRetics = 0
220 cMonospot = 0
230 cRA = 0
240 cASot = 0
250 chkMalaria = 0
260 chkSickledex = 0
270 chkBad = 0

280 tESR = ""
290 tESR.BackColor = &HFFFFFF
300 tESR.ForeColor = &H0&

310 txtEsr1 = ""
320 txtEsr1.BackColor = &HFFFFFF
330 txtEsr1.ForeColor = &H0&

340 tRetA = ""
350 tRetP = ""
360 tRetA.BackColor = vbWhite
370 tRetA.ForeColor = 1
380 tMonospot = ""
390 cFilm = 0
400 tRa = ""
410 tASOt = ""
420 lblMalaria = ""
430 lblSickledex = ""

    'cCoag = 0

440 tWarfarin = ""

450 For n = 0 To 5
460     ipflag(n).Visible = False
470 Next

480 Exit Sub

ClearHaematologyResults_Error:

    Dim strES As String
    Dim intEL As Integer

490 intEL = Erl
500 strES = Err.Description
510 LogError "frmEditAll", "ClearHaematologyResults", intEL, strES

End Sub

Private Sub ClearHaemDiffGrid()

10  On Error GoTo ClearHaemDiffGrid_Error

20  With grdH

30      .Rows = 2
40      .AddItem ""
50      .RemoveItem 1

60      .AddItem vbTab & vbTab & "Neut"
70      .AddItem vbTab & vbTab & "Lymph"
80      .AddItem vbTab & vbTab & "Mono"
90      .AddItem vbTab & vbTab & "Eos"
100     .AddItem vbTab & vbTab & "Bas"
110     .AddItem vbTab & vbTab & "Luc"

120     .RemoveItem 1

130 End With

140 Exit Sub

ClearHaemDiffGrid_Error:

    Dim strES As String
    Dim intEL As Integer

150 intEL = Erl
160 strES = Err.Description
170 LogError "frmEditAll", "ClearHaemDiffGrid", intEL, strES

End Sub

Private Sub ClearHgb()

    Dim n As Long

10  On Error GoTo ClearHgb_Error

20  pBar = 0

30  bcleardiff_click

'40  gRBC.TextMatrix(1, 1) = ""
50  gRbc.Row = 1
60  gRbc.Col = 1
70  gRbc.CellBackColor = vbWhite
80  gRbc.CellForeColor = 1
90  gRbc.Col = 2
'100 gRBC.TextMatrix(1, 2) = ""
110 gRbc.CellBackColor = &H8000000F
120 gRbc.CellForeColor = 1
'Zyam changed the n to start from 1 instead of 3 to clear all the fbc results 9-6-24

130 For n = 1 To gRbc.Rows - 1
140     gRbc.Row = n
150     gRbc.Col = 1
160     gRbc = ""
170     gRbc.CellBackColor = vbWhite
180     gRbc.CellForeColor = 1
190     gRbc.Col = 2
200     gRbc = ""
210     gRbc.CellBackColor = &H8000000F
220     gRbc.CellForeColor = 1
        
230 Next
'Zyam 9-6-24
240 tWBC = ""
250 tWBC.BackColor = &HFFFFFF
260 tWBC.ForeColor = 1
270 tPlt = ""
280 tPlt.BackColor = &HFFFFFF
290 tPlt.ForeColor = 1
300 txtMPXI = ""
310 txtMPXI.BackColor = &HFFFFFF
320 txtMPXI.ForeColor = 1
330 lWIC = ""
340 lWOC = ""
350 tMPV = ""
360 tMPV.BackColor = &HFFFFFF
370 tMPV.ForeColor = 1
380 txtLI = ""
390 txtLI.BackColor = &HFFFFFF
400 txtLI.ForeColor = 1

410 cmdSaveHaem.Enabled = True: cmdHSaveH.Enabled = True
420 bValidateHaem.Enabled = True

430 Exit Sub

ClearHgb_Error:

    Dim strES As String
    Dim intEL As Integer

440 intEL = Erl
450 strES = Err.Description
460 LogError "frmEditAll", "ClearHgb", intEL, strES

End Sub

Private Sub ClearImmFlags()

10  On Error GoTo ClearImmFlags_Error

20  Ih(1) = 0
30  Iis(1) = 0
40  Il(1) = 0
50  Io(1) = 0
60  Ig(1) = 0
70  Ij(1) = 0

80  Exit Sub

ClearImmFlags_Error:

    Dim strES As String
    Dim intEL As Integer

90  intEL = Erl
100 strES = Err.Description
110 LogError "frmEditAll", "ClearImmFlags", intEL, strES

End Sub


Private Sub ClearOutstanding(ByVal grd As MSFlexGrid)

10  With grd
20      .Rows = 2
30      .AddItem ""
40      .RemoveItem 1
50  End With

End Sub


Private Sub ClearRbcGrid()

    Dim n As Long

10  On Error GoTo ClearRbcGrid_Error

20  With gRbc

30      .Rows = 2
40      .AddItem ""
50      .RemoveItem 1

60      .AddItem "RBC"
70      .AddItem "Hgb"
80      .AddItem "HCT"
90      .AddItem "MCV"
100     .AddItem "HDW"
110     .AddItem "MCH"
120     .AddItem "MCHC"
130     .AddItem "CHCM"
140     .AddItem "RDW"
150     .AddItem "NRBC%"
160     .AddItem "HYPO%"

170     .RemoveItem 1

180     For n = 1 To .Rows - 1
190         .Row = n
200         .Col = 0
210         .CellFontBold = True
220         .CellBackColor = &H8000000F
230         .CellForeColor = &HC0&
240         .Col = 1
250         .CellFontBold = True
260         .CellBackColor = &H80000005
270         .CellForeColor = vbBlack
280         .Col = 2
290         .CellFontBold = True
300         .CellBackColor = &H80000005
310         .CellForeColor = vbBlack
320     Next
330 End With

340 Exit Sub

ClearRbcGrid_Error:

    Dim strES As String
    Dim intEL As Integer

350 intEL = Erl
360 strES = Err.Description
370 LogError "frmEditAll", "ClearRbcGrid", intEL, strES

End Sub

Private Sub cISampleType_KeyPress(Index As Integer, KeyAscii As Integer)
10  KeyAscii = 0
End Sub

Private Sub cmbClinician_Change()

10  On Error GoTo cmbClinician_Change_Error

20  SetWardClinGP

30  Exit Sub

cmbClinician_Change_Error:

    Dim strES As String
    Dim intEL As Integer

40  intEL = Erl
50  strES = Err.Description
60  LogError "frmEditAll", "cmbClinician_Change", intEL, strES

End Sub

Private Sub cmbClinician_Click()

10  On Error GoTo cmbClinician_Click_Error

20  cmdSaveDemographics.Enabled = True
30  cmdSaveInc.Enabled = True

40  Exit Sub

cmbClinician_Click_Error:

    Dim strES As String
    Dim intEL As Integer

50  intEL = Erl
60  strES = Err.Description
70  LogError "frmEditAll", "cmbClinician_Click", intEL, strES

End Sub

Private Sub cmbClinician_KeyPress(KeyAscii As Integer)

10  On Error GoTo cmbClinician_KeyPress_Error

    '20        KeyAscii = AutoComplete(cmbClinician, KeyAscii, False)
20  cmdSaveDemographics.Enabled = True
30  cmdSaveInc.Enabled = True

40  Exit Sub

cmbClinician_KeyPress_Error:

    Dim strES As String
    Dim intEL As Integer

50  intEL = Erl
60  strES = Err.Description
70  LogError "frmEditAll", "cmbClinician_KeyPress", intEL, strES

End Sub

Private Sub cmbClinician_LostFocus()

10  On Error GoTo cmbClinician_LostFocus_Error

20  pBar = 0
30  cmbClinician = QueryKnown("Clin", cmbClinician, UCase(cmbHospital))

40  Exit Sub

cmbClinician_LostFocus_Error:

    Dim strES As String
    Dim intEL As Integer

50  intEL = Erl
60  strES = Err.Description
70  LogError "frmEditAll", "cmbClinician_LostFocus", intEL, strES

End Sub



Private Sub cmbEndResults_Click()

    Dim sql As String

10  On Error GoTo cmbEndResults_Click_Error

20  If LTrim(RTrim(cmbEndResults)) = "" Then
30      cmbEndResults.Visible = False
40      Exit Sub
50  End If

60  gImm(0) = cmbEndResults.Text
70  sql = "UPDATE EndResults " & _
          "SET Result = '" & AddTicks(cmbEndResults.Text) & "' " & _
          "WHERE sampleid = '" & txtSampleID & "' " & _
          "AND Code = '" & eCodeForShortName(gImm(0).TextMatrix(gImm(0).Row, 0)) & "'"
80  Cnxn(0).Execute sql
90  cmbEndResults.Visible = False

100 Exit Sub

cmbEndResults_Click_Error:

    Dim strES As String
    Dim intEL As Integer

110 intEL = Erl
120 strES = Err.Description
130 LogError "frmEditAll", "cmbEndResults_Click", intEL, strES, sql

End Sub

Private Sub cmbEndResults_KeyPress(KeyAscii As Integer)
10  KeyAscii = 0
End Sub

Private Sub cmbGP_Change()

10  On Error GoTo cmbGP_Change_Error

20  SetWardClinGP

30  cmbWard = "GP"

40  Exit Sub

cmbGP_Change_Error:

    Dim strES As String
    Dim intEL As Integer

50  intEL = Erl
60  strES = Err.Description
70  LogError "frmEditAll", "cmbGP_Change", intEL, strES

End Sub

Private Sub cmbGP_Click()

10  On Error GoTo cmbGP_Click_Error

20  pBar = 0

30  SetWardClinGP

40  cmbWard = "GP"
50  cmdSaveDemographics.Enabled = True
60  cmdSaveInc.Enabled = True

70  Exit Sub

cmbGP_Click_Error:

    Dim strES As String
    Dim intEL As Integer

80  intEL = Erl
90  strES = Err.Description
100 LogError "frmEditAll", "cmbGP_Click", intEL, strES

End Sub

Private Sub cmbGP_KeyPress(KeyAscii As Integer)

10  On Error GoTo cmbGP_KeyPress_Error

    '20        KeyAscii = AutoComplete(cmbGP, KeyAscii, False)
20  cmdSaveDemographics.Enabled = True
30  cmdSaveInc.Enabled = True

40  Exit Sub

cmbGP_KeyPress_Error:

    Dim strES As String
    Dim intEL As Integer

50  intEL = Erl
60  strES = Err.Description
70  LogError "frmEditAll", "cmbGP_KeyPress", intEL, strES

End Sub

Private Sub cmbGP_LostFocus()

10  On Error GoTo cmbGP_LostFocus_Error

20  cmbGP = QueryKnown("GP", cmbGP, cmbHospital)

30  Exit Sub

cmbGP_LostFocus_Error:

    Dim strES As String
    Dim intEL As Integer

40  intEL = Erl
50  strES = Err.Description
60  LogError "frmEditAll", "cmbGP_LostFocus", intEL, strES

End Sub

Private Sub cmbHospital_Click()

10  On Error GoTo cmbHospital_Click_Error

20  cmbWard.Clear
30  cmbGP.Clear
40  cmbClinician.Clear

50  FillGPsClinWard Me, cmbHospital

60  cmdSaveDemographics.Enabled = True
70  cmdSaveInc.Enabled = True


80  Exit Sub

cmbHospital_Click_Error:

    Dim strES As String
    Dim intEL As Integer

90  intEL = Erl
100 strES = Err.Description
110 LogError "frmEditAll", "cmbHospital_Click", intEL, strES

End Sub

Private Sub cmbHospital_KeyPress(KeyAscii As Integer)
'10        KeyAscii = AutoComplete(cmbHospital, KeyAscii, False)
End Sub

Private Sub cmbHospital_LostFocus()
    Dim n As Long

10  On Error GoTo cmbHospital_LostFocus_Error

20  For n = 0 To cmbHospital.ListCount
30      If UCase(cmbHospital) = UCase(Left(cmbHospital.List(n), Len(cmbHospital))) Then
40          cmbHospital.ListIndex = n
50      End If
60  Next

70  Exit Sub

cmbHospital_LostFocus_Error:

    Dim strES As String
    Dim intEL As Integer

80  intEL = Erl
90  strES = Err.Description
100 LogError "frmEditAll", "cmbHospital_LostFocus", intEL, strES

End Sub

Private Sub cmbWard_Change()

10  On Error GoTo cmbWard_Change_Error

20  SetWardClinGP

30  Exit Sub

cmbWard_Change_Error:

    Dim strES As String
    Dim intEL As Integer

40  intEL = Erl
50  strES = Err.Description
60  LogError "frmEditAll", "cmbWard_Change", intEL, strES

End Sub

Private Sub cmbWard_Click()

10  On Error GoTo cmbWard_Click_Error

20  SetWardClinGP

30  cmdSaveDemographics.Enabled = True
40  cmdSaveInc.Enabled = True

50  Exit Sub

cmbWard_Click_Error:

    Dim strES As String
    Dim intEL As Integer

60  intEL = Erl
70  strES = Err.Description
80  LogError "frmEditAll", "cmbWard_Click", intEL, strES

End Sub

Private Sub cmbWard_KeyPress(KeyAscii As Integer)

10  On Error GoTo cmbWard_KeyPress_Error

    '20        KeyAscii = AutoComplete(cmbWard, KeyAscii, False)
20  cmdSaveDemographics.Enabled = True
30  cmdSaveInc.Enabled = True

40  Exit Sub

cmbWard_KeyPress_Error:

    Dim strES As String
    Dim intEL As Integer

50  intEL = Erl
60  strES = Err.Description
70  LogError "frmEditAll", "cmbWard_KeyPress", intEL, strES

End Sub

Private Sub cmbWard_LostFocus()

    Dim tb As New Recordset
    Dim sql As String

10  On Error GoTo cmbWard_LostFocus_Error

20  If Trim$(cmbWard) = "" Then
30      cmbWard = "GP"
40      Exit Sub
50  End If

60  sql = "SELECT Text FROM Wards WHERE " & _
          "(Text = '" & AddTicks(cmbWard) & "' " & _
          "OR Code = '" & AddTicks(cmbWard) & "') " & _
          "AND HospitalCode = '" & ListCodeFor("HO", cmbHospital) & "' " & _
          "AND InUse = 1"
70  Set tb = New Recordset
80  RecOpenServer 0, tb, sql
90  If Not tb.EOF Then
100     cmbWard = Trim(tb!Text)
110 Else
120     cmbWard = "GP"
130 End If

140 Exit Sub

cmbWard_LostFocus_Error:

    Dim strES As String
    Dim intEL As Integer

150 intEL = Erl
160 strES = Err.Description
170 LogError "frmEditAll", "cmbWard_LostFocus", intEL, strES, sql

End Sub

Private Sub cmdAudit_Click()

    Dim Disp As String

10  On Error GoTo cmdAudit_Click_Error


20  If ssTabAll.Tab = 0 Then
30      With frmArchiveNew
40          .TableName = "Demographics"
50          .SampleID = txtSampleID
60          .Show 1
70      End With
80  Else
90      Disp = Choose(ssTabAll.Tab, "", "Bio", "Coag", "End", "Bga", "Imm")
100     With frmAudit
110         .TableName = Disp & "Results"
120         .SampleID = txtSampleID
130         .Show 1
140     End With

150     Select Case Disp
        Case "Bio": LoadBiochemistry
160     Case "Coag": LoadCoagulation
170     Case "End": LoadEndocrinology
180     Case "Bga": LoadBloodGas
190     Case "Imm": LoadImmunology
200     End Select
210 End If



220 Exit Sub

cmdAudit_Click_Error:

    Dim strES As String
    Dim intEL As Integer

230 intEL = Erl
240 strES = Err.Description
250 LogError "frmEditAll", "cmdAudit_Click", intEL, strES

End Sub

Private Sub cmdCopyTo_Click()
    Dim s As String

10  On Error GoTo cmdCopyTo_Click_Error

20  s = cmbWard & " " & cmbClinician & " " & cmbGP
30  s = Trim$(s)

40  frmCopyTo.EditScreen = Me
50  frmCopyTo.lblOriginal = s
60  frmCopyTo.lblSampleID = txtSampleID
70  frmCopyTo.Show 1

80  CheckCC


90  Exit Sub

cmdCopyTo_Click_Error:

    Dim strES As String
    Dim intEL As Integer

100 intEL = Erl
110 strES = Err.Description
120 LogError "frmEditAll", "cmdCopyTo_Click", intEL, strES

End Sub

Private Sub cmdDartViewer_Click()
10  On Error GoTo cmdDartViewer_Click_Error

20  If Dir("C:\Program Files\The PlumTree Group\Dartviewer\Dartviewer.exe") = "" Then
30      iMsg "Dart client not installed on this machine. Please contact you system administrator", vbInformation
40      Exit Sub
50  End If

60  Shell "C:\Program Files\The PlumTree Group\Dartviewer\Dartviewer.exe " & txtSampleID, vbNormalFocus

70  Exit Sub

cmdDartViewer_Click_Error:

    Dim strES As String
    Dim intEL As Integer

80  intEL = Erl
90  strES = Err.Description
100 LogError "frmEditAll", "cmdDartViewer_Click", intEL, strES
End Sub

Private Sub cmdDel_Click()

    Dim Str As String

10  On Error GoTo cmdDel_Click_Error

20  If grdExt.TextMatrix(grdExt.Row, 0) = "" Then Exit Sub

30  Str = "  Test Name : " & grdExt.TextMatrix(grdExt.Row, 0) & vbCrLf & _
          "DELETE this test?"
40  If iMsg(Str, vbQuestion + vbYesNo, "Confirm Deletion") = vbYes Then
50      Str = "DELETE from extresults WHERE " & _
              "sampleid = '" & txtSampleID & "' " & _
              "and Analyte = '" & grdExt.TextMatrix(grdExt.Row, 0) & "'"
60      Cnxn(0).Execute Str
70      LoadExt
80  End If

90  Exit Sub

cmdDel_Click_Error:

    Dim strES As String
    Dim intEL As Integer

100 intEL = Erl
110 strES = Err.Description
120 LogError "frmEditAll", "cmdDel_Click", intEL, strES, Str

End Sub

Private Sub cmdDemoVal_Click()

    Dim Validating As Boolean

10  On Error GoTo cmdDemoVal_Click_Error

20  If Not EntriesOK(txtSampleID, txtName, txtSex, cmbWard.Text, cmbGP.Text) Then
30      Exit Sub
40  End If

50  Validating = cmdDemoVal.Caption = "&Validate"

60  If Not Validating And UCase(UserMemberOf) = "SECRETARYS" And GetOptionSetting("DEMOUNVALIDATE", "0") = 1 Then
70      iMsg "You do not have permissions to unvalidate this record"
80      Exit Sub
90  End If

100 If Validating Then
110     cmdSaveDemographics_Click
120 End If

130 ValidateDemographics Validating

140 If Validating Then
150     If GetOptionSetting("RollForward", "0") = "1" Then
160         txtSampleID = Format$(Val(txtSampleID) + 1)
170         txtSampleID.SelStart = 0
180         txtSampleID.SelLength = Len(txtSampleID)
190         txtSampleID.SetFocus

200     Else
210         If UserMemberOf = "Secretarys" Then
220             txtSampleID = ""
230         Else
240             txtSampleID = Format$(Val(txtSampleID) + 1)
250         End If
260     End If
270 End If

    '          If UserMemberOf = "Secretarys" Then
    '110         txtSampleID = ""
    '120       ElseIf Validating Then
    '130         txtSampleID = Format$(Val(txtSampleID) + 1)
    '140       End If



280 LoadAllDetails

290 cmdSaveHaem.Enabled = False
300 cmdHSaveH.Enabled = False
310 cmdSaveBio.Enabled = False
320 cmdSaveCoag.Enabled = False
330 cmdSaveImm(0).Enabled = False
340 cmdSaveImm(1).Enabled = False
350 cmdSaveBGa.Enabled = False

360 Exit Sub

cmdDemoVal_Click_Error:

    Dim strES As String
    Dim intEL As Integer

370 intEL = Erl
380 strES = Err.Description
390 LogError "frmEditAll", "cmdDemoVal_Click", intEL, strES

End Sub

Private Sub cmdExcel_Click()

10  On Error GoTo cmdExcel_Click_Error

20  ExportFlexGrid grdExt, Me

30  Exit Sub

cmdExcel_Click_Error:

    Dim strES As String
    Dim intEL As Integer

40  intEL = Erl
50  strES = Err.Description
60  LogError "frmEditAll", "cmdExcel_Click", intEL, strES

End Sub

Private Sub cmdGetBio_Click()

10  On Error GoTo cmdGetBio_Click_Error

20  LoadBiochemistry
30  frmBio2Imm.Show 1
40  LoadImmunology

50  Exit Sub

cmdGetBio_Click_Error:

    Dim strES As String
    Dim intEL As Integer

60  intEL = Erl
70  strES = Err.Description
80  LogError "frmEditAll", "cmdGetBio_Click", intEL, strES

End Sub

Private Sub cmdGetBioEnd_Click()

    Dim n As Integer
    Dim ResultsFound As Boolean

10  On Error GoTo cmdGetBioEnd_Click_Error

20  LoadBiochemistry

30  ResultsFound = False
40  For n = 1 To gBio.Rows - 1
50      If InStr(1, gBio.TextMatrix(n, 6), "V") > 0 Then
60          ResultsFound = True
70      End If

80  Next n

90  If Not ResultsFound Then
100     iMsg "No validated results found", vbInformation
110     Exit Sub
120 End If


130 frmBio2End.Show 1
140 LoadEndocrinology

150 Exit Sub

cmdGetBioEnd_Click_Error:

    Dim strES As String
    Dim intEL As Integer

160 intEL = Erl
170 strES = Err.Description
180 LogError "frmEditAll", "cmdGetBioEnd_Click", intEL, strES

End Sub

Private Sub cmdHSaveH_Click()

10  On Error GoTo cmdHSaveH_Click_Error

20  pBar = 0

30  txtSampleID = Format(Val(txtSampleID))
40  If Val(txtSampleID) = 0 Then Exit Sub

    'Added 15/Jul/2004

50  If bValidateHaem.Caption = "&Validate" Then
60      SaveHaematology 0
70  Else
80      SaveHaematology 1
90  End If

100 SaveComments
110 UPDATEMRU txtSampleID, cMRU

120 LoadAllDetails

130 cmdSaveHaem.Enabled = False
140 cmdHSaveH.Enabled = False


150 Exit Sub

cmdHSaveH_Click_Error:

    Dim strES As String
    Dim intEL As Integer

160 intEL = Erl
170 strES = Err.Description
180 LogError "frmEditAll", "cmdHSaveH_Click", intEL, strES

End Sub

Private Sub cmdIAdd_Click(Index As Integer)
    Dim sql As String
    Dim s As String
    Dim n As Long
    Dim Code As String

10  On Error GoTo cmdIAdd_Click_Error

20  pBar = 0

30  If Index = 0 Then

40      If cIAdd(0).Text = "" Then Exit Sub
50      If Trim$(tINewValue(0)) = "" Then Exit Sub
60      If Trim$(txtSampleID) = "" Then Exit Sub

70      For n = 1 To gImm(0).Rows - 1
80          If cIAdd(0) = gImm(0).TextMatrix(n, 0) Then
90              iMsg "Test already Exists. Please delete before adding!"
100             Exit Sub
110         End If
120     Next
130     s = Check_End(cIAdd(0).Text, cIUnits(0), cISampleType(0))
140     If s <> "" Then
150         iMsg s & " is incorrect!"
160         Exit Sub
170     End If

180     Code = eCodeForShortName(cIAdd(0).Text)
190     sql = "INSERT into endResults " & _
              "(RunDate, SampleID, Code, Result, RunTime, Units, SampleType, Valid, Printed) VALUES " & _
              "('" & Format$(dtRunDate, "dd/mmm/yyyy") & "', " & _
              "'" & txtSampleID & "', " & _
              "'" & Code & "', " & _
              "'" & tINewValue(0) & "', " & _
              "'" & Format$(Now, "dd/mmm/yyyy hh:mm") & "', " & _
              "'" & cIUnits(0) & "', " & _
              "'" & ListCodeFor("ST", cISampleType(0)) & "', 0, 0);"

200     Cnxn(0).Execute sql

210     sql = "DELETE FROM EndRequests " & _
              "WHERE SampleID = '" & txtSampleID & "' " & _
              "AND Code = '" & Code & "'"
220     Cnxn(0).Execute sql

230     LoadEndocrinology

240     cIAdd(0) = ""
250     tINewValue(0) = ""
260     cIUnits(0) = ""

270 ElseIf Index = 1 Then

280     If cIAdd(1).Text = "" Then Exit Sub
290     If Trim$(tINewValue(1)) = "" Then Exit Sub
300     If Trim$(txtSampleID) = "" Then Exit Sub

310     For n = 1 To gImm(1).Rows - 1
320         If gImm(1).TextMatrix(n, 0) = cIAdd(1) Then
330             iMsg "Test already in List!"
340             Exit Sub
350         End If
360     Next

370     s = Check_Imm(cIAdd(1).Text, cIUnits(1), cISampleType(1))
380     If s <> "" Then
390         iMsg s & " is incorrect!"
400         Exit Sub
410     End If

420     Code = ICodeForShortName(cIAdd(1).Text)
430     sql = "INSERT into ImmResults " & _
              "(RunDate, SampleID, Code, Result, RunTime, Units, SampleType, Valid, Printed) VALUES " & _
              "('" & Format$(dtRunDate, "dd/mmm/yyyy") & "', " & _
              "'" & txtSampleID & "', " & _
              "'" & Code & "', " & _
              "'" & AddTicks(tINewValue(1)) & "', " & _
              "'" & Format$(Now, "dd/mmm/yyyy hh:mm") & "', " & _
              "'" & cIUnits(1) & "', " & _
              "'" & ListCodeFor("ST", cISampleType(1)) & "', 0, 0);"

440     Cnxn(0).Execute sql

450     sql = "DELETE FROM ImmRequests " & _
              "WHERE SampleID = '" & txtSampleID & "' " & _
              "AND Code = '" & Code & "'"
460     Cnxn(0).Execute sql

470     LoadImmunology

480     cIAdd(1) = ""
490     tINewValue(1) = ""
500     cIUnits(1) = ""

510 ElseIf Index = 2 Then

520     If cIAdd(2).Text = "" Then Exit Sub
530     If Trim$(tINewValue(2)) = "" Then Exit Sub
540     If Trim$(txtSampleID) = "" Then Exit Sub

550     For n = 1 To gBga.Rows - 1
560         If gBga.TextMatrix(n, 0) = cIAdd(2) Then
570             iMsg "Test already in List!"
580             Exit Sub
590         End If
600     Next

610     sql = "INSERT into bgaResults " & _
              "(RunDate, SampleID, Code, Result, RunTime, Units, SampleType, Valid, Printed) VALUES " & _
              "('" & Format$(dtRunDate, "dd/mmm/yyyy") & "', " & _
              "'" & txtSampleID & "', " & _
              "'" & BgaCodeForShortName(cIAdd(2).Text) & "', " & _
              "'" & tINewValue(2) & "', " & _
              "'" & Format$(Now, "dd/mmm/yyyy hh:mm") & "', " & _
              "'" & cIUnits(2) & "', " & _
              "'S', 0, 0);"

620     Cnxn(0).Execute sql

630     sql = "UPDATE Demographics set Forbga = 1 WHERE sampleID = '" & txtSampleID & "'"
640     Cnxn(0).Execute sql

650     LoadBloodGas

660     cIAdd(2) = ""
670     tINewValue(2) = ""
680     cIUnits(2) = ""

690 End If

700 Exit Sub

cmdIAdd_Click_Error:

    Dim strES As String
    Dim intEL As Integer

710 intEL = Erl
720 strES = Err.Description
730 LogError "frmEditAll", "cmdIAdd_Click", intEL, strES, sql

End Sub

Private Sub cmdIremoveduplicates_Click(Index As Integer)
    Dim tb As New Recordset
    Dim sql As String
    Dim Y As Long
    Dim Code As String


10  On Error GoTo cmdIremoveduplicates_Click_Error

20  pBar = 0

30  If Index = 0 Then

40      If gImm(0).Rows < 3 Then Exit Sub

50      For Y = 1 To gImm(0).Rows - 1
60          Code = eCodeForShortName(gImm(0).TextMatrix(Y, 0))
70          sql = "SELECT * from Endresults WHERE " & _
                  "sampleid = '" & txtSampleID & "' " & _
                  "and code = '" & Code & "' order by runtime asc"
80          Set tb = New Recordset
90          RecOpenClient 0, tb, sql
100         Do While tb.recordCount > 1
110             tb.DELETE
120             tb.MoveNext
130         Loop
140     Next

150     LoadEndocrinology
160 Else

170     If gImm(1).Rows < 3 Then Exit Sub

180     For Y = 1 To gImm(1).Rows - 1
190         Code = ICodeForShortName(gImm(1).TextMatrix(Y, 0))
200         sql = "SELECT * from Immresults WHERE " & _
                  "sampleid = '" & txtSampleID & "' " & _
                  "and code = '" & Code & "' order by runtime asc"
210         Set tb = New Recordset
220         RecOpenClient 0, tb, sql
230         Do While tb.recordCount > 1
240             tb.DELETE
250             tb.MoveNext
260         Loop
270     Next

280     LoadImmunology
290 End If

300 Exit Sub

cmdIremoveduplicates_Click_Error:

    Dim strES As String
    Dim intEL As Integer

310 intEL = Erl
320 strES = Err.Description
330 LogError "frmEditAll", "cmdIremoveduplicates_Click", intEL, strES, sql

End Sub

Private Sub cmdOrderPhoresis_Click()

10  On Error GoTo cmdOrderPhoresis_Click_Error

20  With frmOrderPhoresis
30      .lblName = txtName
40      .lblChart = txtChart
50      .lblAandE = txtAandE
60      .lblDoB = txtDoB
70      .lblAge = txtAge
80      .lblSex = txtSex
90      .lblClinician = cmbClinician
100     .lblSampleDate = lblSampleDate
110     .lblSampleID = txtSampleID

120     .Show 1
130 End With

140 Exit Sub

cmdOrderPhoresis_Click_Error:

    Dim strES As String
    Dim intEL As Integer

150 intEL = Erl
160 strES = Err.Description
170 LogError "frmEditAll", "cmdOrderPhoresis_Click", intEL, strES

End Sub

Private Sub cmdPatientNotePad_Click(Index As Integer)
'Index 0 = Patient notepad
'Index 1 = Barcode Printing

On Error GoTo cmdPatientNotePad_Click_Error

Select Case Index
Case 0
    frmPatientNotePad.SampleID = txtSampleID
    frmPatientNotePad.Caller = "General"
    frmPatientNotePad.Show 1
Case 1
    '
    'Dim dummyval As String
    'Dim temp as long
    '
    'dummyval = " "
    'temp = winhelp(acquire.hWnd, "c:\windows\help\windows.hlp", HELP_PARTIALKEY, dummyval)
'    Dim tb As New Recordset
'    Dim tbSN As New Recordset    'to get short name
'    Dim sql As String
'
'    Set tb = New Recordset

    If Trim$(txtSampleID) = "" Or Trim(txtName) = "" Then
        Exit Sub
    End If

    With frmPrintBarcodeLabel
'        Select Case ssTabAll.Tab
'            Case 1
'                .lblDepartment.Caption = "HAEM"
'            Case 2
'                .lblDepartment.Caption = "BIO"
'            Case 3
'                .lblDepartment.Caption = "COAG"
'            Case 4
'                .lblDepartment.Caption = "END"
'            Case 5
'                .lblDepartment.Caption = "BGA"
'            Case 6
'                .lblDepartment.Caption = "IMM"
'        End Select
        .lblDepartment = ssTabAll.Caption
        .lblSampleID.Caption = txtSampleID.Text
        .lblPatName.Caption = txtName
        .lblSampleDate.Caption = Format(dtSampleDate & " " & tSampleTime, "dd/MM/yyyy HH:mm:ss")
        .lblAgeSexDoB = "A/S/DOB: " & _
                    IIf(txtAge = "", "", Left(txtAge, Len(txtAge) - 2)) & " " & _
                    IIf(txtSex = "", "", Left(txtSex, 1)) & " " & _
                    IIf(txtDoB = "", "", Format(txtDoB, "yyyyMMdd"))
    End With
    
    
    '180           With tb
    '190               If (.EOF And .BOF) Or IsNull(!Requested) Then
    '200               Else
    '210                   .MoveFirst
    '220                   Do Until .EOF
    '230                       If Len(frmPrintBarcodeLabel.lblTestOrder.Caption) = 0 Then
    '240                           frmPrintBarcodeLabel.lblTestOrder.Caption = ShortNameFromLongName(!Requested & "")
    '250                       Else
    '260                           If Len(ShortNameFromLongName(!Requested & "")) > 0 Then
    '270                               frmPrintBarcodeLabel.lblTestOrder.Caption = frmPrintBarcodeLabel.lblTestOrder.Caption & "," & ShortNameFromLongName(!Requested & "") & ""
    '280                           End If
    '290                       End If
    '300                       .MoveNext
    '310                   Loop
    '320               End If
    '330           End With
    frmPrintBarcodeLabel.Show 1
    'End If


End Select

Exit Sub
cmdPatientNotePad_Click_Error:

LogError "frmEditAll", "cmdPatientNotePad_Click", Erl, Err.Description
End Sub

Private Sub cmdPhone_Click()

10  On Error GoTo cmdPhone_Click_Error

20  With frmPhoneLog
30      .SampleID = txtSampleID
40      .Caller = "General"
50      If cmbGP <> "" Then
60          .GP = cmbGP
70          .WardOrGP = "GP"
80      Else
90          .GP = cmbWard
100         .WardOrGP = "Ward"
110     End If
120     .Show 1
130 End With

140 CheckIfPhoned

150 Exit Sub

cmdPhone_Click_Error:

    Dim strES As String
    Dim intEL As Integer

160 intEL = Erl
170 strES = Err.Description
180 LogError "frmEditAll", "cmdPhone_Click", intEL, strES

End Sub

Private Sub cmdPhoresisComments_Click()

10  On Error GoTo cmdPhoresisComments_Click_Error

20  frmPhoresisComments.SampleID = txtSampleID
30  frmPhoresisComments.Show 1

40  Exit Sub

cmdPhoresisComments_Click_Error:

    Dim strES As String
    Dim intEL As Integer

50  intEL = Erl
60  strES = Err.Description
70  LogError "frmEditAll", "cmdPhoresisComments_Click", intEL, strES

End Sub

Private Sub cmdPrint_Click()

    Dim tb As New Recordset
    Dim sql As String

10  On Error GoTo cmdPrint_Click_Error

20  pBar = 0

30  txtSampleID = Format(Val(txtSampleID))

40  If Not EntriesOK(txtSampleID, txtName, txtSex, cmbWard.Text, cmbGP.Text) Then
50      Exit Sub
60  End If

70  If SysOptDemoVal(0) And cmdDemoVal.Caption = "&Validate" Then
80      If iMsg("Do you wish to validate demographics?", vbQuestion + vbYesNo) = vbNo Then
90          Exit Sub
100     Else
110         ValidateDemographics True
120     End If
130 End If

140 SaveDemographics

150 If SavePrintInhibit() Then
160     If ssTabAll.Tab = 1 Then
170         If lblHaemValid.Visible = False Then
180             SaveHaematology 1
190         End If
200         sql = "SELECT * from HaemResults WHERE " & _
                  "SampleID = '" & txtSampleID & "'"
210         Set tb = New Recordset
220         RecOpenClient 0, tb, sql
230         If Not tb.EOF Then
240             tb!Printed = 1
250             tb.Update
260         End If
270     ElseIf ssTabAll.Tab = 2 Then
280         If cmdSaveBio.Enabled = True Then SaveBiochemistry True
290         ValidateTests "Bio", gBio
300     ElseIf ssTabAll.Tab = 3 Then
310         If cmdSaveCoag.Enabled = True Then
320             SaveCoag False
330             cmdSaveCoag.Enabled = False
340         End If
350         If cmdValidateCoag.Caption = "&Validate" Then
360             ValidateTests "Coag", grdCoag
370         End If
380         sql = "UPDATE CoagResults " & _
                  "Set Printed = 0 WHERE " & _
                  "SampleID = '" & txtSampleID & "'"
390         Cnxn(0).Execute sql
400     ElseIf ssTabAll.Tab = 4 Then
410         ValidateTests "End", gImm(0)
420     ElseIf ssTabAll.Tab = 5 And cmdValBG.Caption = "&Validate" Then
430         sql = "UPDATE bgaResults " & _
                  "Set valid = 1 , operator = '" & AddTicks(UserCode) & "' WHERE " & _
                  "SampleID = '" & txtSampleID & "' and valid <> 1"
440         Cnxn(0).Execute sql
450     ElseIf ssTabAll.Tab = 6 Then
460         ValidateTests "Imm", gImm(1)
470     End If

480     If ssTabAll.Tab <> 0 Then
490         LogTimeOfPrinting txtSampleID, Choose(ssTabAll.Tab, "H", "B", "C", "E", "Q", "I", "")
500         sql = "SELECT * FROM PrintPending WHERE " & _
                  "Department = '" & Choose(ssTabAll.Tab, "H", "B", "C", "E", "Q", "I", "X") & "' " & _
                  "AND SampleID = '" & txtSampleID & "' " & _
                  "AND (FaxNumber = '' OR FaxNumber IS NULL)"
510         Set tb = New Recordset
520         RecOpenClient 0, tb, sql
530         If tb.EOF Then
540             tb.AddNew
550         End If
560         tb!SampleID = txtSampleID
570         tb!Department = Choose(ssTabAll.Tab, "H", "B", "C", "E", "Q", "I", "X")
580         If tb!Department = "I" And IsAllergy() Then tb!Department = "W"
590         If SysOptRealImm(0) And tb!Department = "I" Then tb!Department = "J"
600         tb!Initiator = UserName
610         tb!Ward = cmbWard
620         tb!Clinician = cmbClinician
630         tb!GP = cmbGP
640         tb!UsePrinter = pPrintToPrinter
650         tb!pTime = Now
660         tb.Update
670     End If




680     txtSampleID = Format$(Val(txtSampleID) + 1)
690     LoadAllDetails

700 End If

710 Exit Sub

cmdPrint_Click_Error:

    Dim strES As String
    Dim intEL As Integer

720 intEL = Erl
730 strES = Err.Description
740 LogError "frmEditAll", "cmdPrint_Click", intEL, strES, sql

End Sub

Private Sub cmdPrintAll_Click()

    Dim tb As New Recordset
    Dim sql As String

10  On Error GoTo cmdPrintAll_Click_Error

20  pBar = 0

30  txtSampleID = Format(Val(txtSampleID))
40  If Val(txtSampleID) = 0 Then Exit Sub

50  If Trim$(txtSex) = "" Then
60      If iMsg("Sex not entered." & vbCrLf & "Do you want to enter sex now?", vbQuestion + vbYesNo) = vbYes Then
70          Exit Sub
80      End If
90  End If

100 If Trim$(txtSampleID) = "" Then
110     iMsg "Must have Lab Number.", vbCritical
120     Exit Sub
130 End If

140 If Trim$(cmbWard) = "" Then
150     iMsg "Must have Ward entry.", vbCritical
160     Exit Sub
170 End If

180 If Trim$(cmbWard) = "GP" Then
190     If Trim$(cmbGP) = "" Then
200         iMsg "Must have Ward or GP entry.", vbCritical
210         Exit Sub
220     End If
230 End If

240 SaveDemographics

250 If ssTabAll.Tab <> 0 Then
260     sql = "SELECT * FROM PrintPending WHERE " & _
              "Department = 'D' " & _
              "AND SampleID = '" & txtSampleID & "'"
270     Set tb = New Recordset
280     RecOpenClient 0, tb, sql
290     If tb.EOF Then
300         tb.AddNew
310     End If
320     tb!SampleID = txtSampleID
330     tb!Ward = cmbWard
340     tb!Clinician = cmbClinician
350     tb!GP = cmbGP
360     tb!Department = "D"
370     tb!Initiator = UserName
380     tb!UsePrinter = pPrintToPrinter
390     tb.Update
400 End If

410 SaveCoag 1
420 sql = "UPDATE CoagResults " & _
          "Set Valid = 1, Printed = 1 WHERE " & _
          "SampleID = '" & txtSampleID & "'"
430 Cnxn(0).Execute sql

440 txtSampleID = Format$(Val(txtSampleID) + 1)
450 LoadAllDetails

460 Exit Sub

cmdPrintAll_Click_Error:

    Dim strES As String
    Dim intEL As Integer

470 intEL = Erl
480 strES = Err.Description
490 LogError "frmEditAll", "cmdPrintAll_Click", intEL, strES, sql

End Sub

Private Sub cmdPrintesr_Click()

    Dim sql As String
    Dim tb As New Recordset


10  On Error GoTo cmdPrintesr_Click_Error

20  txtSampleID = Format(Val(txtSampleID))
30  If Val(txtSampleID) = 0 Then Exit Sub

40  pBar = 0

50  If Trim$(txtSex) = "" Then
60      If iMsg("Sex not entered." & vbCrLf & "Do you want to enter sex now?", vbQuestion + vbYesNo) = vbYes Then
70          Exit Sub
80      End If
90  End If

100 If Trim$(txtSampleID) = "" Then
110     iMsg "Must have Lab Number.", vbCritical
120     Exit Sub
130 End If

140 If Trim$(cmbWard) = "" Then
150     iMsg "Must have Ward entry.", vbCritical
160     Exit Sub
170 End If

180 If Trim$(cmbWard) = "GP" Then
190     If Trim$(cmbGP) = "" Then
200         iMsg "Must have GP entry.", vbCritical
210         Exit Sub
220     End If
230 End If

240 SaveDemographics
250 SaveHaematology 1

    'PrintResultESRWin txtSampleID

260 sql = "SELECT * from HaemResults WHERE " & _
          "SampleID = '" & txtSampleID & "'"
270 Set tb = New Recordset
280 RecOpenClient 0, tb, sql

290 If Not tb.EOF Then
300     tb!Printed = 1
310     tb!Valid = 1
320     tb.Update
330 End If

340 Exit Sub

cmdPrintesr_Click_Error:

    Dim strES As String
    Dim intEL As Integer

350 intEL = Erl
360 strES = Err.Description
370 LogError "frmEditAll", "cmdPrintesr_Click", intEL, strES, sql

End Sub

Private Sub cmdPrintHold_Click()

    Dim sql As String
    Dim Department As String
    Dim NewDepartment As String

10  On Error GoTo cmdPrintHold_Click_Error

20  pBar = 0

30  txtSampleID = Format(Val(txtSampleID))
40  If Val(txtSampleID) = 0 Then Exit Sub

50  If Trim$(txtSex) = "" Then
60      If iMsg("Sex not entered." & vbCrLf & "Do you want to enter sex now?", vbQuestion + vbYesNo) = vbYes Then
70          Exit Sub
80      End If
90  End If

100 If Trim$(txtSampleID) = "" Then
110     iMsg "Must have Lab Number.", vbCritical
120     Exit Sub
130 End If

140 If Len(cmbWard) = 0 Then
150     iMsg "Must have Ward entry.", vbCritical
160     Exit Sub
170 End If

180 If Trim$(cmbWard) = "GP" Then
190     If Len(cmbGP) = 0 Then
200         iMsg "Must have Ward or GP entry.", vbCritical
210         Exit Sub
220     End If
230 End If

240 If SysOptDemoVal(0) And cmdDemoVal.Caption = "&Validate" Then
250     If iMsg("Do you wish to validate demographics?", vbQuestion + vbYesNo) = vbNo Then
260         Exit Sub
270     Else
280         ValidateDemographics True
290     End If
300 End If

310 SaveDemographics

320 If SavePrintInhibit() Then

330     If ssTabAll.Tab = 1 Then
340         If lblHaemValid.Visible = False Then
350             SaveHaematology 1
360         End If
370         sql = "IF EXISTS (SELECT * from HaemResults WHERE " & _
                "           SampleID = '" & txtSampleID & "') " & _
                "  UPDATE HaemResults " & _
                "  SET Printed = 1 " & _
                "  WHERE SampleID = '" & txtSampleID & "'"
380         Cnxn(0).Execute sql
390         LoadHaematology
400     ElseIf ssTabAll.Tab = 2 Then
410         If cmdSaveBio.Enabled = True Then SaveBiochemistry True
420         ValidateTests "Bio", gBio
430         LoadBiochemistry
440     ElseIf ssTabAll.Tab = 3 Then
450         If cmdSaveCoag.Enabled = True Then
460             SaveCoag False
470             cmdSaveCoag.Enabled = False
480         End If
490         If cmdValidateCoag.Caption = "&Validate" Then
500             ValidateTests "Coag", grdCoag
510             LoadCoagulation
520         End If

530     ElseIf ssTabAll.Tab = 4 Then
540         ValidateTests "End", gImm(0)
550         LoadEndocrinology
560     ElseIf ssTabAll.Tab = 5 Then
570         sql = "UPDATE BgaResults " & _
                  "Set valid = 1 WHERE " & _
                  "SampleID = '" & txtSampleID & "'"
580         Cnxn(0).Execute sql
590         LoadBloodGas
600     ElseIf ssTabAll.Tab = 6 Then
610         ValidateTests "Imm", gImm(1)
620         LoadImmunology
630     End If

640     If ssTabAll.Tab <> 0 Then
650         Department = Choose(ssTabAll.Tab, "H", "B", "C", "E", "Q", "I", "X")
660         If Department = "I" And IsAllergy() Then Department = "W"
670         If SysOptRealImm(0) And Department = "I" Then
680             NewDepartment = "J"
690         Else
700             NewDepartment = Department
710         End If

720         LogTimeOfPrinting txtSampleID, Department
730         sql = "IF EXISTS (SELECT * FROM PrintPending WHERE " & _
                "           Department = '" & Department & "' " & _
                "           AND SampleID = '" & txtSampleID & "' " & _
                "           AND COALESCE(FaxNumber, '') = '') " & _
                "    UPDATE PrintPending " & _
                "    SET Department = '" & NewDepartment & "', " & _
                "    Initiator = '" & UserName & "', " & _
                "    Ward = '" & AddTicks(cmbWard) & "', " & _
                "    Clinician = '" & AddTicks(cmbClinician) & "', " & _
                "    GP = '" & AddTicks(cmbGP) & "', " & _
                "    UsePrinter = '" & pPrintToPrinter & "', " & _
                "    pTime = getdate() " & _
                "    WHERE Department = '" & Department & "' " & _
                "    AND SampleID = '" & txtSampleID & "' " & _
                "    AND COALESCE(FaxNumber, '') = '' " & _
                  "ELSE " & _
                "    INSERT INTO PrintPending " & _
                    "    (SampleID, Department, Initiator, Ward, Clinician, GP, UsePrinter, pTime) "
740         sql = sql & _
                "    VALUES ( " & _
                "    '" & txtSampleID & "', " & _
                "    '" & NewDepartment & "', " & _
                "    '" & UserName & "', " & _
                "    '" & AddTicks(cmbWard) & "', " & _
                "    '" & AddTicks(cmbClinician) & "', " & _
                "    '" & AddTicks(cmbGP) & "', " & _
                "    '" & pPrintToPrinter & "', " & _
                "    getdate() )"
750         Cnxn(0).Execute sql
760     End If


770 End If

780 Exit Sub

cmdPrintHold_Click_Error:

    Dim strES As String
    Dim intEL As Integer

790 intEL = Erl
800 strES = Err.Description
810 LogError "frmEditAll", "cmdPrintHold_Click", intEL, strES, sql

End Sub

Private Sub cmdResend_Click()

      Dim sql As String
      Dim Disp As String

10    On Error GoTo cmdResend_Click_Error

20    Select Case ssTabAll.Tab
          Case 0, 1, 3, 5, 7:
30            Disp = ""
40        Case 2:
50            Disp = "Biochemistry"
60        Case 4:
70            Disp = "Endocrinology"
80        Case 6:
90            Disp = "Immunology"
100   End Select

110   sql = "UPDATE LabLinkCommunication SET MessageState = 3, Status = 'Request Received' WHERE SampleID = " & txtSampleID & " And Department = '" & Disp & "'"
120   Cnxn(0).Execute sql
130   cmdResend.Visible = False

140   Exit Sub

cmdResend_Click_Error:

       Dim strES As String
       Dim intEL As Integer

150    intEL = Erl
160    strES = Err.Description
170    LogError "frmEditAll", "cmdResend_Click", intEL, strES, sql
          
End Sub

Private Sub cmdSaveBGa_Click()

10  On Error GoTo cmdSaveBGa_Click_Error

20  pBar = 0

30  txtSampleID = Format(Val(txtSampleID))
40  If Val(txtSampleID) = 0 Then Exit Sub

50  If cmdValBG.Caption = "&Validate" Then
60      SaveBloodGas False
70  Else
80      SaveBloodGas True
90  End If
100 SaveComments
110 UPDATEMRU txtSampleID, cMRU

120 cmdSaveBGa.Enabled = False

130 Exit Sub

cmdSaveBGa_Click_Error:

    Dim strES As String
    Dim intEL As Integer

140 intEL = Erl
150 strES = Err.Description
160 LogError "frmEditAll", "cmdSaveBGa_Click", intEL, strES

End Sub

Private Sub cmdSaveBio_Click()

10  On Error GoTo cmdSaveBio_Click_Error

20  pBar = 0

30  txtSampleID = Format(Val(txtSampleID))
40  If Val(txtSampleID) = 0 Then Exit Sub

50  If bValidateBio.Caption = "&Validate" Then
60      SaveBiochemistry False
70  Else
80      SaveBiochemistry True
90  End If
100 SaveComments
110 UPDATEMRU txtSampleID, cMRU

120 cmdSaveBio.Enabled = False

130 Exit Sub

cmdSaveBio_Click_Error:

    Dim strES As String
    Dim intEL As Integer

140 intEL = Erl
150 strES = Err.Description
160 LogError "frmEditAll", "cmdSaveBio_Click", intEL, strES

End Sub

Private Sub cmdSaveCoag_Click()

10  On Error GoTo cmdSaveCoag_Click_Error

20  pBar = 0

30  txtSampleID = Format(Val(txtSampleID))
40  If Val(txtSampleID) = 0 Then Exit Sub

    'added 15/Jul/2004
50  If cmdValidateCoag.Caption = "&Validate" Then
60      SaveCoag 0
70  Else
80      SaveCoag 1
90  End If

100 SaveComments
110 UPDATEMRU txtSampleID, cMRU

120 cmdSaveCoag.Enabled = False

130 Exit Sub

cmdSaveCoag_Click_Error:

    Dim strES As String
    Dim intEL As Integer

140 intEL = Erl
150 strES = Err.Description
160 LogError "frmEditAll", "cmdSaveCoag_Click", intEL, strES

End Sub

Private Sub cmdSaveComm_Click()
    Dim sql As String

10  On Error GoTo cmdSaveComm_Click_Error

20  sql = "UPDATE HaemResults " & _
          "SET HealthLink = 0 WHERE " & _
          "SampleID = '" & txtSampleID & "'"
30  Cnxn(0).Execute sql

40  SaveComments

50  Exit Sub

cmdSaveComm_Click_Error:

    Dim strES As String
    Dim intEL As Integer

60  intEL = Erl
70  strES = Err.Description
80  LogError "frmEditAll", "cmdSaveComm_Click", intEL, strES, sql

End Sub

Private Sub cmdSaveDemographics_Click()

10  On Error GoTo cmdSaveDemographics_Click_Error

20  pBar = 0

30  txtSampleID = Format(Val(txtSampleID))
40  If Val(txtSampleID) = 0 Then Exit Sub

50  If Trim$(txtSex) = "" Then
60      If iMsg("Sex not entered." & vbCrLf & "Do you want to enter sex now?", vbQuestion + vbYesNo) = vbYes Then
70          Exit Sub
80      End If
90  End If

100 If Trim$(txtSampleID) = "" Then
110     iMsg "Must have Lab Number.", vbCritical
120     Exit Sub
130 End If

140 If Trim$(txtName) <> "" Then
150     If Trim$(cmbWard) = "" Then
160         iMsg "Must have Ward entry.", vbCritical
170         Exit Sub
180     End If

190     If Trim$(cmbWard) = "GP" Then
200         If Trim$(cmbGP) = "" Then
210             iMsg "Must have GP entry.", vbCritical
220             Exit Sub
230         End If
240     End If
250 End If

260 If dtRunDate < dtSampleDate Then
270     iMsg "Sample Date After Run Date. Please Amend!"
280     Exit Sub
290 End If

300 If dtRunDate < dtRecDate Then
310     iMsg "Rec. Date After Run Date. Please Amend!"
320     Exit Sub
330 End If

340 If dtRecDate < dtSampleDate Then
350     iMsg "Sample Date After Rec. Date. Please Amend!"
360     Exit Sub
370 End If

380 If Format(dtRunDate, "dd/MM/yyyy") <> Format(Now, "dd/MM/yyyy") Then
390     If iMsg("Rundate not today. Proceed ?", vbYesNo) = vbNo Then
400         Exit Sub
410     End If
420 End If

430 cmdSaveDemographics.Caption = "Saving"

440 SaveDemographics
450 UPDATEMRU txtSampleID, cMRU
460 LoadDemographics
470 cmdSaveDemographics.Caption = "Save && &Hold"
480 cmdSaveDemographics.Enabled = False
490 cmdSaveInc.Enabled = False

500 If txtSampleID.Visible And txtSampleID.Enabled Then
510     txtSampleID.SetFocus
520 End If

530 Exit Sub

cmdSaveDemographics_Click_Error:

    Dim strES As String
    Dim intEL As Integer

540 intEL = Erl
550 strES = Err.Description
560 LogError "frmEditAll", "cmdSaveDemographics_Click", intEL, strES

End Sub

'Private Sub cmdSaveExt_Click()
'
'    On Error GoTo cmdSaveExt_Click_Error
'
'    txtSampleID = Format(Val(txtSampleID))
'    If Val(txtSampleID) = 0 Then Exit Sub
'    UpDown1.Enabled = True
'
'    SaveExtern False
'    UPDATEMRU txtSampleID, cMRU
'    cmdSaveExt.Enabled = False
'
'    LoadExt
'
'    Exit Sub
'
'cmdSaveExt_Click_Error:
'
'    Dim strES As String
'    Dim intEL As Integer
'
'    intEL = Erl
'    strES = Err.Description
'    LogError "frmEditAll", "cmdSaveExt_Click", intEL, strES
'
'End Sub

Private Sub cmdSaveHaem_Click()



10  On Error GoTo cmdSaveHaem_Click_Error

20  pBar = 0

30  txtSampleID = Format(Val(txtSampleID))
40  If Val(txtSampleID) = 0 Then Exit Sub

    'Added 15/Jul/2004

50  If bValidateHaem.Caption = "&Validate" Then
60      SaveHaematology 0
70  Else
80      SaveHaematology 1
90  End If

100 SaveComments
110 UPDATEMRU txtSampleID, cMRU

120 txtSampleID = Format$(Val(txtSampleID) + 1)
130 LoadAllDetails

140 cmdSaveHaem.Enabled = False
150 cmdHSaveH.Enabled = False


160 Exit Sub

cmdSaveHaem_Click_Error:

    Dim strES As String
    Dim intEL As Integer

170 intEL = Erl
180 strES = Err.Description
190 LogError "frmEditAll", "cmdSaveHaem_Click", intEL, strES

End Sub

Private Sub cmdSaveImm_Click(Index As Integer)



10  On Error GoTo cmdSaveImm_Click_Error

20  pBar = 0


30  If Index = 0 Then
        'Endocrinology

40      txtSampleID = Format(Val(txtSampleID))
50      If Val(txtSampleID) = 0 Then Exit Sub

        'added 15/Jul/2004
60      If bValidateImm(0).Caption = "&Validate" Then
70          SaveEndocrinology False
80      Else
90          SaveEndocrinology True
100     End If
110     SaveComments
120     UPDATEMRU txtSampleID, cMRU

130     cmdSaveImm(0).Enabled = False
140 ElseIf Index = 1 Then
        'Immunology
150     txtSampleID = Format(Val(txtSampleID))
160     If Val(txtSampleID) = 0 Then Exit Sub

        'added 15/Jul/2004
170     If bValidateImm(1).Caption = "&Validate" Then
180         SaveImmunology False
190     Else
200         SaveImmunology True
210     End If
220     SaveComments
230     UPDATEMRU txtSampleID, cMRU

240     cmdSaveImm(1).Enabled = False
250 ElseIf Index = 2 Then
        'External
260     txtSampleID = Format(Val(txtSampleID))
270     If Val(txtSampleID) = 0 Then Exit Sub
280     UpDown1.Enabled = True

290     SaveExtern False
300     UPDATEMRU txtSampleID, cMRU
310     cmdSaveImm(2).Enabled = False

320     LoadExt
330 End If

340 Exit Sub

cmdSaveImm_Click_Error:

    Dim strES As String
    Dim intEL As Integer

350 intEL = Erl
360 strES = Err.Description
370 LogError "frmEditAll", "cmdSaveImm_Click", intEL, strES

End Sub

Private Sub cmdSaveInc_Click()

10  On Error GoTo cmdSaveInc_Click_Error

20  pBar = 0

30  txtSampleID = Format(Val(txtSampleID))

40  If Not EntriesOK(txtSampleID, txtName, txtSex, cmbWard.Text, cmbGP.Text) Then
50      Exit Sub
60  End If

70  If lblChartNumber.BackColor = vbRed And Trim(txtChart) <> "" Then
80      If iMsg("Confirm this Patient has" & vbCrLf & _
                lblChartNumber.Caption, vbYesNo + vbQuestion, , vbRed, 12) = vbNo Then
90          Exit Sub
100     End If
110 End If

120 If dtRunDate < dtSampleDate Then
130     iMsg "Sample Date After Run Date. Please Amend!"
140     Exit Sub
150 End If

160 If dtRunDate < dtRecDate Then
170     iMsg "Rec. Date After Run Date. Please Amend!"
180     Exit Sub
190 End If

200 If dtRecDate < dtSampleDate Then
210     iMsg "Sample Date After Rec. Date. Please Amend!"
220     Exit Sub
230 End If

240 If Format(dtRunDate, "dd/MMM/yyyy") <> Format(Now, "dd/MMM/yyyy") Then
250     If iMsg("Rundate not today. Proceed ?", vbYesNo) = vbNo Then
260         Exit Sub
270     End If
280 End If

290 cmdSaveDemographics.Caption = "Saving"

300 SaveDemographics
310 UPDATEMRU txtSampleID, cMRU

320 cmdSaveDemographics.Caption = "Save && &Hold"
330 cmdSaveDemographics.Enabled = False
340 cmdSaveInc.Enabled = False

350 txtSampleID = Format$(Val(txtSampleID) + 1)

360 LoadAllDetails

370 cmdSaveHaem.Enabled = False
380 cmdHSaveH.Enabled = False
390 cmdSaveBio.Enabled = False
400 cmdSaveCoag.Enabled = False
410 cmdSaveImm(0).Enabled = False
420 cmdSaveImm(1).Enabled = False
430 cmdSaveBGa.Enabled = False

440 txtSampleID.SelStart = 0
450 txtSampleID.SelLength = Len(txtSampleID)
460 txtSampleID.SetFocus

470 Exit Sub

cmdSaveInc_Click_Error:

    Dim strES As String
    Dim intEL As Integer

480 intEL = Erl
490 strES = Err.Description
500 LogError "frmEditAll", "cmdSaveInc_Click", intEL, strES

End Sub

Private Sub cmdSetPrinter_Click()

10  On Error GoTo cmdSetPrinter_Click_Error

20  Set frmForcePrinter.f = frmEditAll
30  frmForcePrinter.Show 1

40  If pPrintToPrinter = "Automatic Selection" Then
50      pPrintToPrinter = ""
60  End If

70  If pPrintToPrinter <> "" Then
80      cmdSetPrinter.BackColor = vbRed
90      cmdSetPrinter.ToolTipText = "Print Forced to " & pPrintToPrinter
100 Else
110     cmdSetPrinter.BackColor = vbButtonFace
120     pPrintToPrinter = ""
130     cmdSetPrinter.ToolTipText = "Printer Selected Automatically"
140 End If

150 Exit Sub

cmdSetPrinter_Click_Error:

    Dim strES As String
    Dim intEL As Integer

160 intEL = Erl
170 strES = Err.Description
180 LogError "frmEditAll", "cmdSetPrinter_Click", intEL, strES

End Sub

Private Sub cmdUnvalPrint_Click()

    Dim tb As New Recordset
    Dim sql As String

10  On Error GoTo cmdUnvalPrint_Click_Error

20  pBar = 0

30  txtSampleID = Format(Val(txtSampleID))
40  If Val(txtSampleID) = 0 Then Exit Sub

50  If Trim$(txtSex) = "" Then
60      If iMsg("Sex not entered." & vbCrLf & "Do you want to enter sex now?", vbQuestion + vbYesNo) = vbYes Then
70          Exit Sub
80      End If
90  End If

100 If Trim$(txtSampleID) = "" Then
110     iMsg "Must have Lab Number.", vbCritical
120     Exit Sub
130 End If


140 If Len(cmbWard) = 0 Then
150     iMsg "Must have Ward entry.", vbCritical
160     Exit Sub
170 End If

180 If Trim$(cmbWard) = "GP" Then
190     If Len(cmbGP) = 0 Then
200         iMsg "Must have Ward or GP entry.", vbCritical
210         Exit Sub
220     End If
230 End If

240 sql = "SELECT * FROM PrintPending WHERE " & _
          "Department = 'K' " & _
          "AND SampleID = '" & txtSampleID & "'"
250 Set tb = New Recordset
260 RecOpenClient 0, tb, sql
270 If tb.EOF Then
280     tb.AddNew
290 End If
300 tb!SampleID = txtSampleID
310 tb!Department = "K"
320 tb!Initiator = UserName
330 tb!Ward = cmbWard
340 tb!Clinician = cmbClinician
350 tb!GP = cmbGP
360 tb!UsePrinter = pPrintToPrinter
370 tb!pTime = Now
380 tb.Update

390 Exit Sub

cmdUnvalPrint_Click_Error:

    Dim strES As String
    Dim intEL As Integer

400 intEL = Erl
410 strES = Err.Description
420 LogError "frmEditAll", "cmdUnvalPrint_Click", intEL, strES

End Sub

Private Sub cmdValBG_Click()

    Dim sql As String

10  On Error GoTo cmdValBG_Click_Error

20  If cmdDemoVal.Caption = "&Validate" Then
30      If iMsg("Do you wish to validate demographics?", vbQuestion + vbYesNo) = vbNo Then
40          Exit Sub
50      Else
60          ValidateDemographics True
70      End If
80  End If

    'If Trim(txtDoB) = "" Then
    '    iMsg "No Date of Birth specified." & vbCrLf & "Adult Age 25 used for Normal Ranges!", vbInformation
    'End If

90  sql = "UPDATE BgaResults SET Valid = '1' WHERE SampleID = '" & txtSampleID & "'"
100 Cnxn(0).Execute sql

110 sql = "UPDATE Demographics SET ForBGA = '1' WHERE SampleID = '" & txtSampleID & "'"
120 Cnxn(0).Execute sql

130 cmdValBG.Caption = "VALID"

140 Exit Sub

cmdValBG_Click_Error:

    Dim strES As String
    Dim intEL As Integer

150 intEL = Erl
160 strES = Err.Description
170 LogError "frmEditAll", "cmdValBG_Click", intEL, strES, sql

End Sub

Private Sub cmdValidateCoag_Click()

10  On Error GoTo cmdValidateCoag_Click_Error

20  pBar = 0

30  txtSampleID = Format(Val(txtSampleID))
40  If Val(txtSampleID) = 0 Then Exit Sub

50  If cmdValidateCoag.Caption = "VALID" Then
60      If UCase(iBOX("Unvalidate ! Enter Password", , , True)) = UserPass Then
70          SaveCoag False
80          SaveComments
            'txtCoagComment.Locked = False
90          cmdValidateCoag.Caption = "&Validate"
100         Me.Refresh
110     End If
120 Else
130     If cmdDemoVal.Caption = "&Validate" Then
140         If iMsg("Do you wish to validate demographics !", vbYesNo) = vbNo Then
150             Exit Sub
160         Else
170             ValidateDemographics True
180         End If
190     End If
        '    If Trim(txtDoB) = "" Then
        '        iMsg "No Date of Birth Specified." & vbCrLf & "Adult Age 25 used for Normal Ranges!", vbInformation
        '    End If
200     SaveCoag True
210     SaveComments
220     UPDATEMRU txtSampleID, cMRU
        'txtCoagComment.Locked = True
230     txtSampleID = Format(Val(txtSampleID)) + 1
240     Me.Refresh
250 End If

260 LoadAllDetails

270 Exit Sub

cmdValidateCoag_Click_Error:

    Dim strES As String
    Dim intEL As Integer

280 intEL = Erl
290 strES = Err.Description
300 LogError "frmEditAll", "cmdValidateCoag_Click", intEL, strES

End Sub

Private Sub cmdViewBioReps_Click()

10  On Error GoTo cmdViewBioReps_Click_Error

20  frmRFT.SampleID = Val(txtSampleID)
30  frmRFT.Dept = "B"
40  frmRFT.Show 1

50  Exit Sub

cmdViewBioReps_Click_Error:

    Dim strES As String
    Dim intEL As Integer

60  intEL = Erl
70  strES = Err.Description
80  LogError "frmEditAll", "cmdViewBioReps_Click", intEL, strES

End Sub

Private Sub cmdViewCoagRep_Click()

10  On Error GoTo cmdViewCoagRep_Click_Error

20  frmRFT.SampleID = Val(txtSampleID)
30  frmRFT.Dept = "C"
40  frmRFT.Show 1

50  Exit Sub

cmdViewCoagRep_Click_Error:

    Dim strES As String
    Dim intEL As Integer

60  intEL = Erl
70  strES = Err.Description
80  LogError "frmEditAll", "cmdViewCoagRep_Click", intEL, strES

End Sub

Private Sub cmdViewExtReport_Click()

10  On Error GoTo cmdViewExtReport_Click_Error

20  frmRFT.SampleID = Val(txtSampleID)
30  frmRFT.Dept = "X"
40  frmRFT.Show 1

50  Exit Sub

cmdViewExtReport_Click_Error:

    Dim strES As String
    Dim intEL As Integer

60  intEL = Erl
70  strES = Err.Description
80  LogError "frmEditAll", "cmdViewExtReport_Click", intEL, strES

End Sub

Private Sub cmdViewHaemRep_Click()

10  On Error GoTo cmdViewHaemRep_Click_Error

20  frmRFT.SampleID = Val(txtSampleID)
30  frmRFT.Dept = "H"
40  frmRFT.Show 1

50  Exit Sub

cmdViewHaemRep_Click_Error:

    Dim strES As String
    Dim intEL As Integer

60  intEL = Erl
70  strES = Err.Description
80  LogError "frmEditAll", "cmdViewHaemRep_Click", intEL, strES

End Sub

Private Sub cmdViewImmRep_Click()

10  On Error GoTo cmdViewImmRep_Click_Error

20  frmRFT.SampleID = Val(txtSampleID)
30  frmRFT.Dept = "I"
40  frmRFT.Show 1

50  Exit Sub

cmdViewImmRep_Click_Error:

    Dim strES As String
    Dim intEL As Integer

60  intEL = Erl
70  strES = Err.Description
80  LogError "frmEditAll", "cmdViewImmRep_Click", intEL, strES

End Sub

Private Sub cmdViewReports_Click()

10  On Error GoTo cmdViewReports_Click_Error

20  frmRFT.SampleID = Val(txtSampleID)
30  frmRFT.Dept = "E"
40  frmRFT.Show 1

50  Exit Sub

cmdViewReports_Click_Error:

    Dim strES As String
    Dim intEL As Integer

60  intEL = Erl
70  strES = Err.Description
80  LogError "frmEditAll", "cmdViewReports_Click", intEL, strES

End Sub

Private Sub cMonospot_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

10  On Error GoTo cMonospot_MouseUp_Error

20  If cMonospot = 0 Then
30      If Trim$(tMonospot) = "?" Then
40          tMonospot = ""
50      ElseIf Trim$(tMonospot) <> "" Then
60          cMonospot = 1
70      End If
80  Else
90      If Trim$(tMonospot) = "" Then
100         tMonospot = "?"
110     End If
120 End If

130 cmdSaveHaem.Enabled = True: cmdHSaveH.Enabled = True

140 Exit Sub

cMonospot_MouseUp_Error:

    Dim strES As String
    Dim intEL As Integer

150 intEL = Erl
160 strES = Err.Description
170 LogError "frmEditAll", "cMonospot_MouseUp", intEL, strES

End Sub

Private Sub cMRU_Click()

10  On Error GoTo cMRU_Click_Error

20  txtSampleID = cMRU

30  LoadAllDetails

40  cmdSaveDemographics.Enabled = False
50  cmdSaveInc.Enabled = False
60  cmdSaveHaem.Enabled = False
70  cmdSaveComm.Enabled = False
80  cmdHSaveH.Enabled = False
90  cmdSaveBio.Enabled = False
100 cmdSaveCoag.Enabled = False
110 cmdSaveImm(0).Enabled = False
120 cmdSaveImm(1).Enabled = False
130 cmdSaveBGa.Enabled = False

140 Exit Sub

cMRU_Click_Error:

    Dim strES As String
    Dim intEL As Integer

150 intEL = Erl
160 strES = Err.Description
170 LogError "frmEditAll", "cMRU_Click", intEL, strES

End Sub

Private Sub cMRU_KeyPress(KeyAscii As Integer)

10  KeyAscii = 0

End Sub

Private Sub Colourise(ByVal Analyte As String, _
                      ByVal Destination As TextBox, _
                      ByVal strValue As String, _
                      ByVal sex As String, _
                      ByVal Dob As String)

    Dim Value As Single
    Dim sql As String
    Dim tb As Recordset
    Dim x As Long


10  On Error GoTo Colourise_Error


20  Value = Val(strValue)

30  If InStr(strValue, ">") Then
40      x = InStr(strValue, ">")
50      Value = Mid(strValue, x + 1)
60  End If

70  sql = "SELECT PrintFormat FROM HaemTestDefinitions WHERE " & _
          "AnalyteName = '" & Analyte & "'"
80  Set tb = New Recordset
90  RecOpenServer 0, tb, sql
100 If Not tb.EOF Then
110     Select Case Val(tb!Printformat & "")
        Case 0:
120         Destination = strValue
130     Case 1:
140         Destination = Format(strValue, "##0.0")
150     Case 2:
160         Destination = Format(strValue, "##0.00")
170     Case 3:
180         Destination = Format(strValue, "##0.000")
190     End Select
200 Else
210     Destination = strValue
220 End If

230 If Trim$(strValue) = "" Then
240     Destination.BackColor = &HFFFFFF
250     Destination.ForeColor = &H0&
260     Exit Sub
270 End If


280 If sex <> "" Then        'QMS reference number #817982
290     Select Case InterpH(Value, Analyte, sex, Dob, 0, lblSampleDate)
        Case "X":
300         Destination.BackColor = SysOptPlasBack(0)
310         Destination.ForeColor = SysOptPlasFore(0)
320     Case "H":
330         Destination.BackColor = SysOptHighBack(0)
340         Destination.ForeColor = SysOptHighFore(0)
350     Case "L"
360         Destination.BackColor = SysOptLowBack(0)
370         Destination.ForeColor = SysOptLowFore(0)
380     Case Else
390         Destination.BackColor = &HFFFFFF
400         Destination.ForeColor = &H0&
410     End Select
420 End If
430 Exit Sub

Colourise_Error:

    Dim strES As String
    Dim intEL As Integer

440 intEL = Erl
450 strES = Err.Description
460 LogError "frmEditAll", "Colourise", intEL, strES, sql

End Sub

Private Sub ColouriseG(ByVal Analyte As String, _
                       ByVal Destination As MSFlexGrid, _
                       ByVal x As Long, _
                       ByVal Y As Long, _
                       ByVal strValue As String, _
                       ByVal sex As String, _
                       ByVal Dob As String)

    Dim Value As Single
    Dim sql As String
    Dim tb As Recordset

10  On Error GoTo ColouriseG_Error

20  Value = Trim(Val(strValue))

30  sql = "SELECT * from haemtestdefinitions WHERE analytename = '" & Analyte & "'"
40  Set tb = New Recordset
50  RecOpenServer 0, tb, sql
60  If Not tb.EOF Then
70      Select Case Val(tb!Printformat & "")
        Case 0:
80          Destination.TextMatrix(x, Y) = strValue
90      Case 1:
100         Destination.TextMatrix(x, Y) = Format(strValue, "##0.0")
110     Case 2:
120         Destination.TextMatrix(x, Y) = Format(strValue, "##0.00")
130     Case 3:
140         Destination.TextMatrix(x, Y) = Format(strValue, "##0.000")
150     End Select
160 Else
170     Destination.TextMatrix(x, Y) = strValue
180 End If

190 Destination.Col = Y
200 Destination.Row = x

210 If Trim$(strValue) = "" Then
220     Destination.CellBackColor = &HFFFFFF
230     Destination.CellForeColor = 1
240     Exit Sub
250 End If

260 Select Case InterpH(Value, Analyte, sex, Dob, 0, Format(lblSampleDate, "dd/MMM/yyyy"))
    Case "X":
270     Destination.CellBackColor = SysOptPlasBack(0)
280     Destination.CellForeColor = SysOptPlasFore(0)
290 Case "H":
300     Destination.CellBackColor = SysOptHighBack(0)
310     Destination.CellForeColor = SysOptHighFore(0)
320 Case "L"
330     Destination.CellBackColor = SysOptLowBack(0)
340     Destination.CellForeColor = SysOptLowFore(0)
350 Case Else
360     Destination.CellBackColor = &HFFFFFF
370     Destination.CellForeColor = 1
380 End Select

390 Exit Sub

ColouriseG_Error:

    Dim strES As String
    Dim intEL As Integer

400 intEL = Erl
410 strES = Err.Description
420 LogError "frmEditAll", "ColouriseG", intEL, strES, sql

End Sub



Private Sub cParameter_Click()


10  On Error GoTo cParameter_Click_Error

20  pBar = 0
    Dim n As Long
    Dim Unit As String

30  cCunits.Enabled = True

    Dim SampleType As String
    Dim Code As String

40  SampleType = ListCodeFor("ST", cISampleType(3))
50  Code = CoagCodeFor(cParameter)
60  Unit = ACoagUnitsFor(cParameter)

70  If Unit <> "" Then
80      For n = 1 To cCunits.ListCount
90          If cCunits.List(n) = Trim(Unit) Then
100             cCunits.ListIndex = n
110             Exit For
120         End If
130     Next
140 End If

150 If cParameter = "PT" Then cCunits.Enabled = True Else cCunits.Enabled = False

160 Exit Sub

cParameter_Click_Error:

    Dim strES As String
    Dim intEL As Integer

170 intEL = Erl
180 strES = Err.Description
190 LogError "frmEditAll", "cParameter_Click", intEL, strES

End Sub

Private Sub cParameter_KeyPress(KeyAscii As Integer)

10  On Error GoTo cParameter_KeyPress_Error

20  KeyAscii = AutoComplete(cParameter, KeyAscii, False)

30  Exit Sub

cParameter_KeyPress_Error:

    Dim strES As String
    Dim intEL As Integer

40  intEL = Erl
50  strES = Err.Description
60  LogError "frmEditAll", "cParameter_KeyPress", intEL, strES

End Sub

Private Sub cParameter_LostFocus()

10  On Error GoTo cParameter_LostFocus_Error

20  cParameter.Text = QueryCombo(cParameter)

30  Exit Sub

cParameter_LostFocus_Error:

    Dim strES As String
    Dim intEL As Integer

40  intEL = Erl
50  strES = Err.Description
60  LogError "frmEditAll", "cParameter_LostFocus", intEL, strES

End Sub

Private Sub cRA_Click()

10  On Error GoTo cRA_Click_Error

20  If cRA = 0 Then
30      If Trim$(tRa) = "?" Then
40          tRa = ""
50      ElseIf Trim$(tRa) <> "" Then
60          cRA = 1
70      End If
80  Else
90      If Trim$(tRa) = "" Then
100         tRa = "?"
110     End If
120 End If

130 cmdSaveHaem.Enabled = True: cmdHSaveH.Enabled = True

140 Exit Sub

cRA_Click_Error:

    Dim strES As String
    Dim intEL As Integer

150 intEL = Erl
160 strES = Err.Description
170 LogError "frmEditAll", "cRA_Click", intEL, strES

End Sub

Private Function CreateHist(ByVal Dept As String) As String

    Dim sql As String

10  On Error GoTo CreateHist_Error

20  If Trim(txtSampleID) = "" Then Exit Function

30  sql = "SELECT TOP 1 D.SampleID, D.RunDate, D.SampleDate " & _
          "FROM Demographics AS D, " & Dept & "Results AS R WHERE "

40  If Trim(txtChart) <> "" Then sql = sql & "D.Chart = '" & AddTicks(txtChart) & "' AND "
50  If Trim(txtAandE) <> "" And UCase(HospName(0)) = "MULLINGAR" Then sql = sql & "D.AandE = '" & AddTicks(txtAandE) & "' AND "

60  sql = sql & "D.PatName = '" & AddTicks(txtName) & "' " & _
          "AND D.DOB = '" & Format(txtDoB, "dd/MMM/yyyy") & "' AND "

70  If SysOptHistView(0) = False Then
80      sql = sql & "D.SampleID < '" & txtSampleID & "' AND R.SampleID = D.SampleID " & _
              "ORDER BY D.SampleID desc"
90  Else
100     sql = sql & " D.SampleID <> '" & txtSampleID & "' AND R.SampleID = D.SampleID " & _
              "ORDER BY D.SampleDate desc"
110 End If

120 CreateHist = sql

130 Exit Function

CreateHist_Error:

    Dim strES As String
    Dim intEL As Integer

140 intEL = Erl
150 strES = Err.Description
160 LogError "frmEditAll", "CreateHist", intEL, strES, sql

End Function

Private Function CreateSql(ByVal Dept As String) As String

    Dim sql As String
    Dim SampleDate As String

10  On Error GoTo CreateSql_Error

20  SampleDate = Format$(dtSampleDate, "dd/MMM/yyyy")
30  If IsDate(tSampleTime) Then
40      SampleDate = SampleDate & " " & Format$(tSampleTime, "HH:nn")
50  End If

60  sql = "SELECT TOP 1 D.SampleID, D.RunDate, D.SampleDate FROM Demographics AS D, " & Dept & "Results AS R WHERE ("

70  If Trim(txtChart) <> "" Then
80      sql = sql & "D.Chart = '" & AddTicks(txtChart) & "' AND "
90  End If
100 If Trim(txtAandE) <> "" And UCase(HospName(0)) = "MULLINGAR" Then
110     sql = sql & "D.AandE = '" & AddTicks(txtAandE) & "' AND"
120 End If

130 sql = sql & " D.PatName = '" & AddTicks(txtName) & "' " & _
          "AND D.DOB = '" & Format(txtDoB, "dd/MMM/yyyy") & "' " & _
          "AND D.SampleID < '" & txtSampleID & "' " & _
          "AND R.SampleID = D.SampleID " & _
          "AND D.SampleDate < '" & SampleDate & "') " & _
          "ORDER BY D.RunDate desc"

140 CreateSql = sql

150 Exit Function

CreateSql_Error:

    Dim strES As String
    Dim intEL As Integer

160 intEL = Erl
170 strES = Err.Description
180 LogError "frmEditAll", "CreateSql", intEL, strES

End Function

Private Sub cRetics_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

10  On Error GoTo cRetics_MouseUp_Error

20  If cRetics = 0 Then
30      If Trim$(tRetA) = "?" Then
40          tRetA = ""
50          tRetP = ""
60      ElseIf Trim$(tRetA) <> "" Then
70          cRetics = 1
80      End If
90  Else
100     If Trim$(tRetA) = "" Then
110         tRetA = "?"
120         tRetP = "?"
130     End If
140 End If

150 cmdSaveHaem.Enabled = True: cmdHSaveH.Enabled = True

160 Exit Sub

cRetics_MouseUp_Error:

    Dim strES As String
    Dim intEL As Integer

170 intEL = Erl
180 strES = Err.Description
190 LogError "frmEditAll", "cRetics_MouseUp", intEL, strES

End Sub

Private Sub cRooH_Click(Index As Integer)

10  On Error GoTo cRooH_Click_Error

20  cmdSaveDemographics.Enabled = True
30  cmdSaveInc.Enabled = True

40  Exit Sub

cRooH_Click_Error:

    Dim strES As String
    Dim intEL As Integer

50  intEL = Erl
60  strES = Err.Description
70  LogError "frmEditAll", "cRooH_Click", intEL, strES

End Sub




Private Sub cUnits_KeyPress(KeyAscii As Integer)

10  KeyAscii = 0

End Sub

Private Sub DeltaCheck(ByVal Analyte As String, _
                       ByVal Value As String, _
                       ByVal PreviousValue As String, _
                       ByVal PreviousDate As String, _
                       ByVal PreviousID As String)

    Dim sql As String
    Dim tb As Recordset
    Dim CheckTime As Integer

10  On Error GoTo DeltaCheck_Error

20  sql = "SELECT * FROM HaemTestDefinitions WHERE " & _
          "AnalyteName = '" & Analyte & "' " & _
          "AND DoDelta = 1"
30  Set tb = New Recordset
40  RecOpenServer 0, tb, sql
50  Do While Not tb.EOF
60      If IsNull(tb!CheckTime) Then
70          CheckTime = 1
80      Else
90          CheckTime = tb!CheckTime
100     End If
110     If tb!AnalyteName = Analyte Then
120         If (dtSampleDate - CDate(PreviousDate)) <= CheckTime Then
130             If PreviousValue <> 0 Then
140                 If tb!AgeFromDays > 0 And tb!AgeToDays >= MaxAgeToDays Then
150                     If Abs(Val(PreviousValue) - Val(Value)) > tb!DeltaValue Then
160                         pdelta.ForeColor = vbBlue
170                         pdelta.Print Left$(Format$(PreviousDate, "dd/mm/yyyy") & _
                                               "(" & PreviousID & ") " & _
                                               Analyte & ":" & Space(25), 25); PreviousValue
180                         Exit Do
190                     End If
200                 End If
210             End If
220         End If
230     End If
240     tb.MoveNext
250 Loop

260 Exit Sub

DeltaCheck_Error:

    Dim strES As String
    Dim intEL As Integer

270 intEL = Erl
280 strES = Err.Description
290 LogError "frmEditAll", "DeltaCheck", intEL, strES, sql

End Sub

Private Sub ValidateDemographics(ByVal Validate As Boolean)

    Dim sql As String
    Dim tb As New Recordset

10  On Error GoTo ValidateDemographics_Error

20  If Validate Then
30      If cmdSaveDemographics.Enabled Then
40          SaveDemographics
50      End If
60      sql = "SELECT * FROM Demographics WHERE " & _
              "SampleID = '" & Val(txtSampleID) & "'"
70      Set tb = New Recordset
80      RecOpenServer 0, tb, sql
90      If Not tb.EOF Then
100         sql = "UPDATE Demographics SET Valid = 1, " & _
                  "UserName = '" & UserName & "' WHERE " & _
                  "SampleID = '" & Val(txtSampleID) & "'"
110         Cnxn(0).Execute sql
120         EnableDemographicEntry False
130         cmdDemoVal.Caption = "VALID"
140         cmdSaveDemographics.Enabled = False
150         cmdSaveInc.Enabled = False
160     End If
170 Else
180     If UCase(iBOX("Enter password to unValidate ?", , , True)) = UserPass Then
190         sql = "SELECT * FROM Demographics WHERE " & _
                  "SampleID = '" & Val(txtSampleID) & "'"

200         Set tb = New Recordset
210         RecOpenServer 0, tb, sql
220         If Not tb.EOF Then
230             sql = "UPDATE Demographics SET valid = 0, " & _
                      "UserName = '" & UserName & "' WHERE " & _
                      "SampleID = '" & Val(txtSampleID) & "'"
240             Cnxn(0).Execute sql
250             EnableDemographicEntry True
260             cmdDemoVal.Caption = "&Validate"
270         End If
280     End If
290 End If

300 Exit Sub

ValidateDemographics_Error:

    Dim strES As String
    Dim intEL As Integer

310 intEL = Erl
320 strES = Err.Description
330 LogError "frmEditAll", "ValidateDemographics", intEL, strES, sql

End Sub
Private Sub dtRecDate_CloseUp()
10  On Error GoTo dtRecDate_CloseUp_Error

20  pBar = 0

30  cmdSaveDemographics.Enabled = True
40  cmdSaveInc.Enabled = True

50  Exit Sub

dtRecDate_CloseUp_Error:

    Dim strES As String
    Dim intEL As Integer

60  intEL = Erl
70  strES = Err.Description
80  LogError "frmEditAll", "dtRecDate_CloseUp", intEL, strES

End Sub

Private Sub dtRecDate_LostFocus()

10  SetDatesColour Me

End Sub


Private Sub dtRunDate_CloseUp()

10  On Error GoTo dtRunDate_CloseUp_Error

20  pBar = 0

30  cmdSaveDemographics.Enabled = True
40  cmdSaveInc.Enabled = True

50  Exit Sub

dtRunDate_CloseUp_Error:

    Dim strES As String
    Dim intEL As Integer

60  intEL = Erl
70  strES = Err.Description
80  LogError "frmEditAll", "dtRunDate_CloseUp", intEL, strES

End Sub

Private Sub dtRunDate_LostFocus()

10  SetDatesColour Me

End Sub


Private Sub dtSampleDate_CloseUp()

10  On Error GoTo dtSampleDate_CloseUp_Error

20  pBar = 0

30  lblSampleDate = dtSampleDate

40  cmdSaveDemographics.Enabled = True
50  cmdSaveInc.Enabled = True

60  Exit Sub

dtSampleDate_CloseUp_Error:

    Dim strES As String
    Dim intEL As Integer

70  intEL = Erl
80  strES = Err.Description
90  LogError "frmEditAll", "dtSampleDate_CloseUp", intEL, strES

End Sub

Private Sub FillcAdd()

    Dim tb As New Recordset
    Dim sql As String

10  On Error GoTo FillcAdd_Error

20  sql = "SELECT distinct B.ShortName, B.PrintPriority " & _
          "from BioTestDefinitions as B, Lists as L " & _
          "WHERE B.SampleType = L.Code " & _
          "and L.ListType = 'ST' " & _
          "and L.Text like '" & cISampleType(3) & "%' and b.inuse = '1' " & _
          "order by B.PrintPriority"
30  Set tb = New Recordset
40  RecOpenServer 0, tb, sql

50  cAdd.Clear

60  Do While Not tb.EOF
70      cAdd.AddItem tb!ShortName & ""
80      tb.MoveNext
90  Loop

100 Exit Sub

FillcAdd_Error:

    Dim strES As String
    Dim intEL As Integer

110 intEL = Erl
120 strES = Err.Description
130 LogError "frmEditAll", "FillcAdd", intEL, strES, sql

End Sub

Private Sub FillCats()

    Dim sql As String
    Dim tb As New Recordset

10  On Error GoTo FillCats_Error

20  cCat(0).Clear
30  cCat(1).Clear

40  sql = "SELECT * from categorys"
50  Set tb = New Recordset
60  RecOpenServer 0, tb, sql
70  Do While Not tb.EOF
80      cCat(0).AddItem Trim(tb!Cat)
90      cCat(1).AddItem Trim(tb!Cat)
100     tb.MoveNext
110 Loop

120 If cCat(0).ListCount > 0 Then
130     cCat(0).ListIndex = 0
140     cCat(1).ListIndex = 0
150 End If

160 Exit Sub

FillCats_Error:

    Dim strES As String
    Dim intEL As Integer

170 intEL = Erl
180 strES = Err.Description
190 LogError "frmEditAll", "FillCats", intEL, strES, sql

End Sub

Private Sub FillcbAdd()

    Dim tb As New Recordset
    Dim sql As String
    Dim Found As Boolean
    Dim n As Long

10  On Error GoTo FillcbAdd_Error

20  sql = "SELECT distinct ShortName, PrintPriority " & _
          "from bgaTestDefinitions " & _
          "WHERE InUse = '1' and sampletype = '" & Left(cISampleType(2), 1) & "' " & _
          "order by shortname"
30  Set tb = New Recordset
40  RecOpenServer 0, tb, sql

50  cIAdd(2).Clear
60  Do While Not tb.EOF
70      Found = False
80      For n = 0 To cIAdd(2).ListCount - 1
90          If cIAdd(2).List(n) = tb!ShortName Then
100             Found = True
110         End If
120     Next
130     If Not Found Then
140         cIAdd(2).AddItem tb!ShortName
150     End If
160     Found = False
170     tb.MoveNext
180 Loop

190 Exit Sub

FillcbAdd_Error:

    Dim strES As String
    Dim intEL As Integer

200 intEL = Erl
210 strES = Err.Description
220 LogError "frmEditAll", "FillcbAdd", intEL, strES, sql

End Sub

Private Sub FillcEAdd()

    Dim tb As New Recordset
    Dim sql As String
    Dim Found As Boolean
    Dim n As Long

10  On Error GoTo FillcEAdd_Error

20  sql = "SELECT distinct ShortName, PrintPriority " & _
          "from EndTestDefinitions " & _
          "WHERE InUse = '1' and sampletype = '" & ListCodeFor("ST", cISampleType(0)) & "' " & _
          "order by PrintPriority"
30  Set tb = New Recordset
40  RecOpenServer 0, tb, sql

50  cIAdd(0).Clear
60  Do While Not tb.EOF
70      Found = False
80      For n = 0 To cIAdd(0).ListCount - 1
90          If cIAdd(0).List(n) = tb!ShortName Then
100             Found = True
110         End If
120     Next
130     If Not Found Then
140         cIAdd(0).AddItem tb!ShortName
150     End If
160     tb.MoveNext
170 Loop

180 Exit Sub

FillcEAdd_Error:

    Dim strES As String
    Dim intEL As Integer

190 intEL = Erl
200 strES = Err.Description
210 LogError "frmEditAll", "FillcEAdd", intEL, strES, sql

End Sub

Private Sub FillcIAdd()

    Dim tb As New Recordset
    Dim sql As String
    Dim Found As Boolean
    Dim n As Long

10  On Error GoTo FillcIAdd_Error

20  sql = "SELECT distinct ShortName, PrintPriority " & _
          "from ImmTestDefinitions " & _
          "WHERE InUse = '1' and sampletype = '" & ListCodeFor("ST", cISampleType(1)) & "' and hospital = '" & HospName(0) & "' " & _
          "order by shortname"
30  Set tb = New Recordset
40  RecOpenServer 0, tb, sql

50  cIAdd(1).Clear
60  Do While Not tb.EOF
70      Found = False
80      For n = 0 To cIAdd(1).ListCount - 1
90          If cIAdd(1).List(n) = tb!ShortName Then
100             Found = True
110         End If
120     Next
130     If Not Found Then
140         cIAdd(1).AddItem tb!ShortName
150     End If
160     Found = False
170     tb.MoveNext
180 Loop

190 Exit Sub

FillcIAdd_Error:

    Dim strES As String
    Dim intEL As Integer

200 intEL = Erl
210 strES = Err.Description
220 LogError "frmEditAll", "FillcIAdd", intEL, strES, sql

End Sub

Private Sub FillcParameter()

    Dim tb As New Recordset
    Dim n As Long
    Dim InList As Boolean
    Dim InUList As Boolean
    Dim sql As String

10  On Error GoTo FillcParameter_Error

20  cParameter.Clear

30  sql = "SELECT * from coagtestdefinitions"
40  Set tb = New Recordset
50  RecOpenServer 0, tb, sql

60  Do While Not tb.EOF
70      InList = False
80      For n = 0 To cParameter.ListCount - 1
90          If cParameter.List(n) = Trim(tb!TestName) Then
100             InList = True
110         End If
120     Next
130     If Not InList Then
140         cParameter.AddItem Trim(tb!TestName)
150     End If
160     InUList = False
170     For n = 0 To cCunits.ListCount - 1
180         If cCunits.List(n) = Trim(tb!Units) Then
190             InUList = True
200         End If
210     Next
220     If Not InUList Then
230         cCunits.AddItem Trim(tb!Units)
240     End If
250     tb.MoveNext
260 Loop

270 Exit Sub

FillcParameter_Error:

    Dim strES As String
    Dim intEL As Integer

280 intEL = Erl
290 strES = Err.Description
300 LogError "frmEditAll", "FillcParameter", intEL, strES, sql

End Sub

Private Sub FillcSampleType()

    Dim sql As String
    Dim tb As New Recordset

10  On Error GoTo FillcSampleType_Error

20  cISampleType(3).Clear

30  sql = "SELECT * from lists WHERE listtype = 'ST' order by listorder"
40  Set tb = New Recordset
50  RecOpenServer 0, tb, sql
60  Do While Not tb.EOF

70      cISampleType(0).AddItem Trim(tb!Text)
80      cISampleType(1).AddItem Trim(tb!Text)
90      cISampleType(2).AddItem Trim(tb!Text)
100     cISampleType(3).AddItem Trim(tb!Text)

110     tb.MoveNext
120 Loop

130 If cISampleType(3).ListCount > 0 Then
        '  For n = 1 To cSampleType.ListCount - 1
        '    If InStr(UCase(cSampleType.List(n)), "SERUM") > 0 Then

140     cISampleType(0).ListIndex = 0
150     cISampleType(1).ListIndex = 0
160     cISampleType(3).ListIndex = 0

        '    End If
        '  Next
170     FillcAdd
180     FillcEAdd
190     FillcIAdd
200     cISampleType(2).ListIndex = 0
210 End If

220 Exit Sub

FillcSampleType_Error:

    Dim strES As String
    Dim intEL As Integer

230 intEL = Erl
240 strES = Err.Description
250 LogError "frmEditAll", "FillcSampleType", intEL, strES, sql

End Sub

Private Sub FillLists()

    Dim tb As New Recordset
    Dim sql As String

10  On Error GoTo FillLists_Error

20  FillGPsClinWard Me, HospName(0)

30  FillUnits

40  cClDetails.Clear
50  cmbHospital.Clear

60  sql = "SELECT ListType, Text FROM Lists WHERE " & _
          "ListType = 'CD' " & _
          "OR ListType = 'HO' " & _
          "ORDER BY ListOrder"
70  Set tb = New Recordset
80  RecOpenServer 0, tb, sql
90  Do While Not tb.EOF
100     If Trim(tb!ListType) = "CD" Then
110         cClDetails.AddItem Trim$(tb!Text & "")
120     ElseIf Trim(tb!ListType) = "HO" Then
130         cmbHospital.AddItem Trim$(tb!Text & "")
140     End If
150     tb.MoveNext
160 Loop

170 cClDetails.ListIndex = -1
180 cmbHospital.ListIndex = -1

190 Exit Sub

FillLists_Error:

    Dim strES As String
    Dim intEL As Integer

200 intEL = Erl
210 strES = Err.Description
220 LogError "frmEditAll", "FillLists", intEL, strES, sql

End Sub

Private Sub FillUnits()

    Dim tb As New Recordset
    Dim sql As String

10  On Error GoTo FillUnits_Error

20  cUnits.Clear
30  cCunits.Clear
40  cIUnits(0).Clear
50  cIUnits(1).Clear

60  sql = "SELECT * from lists WHERE listtype = 'UN'"
70  Set tb = New Recordset
80  RecOpenServer 0, tb, sql
90  Do While Not tb.EOF
100     cUnits.AddItem Trim(tb!Text)
110     cIUnits(0).AddItem Trim(tb!Text)
120     cIUnits(1).AddItem Trim(tb!Text)
130     cCunits.AddItem Trim(tb!Text)
140     tb.MoveNext
150 Loop
160 cUnits.ListIndex = -1
170 cCunits.ListIndex = -1
180 cIUnits(0).ListIndex = -1
190 cIUnits(1).ListIndex = -1

200 Exit Sub

FillUnits_Error:

    Dim strES As String
    Dim intEL As Integer

210 intEL = Erl
220 strES = Err.Description
230 LogError "frmEditAll", "FillUnits", intEL, strES, sql

End Sub

Private Sub dtSampleDate_LostFocus()

10  SetDatesColour Me

End Sub

Private Sub Form_Activate()

10  On Error GoTo Form_Activate_Error

20  If Trim$(txtSampleID) = "" Then txtSampleID.SetFocus

30  TimerBar.Enabled = True
40  pBar = 0

50  Set_Font Me
60  SetFormCaption
70  SetToolTip SysOptToolTip(0), Me

80  UpDown1.Max = 99999999

90  cmdSaveDemographics.Visible = GetOptionSetting("SaveDemographicHidden", "0") = "0"
100 cmdSaveInc.Visible = GetOptionSetting("SaveDemographicHidden", "0") = "0"

110 Exit Sub

Form_Activate_Error:

    Dim strES As String
    Dim intEL As Integer

120 intEL = Erl
130 strES = Err.Description
140 LogError "frmEditAll", "Form_Activate", intEL, strES

End Sub

Private Sub Form_Deactivate()

10  On Error GoTo Form_Deactivate_Error

20  Me.Refresh
30  pBar = 0
40  TimerBar.Enabled = False

50  Exit Sub

Form_Deactivate_Error:

    Dim strES As String
    Dim intEL As Integer

60  intEL = Erl
70  strES = Err.Description
80  LogError "frmEditAll", "Form_Deactivate", intEL, strES


End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

10  pBar = 0

End Sub

Private Sub Form_Load()
          Dim n As Long
          Dim sql As String
          Dim tb As New Recordset
          Dim ax As Control



10    On Error GoTo Form_Load_Error

20    SampleType = ""

30    sql = "SELECT * from options WHERE " & _
          "username = '" & UserName & "' " & _
          "and description like 'frmEditAll.%' order by contents desc"
40    Set tb = New Recordset
50    RecOpenServer 0, tb, sql
60    Do While Not tb.EOF
70      For Each ax In Me
80          If UCase("frmEditAll" & ax.Name) = UCase(Trim(tb!Description)) Then
90              ax.TabIndex = tb!Contents
100         End If
110     Next
120     tb.MoveNext
130   Loop

140   n = n + 1

150   UpDown1.Max = (2 ^ 31) - 1    '2147483647 '999999999

160   If SysOptDontShowPrevCoag(0) = True Then
170     grdPrev.Visible = False
180     lblPrevCoag.Visible = False
190   End If

200   EndLoaded = False
210   ImmLoaded = False

220   StatusBar1.Panels(1).Text = UserName

230   If SysOptDemoVal(0) = False Then cmdDemoVal.Visible = False
240   If SysOptDeptBio(0) = False Then ssTabAll.TabVisible(2) = False Else n = n + 1
250   If SysOptDeptHaem(0) = False Then ssTabAll.TabVisible(1) = False Else n = n + 1
260   If SysOptDeptCoag(0) = False Then ssTabAll.TabVisible(3) = False Else n = n + 1
270   If SysOptDeptEnd(0) = False Then ssTabAll.TabVisible(4) = False Else n = n + 1
280   If SysOptDeptBga(0) = False Then ssTabAll.TabVisible(5) = False Else n = n + 1
290   If SysOptDeptImm(0) = False Then ssTabAll.TabVisible(6) = False Else n = n + 1
300   If SysOptDeptExt(0) = False Then ssTabAll.TabVisible(7) = False Else n = n + 1
310   If PrnAll(0) = False Then cmdPrintAll.Visible = False
320   If SysOptPhone(0) = False Then cmdPhone.Visible = False

330   If SysOptHaemAn1(0) = "ADVIA" Then
340     Label1(17) = "WBCP"
350     Label1(18) = "WBCB"
360     cmdPrintEsr.Visible = False
370   End If

380   With lblChartNumber
390     .BackColor = &H8000000F
400     .ForeColor = vbBlack
410     Select Case UCase(HospName(0))
        Case "PORTLAOISE", "DEMONSTRATION"
420         .Caption = initial2upper(HospName(0)) & " Chart #"
430         lblAandE.Visible = False
440         txtAandE.Visible = False
            'lblNameTitle.Left = txtName.Left
450         txtName.Left = txtAandE.Left
460         txtName.Width = txtName.Width + txtAandE.Width
470     Case "MULLINGAR", "TULLAMORE"
480         .Caption = initial2upper(HospName(0)) & " Chart #"
490         lblAandE.Visible = True
500         txtAandE.Visible = True
            'txtAandE.Width = 2000
            'lblNameTitle.Left = txtName.Left
510         txtName.Left = 2595
520         txtName.Width = 4335
530     End Select
540   End With

550   ssTabAll.TabsPerRow = n

560   With lblViewSplit
570     Select Case GetSetting("NetAcquire", "StartUp", "Split", "All")
        Case "All":
580         .Caption = "Viewing All"
590         .BackColor = &H8000000F
600         .ForeColor = vbBlack
610     Case "Pri":
620         .Caption = "Viewing Primary Split"
630         .BackColor = &H800080
640         .ForeColor = &HFF00&
650     Case "Viewing Sec":
660         .Caption = "Viewing Secondary Split"
670         .BackColor = &H800080
680         .ForeColor = &HFF00&
690     End Select
700   End With

710   With lblImmViewSplit(0)
720     Select Case GetSetting("NetAcquire", "StartUp", "EndSplit", "All")
        Case "All":
730         .Caption = "Viewing All"
740         .BackColor = &H8000000F
750         .ForeColor = vbBlack
760     Case "Pri":
770         .Caption = "Viewing Primary Split"
780         .BackColor = &H800080
790         .ForeColor = &HFF00&
800     Case "Viewing Sec":
810         .Caption = "Viewing Secondary Split"
820         .BackColor = &H800080
830         .ForeColor = &HFF00&
840     End Select
850   End With
860   With lblImmViewSplit(1)
870     Select Case GetSetting("NetAcquire", "StartUp", "ImmSplit", "All")
        Case "All":
880         .Caption = "Viewing All"
890         .BackColor = &H8000000F
900         .ForeColor = vbBlack
910     Case "Pri":
920         .Caption = "Viewing Primary Split"
930         .BackColor = &H800080
940         .ForeColor = &HFF00&
950     Case "Viewing Sec":
960         .Caption = "Viewing Secondary Split"
970         .BackColor = &H800080
980         .ForeColor = &HFF00&
990     End Select
1000  End With

1010  cmdViewBioReps.Visible = SysOptRTFView(0)
1020  cmdViewReports.Visible = SysOptRTFView(0)
1030  cmdViewCoagRep.Visible = SysOptRTFView(0)
1040  cmdViewHaemRep.Visible = SysOptRTFView(0)
1050  cmdViewHaemRep.Visible = SysOptRTFView(0)
1060  cmdViewImmRep.Visible = SysOptRTFView(0)
1070  cmdViewExtReport.Visible = SysOptRTFView(0)

1080  FillcSampleType
1090  FillcParameter
1100  FillLists
1110  FillCats
1120  ClearHaemDiffGrid

1130  FillMRU cMRU
1140  ClearRbcGrid

1150  With lblChartNumber
1160    .BackColor = &H8000000F
1170    .ForeColor = vbBlack
1180    Select Case Entity(0)
        Case "03"
1190        .Caption = "Portlaoise Chart #"
1200    Case "04"
1210        .Caption = "Tullamore Chart #"
1220    End Select
1230  End With

1240  dtRunDate = Format$(Now, "dd/mm/yyyy")
1250  lblRundate = dtRunDate
1260  dtSampleDate = Format$(Now, "dd/mm/yyyy")

1270  UpDown1.Max = 99999999

1280  txtSampleID = GetSetting("NetAcquire", "StartUp", "LastUsed", "1")

1290  LoadAllDetails

1300  pBar.Max = LogOffDelaySecs

1310  If UserMemberOf = "Secretarys" Then
1320    For n = 1 To 6
1330        ssTabAll.TabVisible(n) = False
1340    Next
1350  Else
1360    cmdSaveDemographics.Enabled = False
1370    cmdSaveInc.Enabled = False
1380    cmdSaveHaem.Enabled = False
1390    cmdSaveComm.Enabled = False
1400    cmdHSaveH.Enabled = False
1410    cmdSaveBio.Enabled = False
1420    cmdSaveCoag.Enabled = False
1430    cmdSaveImm(0).Enabled = False
1440    cmdSaveImm(1).Enabled = False
1450  End If

1460  Activated = False

1470  Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

1480  intEL = Erl
1490  strES = Err.Description
1500  LogError "frmEditAll", "Form_Load", intEL, strES

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

10  pBar = 0

End Sub

Private Sub Form_Paint()

    Dim TabNumber As Long
    Dim ax As Control

10  On Error GoTo Form_Paint_Error

20  If Activated Then Exit Sub

30  Activated = True

40  If SysOptDefaultTab(0) <> "" Then
50      TabNumber = Val(SysOptDefaultTab(0))
60  Else
70      TabNumber = Val(GetSetting("NetAcquire", "StartUp", "LastDepartment", "0"))
80  End If

90  If SysOptDontShowPrevCoag(0) = True Then
100     grdPrev.Visible = False
110     lblPrevCoag.Visible = False
120 End If

130 If ssTabAll.TabVisible(TabNumber) = False Then TabNumber = 1

140 If UserMemberOf = "Secretarys" Then TabNumber = 0

150 ssTabAll.Tab = TabNumber

160 For Each ax In Me
170     If ax.Name = SysSetFoc(0) Then
180         ax.SetFocus
190     End If
200 Next

210 Exit Sub

Form_Paint_Error:

    Dim strES As String
    Dim intEL As Integer

220 intEL = Erl
230 strES = Err.Description
240 LogError "frmEditAll", "Form_Paint", intEL, strES

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Dim strSplitStatus As String

10  On Error GoTo Form_Unload_Error

20  If Val(txtSampleID) <> Val(GetSetting("NetAcquire", "StartUp", "LastUsed", "1")) Then
30      SaveSetting "NetAcquire", "StartUp", "LastUsed", txtSampleID
40  End If

50  SaveSetting "NetAcquire", "StartUp", "LastDepartment", CStr(ssTabAll.Tab)

60  With lblViewSplit
70      If InStr(.Caption, "All") Then
80          strSplitStatus = "All"
90      ElseIf InStr(.Caption, "Pri") Then
100         strSplitStatus = "Pri"
110     ElseIf InStr(.Caption, "Sec") Then
120         strSplitStatus = "Sec"
130     End If
140     SaveSetting "NetAcquire", "StartUp", "Split", strSplitStatus
150 End With


160 With lblImmViewSplit(0)
170     If InStr(.Caption, "All") Then
180         strSplitStatus = "All"
190     ElseIf InStr(.Caption, "Pri") Then
200         strSplitStatus = "Pri"
210     ElseIf InStr(.Caption, "Sec") Then
220         strSplitStatus = "Sec"
230     End If
240     SaveSetting "NetAcquire", "StartUp", "EndSplit", strSplitStatus
250 End With

260 With lblImmViewSplit(1)
270     If InStr(.Caption, "All") Then
280         strSplitStatus = "All"
290     ElseIf InStr(.Caption, "Pri") Then
300         strSplitStatus = "Pri"
310     ElseIf InStr(.Caption, "Sec") Then
320         strSplitStatus = "Sec"
330     End If
340     SaveSetting "NetAcquire", "StartUp", "ImmSplit", strSplitStatus
350 End With


360 pPrintToPrinter = ""

370 Activated = False

380 Exit Sub

Form_Unload_Error:

    Dim strES As String
    Dim intEL As Integer

390 intEL = Erl
400 strES = Err.Description
410 LogError "frmEditAll", "Form_Unload", intEL, strES

End Sub

Private Sub fraDate_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

10  pBar = 0

End Sub


Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

10  pBar = 0

End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

10  pBar = 0

End Sub

Private Sub Frame3_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

10  pBar = 0

End Sub

Private Sub Frame4_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

10  pBar = 0

End Sub

Private Sub Frame5_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

10  pBar = 0

End Sub

Private Sub Frame6_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

10  pBar = 0

End Sub

Private Sub Frame8_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

10  pBar = 0

End Sub

Private Sub Frame9_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

10  pBar = 0

End Sub

Private Sub gBga_Click()
    Dim sql As String

10  On Error GoTo gBga_Click_Error

20  If gBga.MouseRow = 0 Then Exit Sub

30  If iMsg("DELETE " & gBga.TextMatrix(gBga.Row, 0) & " !", vbYesNo) <> vbYes Then
40      Exit Sub
50  End If

60  If InStr(gBga.TextMatrix(gBga.Row, 6), "V") > 0 Then
70      If UCase(iBOX("Valid Result! Enter Password", , , True)) <> UserPass Then
80          Exit Sub
90      End If
100 End If

110 tINewValue(2) = gBga.TextMatrix(gBga.Row, 1)
120 cIUnits(2) = gBga.TextMatrix(gBga.Row, 2)

130 gBga.Col = 0
140 cIAdd(2) = gBga

150 sql = "DELETE from bgaresults WHERE " & _
          "sampleid = '" & txtSampleID & "' " & _
          "and code = '" & BgaCodeForShortName(gBga) & "'"
160 Cnxn(0).Execute sql

170 LoadBloodGas

180 tINewValue(2).SetFocus

190 Exit Sub

gBga_Click_Error:

    Dim strES As String
    Dim intEL As Integer

200 intEL = Erl
210 strES = Err.Description
220 LogError "frmEditAll", "gBga_Click", intEL, strES, sql

End Sub

Private Sub gBio_Click()

    Dim tb As New Recordset
    Dim sql As String
    Dim s As String
    Dim R As Integer
    Dim f As Form

10  On Error GoTo gBio_Click_Error

20  If gBio.MouseRow = 0 Then Exit Sub
30  If gBio.TextMatrix(gBio.Row, 0) = "" Then Exit Sub

40  R = gBio.Row

50  If gBio.TextMatrix(R, 0) = "HbA1c" Then
60      frmViewFullDataHBA.SampleID = txtSampleID
70      frmViewFullDataHBA.Show 1
80      Exit Sub
90  End If

100 If gBio.Col = 10 Then
110     If gBio.CellBackColor <> vbRed Then
120         If gBio.CellPicture = imgGreenTick.Picture Then
130             Set gBio.CellPicture = imgRedCross.Picture
140         Else
150             Set gBio.CellPicture = imgGreenTick.Picture
160         End If
170     End If
180     Exit Sub
190 End If

200 If gBio.Col = 5 Then
210     Select Case gBio
        Case "": Exit Sub
220     Case "AE": s = "ADC Error"
230     Case "AH": s = "Initial Absorbance High"
240     Case "BH": s = "Blank Absorbance High"
250     Case "BL": s = "Blank Absorbance Low"
260     Case "BN": s = "Blank Mean Deviation"
270     Case "BO": s = "Blank Maximum Deviation"
280     Case "DH": s = "Dynamic Range High"
290     Case "DL": s = "Dynamic range Low"
300     Case "DR": s = "Reference Drift (ISE)"
310     Case "EA": s = "Erratic ADC (ISE)"
320     Case "HR": s = "Reaction Absorbance High"
330     Case "IR": s = "Initial Absorbance High"
340     Case "IT": s = "Iteration Tolerance"
350     Case "LR": s = "Reaction Absorbance Low"
360     Case "NT": s = "Noise Threshold"
370     Case "OH": s = "ORDAC High"
380     Case "OL": s = "ORDAC Low"
390     Case "OT": s = "Outliers Threshold"
400     Case "RH": s = "Reaction Rate High"
410     Case "RL": s = "Reaction Rate Low"
420     Case "RN": s = "Reaction Mean Deviation"
430     Case "RO": s = "Reaction Maximum Deviation"
440     Case "SD": s = "Substrate Depleted"
450     Case "SH": s = "Blank Rate High"
460     Case "SL": s = "Blank Rate Low"
470     Case "TM": s = "Temperature"
480     Case Else: s = "Unknown Error"
490     End Select
500     iMsg s, vbInformation
510     Exit Sub
520 End If
    'By Farhan Waheed 18/04/2016
    '---------------
530 Debug.Print UCase(GetAnyFieldFromTestDefinitions("Hospital", gBio.TextMatrix(R, 0), txtSampleID, "Bio", ""))
540 Debug.Print UCase(HospName(0))
550 If Trim(UCase(GetAnyFieldFromTestDefinitions("Hospital", gBio.TextMatrix(R, 0), txtSampleID, "Bio", ""))) <> Trim(UCase(HospName(0))) Then
560     iMsg "You are not authorised to change this test"
570     Exit Sub
580 End If
    '===============
590 If InStr(gBio.TextMatrix(R, 6), "V") > 0 Then
600     If UCase(iBOX("Valid Result! Enter Password", , , True)) <> UserPass Then
610         Exit Sub
620     End If
630 End If

640 If gBio.Col = 7 Then
650     Select Case gBio.TextMatrix(R, 7)
        Case "": gBio.TextMatrix(R, 7) = "P"
660     Case "P": gBio.TextMatrix(R, 7) = "C"
670     Case "C": gBio.TextMatrix(R, 7) = "PC"
680     Case Else: gBio.TextMatrix(R, 7) = ""
690     End Select
700     sql = "UPDATE BioResults " & _
              "SET PC = '" & gBio.TextMatrix(R, 7) & "' WHERE " & _
              "SampleID = '" & txtSampleID & "' " & _
              "AND Code = " & _
            "  (SELECT Top 1 Code FROM BioTestDefinitions WHERE " & _
            "   ShortName = '" & gBio.TextMatrix(R, 0) & "' )"
710     Cnxn(0).Execute sql
720     Exit Sub
730 End If

740 If gBio.Col = 9 Then
750     Set f = New frmComment
760     With f
770         .Discipline = "BIO"
780         .Comment = gBio.TextMatrix(R, 9)
790         .Show 1
800         gBio.TextMatrix(R, 9) = .Comment
810     End With
820     Unload f
830     Set f = Nothing

840     sql = "UPDATE BioResults " & _
              "SET Comment = '" & Left$(gBio.TextMatrix(R, 9), 100) & "' WHERE " & _
              "SampleID = '" & txtSampleID & "' " & _
              "AND Code = " & _
            "  (SELECT Top 1 Code FROM BioTestDefinitions WHERE " & _
            "   ShortName = '" & gBio.TextMatrix(R, 0) & "' )"
850     Cnxn(0).Execute sql
860     Exit Sub
870 End If

880 If gBio.Col = 1 And gBio.TextMatrix(R, 0) <> "" Then
890     If iMsg("DELETE " & gBio.TextMatrix(R, 0) & " ?", vbYesNo) = vbYes Then
900         tnewvalue = gBio.TextMatrix(R, 1)
910         cUnits = gBio.TextMatrix(R, 2)
920         cAdd = gBio.TextMatrix(R, 0)
930         sql = "SELECT * FROM BioResults WHERE " & _
                  "SampleID = '" & txtSampleID & "' " & _
                  "AND Code = " & _
                "  (SELECT DISTINCT Code FROM BioTestDefinitions WHERE InUse = 1 AND " & _
                "   ShortName = '" & gBio.TextMatrix(R, 0) & "' )"
940         Set tb = New Recordset
950         RecOpenServer 0, tb, sql
960         Frame2.Enabled = True
970         sql = "DELETE FROM BioResults WHERE " & _
                  "SampleID = '" & txtSampleID & "' " & _
                  "AND Code = " & _
                "  (SELECT DISTINCT Code FROM BioTestDefinitions WHERE " & _
                "   ShortName = '" & gBio.TextMatrix(R, 0) & "' )"
980         Cnxn(0).Execute sql
990         LoadBiochemistry

1000    End If
1010 End If

1020 Exit Sub

gBio_Click_Error:

    Dim strES As String
    Dim intEL As Integer

1030 intEL = Erl
1040 strES = Err.Description
1050 LogError "frmEditAll", "gBio_Click", intEL, strES, sql

End Sub

Private Sub gBio_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

    Dim xx As Integer
    Dim yy As Integer

10  On Error GoTo gBio_MouseMove_Error

20  xx = gBio.MouseCol
30  yy = gBio.MouseRow
40  gBio.ToolTipText = "Biochemistry Results"

50  If xx = 9 Then
60      If Trim(gBio.TextMatrix(yy, xx)) <> "" Then gBio.ToolTipText = gBio.TextMatrix(yy, xx)
70  ElseIf xx = 7 Then
80      If gBio.TextMatrix(yy, xx) = "P" Then
90          gBio.ToolTipText = "Phoned"
100     ElseIf gBio.TextMatrix(yy, xx) = "C" Then
110         gBio.ToolTipText = "Checked"
120     ElseIf gBio.TextMatrix(yy, xx) = "PC" Then
130         gBio.ToolTipText = "Checked & Phoned"
140     Else
150         gBio.ToolTipText = "Biochemistry Results"
160     End If
170 End If

180 pBar = 0

190 Exit Sub

gBio_MouseMove_Error:

    Dim strES As String
    Dim intEL As Integer

200 intEL = Erl
210 strES = Err.Description
220 LogError "frmEditAll", "gBio_MouseMove", intEL, strES

End Sub

Private Sub gImm_Click(Index As Integer)

    Dim sql As String
    Dim gy As Long
    Dim gX As Integer
    Dim f As Form
    Dim Cat As String


10    On Error GoTo gImm_Click_Error

20    If gImm(Index).MouseRow = 0 Then Exit Sub
30    If gImm(Index).TextMatrix(gImm(Index).Row, 0) = "" Then Exit Sub

40    gy = gImm(Index).Row
50    gX = gImm(Index).Col

60    If gX = 8 + Index Then
70      If gImm(Index).CellPicture = imgGreenTick.Picture Then
80          Set gImm(Index).CellPicture = imgRedCross.Picture
90      Else
100         Set gImm(Index).CellPicture = imgGreenTick.Picture
110     End If
120     Exit Sub
130   End If
    'By Farhan Waheed 18/04/2016
    '---------------


140   If cCat(0) = "" Then Cat = "Default" Else Cat = cCat(0)
150   Select Case ssTabAll.Tab
    Case 4
160     Debug.Print UCase(GetAnyFieldFromTestDefinitions("Hospital", gImm(Index).TextMatrix(gy, 0), txtSampleID, "End", Cat))
170     Debug.Print UCase(HospName(0))
180     If Trim(UCase(GetAnyFieldFromTestDefinitions("Hospital", gImm(Index).TextMatrix(gy, 0), txtSampleID, "End", Cat))) <> Trim(UCase(HospName(0))) Then
190         iMsg "You are not authorised to change this test"
200         Exit Sub
210     End If
220       Case 6
230     Debug.Print UCase(GetAnyFieldFromTestDefinitions("Hospital", gImm(Index).TextMatrix(gy, 0), txtSampleID, "Imm", Cat))
240     Debug.Print UCase(HospName(0))
250     If Trim(UCase(GetAnyFieldFromTestDefinitions("Hospital", gImm(Index).TextMatrix(gy, 0), txtSampleID, "Imm", Cat))) <> Trim(UCase(HospName(0))) Then
260         iMsg "You are not authorised to change this test"
270         Exit Sub
280     End If
290       End Select
    '===============
300   If InStr(gImm(Index).TextMatrix(gy, 6), "V") > 0 Then
310     If UCase(iBOX("Valid Result! Enter Password", , , True)) <> UserPass Then
320         Exit Sub
330     End If
340   End If

350   If Index = 0 Then           'Endocrinology Grid
360     If gX = 7 Then
370         Set f = New frmComment
380         With f
390             .Discipline = "END"
400             .Comment = gImm(0)
410             .Show 1
420             gImm(0) = .Comment
430         End With
440         Unload f
450         Set f = Nothing
460         sql = "UPDATE EndResults " & _
                  "SET Comment = '" & gImm(0).TextMatrix(gy, 7) & "' " & _
                  "WHERE sampleid = '" & txtSampleID & "' " & _
                  "AND Code = '" & eCodeForShortName(gImm(0).TextMatrix(gy, 0)) & "'"
470         Cnxn(0).Execute sql
480         Exit Sub
490     End If
        '    If gX = 1 And UCase$(gImm(0).Tag) = "VIROLOGY" Then
        '        LoadListGeneric cmbEndResults, "AxsymResults"
        '        ComboTop = (ssTabAll.Top + gImm(0).Top + gImm(0).Row * gImm(0).RowHeight(1)) + 50
        '        ComboLeft = gImm(0).Left + 150
        '
        '        For i = 0 To gImm(0).Col - 1
        '            ComboLeft = ComboLeft + gImm(0).ColWidth(i)
        '        Next i
        '
        '        If Not cmbEndResults Is Nothing Then
        '            cmbEndResults.Move ComboLeft, ComboTop, gImm(0).ColWidth(gImm(0).Col)
        '            If gImm(0).TextMatrix(gImm(0).Row, gImm(0).Col) <> "" Then
        '                cmbEndResults.Text = gImm(0).TextMatrix(gImm(0).Row, gImm(0).Col)
        '            End If
        '            cmbEndResults.Visible = True
        '        End If
        '        Exit Sub
        '    End If

500     If iMsg("DELETE " & gImm(0).TextMatrix(gy, 0) & " !", vbYesNo) = vbYes Then
510         tINewValue(0) = gImm(0).TextMatrix(gy, 1)
520         cIUnits(0) = gImm(0).TextMatrix(gy, 2)
530         gImm(0).Col = 0
540         cIAdd(0) = gImm(0)
550         sql = "DELETE from endresults WHERE " & _
                  "sampleid = '" & txtSampleID & "' " & _
                  "and code = '" & eCodeForShortName(gImm(0)) & "'"
560         Cnxn(0).Execute sql

570         LoadEndocrinology

580         If tINewValue(0).Enabled And tINewValue(0).Visible Then
590             tINewValue(0).SetFocus
600         End If
610     End If

620   Else

630     If gX = 7 Then
640         If gImm(1) = "" Then
650             gImm(1) = "P"
660         ElseIf gImm(1) = "P" Then
670             gImm(1) = "C"
680         ElseIf gImm(1) = "C" Then
690             gImm(1) = "PC"
700         ElseIf gImm(1) = "PC" Then
710             gImm(1) = ""
720         End If
730         sql = "UPDATE ImmResults " & _
                  "SET PC = '" & gImm(1).TextMatrix(gy, 7) & "' " & _
                  "WHERE SampleID = '" & txtSampleID & "' " & _
                  "AND Code = '" & ICodeForShortName(gImm(1).TextMatrix(gy, 0)) & "'"
740         Cnxn(0).Execute sql
750         Exit Sub
760     End If

770     If gX = 8 Then
780         Set f = New frmComment
790         With f
800             .Discipline = "IMM"
810             .Comment = gImm(1)
820             .Show 1
830             gImm(1) = .Comment
840         End With
850         Unload f
860         Set f = Nothing
870         sql = "UPDATE ImmResults SET Comment = '" & gImm(1).TextMatrix(gy, 8) & "' " & _
                  "WHERE SampleID = '" & txtSampleID & "' " & _
                  "AND Code = '" & ICodeForShortName(gImm(1).TextMatrix(gy, 0)) & "'"
880         Cnxn(0).Execute sql
890         Exit Sub
900     End If

910     If iMsg("DELETE " & gImm(1).TextMatrix(gy, 0) & " !", vbYesNo) = vbYes Then
920         tINewValue(1) = gImm(1).TextMatrix(gy, 1)
930         cIUnits(1) = gImm(1).TextMatrix(gy, 2)
940         gImm(1).Col = 0
950         cIAdd(1) = gImm(1)
960         sql = "DELETE FROM ImmResults WHERE " & _
                  "SampleID = '" & txtSampleID & "' " & _
                  "AND Code = '" & ICodeForShortName(gImm(1)) & "'"
970         Cnxn(0).Execute sql

980         LoadImmunology

990         If tINewValue(1).Visible And tINewValue(1).Enabled Then
1000            tINewValue(1).SetFocus
1010        End If
1020    End If
1030  End If

1040  Exit Sub

gImm_Click_Error:

    Dim strES As String
    Dim intEL As Integer

1050  intEL = Erl
1060  strES = Err.Description
1070  LogError "frmEditAll", "gImm_Click", intEL, strES, sql

End Sub

Private Sub gImm_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
10  cmbEndResults.Visible = False
End Sub

Private Sub gImm_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)

    Dim gX As Integer
    Dim gy As Integer

10  On Error GoTo gImm_MouseMove_Error

20  If Index = 1 Then
30      gX = gImm(1).MouseCol
40      gy = gImm(1).MouseRow
50      gImm(1).ToolTipText = "Immunology Results"

60      If gX = 8 Then
70          If Trim(gImm(1).TextMatrix(gy, gX)) <> "" Then gImm(1).ToolTipText = gImm(1).TextMatrix(gy, gX)
80      ElseIf gX = 1 Then
90          If Trim(gImm(1).TextMatrix(gy, gX)) <> "" Then gImm(1).ToolTipText = gImm(1).TextMatrix(gy, gX)
100     ElseIf gX = 7 Then
110         If gImm(1).TextMatrix(gy, gX) = "P" Then
120             gImm(1).ToolTipText = "Phoned"
130         ElseIf gImm(1).TextMatrix(gy, gX) = "C" Then
140             gImm(1).ToolTipText = "Checked"
150         ElseIf gImm(1).TextMatrix(gy, gX) = "PC" Then
160             gImm(1).ToolTipText = "Checked & Phoned"
170         Else
180             gImm(1).ToolTipText = "Immunology Results"
190         End If
200     ElseIf gX = 9 Then
210         gImm(1).ToolTipText = gImm(1).TextMatrix(gy, 0)
220     End If
230 Else
240     gX = gImm(0).MouseCol
250     gy = gImm(0).MouseRow
260     gImm(0).ToolTipText = "Endocrinology Results"

270     If gImm(0).MouseCol = 7 Then
280         If Trim(gImm(0).TextMatrix(gy, gX)) <> "" Then gImm(0).ToolTipText = gImm(0).TextMatrix(gy, gX)
290     ElseIf gImm(0).MouseCol = 1 Then
300         If Trim(gImm(0).TextMatrix(gy, gX)) <> "" Then gImm(0).ToolTipText = gImm(0).TextMatrix(gy, gX)
310     ElseIf gImm(0).MouseCol = 0 Then
320         If Trim(gImm(0).TextMatrix(gy, gX)) <> "" Then gImm(0).ToolTipText = EndLongNameFor(eCodeForShortName(gImm(0).TextMatrix(gy, gX)))
330     ElseIf gImm(0).MouseCol = 5 Then
340         If gImm(0).TextMatrix(gy, gX) = "P" Then
350             gImm(0).ToolTipText = "Phoned"
360         ElseIf gImm(0).TextMatrix(gy, gX) = "C" Then
370             gImm(0).ToolTipText = "Checked"
380         ElseIf gImm(0).TextMatrix(gy, gX) = "PC" Then
390             gImm(0).ToolTipText = "Checked & Phoned"
400         Else
410             gImm(0).ToolTipText = "Endocrinology Results"
420         End If
430     End If
440 End If

450 Exit Sub

gImm_MouseMove_Error:

    Dim strES As String
    Dim intEL As Integer

460 intEL = Erl
470 strES = Err.Description
480 LogError "frmEditAll", "gImm_MouseMove", intEL, strES

End Sub

Private Sub gRBC_Click()

10  On Error GoTo gRBC_Click_Error

20  If gRbc.ColSel = 0 And gRbc.RowSel = 2 Then
30      ClearHgb
40  End If

50  If gRbc.ColSel = 1 Then
60      If gRbc.MouseRow > 0 Then
70          Set grd = gRbc
80          grd.Row = grd.MouseRow
90          grd.Col = grd.MouseCol
100         LoadControls
110     End If
        '    txtInput.Text = gRbc.TextMatrix(gRbc.RowSel, 1)
        '    txtInput.SetFocus
120     Exit Sub
130 End If

140 If SysOptHaemAn1(0) = "ADVIA" Then
150     If Trim(gRbc.TextMatrix(11, 1)) = "" Then
160         Exit Sub
170     End If
        '  n = 100 - Val(gRbc.TextMatrix(12, 1))
        '
        '  tWBC = (tWBC / 100) * n
180 End If


190 Exit Sub

gRBC_Click_Error:

    Dim strES As String
    Dim intEL As Integer

200 intEL = Erl
210 strES = Err.Description
220 LogError "frmEditAll", "gRBC_Click", intEL, strES

End Sub

Private Sub gRbc_LeaveCell()
10  txtText.Visible = False
End Sub

Private Sub gRbc_Scroll()
10  grd = gRbc
20  LoadControls
End Sub

Private Sub grdCoag_Click()
    Dim tb As New Recordset
    Dim sql As String
    Dim Code As String


10  On Error GoTo grdCoag_Click_Error

20  If txtSampleID = "" Then Exit Sub

30  If grdCoag.MouseRow = 0 Then Exit Sub

40  If grdCoag.TextMatrix(grdCoag.Row, 0) = "" Then Exit Sub

50  If grdCoag.Col = 8 Then
60      grdCoag.CellPictureAlignment = flexAlignCenterCenter
70      If grdCoag.CellPicture = imgGreenTick.Picture Then
80          Set grdCoag.CellPicture = imgRedCross.Picture
90      Else
100         Set grdCoag.CellPicture = imgGreenTick.Picture
110     End If
120     Exit Sub
130 End If

140 If InStr(grdCoag.TextMatrix(grdCoag.Row, 5), "V") > 0 Then
150     If UCase(iBOX("Valid Result! Enter Password", , , True)) <> UserPass Then
160         Exit Sub
170     End If
180 End If

190 With grdCoag

200     Select Case .Col

        Case 0:
210         If iMsg("DELETE " & .Text & "?", vbQuestion + vbYesNo) = vbYes Then
220             cParameter = .Text
230             tResult = .TextMatrix(.Row, 1)
240             If Trim$(.TextMatrix(.Row, 2)) <> "" Then cCunits = .TextMatrix(.Row, 2)
250             If cCunits = "" Then cCunits = CoagUnitsFor(CoagCodeFor(cParameter))
260             If .Rows = 2 Then
270                 .AddItem ""
280                 .RemoveItem 1
290             Else
300                 .RemoveItem .Row
310             End If
320             Code = CoagCodeFor(cParameter)
330             sql = "SELECT * from coagresults WHERE " & _
                      "sampleid = '" & txtSampleID & "' " & _
                      "and Code = '" & Code & "'"
340             Set tb = New Recordset
350             RecOpenServer 0, tb, sql
360             sql = "DELETE from CoagResults WHERE " & _
                      "SampleID = '" & txtSampleID & "' " & _
                      "and Code = '" & Code & "'"
370             cmdSaveCoag.Enabled = True
380             Cnxn(0).Execute sql
390         End If

400     Case 1:
410         .Text = iBOX("Enter new Value for " & .TextMatrix(.Row, 0), , .Text)
420         cmdSaveCoag.Enabled = True

430     Case 5:
440         .Text = IIf(.Text = "", "V", "")
450         cmdSaveCoag.Enabled = True

460     End Select

470 End With

480 Exit Sub

grdCoag_Click_Error:

    Dim strES As String
    Dim intEL As Integer

490 intEL = Erl
500 strES = Err.Description
510 LogError "frmEditAll", "grdCoag_Click", intEL, strES, sql

End Sub

Private Sub grdCoag_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

10  pBar = 0

End Sub

Private Sub grdExt_Click()
    Dim Str As String
    Dim Prompt As String

10  On Error GoTo grdExt_Click_Error
20  grdExt.Col = grdExt.MouseCol
30  grdExt.Row = grdExt.MouseRow
40  If grdExt.MouseRow = 0 Then Exit Sub
50  If grdExt.Col = 1 Then
60      Prompt = "Enter result for " & grdExt.TextMatrix(grdExt.Row, 0)
70      Str = iBOX(Prompt, , grdExt.TextMatrix(grdExt.Row, 1))
80      If Str <> "" Then
90          grdExt.TextMatrix(grdExt.Row, 1) = Str
100         grdExt.TextMatrix(grdExt.Row, 6) = Format(Now, "dd/mmm/yyyy")
110     End If
120 ElseIf grdExt.Col = 7 Then
130     Prompt = "Enter Sap Code for " & grdExt.TextMatrix(grdExt.Row, 0)
140     Str = iBOX(Prompt, , grdExt.TextMatrix(grdExt.Row, 1))
150     If Str <> "" Then
160         grdExt.TextMatrix(grdExt.Row, 7) = Str
170     End If

180 End If
190 cmdSaveImm(2).Enabled = True
200 UpDown1.Enabled = False

210 Exit Sub

grdExt_Click_Error:

    Dim strES As String
    Dim intEL As Integer

220 intEL = Erl
230 strES = Err.Description
240 LogError "frmEditAll", "grdExt_Click", intEL, strES

End Sub

Private Sub grdH_Click()

10  On Error GoTo grdH_Click_Error

20  If grdH.Height = 360 Then
30      grdH.Height = 2000
40  End If

50  Exit Sub

grdH_Click_Error:

    Dim strES As String
    Dim intEL As Integer

60  intEL = Erl
70  strES = Err.Description
80  LogError "frmEditAll", "grdH_Click", intEL, strES

End Sub

Private Sub grdOutstanding_Click()

    Dim tb As New Recordset
    Dim sql As String



10  On Error GoTo grdOutstanding_Click_Error

20  With grdOutstanding
30      If .MouseRow = 0 Then Exit Sub
40      If .Text = "" Then Exit Sub
50      If iMsg("Remove " & .Text & " from Requests?", vbQuestion + vbYesNo) = vbYes Then
60          sql = "DELETE from BioRequests WHERE " & _
                  "SampleID = '" & txtSampleID & "' " & _
                  "and code = '" & CodeForShortName(.Text) & "'"
70          Set tb = New Recordset
80          RecOpenServer 0, tb, sql
90          If .Rows > 2 Then
100             .RemoveItem .Row
110         Else
120             .AddItem ""
130             .RemoveItem 1
140         End If
150     End If
160 End With



170 Exit Sub

grdOutstanding_Click_Error:

    Dim strES As String
    Dim intEL As Integer

180 intEL = Erl
190 strES = Err.Description
200 LogError "frmEditAll", "grdOutstanding_Click", intEL, strES

End Sub

Private Sub grdOutstandingCoag_Click()
    Dim tb As New Recordset
    Dim sql As String

10  On Error GoTo grdOutstandingCoag_Click_Error

20  With grdOutstandingCoag
30      If .MouseRow = 0 Then Exit Sub
40      If .Text = "" Then Exit Sub
50      If iMsg("Remove " & .Text & " from Requests?", vbQuestion + vbYesNo) = vbYes Then
60          sql = "DELETE from coagRequests WHERE " & _
                  "SampleID = '" & txtSampleID & "' " & _
                  "and code = '" & CoagCodeFor(.Text) & "'"
70          Set tb = New Recordset
80          RecOpenClient 0, tb, sql
90          If .Rows > 2 Then
100             .RemoveItem .Row
110         Else
120             .AddItem ""
130             .RemoveItem 1
140         End If
150     End If
160 End With

170 Exit Sub

grdOutstandingCoag_Click_Error:

    Dim strES As String
    Dim intEL As Integer

180 intEL = Erl
190 strES = Err.Description
200 LogError "frmEditAll", "grdOutstandingCoag_Click", intEL, strES, sql

End Sub

Private Sub grdOutstandings_Click(Index As Integer)


    Dim tb As New Recordset
    Dim sql As String


10  On Error GoTo grdOutstandings_Click_Error

20  If Index = 0 Then
30      With grdOutstandings(0)
40          If .MouseRow = 0 Then Exit Sub
50          If .Text = "" Then Exit Sub
60          If iMsg("Remove " & .Text & " from Requests?", vbQuestion + vbYesNo) = vbYes Then
70              sql = "DELETE from EndRequests WHERE " & _
                      "SampleID = '" & txtSampleID & "' " & _
                      "and code = '" & eCodeForShortName(.Text) & "'"
80              Set tb = New Recordset
90              RecOpenClient 0, tb, sql
100             If .Rows > 2 Then
110                 .RemoveItem .Row
120             Else
130                 .AddItem ""
140                 .RemoveItem 1
150             End If
160         End If
170     End With
180 Else
190     With grdOutstandings(1)
200         If .MouseRow = 0 Then Exit Sub
210         If .Text = "" Then Exit Sub
220         If iMsg("Remove " & .Text & " from Requests?", vbQuestion + vbYesNo) = vbYes Then
230             sql = "DELETE from ImmRequests WHERE " & _
                      "SampleID = '" & txtSampleID & "' " & _
                      "and code = '" & ICodeForShortName(.Text) & "'"
240             Set tb = New Recordset
250             RecOpenClient 0, tb, sql
260             If .Rows > 2 Then
270                 .RemoveItem .Row
280             Else
290                 .AddItem ""
300                 .RemoveItem 1
310             End If
320         End If
330     End With
340 End If

350 Exit Sub

grdOutstandings_Click_Error:

    Dim strES As String
    Dim intEL As Integer

360 intEL = Erl
370 strES = Err.Description
380 LogError "frmEditAll", "grdOutstandings_Click", intEL, strES, sql

End Sub

Private Sub grdOutstandings_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)

10  pBar = 0

End Sub

Private Sub Ig_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
10  On Error GoTo Ig_MouseUp_Error

20  If Index = 0 Then
30      EndChanged = True
40      cmdSaveImm(0).Enabled = True
50  Else
60      ImmChanged = True
70      cmdSaveImm(1).Enabled = True
80  End If

90  If Ig(Index).Value = 1 Then txtImmComment(Index) = Trim(txtImmComment(Index) & " " & Ig(Index).Caption)

100 Exit Sub

Ig_MouseUp_Error:

    Dim strES As String
    Dim intEL As Integer

110 intEL = Erl
120 strES = Err.Description
130 LogError "frmEditAll", "Ig_MouseUp", intEL, strES

End Sub

Private Sub Ih_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)

10  On Error GoTo Ih_MouseUp_Error

20  If Index = 0 Then
30      EndChanged = True
40      cmdSaveImm(0).Enabled = True

50  Else
60      ImmChanged = True
70      cmdSaveImm(1).Enabled = True
80  End If

90  If Ih(Index).Value = 1 Then txtImmComment(Index) = Trim(txtImmComment(Index) & " " & Ih(Index).Caption)

100 Exit Sub

Ih_MouseUp_Error:

    Dim strES As String
    Dim intEL As Integer

110 intEL = Erl
120 strES = Err.Description
130 LogError "frmEditAll", "Ih_MouseUp", intEL, strES

End Sub

Private Sub Iis_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)

10  On Error GoTo Iis_MouseUp_Error

20  If Index = 0 Then
30      EndChanged = True
40      cmdSaveImm(0).Enabled = True
50  Else
60      ImmChanged = True
70      cmdSaveImm(1).Enabled = True
80  End If

90  If Iis(Index).Value = 1 Then txtImmComment(Index) = Trim(txtImmComment(Index) & " " & Iis(Index).Caption)

100 Exit Sub

Iis_MouseUp_Error:

    Dim strES As String
    Dim intEL As Integer

110 intEL = Erl
120 strES = Err.Description
130 LogError "frmEditAll", "Iis_MouseUp", intEL, strES

End Sub

Private Sub Ij_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)

10  On Error GoTo Ij_MouseUp_Error

20  If Index = 0 Then
30      EndChanged = True
40      cmdSaveImm(0).Enabled = True
50  Else
60      ImmChanged = True
70      cmdSaveImm(1).Enabled = True
80  End If

90  If Ij(Index).Value = 1 Then txtImmComment(Index) = Trim(txtImmComment(Index) & " " & Ij(Index).Caption)

100 Exit Sub

Ij_MouseUp_Error:

    Dim strES As String
    Dim intEL As Integer

110 intEL = Erl
120 strES = Err.Description
130 LogError "frmEditAll", "Ij_MouseUp", intEL, strES

End Sub

Private Sub Il_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
10  On Error GoTo Il_MouseUp_Error

20  If Index = 0 Then
30      EndChanged = True
40      cmdSaveImm(0).Enabled = True
50  Else
60      ImmChanged = True
70      cmdSaveImm(1).Enabled = True
80  End If

90  If Il(Index).Value = 1 Then txtImmComment(Index) = Trim(txtImmComment(Index) & " " & Il(Index).Caption)

100 Exit Sub

Il_MouseUp_Error:

    Dim strES As String
    Dim intEL As Integer

110 intEL = Erl
120 strES = Err.Description
130 LogError "frmEditAll", "Il_MouseUp", intEL, strES

End Sub

Private Sub cmdGreenTick_Click(Index As Integer)

    Dim Y As Integer

10  Select Case Index
    Case 0    'Imm
20      gImm(1).Col = 9
30      For Y = 1 To gImm(1).Rows - 1
40          gImm(1).Row = Y
50          Set gImm(1).CellPicture = imgGreenTick.Picture
60      Next
70  Case 1    'Bio
80      gBio.Col = 10
90      For Y = 1 To gBio.Rows - 1
100         gBio.Row = Y
110         If gBio.CellBackColor <> vbRed Then
120             Set gBio.CellPicture = imgGreenTick.Picture
130         End If
140     Next
150 Case 2:    'Coag
160     grdCoag.Col = 8
170     For Y = 1 To grdCoag.Rows - 1
180         grdCoag.Row = Y
190         Set grdCoag.CellPicture = imgGreenTick.Picture
200     Next
210 Case 3    'End
220     gImm(0).Col = 8
230     For Y = 1 To gImm(0).Rows - 1
240         gImm(0).Row = Y
250         Set gImm(0).CellPicture = imgGreenTick.Picture
260     Next
270 End Select

End Sub

Private Sub imgLast_Click()

    Dim sql As String
    Dim tb As New Recordset
    Dim strDept As String
    Dim strSplitSELECT As String




10  On Error GoTo imgLast_Click_Error

20  Select Case ssTabAll.Tab
    Case 0:
30      txtSampleID = Format$(Val(txtSampleID) + 1)
40      LoadAllDetails

50      cmdSaveDemographics.Enabled = False
60      cmdSaveInc.Enabled = False
70      cmdSaveHaem.Enabled = False
80      cmdSaveComm.Enabled = False
90      cmdHSaveH.Enabled = False
100     cmdSaveBio.Enabled = False
110     cmdSaveCoag.Enabled = False
120     cmdSaveImm(0).Enabled = False
130     cmdSaveImm(1).Enabled = False
140     cmdSaveBGa.Enabled = False
150     Exit Sub

160 Case 1: strDept = "Haem"
170 Case 2: strDept = "Bio"
180 Case 3: strDept = "Coag"
190 Case 4: strDept = "End"
200 Case 5: strDept = "Bga"
210 Case 6: strDept = "Imm"
220 Case 7: strDept = "Ext"
230 End Select

240 sql = "SELECT top 1 SampleID from " & strDept & "Results "
250 If HospName(0) = "PORTLAOISE" Then
260     sql = sql & "WHERE sampleid < 9000000 "
270 End If
280 sql = sql & "Order by SampleID desc"

290 If strDept = "Bio" Then
300     If InStr(lblViewSplit, "Pri") Then
310         strSplitSELECT = LoadSplitList(1)
320     ElseIf InStr(lblViewSplit, "Sec") Then
330         strSplitSELECT = LoadSplitList(2)
340     End If
350     If strSplitSELECT <> "" Then
360         sql = "SELECT top 1 SampleID from " & strDept & "Results WHERE " & _
                  "(" & strSplitSELECT & ") "
370         If HospName(0) = "PORTLAOISE" Then
380             sql = sql & "and sampleid < 9000000 "
390         End If
400         sql = sql & "Order by SampleID desc"
410     End If
420 ElseIf strDept = "Imm" Then
430     If InStr(lblImmViewSplit(1), "Pri") Then
440         strSplitSELECT = LoadImmSplitList(1)
450     ElseIf InStr(lblImmViewSplit(1), "Sec") Then
460         strSplitSELECT = LoadImmSplitList(2)
470     End If
480     If strSplitSELECT <> "" Then
490         sql = "SELECT top 1 SampleID from " & strDept & "Results WHERE " & _
                  "(" & strSplitSELECT & ") "
500         If HospName(0) = "PORTLAOISE" Then
510             sql = sql & "and sampleid < 9000000 "
520         End If
530         sql = sql & "Order by SampleID desc"
540     End If

550 End If

560 Set tb = New Recordset
570 RecOpenServer 0, tb, sql
580 If Not tb.EOF Then
590     txtSampleID = tb!SampleID & ""
600 End If

610 LoadAllDetails

620 cmdSaveDemographics.Enabled = False
630 cmdSaveInc.Enabled = False
640 cmdSaveHaem.Enabled = False
650 cmdSaveComm.Enabled = False
660 cmdHSaveH.Enabled = False
670 cmdSaveBio.Enabled = False
680 cmdSaveCoag.Enabled = False
690 cmdSaveImm(0).Enabled = False
700 cmdSaveImm(1).Enabled = False
710 cmdSaveBGa.Enabled = False

720 Exit Sub

imgLast_Click_Error:

    Dim strES As String
    Dim intEL As Integer

730 intEL = Erl
740 strES = Err.Description
750 LogError "frmEditAll", "imgLast_Click", intEL, strES, sql

End Sub

Private Sub cmdRedCross_Click(Index As Integer)

    Dim Y As Integer

10  Select Case Index
    Case 0    'Imm
20      gImm(1).Col = 9
30      For Y = 1 To gImm(1).Rows - 1
40          gImm(1).Row = Y
50          Set gImm(1).CellPicture = imgRedCross.Picture
60      Next
70  Case 1    'Bio
80      gBio.Col = 10
90      For Y = 1 To gBio.Rows - 1
100         gBio.Row = Y
110         Set gBio.CellPicture = imgRedCross.Picture
120     Next
130 Case 2:    'Coag
140     grdCoag.Col = 8
150     For Y = 1 To grdCoag.Rows - 1
160         grdCoag.Row = Y
170         Set grdCoag.CellPicture = imgRedCross.Picture
180     Next
190 Case 3    'Imm
200     gImm(0).Col = 8
210     For Y = 1 To gImm(0).Rows - 1
220         gImm(0).Row = Y
230         Set gImm(0).CellPicture = imgRedCross.Picture
240     Next

250 End Select

End Sub

Private Sub Io_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)

10  On Error GoTo Io_MouseUp_Error

20  If Index = 0 Then
30      EndChanged = True
40      cmdSaveImm(0).Enabled = True
50  Else
60      ImmChanged = True
70      cmdSaveImm(1).Enabled = True
80  End If

90  If Io(Index).Value = 1 Then txtImmComment(Index) = Trim(txtImmComment(Index) & " " & Io(Index).Caption)

100 Exit Sub

Io_MouseUp_Error:

    Dim strES As String
    Dim intEL As Integer

110 intEL = Erl
120 strES = Err.Description
130 LogError "frmEditAll", "Io_MouseUp", intEL, strES

End Sub

Private Sub iRecDate_Click(Index As Integer)

10  On Error GoTo iRecDate_Click_Error

20  If Index = 0 Then
30      dtRecDate = DateAdd("d", -1, dtRecDate)
40  Else
50      If DateDiff("d", dtRecDate, Now) > 0 Then
60          dtRecDate = DateAdd("d", 1, dtRecDate)
70      End If
80  End If

90  SetDatesColour Me

100 cmdSaveInc.Enabled = True
110 cmdSaveDemographics.Enabled = True

120 Exit Sub

iRecDate_Click_Error:

    Dim strES As String
    Dim intEL As Integer

130 intEL = Erl
140 strES = Err.Description
150 LogError "frmEditAll", "iRecDate_Click", intEL, strES

End Sub

Private Sub irelevant_Click(Index As Integer)

    Dim sql As String
    Dim tb As New Recordset
    Dim strDept As String
    Dim strDirection As String
    Dim strSplitSELECT As String
    Dim strArrow As String



10  On Error GoTo irelevant_Click_Error

20  If txtSampleID = "" Then Exit Sub

30  If cmdSaveImm(2).Enabled Then
40      iMsg "External Save Enabled!"
50      Exit Sub
60  End If


70  Select Case ssTabAll.Tab
    Case 0:
80      If Index = 0 Then
90          txtSampleID = Format$(Val(txtSampleID) - 1)
100     Else
110         txtSampleID = Format$(Val(txtSampleID) + 1)
120     End If

130     If SysOptNumLen(0) > 0 Then
140         If Len(txtSampleID) > SysOptNumLen(0) Then
150             iMsg "Sample Id longer then recommended!"
160         End If
170     End If

180     LoadAllDetails

190     cmdSaveDemographics.Enabled = False
200     cmdSaveInc.Enabled = False
210     cmdSaveHaem.Enabled = False
220     cmdSaveComm.Enabled = False
230     cmdHSaveH.Enabled = False
240     cmdSaveBio.Enabled = False
250     cmdSaveCoag.Enabled = False
260     cmdSaveImm(0).Enabled = False
270     cmdSaveImm(1).Enabled = False
280     cmdSaveBGa.Enabled = False
290     Exit Sub

300 Case 1: strDept = "Haem"
310 Case 2: strDept = "Bio"
320 Case 3: strDept = "Coag"
330 Case 4: strDept = "End"
340 Case 5: strDept = "Bga"
350 Case 6: strDept = "Imm"
360 Case 7: strDept = "Ext"
370 End Select

380 strDirection = IIf(Index = 0, "Desc", "Asc")
390 strArrow = IIf(Index = 0, "<", ">")


400 If lblResultOrRequest = "Film" And strDept <> "Haem" Then lblResultOrRequest = "Results"

410 If lblResultOrRequest = "Results" Then
420     sql = "SELECT top 1 SampleID from " & strDept & "Results WHERE " & _
              "SampleID " & strArrow & " " & txtSampleID & " " & _
              "Order by SampleID " & strDirection
430 ElseIf ssTabAll.Tab = 7 Then    'ext
440     sql = "SELECT top 1 SampleID from " & strDept & "Results WHERE " & _
              "SampleID " & strArrow & " " & txtSampleID & " " & _
              "Order by SampleID " & strDirection
450 ElseIf lblResultOrRequest = "Requests" Then
460     sql = "SELECT top 1 SampleID from " & strDept & "Requests WHERE " & _
              "SampleID " & strArrow & " " & txtSampleID & " " & _
              "Order by SampleID " & strDirection
470 Else
480     sql = "SELECT top 1 SampleID from " & strDept & "Results WHERE " & _
              "SampleID " & strArrow & " " & txtSampleID & " and valid <> 1 " & _
              "Order by SampleID " & strDirection
490 End If

500 If strDept = "Bio" Then
510     If InStr(lblViewSplit, "Pri") Then
520         strSplitSELECT = LoadSplitList(1)
530     ElseIf InStr(lblViewSplit, "Sec") Then
540         strSplitSELECT = LoadSplitList(2)
550     End If
560     If lblResultOrRequest = "Results" Then
570         If strSplitSELECT <> "" Then
580             sql = "SELECT top 1 SampleID from BioResults WHERE " & _
                      "SampleID " & strArrow & " " & txtSampleID & " " & _
                      "and (" & strSplitSELECT & ") " & _
                      "Order by SampleID " & strDirection
590         End If
600     Else
610         If strSplitSELECT <> "" Then
620             sql = "SELECT top 1 SampleID from BioRequests WHERE " & _
                      "SampleID " & strArrow & " " & txtSampleID & " " & _
                      "and (" & strSplitSELECT & ") " & _
                      "Order by SampleID " & strDirection
630         End If
640     End If
650 ElseIf strDept = "Imm" Then
660     If InStr(lblImmViewSplit(1), "Pri") Then
670         strSplitSELECT = LoadImmSplitList(1)
680     ElseIf InStr(lblImmViewSplit(1), "Sec") Then
690         strSplitSELECT = LoadImmSplitList(2)
700     End If
710     If lblResultOrRequest = "Results" Then
720         If strSplitSELECT <> "" Then
730             sql = "SELECT top 1 SampleID from ImmResults WHERE " & _
                      "SampleID " & strArrow & " " & txtSampleID & " " & _
                      "and (" & strSplitSELECT & ") " & _
                      "Order by SampleID " & strDirection
740         End If
750     Else
760         If strSplitSELECT <> "" Then
770             sql = "SELECT top 1 SampleID from ImmRequests WHERE " & _
                      "SampleID " & strArrow & " " & txtSampleID & " " & _
                      "and (" & strSplitSELECT & ") " & _
                      "Order by SampleID " & strDirection
780         End If
790     End If
800 ElseIf strDept = "End" Then
810     If InStr(lblImmViewSplit(0), "Pri") Then
820         strSplitSELECT = LoadEndSplitList(1)
830     ElseIf InStr(lblImmViewSplit(0), "Sec") Then
840         strSplitSELECT = LoadEndSplitList(2)
850     End If
860     If lblResultOrRequest = "Results" Then
870         If strSplitSELECT <> "" Then
880             sql = "SELECT top 1 SampleID from endResults WHERE " & _
                      "SampleID " & strArrow & " " & txtSampleID & " " & _
                      "and (" & strSplitSELECT & ") " & _
                      "Order by SampleID " & strDirection
890         End If
900     Else
910         If strSplitSELECT <> "" Then
920             sql = "SELECT top 1 SampleID from endRequests WHERE " & _
                      "SampleID " & strArrow & " " & txtSampleID & " " & _
                      "and (" & strSplitSELECT & ") " & _
                      "Order by SampleID " & strDirection
930         End If
940     End If
950 ElseIf strDept = "Haem" And lblResultOrRequest = "Film" Then
960     sql = "SELECT top 1 SampleID from HaemResults WHERE " & _
              "SampleID " & strArrow & " " & txtSampleID & " " & _
              "and cfilm = 1 " & _
              "Order by SampleID " & strDirection
970 End If

980 Set tb = New Recordset
990 RecOpenServer 0, tb, sql
1000 If Not tb.EOF Then
1010    txtSampleID = tb!SampleID & ""
1020 End If

1030 LoadAllDetails

1040 cmdSaveDemographics.Enabled = False
1050 cmdSaveInc.Enabled = False
1060 cmdSaveHaem.Enabled = False
1070 cmdSaveComm.Enabled = False
1080 cmdHSaveH.Enabled = False
1090 cmdSaveBio.Enabled = False
1100 cmdSaveCoag.Enabled = False
1110 cmdSaveImm(0).Enabled = False
1120 cmdSaveImm(1).Enabled = False
1130 cmdSaveBGa.Enabled = False

1140 Exit Sub

irelevant_Click_Error:

    Dim strES As String
    Dim intEL As Integer

1150 intEL = Erl
1160 strES = Err.Description
1170 LogError "frmEditAll", "irelevant_Click", intEL, strES, sql


End Sub

Private Sub iRunDate_Click(Index As Integer)

10  On Error GoTo iRunDate_Click_Error

20  If Index = 0 Then
30      dtRunDate = DateAdd("d", -1, dtRunDate)
40  Else
50      If DateDiff("d", dtRunDate, Now) > 0 Then
60          dtRunDate = DateAdd("d", 1, dtRunDate)
70      End If
80  End If

90  SetDatesColour Me

100 cmdSaveInc.Enabled = True
110 cmdSaveDemographics.Enabled = True

120 Exit Sub

iRunDate_Click_Error:

    Dim strES As String
    Dim intEL As Integer

130 intEL = Erl
140 strES = Err.Description
150 LogError "frmEditAll", "iRunDate_Click", intEL, strES

End Sub

Private Sub iSampleDate_Click(Index As Integer)

10  On Error GoTo iSampleDate_Click_Error

20  If Index = 0 Then
30      dtSampleDate = DateAdd("d", -1, dtSampleDate)
40  Else
50      If DateDiff("d", dtSampleDate, Now) > 0 Then
60          dtSampleDate = DateAdd("d", 1, dtSampleDate)
70      End If
80  End If

90  SetDatesColour Me

100 cmdSaveInc.Enabled = True
110 cmdSaveDemographics.Enabled = True

120 Exit Sub

iSampleDate_Click_Error:

    Dim strES As String
    Dim intEL As Integer

130 intEL = Erl
140 strES = Err.Description
150 LogError "frmEditAll", "iSampleDate_Click", intEL, strES

End Sub

Private Function IsControl(ByVal Chart As String) As Boolean

    Dim n As Long

10  On Error GoTo IsControl_Error

20  IsControl = False

30  If Trim(Chart) <> "" Then
40      For n = 0 To UBound(ControlName)
50          If Trim(UCase(Chart)) = UCase(ControlName(n)) Then
60              IsControl = True
70              Exit For
80          End If
90      Next
100 End If

110 Exit Function

IsControl_Error:

    Dim strES As String
    Dim intEL As Integer

120 intEL = Erl
130 strES = Err.Description
140 LogError "frmEditAll", "IsControl", intEL, strES

End Function

Private Sub iToday_Click(Index As Integer)

10  On Error GoTo iToday_Click_Error

20  If Index = 0 Then
30      dtRunDate = Format$(Now, "dd/mm/yyyy")
40  ElseIf Index = 1 Then
50      If DateDiff("d", dtRunDate, Now) > 0 Then
60          dtSampleDate = dtRunDate
70      Else
80          dtSampleDate = Format$(Now, "dd/mm/yyyy")
90      End If
100 ElseIf Index = 2 Then
110     If DateDiff("d", dtRunDate, Now) > 0 Then
120         dtRecDate = dtRunDate
130     Else
140         dtRecDate = Format$(Now, "dd/mm/yyyy")
150     End If
160 End If

170 SetDatesColour Me

180 cmdSaveInc.Enabled = True
190 cmdSaveDemographics.Enabled = True

200 Exit Sub

iToday_Click_Error:

    Dim strES As String
    Dim intEL As Integer

210 intEL = Erl
220 strES = Err.Description
230 LogError "frmEditAll", "iToday_Click", intEL, strES

End Sub



Private Sub lblAss_Click()
    Dim Num As Long
    Dim Numx As Long

10  On Error GoTo lblAss_Click_Error

20  For Num = Len(lblAss) To 1 Step -1
30      If Mid(lblAss, Num, 1) = " " Then
40          Numx = Num
50          Exit For
60      End If
70  Next

80  txtSampleID = Trim(Mid(lblAss, Numx))
90  txtSampleID_LostFocus

100 Exit Sub

lblAss_Click_Error:

    Dim strES As String
    Dim intEL As Integer

110 intEL = Erl
120 strES = Err.Description
130 LogError "frmEditAll", "lblAss_Click", intEL, strES

End Sub

Private Sub lblChartNumber_Click()

10  On Error GoTo lblChartNumber_Click_Error

20  With lblChartNumber
30      If InStr(.Caption, HospName(0)) = 0 Then
40          .BackColor = vbRed
50          .ForeColor = vbYellow
60      Else
70          .BackColor = &H8000000F
80          .ForeColor = vbBlack
90      End If

100 End With

110 If Trim$(txtChart) <> "" Then
120     LoadPatientFromChart frmEditAll, True
130     cmdSaveDemographics.Enabled = True
140     cmdSaveInc.Enabled = True
150 End If

160 Exit Sub

lblChartNumber_Click_Error:

    Dim strES As String
    Dim intEL As Integer

170 intEL = Erl
180 strES = Err.Description
190 LogError "frmEditAll", "lblChartNumber_Click", intEL, strES

End Sub

Private Sub lblImmViewSplit_Click(Index As Integer)


10  On Error GoTo lblImmViewSplit_Click_Error

20  If Index = 0 Then
30      With lblImmViewSplit(0)
40          Select Case .Caption
            Case "Viewing All":
50              .Caption = "Viewing Primary Split"
60              .BackColor = &H800080
70              .ForeColor = &HFF00&
80          Case "Viewing Primary Split":
90              .Caption = "Viewing Secondary Split"
100             .BackColor = &H800080
110             .ForeColor = &HFF00&
120         Case "Viewing Secondary Split":
130             .Caption = "Viewing All"
140             .BackColor = &H8000000F
150             .ForeColor = vbBlack
160         End Select
170     End With
180 Else
190     With lblImmViewSplit(1)
200         Select Case .Caption
            Case "Viewing All":
210             .Caption = "Viewing Primary Split"
220             .BackColor = &H800080
230             .ForeColor = &HFF00&
240         Case "Viewing Primary Split":
250             .Caption = "Viewing Secondary Split"
260             .BackColor = &H800080
270             .ForeColor = &HFF00&
280         Case "Viewing Secondary Split":
290             .Caption = "Viewing All"
300             .BackColor = &H8000000F
310             .ForeColor = vbBlack
320         End Select
330     End With
340 End If

350 Exit Sub

lblImmViewSplit_Click_Error:

    Dim strES As String
    Dim intEL As Integer

360 intEL = Erl
370 strES = Err.Description
380 LogError "frmEditAll", "lblImmViewSplit_Click", intEL, strES

End Sub

Private Sub lblMalaria_Change()

10  On Error GoTo lblMalaria_Change_Error

20  cmdSaveHaem.Enabled = True: cmdHSaveH.Enabled = True

30  If Trim$(lblMalaria) <> "" Then
40      chkMalaria = 1
50  Else
60      chkMalaria = 0
70  End If

80  Exit Sub

lblMalaria_Change_Error:

    Dim strES As String
    Dim intEL As Integer

90  intEL = Erl
100 strES = Err.Description
110 LogError "frmEditAll", "lblMalaria_Change", intEL, strES

End Sub

Private Sub lblMalaria_Click()

10  On Error GoTo lblMalaria_Click_Error

20  If lblMalaria = "" Then
30      lblMalaria = "Positive"
40  ElseIf lblMalaria = "Positive" Then
50      lblMalaria = "Negative"
60  ElseIf lblMalaria = "Negative" Then
70      lblMalaria = "Inconclusive"
80  ElseIf lblMalaria = "Inconclusive" Then
90      lblMalaria = ""
100 End If

110 Exit Sub

lblMalaria_Click_Error:

    Dim strES As String
    Dim intEL As Integer

120 intEL = Erl
130 strES = Err.Description
140 LogError "frmEditAll", "lblMalaria_Click", intEL, strES

End Sub

Private Sub lblResultOrRequest_Click()

10  On Error GoTo lblResultOrRequest_Click_Error

20  If ssTabAll.Tab <> 0 Then
30      If lblResultOrRequest = "Results" Then
40          lblResultOrRequest = "Request"
50      ElseIf lblResultOrRequest = "Request" Then
60          lblResultOrRequest = "UnValid"
70      ElseIf ssTabAll.Tab = 1 And lblResultOrRequest = "UnValid" Then
80          lblResultOrRequest = "Film"
90      Else
100         lblResultOrRequest = "Results"
110     End If
120 End If

130 Exit Sub

lblResultOrRequest_Click_Error:

    Dim strES As String
    Dim intEL As Integer

140 intEL = Erl
150 strES = Err.Description
160 LogError "frmEditAll", "lblResultOrRequest_Click", intEL, strES


End Sub

Private Sub lblSickledex_Change()

10  On Error GoTo lblSickledex_Change_Error

20  cmdSaveHaem.Enabled = True: cmdHSaveH.Enabled = True

30  If Trim$(lblSickledex) <> "" Then
40      chkSickledex = 1
50  Else
60      chkSickledex = 0
70  End If

80  Exit Sub

lblSickledex_Change_Error:

    Dim strES As String
    Dim intEL As Integer

90  intEL = Erl
100 strES = Err.Description
110 LogError "frmEditAll", "lblSickledex_Change", intEL, strES

End Sub

Private Sub lblSickledex_Click()

10  On Error GoTo lblSickledex_Click_Error

20  If lblSickledex = "" Then
30      lblSickledex = "Positive"
40  ElseIf lblSickledex = "Positive" Then
50      lblSickledex = "Negative"
60  ElseIf lblSickledex = "Negative" Then
70      lblSickledex = "Inconclusive"
80  ElseIf lblSickledex = "Inconclusive" Then
90      lblSickledex = ""
100 End If

110 Exit Sub

lblSickledex_Click_Error:

    Dim strES As String
    Dim intEL As Integer

120 intEL = Erl
130 strES = Err.Description
140 LogError "frmEditAll", "lblSickledex_Click", intEL, strES

End Sub

Private Sub lblUrgent_Click()
    Dim sql As String

10  On Error GoTo lblUrgent_Click_Error

20  lblUrgent.Visible = False

30  sql = "UPDATE demographics set urgent = 0 WHERE sampleid = '" & txtSampleID & "'"
40  Cnxn(0).Execute sql

50  Exit Sub

lblUrgent_Click_Error:

    Dim strES As String
    Dim intEL As Integer

60  intEL = Erl
70  strES = Err.Description
80  LogError "frmEditAll", "lblUrgent_Click", intEL, strES, sql


End Sub

Private Sub lblViewSplit_Click()



10  On Error GoTo lblViewSplit_Click_Error

20  With lblViewSplit
30      Select Case .Caption
        Case "Viewing All":
40          .Caption = "Viewing Primary Split"
50          .BackColor = &H800080
60          .ForeColor = &HFF00&
70      Case "Viewing Primary Split":
80          .Caption = "Viewing Secondary Split"
90          .BackColor = &H800080
100         .ForeColor = &HFF00&
110     Case "Viewing Secondary Split":
120         .Caption = "Viewing All"
130         .BackColor = &H8000000F
140         .ForeColor = vbBlack
150     End Select
160 End With

170 Exit Sub

lblViewSplit_Click_Error:

    Dim strES As String
    Dim intEL As Integer

180 intEL = Erl
190 strES = Err.Description
200 LogError "frmEditAll", "lblViewSplit_Click", intEL, strES

End Sub

Private Sub lHaemErrors_Click()

10  On Error GoTo lHaemErrors_Click_Error

20  Unload frmHaemErrors

30  With frmHaemErrors
40      .Analyser = HaemAnalyser
50      .ErrorNumber = lHaemErrors.Tag
60      .Show 1
70  End With

80  Exit Sub

lHaemErrors_Click_Error:

    Dim strES As String
    Dim intEL As Integer

90  intEL = Erl
100 strES = Err.Description
110 LogError "frmEditAll", "lHaemErrors_Click", intEL, strES


End Sub

Private Sub lImmRan_Click(Index As Integer)

10  On Error GoTo lImmRan_Click_Error

20  If Index = 0 Then
30      If lImmRan(0) = "Random Sample" Then
40          lImmRan(0) = "Fasting Sample"
50      Else
60          lImmRan(0) = "Random Sample"
70      End If

80      LoadEndocrinology

90      cmdSaveImm(0).Enabled = True
100 Else
110     If lImmRan(1) = "Random Sample" Then
120         lImmRan(1) = "Fasting Sample"
130     Else
140         lImmRan(1) = "Random Sample"
150     End If

160     LoadImmunology

170     cmdSaveImm(1).Enabled = True
180 End If

190 Exit Sub

lImmRan_Click_Error:

    Dim strES As String
    Dim intEL As Integer

200 intEL = Erl
210 strES = Err.Description
220 LogError "frmEditAll", "lImmRan_Click", intEL, strES


End Sub

Private Sub LoadAllDetails()

10        On Error GoTo LoadAllDetails_Error

20        HaemLoaded = False
30        BioLoaded = False
40        CoagLoaded = False
50        ImmLoaded = False
60        BgaLoaded = False
70        ExtLoaded = False
80        EndLoaded = False

90        cAdd = ""
100       cUnits = ""
110       tnewvalue = ""

120       cIAdd(0) = ""
130       cIUnits(0) = ""
140       tINewValue(0) = ""

150       cIAdd(1) = ""
160       cIUnits(1) = ""
170       tINewValue(1) = ""

180       ClearDemographics
190       ClearHaematologyResults
200       ClearCoagulation
210       ClearOutstanding grdOutstandings(1)
220       ClearOutstanding grdOutstanding
230       ClearImmFlags
240       ClearEndFlags
          'ClearBga
250       ClearExt

260       ssTabAll.TabCaption(1) = "Haematology"
270       ssTabAll.TabCaption(2) = "Biochemistry"
280       ssTabAll.TabCaption(3) = "Coagulation"
290       ssTabAll.TabCaption(4) = "Endocrinology"
300       ssTabAll.TabCaption(5) = "Blood Gas"
310       ssTabAll.TabCaption(6) = "Immunology"
320       ssTabAll.TabCaption(7) = "Externals"

          'SetSampleType
330       LoadDemographics
340       CheckDepartments
          'LoadComments

350       CheckIfPhoned

360       Select Case ssTabAll.Tab
              Case 0:
370           Case 1: LoadHaematology
380               HaemLoaded = True
390           Case 2: LoadBiochemistry
400               BioLoaded = True
410           Case 3: LoadCoagulation
420               CoagLoaded = True
430           Case 4: LoadEndocrinology

440               EndLoaded = True
450           Case 5: LoadBloodGas
460               BgaLoaded = True
470           Case 6: LoadImmunology
480               ImmLoaded = True
490           Case 7: LoadExt
500               ExtLoaded = True
510       End Select

520       LoadComments
530       CheckPatientNotePad (Trim$(txtSampleID))

540       SetViewHistory

550       cmdSaveHaem.Enabled = False
560       cmdSaveComm.Enabled = False
570       cmdHSaveH.Enabled = False
580       cmdSaveBio.Enabled = False
590       cmdSaveImm(0).Enabled = False
600       cmdSaveImm(1).Enabled = False

          'SetDefaultSampleType
610       CheckAuditTrail
620       EnableBarCodePrinting
630       CheckLabLinkStatus
640       Exit Sub

LoadAllDetails_Error:

          Dim strES As String
          Dim intEL As Integer

650       intEL = Erl
660       strES = Err.Description
670       LogError "frmEditAll", "LoadAllDetails", intEL, strES

End Sub

Public Sub LoadBiochemistry()

    Dim tb As New Recordset
    Dim sql As String
    Dim s As String
    Dim Value As Single
    Dim valu As String
    Dim n As Long
    Dim e As String
    Dim SampleType As String
    Dim BRs As New BIEResults
    Dim BRres As BIEResults
    Dim br As BIEResult
    Dim Fasting As Boolean
    Dim Flag As String
    Dim T As String
    Dim Code As String
    Dim CodeTb As Recordset
    Dim FormatedResult As String

10  On Error GoTo LoadBiochemistry_Error


20  s = CheckAutoComments(txtSampleID, "Biochemistry")
    'If InStr(txtBioComment, s) = 0 And UCase(bValidateBio.Caption) = "&VALIDATE" Then
    '    txtBioComment = s & vbCrLf & txtBioComment
    'End If

30  HistBio = False

40  If txtName <> "" And txtDoB <> "" And IsDate(txtDoB) Then
50      sql = CreateHist("Bio")
60      Set tb = New Recordset
70      RecOpenClient 0, tb, sql
80      If Not tb.EOF Then
90          HistBio = True
100     End If
110 End If

120 If txtSampleID = "" Then Exit Sub

130 Frame2.Enabled = True
140 lRandom.Enabled = True
150 txtBioComment.Locked = False

160 Fasting = lRandom = "Fasting Sample"
170 lblAss.Visible = False

180 ClearFGrid gBio

190 oH = 0
200 oS = 0
210 oL = 0
220 oO = 0
230 oG = 0
240 oJ = 0
250 lBDate = ""
260 ldelta = ""
270 bViewBioRepeat.Visible = False

280 ssTabAll.TabCaption(2) = "Biochemistry"

290 gBio.Visible = False

300 Set BRres = BRs.Load("Bio", txtSampleID, "Results", gDONTCARE, gDONTCARE, 0, cCat(0), dtRunDate)

310 If Not BRres Is Nothing Then
320     If SysOptDoAssGlucose(0) Then
330         CheckAssGlucose BRres
340     End If
350     CheckCalcPSA BRres
360     If CheckEGFR(BRres) Or CheckGPCR(BRres) Then
370         Set BRres = BRs.Load("Bio", txtSampleID, "Results", gDONTCARE, gDONTCARE, 0, cCat(0), dtRunDate)
380     End If
390     If SysOptCheckCholHDLRatio(0) Then CheckCholHDL BRres
400 End If

410 If IsControl(txtChart) Then
420     ldelta = ""
430     gBio.Rows = 2
440     gBio.AddItem ""
450     gBio.RemoveItem 1
460     If Not BRres Is Nothing Then
470         For Each br In BRres
                '    CheckBioNormalStatus br
480             Code = br.Code

490             If GetOptionSetting("BioAn1", "") = "ROCHE" Then
500                 If Code = 91 Then
510                     If Val(br.Result) > 200 Then
520                         oG.Value = True
530                     ElseIf Val(br.Result) > 80 Then
540                         oH.Value = True
550                     ElseIf Val(br.Result) > 30 Then
560                         oS.Value = True
570                     End If
580                 End If
590             End If
600             s = br.LongName & vbTab
610             If IsNumeric(br.Result) Then
620                 If Not IsNull(br.Result) Then
630                     Value = br.Result
640                 Else
650                     Value = 0
660                 End If
670                 If Value <= 1 Then
680                     valu = Format(Value, "0.00")
690                 ElseIf Value > 1 And Value <= 10 Then
700                     valu = Format(Value, "0.0")
710                 Else
720                     valu = Format(Value)
730                 End If
740             Else
750                 valu = br.Result
760             End If

770             s = s & valu & vbTab
780             sql = "SELECT * from controls WHERE controlname = '" & txtChart & "' and parameter = '" & Code & "'"
790             Set CodeTb = New Recordset
800             RecOpenServer 0, CodeTb, sql
810             If Not CodeTb.EOF Then
820                 If Not IsNull(CodeTb!mean) And Not IsNull(CodeTb("1sd")) Then
830                     s = s & InterC(Value, CodeTb!mean - CodeTb("1sd") * 2, CodeTb!mean + CodeTb("1sd") * 2) & vbTab & _
                            (CodeTb!mean - CodeTb("1sd") * 2) & _
                          "  -  " & _
                            (CodeTb!mean + CodeTb("1sd") * 2) & vbTab & vbTab & vbTab & vbTab
840                 End If
850             End If
860             s = s & br.Pc & vbTab
870             Select Case Trim(br.Analyser)
                Case "4": s = s & "Immuno"
880             Case "A": s = s & "Bio (A)"
890             Case "B": s = s & "Bio (B)"
900             Case "P1": s = s & SysOptBioN1(0)
910             Case "P2": s = s & SysOptBioN2(0)
920             Case Else: s = s & "General"
930             End Select
940             s = s & vbTab & br.Comment

950             gBio.AddItem s
960             If br.Printable = False Then
970                 gBio.Row = gBio.Rows - 1
980                 gBio.Col = 10
990                 gBio.CellBackColor = vbRed
1000            End If
1010        Next
1020    End If
1030    gBio.Visible = True
1040    If gBio.Rows > 2 Then
1050        gBio.RemoveItem 1
1060    End If
1070    SetPrintInhibit "Bio"
1080    Exit Sub
1090 End If


1100 If Not BRres Is Nothing Then
1110    ssTabAll.TabCaption(2) = ">>Biochemistry<<"
1120    SampleType = ""
1130    For Each br In BRres
1140        Flag = ""
1150        If SampleType = "" Then
1160            SampleType = br.SampleType
1170            cISampleType(3) = ListText("ST", br.SampleType)
1180            If Len(SampleType) = 0 Then
1190                SampleType = "S"
1200            End If
1210        End If
1220        SampleType = br.SampleType
1230        s = br.ShortName & vbTab
1240        lBDate = Format(GetLatestRunDateTime("Bio", br.SampleID, br.RunTime), "dd/MM/yyyy hh:mm:ss")
1250        If IsNumeric(br.Result) Then
1260            Value = Val(br.Result)
1270            Select Case br.Printformat
                Case 0: valu = Format$(Value, "######")
1280            Case 1: valu = Format$(Value, "###0.0")
1290            Case 2: valu = Format$(Value, "##0.00")
1300            Case 3: valu = Format$(Value, "#0.000")
1310            Case Else: valu = Format$(Value, "0.000")
1320            End Select
1330        Else
1340            valu = br.Result
1350        End If
1360        s = s & valu & vbTab
1370        If ListText("UN", br.Units) <> "" Then
1380            s = s & ListText("UN", br.Units)
1390        Else
1400            s = s & br.Units
1410        End If
1420        s = s & vbTab
1430        T = ""
1440        If Trim(UCase(br.Code)) = "418" Or txtSex = "" Then   'QMS Ref No. #817982
1450            s = s & "" & vbTab
1460        Else

1470            If IsNumeric(br.Result) Then
1480                If Value > Val(br.PlausibleHigh) Then
1490                    Flag = "X"
1500                    s = s & br.Low & " - " & br.High & vbTab
1510                    s = s & "X"
1520                ElseIf Value < Val(br.PlausibleLow) Then
1530                    Flag = "X"
1540                    s = s & br.Low & " - " & br.High & vbTab
1550                    s = s & "X"
1560                ElseIf br.Code = SysOptBioCodeForGlucose(0) Or _
                           br.Code = SysOptBioCodeForChol(0) Or _
                           br.Code = SysOptBioCodeForTrig(0) Or _
                           br.Code = SysOptBioCodeForGlucoseP(0) Or _
                           br.Code = SysOptBioCodeForCholP(0) Or _
                           br.Code = SysOptBioCodeForTrigP(0) Then
1570                    If Fasting Then
1580                        If br.Code = SysOptBioCodeForGlucose(0) Or br.Code = SysOptBioCodeForGlucoseP(0) Then
1590                            sql = "SELECT * from fastings WHERE testname = '" & "GLU" & "'"
1600                        ElseIf br.Code = SysOptBioCodeForChol(0) Or br.Code = SysOptBioCodeForCholP(0) Then
1610                            sql = "SELECT * from fastings WHERE testname = '" & "CHO" & "'"
1620                        ElseIf br.Code = SysOptBioCodeForTrig(0) Or br.Code = SysOptBioCodeForTrigP(0) Then
1630                            sql = "SELECT * from fastings WHERE testname = '" & "TRI" & "'"
1640                        End If
1650                        Set tb = New Recordset
1660                        RecOpenServer 0, tb, sql
1670                        If Not tb.EOF Then
1680                            s = s & tb!FastingText & vbTab
1690                            If Value > tb!FastingHigh Then
1700                                Flag = "H"
1710                                s = s & "H"
1720                            ElseIf Value < tb!FastingLow Then
1730                                Flag = "L"
1740                                s = s & "L"
1750                            End If
1760                        Else
1770                        End If
1780                    Else
1790                        s = s & br.Low & " - " & br.High & vbTab
1800                        If Value < Val(br.FlagLow) Then
1810                            Flag = "FL"
1820                            T = "FL"
1830                        ElseIf Value > Val(br.FlagHigh) Then
1840                            Flag = "FH"
1850                            T = "FH"
1860                        End If
1870                        If Value < Val(br.Low) Then
1880                            Flag = "L"
1890                            T = "L"
1900                        ElseIf Value > Val(br.High) Then
1910                            Flag = "H"
1920                            T = "H"
1930                        End If
1940                    End If
1950                Else
1960                    If (Val(br.Low) = 0 And Val(br.High) = 0) Or (Val(br.Low) = 0 And Val(br.High) = 999) Or (Val(br.Low) = 0 And Val(br.High) = 9999) Then
1970                        s = s & vbTab
1980                    Else
1990                        s = s & br.Low & " - " & br.High & vbTab
2000                        If Value < Val(br.FlagLow) Then
2010                            Flag = "FL"
2020                            T = "FL"
2030                        ElseIf Value > Val(br.FlagHigh) Then
2040                            Flag = "FH"
2050                            T = "FH"
2060                        End If
2070                        If Value < Val(br.Low) Then
2080                            Flag = "L"
2090                            T = "L"
2100                        ElseIf Value > Val(br.High) Then
2110                            Flag = "H"
2120                            T = "H"
2130                        End If
2140                    End If
2150                End If
2160            ElseIf InStr(1, br.Result, ">") > 0 Then
                    s = s & br.Low & " - " & br.High & vbTab
                    FormatedResult = Replace(br.Result, ">", "")
                    If Val(FormatedResult) >= Val(br.High) Then
                        Flag = "H"
                        T = "H"
                    End If
                ElseIf InStr(1, br.Result, "<") > 0 Then
                    s = s & br.Low & " - " & br.High & vbTab
                    FormatedResult = Replace(br.Result, "<", "")
                    If Val(FormatedResult) <= Val(br.Low) Then
                        Flag = "L"
                        T = "L"
                    End If
                Else
2170                If (Val(br.Low) = 0 And Val(br.High) = 0) Or (Val(br.Low) = 0 And Val(br.High) = 999) Or (Val(br.Low) = 0 And Val(br.High) = 9999) Then
2180                    s = s & vbTab
2190                Else
2200                    s = s & br.Low & " - " & br.High & vbTab
2210                End If
2220            End If
2230        End If
2240        e = ""
2250        e = Trim(br.Flags & "")
2260        s = s & T & vbTab & _
                IIf(e <> "", e, "") & vbTab & _
                IIf(br.Valid, "V", " ") & _
                IIf(br.Printed, "P", " ") & vbTab
2270        If br.Valid = True Then
2280            Frame2.Enabled = False
2290            lRandom.Enabled = False
2300            txtBioComment.Locked = True
2310        End If
2320        s = s & br.Pc & vbTab
2330        Select Case Trim(br.Analyser)
            Case "4": s = s & "Immuno"
2340        Case "A": s = s & "Bio (A)"
2350        Case "B": s = s & "Bio (B)"
2360        Case "P1": s = s & SysOptBioN1(0)
2370        Case "P2": s = s & SysOptBioN2(0)
2380        Case Else: s = s & "General"
2390        End Select
2400        s = s & vbTab & br.Comment
2410        gBio.AddItem s
2420        If br.Printable = False Then
2430            gBio.Row = gBio.Rows - 1
2440            gBio.Col = 10
2450            gBio.CellBackColor = vbRed
2460        End If

2470        If Flag <> "" Then
2480            gBio.Row = gBio.Rows - 1
2490            gBio.Col = 1
2500            Select Case Flag
                Case "H":
2510                For n = 0 To 9
2520                    gBio.Col = n
2530                    gBio.CellBackColor = SysOptHighBack(0)
2540                    gBio.CellForeColor = SysOptHighFore(0)
2550                Next
2560            Case "L":
2570                For n = 0 To 9
2580                    gBio.Col = n
2590                    gBio.CellBackColor = SysOptLowBack(0)
2600                    gBio.CellForeColor = SysOptLowFore(0)
2610                Next
2620            Case "X":
2630                For n = 0 To 9
2640                    gBio.Col = n
2650                    gBio.CellBackColor = SysOptPlasBack(0)
2660                    gBio.CellForeColor = SysOptPlasFore(0)
2670                Next
2680            End Select
2690        End If
2700    Next
2710 End If

2720 If Trim(txtChart) <> "" Then
2730    DoDeltaCheckBio
2740 End If

2750 FixG gBio

2760 With gBio
2770    bValidateBio.Caption = "VALID"
2780    lblUrgent.Visible = False
2790    For n = 1 To .Rows - 1
2800        If .TextMatrix(n, 3) = "X" Then
2810            .Row = n
2820            .Col = 1
2830            .CellForeColor = vbWhite
2840            .CellBackColor = vbBlack
2850        End If
2860        If InStr(.TextMatrix(n, 6), "V") = "0" Then
2870            bValidateBio.Caption = "&Validate"
2880            lblUrgent.Visible = UrgentTest
2890        End If
2900    Next
2910 End With

2920 LoadOutstandingBio

2930 sql = "SELECT COUNT(*) Tot FROM BioRepeats WHERE " & _
           "SampleID = '" & txtSampleID & "'"
2940 Set tb = New Recordset
2950 RecOpenServer 0, tb, sql
2960 bViewBioRepeat.Visible = False
2970 If tb!Tot > 0 Then
2980    bViewBioRepeat.Visible = True
2990 End If

3000 sql = "SELECT * from Masks WHERE " & _
           "SampleID = '" & txtSampleID & "'"
3010 Set tb = New Recordset
3020 RecOpenServer 0, tb, sql
3030 If Not tb.EOF Then
3040    oH = IIf(tb!h, 1, 0)
3050    oS = IIf(tb!s, 1, 0)
3060    oL = IIf(tb!l, 1, 0)
3070    oO = IIf(tb!o, 1, 0)
3080    oG = IIf(tb!g, 1, 0)
3090    oJ = IIf(tb!J, 1, 0)
3100 End If

3110 SetPrintInhibit "Bio"
3120 CheckAuditTrail
3130 EnableBarCodePrinting
3140 CheckLabLinkStatus
3150 bFAX.Enabled = (UCase$(bValidateBio.Caption) = "VALID")

3160 LoadComments

3170 If txtName <> "" Then txtName.SetFocus

3180 Exit Sub

LoadBiochemistry_Error:

    Dim strES As String
    Dim intEL As Integer

3190 intEL = Erl
3200 strES = Err.Description
3210 LogError "frmEditAll", "LoadBiochemistry", intEL, strES, sql

End Sub


Public Sub LoadBloodGas()

    Dim Deltasn As Recordset
    Dim Deltatb As Recordset
    Dim tb As New Recordset
    Dim sql As String
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
    Dim sn As New Recordset

10  On Error GoTo LoadBloodGas_Error

20  If txtSampleID = "" Then Exit Sub

30  lblBgaDate = ""
40  lBgaDelta = ""

50  ClearFGrid gBga

60  bViewBgaRepeat.Visible = False

70  ssTabAll.TabCaption(5) = "Blood Gas"

    'get date & run number of previous record
80  PreviousBga = False
90  HistBga = False

100 If txtName <> "" And txtDoB <> "" Then
110     sql = CreateHist("bga")
120     Set sn = New Recordset
130     RecOpenServer 0, sn, sql
140     If Not sn.EOF Then
150         HistBga = True
160     End If

170     sql = CreateSql("Bga")
180     Set Deltatb = New Recordset
190     RecOpenServer 0, Deltatb, sql
200     If Not Deltatb.EOF Then
210         PreviousDate = Deltatb!Rundate & ""
220         PreviousRec = Deltatb!SampleID & ""
230         PreviousBga = True
240     End If
250 End If

260 Set BRres = BRs.Load("Bga", txtSampleID, "Results", gDONTCARE, gDONTCARE, 0, cCat(0), dtRunDate)

270 With gBga
280     .Rows = 2
290     .AddItem ""
300     .RemoveItem 1
310 End With

320 If Not BRres Is Nothing Then
330     ssTabAll.TabCaption(5) = ">>Blood Gas<<"
340     For Each br In BRres
350         lblBgaDate = Format(GetLatestRunDateTime("Bga", br.SampleID, br!RunTime), "dd/MM/yyyy hh:mm:ss") 'Format(br.RunTime, "dd/MMM/yyyy hh:mm")
360         Flag = ""
370         SampleType = br.SampleType
380         If Len(SampleType) = 0 Then SampleType = "S"
390         s = br.ShortName & vbTab
400         lBDate = Format(GetLatestRunDateTime("Bio", br.SampleID, br.RunTime), "dd/MM/yyyy hh:mm:ss")
410         If IsNumeric(br.Result) Then
420             Value = Val(br.Result)
430             Select Case br.Printformat
                Case 0: valu = Format$(Value, "0")
440             Case 1: valu = Format$(Value, "0.0")
450             Case 2: valu = Format$(Value, "0.00")
460             Case 3: valu = Format$(Value, "0.000")
470             Case Else: valu = Format$(Value, "0.000")
480             End Select
490         Else
500             valu = br.Result
510         End If
520         s = s & valu & vbTab
530         If ListText("UN", br.Units) <> "" Then
540             s = s & ListText("UN", br.Units)
550         Else
560             s = s & br.Units
570         End If
580         s = s & vbTab
590         If txtSex = "" Then    'QMS Ref No. #817982
600             s = s & vbTab
610         Else
620             s = s & br.Low & " - " & br.High & vbTab
630             T = ""
640             If IsNumeric(br.Result) Then
650                 If Value > Val(br.PlausibleHigh) Then
660                     Flag = "X"
670                     s = s & "X"
680                 ElseIf Value < Val(br.PlausibleLow) Then
690                     Flag = "X"
700                     s = s & "X"
710                 Else
720                     If Value < Val(br.FlagLow) Then
730                         Flag = "FL"
740                         T = "FL"
750                     ElseIf Value > Val(br.FlagHigh) Then
760                         Flag = "FH"
770                         T = "FH"
780                     End If
790                     If Value < Val(br.Low) Then
800                         Flag = "L"
810                         T = "L"
820                     ElseIf Value > Val(br.High) Then
830                         Flag = "H"
840                         T = "H"
850                     End If
860                 End If
870             End If
880         End If
890         e = ""
900         e = Trim(br.Flags & "")
910         s = s & T
920         s = s & vbTab & _
                IIf(br.Valid, "V", " ") & vbTab & _
                IIf(br.Printed, "P", " ") & vbTab
930         gBga.AddItem s

940         If Flag <> "" Then
950             gBga.Row = gBga.Rows - 1
960             gBga.Col = 1
970             Select Case Flag
                Case "H":
980                 For n = 0 To 6
990                     gBga.Col = n
1000                    gBga.CellBackColor = SysOptHighBack(0)
1010                    gBga.CellForeColor = SysOptHighFore(0)
1020                Next
1030            Case "L":
1040                For n = 0 To 6
1050                    gBga.Col = n
1060                    gBga.CellBackColor = SysOptLowBack(0)
1070                    gBga.CellForeColor = SysOptLowFore(0)
1080                Next
1090            Case "X":
1100                For n = 0 To 7
1110                    gBga.Col = n
1120                    gBga.CellBackColor = SysOptPlasBack(0)
1130                    gBga.CellForeColor = SysOptPlasFore(0)
1140                Next
1150            End Select
1160        End If

1170        If br.DoDelta And PreviousBga Then
1180            sql = "SELECT * from bgaresults WHERE " & _
                      "sampleid = '" & PreviousRec & "' " & _
                      "and code = '" & br.Code & "'"
1190            Set Deltasn = New Recordset
1200            RecOpenClient 0, Deltasn, sql
1210            If Not Deltasn.EOF Then
1220                OldValue = Val(Deltasn!Result)
1230                If OldValue <> 0 Then
1240                    DeltaLimit = br.DeltaLimit
1250                    If Abs(OldValue - Value) > DeltaLimit Then
1260                        Res = Format$(PreviousDate, "dd/mm/yyyy") & " (" & PreviousRec & ") " & _
                                  br.ShortName & " " & _
                                  OldValue & vbCr
1270                        lBgaDelta = lBgaDelta & Res
1280                    End If
1290                End If
1300            End If
1310        End If
1320        OldValue = 0
1330    Next
1340 End If

1350 FixG gBga

1360 With gBga
1370    cmdValBG.Caption = "VALID"
1380    lblUrgent.Visible = False
1390    For n = 1 To .Rows - 1
1400        If .TextMatrix(n, 3) = "X" Then
1410            .Row = n
1420            .Col = 1
1430            .CellForeColor = vbWhite
1440            .CellBackColor = vbBlack
1450        End If
1460        If InStr(.TextMatrix(n, 5), "V") = "0" Then
1470            cmdValBG.Caption = "&Validate"
1480            lblUrgent.Visible = UrgentTest
1490        End If
1500    Next
1510 End With

1520 sql = "SELECT * from BgaRepeats WHERE " & _
           "SampleID = '" & txtSampleID & "'"
1530 Set tb = New Recordset
1540 RecOpenServer 0, tb, sql
1550 bViewBgaRepeat.Visible = False
1560 If Not tb.EOF Then
1570    bViewBgaRepeat.Visible = True
1580 End If
1590 bFAX.Enabled = (UCase$(cmdValBG.Caption) = "VALID")

1600 LoadComments

1610 Exit Sub

LoadBloodGas_Error:

    Dim strES As String
    Dim intEL As Integer

1620 intEL = Erl
1630 strES = Err.Description
1640 LogError "frmEditAll", "LoadBloodGas", intEL, strES, sql

End Sub

Public Sub LoadCoagulation()

          Dim CRs As New CoagResults
          Dim cRR As New CoagResults
          Dim CR As CoagResult
          Dim s As String
          Dim n As Long
          Dim x As Long
          Dim sql As String
          Dim tb As New Recordset
          Dim g As String
          Dim sex As String
          Dim sn As New Recordset
          Dim Deltasn As Recordset
          Dim DaysOld As String
          Dim Dob As String
          Dim Value As Single
          Dim OldValue As Single
          Dim DeltaLimit As Single
          Dim Res As String
          Dim resultFlag As Boolean

10    On Error GoTo LoadCoagulation_Error

20    If txtSampleID = "" Then Exit Sub

30    s = CheckAutoComments(txtSampleID, "Coagulation")


40    ClearFGrid grdCoag
50    lIDelta(2) = ""

60    cmdValidateCoag.Caption = "&Validate"
70    txtCoagComment.Locked = False

80    HistCoag = False


90    If txtName <> "" And txtDoB <> "" And IsDate(txtDoB) Then
100     sql = CreateHist("Coag")
110     Set sn = New Recordset
120     RecOpenServer 0, sn, sql
130     If Not sn.EOF Then
140         HistCoag = True
150     End If
        '
        '    sql = CreateSql("Coag")
        '    Set Deltasn = New Recordset
        '    RecOpenServer 0, Deltasn, sql
        '    If Not Deltasn.EOF Then
        '        PreviousDate = Deltasn!Rundate & ""
        '        PreviousRec = Deltasn!SampleID & ""
        '        PreviousEnd = True
        '    End If
160   End If

170   Dob = txtDoB

180   If Dob <> "" And Len(txtDoB) > 9 Then DaysOld = Abs(DateDiff("d", dtRunDate, txtDoB)) Else DaysOld = 12783

190   If DaysOld = 0 Then DaysOld = 1

200   Set CRs = CRs.Load(txtSampleID, gDONTCARE, gDONTCARE, Trim(SysOptExp(0)), 0)
210   Set cRR = cRR.LoadRepeats(txtSampleID, gDONTCARE, gDONTCARE, Trim(SysOptExp(0)))

220   ClearCoagulation

230   ssTabAll.TabCaption(3) = "Coagulation"

240   sql = "SELECT * FROM Demographics WHERE " & _
          "SampleID = '" & txtSampleID & "'"

250   Set tb = New Recordset
260   RecOpenServer 0, tb, sql
270   If Not tb.EOF Then
280     sex = tb!sex & ""
290   End If

300   sql = "SELECT * from coagresults WHERE sampleid = '" & txtSampleID & "'"
310   Set tb = New Recordset
320   RecOpenServer 0, tb, sql
330   If Not tb.EOF Then
340     ssTabAll.TabCaption(3) = ">>Coagulation<<"
350     If tb!Valid = True Then cmdValidateCoag.Caption = "VALID"
360   End If

370   For Each CR In CRs
380     lCDate = CR.Rundate & " " & Format(CR.RunTime, "hh:mm")
390     sql = "SELECT * from coagtestdefinitions WHERE " & _
              "(code = '" & Trim(CR.Code) & "' OR TestName = '" & Trim$(CR.Code) & "') " & _
              "and agefromdays <= '" & DaysOld & "' and agetodays >= '" & DaysOld & "'"
        'Zyam 14-3-24
        If InStr(1, CR.Result, ">") Then
                resultFlag = True
        End If
        'Zyam 14-3-24
400     Set tb = New Recordset
410     RecOpenServer 0, tb, sql
420     If Not tb.EOF Then
430         If HospName(0) = "Mullingar" Or tb!InUse = True Then
                'If Trim(CR.Units) = "INR" Then
                    's = "INR" & vbTab
                'Else
                    's = CoagNameFor(CR.Code) & vbTab
                'End If
440             s = CoagNameFor(CR.Code) & vbTab
450             If UserMemberOf = "The World" And Not CR.Valid Then
460                 s = s & vbTab
470             Else
480                 Select Case CoagPrintFormat(Trim(CR.Code) & "")
                    Case 0: g = Format$(CR.Result, "0")
490                 Case 1: g = Format$(CR.Result, "0.0")
500                 Case 2: g = Format$(CR.Result, "0.00")
510                 End Select
520                 If g = "0" Or g = "0.0" Or g = "0.00" Then
530                     g = "Check"
540                 End If
550                 s = s & g & vbTab & _
                        IIf(CR.Code = "1", "R", UnitConv(CR.Units)) & vbTab
560                 If Trim(UCase(CR.Code)) = "1" Or txtSex = "" Then   'QMS Ref No. #817982
570                     s = s & vbTab
580                 Else
                        'Removed Ranges for some tests
                        'Zyam
                        If CR.Code = "1" Or CR.Code = "13" Or CR.Code = "14" Or CR.Code = "27" Or CR.Code = "94" Or CR.Code = "95" Then
                            s = s & vbTab
                        Else
                            If sex = "M" Then
600                             s = s & Trim(tb!MaleLow) & " - " & Trim(tb!MaleHigh) & vbTab
610                         ElseIf sex = "F" Then
620                             s = s & Trim(tb!FemaleLow) & " - " & Trim(tb!FemaleHigh) & vbTab
630                         Else
640                             s = s & Trim(tb!FemaleLow) & " - " & Trim(tb!MaleHigh) & vbTab
650                         End If
                        End If
                        'Zyam
590
660                 End If
670                 If Trim(UCase(CR.Code)) = "1" Or txtSex = "" Then   'QMS Ref No. #817982
680                     s = s & vbTab
690                 Else
                         'Zyam removed flags for some tests
                         If CR.Code = "1" Or CR.Code = "13" Or CR.Code = "14" Or CR.Code = "27" Or CR.Code = "94" Or CR.Code = "95" Then
                            s = s & vbTab
                         Else
                            s = s & InterpCoag(sex, CR.Code, CR.Result, DaysOld, resultFlag) & vbTab
                         End If
                         'Zyam
700
710                 End If
720                 s = s & IIf(CR.Valid, "V", "") & vbTab & _
                        IIf(CR.Printed, "P", "")
730                 s = s & vbTab & CR.Analyser
740             End If
750             grdCoag.AddItem s

760             If txtName <> "" And txtDoB <> "" And IsDate(txtDoB) Then
770                 Set Deltasn = DoDeltaCheck("Coag", CR.Code)
780                 If CR.DoDelta And (Not Deltasn.EOF) Then
790                     If (dtSampleDate - CDate(Format(Deltasn!SampleDate, "dd/mm/yyyy"))) <= CR.CheckTime Then
800                         OldValue = Val(Deltasn!Result)
810                         If OldValue <> 0 Then
820                             DeltaLimit = CR.DeltaLimit
830                             If Abs(OldValue - Value) > DeltaLimit Then
840                                 Res = Format$(Deltasn!SampleDate, "dd/mm/yyyy") & " (" & Deltasn!SampleID & ") " & _
                                          CoagNameFor(CR.Code) & " " & _
                                          OldValue & vbCr
850                                 lIDelta(2) = lIDelta(2) & Res
860                             End If
870                         End If
880                     End If
890                 End If
900             End If
910             OldValue = 0
920         End If
930     End If
940   Next

950   FixG grdCoag

960   If grdCoag.Rows > 2 Then
970     ssTabAll.TabCaption(3) = ">>Coagulation<<"
980   End If

990   If txtCoagComment <> "" Then
1000    ssTabAll.TabCaption(3) = ">>Coagulation<<"
1010  End If

1020  With grdCoag
1030    If grdCoag.TextMatrix(1, 0) <> "" Then
1040        cmdValidateCoag.Caption = "VALID"
1050        txtCoagComment.Locked = True
1060        lblUrgent.Visible = False
1070        For n = 1 To .Rows - 1
1080            If .TextMatrix(n, 0) <> "" Then
1090                If InStr(.TextMatrix(n, 5), "V") = "0" Then
1100                    cmdValidateCoag.Caption = "&Validate"
1110                    lblUrgent.Visible = UrgentTest
1120                    txtCoagComment.Locked = False
1130                End If
1140            End If
1150        Next
1160    End If
1170  End With

1180  For n = 1 To grdCoag.Rows - 1
1190    Select Case Left(grdCoag.TextMatrix(n, 4), 1)
        Case "X":
1200        grdCoag.Row = n
1210        For x = 0 To grdCoag.Cols - 1
1220            grdCoag.Col = x
1230            grdCoag.CellBackColor = SysOptPlasBack(0)
1240            grdCoag.CellForeColor = SysOptPlasFore(0)
1250        Next
1260    Case "H":
1270        grdCoag.Row = n
1280        For x = 0 To grdCoag.Cols - 1
1290            grdCoag.Col = x
1300            grdCoag.CellBackColor = SysOptHighBack(0)
1310            grdCoag.CellForeColor = SysOptHighFore(0)
1320        Next
1330    Case "L":
1340        grdCoag.Row = n
1350        For x = 0 To grdCoag.Cols - 1
1360            grdCoag.Col = x
1370            grdCoag.CellBackColor = SysOptLowBack(0)
1380            grdCoag.CellForeColor = SysOptLowFore(0)
1390        Next
1400    Case Else
1410        grdCoag.Row = n
1420        For x = 0 To grdCoag.Cols - 1
1430            grdCoag.Col = x
1440            grdCoag.CellBackColor = vbWhite
1450            grdCoag.CellForeColor = vbBlack
1460        Next
1470    End Select
1480    If grdCoag.TextMatrix(n, 1) = "Check" Then
1490        grdCoag.Row = n
1500        For x = 0 To grdCoag.Cols - 1
1510            grdCoag.Col = x
1520            grdCoag.CellBackColor = vbBlue
1530            grdCoag.CellForeColor = vbYellow
1540        Next
1550    End If
1560  Next

1570  bViewCoagRepeat.Visible = cRR.Count <> 0

1580  LoadOutstandingrdCoag

1590  If SysOptDontShowPrevCoag(0) = True Then
1600    grdPrev.Visible = False
1610    lblPrevCoag.Visible = False
1620  Else
1630    LoadPreviousCoag
1640  End If

1650  SetPrintInhibit "Coa"
1660  CheckAuditTrail
1670  EnableBarCodePrinting
1680  CheckLabLinkStatus
1690  bFAX.Enabled = (UCase$(cmdValidateCoag.Caption) = "VALID")
1700  Exit Sub

          'If InStr(txtCoagComment, s) = 0 Then
          '    txtCoagComment = s & vbCrLf & txtCoagComment
          'End If
          'SaveComments

1710  LoadComments
      resultFlag = False
      grdCoag.ColWidth(1) = 3000

1720  If txtName <> "" Then txtName.SetFocus

LoadCoagulation_Error:

          Dim strES As String
          Dim intEL As Integer

1730  intEL = Erl
1740  strES = Err.Description
1750  LogError "frmEditAll", "LoadCoagulation", intEL, strES, sql

End Sub
Private Sub CheckPatientNotePad(SampleID As String)

    Dim tb As New Recordset
    Dim sql As String

   On Error GoTo CheckPatientNotePad_Error

    sql = "SELECT * from PatientNotePad WHERE " & _
          "SampleID = '" & txtSampleID & "'"
    Set tb = New Recordset
    RecOpenServer 0, tb, sql
    If Not tb.EOF Then
        cmdPatientNotePad(0).BackColor = &HFF00&
    Else
        cmdPatientNotePad(0).BackColor = &H8000000F
    End If

   On Error GoTo 0
   Exit Sub

CheckPatientNotePad_Error:

    Dim strES As String
    Dim intEL As Integer

 intEL = Erl
 strES = Err.Description
 LogError "frmEditAll", "CheckPatientNotePad", intEL, strES, sql

End Sub

Private Sub LoadComments()

    Dim Ob As Observation
    Dim Obs As Observations

10  On Error GoTo LoadComments_Error

20  txtBioComment = ""
30  txtHaemComment = ""
40  txtDemographicComment = ""
50  lblDemographicComment = ""
60  txtCoagComment = ""
70  txtImmComment(0) = ""
80  txtImmComment(1) = ""
90  txtBGaComment = ""

100 If Trim$(txtSampleID) = "" Then Exit Sub

110 Set Obs = New Observations
120 Set Obs = Obs.Load(txtSampleID, "Biochemistry", "Demographic", "Haematology", "Coagulation", _
                       "Immunology", "Endocrinology", "BloodGas")
130 If Not Obs Is Nothing Then
140     For Each Ob In Obs
150         Select Case UCase$(Ob.Discipline)
            Case "BIOCHEMISTRY": txtBioComment = Split_Comm(Ob.Comment)
160         Case "HAEMATOLOGY": txtHaemComment = Split_Comm(Ob.Comment)
170         Case "DEMOGRAPHIC": txtDemographicComment = Split_Comm(Ob.Comment)
180             lblDemographicComment = txtDemographicComment
190         Case "COAGULATION": txtCoagComment = Split_Comm(Ob.Comment)
200         Case "IMMUNOLOGY": txtImmComment(1) = Split_Comm(Ob.Comment)
210         Case "ENDOCRINOLOGY": txtImmComment(0) = Split_Comm(Ob.Comment)
220         Case "BLOODGAS": txtBGaComment = Split_Comm(Ob.Comment)
230         End Select
240     Next
250 End If

260 Exit Sub

LoadComments_Error:

    Dim strES As String
    Dim intEL As Integer

270 intEL = Erl
280 strES = Err.Description
290 LogError "frmEditAll", "LoadComments", intEL, strES

End Sub

Private Sub LoadDemo(ByVal IDNumber As String)

    Dim tb As New Recordset
    Dim sql As String
    Dim IDType As String
    Dim n As Long

10  On Error GoTo LoadDemo_Error

20  IDType = CheckDemographics(IDNumber)
30  If IDType = "" Then
        'clearpatient
40      Exit Sub
50  End If

    'Rem Code Change 16/01/2006
60  sql = "SELECT * from patientifs WHERE " & _
          IDType & " = '" & AddTicks(IDNumber) & "' "

70  Set tb = New Recordset
80  RecOpenServer 0, tb, sql
90  If tb.EOF = True Then
        '   clearpatient
100 Else
110     If Trim(tb!Chart & "") = "" Then txtChart = tb!Mrn & "" Else txtChart = tb!Chart & ""
120     txtAandE = tb!AandE & ""
130     n = InStr(tb!PatName & "", "''")
140     If n <> 0 Then
150         tb!PatName = Left$(tb!PatName, n) & Mid$(tb!PatName, n + 2)
160         tb.Update
170     End If
180     txtName = initial2upper(tb!PatName & "")
190     If Not IsNull(tb!Dob) Then
200         lDoB = Format(tb!Dob, "DD/MM/YYYY")
210         txtDoB = Format(tb!Dob, "DD/MM/YYYY")
220     Else
230         lDoB = ""
240         txtDoB = ""
250     End If
260     lAge = CalcAge(tb!Dob & "", dtSampleDate)
270     txtAge = lAge
280     Select Case tb!sex & ""
        Case "M": lSex = "Male"
290     Case "F": lSex = "Female"
300     Case Else: lSex = ""
310     End Select
320     txtSex = lSex
330     n = InStr(tb!Addr0 & "", "''")
340     If n <> 0 Then
350         tb!Addr0 = Left$(tb!Addr0, n) & Mid$(tb!Addr0, n + 2)
360         tb.Update
370     End If

380     taddress(0) = initial2upper(Trim(tb!Addr0 & ""))
390     taddress(1) = initial2upper(Trim(tb!Addr1 & ""))
400     cmbWard.Text = initial2upper(tb!Ward & "")
410     cmbClinician.Text = initial2upper(tb!Clinician & "")
420 End If
430 tb.Close

440 Exit Sub

LoadDemo_Error:

    Dim strES As String
    Dim intEL As Integer

450 intEL = Erl
460 strES = Err.Description
470 LogError "frmEditAll", "LoadDemo", intEL, strES, sql

End Sub

Public Sub LoadDemographics()

          Dim sql As String
          Dim tb As New Recordset
          Dim SampleDate As String
          Dim RooH As Boolean

10    On Error GoTo LoadDemographics_Error

20    UrgentTest = False
30    RooH = IsRoutine()
40    cRooH(0) = RooH
50    cRooH(1) = Not RooH
60    bViewBB.Enabled = False
70    txtAge = ""
80    lAge = ""
90    If Trim$(txtSampleID) = "" Then Exit Sub

100   lRandom = "Random Sample"
110   lImmRan(0) = "Random Sample"
120   lImmRan(1) = "Random Sample"

130   sql = "SELECT * FROM Demographics WHERE " & _
          "SampleID = '" & txtSampleID & "'"

140   Set tb = New Recordset
150   RecOpenServer 0, tb, sql
160   If Not tb.EOF Then
170     If Trim$(tb!Hospital & "") <> "" Then
180         lblChartNumber = Trim$(UCase(tb!Hospital)) & " Chart #"
190         If UCase(tb!Hospital) = HospName(0) Then
200             lblChartNumber.BackColor = &H8000000F
210             lblChartNumber.ForeColor = vbBlack
220         Else
230             lblChartNumber.BackColor = vbRed
240             lblChartNumber.ForeColor = vbYellow
250         End If
260     Else
270         lblChartNumber.Caption = HospName(0) & " Chart #"
280         lblChartNumber.BackColor = &H8000000F
290         lblChartNumber.ForeColor = vbBlack
300     End If
310     If IsDate(tb!SampleDate) Then
320         dtSampleDate = Format$(tb!SampleDate, "dd/mm/yyyy")
330         lblSampleDate = dtSampleDate
340     Else
350         dtSampleDate = Format$(Now, "dd/mm/yyyy")
360         lblSampleDate = dtSampleDate
370     End If
380     If IsDate(tb!Rundate) Then
390         dtRunDate = Format$(tb!Rundate, "dd/mm/yyyy")
400     Else
410         dtRunDate = Format$(Now, "dd/mm/yyyy")
420     End If
430     StatusBar1.Panels(4).Text = dtRunDate
440     mNewRecord = False
450     If Trim$(tb!RooH & "") <> "" Then cRooH(0) = tb!RooH
460     If Trim$(tb!RooH & "") <> "" Then cRooH(1) = Not tb!RooH
470     txtChart = Trim(tb!Chart & "")
480     txtName = Trim(initial2upper(tb!PatName & ""))
490     txtAandE = Trim(tb!AandE & "")
500     taddress(0) = tb!Addr0 & ""
510     taddress(1) = tb!Addr1 & ""
520     Select Case Left$(Trim$(UCase$(tb!sex & "")), 1)
        Case "M": txtSex = "Male"
530     Case "F": txtSex = "Female"
540     Case Else: txtSex = ""
550     End Select
560     If Trim(tb!Dob & "") <> "" Then txtDoB = Format$(tb!Dob, "dd/mm/yyyy") Else txtDoB = ""
570     If tb!Age & "" <> "" Then
580         txtAge = Trim(tb!Age)
590     Else
600         If Trim(tb!Dob & "") <> "" Then txtAge = CalcOldAge(tb!Dob, dtRunDate)
610     End If
620     lAge = txtAge & ""
630     lDoB = txtDoB
640     If Trim(tb!Hospital) & "" <> "" Then cmbHospital = tb!Hospital Else cmbHospital = HospName(0)
650     cmbClinician = Trim(tb!Clinician & "")
660     cmbGP = Trim(tb!GP & "")
670     cmbWard = Trim(tb!Ward & "")
680     cClDetails = Trim(tb!ClDetails & "")
690     If Trim$(tb!Category & "") <> "" Then
700         cCat(0) = Trim(tb!Category & "")
710         cCat(1) = Trim(tb!Category & "")
720     Else
730         cCat(0) = "Default"
740         cCat(1) = "Default"
750     End If
760     If IsDate(tb!SampleDate) Then
770         dtSampleDate = Format$(tb!SampleDate, "dd/mm/yyyy")
780         If Format$(tb!SampleDate, "hh:mm") <> "00:00" Then
790             tSampleTime = Format$(tb!SampleDate, "hh:mm")
800         Else
810             tSampleTime.Mask = ""
820             tSampleTime.Text = ""
830             tSampleTime.Mask = "##:##"
840         End If
850     ElseIf IsDate(tb!RecDate) Then
860         dtSampleDate = Format$(tb!RecDate, "dd/mm/yyyy")
870         tSampleTime.Mask = ""
880         tSampleTime.Text = ""
890         tSampleTime.Mask = "##:##"
900     ElseIf IsDate(tb!Rundate & "") Then
910         dtSampleDate = Format$(tb!Rundate, "dd/mm/yyyy")
920         tSampleTime.Mask = ""
930         tSampleTime.Text = ""
940         tSampleTime.Mask = "##:##"
950     End If
960     If SysOptDemoVal(0) = True Then
970         If tb!Valid <> 0 Then
980             cmdDemoVal.Caption = "VALID"
990             EnableDemographicEntry False
1000        Else
1010            cmdDemoVal.Caption = "&Validate"
1020            EnableDemographicEntry True
1030        End If
1040    End If
1050    If IsDate(tb!RecDate & "") Then
1060        dtRecDate = Format$(tb!RecDate, "dd/mm/yyyy")
1070        If Format$(tb!RecDate, "hh:mm") <> "00:00" Then
1080            tRecTime = Format$(tb!RecDate, "hh:mm")
1090        Else
1100            tRecTime.Mask = ""
1110            tRecTime.Text = ""
1120            tRecTime.Mask = "##:##"
1130        End If
1140    Else
1150        If Trim(tb!Rundate & "") <> "" Then dtRecDate = Format$(tb!Rundate, "dd/mm/yyyy")
1160        tRecTime.Mask = ""
1170        tRecTime.Text = ""
1180        tRecTime.Mask = "##:##"
1190    End If
1200    If Trim$(tb!Fasting & "") <> "" Then
1210        If tb!Fasting Then
1220            lRandom = "Fasting Sample"
1230            lImmRan(0) = "Fasting Sample"
1240            lImmRan(1) = "Fasting Sample"
1250        End If
1260    End If
1270    If SysOptUrgent(0) Then
1280        If tb!Urgent = 1 Then
1290            lblUrgent.Visible = True
1300            chkUrgent.Value = 1
1310            UrgentTest = True
1320        Else
1330            chkUrgent.Value = 0
1340            UrgentTest = False
1350        End If
1360    End If
1370  End If

1380  cmdSaveDemographics.Enabled = False
1390  cmdSaveInc.Enabled = False

1400  If SysOptViewTrans(0) = True Then
1410    bViewBB.Visible = True
1420    If CnxnBB(0) Is Nothing Then
1430    Else
1440        If Trim$(txtChart) <> "" And Right(CnxnBB(0), 2) <> "=;" Then
1450            sql = "SELECT  * from PatientDetails WHERE " & _
                      "PatNum = '" & txtChart & "'"
1460            Set tb = New Recordset
1470            RecOpenClientBB tb, sql
1480            bViewBB.Enabled = Not tb.EOF
1490        End If
1500    End If
1510  End If

1520  CheckCC
1530  CheckAuditTrail
1540  EnableBarCodePrinting
1550  CheckLabLinkStatus

1560  Exit Sub

LoadDemographics_Error:

          Dim strES As String
          Dim intEL As Integer

1570  intEL = Erl
1580  strES = Err.Description
1590  LogError "frmEditAll", "LoadDemographics", intEL, strES, sql

End Sub

Public Sub LoadEndocrinology()

          Dim Deltasn As Recordset
          Dim tb As New Recordset
          Dim sql As String
          Dim s As String
          Dim Value As Single
          Dim OldValue As Single
          Dim valu As String
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

10    On Error GoTo LoadEndocrinology_Error

20    If txtSampleID = "" Then Exit Sub

30    s = CheckAutoComments(txtSampleID, "Endocrinology")


40    PreviousEnd = False
50    HistEnd = False

60    Fasting = lImmRan(0) = "Fasting Sample"

70    lblEDate = ""
80    lIDelta(0) = ""
90    bViewImmRepeat(0).Visible = False

100   ssTabAll.TabCaption(4) = "Endocrinology"

110   ClearFGrid gImm(0)

          'get date & run number of previous record

120   If txtName <> "" And txtDoB <> "" And IsDate(txtDoB) Then
130     sql = CreateHist("end")
140     Set sn = New Recordset
150     RecOpenServer 0, sn, sql
160     If Not sn.EOF Then
170         HistEnd = True
180     End If
        '
        '    sql = CreateSql("End")
        '    Set Deltatb = New Recordset
        '    RecOpenServer 0, Deltatb, sql
        '    If Not Deltatb.EOF Then
        '        PreviousDate = Deltatb!Rundate & ""
        '        PreviousRec = Deltatb!SampleID & ""
        '        PreviousEnd = True
        '    End If
190   End If

200   If cCat(0) = "" Then Cat = "Default" Else Cat = cCat(0)

210   Set IMres = Ims.Load("End", txtSampleID, "Results", gDONTCARE, gDONTCARE, 0, Cat, dtRunDate)

220   CheckCalcEPSA IMres

230   If Not IMres Is Nothing Then
240     ssTabAll.TabCaption(4) = ">>Endocrinology<<"
250     For Each Im In IMres
260         SampleType = Im.SampleType
270         If Len(SampleType) = 0 Then SampleType = "S"
280         s = Im.ShortName & vbTab
290         lblEDate = Format(GetLatestRunDateTime("End", Im.SampleID, Im.RunTime), "dd/MM/yyyy hh:mm:ss")
300         If UCase(Im.Analyser) = "VIROLOGY" Then
                'if AxSym virology then translate result here
310             gImm(0).Tag = Im.Analyser
320             Im.Result = TranslateEndResultVirology(Im.Code, Im.Result)
                'now result is non numeric. so it won't generate any flags or apply any rules
330         End If
340         If IsNumeric(Im.Result) Then
350             Value = Val(Im.Result)
360             Select Case Im.Printformat
                Case 0: valu = Format$(Value, "0")
370             Case 1: valu = Format$(Value, "0.0")
380             Case 2: valu = Format$(Value, "0.00")
390             Case 3: valu = Format$(Value, "0.000")
400             Case Else: valu = Format$(Value, "0.000")
410             End Select
420         Else
430             valu = Im.Result
440         End If
            '   If UserMemberOf = "The World" And Not BR.Valid Then
            '     s = s & "" & vbTab
            '   Else
450         s = s & valu & vbTab
460         If ListText("UN", Im.Units) <> "" Then
470             s = s & ListText("UN", Im.Units)
480         Else
490             s = s & Im.Units
500         End If
510         s = s & vbTab
520         If txtSex = "" Then   'QMS Ref No. #817982
530             s = s & vbTab
540         Else
550             s = s & Im.Low & " - " & Im.High & vbTab
560             If IsNumeric(Im.Result) Then
570                 If Value > Val(Im.PlausibleHigh) Then
580                     Flag = "X"
590                     s = s & "X"
600                 ElseIf Value < Val(Im.PlausibleLow) Then
610                     Flag = "X"
620                     s = s & "X"
630                 Else
640                     If Value < Val(Im.Low) Then
650                         Flag = "L"
660                         s = s & "L"
670                     ElseIf Value > Val(Im.High) Then
680                         Flag = "H"
690                         s = s & "H"
700                     End If
710                 End If
720             Else
730                 If Left(Im.Result, 1) = "<" Then
740                     Flag = "L"
750                     s = s & "L"
760                 ElseIf Left(Im.Result, 1) = ">" Then
770                     Flag = "H"
780                     s = s & "H"
790                 End If
800             End If
810         End If
820         If Im.Flags = "1" Then e = "C" Else e = ""
830         s = s & vbTab & _
                IIf(e <> "", e, "") & vbTab & _
                IIf(Im.Valid, "V", " ") & _
                IIf(Im.Printed, "P", " ") & vbTab & Trim(Im.Comment)
840         gImm(0).AddItem s
850         If UCase(Im.Analyser) <> "VIROLOGY" Then
860             If Flag <> "" Then
870                 gImm(0).Row = gImm(0).Rows - 1
880                 gImm(0).Col = 1
890                 Select Case Flag
                    Case "H":
900                     For n = 0 To 7
910                         gImm(0).Col = n
920                         gImm(0).CellBackColor = SysOptHighBack(0)
930                         gImm(0).CellForeColor = SysOptHighFore(0)
940                     Next
950                 Case "L":
960                     For n = 0 To 7
970                         gImm(0).Col = n
980                         gImm(0).CellBackColor = SysOptLowBack(0)
990                         gImm(0).CellForeColor = SysOptLowFore(0)
1000                    Next
1010                Case "X":
1020                    For n = 0 To 7
1030                        gImm(0).Col = n
1040                        gImm(0).CellBackColor = SysOptPlasBack(0)
1050                        gImm(0).CellForeColor = SysOptPlasFore(0)
1060                    Next
1070                End Select
1080            End If
1090        End If
1100        Flag = ""
1110        If txtName <> "" And txtDoB <> "" And IsDate(txtDoB) Then
1120            Set Deltasn = DoDeltaCheck("End", Im.Code)
1130            If Im.DoDelta And (Not Deltasn.EOF) Then
1140                If (dtSampleDate - CDate(Format(Deltasn!SampleDate, "dd/mm/yyyy"))) <= Im.CheckTime Then
1150                    OldValue = Val(Deltasn!Result)
1160                    If OldValue <> 0 Then
1170                        DeltaLimit = Im.DeltaLimit
1180                        If Abs(OldValue - Value) > DeltaLimit Then
1190                            Res = Format$(Deltasn!SampleDate, "dd/mm/yyyy") & " (" & Deltasn!SampleID & ") " & _
                                      Im.ShortName & " " & _
                                      OldValue & vbCr
1200                            lIDelta(0) = lIDelta(0) & Res
1210                        End If
1220                    End If
1230                End If
1240            End If
1250        End If
1260        OldValue = 0
1270    Next
1280  End If

1290  FixG gImm(0)

1300  With gImm(0)
1310    bValidateImm(0).Caption = "VALID"
1320    lblUrgent.Visible = False
1330    For n = 1 To .Rows - 1
1340        If .TextMatrix(n, 3) = "X" Then
1350            .Row = n
1360            .Col = 1
1370            .CellForeColor = vbWhite
1380            .CellBackColor = vbBlack
1390        End If
1400        If InStr(.TextMatrix(n, 6), "V") = "0" Then
1410            bValidateImm(0).Caption = "&Validate"
1420            lblUrgent.Visible = UrgentTest
1430        End If
1440    Next
1450  End With

1460  LoadOutstandingEnd

1470  sql = "SELECT * from endRepeats WHERE " & _
           "SampleID = '" & txtSampleID & "'"
1480  Set tb = New Recordset
1490  RecOpenServer 0, tb, sql
1500  bViewImmRepeat(0).Visible = False
1510  If Not tb.EOF Then
1520    bViewImmRepeat(0).Visible = True
1530  End If

1540  sql = "SELECT * from EndMasks WHERE " & _
           "SampleID = '" & txtSampleID & "'"
1550  Set tb = New Recordset
1560  RecOpenServer 0, tb, sql
1570  If Not tb.EOF Then
1580    Ih(0) = IIf(tb!h, 1, 0)
1590    Iis(0) = IIf(tb!s, 1, 0)
1600    Il(0) = IIf(tb!l, 1, 0)
1610    Io(0) = IIf(tb!o, 1, 0)
1620    Ig(0) = IIf(tb!g, 1, 0)
1630    Ij(0) = IIf(tb!J, 1, 0)
1640  End If
1650  SetPrintInhibit "End"
1660  CheckAuditTrail
1670  CheckLabLinkStatus
1680  EnableBarCodePrinting
1690  bFAX.Enabled = (UCase$(bValidateImm(0).Caption) = "VALID")

          'If InStr(txtImmComment(0), s) = 0 And UCase(bValidateImm(0).Caption) = "&VALIDATE" Then
          '    txtImmComment(0) = s & vbCrLf & txtImmComment(0)
          'End If
          'SaveComments

1700  LoadComments

1710  If txtName <> "" Then txtName.SetFocus

1720  Exit Sub

LoadEndocrinology_Error:

          Dim strES As String
          Dim intEL As Integer

1730  intEL = Erl
1740  strES = Err.Description
1750  LogError "frmEditAll", "LoadEndocrinology", intEL, strES, sql

End Sub

Private Function LoadEndSplitList(ByVal Index As Integer) As String

    Dim tb As New Recordset
    Dim sql As String
    Dim strIndex As String
    Dim strReturn As String

10  On Error GoTo LoadEndSplitList_Error

20  strIndex = Index

30  sql = "SELECT distinct Code, PrintPriority, SplitList " & _
          "from EndTestDefinitions " & _
          "WHERE SplitList = " & strIndex & " " & _
          "order by PrintPriority"

40  Set tb = New Recordset
50  RecOpenServer 0, tb, sql

60  strReturn = ""
70  Do While Not tb.EOF
80      strReturn = strReturn & "Code = '" & tb!Code & "' or "
90      tb.MoveNext
100 Loop
110 If strReturn <> "" Then
120     strReturn = Left$(strReturn, Len(strReturn) - 3)
130 End If

140 LoadEndSplitList = strReturn

150 Exit Function

LoadEndSplitList_Error:

    Dim strES As String
    Dim intEL As Integer

160 intEL = Erl
170 strES = Err.Description
180 LogError "frmEditAll", "LoadEndSplitList", intEL, strES, sql

End Function

Private Sub LoadExt()

    Dim sql As String
    Dim tb As New Recordset
    Dim Deltatb As Recordset
    Dim s As String
    Dim TestName As String
    Dim PreviousDate As String
    Dim PreviousRec As Long
    Dim sn As New Recordset
    Dim n As Integer
    Dim i As Integer

10  On Error GoTo LoadExt_Error

20  If txtSampleID = "" Then Exit Sub

30  ClearFGrid grdExt

40  ssTabAll.TabCaption(7) = "Externals"

50  PreviousExt = False
60  HistExt = False

70  sql = CreateHist("Ext")
80  Set sn = New Recordset
90  RecOpenServer 0, sn, sql
100 If Not sn.EOF Then
110     HistExt = True
120 End If

130 sql = CreateSql("Ext")

140 Set Deltatb = New Recordset
150 RecOpenServer 0, Deltatb, sql
160 If Not Deltatb.EOF Then
170     PreviousDate = Deltatb!Rundate & ""
180     PreviousRec = Deltatb!SampleID & ""
190     PreviousExt = True
200 End If

210 sql = "SELECT * FROM Extresults WHERE sampleid = " & txtSampleID & " order by orderlist"
220 Set tb = New Recordset
230 RecOpenServer 0, tb, sql
240 Do While Not tb.EOF
250     ssTabAll.TabCaption(7) = ">>Externals<<"
260     TestName = tb!Analyte & ""
270     s = TestName & vbTab & _
            Trim(tb!Result) & vbTab & _
            tb!NormalRange & vbTab & _
            tb!Units & vbTab & _
            tb!SendTo & vbTab
280     If Not IsNull(tb!SentDate) Then
290         s = s & Format(tb!SentDate, "dd/mmm/yyyy hh:mm:ss")
300     End If
310     s = s & vbTab
320     If Not IsNull(tb!RetDate) Then
330         s = s & Format(tb!RetDate, "dd/mmm/yyyy")
340     End If
350     s = s & vbTab & Trim(tb!SapCode & "")
360     If Trim(tb!Valid & "") <> "" And tb!Valid & "" = 1 Then
370         s = s & vbTab & "V"
380     End If
390     grdExt.AddItem s


400     tb.MoveNext
410 Loop
420 FixG grdExt

430 With grdExt
440     bValidateImm(2).Caption = "VALID"
450     For n = 0 To 8
460         txtEtc(n).Locked = True
470     Next n
480     For i = 1 To .Rows - 1
            '        Frame2.Enabled = False
            '        If .TextMatrix(n, 3) = "X" Then
            '            .Row = n
            '            .Col = 1
            '            .CellForeColor = vbWhite
            '            .CellBackColor = vbBlack
            '        End If
490         If InStr(.TextMatrix(i, 8), "V") = "0" Then
500             bValidateImm(2).Caption = "&Validate"
510             For n = 0 To 8
520                 txtEtc(n).Locked = False
530             Next n
540         End If
550     Next
560 End With

570 For n = 0 To 8
580     txtEtc(n) = ""
590 Next

600 sql = "SELECT * from etc WHERE sampleid = '" & txtSampleID & "'"
610 Set tb = New Recordset
620 RecOpenServer 0, tb, sql
630 If Not tb.EOF Then
640     txtEtc(0) = tb!etc0 & ""
650     txtEtc(1) = tb!etc1 & ""
660     txtEtc(2) = tb!etc2 & ""
670     txtEtc(3) = tb!etc3 & ""
680     txtEtc(4) = tb!etc4 & ""
690     txtEtc(5) = tb!etc5 & ""
700     txtEtc(6) = tb!etc6 & ""
710     txtEtc(7) = tb!etc7 & ""
720     txtEtc(8) = tb!etc8 & ""
730 End If

740 sql = "SELECT Count(*) AS Cnt FROM MediBridgeResults WHERE SampleID = " & txtSampleID
750 Set tb = New Recordset
760 RecOpenServer 0, tb, sql
770 If tb!Cnt > 0 Then
780     baddtotests(1).BackColor = vbYellow
790 Else
800     baddtotests(1).BackColor = &H8000000F
810 End If

820 cmdSaveImm(2).Enabled = False
830 UpDown1.Enabled = True

840 Exit Sub

LoadExt_Error:

    Dim strES As String
    Dim intEL As Integer

850 intEL = Erl
860 strES = Err.Description
870 LogError "frmEditAll", "LoadExt", intEL, strES, sql

End Sub

Public Sub LoadHaematology()

      Dim tb As New Recordset
      Dim sn As New Recordset
      Dim n As Long
      Dim ip As String
      Dim e As String
      Dim PrevDate As String
      Dim PrevID As String
      Dim sql As String

10    On Error GoTo LoadHaematology_Error

20    ReDim i(0 To 6) As String
      'Dim HD As HaemTestDefinition
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
      Dim Dob As String
      Dim ThisValid As Long
      Dim GR As Long
      Dim Asql As String
      Dim Csql As String
      Dim Nsql As String
      Dim DaysOld As Long
      Dim SA As Long
      Dim Plt As String
      Dim x As Long
      Dim g As GenericResult


30    ClearHgb

40    Panel3D4.Enabled = True
50    Panel3D5.Enabled = True
60    Panel3D6.Enabled = True
      'Haemlock
70    txtHaemComment.Locked = False
80    Panel3D7.Enabled = True
90    cmdUnvalPrint.Visible = True

100   bHaemGraphs.Visible = False
110   bViewHaemRepeat.Visible = False
120   PreviousHaem = False
130   HistHaem = False
140   ssTabAll.TabCaption(1) = "Haematology"
150   lHaemErrors.Visible = False
160   bValidateHaem.Caption = "&Validate"
170   SA = 0
180   lblHaemPrinted = ""
190   txtCondition = ""
200   lblHaemValid.Visible = True
210   txtEsr1 = ""
220   txtEsr1.Visible = False
230   lWIC = ""
240   lWOC = ""
250   lblRepeats.Visible = False

260   grdH.Height = 2000

270   If Trim$(txtSampleID) = "" Then Exit Sub

280   Dob = txtDoB
290   If Dob <> "" Then DaysOld = Abs(DateDiff("d", Format(lblSampleDate, "dd/MMM/yyyy"), txtDoB)) Else DaysOld = 12783

300   If txtName <> "" And txtDoB <> "" And IsDate(txtDoB) Then
310       sql = CreateHist("haem")
320       Set sn = New Recordset
330       RecOpenServer 0, sn, sql
340       If Not sn.EOF Then
350           HistHaem = True
360       End If

370       If HistHaem = True Then
380           sql = CreateSql("haem")
390           Set sn = New Recordset
400           RecOpenServer 0, sn, sql
410           If Not sn.EOF Then
420               PrevDate = sn!SampleDate
430               PrevID = sn!SampleID
440               sql = "SELECT * from HaemResults WHERE " & _
                        "SampleID = '" & PrevID & "'"
450               Set tb = New Recordset
460               RecOpenServer 0, tb, sql
470               If Not tb.EOF Then
480                   PreviousHaem = True
490                   PrevRBC = Val(tb!rbc & "")
500                   PrevHgb = Val(tb!Hgb & "")
510                   PrevMCV = Val(tb!MCV & "")
520                   PrevHct = Val(tb!hct & "")
530                   PrevRDWCV = Val(tb!RDWCV & "")
540                   PrevRDWSD = Val(tb!rdwsd & "")
550                   PrevMCH = Val(tb!mch & "")
560                   PrevMCHC = Val(tb!mchc & "")
570                   Prevplt = Val(tb!Plt & "")
580                   PrevMPV = Val(tb!mpv & "")
590                   PrevPLCR = Val(tb!plcr & "")
600                   PrevPdw = Val(tb!pdw & "")
610                   PrevWBC = Val(tb!wbc & "")
620                   PrevLymA = Val(tb!LymA & "")
630                   PrevLymP = Val(tb!LymP & "")
640                   PrevMonoA = Val(tb!MonoA & "")
650                   PrevMonoP = Val(tb!MonoP & "")
660                   PrevNeutA = Val(tb!NeutA & "")
670                   PrevNeutP = Val(tb!NeutP & "")
680                   PrevEosA = Val(tb!EosA & "")
690                   PrevEosP = Val(tb!EosP & "")
700                   PrevBasA = Val(tb!BasA & "")
710                   PrevBasP = Val(tb!BasP & "")
720                   PrevLucA = Val(tb!luca & "")
730                   PrevLucp = Val(tb!lucp & "")
740                   PrevChcm = Val(tb!cH & "")
750                   PrevHDW = Val(tb!HDW & "")
760               End If
770           End If
780       End If
790   End If

800   sql = "SELECT * from HaemResults WHERE " & _
            "SampleID = '" & txtSampleID & "'"

810   Set tb = New Recordset
820   RecOpenServer 0, tb, sql

830   If tb.EOF Then
840       bValidateHaem.Enabled = False
850       lblAgeSex.Visible = False
860       lblHaemValid.Visible = False
870   Else
880       If txtDoB = "" Then
890           SA = 2
900       End If

910       If txtSex = "" Then
920           SA = SA + 1
930       End If
940       Select Case SA
              Case 3: lblAgeSex = "Ref Range Not Age/Sex Related"
950               lblAgeSex.Visible = True
960           Case 2: lblAgeSex = "Ref Range Not Age Related"
970               lblAgeSex.Visible = True
980           Case 1: lblAgeSex = "Ref Range Not Sex Related"
990               lblAgeSex.Visible = True
1000          Case Else
1010              lblAgeSex.Visible = False
1020      End Select
1030      bValidateHaem.Enabled = True
1040      If SysOptHaemAn1(0) = "ADVIA" And Trim(tb!Analyser) & "" = "1" Or SysOptHaemAn2(0) = "ADVIA" Or SysOptHaemAn2(0) = "ADVIA60" Then
1050          sql = "SELECT * from HaemFlags WHERE " & _
                    "sampleid = '" & txtSampleID & "'"
1060          Set sn = New Recordset
1070          RecOpenServer 0, sn, sql
1080          If Not sn.EOF Then
1090              If Trim(sn!Flags) = "" Or IsNull(sn!Flags) Then
1100                  lHaemErrors.Visible = False
1110              Else
1120                  lHaemErrors.Visible = True
1130              End If
1140          Else
1150              lHaemErrors.Visible = False
1160          End If
1170      Else
1180          If Not IsNull(tb!LongError) Then
1190              If Val(tb!LongError) > 1 Then
1200                  lHaemErrors.Visible = True
1210                  lHaemErrors.Tag = Format$(tb!LongError)
1220              End If
1230          End If
1240      End If
1250      If Not IsNull(tb!gwb1) Or Not IsNull(tb!gwb2) Or Not IsNull(tb!gRbc) Or Not IsNull(tb!gplt) Then
1260          bHaemGraphs.Visible = True
1270      Else
1280          bHaemGraphs.Visible = True
1290      End If

1300      pdelta.Cls
1310      lHDate = Format(tb!RunDateTime, "dd/MM/yyyy hh:mm:ss")
1320      If tb!wic & "" <> "" Then
1330          lWIC = Trim(tb!wic & "")
1340          lWOC = Trim(tb!woc & "")
1350          Label1(17) = "WIC"
1360          Label1(18) = "WOC"
1370      Else
1380          Label1(17) = "WCBC"
1390          Label1(18) = "WCPC"
1400      End If
1410      If lWIC = "" Then lWIC = Trim(tb!wb & "")
1420      If lWOC = "" Then lWOC = Trim(tb!wp & "")

1430      cFilm = 0
1440      If Not IsNull(tb!cFilm) Then
1450          cFilm = IIf(tb!cFilm, 1, 0)
1460      End If

1470      gRbc.Visible = False

1480      If Not IsNull(tb!rbc) Then
1490          ColouriseG "RBC", gRbc, 1, 1, Trim(tb!rbc), txtSex, Dob
1500          If PreviousHaem Then DeltaCheck "RBC", tb!rbc, PrevRBC, PrevDate, PrevID
1510          gRbc.TextMatrix(1, 2) = HNR("RBC", DaysOld, txtSex)
1520      End If

1530      If Not IsNull(tb!Hgb) Then
1540          ColouriseG "Hgb", gRbc, 2, 1, Trim(tb!Hgb), txtSex, Dob
1550          gRbc.Row = 2
1560          gRbc.Col = 1
1570          gRbc.CellFontSize = 12
1580          If PreviousHaem Then DeltaCheck "Hgb", tb!Hgb, PrevHgb, PrevDate, PrevID
1590          gRbc.TextMatrix(2, 2) = HNR("hgb", DaysOld, txtSex)
1600      End If

1610      If Not IsNull(tb!hct) Then
1620          ColouriseG "Hct", gRbc, 3, 1, Trim(tb!hct), txtSex, Dob
1630          If PreviousHaem Then DeltaCheck "Hct", tb!hct, PrevHct, PrevDate, PrevID
1640          gRbc.TextMatrix(3, 2) = HNR("hct", DaysOld, txtSex)
1650      End If

1660      If Not IsNull(tb!MCV) Then
1670          ColouriseG "MCV", gRbc, 4, 1, Trim(tb!MCV), txtSex, Dob
1680          If PreviousHaem Then DeltaCheck "MCV", tb!MCV, PrevMCV, PrevDate, PrevID
1690          gRbc.TextMatrix(4, 2) = HNR("MCV", DaysOld, txtSex)
1700      End If

1710      If SysOptHaemAn1(0) <> "" Then
1720          If Not IsNull(tb!HDW) Then
1730              ColouriseG "HDW", gRbc, 5, 1, Trim(tb!HDW), txtSex, Dob
1740              If PreviousHaem Then DeltaCheck "HDW", tb!HDW, PrevHDW, PrevDate, PrevID
1750              gRbc.TextMatrix(5, 2) = HNR("hdw", DaysOld, txtSex)
1760          End If
1770      End If

1780      If Not IsNull(tb!mch) Then
1790          ColouriseG "MCH", gRbc, 6, 1, Trim(tb!mch), txtSex, Dob
1800          If PreviousHaem Then DeltaCheck "MCH", tb!mch, PrevMCH, PrevDate, PrevID
1810          gRbc.TextMatrix(6, 2) = HNR("mch", DaysOld, txtSex)
1820      End If

1830      If Not IsNull(tb!mchc) Then
1840          ColouriseG "MCHC", gRbc, 7, 1, Trim(tb!mchc), txtSex, Dob
1850          If PreviousHaem Then DeltaCheck "MCHC", tb!mchc, PrevMCHC, PrevDate, PrevID
1860          gRbc.TextMatrix(7, 2) = HNR("mchc", DaysOld, txtSex)
1870          If IsNumeric(tb!mchc) Then
1880              If Val(tb!mchc & "") > 36 Then
1890                  If tb!he & "" = "" Then
1900                      sql = "UPDATE HaemResults SET he = '+' WHERE SampleID = '" & txtSampleID & "'"
1910                      Cnxn(0).Execute sql
1920                  End If
                      '                tb!he = "+"
                      '                tb.Update
1930              End If
1940          End If
1950      End If

1960      If SysOptHaemAn1(0) <> "" Then
1970          If Not IsNull(tb!cH) Then
1980              ColouriseG "CHCM", gRbc, 8, 1, Trim(tb!cH), txtSex, Dob
1990              If PreviousHaem Then DeltaCheck "CHCM", tb!cH, PrevLucp, PrevDate, PrevID
2000              gRbc.TextMatrix(8, 2) = HNR("CHCM", DaysOld, txtSex)
2010          End If
2020      End If

2030      If Not IsNull(tb!RDWCV) And Val(tb!RDWCV & "") <> 0 Then
2040          ColouriseG "RDW", gRbc, 9, 1, Trim(tb!RDWCV), txtSex, Dob
2050          If PreviousHaem Then DeltaCheck "RDW", tb!RDWCV, PrevRDWCV, PrevDate, PrevID
2060          gRbc.TextMatrix(9, 2) = HNR("rdw", DaysOld, txtSex)
2070      End If

2080      If Not IsNull(tb!Plt) Then
2090          Plt = Trim(tb!Plt)
2100          Colourise "Plt", tPlt, Trim(Plt), txtSex, Dob
2110          If PreviousHaem Then DeltaCheck "plt", Plt, Prevplt, PrevDate, PrevID
2120      End If

2130      If Not IsNull(tb!mpv) Then
2140          Colourise "MPV", tMPV, Trim(tb!mpv), txtSex, Dob
2150          If PreviousHaem Then DeltaCheck "MPV", tb!mpv, PrevMPV, PrevDate, PrevID
2160      End If

2170      grdH.Visible = False

2180      If Not IsNull(tb!wbc) Then
2190          Colourise "WBC", tWBC, tb!wbc, txtSex, Dob
2200          If PreviousHaem Then DeltaCheck "WBC", tb!wbc, PrevWBC, PrevDate, PrevID
2210          If SysOptWBCDC(0) = True Then
2220              tWBC = Format(tb!wbc, "##0.0")
2230          Else
2240              tWBC = Format(tb!wbc, "##0.00")
2250          End If
2260      End If

          'Diff

2270      If Not IsNull(tb!NeutA) Then
2280          ColouriseG "NeutA", grdH, 1, 0, Trim(tb!NeutA & ""), txtSex, Dob
2290          If PreviousHaem Then DeltaCheck "NeutA", tb!NeutA, PrevNeutA, PrevDate, PrevID
2300          grdH.TextMatrix(1, 1) = HNR("neuta", DaysOld, txtSex)
2310      End If

2320      If Not IsNull(tb!NeutP) Then
2330          ColouriseG "NeutP", grdH, 1, 3, Trim(tb!NeutP & ""), txtSex, Dob
2340          If PreviousHaem Then DeltaCheck "NeutP", tb!NeutP, PrevNeutP, PrevDate, PrevID
2350      End If

2360      If Not IsNull(tb!LymA) Then
2370          ColouriseG "LymA", grdH, 2, 0, Trim(tb!LymA & ""), txtSex, Dob
2380          If PreviousHaem Then DeltaCheck "LymA", tb!LymA, PrevLymA, PrevDate, PrevID
2390          grdH.TextMatrix(2, 1) = HNR("lyma", DaysOld, txtSex)
2400      End If

2410      If Not IsNull(tb!LymP) Then
2420          ColouriseG "LymP", grdH, 2, 3, Trim(tb!LymP & ""), txtSex, Dob
2430          If PreviousHaem Then DeltaCheck "LymP", tb!LymP, PrevLymP, PrevDate, PrevID
2440      End If

2450      If Not IsNull(tb!MonoA) Then
2460          ColouriseG "MonoA", grdH, 3, 0, Trim(tb!MonoA & ""), txtSex, Dob
2470          If PreviousHaem Then DeltaCheck "MonoA", tb!MonoA, PrevMonoA, PrevDate, PrevID
2480          grdH.TextMatrix(3, 1) = HNR("monoa", DaysOld, txtSex)
2490      End If

2500      If Not IsNull(tb!MonoP) Then
2510          ColouriseG "MonoP", grdH, 3, 3, Trim(tb!MonoP & ""), txtSex, Dob
2520          If PreviousHaem Then DeltaCheck "MonoP", tb!MonoP, PrevMonoP, PrevDate, PrevID
2530      End If

2540      If Not IsNull(tb!EosA) Then
2550          ColouriseG "EosA", grdH, 4, 0, Trim(tb!EosA & ""), txtSex, Dob
2560          If PreviousHaem Then DeltaCheck "EosA", tb!EosA, PrevEosA, PrevDate, PrevID
2570          grdH.TextMatrix(4, 1) = HNR("eosa", DaysOld, txtSex)
2580      End If

2590      If Not IsNull(tb!EosP) Then
2600          ColouriseG "EosP", grdH, 4, 3, Trim(tb!EosP & ""), txtSex, Dob
2610          If PreviousHaem Then DeltaCheck "EosP", tb!EosP, PrevEosP, PrevDate, PrevID
2620      End If

2630      If Not IsNull(tb!BasA) Then
2640          ColouriseG "BasA", grdH, 5, 0, Trim(tb!BasA & ""), txtSex, Dob
2650          If PreviousHaem Then DeltaCheck "BasA", tb!BasA, PrevBasA, PrevDate, PrevID
2660          grdH.TextMatrix(5, 1) = HNR("Basa", DaysOld, txtSex)
2670      End If

2680      If Not IsNull(tb!BasP) Then
2690          ColouriseG "BasP", grdH, 5, 3, Trim(tb!BasP & ""), txtSex, Dob
2700          If PreviousHaem Then DeltaCheck "BasP", tb!BasP, PrevBasP, PrevDate, PrevID
2710      End If

2720      If SysOptHaemAn1(0) <> "" Then
2730          If Not IsNull(tb!luca) Then
2740              ColouriseG "LucA", grdH, 6, 0, Trim(tb!luca & ""), txtSex, Dob
2750              If PreviousHaem Then DeltaCheck "LucA", tb!luca, PrevLucA, PrevDate, PrevID
2760              grdH.TextMatrix(6, 1) = HNR("luca", DaysOld, txtSex)
2770          End If

2780          If Not IsNull(tb!lucp) Then
2790              ColouriseG "LucP", grdH, 6, 3, Trim(tb!lucp & ""), txtSex, Dob
2800              If PreviousHaem Then DeltaCheck "LucP", tb!lucp, PrevLucp, PrevDate, PrevID
2810          End If
2820          tASOt = Trim(tb!tASOt & "")
2830          tRa = Trim(tb!tRa & "")
2840          If Not IsNull(tb!cASot) Then
2850              cASot = IIf(tb!cASot, 1, 0)
2860          Else
2870              cASot = 0
2880          End If
              '            If Trim(tb!Analyser) = "1" Then
              '                lblAnalyser = "Analyser : " & SysOptHaemN1(0)
              '                HaemAnalyser = "1"
              '            ElseIf Trim(tb!Analyser) = "2" Then
              '                HaemAnalyser = "2"
              '                lblAnalyser = "Analyser : " & SysOptHaemN2(0)
              '            End If
2890          lblAnalyser = "Analyser : " & Trim(tb!Analyser)
2900      End If

2910      If SysOptHaemAn1(0) = "ADVIA" Then
2920          If GetOptionSetting("EnableAdviaOldFlags", 1) = 1 Then
2930            If Trim(tb!LS & "") <> "" Or Trim(tb!va & "") <> "" _
                    Or Trim(tb!At & "") <> "" Or Trim(tb!bl & "") <> "" _
                    Or Trim(tb!An & "") <> "" Or Trim(tb!mi & "") <> "" _
                    Or Trim(tb!ca & "") <> "" Or Trim(tb!ho & "") <> "" _
                    Or Trim(tb!he & "") <> "" Or Trim(tb!Ig & "") <> "" _
                    Or Trim(tb!mpo & "") <> "" Or Trim(tb!lplt & "") <> "" _
                    Or Trim(tb!pclm & "") <> "" Or Trim(tb!rbcf & "") <> "" _
                    Or Trim(tb!rbcg & "") <> "" Then
2940              lHaemErrors.Visible = True
2950            End If
2960          End If
2970          txtLI = Trim(tb!Li & "")
2980          txtMPXI = Trim(tb!mpxi & "")
              'gRBC.TextMatrix(11, 1) = Trim(tb!ho & "")
2990      End If

          '2950      If SysOptBadRes(0) Then
3000      If Not IsNull(tb!cbad) Then
3010          chkBad = IIf(tb!cbad, 1, 0)
3020      Else
3030          chkBad = 0
3040      End If
          '3010      End If
          '
3050      gRbc.TextMatrix(10, 1) = Trim(tb!nrbcp & "")

          '  tESR = Trim(tb!esr & "")
3060      If Trim(tb!esr & "") <> "" Then
3070            Colourise "ESR", tESR, Trim(tb!esr), txtSex, Dob
3080      Else
3090            tESR = ""
3100      End If



3110      If Trim(tb!reta & "") <> "" Then
3120          Colourise "RETA", tRetA, Trim(tb!reta), txtSex, Dob
3130      Else
3140          tRetA = ""
3150      End If

3160      tRetP = Trim(tb!RetP) & ""

3170      Select Case Trim$(tb!Monospot & "")
              Case "P": tMonospot = "Positive"
3180          Case "N": tMonospot = "Negative"
3190          Case "I": tMonospot = "Inconclusive"
3200          Case Else: tMonospot = ""
3210      End Select



3270      If Not IsNull(tb!cRA) Then
3280          cRA = IIf(tb!cRA, 1, 0)
3290      Else
3300          cRA = 0
3310      End If

          '3210      If SysOptBadRes(0) Then
          '3220          If Not IsNull(tb!cbad) Then
          '3230              chkBad = IIf(tb!cbad, 1, 0)
          '3240          Else
          '3250              chkBad = 0
          '3260          End If
          '3270      End If

3320      If Not IsNull(tb!cRetics) Then
3330          cRetics = IIf(tb!cRetics, 1, 0)
3340      Else
3350          cRetics = 0
3360      End If
3220      If Not IsNull(tb!cESR) Then
3230          cESR = IIf(tb!cESR, 1, 0)
3240      Else
3250          cESR = 0
3260      End If

3370      If Not IsNull(tb!cMonospot) Then
3380          cMonospot = IIf(tb!cMonospot, 1, 0)
3390      Else
3400          cMonospot = 0
3410      End If

3420      If Not IsNull(tb!cmalaria) Then
3430          chkMalaria = IIf(tb!cmalaria, 1, 0)
3440      Else
3450          chkMalaria = 0
3460      End If
3470      lblMalaria = Trim(tb!Malaria & "")

3480      If Not IsNull(tb!csickledex) Then
3490          chkSickledex = IIf(tb!csickledex, 1, 0)
3500      Else
3510          chkSickledex = 0
3520      End If
3530      lblSickledex = Trim(tb!Sickledex & "")

3540      tWarfarin = Trim(tb!Warfarin & "")

3550      ip = Left$(tb!ipmessage & "000000", 6)
3560      For n = 0 To 5
3570          ipflag(n).Enabled = Mid$(ip, n + 1, 1) = "1"
3580      Next


3590      e = tb!negposerror & ""

3600      buildinterp tb, i()
3610      If i(0) <> "" Then pdelta.Print
3620      For n = 0 To 6
3630          pdelta.ForeColor = vbRed
3640          pdelta.Print i(n)
3650      Next

3660      ThisValid = False
3670      If Not IsNull(tb!Valid) Then
3680          ThisValid = IIf(tb!Valid, 1, 0)
3690      End If
3700      If ThisValid = 1 Then
3710          bValidateHaem.Caption = "VALID"
3720          lblUrgent.Visible = False
3730      Else
3740          bValidateHaem.Caption = "&Validate"
3750          lblUrgent.Visible = UrgentTest
3760      End If
          'Zyam added a condition for ESR for portloaise 20-05-24
'          If ThisValid = 1 And Trim(tb!rbc) & "" <> "" Then
'            tESR.Text = ""
'          End If
          'Zyam
3770      If ThisValid = 1 Then lblHaemValid.Visible = True Else lblHaemValid.Visible = False
3780      If ThisValid = 1 Then Panel3D4.Enabled = False Else Panel3D4.Enabled = True
3790      If ThisValid = 1 Then Panel3D5.Enabled = False Else Panel3D5.Enabled = True
3800      If ThisValid = 1 Then Panel3D6.Enabled = False Else Panel3D6.Enabled = True
          'Haemlock
3810      If ThisValid = 1 Then txtHaemComment.Locked = True Else txtHaemComment.Locked = False
3820      If ThisValid = 1 Then Panel3D7.Enabled = False Else Panel3D7.Enabled = True
3830      If ThisValid = 1 Then cmdUnvalPrint.Visible = False Else cmdUnvalPrint.Visible = True

3840      If Not IsNull(tb!Printed) Then
3850          If tb!Printed = 1 Then
3860              lblHaemPrinted = "Already Printed"
3870          Else
3880              lblHaemPrinted = "Not Printed"
3890          End If
3900      Else
3910          lblHaemPrinted = "Not Printed"
3920      End If

          'QMS Ref # 818576
          '***BLR: If haemrepeats are available, always show repeat button with yellow color
          '***    if haemrepeats are available and either it's unvalidated or VIEW options in
          '***    option table is set to 1 then enable repeat button

3930      sql = "SELECT * from HaemRepeats WHERE " & _
                "SampleID = '" & txtSampleID & "'"
3940      Set tb = New Recordset
3950      RecOpenServer 0, tb, sql
3960      If Not tb.EOF Then
3970          bViewHaemRepeat.Visible = True
3980          lblRepeats.Visible = True
3990          If SysOptView(0) = True Or ThisValid = 0 Then
'+++Zyam 14-5-24
'4000              If tb!wbc & "" <> "" Or tb!reta & "" <> "" Then
4010                  bViewHaemRepeat.Enabled = True
'4020              End If
'+++Zyam 14-5-24
4030          Else
4040              bViewHaemRepeat.Enabled = False
4050          End If

4060      End If

          '    If ThisValid = 0 Then
          '        sql = "SELECT * from HaemRepeats WHERE " & _
                   '              "SampleID = '" & txtSampleID & "'"
          '        Set tb = New Recordset
          '        RecOpenServer 0, tb, sql
          '        If Not tb.EOF Then
          '            If tb!wbc & "" <> "" Or tb!reta & "" <> "" Then bViewHaemRepeat.Visible = True
          '        End If
          '    End If



4070      If Trim(txtChart) <> "" Then
4080          sql = "SELECT * from HaemCondition WHERE " & _
                    "chart = '" & txtChart & "'"
4090          Set tb = New Recordset
4100          RecOpenServer 0, tb, sql
4110          If Not tb.EOF Then
4120              txtCondition = Trim(tb!condition)
4130          End If
4140      End If

4150      ssTabAll.TabCaption(1) = ">>Haematology<<"

4160  End If

4170  grdH.Visible = True
4180  gRbc.Visible = True

4190  sql = "SELECT * from Differentials WHERE " & _
            "runnumber = '" & txtSampleID & "'"
4200  Set tb = New Recordset
4210  RecOpenServer 0, tb, sql
4220  If tb.EOF Then
4230      bFilm.BackColor = &H8000000F
4240  Else
4250      bFilm.BackColor = vbBlue
4260      If tb!prndiff = True Then
4270          grdH.Height = 360
4280      End If
4290  End If

4300  FixG gRbc
4310  g = LoadGenericResult(txtSampleID, "Viscosity37")
4320  txtViscosity = g.Result
4330  txtReadingDateTime = g.TestDateTime


4340  cmdSaveHaem.Enabled = False
4350  cmdSaveComm.Enabled = False
4360  cmdHSaveH.Enabled = False

4370  bFAX.Enabled = (UCase$(bValidateHaem.Caption) = "VALID")

4380  LoadComments

4390  Exit Sub

LoadHaematology_Error:

      Dim strES As String
      Dim intEL As Integer

4400  intEL = Erl
4410  strES = Err.Description
4420  LogError "frmEditAll", "LoadHaematology", intEL, strES, sql
4430  grdH.Visible = True
4440  gRbc.Visible = True

End Sub

Private Function LoadImmSplitList(ByVal Index As Integer) As String

    Dim tb As New Recordset
    Dim sql As String
    Dim strIndex As String
    Dim strReturn As String

10  On Error GoTo LoadImmSplitList_Error

20  strIndex = Index

30  sql = "SELECT distinct Code, PrintPriority, SplitList " & _
          "from ImmTestDefinitions " & _
          "WHERE SplitList = " & strIndex & " " & _
          "order by PrintPriority"

40  Set tb = New Recordset
50  RecOpenServer 0, tb, sql

60  strReturn = ""
70  Do While Not tb.EOF
80      strReturn = strReturn & "Code = '" & tb!Code & "' or "
90      tb.MoveNext
100 Loop
110 If strReturn <> "" Then
120     strReturn = Left$(strReturn, Len(strReturn) - 3)
130 End If

140 LoadImmSplitList = strReturn

150 Exit Function

LoadImmSplitList_Error:

    Dim strES As String
    Dim intEL As Integer

160 intEL = Erl
170 strES = Err.Description
180 LogError "frmEditAll", "LoadImmSplitList", intEL, strES, sql

End Function

Public Sub LoadImmunology()

    Dim Deltasn As Recordset
    Dim tb As New Recordset
    Dim sql As String
    Dim s As String
    Dim Value As Single
    Dim OldValue As Single
    Dim valu As String
    Dim OldValu As String
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

s = CheckAutoComments(txtSampleID, "Immunology")


gImm(1).ColWidth(10) = 0

PreviousImm = False
HistImm = False
lblIRundate = ""
cmdGetBio.Visible = True
Frame12(1).Enabled = True
txtImmComment(1).Locked = False
    '
Fasting = lImmRan(1) = "Fasting Sample"

ClearFGrid gImm(1)

lIDelta(1) = ""
bViewImmRepeat(1).Visible = False

ssTabAll.TabCaption(6) = "Immunology"

    'get date & run number of previous record
PreviousImm = False

If txtName <> "" And txtDoB <> "" And IsDate(txtDoB) Then

  sql = CreateHist("imm")
  Set sn = New Recordset
  RecOpenServer 0, sn, sql
  If Not sn.EOF Then
      HistImm = True
  End If
  '
  '    sql = CreateSql("Imm")
  '    Set Deltatb = New Recordset
  '    RecOpenServer 0, Deltatb, sql
  '    If Not Deltatb.EOF Then
  '        PreviousDate = Deltatb!Rundate & ""
  '        PreviousRec = Deltatb!SampleID & ""
  '        PreviousImm = True
  '    End If
End If

If cCat(0) = "" Then Cat = "Default" Else Cat = cCat(0)

Set IMres = Ims.Load("Imm", txtSampleID, "Results", gDONTCARE, gDONTCARE, 0, Cat, dtRunDate)

CheckCalcIPSA IMres

If Not IMres Is Nothing Then
  ssTabAll.TabCaption(6) = ">>Immunology<<"
  For Each Im In IMres
      lblIRundate = Format(GetLatestRunDateTime("Imm", Im.SampleID, Im.RunTime), "dd/MM/yyyy hh:mm:ss")
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
      If txtSex = "" Or (Val(Im.Low) = 0 And Val(Im.High) = 0) Or (Val(Im.Low) = 0 And Val(Im.High) = 999) Or (Val(Im.Low) = 0 And Val(Im.High) = 9999) Then      'QMS Ref No. #817982
          s = s & vbTab
      Else
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
              ElseIf InStr(1, Im.Result, "<") > 0 Then
                  Flag = "L"
                  s = s & "L"
              End
              ElseIf InStr(1, Im.Result, ">") > 0 Then
                  Flag = "H"
                  s = s & "H"
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
      End If
      e = ""
      s = s & vbTab & _
          IIf(e <> "", e, "") & vbTab & _
          IIf(Im.Valid, "V", " ") & _
          IIf(Im.Printed, "P", " ")
      s = s & vbTab & Im.Pc & vbTab & Im.Comment
      s = s & vbTab & vbTab & Im.LongName
      gImm(1).AddItem s
      If Flag <> "" Then
          gImm(1).Row = gImm(1).Rows - 1
          gImm(1).Col = 1
          Select Case Flag
          Case "N", "H":
              For n = 0 To 8
                  gImm(1).Col = n
                  gImm(1).CellBackColor = SysOptHighBack(0)
                  gImm(1).CellForeColor = SysOptHighFore(0)
              Next
          Case "E", "L":
              For n = 0 To 8
                  gImm(1).Col = n
                  gImm(1).CellBackColor = SysOptLowBack(0)
                  gImm(1).CellForeColor = SysOptLowFore(0)
              Next
          Case "X":
              For n = 0 To 8
                  gImm(1).Col = n
                  gImm(1).CellBackColor = SysOptPlasBack(0)
                  gImm(1).CellForeColor = SysOptPlasFore(0)
              Next
          End Select
      End If
      Flag = ""
      If txtName <> "" And txtDoB <> "" And IsDate(txtDoB) Then
          Set Deltasn = DoDeltaCheck("Imm", Im.Code)
          If Im.DoDelta And (Not Deltasn.EOF) Then
              If (dtSampleDate - CDate(Format(Deltasn!SampleDate, "dd/mm/yyyy"))) <= Im.CheckTime Then
                  If IsNumeric(valu) Then
                      OldValue = Val(Deltasn!Result)
                  Else
                      If Not IsNumeric(Deltasn!Result) Then
                          OldValu = Deltasn!Result
                      Else
                          OldValu = ""
                      End If
                  End If
                  If OldValue <> 0 Then
                      DeltaLimit = Im.DeltaLimit
                      If Abs(OldValue - Value) > DeltaLimit Then
                          Res = Format$(Deltasn!SampleDate, "dd/mm/yyyy") & " (" & Deltasn!SampleID & ") " & _
                                Im.ShortName & " " & _
                                OldValue & vbCr
                          lIDelta(1) = lIDelta(1) & Res
                      End If
                  Else
                      If UCase(OldValu) <> UCase(valu) Then
                          Res = Format$(Deltasn!SampleDate, "dd/mm/yyyy") & " (" & Deltasn!SampleID & ") " & _
                                Im.ShortName & " " & _
                                OldValu & vbCr
                          lIDelta(1) = lIDelta(1) & Res
                      End If

                  End If

              End If
          End If
      End If
      OldValu = ""
      OldValue = 0

  Next
End If

FixG gImm(1)

With gImm(1)
  bValidateImm(1).Caption = "VALID"
  lblUrgent.Visible = False
  txtImmComment(1).Locked = True
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
          txtImmComment(1).Locked = False
      End If
  Next
End With

LoadOutstandingImm

sql = "SELECT * from ImmRepeats WHERE " & _
     "SampleID = '" & Val(txtSampleID) & "'"
Set tb = New Recordset
RecOpenServer 0, tb, sql
bViewImmRepeat(1).Visible = False
If Not tb.EOF Then
  bViewImmRepeat(1).Visible = True
End If

sql = "SELECT * from ImmMasks WHERE " & _
     "SampleID = '" & txtSampleID & "'"
Set tb = New Recordset
RecOpenServer 0, tb, sql
If Not tb.EOF Then
  Ih(1) = IIf(tb!h, 1, 0)
  Iis(1) = IIf(tb!s, 1, 0)
  Il(1) = IIf(tb!l, 1, 0)
  Io(1) = IIf(tb!o, 1, 0)
  Ig(1) = IIf(tb!g, 1, 0)
  Ij(1) = IIf(tb!J, 1, 0)
End If

If gImm(1).Rows > 2 And gImm(1).TextMatrix(1, 0) <> "" Then
  sql = "SELECT * from bioresults WHERE sampleid = '" & txtSampleID & "' and " & _
      " (code = '" & SysOptBioCodeForAlb(0) & "' or " & _
      " code = '" & SysOptBioCodeForUProt(0) & "' or " & _
      " code = '" & SysOptBioCodeForGlob(0) & "' or " & _
      " code = '" & SysOptBioCodeFor24UProt(0) & "' or " & _
      " code = '" & SysOptBioCodeForTProt(0) & "' or " & _
      " code = '" & SysOptBioCodeFor24Vol(0) & "') "

  Set tb = New Recordset
  RecOpenServer 0, tb, sql
  If Not tb.EOF Then cmdGetBio.Visible = True Else cmdGetBio.Visible = False
End If

SetPrintInhibit "Imm"
CheckAuditTrail
EnableBarCodePrinting
CheckLabLinkStatus
bFAX.Enabled = (UCase$(bValidateImm(1).Caption) = "VALID")

    'If InStr(txtImmComment(1), s) = 0 And UCase(bValidateImm(1).Caption) = "&VALIDATE" Then
    '    txtImmComment(1) = s & vbCrLf & txtImmComment(1)
    'End If
    'SaveComments

LoadComments

If txtName <> "" Then txtName.SetFocus

Exit Sub

LoadImmunology_Error:

    Dim strES As String
    Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "frmEditAll", "LoadImmunology", intEL, strES, sql

End Sub
Private Sub LoadOutstandingBio()

    Dim tb As New Recordset
    Dim sql As String

10  On Error GoTo LoadOutstandingBio_Error

20  ClearOutstanding grdOutstanding

30  sql = "SELECT DISTINCT(ShortName) FROM BioRequests AS R, BioTestDefinitions AS D WHERE " & _
          "R.SampleID = '" & txtSampleID & "' " & _
          "AND D.Code = R.Code " & _
          "AND D.SampleType = R.SampleType"
40  Set tb = New Recordset
50  RecOpenServer 0, tb, sql
60  Do While Not tb.EOF
70      grdOutstanding.AddItem Trim(tb!ShortName & "")
80      tb.MoveNext
90  Loop

100 If grdOutstanding.Rows > 2 Then
110     grdOutstanding.RemoveItem 1
120 End If

130 Exit Sub

LoadOutstandingBio_Error:

    Dim strES As String
    Dim intEL As Integer

140 intEL = Erl
150 strES = Err.Description
160 LogError "frmEditAll", "LoadOutstandingBio", intEL, strES, sql

End Sub

Private Sub LoadOutstandingEnd()

    Dim tb As New Recordset
    Dim sql As String

10  On Error GoTo LoadOutstandingEnd_Error

20  ClearOutstanding grdOutstandings(0)

30  sql = "SELECT * FROM EndRequests WHERE " & _
          "SampleID = '" & txtSampleID & "'"
40  Set tb = New Recordset
50  RecOpenServer 0, tb, sql
60  Do While Not tb.EOF
70      grdOutstandings(0).AddItem EndShortNameFor(tb!Code & "")
80      tb.MoveNext
90  Loop

100 If grdOutstandings(0).Rows > 2 Then
110     grdOutstandings(0).RemoveItem 1
120 End If

130 Exit Sub

LoadOutstandingEnd_Error:

    Dim strES As String
    Dim intEL As Integer

140 intEL = Erl
150 strES = Err.Description
160 LogError "frmEditAll", "LoadOutstandingEnd", intEL, strES, sql

End Sub

Private Sub LoadOutstandingImm()

    Dim tb As New Recordset
    Dim sql As String

10  On Error GoTo LoadOutstandingImm_Error

20  ClearOutstanding grdOutstandings(1)

30  If txtSampleID = "" Then Exit Sub

40  sql = "SELECT * from ImmRequests WHERE " & _
          "sampleid = '" & txtSampleID & "'"
50  Set tb = New Recordset
60  RecOpenServer 0, tb, sql
70  Do While Not tb.EOF
80      grdOutstandings(1).AddItem ImmShortNameFor(tb!Code & "")
90      tb.MoveNext
100 Loop

110 If grdOutstandings(1).Rows > 2 Then
120     grdOutstandings(1).RemoveItem 1
130 End If

140 Exit Sub

LoadOutstandingImm_Error:

    Dim strES As String
    Dim intEL As Integer

150 intEL = Erl
160 strES = Err.Description
170 LogError "frmEditAll", "LoadOutstandingImm", intEL, strES, sql

End Sub

Private Sub LoadOutstandingrdCoag()

    Dim tb As New Recordset
    Dim sql As String

10  On Error GoTo LoadOutstandingrdCoag_Error

20  With grdOutstandingCoag
30      .Rows = 2
40      .AddItem ""
50      .RemoveItem 1
60  End With

70  sql = "SELECT * FROM CoagRequests WHERE " & _
          "SampleID = '" & txtSampleID & "'"
80  Set tb = New Recordset
90  RecOpenServer 0, tb, sql
100 Do While Not tb.EOF
110     grdOutstandingCoag.AddItem CoagNameFor(tb!Code & "")
120     tb.MoveNext
130 Loop

140 If grdOutstandingCoag.Rows > 2 Then
150     grdOutstandingCoag.RemoveItem 1
160 End If

170 Exit Sub

LoadOutstandingrdCoag_Error:

    Dim strES As String
    Dim intEL As Integer

180 intEL = Erl
190 strES = Err.Description
200 LogError "frmEditAll", "LoadOutstandingrdCoag", intEL, strES, sql

End Sub

Private Sub LoadPreviousCoag()

          Dim tb As New Recordset
          Dim sql As String
          Dim CRs As CoagResults
          Dim CR As CoagResult
          Dim PrevDate As String
          Dim PrevID As String
          Dim s As String
          Dim g As String

10        On Error GoTo LoadPreviousCoag_Error

20        PreviousCoag = False

30        ClearFGrid grdPrev

40        sql = CreateSql("Coag")
50        Set tb = New Recordset
60        RecOpenServer 0, tb, sql
70        If Not tb.EOF Then

80            PreviousCoag = True

90            PrevDate = Format$(tb!Rundate, "dd/mm/yy")
100           PrevID = tb!SampleID

110           Set CRs = New CoagResults
120           Set CRs = CRs.Load(PrevID, gDONTCARE, gDONTCARE, Trim(SysOptExp(0)), 0)

130           If Not CRs Is Nothing Then
140               For Each CR In CRs
                      'If Trim(CR.Units) = "INR" Then
                      's = "INR" & vbTab
                      'Else
                      's = CoagNameFor(CR.Code) & vbTab
                      'End If
150                   s = CoagNameFor(CR.Code) & vbTab
160                   Select Case CoagPrintFormat(Trim(CR.Code) & "")
                      Case 0: g = Format$(CR.Result, "0")
170                   Case 1: g = Format$(CR.Result, "0.0")
180                   Case 2: g = Format$(CR.Result, "0.00")
190                   End Select
200                   s = s & g & vbTab & UnitConv(CR.Units)
210                   grdPrev.AddItem s
220               Next
230               lblPrevCoag = PrevDate & " Result for " & txtChart
240           Else
250               lblPrevCoag = "No Previous Coag Details"
260           End If
270       Else
280           lblPrevCoag = "No Previous Coag Details"
290       End If

300       FixG grdPrev

310       Exit Sub

LoadPreviousCoag_Error:

          Dim strES As String
          Dim intEL As Integer

320       intEL = Erl
330       strES = Err.Description
340       LogError "frmEditAll", "LoadPreviousCoag", intEL, strES, sql

End Sub

Private Function LoadSplitList(ByVal Index As Integer) As String

    Dim tb As New Recordset
    Dim sql As String
    Dim strIndex As String
    Dim strReturn As String

10  On Error GoTo LoadSplitList_Error

20  strIndex = Index

30  sql = "SELECT distinct Code, PrintPriority, SplitList " & _
          "from BioTestDefinitions " & _
          "WHERE SplitList = " & strIndex & " " & _
          "order by PrintPriority"

40  Set tb = New Recordset
50  RecOpenClient 0, tb, sql

60  strReturn = ""
70  Do While Not tb.EOF
80      strReturn = strReturn & "Code = '" & tb!Code & "' or "
90      tb.MoveNext
100 Loop
110 If strReturn <> "" Then
120     strReturn = Left$(strReturn, Len(strReturn) - 3)
130 End If

140 LoadSplitList = strReturn

150 Exit Function

LoadSplitList_Error:

    Dim strES As String
    Dim intEL As Integer

160 intEL = Erl
170 strES = Err.Description
180 LogError "frmEditAll", "LoadSplitList", intEL, strES, sql

End Function

Private Sub lRandom_Click()


10  On Error GoTo lRandom_Click_Error

20  If lRandom = "Random Sample" Then
30      lRandom = "Fasting Sample"
40  Else
50      lRandom = "Random Sample"
60  End If

70  LoadBiochemistry

80  cmdSaveBio.Enabled = True



90  Exit Sub

lRandom_Click_Error:

    Dim strES As String
    Dim intEL As Integer

100 intEL = Erl
110 strES = Err.Description
120 LogError "frmEditAll", "lRandom_Click", intEL, strES


End Sub

Private Sub oG_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

10  On Error GoTo oG_MouseUp_Error

20  If oG.Value = 0 Then
30      If iMsg("Are you sure you want to unmask this sample?", vbYesNo, "NetAcquire") = vbNo Then
40          oG.Value = 1
50          Exit Sub
60      End If
70  End If

80  If oG.Value = 1 And InStr(1, txtBioComment, oG.Caption) = 0 Then txtBioComment = Trim(txtBioComment & " " & oG.Caption)
90  BioChanged = True
100 cmdSaveBio.Enabled = True

110 Exit Sub

oG_MouseUp_Error:

    Dim strES As String
    Dim intEL As Integer

120 intEL = Erl
130 strES = Err.Description
140 LogError "frmEditAll", "oG_MouseUp", intEL, strES


End Sub

Private Sub oH_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

10  On Error GoTo oH_MouseUp_Error

20  If oH.Value = 0 Then
30      If iMsg("Are you sure you want to unmask this sample?", vbYesNo, "NetAcquire") = vbNo Then
40          oH.Value = 1
50          Exit Sub
60      End If
70  End If

80  BioChanged = True
90  cmdSaveBio.Enabled = True
100 If oH.Value = 1 And InStr(1, txtBioComment, oH.Caption) = 0 Then txtBioComment = Trim(txtBioComment & " " & oH.Caption)

110 Exit Sub

oH_MouseUp_Error:

    Dim strES As String
    Dim intEL As Integer

120 intEL = Erl
130 strES = Err.Description
140 LogError "frmEditAll", "oH_MouseUp", intEL, strES


End Sub

Private Sub oJ_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

10  On Error GoTo oJ_MouseUp_Error

20  If oJ.Value = 0 Then
30      If iMsg("Are you sure you want to unmask this sample?", vbYesNo, "NetAcquire") = vbNo Then
40          oJ.Value = 1
50          Exit Sub
60      End If
70  End If

80  BioChanged = True
90  cmdSaveBio.Enabled = True
100 If oJ.Value = 1 And InStr(1, txtBioComment, oJ.Caption) = 0 Then txtBioComment = Trim(txtBioComment & " " & oJ.Caption)

110 Exit Sub

oJ_MouseUp_Error:

    Dim strES As String
    Dim intEL As Integer

120 intEL = Erl
130 strES = Err.Description
140 LogError "frmEditAll", "oJ_MouseUp", intEL, strES


End Sub

Private Sub oL_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

10  On Error GoTo oL_MouseUp_Error

20  If oL.Value = 0 Then
30      If iMsg("Are you sure you want to unmask this sample?", vbYesNo, "NetAcquire") = vbNo Then
40          oL.Value = 1
50          Exit Sub
60      End If
70  End If


80  BioChanged = True
90  cmdSaveBio.Enabled = True
100 If oL.Value = 1 And InStr(1, txtBioComment, oL.Caption) = 0 Then
110     txtBioComment = Trim(txtBioComment & " " & oL.Caption)
120 End If

130 Exit Sub

oL_MouseUp_Error:

    Dim strES As String
    Dim intEL As Integer

140 intEL = Erl
150 strES = Err.Description
160 LogError "frmEditAll", "oL_MouseUp", intEL, strES


End Sub

Private Sub oO_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

10  On Error GoTo oO_MouseUp_Error

20  If oO.Value = 0 Then
30      If iMsg("Are you sure you want to unmask this sample?", vbYesNo, "NetAcquire") = vbNo Then
40          oO.Value = 1
50          Exit Sub
60      End If
70  End If


80  BioChanged = True
90  cmdSaveBio.Enabled = True
100 If oO.Value = 1 And InStr(1, txtBioComment, oO.Caption) = 0 Then txtBioComment = Trim(txtBioComment & " " & oO.Caption)

110 Exit Sub

oO_MouseUp_Error:

    Dim strES As String
    Dim intEL As Integer

120 intEL = Erl
130 strES = Err.Description
140 LogError "frmEditAll", "oO_MouseUp", intEL, strES


End Sub

Private Sub oS_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

10  On Error GoTo oS_MouseUp_Error

20  If oS.Value = 0 Then
30      If iMsg("Are you sure you want to unmask this sample?", vbYesNo, "NetAcquire") = vbNo Then
40          oS.Value = 1
50          Exit Sub
60      End If
70  End If

80  BioChanged = True
90  cmdSaveBio.Enabled = True
100 If oS.Value = 1 And InStr(1, txtBioComment, oS.Caption) = 0 Then txtBioComment = Trim(txtBioComment & " " & oS.Caption)

110 Exit Sub

oS_MouseUp_Error:

    Dim strES As String
    Dim intEL As Integer

120 intEL = Erl
130 strES = Err.Description
140 LogError "frmEditAll", "oS_MouseUp", intEL, strES


End Sub

Public Property Let PrintToPrinter(ByVal strNewValue As String)

10  On Error GoTo PrintToPrinter_Error

20  pPrintToPrinter = strNewValue

30  Exit Property

PrintToPrinter_Error:

    Dim strES As String
    Dim intEL As Integer

40  intEL = Erl
50  strES = Err.Description
60  LogError "frmEditAll", "PrintToPrinter", intEL, strES

End Property

Public Property Get PrintToPrinter() As String

10  On Error GoTo PrintToPrinter_Error

20  PrintToPrinter = pPrintToPrinter

30  Exit Property

PrintToPrinter_Error:

    Dim strES As String
    Dim intEL As Integer

40  intEL = Erl
50  strES = Err.Description
60  LogError "frmEditAll", "PrintToPrinter", intEL, strES

End Property

Private Sub SaveBiochemistry(ByVal Validate As Boolean, Optional ByVal UnVal As Boolean)

    Dim sql As String
    Dim tb As New Recordset

10  On Error GoTo SaveBiochemistry_Error

20  txtSampleID = Format(Val(txtSampleID))
30  If Val(txtSampleID) = 0 Then Exit Sub

40  If Validate Then
50      sql = "UPDATE BioResults " & _
              "SET Valid = 1, " & _
              "Operator = '" & UserCode & "' " & _
              "WHERE SampleID = '" & txtSampleID & "'"
60      Cnxn(0).Execute sql
70  ElseIf UnVal = True Then
80      sql = "UPDATE BioResults " & _
              "SET Valid = 0, " & _
              "HealthLink = 0 " & _
              "WHERE SampleID = '" & txtSampleID & "'"
90      Cnxn(0).Execute sql
100 End If

110 If oH Or oS Or oL Or oO Or oG Or oJ Then
120     sql = "SELECT * from Masks WHERE " & _
              "SampleID = '" & txtSampleID & "'"
130     Set tb = New Recordset
140     RecOpenClient 0, tb, sql
150     If tb.EOF Then tb.AddNew
160     tb!SampleID = txtSampleID
170     If Trim(tb!Rundate) & "" = "" Then
180         tb!Rundate = Format$(dtRunDate, "dd/mmm/yyyy")
190     End If
200     tb!h = oH
210     tb!s = oS
220     tb!l = oL
230     tb!o = oO
240     tb!g = oG
250     tb!J = oJ
260     tb.Update
270 Else
280     sql = "DELETE from Masks WHERE " & _
              "SampleID = '" & txtSampleID & "'"
290     Cnxn(0).Execute sql
300 End If

310 sql = "SELECT * FROM Demographics WHERE " & _
          "SampleID = '" & txtSampleID & "'"

320 Set tb = New Recordset
330 RecOpenClient 0, tb, sql
340 If tb.EOF Then
350     tb.AddNew
360 End If
370 If lRandom = "Fasting Sample" Then
380     tb!Fasting = 1
390 Else
400     tb!Fasting = 0
410 End If
420 tb!Faxed = 0
430 tb!RooH = cRooH(0)
440 If Trim(tb!Rundate) & "" = "" Then tb!Rundate = Format$(dtRunDate, "dd/mmm/yyyy")
450 If IsDate(tSampleTime) Then
460     tb!SampleDate = Format$(dtSampleDate, "dd/mmm/yyyy") & " " & Format$(tSampleTime, "hh:mm")
470 Else
480     tb!SampleDate = Format$(dtSampleDate, "dd/mmm/yyyy")
490 End If
500 tb!SampleID = txtSampleID
510 tb.Update

520 Exit Sub

SaveBiochemistry_Error:

    Dim strES As String
    Dim intEL As Integer

530 intEL = Erl
540 strES = Err.Description
550 LogError "frmEditAll", "SaveBiochemistry", intEL, strES, sql

End Sub

Private Sub SaveBloodGas(ByVal Validate As Boolean, Optional ByVal UnVal As Boolean)

    Dim sql As String
    Dim tb As New Recordset

10  On Error GoTo SaveBloodGas_Error

20  txtSampleID = Format(Val(txtSampleID))
30  If Val(txtSampleID) = 0 Then Exit Sub

40  If Validate Then
50      sql = "UPDATE BgaResults set valid = 1 WHERE " & _
              "sampleid = '" & txtSampleID & "'"
60      Cnxn(0).Execute sql
70  ElseIf UnVal = True Then
80      sql = "UPDATE BgaResults set valid = 0 WHERE " & _
              "sampleid = '" & txtSampleID & "'"
90      Cnxn(0).Execute sql
100 End If
110 If Validate Then
120     sql = "UPDATE BgaResults " & _
              "set operator = '" & UserCode & "' WHERE " & _
              "SampleID = '" & txtSampleID & "' "
130     Cnxn(0).Execute sql
140 End If

150 sql = "SELECT * FROM Demographics WHERE " & _
          "SampleID = '" & txtSampleID & "'"

160 Set tb = New Recordset
170 RecOpenClient 0, tb, sql
180 If tb.EOF Then
190     tb.AddNew
200     tb!ForESR = 0
210 End If
220 If lRandom = "Fasting Sample" Then
230     tb!Fasting = 1
240 Else
250     tb!Fasting = 0
260 End If
270 tb!Faxed = 0
280 tb!RooH = cRooH(0)
290 If Trim(tb!Rundate) & "" = "" Then tb!Rundate = Format$(dtRunDate, "dd/mmm/yyyy")
300 If IsDate(tSampleTime) Then
310     tb!SampleDate = Format$(dtSampleDate, "dd/mmm/yyyy") & " " & Format$(tSampleTime, "hh:mm")
320 Else
330     tb!SampleDate = Format$(dtSampleDate, "dd/mmm/yyyy")
340 End If
350 tb!SampleID = txtSampleID
360 tb.Update

370 Exit Sub

SaveBloodGas_Error:

    Dim strES As String
    Dim intEL As Integer

380 intEL = Erl
390 strES = Err.Description
400 LogError "frmEditAll", "SaveBloodGas", intEL, strES, sql

End Sub

Private Sub SaveCoag(ByVal Validate As Boolean)

          Dim sql As String
          Dim tb As New Recordset
          Dim n As Long
          Dim Code As String
          Dim Unit As String

10        On Error GoTo SaveCoag_Error

20        txtSampleID = Format(Val(txtSampleID))
30        If Val(txtSampleID) = 0 Then Exit Sub

40        If grdCoag.Rows = 2 And grdCoag.TextMatrix(1, 0) = "" And txtCoagComment = "" Then Exit Sub

50        If grdCoag.Rows > 1 And grdCoag.TextMatrix(1, 0) <> "" Then
60            For n = 1 To grdCoag.Rows - 1
70                Code = CoagCodeFor(grdCoag.TextMatrix(n, 0))
                  'If grdCoag.TextMatrix(n, 0) = "INR" Then
                  'Unit = "INR"
                  'Else
                  'Unit = grdCoag.TextMatrix(n, 2)
                  'End If
80                Unit = grdCoag.TextMatrix(n, 2)
90                sql = "SELECT * FROM CoagResults WHERE " & _
                        "SampleID = '" & txtSampleID & "' " & _
                        "AND Code = '" & Trim(Code) & "' " & _
                        "AND Units = '" & Unit & "'"
100               Set tb = New Recordset
110               RecOpenClient 0, tb, sql
120               If tb.EOF And SysOptExp(0) = False Then
130                   sql = "SELECT * FROM CoagResults WHERE " & _
                            "SampleID = '" & txtSampleID & "' " & _
                            "AND Code = '" & Code & "'"
140                   Set tb = New Recordset
150                   RecOpenClient 0, tb, sql
160               End If
170               If tb.EOF Then
180                   tb.AddNew
190                   tb!Rundate = Format$(dtRunDate, "dd/mmm/yyyy")
200                   tb!RunTime = Format$(Now, "dd/mmm/yyyy hh:mm")

210               End If
220               tb!Code = Trim(Code)
230               tb!Result = Trim(grdCoag.TextMatrix(n, 1))
240               tb!SampleID = txtSampleID
250               tb!Units = Unit
260               If Validate Then
270                   tb!Valid = 1
280                   tb!UserName = UserCode
290                   tb!Printed = IIf(grdCoag.TextMatrix(n, 6) = "P", 1, 0)
300               ElseIf Validate = False Then
310                   tb!Valid = 0
320                   tb!HealthLink = 0
330                   tb!UserName = UserCode
340                   tb!Printed = IIf(grdCoag.TextMatrix(n, 6) = "P", 1, 0)
350               Else
360                   tb!Valid = IIf(grdCoag.TextMatrix(n, 5) = "V", 1, 0)
370                   tb!Printed = IIf(grdCoag.TextMatrix(n, 6) = "P", 1, 0)
380               End If
390               tb.Update
400           Next
410           tb.Close
420       End If

430       If Trim(tWarfarin) <> "" Then
440           sql = "SELECT * from HaemResults WHERE " & _
                    "SampleID = '" & txtSampleID & "'"
450           Set tb = New Recordset
460           RecOpenClient 0, tb, sql
470           If tb.EOF Then
480               tb.AddNew
490               tb!SampleID = txtSampleID
500           End If
510           tb!Warfarin = Trim$(tWarfarin)
520           tb.Update
530       End If

540       Exit Sub

SaveCoag_Error:

          Dim strES As String
          Dim intEL As Integer

550       intEL = Erl
560       strES = Err.Description
570       LogError "frmEditAll", "SaveCoag", intEL, strES, sql

End Sub

Private Sub SaveComments()

    Dim Obs As New Observations

10  On Error GoTo SaveComments_Error

20  txtSampleID = Format(Val(txtSampleID))
30  If Val(txtSampleID) = 0 Then Exit Sub



40  Obs.Save txtSampleID, True, _
             "Biochemistry", Trim$(txtBioComment), _
             "Demographic", Trim$(txtDemographicComment), _
             "Haematology", Trim$(txtHaemComment), _
             "Coagulation", Trim$(txtCoagComment), _
             "Immunology", Trim$(txtImmComment(1)), _
             "Endocrinology", Trim$(txtImmComment(0)), _
             "BloodGas", Trim$(txtBGaComment)

50  Exit Sub

SaveComments_Error:

    Dim strES As String
    Dim intEL As Integer

60  intEL = Erl
70  strES = Err.Description
80  LogError "frmEditAll", "SaveComments", intEL, strES

End Sub

Private Sub SaveDemographics()

    Dim sql As String
    Dim Fasting As Integer
    Dim Rundate As String
    Dim SampleDate As String
    Dim RecDate As String
    Dim Category As String
    Dim sex As String
    Dim Ward As String
    Dim Clinician As String
    Dim GP As String
    Dim ClDetails As String
    Dim Urgent As Integer

10  On Error GoTo SaveDemographics_Error

20  txtSampleID = Format(Val(txtSampleID))
30  If Val(txtSampleID) = 0 Then Exit Sub

40  If Trim$(tSampleTime) <> "__:__" Then
50      If Not IsDate(tSampleTime) Then
60          iMsg "Invalid Time", vbExclamation
70          Exit Sub
80      End If
90  End If

100 Rundate = Format$(dtRunDate, "dd/mmm/yyyy")

110 If lRandom = "Fasting Sample" Then
120     Fasting = 1
130 Else
140     Fasting = 0
150 End If

160 If IsDate(tSampleTime) Then
170     SampleDate = Format$(dtSampleDate, "dd/mmm/yyyy") & " " & Format$(tSampleTime, "HH:nn")
180 Else
190     SampleDate = Format$(dtSampleDate, "dd/mmm/yyyy") & " " & Format$("00:00", "HH:nn")
200 End If

210 If IsDate(tRecTime) Then
220     RecDate = Format$(dtRecDate, "dd/mmm/yyyy") & " " & Format$(tRecTime, "HH:nn")
230 Else
240     RecDate = Format$(dtRecDate, "dd/mmm/yyyy") & " " & Format(Now, "HH:nn")
250 End If

260 If cCat(0) = "Default" Then
270     Category = ""
280 Else
290     Category = cCat(0)
300 End If

310 sex = UCase$(Left$(lSex, 1))

320 Ward = StrConv(Left$(cmbWard, 50), vbProperCase)
330 Clinician = Left$(cmbClinician, 50)
340 GP = Left$(cmbGP, 50)
350 ClDetails = Left$(cClDetails, 30)

360 Urgent = 0
370 If SysOptUrgent(0) Then
380     If chkUrgent.Value = 1 Then
390         Urgent = 1
400     End If
410 End If

420 sql = "IF EXISTS(SELECT * FROM Demographics WHERE " & _
        "          SampleID = '" & txtSampleID & "') " & _
        "  UPDATE Demographics " & _
        "  SET Fasting = '" & Fasting & "', " & _
        "  RooH = '" & IIf(cRooH(0), 1, 0) & "', " & _
        "  RunDate = '" & Rundate & "', " & _
        "  SampleDate = '" & SampleDate & "', " & _
        "  RecDate = '" & RecDate & "', " & _
        "  Chart = '" & txtChart & "', " & _
        "  AandE = '" & txtAandE & "', " & _
        "  PatName = '" & AddTicks(txtName) & "', " & _
        "  DoB = " & IIf(IsDate(lDoB), Format$(lDoB, "'dd/mmm/yyyy'"), "Null") & ", " & _
        "  Category = '" & Category & "', " & _
        "  Age = '" & lAge & "', " & _
        "  sex = '" & sex & "', " & _
        "  Addr0 = '" & AddTicks(taddress(0)) & "', " & _
        "  Addr1 = '" & AddTicks(taddress(1)) & "', " & _
        "  Ward = '" & AddTicks(Ward) & "', " & _
        "  Clinician = '" & AddTicks(Clinician) & "', " & _
        "  GP = '" & GP & "', " & _
        "  ClDetails = '" & AddTicks(ClDetails) & "', " & _
        "  Hospital = '" & cmbHospital & "', " & _
        "  UserName = '" & AddTicks(UserName) & "', " & _
        "  Urgent = '" & Urgent & "' " & _
      "  WHERE SampleID = '" & txtSampleID & "' "
430 sql = sql & "ELSE " & _
        "  INSERT INTO Demographics " & _
        "  (SampleID, PatName, Age, Sex, RunDate, DoB, Addr0, Addr1, Ward, Clinician, GP, " & _
        "  SampleDate, ClDetails, Hospital, RooH, FAXed, Fasting, DateTimeDemographics, AandE, " & _
        "  Chart, Category, RecDate, UserName, Urgent, RecordDateTime, Operator ) " & _
        "  VALUES " & _
        "  ('" & txtSampleID & "', " & _
        "   '" & AddTicks(txtName) & "', " & _
        "   '" & lAge & "', " & _
        "   '" & sex & "', " & _
        "   '" & Rundate & "', " & _
        "   " & IIf(IsDate(lDoB), Format$(lDoB, "'dd/mmm/yyyy'"), "Null") & ", " & _
        "   '" & AddTicks(taddress(0)) & "', " & _
        "   '" & AddTicks(taddress(1)) & "', " & _
        "   '" & AddTicks(Ward) & "', " & _
        "   '" & AddTicks(Clinician) & "', " & _
      "   '" & AddTicks(GP) & "', "
440 sql = sql & "   '" & SampleDate & "', " & _
        "   '" & AddTicks(ClDetails) & "', " & _
        "   '" & cmbHospital & "', " & _
        "   '" & IIf(cRooH(0), 1, 0) & "', " & _
        "   '0', " & _
        "   '" & Fasting & "', " & _
        "   getdate(), " & _
        "   '" & txtAandE & "', " & _
        "   '" & txtChart & "', " & _
        "   '" & Category & "', " & _
        "   '" & RecDate & "', " & _
        "   '" & AddTicks(UserName) & "', " & _
        "   '" & Urgent & "', " & _
        "   getdate(), " & _
        "   '" & UserName & "')"
450 Cnxn(0).Execute sql

460 SaveComments
    '
    'sql = "SELECT * FROM Demographics WHERE " & _
     '      "SampleID = '" & txtSampleID & "'"
    '
    'Set tb = New Recordset
    'RecOpenServer 0, tb, sql
    'If tb.EOF Then
    '    tb.AddNew
    '    tb!DateTimeDemographics = Format(Now, "dd/MMM/yyyy HH:nn:ss")
    '    If lRandom = "Fasting Sample" Then
    '        tb!Fasting = 1
    '    Else
    '        tb!Fasting = 0
    '    End If
    '    tb!Faxed = 0
    'End If
    '
    'tb!RooH = cRooH(0)
    '
    'tb!RunDate = Format$(dtRunDate, "dd/mmm/yyyy")
    '
    'If IsDate(tSampleTime) Then
    '    tb!SampleDate = Format$(dtSampleDate, "dd/mmm/yyyy") & " " & Format$(tSampleTime, "HH:nn")
    'Else
    '    tb!SampleDate = Format$(dtSampleDate, "dd/mmm/yyyy") & " " & Format$(Now, "HH:nn")
    'End If
    '
    'If IsDate(tRecTime) Then
    '    tb!RecDate = Format$(dtRecDate, "dd/mmm/yyyy") & " " & Format$(tRecTime, "HH:nn")
    'Else
    '    tb!RecDate = Format$(dtRecDate, "dd/mmm/yyyy") & " " & Format(Now, "HH:nn")
    'End If
    'tb!SampleID = txtSampleID
    'tb!Chart = txtChart
    'tb!AandE = Trim(txtAandE)
    'tb!PatName = AddTicks(txtName)
    'If IsDate(lDoB) Then
    '    tb!DoB = Format$(lDoB, "dd/mmm/yyyy")
    'Else
    '    tb!DoB = Null
    'End If
    'If cCat(0) = "Default" Then tb!Category = "" Else tb!Category = cCat(0)
    'If cCat(1) = "Default" Then tb!Category = "" Else tb!Category = cCat(1)
    'If Len(lAge) > 5 Then tb!Age = lAge
    'tb!Sex = Left$(lSex, 1)
    'tb!Addr0 = tAddress(0)
    'tb!Addr1 = tAddress(1)
    'tb!Ward = StrConv(Left$(cmbWard, 50), vbProperCase)
    'tb!Clinician = AddTicks(Left$(cmbClinician, 50))
    'tb!GP = AddTicks(Left$(cmbGP, 50))
    'tb!ClDetails = Left$(cClDetails, 30)
    'tb!Hospital = cmbHospital
    'tb!UserName = UserName
    'If SysOptUrgent(0) Then
    '    If chkUrgent.Value = 1 Then tb!Urgent = 1 Else tb!Urgent = 0
    'End If
    'tb.Update

470 LogTimeOfPrinting txtSampleID, "D"

480 Exit Sub

SaveDemographics_Error:

    Dim strES As String
    Dim intEL As Integer

490 intEL = Erl
500 strES = Err.Description
510 LogError "frmEditAll", "SaveDemographics", intEL, strES, sql

End Sub

Private Sub SaveEndocrinology(ByVal Validate As Boolean, Optional ByVal UnVal As Boolean)

    Dim sql As String
    Dim tb As New Recordset

10  On Error GoTo SaveEndocrinology_Error

20  txtSampleID = Format(Val(txtSampleID))
30  If Val(txtSampleID) = 0 Then Exit Sub

40  If Validate Then
50      sql = "UPDATE EndResults " & _
              "SET Valid = 1, " & _
              "Operator = '" & UserCode & "' " & _
              "WHERE SampleID = '" & txtSampleID & "'"
60      Cnxn(0).Execute sql
70  ElseIf UnVal = True Then
80      sql = "UPDATE endResults set valid = 0, healthlink = 0 WHERE " & _
              "sampleid = '" & txtSampleID & "'"
90      Cnxn(0).Execute sql
100 End If

110 If Ih(0) Or Iis(0) Or Il(0) Or Io(0) Or Ig(0) Or Ij(0) Then
120     sql = "SELECT * from EndMasks WHERE " & _
              "SampleID = '" & txtSampleID & "'"
130     Set tb = New Recordset
140     RecOpenClient 0, tb, sql
150     If tb.EOF Then tb.AddNew
160     tb!SampleID = txtSampleID
170     tb!Rundate = Format$(dtRunDate, "dd/mmm/yyyy")
180     tb!h = Ih(0)
190     tb!s = Iis(0)
200     tb!l = Il(0)
210     tb!o = Io(0)
220     tb!g = Ig(0)
230     tb!J = Ij(0)
240     tb.Update
250 Else
260     sql = "DELETE from EndMasks WHERE " & _
              "SampleID = '" & txtSampleID & "'"
270     Cnxn(0).Execute sql
280 End If

290 sql = "SELECT * FROM Demographics WHERE " & _
          "SampleID = '" & txtSampleID & "'"

300 Set tb = New Recordset
310 RecOpenClient 0, tb, sql
320 If tb.EOF Then
330     tb.AddNew
340 End If
350 If lImmRan(0) = "Fasting Sample" Then
360     tb!Fasting = 1
370 Else
380     tb!Fasting = 0
390 End If
400 tb!Faxed = 0
410 tb!RooH = cRooH(0)
420 tb!Rundate = Format$(dtRunDate, "dd/mmm/yyyy")
430 If IsDate(tSampleTime) Then
440     tb!SampleDate = Format$(dtSampleDate, "dd/mmm/yyyy") & " " & Format$(tSampleTime, "hh:mm")
450 Else
460     tb!SampleDate = Format$(dtSampleDate, "dd/mmm/yyyy")
470 End If
480 tb!SampleID = txtSampleID
490 tb.Update

500 Exit Sub

SaveEndocrinology_Error:

    Dim strES As String
    Dim intEL As Integer

510 intEL = Erl
520 strES = Err.Description
530 LogError "frmEditAll", "SaveEndocrinology", intEL, strES, sql

End Sub

Private Sub SaveExtern(ByVal Validate As Boolean, Optional ByVal UnVal As Boolean = False)

    Dim tb As New Recordset
    Dim R As Integer
    Dim TestName As String
    Dim sql As String

10  On Error GoTo SaveExtern_Error

20  txtSampleID = Format(Val(txtSampleID))
30  If Val(txtSampleID) = 0 Then Exit Sub

40  If Validate Then
50      sql = "UPDATE ExtResults " & _
              "SET Valid = 1, " & _
              "UserName = '" & UserCode & "' " & _
              "WHERE SampleID = '" & txtSampleID & "'"
60      Cnxn(0).Execute sql
70  ElseIf UnVal = True Then
80      sql = "UPDATE ExtResults " & _
              "SET Valid = 0, " & _
              "HealthLink = 0 " & _
              "WHERE SampleID = '" & txtSampleID & "'"
90      Cnxn(0).Execute sql
100 End If

110 With grdExt
120     For R = 1 To .Rows - 1
130         If Trim(.TextMatrix(R, 0)) <> "" Then
140             TestName = .TextMatrix(R, 0)
150             sql = "If Exists(Select 1 From ExtResults " & _
                      "Where sampleid = @sampleid0 " & _
                      "And Analyte = '@Analyte1' ) " & _
                      "Begin " & _
                      "Update ExtResults Set " & _
                      "sampleid = @sampleid0, " & _
                      "Analyte = '@Analyte1', " & _
                      "result = '@result2', " & _
                      "sendto = '@sendto3', " & _
                      "units = '@units4', " & _
                      "retdate = '@retdate6', " & _
                      "sentdate = '@sentdate7', " & _
                      "sapcode = '@sapcode8', " & _
                      "OrderList = @OrderList10, " & _
                      "Savetime = '@Savetime11', " & _
                      "Username = '@Username12' " & _
                      "Where sampleid = @sampleid0 " & _
                      "And Analyte = '@Analyte1'  " & _
                      "End  "
160             sql = sql & "Else " & _
                      "Begin  " & _
                      "Insert Into ExtResults (sampleid, Analyte, result, sendto, units, retdate, sentdate, " & _
                      "sapcode, OrderList, Savetime, Username) Values " & _
                      "('@sampleid0', '@Analyte1', '@result2', '@sendto3', '@units4', '@retdate6', '@sentdate7', " & _
                      "'@sapcode8', @OrderList10, '@Savetime11', '@Username12') " & _
                      "End"

170             sql = Replace(sql, "@sampleid0", txtSampleID)
180             sql = Replace(sql, "@Analyte1", TestName)
190             sql = Replace(sql, "@result2", Trim(.TextMatrix(R, 1)))
200             sql = Replace(sql, "@sendto3", AddTicks(Trim(.TextMatrix(R, 4))))
210             sql = Replace(sql, "@units4", .TextMatrix(R, 3))
220             If IsDate(.TextMatrix(R, 6)) Then
230                 .TextMatrix(R, 6) = Format(.TextMatrix(R, 6), "dd/mmm/yyyy")
240                 sql = Replace(sql, "@retdate6", .TextMatrix(R, 6))
250             Else
260                 sql = Replace(sql, "'@retdate6'", "Null")
270             End If
280             If IsDate(.TextMatrix(R, 5)) Then
290                 .TextMatrix(R, 5) = Format(.TextMatrix(R, 5), "dd/mmm/yyyy hh:mm:ss")
300                 sql = Replace(sql, "@sentdate7", .TextMatrix(R, 5))
310             Else
320                 sql = Replace(sql, "@sentdate7", Format(Now, "dd/mmm/yyyy hh:mm:ss"))
330             End If
340             sql = Replace(sql, "@sapcode8", .TextMatrix(R, 7))
350             sql = Replace(sql, "@OrderList10", R)
360             sql = Replace(sql, "@Savetime11", Format(Now, "dd/MMM/yyyy hh:mm:ss"))
370             sql = Replace(sql, "@Username12", UserName)

380             Cnxn(0).Execute sql
                '      sql = "SELECT * FROM ExtResults WHERE " & _
                       '            "SampleID = '" & txtSampleID & "' " & _
                       '            "AND Analyte = '" & TestName & "'"
                '      Set tb = New Recordset
                '      RecOpenServer 0, tb, sql
                '      If tb.EOF Then
                '        tb.AddNew
                '      End If
                '      tb!SampleID = txtSampleID
                '      tb!Analyte = TestName
                '      tb!Result = Trim(.TextMatrix(R, 1))
                '      tb!Units = .TextMatrix(R, 3)
                '      tb!SendTo = Trim(.TextMatrix(R, 4))
                '      If IsDate(.TextMatrix(R, 5)) Then
                '        .TextMatrix(R, 5) = Format(.TextMatrix(R, 5), "dd/mmm/yyyy")
                '        tb!SentDate = .TextMatrix(R, 5)
                '      Else
                '        tb!SentDate = Format(Now, "dd/mmm/yyyy")
                '      End If
                '      If IsDate(.TextMatrix(R, 6)) Then
                '        .TextMatrix(R, 6) = Format(.TextMatrix(R, 6), "dd/mmm/yyyy")
                '        tb!RetDate = .TextMatrix(R, 6)
                '      Else
                '        tb!RetDate = Null
                '      End If
                '      tb!SapCode = .TextMatrix(R, 7)
                '      tb!orderlist = R
                '      tb!Username = Username
                '      tb!savetime = Format(Now, "dd/MMM/yyyy hh:mm:ss")
                '      tb.Update
390         End If
400     Next
410 End With

    'Created on 06/07/2011 10:19:33
    'Autogenerated by SQL Scripting

420 sql = "If Exists(Select 1 From etc " & _
          "Where sampleid = '@sampleid0' ) " & _
          "Begin " & _
          "Update etc Set " & _
          "etc0 = '@etc01', " & _
          "etc1 = '@etc12', " & _
          "etc2 = '@etc23', " & _
          "etc3 = '@etc34', " & _
          "etc4 = '@etc45', " & _
          "etc5 = '@etc56', " & _
          "etc6 = '@etc67', " & _
          "etc7 = '@etc78', " & _
          "etc8 = '@etc89' " & _
          "Where sampleid = '@sampleid0'  " & _
          "End  " & _
          "Else " & _
          "Begin  " & _
          "Insert Into etc (sampleid, etc0, etc1, etc2, etc3, etc4, etc5, etc6, etc7, etc8) Values " & _
          "('@sampleid0', '@etc01', '@etc12', '@etc23', '@etc34', '@etc45', '@etc56', '@etc67', '@etc78', '@etc89') " & _
          "End"

430 sql = Replace(sql, "@sampleid0", txtSampleID)
440 sql = Replace(sql, "@etc01", txtEtc(0))
450 sql = Replace(sql, "@etc12", txtEtc(1))
460 sql = Replace(sql, "@etc23", txtEtc(2))
470 sql = Replace(sql, "@etc34", txtEtc(3))
480 sql = Replace(sql, "@etc45", txtEtc(4))
490 sql = Replace(sql, "@etc56", txtEtc(5))
500 sql = Replace(sql, "@etc67", txtEtc(6))
510 sql = Replace(sql, "@etc78", txtEtc(7))
520 sql = Replace(sql, "@etc89", txtEtc(8))

530 Cnxn(0).Execute sql

    'sql = "SELECT * from etc WHERE " & _
     '      "sampleid = '" & txtSampleID & "'"
    'Set tb = New Recordset
    'RecOpenServer 0, tb, sql
    'If tb.EOF Then
    '    tb.AddNew
    '    tb!SampleID = txtSampleID
    'End If
    'tb!etc0 = txtEtc(0)
    'tb!etc1 = txtEtc(1)
    'tb!etc2 = txtEtc(2)
    'tb!etc3 = txtEtc(3)
    'tb!etc4 = txtEtc(4)
    'tb!etc5 = txtEtc(5)
    'tb!etc6 = txtEtc(6)
    'tb!etc7 = txtEtc(7)
    'tb!etc8 = txtEtc(8)
    'tb.Update

540 Exit Sub

SaveExtern_Error:

    Dim strES As String
    Dim intEL As Integer

550 intEL = Erl
560 strES = Err.Description
570 LogError "frmEditAll", "SaveExtern", intEL, strES, sql

End Sub

Private Sub SaveHaematology(ByVal Validate As Boolean)

    Dim tb As New Recordset
    Dim sql As String

10  On Error GoTo SaveHaematology_Error

20  txtSampleID = Format(Val(txtSampleID))
30  If Val(txtSampleID) = 0 Then Exit Sub

40  If Trim$(txtSampleID) = "" Then
50      iMsg "Must have Sample ID Number.", vbCritical
60      Exit Sub
70  End If

    'Created on 01/02/2011 17:33:41
    'Autogenerated by SQL Scripting

    '80    sql = "SELECT * FROM HaemResults WHERE SampleID = '" & txtSampleID & "'"
    '90    Set tb = New Recordset
    '100   RecOpenServer 0, tb, sql
    '110   If Not tb.EOF Then
80  ArchiveHaem txtSampleID
    '130   End If


90  sql = "If Exists(Select 1 From HaemResults " & _
          "Where sampleid = @sampleid120 ) " & _
          "Begin " & _
          "Update HaemResults Set " & _
          "basa = '@basa6', " & _
          "basp = '@basp7', " & _
          "casot = @casot10, " & _
          "cbad = @cbad11, " & _
          "ccoag = @ccoag12, " & _
          "cesr = @cesr20, " & _
          "cFilm = @cFilm21, " & _
          "ch = '@ch22', " & _
          "cmalaria = @cmalaria23, " & _
          "cmonospot = @cmonospot24, " & _
          "cra = @cra25, " & _
          "cretics = @cretics26, " & _
          "csickledex = @csickledex27, " & _
          "eosa = '@eosa30', " & _
          "eosp = '@eosp31', " & _
          "ESR = '@ESR34', " & _
          "hct = '@hct49', " & _
          "hdw = '@hdw50', " & _
          "HealthLink = @HealthLink52, " & _
    "hgb = '@hgb53', "
100 sql = sql & _
          "ho = '@ho54', " & _
          "li = '@li61', " & _
          "luca = '@luca66', " & _
          "lucp = '@lucp67', " & _
          "lyma = '@lyma68', " & _
          "lymp = '@lymp69', " & _
          "mch = '@mch71', " & _
          "mchc = '@mchc72', " & _
          "mcv = '@mcv73', " & _
          "monoa = '@monoa82', " & _
          "monop = '@monop83', " & _
          "monospot = '@monospot84', " & _
          "mpv = '@mpv86', " & _
          "mpxi = '@mpxi87', " & _
          "neuta = '@neuta89', " & _
          "neutP = '@neutP90', " & _
          "nrbcP = '@nrbcP94', " & _
          "Operator = '@Operator95', " & _
      "plt = '@plt100', "
110 sql = sql & _
          "rbc = '@rbc109', " & _
          "rdwcv = '@rdwcv112', " & _
          "RetA = '@RetA114', " & _
          "RetP = '@RetP116', " & _
          "tasot = '@tasot122', " & _
          "tra = '@tra123', " & _
          "valid = @valid131, " & _
          "Warfarin = '@Warfarin133', " & _
          "Malaria = '@Malaria136', " & _
          "Sickledex = '@Sickledex137', " & _
          "wbc = '@wbc135' " & _
          "Where sampleid = @sampleid120  " & _
          "End  " & _
          "Else " & _
            "Begin  "
120 sql = sql & _
          "Insert Into HaemResults (basa, basp, casot, cbad, ccoag, cesr, cFilm, ch, cmalaria, cmonospot, cra, " & _
          "cretics, csickledex, eosa, eosp, ESR, FAXed, hct, hdw, HealthLink, hgb, ho, li, luca, lucp, lyma, lymp, " & _
          "mch, mchc, mcv, monoa, monop, monospot, mpv, mpxi, neuta, neutP, nrbcP, Operator, plt, printed, rbc, " & _
          "rdwcv, RetA, RetP, RunDate, RunDateTime, sampleid, tasot, tra, valid, Warfarin, wbc, Malaria, Sickledex) Values " & _
          "('@basa6', '@basp7', @casot10, @cbad11, @ccoag12, @cesr20, @cFilm21, '@ch22', @cmalaria23, @cmonospot24, " & _
          "@cra25, @cretics26, @csickledex27, '@eosa30', '@eosp31', '@ESR34', @FAXed35, '@hct49', '@hdw50', @HealthLink52, " & _
          "'@hgb53', '@ho54', '@li61', '@luca66', '@lucp67', '@lyma68', '@lymp69', '@mch71', '@mchc72', '@mcv73', " & _
          "'@monoa82', '@monop83', '@monospot84', '@mpv86', '@mpxi87', '@neuta89', '@neutP90', '@nrbcP94', '@Operator95', " & _
          "'@plt100', @printed105, '@rbc109', '@rdwcv112', '@RetA114', '@RetP116', '@RunDate118', " & _
          "'@RunDateTime119', @sampleid120, '@tasot122', '@tra123', @valid131, '@Warfarin133', '@wbc135', '@Malaria136', '@Sickledex137' ) " & _
          "End"

130 sql = Replace(sql, "@basa6", grdH.TextMatrix(5, 0))
140 sql = Replace(sql, "@basp7", grdH.TextMatrix(5, 3))
150 sql = Replace(sql, "@casot10", IIf(cASot = 1, 1, 0))
    '210   If SysOptBadRes(0) Then
160 sql = Replace(sql, "@cbad11", IIf(chkBad = 1, 1, 0))
    '230   Else
    '240       sql = Replace(sql, "@cbad11", "Null")
    '250   End If
170 sql = Replace(sql, "@ccoag12", 0)
180 sql = Replace(sql, "@cesr20", IIf(cESR = 1, 1, 0))
190 sql = Replace(sql, "@cFilm21", IIf(cFilm = 1, 1, 0))
200 sql = Replace(sql, "@ch22", gRbc.TextMatrix(8, 1))
210 sql = Replace(sql, "@cmalaria23", IIf(chkMalaria = 1, 1, 0))
220 sql = Replace(sql, "@cmonospot24", IIf(cMonospot = 1, 1, 0))
230 sql = Replace(sql, "@cra25", IIf(cRA = 1, 1, 0))
240 sql = Replace(sql, "@cretics26", IIf(cRetics = 1, 1, 0))
250 sql = Replace(sql, "@csickledex27", IIf(chkSickledex = 1, 1, 0))
260 sql = Replace(sql, "@eosa30", grdH.TextMatrix(4, 0))
270 sql = Replace(sql, "@eosp31", grdH.TextMatrix(4, 3))
280 sql = Replace(sql, "@ESR34", tESR)
290 sql = Replace(sql, "@FAXed35", 0)
300 sql = Replace(sql, "@hct49", Left$(gRbc.TextMatrix(3, 1), 5))
310 sql = Replace(sql, "@hdw50", Left$(gRbc.TextMatrix(5, 1), 5))
320 sql = Replace(sql, "@HealthLink52", 0)
330 sql = Replace(sql, "@hgb53", gRbc.TextMatrix(2, 1))
340 sql = Replace(sql, "@ho54", gRbc.TextMatrix(11, 1))
350 sql = Replace(sql, "@li61", txtLI)
360 sql = Replace(sql, "@luca66", grdH.TextMatrix(6, 0))
370 sql = Replace(sql, "@lucp67", grdH.TextMatrix(6, 3))
380 sql = Replace(sql, "@lyma68", grdH.TextMatrix(2, 0))
390 sql = Replace(sql, "@lymp69", grdH.TextMatrix(2, 3))
400 sql = Replace(sql, "@mch71", Left$(gRbc.TextMatrix(6, 1), 5))
410 sql = Replace(sql, "@mchc72", Left$(gRbc.TextMatrix(7, 1), 5))
420 sql = Replace(sql, "@mcv73", Left$(gRbc.TextMatrix(4, 1), 5))
430 sql = Replace(sql, "@monoa82", grdH.TextMatrix(3, 0))
440 sql = Replace(sql, "@monop83", grdH.TextMatrix(3, 3))
450 sql = Replace(sql, "@monospot84", Left$(tMonospot, 1))
460 sql = Replace(sql, "@mpv86", tMPV)
470 sql = Replace(sql, "@mpxi87", txtMPXI)
480 sql = Replace(sql, "@neuta89", grdH.TextMatrix(1, 0))
490 sql = Replace(sql, "@neutP90", grdH.TextMatrix(1, 3))
500 sql = Replace(sql, "@nrbcP94", gRbc.TextMatrix(10, 1))
510 sql = Replace(sql, "@Operator95", UserCode)
520 sql = Replace(sql, "@plt100", tPlt)
530 sql = Replace(sql, "@printed105", 0)
540 sql = Replace(sql, "@rbc109", gRbc.TextMatrix(1, 1))
550 sql = Replace(sql, "@rdwcv112", gRbc.TextMatrix(9, 1))
560 sql = Replace(sql, "@RetA114", Format(tRetA, "###.0"))
570 sql = Replace(sql, "@RetP116", Trim(tRetP))
580 sql = Replace(sql, "@RunDate118", Format$(dtRunDate, "dd/mmm/yyyy"))
590 sql = Replace(sql, "@RunDateTime119", Format$(Now, "dd/mmm/yyyy hh:mm"))
600 sql = Replace(sql, "@sampleid120", txtSampleID)
610 sql = Replace(sql, "@tasot122", tASOt)
620 sql = Replace(sql, "@tra123", tRa)
630 sql = Replace(sql, "@valid131", IIf(Validate, 1, 0))
640 sql = Replace(sql, "@Warfarin133", tWarfarin)
650 sql = Replace(sql, "@wbc135", tWBC)
660 sql = Replace(sql, "@Malaria136", lblMalaria)
670 sql = Replace(sql, "@Sickledex137", lblSickledex)

680 Cnxn(0).Execute sql


    'sql = "SELECT * from HaemResults WHERE " & _
     '      "SampleID = '" & txtSampleId & "'"
    'Set tb = New Recordset
    'RecOpenServer 0, tb, sql
    'If tb.EOF Then
    '    tb.AddNew
    '    tb!Rundate = Format$(dtRunDate, "dd/mmm/yyyy")
    '    tb!RunDateTime = Format$(Now, "dd/mmm/yyyy hh:mm")
    '    tb!SampleID = txtSampleId
    '    tb!Faxed = 0
    '    tb!Printed = 0
    'Else
    '    Archive 0, tb, "archaemresults"
    'End If
    '
    'tb!rbc = gRBC.TextMatrix(1, 1)
    'tb!Hgb = gRBC.TextMatrix(2, 1)
    'tb!Hct = Left$(gRBC.TextMatrix(3, 1), 5)
    'tb!MCV = Left$(gRBC.TextMatrix(4, 1), 5)
    'tb!hdw = Left$(gRBC.TextMatrix(5, 1), 5)
    'tb!mch = Left$(gRBC.TextMatrix(6, 1), 5)
    'tb!mchc = Left$(gRBC.TextMatrix(7, 1), 5)
    'tb!cH = gRBC.TextMatrix(8, 1)
    'tb!RDWCV = gRBC.TextMatrix(9, 1)
    'tb!nrbcp = gRBC.TextMatrix(10, 1)
    'tb!ho = gRBC.TextMatrix(11, 1)
    'tb!Plt = tPlt
    'tb!mpv = tMPV
    'tb!wbc = tWBC
    'tb!LymA = grdH.TextMatrix(2, 0)
    'tb!LymP = grdH.TextMatrix(2, 3)
    'tb!MonoA = grdH.TextMatrix(3, 0)
    'tb!MonoP = grdH.TextMatrix(3, 3)
    'tb!NeutA = grdH.TextMatrix(1, 0)
    'tb!NeutP = grdH.TextMatrix(1, 3)
    'tb!EosA = grdH.TextMatrix(4, 0)
    'tb!EosP = grdH.TextMatrix(4, 3)
    'tb!BasA = grdH.TextMatrix(5, 0)
    'tb!BasP = grdH.TextMatrix(5, 3)
    'tb!luca = grdH.TextMatrix(6, 0)
    'tb!lucp = grdH.TextMatrix(6, 3)
    'If txtLI = "" Then tb!Li = ""
    'If txtMPXI = "" Then tb!mpxi = ""
    'If lWOC = "" Then tb!wp = ""
    'If lWIC = "" Then tb!wb = ""
    'If tWBC = "" Then tb!pdw = ""
    'If tWBC = "" Then tb!rdwsd = ""
    'tb!esr = tESR
    'tb!reta = Format(tRetA, "###.0")
    'tb!retp = Trim(tRetP)
    'tb!Monospot = Left$(tMonospot, 1)
    'tb!tASOt = tASOt
    'tb!tRa = tRa
    'tb!cESR = cESR = 1
    'tb!cRetics = cRetics = 1
    'tb!cMonospot = cMonospot = 1
    'tb!cRA = cRA = 1
    'tb!cASot = cASot = 1
    'tb!cMalaria = chkMalaria = 1
    'tb!csickledex = chkSickledex = 1
    'tb!Malaria = lblMalaria
    'tb!Sickledex = lblSickledex
    'If SysOptBadRes(0) Then
    '    tb!cbad = chkBad = 1
    'End If
    'tb!ccoag = 0
    'tb!cFilm = cFilm = 1
    'tb!Warfarin = tWarfarin
    'If Validate Then
    '    tb!Valid = 1
    'Else
    '    tb!Valid = 0
    '    tb!HealthLink = 0
    'End If
    'tb!Operator = UserCode
    '
    'tb.Update

    'If lWOC = "" Then tb!wp = ""
    'If lWIC = "" Then tb!wb = ""

690 If lWOC = "" Or lWIC = "" Then
700     sql = "UPDATE HaemResults Set"
710     If lWOC = "" Then sql = sql & " wp = '' ,"
720     If lWOC = "" Then sql = sql & " wb = ''"
730     sql = sql & " WHERE SampleID = '" & txtSampleID & "'"
740     Cnxn(0).Execute sql
750 End If

760 If Trim(txtCondition) <> "" Then
        'Created on 18/02/2011 16:04:33
        'Autogenerated by SQL Scripting

770     sql = "If Exists(Select 1 From HaemCondition " & _
              "Where Chart = '@Chart0' ) " & _
              "Begin " & _
              "Update HaemCondition Set " & _
              "Chart = '@Chart0', " & _
              "Condition = '@Condition1' " & _
              "Where Chart = '@Chart0'  " & _
              "End  " & _
              "Else " & _
              "Begin  " & _
              "Insert Into HaemCondition (Chart, Condition) Values " & _
              "('@Chart0', '@Condition1') " & _
              "End"

780     sql = Replace(sql, "@Chart0", txtChart)
790     sql = Replace(sql, "@Condition1", Trim(txtCondition))

800     Cnxn(0).Execute sql
        '    sql = "SELECT * from HaemCondition WHERE " & _
             '          "chart = '" & txtChart & "'"
        '    Set tb = New Recordset
        '    RecOpenClient 0, tb, sql
        '    If tb.EOF Then tb.AddNew
        '    tb!Chart = txtChart
        '    tb!condition = Trim(txtCondition)
        '    tb.Update
810 End If

820 Set tb = Nothing

830 Exit Sub

SaveHaematology_Error:

    Dim strES As String
    Dim intEL As Integer

840 intEL = Erl
850 strES = Err.Description
860 LogError "frmEditAll", "SaveHaematology", intEL, strES, sql

End Sub

Private Sub SaveImmunology(ByVal Validate As Boolean, Optional ByVal UnVal As Boolean)

    Dim sql As String
    Dim tb As New Recordset

10  On Error GoTo SaveImmunology_Error

20  txtSampleID = Format(Val(txtSampleID))
30  If Val(txtSampleID) = 0 Then Exit Sub

40  If Validate Then
50      sql = "UPDATE ImmResults set valid = 1, operator = '" & UserCode & "' WHERE " & _
              "sampleid = '" & txtSampleID & "' "
60      Cnxn(0).Execute sql
70  ElseIf UnVal = True Then
80      sql = "UPDATE ImmResults set valid = 0, healthlink = 0, operator = '" & UserCode & "' WHERE " & _
              "sampleid = '" & txtSampleID & "'"
90      Cnxn(0).Execute sql
100 End If

110 If Ih(1) Or Iis(1) Or Il(1) Or Io(1) Or Ig(1) Or Ij(1) Then
120     sql = "SELECT * from ImmMasks WHERE " & _
              "SampleID = '" & txtSampleID & "'"
130     Set tb = New Recordset
140     RecOpenClient 0, tb, sql
150     If tb.EOF Then tb.AddNew
160     tb!SampleID = txtSampleID
170     tb!Rundate = Format$(dtRunDate, "dd/mmm/yyyy")
180     tb!h = Ih(1)
190     tb!s = Iis(1)
200     tb!l = Il(1)
210     tb!o = Io(1)
220     tb!g = Ig(1)
230     tb!J = Ij(1)
240     tb.Update
250 Else
260     sql = "DELETE from ImmMasks WHERE " & _
              "SampleID = '" & txtSampleID & "'"
270     Cnxn(0).Execute sql
280 End If

290 sql = "SELECT * FROM Demographics WHERE " & _
          "SampleID = '" & txtSampleID & "'"

300 Set tb = New Recordset
310 RecOpenClient 0, tb, sql
320 If tb.EOF Then
330     tb.AddNew
340 End If
350 If lImmRan(1) = "Fasting Sample" Then
360     tb!Fasting = 1
370 Else
380     tb!Fasting = 0
390 End If
400 tb!Faxed = 0
410 tb!RooH = cRooH(0)
420 tb!Rundate = Format$(dtRunDate, "dd/mmm/yyyy")
430 If IsDate(tSampleTime) Then
440     tb!SampleDate = Format$(dtSampleDate, "dd/mmm/yyyy") & " " & Format$(tSampleTime, "hh:mm")
450 Else
460     tb!SampleDate = Format$(dtSampleDate, "dd/mmm/yyyy")
470 End If
480 tb!SampleID = txtSampleID
490 tb.Update

500 Exit Sub

SaveImmunology_Error:

    Dim strES As String
    Dim intEL As Integer

510 intEL = Erl
520 strES = Err.Description
530 LogError "frmEditAll", "SaveImmunology", intEL, strES, sql

End Sub

Private Sub EnableDemographicEntry(ByVal Enable As Boolean)

10  On Error GoTo EnableDemographicEntry_Error

20  Frame4.Enabled = Enable
30  Frame5.Enabled = Enable
40  fraDate.Enabled = Enable
50  Frame10(0).Enabled = Enable
60  txtChart.Locked = False
70  txtAandE.Locked = Not Enable
80  txtName.Locked = Not Enable
90  txtDoB.Locked = Not Enable
100 txtAge.Locked = Not Enable
110 txtSex.Locked = Not Enable

120 If Not Enable Then
130     StatusBar1.Panels(3).Text = "Demographics Validated"
140     StatusBar1.Panels(3).Bevel = sbrInset
150 Else
160     StatusBar1.Panels(3).Text = "Check Demographics"
170     StatusBar1.Panels(3).Bevel = sbrRaised
180 End If

190 Exit Sub

EnableDemographicEntry_Error:

    Dim strES As String
    Dim intEL As Integer

200 intEL = Erl
210 strES = Err.Description
220 LogError "frmEditAll", "EnableDemographicEntry", intEL, strES

End Sub

Private Sub SetViewHistory()

10  On Error GoTo SetViewHistory_Error

20  Select Case ssTabAll.Tab
    Case 0: bHistory.Visible = False
30  Case 1: bHistory.Visible = HistHaem
40  Case 2: bHistory.Visible = HistBio
50  Case 3: bHistory.Visible = HistCoag
60  Case 4: bHistory.Visible = HistEnd
70  Case 5: bHistory.Visible = HistBga
80  Case 6: bHistory.Visible = HistImm
90  Case 7: bHistory.Visible = HistExt
100 End Select

110 Exit Sub

SetViewHistory_Error:

    Dim strES As String
    Dim intEL As Integer

120 intEL = Erl
130 strES = Err.Description
140 LogError "frmEditAll", "SetViewHistory", intEL, strES

End Sub

Private Sub SetWardClinGP()

    Dim GPAddr As String

10  On Error GoTo SetWardClinGP_Error

20  GPAddr = AddressOfGP(cmbGP)

30  lblAddWardGP = Trim$(taddress(0)) & " " & _
                   Trim$(taddress(1)) & " : " & _
                   cmbWard & " : " & _
                   cmbGP & ":" & _
                   GPAddr & " " & _
                   cmbClinician

40  Exit Sub

SetWardClinGP_Error:

    Dim strES As String
    Dim intEL As Integer

50  intEL = Erl
60  strES = Err.Description
70  LogError "frmEditAll", "SetWardClinGP", intEL, strES

End Sub

Private Sub sstabAll_Click(PreviousTab As Integer)

10    On Error GoTo sstabAll_Click_Error

20    Select Case PreviousTab
          Case 0
30      If cmdSaveDemographics.Enabled Then
40          If iMsg("Demographic Details have changed!" & vbCrLf & "Save?", vbQuestion + vbYesNo) = vbYes Then
50              If Trim$(txtName) <> "" Then
60                  If Trim$(cmbWard) = "" Then
70                      ssTabAll.Tab = 0
80                      iMsg "Must have Ward entry.", vbCritical
90                      Exit Sub
100                 End If

110                 If Trim$(cmbWard) = "GP" Then
120                     If Trim$(cmbGP) = "" Then
130                         ssTabAll.Tab = 0
140                         iMsg "Must have GP entry.", vbCritical
150                         Exit Sub
160                     End If
170                 End If
180             End If
190             cmdSaveDemographics_Click
200         End If
210     End If
220   Case 1
230     If cmdSaveHaem.Enabled Then
240         If iMsg("Haematology Details have changed!" & vbCrLf & "Save?", vbQuestion + vbYesNo) = vbYes Then
250             cmdHSaveH_Click
260         End If
270     End If
280     If cmdSaveComm.Enabled Then
290         If iMsg("Haematology Comments have changed!" & vbCrLf & "Save?", vbQuestion + vbYesNo) = vbYes Then
300             cmdSaveComm_Click
310         End If
320     End If
330   Case 2
340     If cmdSaveBio.Enabled Then
350         If iMsg("Biochemistry Details have changed!" & vbCrLf & "Save?", vbQuestion + vbYesNo) = vbYes Then
360             cmdSaveBio_Click
370         End If
380     End If
390   Case 3
400     If cmdSaveCoag.Enabled Then
410         If iMsg("Coagulation Details have changed!" & vbCrLf & "Save?", vbQuestion + vbYesNo) = vbYes Then
420             cmdSaveCoag_Click
430         End If
440     End If
450   Case 4
460     If cmdSaveImm(0).Enabled Then
470         If iMsg("Endocrinology Details have changed!" & vbCrLf & "Save?", vbQuestion + vbYesNo) = vbYes Then
480             cmdSaveImm_Click (0)
490         End If
500     End If
510   Case 5
520     If cmdSaveBGa.Enabled Then
530         If iMsg("Blood Gas Details have changed!" & vbCrLf & "Save?", vbQuestion + vbYesNo) = vbYes Then
540             cmdSaveBGa_Click
550         End If
560     End If
570   Case 6
580     If cmdSaveImm(1).Enabled Then
590         If iMsg("Immunology Details have changed!" & vbCrLf & "Save?", vbQuestion + vbYesNo) = vbYes Then
600             cmdSaveImm_Click (1)
610         End If
620     End If
630   Case 7
640     If cmdSaveImm(2).Enabled Then
650         If iMsg("External Details have changed!" & vbCrLf & "Save?", vbQuestion + vbYesNo) = vbYes Then
660             cmdSaveImm_Click (2)
670         End If
680     End If
690   End Select

700   cmdPrint.Visible = True
710   cmdPrintHold.Visible = True
720   bFAX.Visible = True
730   cmdSetPrinter.Visible = True
740   Select Case ssTabAll.Tab
          Case 0:    'Demographics
750     cmdPrint.Visible = False
760     If SysOptAllowDemoPrint(0) = False Then cmdPrintHold.Visible = False
770     bFAX.Visible = False
780     cmdSetPrinter.Visible = False
790   Case 1:    'Haematology
800     If Not HaemLoaded Then
810         LoadHaematology
820         HaemLoaded = True
830     ElseIf bValidateHaem.Caption = "VALID" Then
840         lblUrgent.Visible = False
850     ElseIf bValidateHaem.Caption <> "VALID" Then
860         lblUrgent.Visible = UrgentTest
870     End If
880   Case 2:    'Biochemistry

890     If Not BioLoaded Then
900         LoadBiochemistry
910         BioLoaded = True
920     ElseIf bValidateBio.Caption = "VALID" Then
930         lblUrgent.Visible = False
940     ElseIf bValidateBio.Caption <> "VALID" Then
950         lblUrgent.Visible = UrgentTest
960     End If

970   Case 3:    'Coagulation
980     If Not CoagLoaded Then
990         LoadCoagulation
1000        CoagLoaded = True
1010    ElseIf cmdValidateCoag.Caption = "VALID" Then
1020        lblUrgent.Visible = False
1030    ElseIf cmdValidateCoag.Caption <> "VALID" Then
1040        lblUrgent.Visible = UrgentTest
1050    End If

1060  Case 4:    'Endocrinology
1070    If Not EndLoaded Then
1080        LoadEndocrinology
1090        EndLoaded = True
1100    ElseIf bValidateImm(0).Caption = "VALID" Then
1110        lblUrgent.Visible = False
1120    ElseIf bValidateImm(0).Caption <> "VALID" Then
1130        lblUrgent.Visible = UrgentTest
1140    End If

1150  Case 5:    'Biochemistry
1160    If Not BgaLoaded Then
1170        LoadBloodGas
1180        BgaLoaded = True
1190    ElseIf cmdValBG.Caption = "VALID" Then
1200        lblUrgent.Visible = False
1210    ElseIf cmdValBG.Caption <> "VALID" Then
1220        lblUrgent.Visible = UrgentTest
1230    End If

1240  Case 6:    'Immunology
1250    If Not ImmLoaded Then
1260        LoadImmunology
1270        ImmLoaded = True
1280    ElseIf bValidateImm(1).Caption = "VALID" Then
1290        lblUrgent.Visible = False
1300    ElseIf bValidateImm(1).Caption <> "VALID" Then
1310        lblUrgent.Visible = UrgentTest
1320    End If

1330  Case 7:
1340    bFAX.Visible = False
1350    If Not ExtLoaded Then
1360        LoadExt
1370        ExtLoaded = True
1380    End If

1390  End Select

1400  SetFormCaption

1410  SetViewHistory

          'SetDefaultSampleType
1420  CheckAuditTrail
1430  EnableBarCodePrinting
1440  CheckLabLinkStatus

1450  Exit Sub

sstabAll_Click_Error:

          Dim strES As String
          Dim intEL As Integer

1460  intEL = Erl
1470  strES = Err.Description
1480  LogError "frmEditAll", "sstabAll_Click", intEL, strES

End Sub

Private Sub sstabAll_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

10  On Error GoTo sstabAll_MouseMove_Error

20  If ssTabAll.Tab = 1 And bValidateHaem.Caption = "VALID" Then
30      ssTabAll.ToolTipText = "Unvalidate to change"
40  ElseIf ssTabAll.Tab = 2 And bValidateBio.Caption = "VALID" Then
50      ssTabAll.ToolTipText = "Unvalidate to change"

60  Else
70      ssTabAll.ToolTipText = ""
80  End If

90  pBar = 0

100 Exit Sub

sstabAll_MouseMove_Error:

    Dim strES As String
    Dim intEL As Integer

110 intEL = Erl
120 strES = Err.Description
130 LogError "frmEditAll", "sstabAll_MouseMove", intEL, strES


End Sub



Private Sub OrderExternal()



30      With frmAddToTests
40          .sex = txtSex
50          .SampleID = txtSampleID
60          .ClinDetails = cClDetails
70          .SampleDate = dtSampleDate
80          .SampleTime = tSampleTime
90          .Department = "General"
100         .Ward = cmbWard
110         .Clinician = cmbClinician
120         .GP = cmbGP
130         .Show 1
140     End With

150     LoadExt

'
'Dim frm As New frmAddToTests
'
'If Val(txtSampleID) = 0 Then
'    Exit Sub
'End If
'
'If txtName = "" And txtDoB = "" Then
'    iMsg "Please provide Surname and DoB first", vbInformation
'    Exit Sub
'End If
'
'SaveDemographics
'frm.SampleID = Format$(Val(txtSampleID))
'frm.sex = txtSex
'If IsDate(tSampleTime) Then
'    frm.SampleDateTime = Format$(dtSampleDate, "dd/MMM/yyyy") & " " & tSampleTime
'Else
'    frm.SampleDateTime = Format$(dtSampleDate, "dd/MMM/yyyy") & " " & "00:01"
'End If
'frm.ClinicalDetails = cClDetails
'frm.Show 1
'
'Unload frm
'Set frm = Nothing
End Sub

Private Sub StatusBar1_PanelClick(ByVal Panel As ComctlLib.Panel)
On Error GoTo StatusBar1_PanelClick_Error

If Panel <> "" Then
    If Panel.Index = 3 And Panel.Bevel = sbrRaised Then
        ssTabAll.Tab = 0
    ElseIf Panel.Index = 6 Then
        OrderExternal
    End If
End If

Exit Sub

StatusBar1_PanelClick_Error:

Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "frmEditAll", "StatusBar1_PanelClick", intEL, strES
End Sub

Private Sub taddress_Change(Index As Integer)

10  On Error GoTo taddress_Change_Error

20  SetWardClinGP

30  Exit Sub

taddress_Change_Error:

    Dim strES As String
    Dim intEL As Integer

40  intEL = Erl
50  strES = Err.Description
60  LogError "frmEditAll", "taddress_Change", intEL, strES

End Sub

Private Sub taddress_KeyPress(Index As Integer, KeyAscii As Integer)

10  On Error GoTo taddress_KeyPress_Error

20  KeyAscii = VI(KeyAscii, AlphaNumeric_NoApos)

30  cmdSaveDemographics.Enabled = True
40  cmdSaveInc.Enabled = True

50  Exit Sub

taddress_KeyPress_Error:

    Dim strES As String
    Dim intEL As Integer

60  intEL = Erl
70  strES = Err.Description
80  LogError "frmEditAll", "taddress_KeyPress", intEL, strES

End Sub

Private Sub taddress_LostFocus(Index As Integer)

10  On Error GoTo taddress_LostFocus_Error

20  taddress(Index) = StrConv(taddress(Index), vbProperCase)

30  Exit Sub

taddress_LostFocus_Error:

    Dim strES As String
    Dim intEL As Integer

40  intEL = Erl
50  strES = Err.Description
60  LogError "frmEditAll", "taddress_LostFocus", intEL, strES


End Sub

Private Sub tASOt_Change()

10  On Error GoTo tASOt_Change_Error

20  If Trim$(tASOt) <> "" Then
30      cASot = 1
40  Else
50      cASot = 0
60  End If

70  Exit Sub

tASOt_Change_Error:

    Dim strES As String
    Dim intEL As Integer

80  intEL = Erl
90  strES = Err.Description
100 LogError "frmEditAll", "tASOt_Change", intEL, strES

End Sub

Private Sub tasot_Click()

10  On Error GoTo tasot_Click_Error

20  cmdSaveHaem.Enabled = True: cmdHSaveH.Enabled = True

30  If Trim$(tASOt) = "" Or tASOt = "?" Then
40      tASOt = "Negative"
50  ElseIf tASOt = "Negative" Then
60      tASOt = "Positive"
70  Else
80      tASOt = ""
90  End If

100 Exit Sub

tasot_Click_Error:

    Dim strES As String
    Dim intEL As Integer

110 intEL = Erl
120 strES = Err.Description
130 LogError "frmEditAll", "tasot_Click", intEL, strES

End Sub

Private Sub tasot_KeyPress(KeyAscii As Integer)

10  On Error GoTo tasot_KeyPress_Error

20  cmdSaveHaem.Enabled = True: cmdHSaveH.Enabled = True

30  If Trim$(tASOt) = "" Then
40      tASOt = "Negative"
50  ElseIf tASOt = "Negative" Then
60      tASOt = "Positive"
70  Else
80      tASOt = ""
90  End If

100 Exit Sub

tasot_KeyPress_Error:

    Dim strES As String
    Dim intEL As Integer

110 intEL = Erl
120 strES = Err.Description
130 LogError "frmEditAll", "tasot_KeyPress", intEL, strES

End Sub

Private Sub tESR_Change()

10  On Error GoTo tESR_Change_Error

20  If Trim$(tESR) <> "" Then
30      cESR = 1
40  Else
50      cESR = 0
60  End If

70  cmdSaveHaem.Enabled = True: cmdHSaveH.Enabled = True: cmdSaveComm.Enabled = True

80  Exit Sub

tESR_Change_Error:

    Dim strES As String
    Dim intEL As Integer

90  intEL = Erl
100 strES = Err.Description
110 LogError "frmEditAll", "tESR_Change", intEL, strES

End Sub

Private Sub tESR_KeyPress(KeyAscii As Integer)

10  On Error GoTo tESR_KeyPress_Error

20  cmdSaveHaem.Enabled = True: cmdHSaveH.Enabled = True

30  Exit Sub

tESR_KeyPress_Error:

    Dim strES As String
    Dim intEL As Integer

40  intEL = Erl
50  strES = Err.Description
60  LogError "frmEditAll", "tESR_KeyPress", intEL, strES

End Sub

Private Sub tESR_LostFocus()

10  On Error GoTo tESR_LostFocus_Error

20  If tESR = "" Then Exit Sub

30  If tESR <> "?" And tESR <> "*" Then
40      If Not IsNumeric(tESR) Then
50          iMsg "Result must be numeric"
60          tESR = "?"
70          Exit Sub
80      End If
90  End If

100 Exit Sub

tESR_LostFocus_Error:

    Dim strES As String
    Dim intEL As Integer

110 intEL = Erl
120 strES = Err.Description
130 LogError "frmEditAll", "tESR_LostFocus", intEL, strES

End Sub

Private Sub TimerBar_Timer()

10  On Error GoTo TimerBar_Timer_Error

20  pBar = pBar + 1

    'code added 22/08/2005
    'not live
    'If pBar = pBar.max / 2 Then
    '  txtSampleID_LostFocus
    '  Exit Sub
    'End If

30  If pBar = pBar.Max Then
40      Unload Me
50      Exit Sub
60  End If

70  Exit Sub

TimerBar_Timer_Error:

    Dim strES As String
    Dim intEL As Integer

80  intEL = Erl
90  strES = Err.Description
100 LogError "frmEditAll", "TimerBar_Timer", intEL, strES

End Sub

Private Sub tINewValue_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim sql As String
    Dim tb As Recordset
    Dim s As String

10  On Error GoTo tINewValue_KeyDown_Error

20  If KeyCode = 113 And Index = 1 Then
30      sql = "SELECT * from lists WHERE listtype = 'IR' and code = '" & tINewValue(1) & "'"
40      Set tb = New Recordset
50      RecOpenServer 0, tb, sql
60      If Not tb.EOF Then
70          tINewValue(1) = Trim(tb!Text)
80          tINewValue(1).SelStart = Len(tINewValue(1)) + 1
90      End If
100 ElseIf KeyCode = 114 And Index = 1 Then
110     sql = "SELECT * from lists WHERE listtype = 'IR'"
120     Set tb = New Recordset
130     RecOpenServer 0, tb, sql
140     Do While Not tb.EOF
150         s = Trim(tb!Text)
160         frmMessages.lstComm.AddItem s
170         tb.MoveNext
180     Loop

190     Set frmMessages.f = Me
200     Set frmMessages.T = tINewValue(1)
210     frmMessages.Show 1
220     tINewValue(1).SelStart = Len(tINewValue(1)) + 1
230 End If

240 Exit Sub

tINewValue_KeyDown_Error:

    Dim strES As String
    Dim intEL As Integer

250 intEL = Erl
260 strES = Err.Description
270 LogError "frmEditAll", "tINewValue_KeyDown", intEL, strES, sql

End Sub

Private Sub tINewValue_LostFocus(Index As Integer)

10  On Error GoTo tINewValue_LostFocus_Error

20  If Not IsNumeric(tINewValue(Index)) Then
30      tINewValue(Index) = Trim(tINewValue(Index))
40  End If

50  Exit Sub

tINewValue_LostFocus_Error:

    Dim strES As String
    Dim intEL As Integer

60  intEL = Erl
70  strES = Err.Description
80  LogError "frmEditAll", "tINewValue_LostFocus", intEL, strES

End Sub

Private Sub tMonospot_Change()

10  On Error GoTo tMonospot_Change_Error

20  If Trim$(tMonospot) <> "" Then
30      cMonospot = 1
40  Else
50      cMonospot = 0
60  End If

70  Exit Sub

tMonospot_Change_Error:

    Dim strES As String
    Dim intEL As Integer

80  intEL = Erl
90  strES = Err.Description
100 LogError "frmEditAll", "tMonospot_Change", intEL, strES

End Sub

Private Sub tMonospot_Click()

10  On Error GoTo tMonospot_Click_Error

20  cmdSaveHaem.Enabled = True: cmdHSaveH.Enabled = True

30  If Trim$(tMonospot) = "" Or tMonospot = "?" Then
40      tMonospot = "Negative"
50  ElseIf tMonospot = "Negative" Then
60      tMonospot = "Positive"
70  ElseIf tMonospot = "Positive" Then
80      tMonospot = "Inconclusive"
90  Else
100     tMonospot = ""
110 End If

120 Exit Sub

tMonospot_Click_Error:

    Dim strES As String
    Dim intEL As Integer

130 intEL = Erl
140 strES = Err.Description
150 LogError "frmEditAll", "tMonospot_Click", intEL, strES

End Sub

Private Sub tMonospot_KeyPress(KeyAscii As Integer)

10  On Error GoTo tMonospot_KeyPress_Error

20  cmdSaveHaem.Enabled = True: cmdHSaveH.Enabled = True

30  If Trim$(tMonospot) = "" Then
40      tMonospot = "Negative"
50  ElseIf tMonospot = "Negative" Then
60      tMonospot = "Positive"
70  Else
80      tMonospot = ""
90  End If

100 Exit Sub

tMonospot_KeyPress_Error:

    Dim strES As String
    Dim intEL As Integer

110 intEL = Erl
120 strES = Err.Description
130 LogError "frmEditAll", "tMonospot_KeyPress", intEL, strES

End Sub

Private Sub tMPV_KeyPress(KeyAscii As Integer)

10  On Error GoTo tMPV_KeyPress_Error

20  cmdSaveHaem.Enabled = True: cmdHSaveH.Enabled = True

30  Exit Sub

tMPV_KeyPress_Error:

    Dim strES As String
    Dim intEL As Integer

40  intEL = Erl
50  strES = Err.Description
60  LogError "frmEditAll", "tMPV_KeyPress", intEL, strES

End Sub

Private Sub tnewvalue_Click()

10  On Error GoTo tnewvalue_Click_Error

20  If InStr(UCase(cAdd), "PREG") > 0 Then
30      If tnewvalue = "" Then
40          tnewvalue = "Neg"
50      ElseIf tnewvalue = "Neg" Then
60          tnewvalue = "Pos"
70      ElseIf tnewvalue = "Pos" Then
80          tnewvalue = "WKPos"
90      ElseIf tnewvalue = "WKPos" Then
100         tnewvalue = "STPos"
110     ElseIf tnewvalue = "STPos" Then
120         tnewvalue = "Equiv"
130     ElseIf tnewvalue = "Equiv" Then
140         tnewvalue = ""
150     End If
160 End If

170 Exit Sub

tnewvalue_Click_Error:

    Dim strES As String
    Dim intEL As Integer

180 intEL = Erl
190 strES = Err.Description
200 LogError "frmEditAll", "tnewvalue_Click", intEL, strES

End Sub

Private Sub tPlt_KeyPress(KeyAscii As Integer)

10  On Error GoTo tPlt_KeyPress_Error

20  cmdSaveHaem.Enabled = True: cmdHSaveH.Enabled = True

30  Exit Sub

tPlt_KeyPress_Error:

    Dim strES As String
    Dim intEL As Integer

40  intEL = Erl
50  strES = Err.Description
60  LogError "frmEditAll", "tPlt_KeyPress", intEL, strES

End Sub

Private Sub tRa_Change()

10  On Error GoTo tRa_Change_Error

20  If Trim$(tRa) <> "" Then
30      cRA = 1
40  Else
50      cRA = 0
60  End If

70  Exit Sub

tRa_Change_Error:

    Dim strES As String
    Dim intEL As Integer

80  intEL = Erl
90  strES = Err.Description
100 LogError "frmEditAll", "tRa_Change", intEL, strES

End Sub

Private Sub tRa_Click()

10  On Error GoTo tRa_Click_Error

20  cmdSaveHaem.Enabled = True: cmdHSaveH.Enabled = True

30  If Trim$(tRa) = "" Or tRa = "?" Then
40      tRa = "Negative"
50  ElseIf tRa = "Negative" Then
60      tRa = "Positive"
70  Else
80      tRa = ""
90  End If

100 Exit Sub

tRa_Click_Error:

    Dim strES As String
    Dim intEL As Integer

110 intEL = Erl
120 strES = Err.Description
130 LogError "frmEditAll", "tRa_Click", intEL, strES

End Sub

Private Sub tRa_KeyPress(KeyAscii As Integer)

10  On Error GoTo tRa_KeyPress_Error

20  cmdSaveHaem.Enabled = True: cmdHSaveH.Enabled = True

30  If Trim$(tRa) = "" Then
40      tRa = "Negative"
50  ElseIf tRa = "Negative" Then
60      tRa = "Positive"
70  Else
80      tRa = ""
90  End If

100 Exit Sub

tRa_KeyPress_Error:

    Dim strES As String
    Dim intEL As Integer

110 intEL = Erl
120 strES = Err.Description
130 LogError "frmEditAll", "tRa_KeyPress", intEL, strES

End Sub

Private Sub tRecTime_KeyPress(KeyAscii As Integer)

10  cmdSaveDemographics.Enabled = True
20  cmdSaveInc.Enabled = True

End Sub


Private Sub tRecTime_LostFocus()

10  SetDatesColour Me

End Sub


Private Sub tResult_KeyPress(KeyAscii As Integer)

10  On Error GoTo tResult_KeyPress_Error

20  KeyAscii = VI(KeyAscii, Numericfullstopdash)

30  Exit Sub

tResult_KeyPress_Error:

    Dim strES As String
    Dim intEL As Integer

40  intEL = Erl
50  strES = Err.Description
60  LogError "frmEditAll", "tResult_KeyPress", intEL, strES

End Sub

Private Sub tRetA_Change()

10  On Error GoTo tRetA_Change_Error

20  If Trim$(tRetA) <> "" Then
30      cRetics = 1
40  Else
50      cRetics = 0
60  End If

70  Exit Sub

tRetA_Change_Error:

    Dim strES As String
    Dim intEL As Integer

80  intEL = Erl
90  strES = Err.Description
100 LogError "frmEditAll", "tRetA_Change", intEL, strES

End Sub

Private Sub tRetA_KeyPress(KeyAscii As Integer)

10  On Error GoTo tRetA_KeyPress_Error

20  cmdSaveHaem.Enabled = True: cmdHSaveH.Enabled = True

30  Exit Sub

tRetA_KeyPress_Error:

    Dim strES As String
    Dim intEL As Integer

40  intEL = Erl
50  strES = Err.Description
60  LogError "frmEditAll", "tRetA_KeyPress", intEL, strES

End Sub

Private Sub tRetA_LostFocus()

10  On Error GoTo tRetA_LostFocus_Error

20  If tRetA = "" Then Exit Sub

30  If tRetA <> "?" Then
40      If Not IsNumeric(tRetA) Then
50          iMsg "Result must be numeric"
60          tRetA = "?"
70          Exit Sub
80      End If
90  End If

100 Exit Sub

tRetA_LostFocus_Error:

    Dim strES As String
    Dim intEL As Integer



110 intEL = Erl
120 strES = Err.Description
130 LogError "frmEditAll", "tRetA_LostFocus", intEL, strES


End Sub

Private Sub tRetP_Change()

10  On Error GoTo tRetP_Change_Error

20  If Trim$(tRetP) <> "" Then
30      cRetics = 1
40  Else
50      cRetics = 0
60  End If

70  Exit Sub

tRetP_Change_Error:

    Dim strES As String
    Dim intEL As Integer

80  intEL = Erl
90  strES = Err.Description
100 LogError "frmEditAll", "tRetP_Change", intEL, strES

End Sub

Private Sub tRetP_KeyPress(KeyAscii As Integer)

10  On Error GoTo tRetP_KeyPress_Error

20  cmdSaveHaem.Enabled = True: cmdHSaveH.Enabled = True

30  Exit Sub

tRetP_KeyPress_Error:

    Dim strES As String
    Dim intEL As Integer

40  intEL = Erl
50  strES = Err.Description
60  LogError "frmEditAll", "tRetP_KeyPress", intEL, strES

End Sub

Private Sub tRetP_LostFocus()

10  On Error GoTo tRetP_LostFocus_Error

20  If tRetP = "" Then Exit Sub

30  If tRetP <> "?" Then
40      If Not IsNumeric(tRetP) Then
50          iMsg "Result must be numeric"
60          tRetP = "?"
70          Exit Sub
80      End If
90  End If

100 Exit Sub

tRetP_LostFocus_Error:

    Dim strES As String
    Dim intEL As Integer

110 intEL = Erl
120 strES = Err.Description
130 LogError "frmEditAll", "tRetP_LostFocus", intEL, strES

End Sub

Private Sub tSampleTime_KeyPress(KeyAscii As Integer)

10  On Error GoTo tSampleTime_KeyPress_Error

20  cmdSaveDemographics.Enabled = True
30  cmdSaveInc.Enabled = True

40  Exit Sub

tSampleTime_KeyPress_Error:

    Dim strES As String
    Dim intEL As Integer

50  intEL = Erl
60  strES = Err.Description
70  LogError "frmEditAll", "tSampleTime_KeyPress", intEL, strES

End Sub

Private Sub tSampleTime_LostFocus()

10  SetDatesColour Me

End Sub


Private Sub tWBC_KeyPress(KeyAscii As Integer)

10  On Error GoTo tWBC_KeyPress_Error

20  cmdSaveHaem.Enabled = True: cmdHSaveH.Enabled = True

30  Exit Sub

tWBC_KeyPress_Error:

    Dim strES As String
    Dim intEL As Integer

40  intEL = Erl
50  strES = Err.Description
60  LogError "frmEditAll", "tWBC_KeyPress", intEL, strES

End Sub

Private Sub txtAandE_LostFocus()

10  On Error GoTo txtAandE_LostFocus_Error

20  txtAandE = Trim$(UCase$(txtAandE))

30  If UCase(HospName(0)) = "MULLINGAR" Then
40      LoadPatientFromAandE Me, True
50  End If

    'If Trim(txtName) = "" Then
    '    LoadDemo txtAandE
    'End If

60  txtAandE = UCase(txtAandE)

70  cmdSaveDemographics.Enabled = True
80  cmdSaveInc.Enabled = True

90  Exit Sub

txtAandE_LostFocus_Error:

    Dim strES As String
    Dim intEL As Integer

100 intEL = Erl
110 strES = Err.Description
120 LogError "frmEditAll", "txtAandE_Lostfocus", intEL, strES

End Sub

Private Sub txtage_Change()

10  On Error GoTo txtage_Change_Error

20  lAge = txtAge

30  Exit Sub

txtage_Change_Error:

    Dim strES As String
    Dim intEL As Integer

40  intEL = Erl
50  strES = Err.Description
60  LogError "frmEditAll", "txtage_Change", intEL, strES

End Sub

Private Sub txtAge_KeyPress(KeyAscii As Integer)

10  On Error GoTo txtAge_KeyPress_Error

20  If txtAge.Locked Then Exit Sub

30  cmdSaveDemographics.Enabled = True
40  cmdSaveInc.Enabled = True

50  Exit Sub

txtAge_KeyPress_Error:

    Dim strES As String
    Dim intEL As Integer

60  intEL = Erl
70  strES = Err.Description
80  LogError "frmEditAll", "txtAge_KeyPress", intEL, strES

End Sub

Private Sub txtBGaComment_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim s As Variant
    Dim n As Long
    Dim tb As New Recordset
    Dim sql As String

10  On Error GoTo txtBGaComment_KeyDown_Error

20  If KeyCode = 113 Then

30      If Len(txtBGaComment) < 2 Then Exit Sub

40      n = txtBGaComment.SelStart

50      s = UCase(Mid(txtBGaComment, n - 1, 2))

        'For n = 0 To UBound(s)
60      If ListText("BG", s) <> "" Then
70          s = ListText("BG", s)
80      End If
        'Next

90      txtBGaComment = Left(txtBGaComment, n - 2)
100     txtBGaComment = txtBGaComment & s

110     txtBGaComment.SelStart = Len(txtBGaComment)

120 ElseIf KeyCode = 114 Then
130     sql = "SELECT * from lists WHERE listtype = 'BG'"
140     Set tb = New Recordset
150     RecOpenServer 0, tb, sql
160     Do While Not tb.EOF
170         s = Trim(tb!Text)
180         frmMessages.lstComm.AddItem s
190         tb.MoveNext
200     Loop
210     Set frmMessages.f = Me
220     Set frmMessages.T = txtBGaComment
230     frmMessages.Show 1

240 End If

250 cmdSaveBGa.Enabled = True

260 Exit Sub

txtBGaComment_KeyDown_Error:

    Dim strES As String
    Dim intEL As Integer

270 intEL = Erl
280 strES = Err.Description
290 LogError "frmEditAll", "txtBGaComment_KeyDown", intEL, strES

End Sub

Private Sub txtBioComment_Change()

10  On Error GoTo txtBioComment_Change_Error

20  If bValidateBio.Caption = "VALID" Then Exit Sub

30  cmdSaveBio.Enabled = True

40  Exit Sub

txtBioComment_Change_Error:

    Dim strES As String
    Dim intEL As Integer

50  intEL = Erl
60  strES = Err.Description
70  LogError "frmEditAll", "txtBioComment_Change", intEL, strES

End Sub

Private Sub txtBioComment_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sql As String
    Dim tb As New Recordset
    Dim s As Variant
    Dim n As Long
    Dim z As Integer

10  On Error GoTo txtBioComment_KeyDown_Error

20  If bValidateBio.Caption = "VALID" Then Exit Sub

30  If KeyCode = vbKeyF2 Then
40      If Trim(txtBioComment) = "" Then Exit Sub    '
50      n = txtBioComment.SelStart
60      If n < 3 Then Exit Sub
70      z = 1
80      s = Mid(txtBioComment, (n - z), z + 1)
90      z = 2

100     If ListText("BI", s) <> "" Then
110         s = ListText("BI", s)
120     Else
130         s = ""
140     End If

150     If s = "" And Len(txtBioComment) > 2 Then
160         z = 2
170         s = Mid(txtBioComment, (n - z), z + 1)
180         z = 3

190         If ListText("BI", s) <> "" Then
200             s = ListText("BI", s)
210         Else
220             s = ""
230         End If
240     End If

250     If s = "" Then
260         z = 1
270         s = Mid(txtBioComment, n, z + 1)

280         If ListText("BI", s) <> "" Then
290             s = ListText("BI", s)
300         End If
310     End If

320     txtBioComment = Left(txtBioComment, (n - (z)))
330     txtBioComment = txtBioComment & s

340     txtBioComment.SelStart = Len(txtBioComment)

350 ElseIf KeyCode = 114 Then

360     sql = "SELECT * from lists WHERE listtype = 'BI' order by listorder"
370     Set tb = New Recordset
380     RecOpenServer 0, tb, sql
390     Do While Not tb.EOF
400         s = Trim(tb!Text)
410         frmMessages.lstComm.AddItem s
420         tb.MoveNext
430     Loop

440     Set frmMessages.f = Me
450     Set frmMessages.T = txtBioComment
460     frmMessages.Show 1

470 End If

480 cmdSaveBio.Enabled = True

490 Exit Sub

txtBioComment_KeyDown_Error:

    Dim strES As String
    Dim intEL As Integer

500 intEL = Erl
510 strES = Err.Description
520 LogError "frmEditAll", "txtBioComment_KeyDown", intEL, strES, sql

End Sub

Private Sub txtBioComment_KeyPress(KeyAscii As Integer)

10  On Error GoTo txtBioComment_KeyPress_Error

20  If bValidateBio.Caption = "VALID" Then Exit Sub

30  KeyAscii = VI(KeyAscii, AlphaNumeric_NoApos)

40  cmdSaveBio.Enabled = True

50  Exit Sub

txtBioComment_KeyPress_Error:

    Dim strES As String
    Dim intEL As Integer

60  intEL = Erl
70  strES = Err.Description
80  LogError "frmEditAll", "txtBioComment_KeyPress", intEL, strES

End Sub

Private Sub txtchart_Change()

10  On Error GoTo txtchart_Change_Error

20  lChart = txtChart

30  Exit Sub

txtchart_Change_Error:

    Dim strES As String
    Dim intEL As Integer

40  intEL = Erl
50  strES = Err.Description
60  LogError "frmEditAll", "txtchart_Change", intEL, strES

End Sub

Private Sub txtChart_KeyPress(KeyAscii As Integer)

10    On Error GoTo txtChart_KeyPress_Error

20    If txtName <> "" Then KeyAscii = 0

30    'If txtChart.Locked Then Exit Sub



40    cmdSaveDemographics.Enabled = True
50    cmdSaveInc.Enabled = True

60    Exit Sub

txtChart_KeyPress_Error:

      Dim strES As String
      Dim intEL As Integer

70    intEL = Erl
80    strES = Err.Description
90    LogError "frmEditAll", "txtChart_KeyPress", intEL, strES

End Sub



Private Sub txtchart_LostFocus()
10  On Error GoTo txtchart_LostFocus_Error

20  'If txtChart.Locked Then Exit Sub

30  txtChart = Trim$(UCase$(txtChart))

40  If txtChart = "" Then Exit Sub
50  If Trim$(txtName) <> "" Then Exit Sub

60  LoadPatientFromChart Me, True
    'Zyam 8-3-24
    'If txtName <> "" Then
    'LoadDemo txtChart
'    End If
    'Zyam

70  Exit Sub

txtchart_LostFocus_Error:

    Dim strES As String
    Dim intEL As Integer

80  intEL = Erl
90  strES = Err.Description
100 LogError "frmEditAll", "txtchart_LostFocus", intEL, strES
End Sub

Private Sub txtCoagComment_Change()

10  On Error GoTo txtCoagComment_Change_Error

20  If cmdValidateCoag.Caption = "VALID" Then Exit Sub

30  cmdSaveCoag.Enabled = True

40  Exit Sub

txtCoagComment_Change_Error:

    Dim strES As String
    Dim intEL As Integer

50  intEL = Erl
60  strES = Err.Description
70  LogError "frmEditAll", "txtCoagComment_Change", intEL, strES

End Sub

Private Sub txtCoagComment_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sql As String
    Dim tb As New Recordset
    Dim s As Variant
    Dim n As Long
    Dim z As Integer

    'allow 3 chars


10  On Error GoTo txtCoagComment_KeyDown_Error

20  If cmdValidateCoag.Caption = "VALID" Then Exit Sub

30  If KeyCode = 113 Then
40      If Trim(txtCoagComment) = "" Then Exit Sub

50      n = txtCoagComment.SelStart
60      If n = 1 Then Exit Sub

70      z = 2
80      s = Mid(txtCoagComment, (n - z), z + 1)
90      z = 3
100     If ListText("CO", s) <> "" Then
110         s = ListText("CO", s)
120     Else
130         s = ""
140     End If

150     If s = "" Then
160         z = 1
170         s = Mid(txtCoagComment, (n - z), z + 1)
180         z = 2
190         If ListText("CO", s) <> "" Then
200             s = ListText("CO", s)
210         Else
220             s = ""
230         End If
240     End If

250     If s = "" Then
260         z = 1
270         s = Mid(txtCoagComment, n, z + 1)

280         If ListText("CO", s) <> "" Then
290             s = ListText("CO", s)
300         End If
310     End If

320     txtCoagComment = Left(txtCoagComment, (n - (z)))
330     txtCoagComment = txtCoagComment & s

340     txtCoagComment.SelStart = Len(txtCoagComment)

350 ElseIf KeyCode = 114 Then

360     sql = "SELECT * from lists WHERE listtype = 'CO' order by listorder"
370     Set tb = New Recordset
380     RecOpenServer 0, tb, sql
390     Do While Not tb.EOF
400         s = Trim(tb!Text)
410         frmMessages.lstComm.AddItem s
420         tb.MoveNext
430     Loop

440     Set frmMessages.f = Me
450     Set frmMessages.T = txtCoagComment
460     frmMessages.Show 1
470     cmdSaveCoag.Enabled = True

480 End If

490 Exit Sub

txtCoagComment_KeyDown_Error:

    Dim strES As String
    Dim intEL As Integer

500 intEL = Erl
510 strES = Err.Description
520 LogError "frmEditAll", "txtCoagComment_KeyDown", intEL, strES

End Sub

Private Sub txtCoagComment_KeyPress(KeyAscii As Integer)

10  On Error GoTo txtCoagComment_KeyPress_Error

20  If cmdValidateCoag.Caption = "VALID" Then Exit Sub

30  KeyAscii = VI(KeyAscii, AlphaNumeric_NoApos)

40  cmdSaveCoag.Enabled = True

50  Exit Sub

txtCoagComment_KeyPress_Error:

    Dim strES As String
    Dim intEL As Integer

60  intEL = Erl
70  strES = Err.Description
80  LogError "frmEditAll", "txtCoagComment_KeyPress", intEL, strES

End Sub

Private Sub txtCondition_KeyPress(KeyAscii As Integer)

10  On Error GoTo txtCondition_KeyPress_Error

20  KeyAscii = VI(KeyAscii, AlphaNumeric_NoApos)

30  cmdSaveHaem.Enabled = True: cmdHSaveH.Enabled = True

40  Exit Sub

txtCondition_KeyPress_Error:

    Dim strES As String
    Dim intEL As Integer

50  intEL = Erl
60  strES = Err.Description
70  LogError "frmEditAll", "txtCondition_KeyPress", intEL, strES

End Sub

Private Sub txtDemographicComment_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim s As Variant
    Dim n As Long
    Dim sql As String
    Dim tb As New Recordset
    Dim z As Long

10  On Error GoTo txtDemographicComment_KeyDown_Error

20  If KeyCode = 113 Then

30      If txtDemographicComment = "" Then Exit Sub

40      If Len(txtDemographicComment) < 2 Then Exit Sub

50      n = txtDemographicComment.SelStart

60      z = 2
70      s = Mid(txtDemographicComment, (n - z) + 1, z + 1)
80      z = 3
90      If ListText("DE", s) <> "" Then
100         s = ListText("DE", s)
110     Else
120         s = ""
130     End If

140     If s = "" Then
150         z = 1
160         s = Mid(txtDemographicComment, n - z, z + 1)
170         z = 2
180         If ListText("DE", s) <> "" Then
190             s = ListText("DE", s)
200         Else
210             s = ""
220         End If
230     End If

240     If s = "" Then
250         z = 1
260         s = Mid(txtDemographicComment, n, z)

270         If ListText("DE", s) <> "" Then
280             s = ListText("DE", s)
290         End If
300     End If

310     txtDemographicComment = Left(txtDemographicComment, (n - (z)) + 1)
320     txtDemographicComment = txtDemographicComment & s
330     txtDemographicComment.SelStart = Len(txtDemographicComment)

340     cmdSaveDemographics.Enabled = True
350     cmdSaveInc.Enabled = True

360 ElseIf KeyCode = 114 Then

370     sql = "SELECT * from lists WHERE listtype = 'DE' order by listorder"
380     Set tb = New Recordset
390     RecOpenServer 0, tb, sql
400     Do While Not tb.EOF
410         s = Trim(tb!Text)
420         frmMessages.lstComm.AddItem s
430         tb.MoveNext
440     Loop

450     Set frmMessages.f = frmEditAll
460     Set frmMessages.T = txtDemographicComment
470     frmMessages.Show 1

480     cmdSaveDemographics.Enabled = True
490     cmdSaveInc.Enabled = True

500 End If

510 Exit Sub

txtDemographicComment_KeyDown_Error:

    Dim strES As String
    Dim intEL As Integer

520 intEL = Erl
530 strES = Err.Description
540 LogError "frmEditAll", "txtDemographicComment_KeyDown", intEL, strES, sql

End Sub

Private Sub txtDemographicComment_KeyPress(KeyAscii As Integer)

10  On Error GoTo txtDemographicComment_KeyPress_Error

20  KeyAscii = VI(KeyAscii, AlphaNumeric_NoApos)

30  cmdSaveDemographics.Enabled = True
40  cmdSaveInc.Enabled = True

50  Exit Sub

txtDemographicComment_KeyPress_Error:

    Dim strES As String
    Dim intEL As Integer

60  intEL = Erl
70  strES = Err.Description
80  LogError "frmEditAll", "txtDemographicComment_KeyPress", intEL, strES

End Sub

Private Sub txtDemographicComment_LostFocus()

10  On Error GoTo txtDemographicComment_LostFocus_Error

20  txtDemographicComment = initial2upper(txtDemographicComment)
30  lblDemographicComment = txtDemographicComment

40  Exit Sub

txtDemographicComment_LostFocus_Error:

    Dim strES As String
    Dim intEL As Integer

50  intEL = Erl
60  strES = Err.Description
70  LogError "frmEditAll", "txtDemographicComment_LostFocus", intEL, strES

End Sub

Private Sub txtDoB_Change()

10  On Error GoTo txtDoB_Change_Error

20  lDoB = txtDoB

30  Exit Sub

txtDoB_Change_Error:

    Dim strES As String
    Dim intEL As Integer

40  intEL = Erl
50  strES = Err.Description
60  LogError "frmEditAll", "txtDoB_Change", intEL, strES

End Sub

Private Sub txtDoB_KeyPress(KeyAscii As Integer)

10  On Error GoTo txtDoB_KeyPress_Error

20  If txtDoB.Locked Then Exit Sub

30  cmdSaveDemographics.Enabled = True
40  cmdSaveInc.Enabled = True

    'If IsDate(txtDoB) Then
    '    If Format(txtDoB, "yyyymmdd") > Format(Now - 40000, "yyyymmdd") Then
    '        LoadBiochemistry
    '        LoadEndocrinology
    '        LoadImmunology
    '    End If
    'End If

50  Exit Sub

txtDoB_KeyPress_Error:

    Dim strES As String
    Dim intEL As Integer

60  intEL = Erl
70  strES = Err.Description
80  LogError "frmEditAll", "txtDoB_KeyPress", intEL, strES

End Sub

Private Sub txtDoB_LostFocus()

10  If txtDoB.Locked Then Exit Sub

20  txtDoB = Convert62Date(txtDoB, BACKWARD)

30  If Not IsDate(txtDoB) Then
40      txtDoB = ""
50      Exit Sub
60  End If

70  txtAge = CalcAge(txtDoB, dtSampleDate)

80  If txtAge = "" Then
90      txtDoB.BackColor = vbRed
100 Else
110     txtDoB.BackColor = vbButtonFace
120 End If

End Sub

Private Sub txtEsr1_Change()

10  On Error GoTo txtEsr1_Change_Error

20  cmdSaveHaem.Enabled = True: cmdHSaveH.Enabled = True: cmdSaveComm.Enabled = True

30  Exit Sub

txtEsr1_Change_Error:

    Dim strES As String
    Dim intEL As Integer

40  intEL = Erl
50  strES = Err.Description
60  LogError "frmEditAll", "txtEsr1_Change", intEL, strES

End Sub

Private Sub txtEtc_Change(Index As Integer)

10  On Error GoTo txtEtc_Change_Error

20  cmdSaveImm(2).Enabled = True
30  UpDown1.Enabled = False

40  Exit Sub

txtEtc_Change_Error:

    Dim strES As String
    Dim intEL As Integer

50  intEL = Erl
60  strES = Err.Description
70  LogError "frmEditAll", "txtEtc_Change", intEL, strES

End Sub

Private Sub txtEtc_KeyPress(Index As Integer, KeyAscii As Integer)

10  On Error GoTo txtEtc_KeyPress_Error

20  KeyAscii = VI(KeyAscii, AlphaNumeric_NoApos)

30  Exit Sub

txtEtc_KeyPress_Error:

    Dim strES As String
    Dim intEL As Integer

40  intEL = Erl
50  strES = Err.Description
60  LogError "frmEditAll", "txtEtc_KeyPress", intEL, strES


End Sub

Private Sub txtHaemComment_Change()

10  On Error GoTo txtHaemComment_Change_Error

20  If bValidateHaem.Caption = "VALID" Then Exit Sub

30  pBar = 0
40  cmdSaveHaem.Enabled = True
50  cmdHSaveH.Enabled = True
60  cmdSaveComm.Enabled = True

70  Exit Sub

txtHaemComment_Change_Error:

    Dim strES As String
    Dim intEL As Integer

80  intEL = Erl
90  strES = Err.Description
100 LogError "frmEditAll", "txtHaemComment_Change", intEL, strES

End Sub

Private Sub txtHaemComment_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim s As Variant
    Dim n As Long
    Dim z As Long
    Dim tb As New Recordset
    Dim sql As String

    'Haemlock
10  On Error GoTo txtHaemComment_KeyDown_Error

20  If bValidateHaem.Caption = "VALID" Then Exit Sub

30  If KeyCode = vbKeyF2 Then

40      If Trim(txtHaemComment) = "" Then Exit Sub    '
50      If Len(txtHaemComment) < 3 Then Exit Sub
60      n = txtHaemComment.SelStart
70      z = 2
80      s = Mid(txtHaemComment, (n - z), z + 1)
90      z = 3
100     If ListText("HA", s) <> "" Then
110         s = ListText("HA", s)
120     Else
130         s = ""
140     End If

150     If s = "" Then
160         z = 1
170         s = Mid(txtHaemComment, (n - z), z + 1)
180         z = 2
190         If ListText("HA", s) <> "" Then
200             s = ListText("HA", s)
210         Else
220             s = ""
230         End If
240     End If

250     If s = "" Then
260         z = 1
270         s = Mid(txtHaemComment, n, z + 1)
280         If ListText("HA", s) <> "" Then
290             s = ListText("HA", s)
300         End If
310     End If

320     txtHaemComment = Left(txtHaemComment, (n - (z)))
330     txtHaemComment = txtHaemComment & s

340     txtHaemComment.SelStart = Len(txtHaemComment)

350 ElseIf KeyCode = vbKeyF3 Then

360     sql = "SELECT * from lists WHERE listtype = 'HA' order by listorder"
370     Set tb = New Recordset
380     RecOpenServer 0, tb, sql
390     Do While Not tb.EOF
400         s = Trim(tb!Text)
410         frmMessages.lstComm.AddItem s
420         tb.MoveNext
430     Loop

440     Set frmMessages.f = Me
450     Set frmMessages.T = txtHaemComment
460     frmMessages.Show 1
470     txtHaemComment.SetFocus

480 End If

490 Exit Sub

txtHaemComment_KeyDown_Error:

    Dim strES As String
    Dim intEL As Integer

500 intEL = Erl
510 strES = Err.Description
520 LogError "frmEditAll", "txtHaemComment_KeyDown", intEL, strES

End Sub

Private Sub txtHaemComment_KeyPress(KeyAscii As Integer)

'Haemlock
10  On Error GoTo txtHaemComment_KeyPress_Error

20  If bValidateHaem.Caption = "VALID" Then Exit Sub

30  KeyAscii = VI(KeyAscii, AlphaNumeric_NoApos)

40  cmdSaveHaem.Enabled = True: cmdHSaveH.Enabled = True

50  Exit Sub

txtHaemComment_KeyPress_Error:

    Dim strES As String
    Dim intEL As Integer

60  intEL = Erl
70  strES = Err.Description
80  LogError "frmEditAll", "txtHaemComment_KeyPress", intEL, strES

End Sub

Private Sub txtImmComment_Change(Index As Integer)

10  On Error GoTo txtImmComment_Change_Error

20  If bValidateImm(Index).Caption = "VALID" Then Exit Sub

30  If Index = 0 Then
40      cmdSaveImm(0).Enabled = True
50  Else
60      cmdSaveImm(1).Enabled = True
70  End If

80  Exit Sub

txtImmComment_Change_Error:

    Dim strES As String
    Dim intEL As Integer

90  intEL = Erl
100 strES = Err.Description
110 LogError "frmEditAll", "txtImmComment_Change", intEL, strES

End Sub

Private Sub txtImmComment_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    Dim sql As String
    Dim tb As New Recordset
    Dim s As Variant
    Dim n As Long
    Dim z As Long

10  On Error GoTo txtImmComment_KeyDown_Error

20  If bValidateImm(Index).Caption = "VALID" Then Exit Sub

30  If Index = 0 Then

40      If KeyCode = vbKeyF2 Then
50          If Len(Trim(txtImmComment(0))) < 3 Then Exit Sub
60          n = txtImmComment(0).SelStart
70          z = 3
80          s = Mid(txtImmComment(0), (n - z) + 1, z + 1)
90          z = 3
100         If ListText("EN", s) <> "" Then
110             s = ListText("EN", s)
120         Else
130             s = ""
140         End If

150         If s = "" Then
160             z = 1
170             s = Mid(txtImmComment(0), n - z, z + 1)
180             z = 2
190             If ListText("EN", s) <> "" Then
200                 s = ListText("EN", s)
210             Else
220                 s = ""
230             End If
240         End If

250         If s = "" Then
260             z = 1
270             s = Mid(txtImmComment(0), n, z)

280             If ListText("EN", s) <> "" Then
290                 s = ListText("EN", s)
300             End If
310         End If
320         txtImmComment(0) = Left(txtImmComment(0), (n - (z)))
330         txtImmComment(0) = txtImmComment(0) & s

340         txtImmComment(0).SelStart = Len(txtImmComment(0))

350     ElseIf KeyCode = vbKeyF3 Then

360         sql = "SELECT * from lists WHERE listtype = 'EN' order by listorder"
370         Set tb = New Recordset
380         RecOpenServer 0, tb, sql
390         Do While Not tb.EOF
400             s = Trim(tb!Text)
410             frmMessages.lstComm.AddItem s
420             tb.MoveNext
430         Loop

440         Set frmMessages.f = Me
450         Set frmMessages.T = txtImmComment(0)
460         frmMessages.Show 1

470     End If

480     cmdSaveImm(0).Enabled = True
490 Else

500     If KeyCode = vbKeyF2 Then
510         If Len(Trim(txtImmComment(1))) < 2 Then Exit Sub
520         n = txtImmComment(1).SelStart
530         If n < 3 Then Exit Sub
540         s = UCase(Mid(txtImmComment(1), (n - 2), 3))
550         If ListText("IM", s) <> "" Then
560             s = ListText("IM", s)
570         End If
580         txtImmComment(1) = Left(txtImmComment(1), (n) - 3)
590         txtImmComment(1) = txtImmComment(1) & s
600         txtImmComment(1).SelStart = Len(txtImmComment(1))
610     ElseIf KeyCode = vbKeyF3 Then
620         sql = "SELECT * from lists WHERE listtype = 'IM'"
630         Set tb = New Recordset
640         RecOpenServer 0, tb, sql
650         Do While Not tb.EOF
660             s = Trim(tb!Text)
670             frmMessages.lstComm.AddItem s
680             tb.MoveNext
690         Loop
700         Set frmMessages.f = Me
710         Set frmMessages.T = txtImmComment(1)
720         frmMessages.Show 1

730     End If
740     cmdSaveImm(1).Enabled = True
750 End If

760 Exit Sub

txtImmComment_KeyDown_Error:

    Dim strES As String
    Dim intEL As Integer

770 intEL = Erl
780 strES = Err.Description
790 LogError "frmEditAll", "txtImmComment_KeyDown", intEL, strES, sql

End Sub

Private Sub txtImmComment_KeyPress(Index As Integer, KeyAscii As Integer)

10  On Error GoTo txtImmComment_KeyPress_Error

20  KeyAscii = VI(KeyAscii, AlphaNumeric_NoApos)

30  Exit Sub

txtImmComment_KeyPress_Error:

    Dim strES As String
    Dim intEL As Integer

40  intEL = Erl
50  strES = Err.Description
60  LogError "frmEditAll", "txtImmComment_KeyPress", intEL, strES


End Sub

Private Sub txtInput_Change()

10  On Error GoTo txtInput_Change_Error

20  txtInput.SelStart = Len(txtInput)

30  gRbc.TextMatrix(gRbc.RowSel, 1) = Trim(txtInput)

40  If gRbc.TextMatrix(gRbc.RowSel, 1) = "" Then
50      gRbc.TextMatrix(gRbc.RowSel, 2) = ""
60  End If

70  cmdSaveHaem.Enabled = True: cmdHSaveH.Enabled = True

80  Exit Sub

txtInput_Change_Error:

    Dim strES As String
    Dim intEL As Integer

90  intEL = Erl
100 strES = Err.Description
110 LogError "frmEditAll", "txtInput_Change", intEL, strES

End Sub

Private Sub txtName_Change()

10  On Error GoTo txtName_Change_Error

20  lName = txtName

30  Exit Sub

txtName_Change_Error:

    Dim strES As String
    Dim intEL As Integer

40  intEL = Erl
50  strES = Err.Description
60  LogError "frmEditAll", "txtName_Change", intEL, strES

End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)

10  On Error GoTo txtName_KeyPress_Error

20  If txtName.Locked Then Exit Sub


30  KeyAscii = VI(KeyAscii, AlphaNumeric_WithApos)

40  cmdSaveDemographics.Enabled = True
50  cmdSaveInc.Enabled = True

60  Exit Sub

txtName_KeyPress_Error:

    Dim strES As String
    Dim intEL As Integer

70  intEL = Erl
80  strES = Err.Description
90  LogError "frmEditAll", "txtName_KeyPress", intEL, strES

End Sub

Private Sub txtname_LostFocus()

    Dim strName As String
    Dim strSex As String

10  On Error GoTo txtname_LostFocus_Error

20  strName = txtName
30  strSex = txtSex

40  If cmdDemoVal.Caption = "&Validate" Then
50      NameLostFocus strName, strSex
60  End If
70  If strName <> "" Then txtName = strName
80  If strSex <> "" Then txtSex = strSex

90  Exit Sub

txtname_LostFocus_Error:

    Dim strES As String
    Dim intEL As Integer

100 intEL = Erl
110 strES = Err.Description
120 LogError "frmEditAll", "txtname_LostFocus", intEL, strES

End Sub

Private Sub txtSampleID_KeyPress(KeyAscii As Integer)

10  KeyAscii = VI(KeyAscii, Numeric_Only)

End Sub

Public Sub txtSampleID_LostFocus()

10  On Error GoTo txtSampleID_LostFocus_Error

20  If Val(txtSampleID) < 1 Or Trim$(txtSampleID) = "" Or Val(txtSampleID) > (2 ^ 31) - 1 Then
30      txtSampleID = ""
40      txtSampleID.SetFocus
50      Exit Sub
60  End If

70  txtSampleID = Val(txtSampleID)
80  txtSampleID = Int(txtSampleID)

90  LoadAllDetails

100 cmdSaveDemographics.Enabled = False
110 cmdSaveInc.Enabled = False
120 cmdSaveHaem.Enabled = False
130 cmdHSaveH.Enabled = False
140 cmdSaveBio.Enabled = False
150 cmdSaveCoag.Enabled = False
160 cmdSaveImm(0).Enabled = False
170 cmdSaveImm(1).Enabled = False
180 cmdSaveBGa.Enabled = False

190 Exit Sub

txtSampleID_LostFocus_Error:

    Dim strES As String
    Dim intEL As Integer

200 intEL = Erl
210 strES = Err.Description
220 LogError "frmEditAll", "txtSampleID_LostFocus", intEL, strES

End Sub

Private Sub txtSex_Change()

10  On Error GoTo txtSex_Change_Error

20  lSex = txtSex

30  Exit Sub

txtSex_Change_Error:

    Dim strES As String
    Dim intEL As Integer

40  intEL = Erl
50  strES = Err.Description
60  LogError "frmEditAll", "txtSex_Change", intEL, strES

End Sub

Private Sub txtsex_Click()

10  On Error GoTo txtsex_Click_Error

20  If txtSex.Locked Then Exit Sub

30  Select Case Trim$(txtSex)
    Case "": txtSex = "Male"
40  Case "Male": txtSex = "Female"
50  Case "Female": txtSex = ""
60  Case Else: txtSex = ""
70  End Select

80  cmdSaveDemographics.Enabled = True
90  cmdSaveInc.Enabled = True

100 Exit Sub

txtsex_Click_Error:

    Dim strES As String
    Dim intEL As Integer

110 intEL = Erl
120 strES = Err.Description
130 LogError "frmEditAll", "txtsex_Click", intEL, strES

End Sub

Private Sub txtsex_KeyPress(KeyAscii As Integer)

10  On Error GoTo txtsex_KeyPress_Error

20  KeyAscii = 0
30  txtsex_Click

40  LoadBiochemistry
50  LoadEndocrinology
60  LoadImmunology

70  Exit Sub

txtsex_KeyPress_Error:

    Dim strES As String
    Dim intEL As Integer

80  intEL = Erl
90  strES = Err.Description
100 LogError "frmEditAll", "txtsex_KeyPress", intEL, strES

End Sub

Private Sub txtSex_LostFocus()

10  On Error GoTo txtSex_LostFocus_Error

20  If txtSex.Locked = True Then Exit Sub

30  SexLostFocus txtSex, txtName

40  Exit Sub

txtSex_LostFocus_Error:

    Dim strES As String
    Dim intEL As Integer

50  intEL = Erl
60  strES = Err.Description
70  LogError "frmEditAll", "txtSex_LostFocus", intEL, strES

End Sub



Private Sub txtText_LostFocus()
10  txtText.Visible = False
End Sub

Private Sub UpDown1_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

10  On Error GoTo UpDown1_MouseUp_Error

20  pBar = 0

30  UpDown1.Enabled = False

40  If SysOptNumLen(0) > 0 Then
50      If Len(txtSampleID) > SysOptNumLen(0) Then
60          iMsg "Sample Id longer then recommended!"
70      End If
80  End If

90  UpDown1.Enabled = True

100 LoadAllDetails

110 cmdSaveDemographics.Enabled = False
120 cmdSaveInc.Enabled = False
130 cmdSaveHaem.Enabled = False
140 cmdHSaveH.Enabled = False
150 cmdSaveBio.Enabled = False
160 cmdSaveCoag.Enabled = False
170 cmdSaveImm(1).Enabled = False
180 cmdSaveImm(0).Enabled = False
190 cmdSaveBGa.Enabled = False
200 cmdSaveImm(2).Enabled = False

210 Exit Sub

UpDown1_MouseUp_Error:

    Dim strES As String
    Dim intEL As Integer

220 intEL = Erl
230 strES = Err.Description
240 LogError "frmEditAll", "UpDown1_MouseUp", intEL, strES

End Sub

Private Sub VScroll1_Change()

10  On Error GoTo VScroll1_Change_Error

20  pdelta.Top = -VScroll1

30  Exit Sub

VScroll1_Change_Error:

    Dim strES As String
    Dim intEL As Integer

40  intEL = Erl
50  strES = Err.Description
60  LogError "frmEditAll", "VScroll1_Change", intEL, strES

End Sub



Private Sub ValidateTests(Disp As String, g As MSFlexGrid)

    Dim sql As String
    Dim TestCodeList As String
    Dim i As Integer
    Dim TestCode As String

10  On Error GoTo ValidateTests_Error

20  TestCode = ""

30  With g
40      For i = 1 To .Rows - 1
50          .Row = i
60          Select Case Disp
            Case "Imm"
70              .Col = 9
80          Case "Bio"
90              .Col = 10
100         Case "End"
110             .Col = 8
120         Case "Coag"
130             .Col = 8

140         End Select

            'If .CellPicture = imgGreenTick.Picture Then
150         Select Case Disp
            Case "Imm"
160             If .CellPicture = imgGreenTick.Picture Then
170                 TestCode = ICodeForShortName(.TextMatrix(i, 0))
180             End If
190         Case "Bio"
200             TestCode = CodeForShortName(.TextMatrix(i, 0))
210         Case "End"
220             TestCode = eCodeForShortName(.TextMatrix(i, 0))
230         Case "Coag"
240             TestCode = CoagCodeFor(.TextMatrix(i, 0))
250         End Select

260         TestCodeList = TestCodeList & "'" & TestCode & "',"
            'End If
270     Next i
280     If Len(TestCodeList) > 1 Then
290         TestCodeList = Left(TestCodeList, Len(TestCodeList) - 1)
300         If Disp = "Coag" Then
                '            sql = "UPDATE CoagResults " & _
                             '                  "SET Valid = 1, " & _
                             '                  "UserName = '" & UserCode & "' " & _
                             '                  "WHERE Code IN " & _
                             '                  "(" & TestCodeList & ") " & _
                             '                  "AND SampleID = '" & txtSampleID & "'"
310             sql = "UPDATE CoagResults " & _
                      "SET Valid = 1, " & _
                      "UserName = '" & UserCode & "' " & _
                      "WHERE Code IN " & _
                      "(SELECT Code FROM CoagResults WHERE SampleId = '" & txtSampleID & "' " & _
                      "AND Code IN  (" & TestCodeList & ") " & _
                      "AND Valid <> 1) " & _
                      "AND SampleID = '" & txtSampleID & "'"
320             Cnxn(0).Execute sql
                'Update operator is already validated without operator information
330             sql = "UPDATE CoagResults " & _
                      "SET UserName = '" & UserCode & "' " & _
                      "WHERE Code IN (" & TestCodeList & ") " & _
                      "AND COALESCE(UserName, '') = '' " & _
                      "AND SampleID = '" & txtSampleID & "'"
340             Cnxn(0).Execute sql
350         Else
                '            sql = "UPDATE " & Disp & "Results " & _
                             '                  "SET Valid = 1, " & _
                             '                  "Operator = '" & UserCode & "' " & _
                             '                  "WHERE Code IN " & _
                             '                  "(" & TestCodeList & ") " & _
                             '                  "AND SampleID = '" & txtSampleID & "'"
360             sql = "UPDATE " & Disp & "Results " & _
                      "SET Valid = 1, " & _
                      "Operator = '" & UserCode & "' " & _
                      "WHERE Code IN " & _
                      "(SELECT Code FROM  " & Disp & "Results WHERE SampleId = '" & txtSampleID & "' " & _
                      "AND Code IN  (" & TestCodeList & ") " & _
                      "AND Valid <> 1) " & _
                      "AND SampleID = '" & txtSampleID & "'"
370             Cnxn(0).Execute sql
                'Update operator is already validated without operator information
380             sql = "UPDATE " & Disp & "Results " & _
                      "SET Operator = '" & UserCode & "' " & _
                      "WHERE Code IN (" & TestCodeList & ") " & _
                      "AND COALESCE(Operator, '') = '' " & _
                      "AND SampleID = '" & txtSampleID & "'"
390             Cnxn(0).Execute sql
400         End If

410     End If
420 End With

430 Exit Sub

ValidateTests_Error:

    Dim strES As String
    Dim intEL As Integer

440 intEL = Erl
450 strES = Err.Description
460 LogError "frmEditAll", "ValidateTests", intEL, strES, sql

End Sub



Private Function GetSampleType(Disp As String, SampleID As String) As String

    Dim tb As Recordset
    Dim sql As String

10  On Error GoTo GetSampleType_Error

20  sql = "Select Top 1 SampleType From %dispResults Where SampleID = '%sampleid'"
30  sql = Replace(sql, "%disp", Disp)
40  sql = Replace(sql, "%sampleid", SampleID)
50  Set tb = New Recordset
60  RecOpenClient 0, tb, sql
70  If tb.EOF Then
80      GetSampleType = ""
90  Else
100     GetSampleType = tb!SampleType & ""
110 End If

120 Exit Function

GetSampleType_Error:

    Dim strES As String
    Dim intEL As Integer

130 intEL = Erl
140 strES = Err.Description
150 LogError "frmEditAll", "GetSampleType", intEL, strES, sql

End Function
Private Sub SetDefaultSampleType()

    Dim DefaultType As String

10  On Error GoTo SetDefaultSampleType_Error

20  DefaultType = ""

30  Select Case ssTabAll.Tab
    Case 2: DefaultType = "Serum"           'Biochemistry   (Serum)
40      cISampleType(3) = DefaultType
50  Case 4: DefaultType = "Serum"           'Endocrinology  (Serum / Blood for HbA1c)
60      cISampleType(0) = DefaultType
70  Case 5: DefaultType = ""           'Blood Gas      (
80  Case 6: DefaultType = "Serum"           'Immunology     (Serum)
90      cISampleType(1) = DefaultType
100 End Select

110 Exit Sub

SetDefaultSampleType_Error:

    Dim strES As String
    Dim intEL As Integer

120 intEL = Erl
130 strES = Err.Description
140 LogError "frmEditAll", "SetDefaultSampleType", intEL, strES

End Sub

Private Sub CheckLabLinkStatus()

      Dim tb As Recordset
      Dim sql As String
      Dim EnableResendResults As Boolean
      Dim Disp As String

10    On Error GoTo CheckLabLinkStatus_Error

20    If txtSampleID = "" Then Exit Sub

30    Disp = ""
40    EnableResendResults = False
50    Select Case ssTabAll.Tab
          Case 0, 1, 3, 5, 7:
60            EnableResendResults = False
70            Disp = ""
80        Case 2:
90            Disp = "Biochemistry"
100       Case 4:
110           Disp = "Endocrinology"
120       Case 6:
130           Disp = "Immunology"
140   End Select


150   If Disp <> "" Then
160       sql = "Select Count(*) As Cnt From LablinkCommunication Where SampleID = '" & txtSampleID & "' AND Department = '" & Disp & "' AND MessageState = 4"
170       Set tb = New Recordset
180       RecOpenClient 0, tb, sql
190       If tb!Cnt > 0 Then
200           EnableResendResults = True
210       End If
220   End If
230   cmdResend.Visible = EnableResendResults

240   Exit Sub

CheckLabLinkStatus_Error:

      Dim strES As String
      Dim intEL As Integer

250   intEL = Erl
260   strES = Err.Description
270   LogError "frmEditAll", "CheckLabLinkStatus", intEL, strES, sql

End Sub

Private Sub EnableBarCodePrinting()
10    On Error GoTo EnableBarCodePrinting_Error

20    Select Case ssTabAll.Tab
          Case 0, 7
30            cmdPatientNotePad(1).Visible = False
40        Case Else
50            cmdPatientNotePad(1).Visible = True
60    End Select

70    Exit Sub
EnableBarCodePrinting_Error:
         
80    LogError "frmEditAll", "EnableBarCodePrinting", Erl, Err.Description
End Sub

Private Sub CheckAuditTrail()

    Dim tb As Recordset
    Dim sql As String
    Dim EnableAuditTrail As Boolean
    Dim Disp As String

10  On Error GoTo CheckAuditTrail_Error

20  If txtSampleID = "" Then Exit Sub

30  Disp = ""
40  EnableAuditTrail = False
50  Select Case ssTabAll.Tab
    Case 1, 5, 7:
60      EnableAuditTrail = False
70      Disp = ""
80  Case 0:
90      EnableAuditTrail = False
100     Disp = "Demographics"
110 Case 2:
120     Disp = "BioResults"
130 Case 3:
140     Disp = "CoagResults"
150 Case 4:
160     Disp = "EndResults"
170 Case 6:
180     Disp = "ImmResults"


190 End Select


200 If Disp <> "" And Disp <> "Demographics" Then
210     sql = "Select Count(*) As Cnt From " & Disp & "Audit Where SampleID = '" & txtSampleID & "'"
220     Set tb = New Recordset
230     RecOpenClient 0, tb, sql
240     If tb!Cnt > 0 Then
250         EnableAuditTrail = True
260     End If
270 End If
280 cmdAudit.Visible = EnableAuditTrail

290 Exit Sub

CheckAuditTrail_Error:

    Dim strES As String
    Dim intEL As Integer

300 intEL = Erl
310 strES = Err.Description
320 LogError "frmEditAll", "CheckAuditTrail", intEL, strES, sql

End Sub


Private Function DoDeltaCheck(Dept As String, Code As String) As Recordset

    Dim tb As Recordset
    Dim sql As String
    Dim SampleTime As String

10  On Error GoTo DoDeltaCheck_Error

20  If tSampleTime = "__:__" Then
30      SampleTime = "00:00"
40  Else
50      SampleTime = tSampleTime
60  End If

70  sql = "SELECT TOP 1 D.SampleID, D.SampleDate, R.Result " & _
          "FROM Demographics D JOIN " & Dept & "Results R ON " & _
          "R.SampleID = D.SampleID " & _
          "WHERE ("
80  If Trim(txtChart) <> "" Then
90      sql = sql & "D.Chart = '" & AddTicks(txtChart) & "' AND "
100 End If
110 If Trim(txtAandE) <> "" Then
120     sql = sql & "D.AandE = '" & AddTicks(txtAandE) & "' AND "
130 End If
140 sql = sql & " D.PatName = '" & AddTicks(txtName) & "' " & _
          "AND D.DOB = '" & Format(txtDoB, "dd/MMM/yyyy") & "' " & _
          "AND D.SampleDate < '" & Format(dtSampleDate & " " & SampleTime, "dd/MMM/yyyy  hh:mm") & "') " & _
          "AND R.Code = '" & Code & "' " & _
          "ORDER BY D.SampleDate desc"

150 Set tb = New Recordset
160 RecOpenClient 0, tb, sql
170 Set DoDeltaCheck = tb

180 Exit Function

DoDeltaCheck_Error:

    Dim strES As String
    Dim intEL As Integer

190 intEL = Erl
200 strES = Err.Description
210 LogError "frmEditAll", "DoDeltaCheck", intEL, strES, sql

End Function

Private Sub DoDeltaCheckBio()

    Dim tb As Recordset
    Dim sql As String
    Dim SampleTime As String
    Dim OldValue As Single
    Dim NewValue As Single
    Dim DeltaLimit As Single
    Dim Res As String

10  On Error GoTo DoDeltaCheckBio_Error

20  If tSampleTime = "__:__" Then
30      SampleTime = "00:00"
40  Else
50      SampleTime = tSampleTime
60  End If

70  sql = "SELECT T.ShortName, D.SampleID, D.SampleDate, R.Result OldV, X.Result NewV, MAX(T.DeltaLimit) DV " & _
          "FROM Demographics D JOIN BioResults R " & _
          "ON R.SampleID = D.SampleID " & _
          "LEFT OUTER JOIN BioResults X ON X.SampleID = '" & txtSampleID & "' " & _
          "AND X.Code = R.Code " & _
          "JOIN BioTestDefinitions T ON T.Code = X.Code " & _
          "WHERE D.SampleID = (SELECT TOP 1 D.SampleID FROM Demographics D JOIN BioResults R " & _
        "                    ON R.SampleID = D.SampleID WHERE "
80  If Trim(txtChart) <> "" Then
90      sql = sql & "D.Chart = '" & AddTicks(txtChart) & "' AND "
100 End If
110 If Trim(txtAandE) <> "" Then
120     sql = sql & "D.AandE = '" & AddTicks(txtAandE) & "' AND "
130 End If
140 sql = sql & "SampleDate < '" & Format(dtSampleDate & " " & SampleTime, "dd/MMM/yyyy  hh:mm") & "' " & _
        "      ORDER BY SampleDate desc) " & _
          "AND T.DoDelta = 1 " & _
          "AND ABS(DATEDIFF(day, '" & Format$(dtSampleDate, "dd/MMM/yyyy") & "', D.SampleDate )) <= COALESCE(T.CheckTime, 43830) " & _
          "GROUP BY D.SampleID, D.SampleDate, R.Code, R.Result, X.Result, T.ShortName, T.CheckTime "

150 Set tb = New Recordset
160 RecOpenClient 0, tb, sql
170 Do While Not tb.EOF
180     OldValue = Val(tb!OldV)
190     NewValue = Val(tb!NewV)
200     If OldValue <> 0 Then
210         DeltaLimit = tb!DV
220         If Abs(OldValue - NewValue) > DeltaLimit Then
230             Res = Format$(tb!SampleDate, "dd/mm/yyyy") & " (" & tb!SampleID & ") " & _
                      tb!ShortName & " " & _
                      OldValue & vbCr
240             ldelta = ldelta & Res
250         End If
260     End If
270     tb.MoveNext
280 Loop

290 Exit Sub

300 Exit Sub

DoDeltaCheckBio_Error:

    Dim strES As String
    Dim intEL As Integer

310 intEL = Erl
320 strES = Err.Description
330 LogError "frmEditAll", "DoDeltaCheckBio", intEL, strES, sql

End Sub


Private Sub SetFormCaption()

10  On Error GoTo SetFormCaption_Error



20  Me.Caption = "NetAcquire"
30  Select Case ssTabAll.Tab
    Case 0: Me.Caption = Me.Caption & " - Demographics"
40  Case 1: Me.Caption = Me.Caption & " - Haematology"
50  Case 2: Me.Caption = Me.Caption & " - Biochemistry"
60  Case 3: Me.Caption = Me.Caption & " - Coagulation"
70  Case 4: Me.Caption = Me.Caption & " - Endocrinology"
80  Case 5: Me.Caption = Me.Caption & " - Blood Gas"
90  Case 6: Me.Caption = Me.Caption & " - Immunology"
100 Case 7: Me.Caption = Me.Caption & " - Externals"

110 End Select

120 Exit Sub

SetFormCaption_Error:

    Dim strES As String
    Dim intEL As Integer

130 intEL = Erl
140 strES = Err.Description
150 LogError "frmEditAll", "SetFormCaption", intEL, strES

End Sub




Private Sub LoadControls()
10  On Error GoTo LoadControls_Error

20  txtText.Visible = False
30  txtText = ""
    'gRD.SetFocus

40  Select Case grd.Col
    Case 1:
50      txtText.Move ssTabAll.Left + Panel3D4.Left + grd.Left + grd.CellLeft + 25, _
                     ssTabAll.Top + Panel3D4.Top + grd.Top + grd.CellTop + 5, _
                     grd.CellWidth, grd.CellHeight - 20
60      txtText.Text = grd.TextMatrix(grd.Row, grd.Col)
70      txtText.Visible = True
80      txtText.SelStart = 0
90      txtText.SelLength = Len(txtText)
100     txtText.SetFocus

110 End Select

120 Exit Sub

LoadControls_Error:

    Dim strES As String
    Dim intEL As Integer

130 intEL = Erl
140 strES = Err.Description
150 LogError "frmOptions", "LoadControls", intEL, strES

End Sub

Private Sub txtText_KeyUp(KeyCode As Integer, Shift As Integer)
10  If KeyCode = vbKeyUp Then
        'GoOneRowUp
20  ElseIf KeyCode = vbKeyDown Then
        'GoOneRowDown
30  ElseIf KeyCode = 13 Then
40      txtText.Visible = False
50  Else
60      grd.TextMatrix(grd.Row, grd.Col) = txtText
70      cmdSaveHaem.Enabled = True: cmdHSaveH.Enabled = True
80      bValidateHaem.Enabled = True
90  End If
End Sub

Private Sub GoOneRowUp()
10  If grd.Row > 1 Then
20      grd.Row = grd.Row - 1
30      LoadControls
40  End If
End Sub
Private Sub GoOneRowDown()
10  If grd.Row < grd.Rows - 1 Then
20      grd.Row = grd.Row + 1
30      LoadControls
40  End If
End Sub


Private Sub SetGridColor()
    Dim sql As String
    Dim tb As New ADODB.Recordset
    Dim i As Byte


10  On Error GoTo SetGridColor_Error

20  With grdExt

30      For i = 1 To .Rows - 1
40          sql = "SELECT MediBridgeResults.*" & _
                " FROM  MedibridgeRequests  INNER JOIN MediBridgeResults" & _
                " ON MedibridgeRequests.SampleID = MediBridgeResults.SampleID" & _
                " AND MedibridgeRequests.TestName = MediBridgeResults.TestName" & _
                " WHERE MedibridgeRequests.sampleid=" & txtSampleID & " AND MedibridgeRequests.TestName='" & .TextMatrix(i, 0) & "'"

50          Set tb = New Recordset
60          RecOpenServer 0, tb, sql
70          .Col = 9
80          .Row = i
90          If tb.EOF And tb.BOF Then
100             .CellBackColor = vbRed
110         Else
120             .CellBackColor = vbGreen
130         End If
140     Next
150 End With

160 On Error GoTo 0
170 Exit Sub

SetGridColor_Error:

    Dim strES As String
    Dim intEL As Integer

180 intEL = Erl
190 strES = Err.Description
200 LogError "frmEditAll", "SetGridColor", intEL, strES, sql
End Sub

Private Function GetAnyFieldFromTestDefinitions(ByVal FieldName As String, _
                                 ByVal TestShortName As String, _
                                 ByVal SampleID As String, _
                                 ByVal Discipline As String, _
                                 ByVal Cat As String) As String

      'Discipline is either "Bio", "Imm" or "End"
          Dim tb As New Recordset
          Dim sql As String
          Dim Dob As String
          Dim DaysOld As Long
          Dim TableName As String


10        On Error GoTo GetAnyFieldFromTestDefinitions_Error

20        If UCase(Discipline) = "BIO" Or UCase(Discipline) = "BGA" Then Cat = ""

30        TableName = Discipline & "TestDefinitions"

40        sql = "SELECT DoB, Sex,rundate from Demographics WHERE " & _
                "SampleID = '" & Val(SampleID) & "'"
          '50        Set tb = Cnxn(Cn).Execute(sql)
50        Set tb = New Recordset
60        RecOpenClient 0, tb, sql
70        If Not tb.EOF Then
80            If IsDate(tb!Dob) Then
90                Dob = Format$(tb!Dob, "dd/mmm/yyyy")
100               DaysOld = DateDiff("d", Dob, tb!Rundate)
110               If DaysOld = 0 Then DaysOld = 1
120           End If
130       End If


140       sql = "SELECT DISTINCT(" & FieldName & ") FROM " & TableName & "" & _
              " WHERE shortname = '" & TestShortName & "' AND (Category = '" & Cat & "') AND (AgeFromDays <= " & DaysOld & ") AND (AgeToDays >= " & DaysOld & ")"

150       Set tb = New Recordset
160       RecOpenClient 0, tb, sql
170       If tb.EOF Then
180           GetAnyFieldFromTestDefinitions = ""
190       Else
200           GetAnyFieldFromTestDefinitions = tb!Hospital & ""
210       End If
220       Exit Function

GetAnyFieldFromTestDefinitions_Error:

          Dim strES As String
          Dim intEL As Integer

230       intEL = Erl
240       strES = Err.Description
250       LogError "frmEditAll", "GetAnyFieldFromTestDefinitions", intEL, strES, sql

End Function
