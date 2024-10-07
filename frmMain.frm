VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NetAcquire - Custom Software"
   ClientHeight    =   8625
   ClientLeft      =   675
   ClientTop       =   1275
   ClientWidth     =   9330
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   204
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00004000&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8625
   ScaleWidth      =   9330
   StartUpPosition =   1  'CenterOwner
   Tag             =   "frmMain"
   Begin VB.CommandButton cmdConsultantList 
      Caption         =   "View &Consultant List"
      Height          =   435
      Left            =   5760
      TabIndex        =   28
      Top             =   6450
      Visible         =   0   'False
      Width           =   3075
   End
   Begin VB.CommandButton cmdUnvalidated 
      Caption         =   "&View Unvalidate/Not Printed"
      Height          =   435
      Left            =   5760
      TabIndex        =   26
      Top             =   5940
      Width           =   3075
   End
   Begin VB.CommandButton cmdMicroSurveillanceSearches 
      Caption         =   "Micro Surveillance Searches"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      Picture         =   "frmMain.frx":030A
      TabIndex        =   25
      Top             =   6900
      Visible         =   0   'False
      Width           =   3075
   End
   Begin VB.CheckBox chkAutoRefresh 
      Caption         =   "Auto-Refresh for last"
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
      Left            =   5700
      TabIndex        =   23
      Top             =   4860
      Width           =   1845
   End
   Begin VB.ComboBox cmbRefreshDays 
      Height          =   315
      ItemData        =   "frmMain.frx":0894
      Left            =   7605
      List            =   "frmMain.frx":08AA
      TabIndex        =   22
      Text            =   "14"
      Top             =   4800
      Width           =   675
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6060
      Top             =   7560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":08C3
            Key             =   "Ring0"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0BDD
            Key             =   "Ring1"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0EF7
            Key             =   "Ring2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1211
            Key             =   "Ring3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":152B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer TimerBar 
      Interval        =   1000
      Left            =   5340
      Top             =   5940
   End
   Begin VB.Timer tmrUrgent 
      Interval        =   60000
      Left            =   5340
      Top             =   6900
   End
   Begin VB.Frame frmUrg 
      Caption         =   "Urgent"
      Enabled         =   0   'False
      Height          =   4470
      Left            =   5490
      TabIndex        =   20
      Top             =   135
      Visible         =   0   'False
      Width           =   3750
      Begin MSFlexGridLib.MSFlexGrid grdUrg 
         Height          =   4020
         Left            =   180
         TabIndex        =   21
         ToolTipText     =   "Click on Number or Discipline"
         Top             =   270
         Width           =   3510
         _ExtentX        =   6191
         _ExtentY        =   7091
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
         ScrollBars      =   2
         FormatString    =   "<Sample Id      |^H  |^B  |^C  |^E  |^G  |^ I  "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Timer timerChk 
      Interval        =   10000
      Left            =   5340
      Top             =   6420
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1065
      Left            =   5970
      TabIndex        =   16
      Top             =   2880
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Timer tmrRefresh 
      Interval        =   30000
      Left            =   6930
      Top             =   4110
   End
   Begin VB.Frame Frame1 
      Caption         =   "Not Printed && No Results"
      Height          =   7845
      Left            =   60
      TabIndex        =   3
      Top             =   150
      Width           =   5325
      Begin VB.ListBox lstImmNotPrinted 
         Height          =   900
         Left            =   1245
         TabIndex        =   18
         Top             =   6210
         Visible         =   0   'False
         WhatsThisHelpID =   6
         Width           =   1875
      End
      Begin VB.ListBox lstImmNoResults 
         Height          =   900
         Left            =   3165
         TabIndex        =   17
         Top             =   6210
         Visible         =   0   'False
         WhatsThisHelpID =   7
         Width           =   1875
      End
      Begin VB.ListBox lstEndNoResults 
         Height          =   900
         Left            =   3150
         TabIndex        =   14
         Top             =   4905
         Visible         =   0   'False
         WhatsThisHelpID =   7
         Width           =   1875
      End
      Begin VB.ListBox lstEndNotPrinted 
         Height          =   900
         Left            =   1230
         TabIndex        =   13
         Top             =   4905
         Visible         =   0   'False
         WhatsThisHelpID =   6
         Width           =   1875
      End
      Begin MSFlexGridLib.MSFlexGrid gBioNoResults 
         Height          =   1635
         Left            =   3150
         TabIndex        =   11
         ToolTipText     =   "Click on Heading to Sort"
         Top             =   675
         WhatsThisHelpID =   2
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   2884
         _Version        =   393216
         FixedCols       =   0
         BackColorBkg    =   -2147483643
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   0
         GridLines       =   0
         GridLinesFixed  =   0
         ScrollBars      =   2
         FormatString    =   "^Analyser |<Sample ID"
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
      Begin VB.Timer tmrNotPrinted 
         Enabled         =   0   'False
         Interval        =   20000
         Left            =   450
         Top             =   810
      End
      Begin VB.ListBox lstCoagNotPrinted 
         Height          =   900
         Left            =   1230
         TabIndex        =   10
         Top             =   3615
         WhatsThisHelpID =   4
         Width           =   1875
      End
      Begin VB.ListBox lstCoagNoResults 
         Height          =   900
         Left            =   3150
         TabIndex        =   9
         Top             =   3615
         WhatsThisHelpID =   5
         Width           =   1875
      End
      Begin VB.ListBox lstHaemNotPrinted 
         Height          =   900
         Left            =   1230
         TabIndex        =   8
         Top             =   2325
         WhatsThisHelpID =   3
         Width           =   1875
      End
      Begin MSFlexGridLib.MSFlexGrid gBioNotPrinted 
         Height          =   1635
         Left            =   1230
         TabIndex        =   12
         ToolTipText     =   "Click on Heading to Sort"
         Top             =   675
         WhatsThisHelpID =   1
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   2884
         _Version        =   393216
         FixedCols       =   0
         BackColorBkg    =   -2147483643
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   0
         GridLines       =   0
         GridLinesFixed  =   0
         ScrollBars      =   2
         FormatString    =   "^Analyser |<Sample ID"
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
      Begin VB.Label lblImmEnd 
         AutoSize        =   -1  'True
         Caption         =   "No Results"
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
         Left            =   3180
         TabIndex        =   27
         Top             =   450
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label lblImmEnd 
         AutoSize        =   -1  'True
         Caption         =   "Immunology"
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
         TabIndex        =   19
         Top             =   6420
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label lblImmEnd 
         AutoSize        =   -1  'True
         Caption         =   "Endocrinology"
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
         Left            =   120
         TabIndex        =   15
         Top             =   5115
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Not Printed"
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
         Left            =   1200
         TabIndex        =   7
         Top             =   465
         Width           =   795
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Biochemistry"
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
         Left            =   120
         TabIndex        =   6
         Top             =   765
         Width           =   885
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Haematology"
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
         Left            =   90
         TabIndex        =   5
         Top             =   2235
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Coagulation"
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
         Left            =   120
         TabIndex        =   4
         Top             =   3825
         Width           =   840
      End
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   165
      Left            =   0
      TabIndex        =   2
      Top             =   8100
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Max             =   600
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   8250
      Width           =   9330
      _ExtentX        =   16457
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "04/06/2024"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            Object.Width           =   1244
            MinWidth        =   1235
            TextSave        =   "20:13"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4762
            MinWidth        =   4762
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Bevel           =   0
            Object.Width           =   10583
            MinWidth        =   10583
            Text            =   "Custom Software Ltd"
            TextSave        =   "Custom Software Ltd"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   5460
      Top             =   7560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2405
            Key             =   "Ring0"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":271F
            Key             =   "Ring1"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2A39
            Key             =   "Ring2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2D53
            Key             =   "Ring3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":306D
            Key             =   "Printer"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":36E7
            Key             =   "Fax"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3B39
            Key             =   "Key"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3F8B
            Key             =   "Locked"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":43DD
            Key             =   "ExtRequests"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5767
            Key             =   "ExtResults"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "day(s)"
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
      Left            =   8340
      TabIndex        =   24
      Top             =   4860
      Width           =   420
   End
   Begin VB.Image Image1 
      Height          =   675
      Left            =   5940
      Top             =   5160
      Width           =   915
   End
   Begin VB.Label lNewEXE 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
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
      Height          =   1665
      Left            =   6780
      TabIndex        =   1
      Top             =   7560
      Visible         =   0   'False
      Width           =   3645
   End
   Begin VB.Menu mfile 
      Caption         =   "&File"
      NegotiatePosition=   2  'Middle
      Begin VB.Menu mLogOn 
         Caption         =   "&Log On"
      End
      Begin VB.Menu mLogOff 
         Caption         =   "Log &Off"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuViewWards 
         Caption         =   "View &Ward Enquiries"
      End
      Begin VB.Menu mResetLastUsed 
         Caption         =   "&Reset 'Last Used'"
         Enabled         =   0   'False
      End
      Begin VB.Menu mshowerror 
         Caption         =   "&Show Error Log"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuArc 
         Caption         =   "&Archive"
         Enabled         =   0   'False
         Begin VB.Menu mnuArchive 
            Caption         =   "&General"
         End
         Begin VB.Menu mnuArchiveMicro 
            Caption         =   "&Microbiology"
         End
      End
      Begin VB.Menu mnuMain 
         Caption         =   "Maintenance"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnull 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHosp 
         Caption         =   "Hospital"
         Visible         =   0   'False
         Begin VB.Menu mnuHospital 
            Caption         =   ""
            Index           =   0
         End
         Begin VB.Menu mnuHospital 
            Caption         =   ""
            Index           =   1
            Visible         =   0   'False
         End
         Begin VB.Menu mnuHospital 
            Caption         =   ""
            Index           =   2
            Visible         =   0   'False
         End
         Begin VB.Menu mnuHospital 
            Caption         =   ""
            Index           =   3
            Visible         =   0   'False
         End
      End
      Begin VB.Menu exitmenu 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mEdit 
      Caption         =   "&Edit"
      Enabled         =   0   'False
      Begin VB.Menu mEditAll 
         Caption         =   "&General"
      End
      Begin VB.Menu mnuEditMicrobiology 
         Caption         =   "&Microbiology"
      End
      Begin VB.Menu mEditSemen 
         Caption         =   "&Semen Analysis"
      End
      Begin VB.Menu mnuOther 
         Caption         =   "&Histology/Cytology"
      End
   End
   Begin VB.Menu mOrd 
      Caption         =   "Order"
      Enabled         =   0   'False
      Begin VB.Menu morder 
         Caption         =   "&Order"
         Enabled         =   0   'False
      End
      Begin VB.Menu mhba1c 
         Caption         =   "&HbA1c/Ferritin/PSA"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMicroOrder 
         Caption         =   "Microbiology"
      End
   End
   Begin VB.Menu mnuBatch 
      Caption         =   "Batches"
      Enabled         =   0   'False
      Begin VB.Menu mnuBatHaem 
         Caption         =   "Haematology"
      End
      Begin VB.Menu mnuHaemFime 
         Caption         =   "Haematology Film"
      End
      Begin VB.Menu mnuHaemImm 
         Caption         =   "Haem Immuno"
      End
      Begin VB.Menu mnuImmBat 
         Caption         =   "Immunology"
         Begin VB.Menu mnuImmHiv 
            Caption         =   "HIV"
         End
         Begin VB.Menu mnuImmAuto 
            Caption         =   "AutoImmune Profile"
         End
      End
      Begin VB.Menu mnuBatMicro 
         Caption         =   "Microbiology"
         Begin VB.Menu mnuFae 
            Caption         =   "Faeces"
            Begin VB.Menu mnuFaecesLogIn 
               Caption         =   "Sample Log In"
            End
            Begin VB.Menu mnuBatFCul 
               Caption         =   "Culture"
            End
            Begin VB.Menu mnuBatchOccult 
               Caption         =   "Occult Blood"
            End
            Begin VB.Menu mnuBatERA 
               Caption         =   "Adeno + Rota + cDiff + HPylori"
            End
            Begin VB.Menu mnuBatOva 
               Caption         =   "Ova + Parasite"
            End
         End
         Begin VB.Menu mnuUru 
            Caption         =   "Urine"
            Begin VB.Menu mnuUrLog 
               Caption         =   "Sample &Log In"
            End
            Begin VB.Menu mnuUrBat 
               Caption         =   "&Microscopy"
            End
            Begin VB.Menu mnuPregnancy 
               Caption         =   "&Pregnancy"
            End
            Begin VB.Menu mnuIQ200Worklist 
               Caption         =   "&Worklist"
            End
         End
         Begin VB.Menu mnuBatchPrinting 
            Caption         =   "Batch Printing"
         End
      End
      Begin VB.Menu mnuBatchExt 
         Caption         =   "External"
      End
   End
   Begin VB.Menu msearch 
      Caption         =   "&Search"
      Enabled         =   0   'False
      Begin VB.Menu msearchmore 
         Caption         =   "&Name"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu msearchmore 
         Caption         =   "&Chart"
         Index           =   1
      End
      Begin VB.Menu msearchmore 
         Caption         =   "&Date of Birth"
         Index           =   2
      End
      Begin VB.Menu msearchmore1 
         Caption         =   "Name && Date of Birth"
         Index           =   3
      End
   End
   Begin VB.Menu mlists 
      Caption         =   "&Lists"
      Enabled         =   0   'False
      Begin VB.Menu mLocations 
         Caption         =   "Locations"
         Begin VB.Menu mListHospitals 
            Caption         =   "&Hospitals"
         End
         Begin VB.Menu mWards 
            Caption         =   "&Wards"
         End
         Begin VB.Menu mClinicians 
            Caption         =   "&Clinicians"
         End
         Begin VB.Menu mGPs 
            Caption         =   "&G.P.'s"
         End
      End
      Begin VB.Menu mComments 
         Caption         =   "Co&mments"
         Enabled         =   0   'False
         Begin VB.Menu mnuAutoComments 
            Caption         =   "&Auto Generated Comments"
            Begin VB.Menu mnuAutoComment 
               Caption         =   "&Biochemistry"
               Index           =   1
            End
            Begin VB.Menu mnuAutoComment 
               Caption         =   "&Immunology"
               Index           =   2
            End
            Begin VB.Menu mnuAutoComment 
               Caption         =   "&Endocrinology"
               Index           =   3
            End
         End
         Begin VB.Menu mnuCommentList 
            Caption         =   "&Biochemistry"
            Index           =   0
         End
         Begin VB.Menu mnuCommentList 
            Caption         =   "&Blood Gas"
            Index           =   1
         End
         Begin VB.Menu mnuCommentList 
            Caption         =   "&Coagulation"
            Index           =   2
         End
         Begin VB.Menu mnuCommentList 
            Caption         =   "&Cytology"
            Index           =   3
         End
         Begin VB.Menu mnuCommentList 
            Caption         =   "&Demographics"
            Index           =   4
         End
         Begin VB.Menu mnuCommentList 
            Caption         =   "&Endocrinology"
            Index           =   5
         End
         Begin VB.Menu mnuCommentList 
            Caption         =   "&Haematology"
            Index           =   6
         End
         Begin VB.Menu mnuCommentList 
            Caption         =   "&Histology"
            Index           =   7
         End
         Begin VB.Menu mnuCommentList 
            Caption         =   "&Immunolgy"
            Index           =   8
         End
         Begin VB.Menu mnuCommentList 
            Caption         =   "&Microbiology"
            Index           =   9
         End
         Begin VB.Menu mnuCommentList 
            Caption         =   "&Semen"
            Index           =   10
         End
         Begin VB.Menu mnuCommentList 
            Caption         =   "Consulta&nt Comments"
            Index           =   15
         End
         Begin VB.Menu mnuCommentList 
            Caption         =   "Me&dical Scientist Comments"
            Index           =   16
         End
      End
      Begin VB.Menu mDefaults 
         Caption         =   "&Defaults"
         Enabled         =   0   'False
         Begin VB.Menu mDefaultsBio 
            Caption         =   "&Biochemistry"
            Begin VB.Menu mnuBioAnalysers 
               Caption         =   "Biochemistry Analysers"
               Index           =   0
            End
            Begin VB.Menu mnuBioAnalysers 
               Caption         =   "Bio Test Code Mapping"
               Index           =   1
            End
            Begin VB.Menu mBioListSplits 
               Caption         =   "&Splits"
            End
            Begin VB.Menu mAddCode 
               Caption         =   "&Add Test"
            End
            Begin VB.Menu mpseq 
               Caption         =   "&Print Sequence"
               Visible         =   0   'False
            End
            Begin VB.Menu mpf 
               Caption         =   "Print &Format"
               Visible         =   0   'False
            End
            Begin VB.Menu mFasting 
               Caption         =   "Fasting Values"
            End
            Begin VB.Menu mdelta 
               Caption         =   "&Delta Check Limits"
               Visible         =   0   'False
            End
            Begin VB.Menu mnormal 
               Caption         =   "&Normal Ranges"
            End
            Begin VB.Menu mBioPlausible 
               Caption         =   "P&lausible Ranges"
            End
            Begin VB.Menu mBarCode 
               Caption         =   "&Bar Codes"
            End
            Begin VB.Menu mPanelsTop 
               Caption         =   "&Panels"
               Begin VB.Menu mPanels 
                  Caption         =   "&Define"
               End
            End
            Begin VB.Menu mneBioContChart 
               Caption         =   "Control Chart"
            End
         End
         Begin VB.Menu mnuBloodGa 
            Caption         =   "Blood &Gas"
            Begin VB.Menu mnuAddBgaTest 
               Caption         =   "Add Test"
            End
            Begin VB.Menu mnuBgaRanges 
               Caption         =   "Normal Ranges"
            End
         End
         Begin VB.Menu mDefaultsCoag 
            Caption         =   "&Coagulation"
            Begin VB.Menu mAddCoagTest 
               Caption         =   "&Add Test"
            End
            Begin VB.Menu mCoagDefinitions 
               Caption         =   "&Normal Ranges"
            End
            Begin VB.Menu mnuCoagContChart 
               Caption         =   "Control Chart"
            End
         End
         Begin VB.Menu mnuEndoDef 
            Caption         =   "Endocrinology"
            Begin VB.Menu mnuEndAnalysers 
               Caption         =   "Endocrinology Analysers"
               Index           =   0
            End
            Begin VB.Menu mnuEndAnalysers 
               Caption         =   "End Test Code Mapping"
               Index           =   1
            End
            Begin VB.Menu mnuEndoSplits 
               Caption         =   "&Splits"
            End
            Begin VB.Menu mnuCatImm 
               Caption         =   "Add &Caegory"
            End
            Begin VB.Menu mnuAddETest 
               Caption         =   "Add &Test"
            End
            Begin VB.Menu mnuNormRanges 
               Caption         =   "&Normal Ranges"
            End
            Begin VB.Menu mnuEndoPlausible 
               Caption         =   "P&lausible Ranges"
            End
            Begin VB.Menu mnuEndPanels 
               Caption         =   "&Panels"
            End
            Begin VB.Menu mnuAxsymResults 
               Caption         =   "A&xsym Results"
               Visible         =   0   'False
            End
         End
         Begin VB.Menu mDefaultsHaem 
            Caption         =   "&Haematology"
            Begin VB.Menu mnuHaeAnalysers 
               Caption         =   "Haematology Analysers"
               Index           =   0
            End
            Begin VB.Menu mnuHaeAnalysers 
               Caption         =   "Haem Test Code Mapping"
               Index           =   1
            End
            Begin VB.Menu mHaemDefinitions 
               Caption         =   "&Normal Ranges"
            End
            Begin VB.Menu mBarCodesH 
               Caption         =   "&Bar Codes"
            End
         End
         Begin VB.Menu mnuImm 
            Caption         =   "Immunology"
            Begin VB.Menu mnuImmSplit 
               Caption         =   "Splits"
            End
            Begin VB.Menu mnuImmCat 
               Caption         =   "Add Cat"
            End
            Begin VB.Menu mnuImmTest 
               Caption         =   "Add Test"
            End
            Begin VB.Menu mnuImmNorm 
               Caption         =   "Normal Ranges"
            End
            Begin VB.Menu mnuPlausiImm 
               Caption         =   "Plausible Ranges"
            End
            Begin VB.Menu mnuImmPanel 
               Caption         =   "Panels"
            End
            Begin VB.Menu mnuAllergyPanels 
               Caption         =   "&Allergy Panels"
            End
            Begin VB.Menu mnuAllergyMethods 
               Caption         =   "Allergy &Methods"
            End
            Begin VB.Menu mnuImmTestCodes 
               Caption         =   "Test &Codes"
            End
         End
         Begin VB.Menu mnuMicrobiology 
            Caption         =   "&Microbiology"
            Begin VB.Menu mnuMicrobiologySub 
               Caption         =   "&Urine"
               Index           =   0
               Begin VB.Menu mnuMicroUrineSub 
                  Caption         =   "Bacteria"
                  Index           =   0
               End
               Begin VB.Menu mnuMicroUrineSub 
                  Caption         =   "&WCC"
                  Index           =   1
               End
               Begin VB.Menu mnuMicroUrineSub 
                  Caption         =   "&RCC"
                  Index           =   2
               End
               Begin VB.Menu mnuMicroUrineSub 
                  Caption         =   "Crys&tal"
                  Index           =   3
               End
               Begin VB.Menu mnuMicroUrineSub 
                  Caption         =   "Ca&sts"
                  Index           =   4
               End
               Begin VB.Menu mnuMicroUrineSub 
                  Caption         =   "Miscellaneo&us"
                  Index           =   5
               End
               Begin VB.Menu mnuMicroUrineSub 
                  Caption         =   "-"
                  Index           =   6
               End
               Begin VB.Menu mnuMicroUrineSub 
                  Caption         =   "Pre&gnancy"
                  Index           =   7
               End
               Begin VB.Menu mnuMicroUrineSub 
                  Caption         =   "HCG Level"
                  Index           =   8
               End
            End
            Begin VB.Menu mnuMicrobiologySub 
               Caption         =   "&Identification"
               Index           =   1
               Begin VB.Menu mnuMicroIdentificationSub 
                  Caption         =   "&Gram Stains"
                  Index           =   0
               End
               Begin VB.Menu mnuMicroIdentificationSub 
                  Caption         =   "Wet Prep"
                  Index           =   1
               End
            End
            Begin VB.Menu mnuMicrobiologySub 
               Caption         =   "&Faeces"
               Index           =   2
               Begin VB.Menu mnuMicroFaecesSub 
                  Caption         =   "&XLD"
                  Index           =   0
               End
               Begin VB.Menu mnuMicroFaecesSub 
                  Caption         =   "&DCA"
                  Index           =   1
               End
               Begin VB.Menu mnuMicroFaecesSub 
                  Caption         =   "&SMAC"
                  Index           =   2
               End
               Begin VB.Menu mnuMicroFaecesSub 
                  Caption         =   "&CROMO"
                  Index           =   3
               End
               Begin VB.Menu mnuMicroFaecesSub 
                  Caption         =   "C&AMP"
                  Index           =   4
               End
               Begin VB.Menu mnuMicroFaecesSub 
                  Caption         =   "STE&C (day1)"
                  Index           =   5
               End
               Begin VB.Menu mnuMicroFaecesSub 
                  Caption         =   "STE&C (day2)"
                  Index           =   6
               End
            End
            Begin VB.Menu mnuMicrobiologySub 
               Caption         =   "&Titles"
               Index           =   3
               Begin VB.Menu mnuMicroTitlesSub 
                  Caption         =   "&FOB"
                  Index           =   0
               End
               Begin VB.Menu mnuMicroTitlesSub 
                  Caption         =   "H.&Pylori"
                  Index           =   1
               End
               Begin VB.Menu mnuMicroTitlesSub 
                  Caption         =   "C. Diff &Culture"
                  Index           =   2
               End
               Begin VB.Menu mnuMicroTitlesSub 
                  Caption         =   "C. Diff &ToxinAB"
                  Index           =   3
               End
               Begin VB.Menu mnuMicroTitlesSub 
                  Caption         =   "&GDH"
                  Index           =   4
               End
               Begin VB.Menu mnuMicroTitlesSub 
                  Caption         =   "GDH &Detail"
                  Index           =   5
               End
               Begin VB.Menu mnuMicroTitlesSub 
                  Caption         =   "&PCR"
                  Index           =   6
               End
               Begin VB.Menu mnuMicroTitlesSub 
                  Caption         =   "PCR &Detail"
                  Index           =   7
               End
               Begin VB.Menu mnuMicroTitlesSub 
                  Caption         =   "&Rota"
                  Index           =   8
               End
               Begin VB.Menu mnuMicroTitlesSub 
                  Caption         =   "&Adeno"
                  Index           =   9
               End
               Begin VB.Menu mnuMicroTitlesSub 
                  Caption         =   "&RSV"
                  Index           =   10
               End
               Begin VB.Menu mnuMicroTitlesSub 
                  Caption         =   "&Cryptosporidium"
                  Index           =   11
               End
               Begin VB.Menu mnuMicroTitlesSub 
                  Caption         =   "OP Co&mments"
                  Index           =   12
               End
               Begin VB.Menu mnuMicroTitlesSub 
                  Caption         =   "Giardia Lamblia"
                  Index           =   13
               End
            End
            Begin VB.Menu mnuMicrobiologySub 
               Caption         =   "&Fluids"
               Index           =   4
               Begin VB.Menu mnuMicroFluidsSub 
                  Caption         =   "&Appearance"
                  Index           =   0
               End
               Begin VB.Menu mnuMicroFluidsSub 
                  Caption         =   "Cell Co&unt"
                  Index           =   1
               End
               Begin VB.Menu mnuMicroFluidsSub 
                  Caption         =   "&Gram Stains"
                  Index           =   2
               End
               Begin VB.Menu mnuMicroFluidsSub 
                  Caption         =   "&ZN Stains"
                  Index           =   3
               End
               Begin VB.Menu mnuMicroFluidsSub 
                  Caption         =   "&Leishman's Stains"
                  Index           =   4
               End
               Begin VB.Menu mnuMicroFluidsSub 
                  Caption         =   "&Wet Prep"
                  Index           =   5
               End
               Begin VB.Menu mnuMicroFluidsSub 
                  Caption         =   "&Crystals"
                  Index           =   6
               End
               Begin VB.Menu mnuMicroFluidsSub 
                  Caption         =   "&Sites"
                  Index           =   7
               End
            End
            Begin VB.Menu mnuMicrobiologySub 
               Caption         =   "&C && S"
               Index           =   5
               Begin VB.Menu mnuMicroCandSSub 
                  Caption         =   "&Sites"
                  Index           =   0
               End
               Begin VB.Menu mnuMicroCandSSub 
                  Caption         =   "Organism &Groups"
                  Index           =   1
               End
               Begin VB.Menu mnuMicroCandSSub 
                  Caption         =   "&Organisms"
                  Index           =   2
               End
               Begin VB.Menu mnuMicroCandSSub 
                  Caption         =   "&Antibiotics"
                  Index           =   3
               End
               Begin VB.Menu mnuMicroCandSSub 
                  Caption         =   "Antibiotic &Panels"
                  Index           =   4
               End
               Begin VB.Menu mnuMicroCandSSub 
                  Caption         =   "Micro Setup"
                  Index           =   5
                  Visible         =   0   'False
               End
            End
            Begin VB.Menu mnuMicrobiologySub 
               Caption         =   "-"
               Index           =   6
            End
            Begin VB.Menu mnuMicrobiologySub 
               Caption         =   "Tab Setup"
               Index           =   7
            End
         End
         Begin VB.Menu mnuExtDef 
            Caption         =   "External"
            Enabled         =   0   'False
            Begin VB.Menu mnuExternalGeneral 
               Caption         =   "&General"
               Begin VB.Menu mnuAddExtTest 
                  Caption         =   "Add External &Address"
               End
               Begin VB.Menu mnuAddTest 
                  Caption         =   "Add &Test"
               End
               Begin VB.Menu mnuExtAddPanel 
                  Caption         =   "Add &Panel"
               End
            End
            Begin VB.Menu mnuExternalMicro 
               Caption         =   "&Microbiology"
               Begin VB.Menu mnuMicroExtAddress 
                  Caption         =   "Add External &Address"
               End
               Begin VB.Menu mnuMicroAddTest 
                  Caption         =   "Add &Test"
               End
               Begin VB.Menu mnuMicroAddPanel 
                  Caption         =   "Add &Panel"
               End
            End
         End
         Begin VB.Menu mnuHisto 
            Caption         =   "Histology"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuSemenList 
            Caption         =   "Semen Analysis"
            Begin VB.Menu mnuSemenConsistencyList 
               Caption         =   "Consis&tency"
            End
            Begin VB.Menu mnuSemenVolumeList 
               Caption         =   "Volume"
            End
            Begin VB.Menu mnuSemenCountList 
               Caption         =   "Co&unt"
            End
            Begin VB.Menu mnuSemenTypeList 
               Caption         =   "Specimen Type"
            End
         End
         Begin VB.Menu mnuMisc 
            Caption         =   "Miscellaneous"
         End
         Begin VB.Menu mPanelBarCodes 
            Caption         =   "&Barcodes"
         End
      End
      Begin VB.Menu mPrinters 
         Caption         =   "&Printers"
         Enabled         =   0   'False
      End
      Begin VB.Menu mGeneralLists 
         Caption         =   "&General"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuDocumentControl 
         Caption         =   "Document Control Numbers"
      End
   End
   Begin VB.Menu mreports 
      Caption         =   "&Reports"
      Enabled         =   0   'False
      Begin VB.Menu mnu24Urn 
         Caption         =   "24 Hr Urine"
      End
      Begin VB.Menu mCreatClear 
         Caption         =   "&Creatinine Clearance"
      End
      Begin VB.Menu mneworklist 
         Caption         =   "&Worklist"
         Begin VB.Menu mworklist 
            Caption         =   "General Worklist"
         End
         Begin VB.Menu MnuHistoWk 
            Caption         =   "Histology"
         End
      End
      Begin VB.Menu mEod 
         Caption         =   "&End Of Day Summary"
         Begin VB.Menu mnueodRpt 
            Caption         =   "Biochemistry"
            Index           =   0
         End
         Begin VB.Menu mnueodRpt 
            Caption         =   "Blood Gas"
            Index           =   1
            Visible         =   0   'False
         End
         Begin VB.Menu mnueodRpt 
            Caption         =   "Coagulation"
            Index           =   2
         End
         Begin VB.Menu mnueodRpt 
            Caption         =   "Endocrinology"
            Index           =   3
         End
         Begin VB.Menu mnueodRpt 
            Caption         =   "Haematology"
            Index           =   4
         End
         Begin VB.Menu mnueodRpt 
            Caption         =   "Immunology"
            Index           =   5
         End
         Begin VB.Menu mnueodRpt 
            Caption         =   "External"
            Index           =   6
         End
      End
      Begin VB.Menu mneMicroRep 
         Caption         =   "Micro Reports"
         Begin VB.Menu mnuOutStand 
            Caption         =   "&Outstanding"
         End
         Begin VB.Menu mnuSiteCount 
            Caption         =   "&Site Count"
         End
         Begin VB.Menu mnuIsoRep 
            Caption         =   "&Isolate Report"
         End
         Begin VB.Menu mnuStatFea 
            Caption         =   "&Faeces Stats"
         End
         Begin VB.Menu mnuStatsFluids 
            Caption         =   "Flui&ds Stats"
         End
         Begin VB.Menu mnuUrnStats 
            Caption         =   "&Urine Stats"
         End
         Begin VB.Menu mnuMicroGeneral 
            Caption         =   "&General"
         End
      End
      Begin VB.Menu mnurHisto 
         Caption         =   "Histology"
         Begin VB.Menu mnuNCRI 
            Caption         =   "NCRI Report"
         End
         Begin VB.Menu mnuHistoCytoYear 
            Caption         =   "Histology Yearly Report"
         End
      End
      Begin VB.Menu Sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrintPriorities 
         Caption         =   "&Print Priorities"
      End
      Begin VB.Menu mnuPhoneLog 
         Caption         =   "Phone Log History"
      End
   End
   Begin VB.Menu mqc 
      Caption         =   "&Q.C."
      Enabled         =   0   'False
      Begin VB.Menu mnuBioLimits 
         Caption         =   "Biochemistry"
         Begin VB.Menu mqcview 
            Caption         =   "&View"
         End
         Begin VB.Menu mqclimits 
            Caption         =   "&Limits"
         End
      End
      Begin VB.Menu mnuCoagQc 
         Caption         =   "Coagulation"
         Begin VB.Menu mnuCoagView 
            Caption         =   "View"
         End
         Begin VB.Menu mnuCoagLimits 
            Caption         =   "Limits"
         End
      End
      Begin VB.Menu mmeans 
         Caption         =   "&Running Means"
      End
   End
   Begin VB.Menu mPrint 
      Caption         =   "&Print"
      Enabled         =   0   'False
      Begin VB.Menu mBatch 
         Caption         =   "&Batch"
      End
      Begin VB.Menu mG 
         Caption         =   "&Glucose"
         Begin VB.Menu mglucose 
            Caption         =   "By &Date"
         End
         Begin VB.Menu mGluByName 
            Caption         =   "By &Name"
         End
      End
   End
   Begin VB.Menu mstock 
      Caption         =   "&Stock"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
   Begin VB.Menu mS 
      Caption         =   "&Statistics"
      Enabled         =   0   'False
      Begin VB.Menu mnuSuperStats 
         Caption         =   "General Stats"
      End
      Begin VB.Menu mnuStBio 
         Caption         =   "Biochemistry"
         Begin VB.Menu mtotbio 
            Caption         =   "Totals for Biochemistry"
         End
         Begin VB.Menu mnuTotBio 
            Caption         =   "Total Tests"
         End
         Begin VB.Menu mnuAbBio 
            Caption         =   "Abnormals"
         End
         Begin VB.Menu mnuBioUsa 
            Caption         =   "Usage"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnuStCoag 
         Caption         =   "Coagulation"
         Begin VB.Menu mnuStTotCoag 
            Caption         =   "Total For Coa&gulation"
         End
         Begin VB.Menu mnuCoToTe 
            Caption         =   "Total Tests"
         End
         Begin VB.Menu mnuCoSoTo 
            Caption         =   "Source Totals"
         End
      End
      Begin VB.Menu mnuEndoSt 
         Caption         =   "Endocrinology"
         Begin VB.Menu mnuStTotEndo 
            Caption         =   "Total For Endocrinology"
         End
         Begin VB.Menu mnuStTotEn 
            Caption         =   "Totals"
         End
         Begin VB.Menu mnuEndAb 
            Caption         =   "Abnormals"
         End
      End
      Begin VB.Menu mnuStHaem 
         Caption         =   "Haematology"
         Begin VB.Menu mtothaem 
            Caption         =   "Totals for Haematology"
         End
      End
      Begin VB.Menu mnuStatImm 
         Caption         =   "Immunology"
         Begin VB.Menu mnuStTotImm 
            Caption         =   "Totals for Immunology"
         End
         Begin VB.Menu mnuStTotIm 
            Caption         =   "Totals"
         End
         Begin VB.Menu mnuImmAbn 
            Caption         =   "Abnormals"
         End
      End
      Begin VB.Menu mnuMicroStatistics 
         Caption         =   "&Microbiology"
      End
      Begin VB.Menu mneExte 
         Caption         =   "External"
         Begin VB.Menu mnuExtStats 
            Caption         =   "External Stats"
         End
         Begin VB.Menu mnuOut 
            Caption         =   "Outstanding External "
         End
         Begin VB.Menu mnuExtSou 
            Caption         =   "Stats by Source"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnurepHisto 
         Caption         =   "Histology/Cytology"
         Begin VB.Menu mnuMissing 
            Caption         =   "Missing"
         End
         Begin VB.Menu mnuHistoStat 
            Caption         =   "Histology Stats"
         End
         Begin VB.Menu mnuCytoStat 
            Caption         =   "Cytology Stats"
         End
         Begin VB.Menu mnuFrozen 
            Caption         =   "Frozen Section"
         End
      End
      Begin VB.Menu mViewStats 
         Caption         =   "View"
      End
      Begin VB.Menu mSetSourceNames 
         Caption         =   "Set Source Names"
      End
      Begin VB.Menu mnuBad 
         Caption         =   "Bad Results"
      End
      Begin VB.Menu mnuGPClinWard 
         Caption         =   "GP/Clin/Ward"
      End
      Begin VB.Menu mnuStatsCol 
         Caption         =   "Statistics Collection"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuStoCon 
      Caption         =   "Stock Control"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnuStock 
         Caption         =   "Check Stock"
      End
      Begin VB.Menu mnuAdReg 
         Caption         =   "Administer Reagents"
      End
      Begin VB.Menu mnuUpdatestock 
         Caption         =   "Update Stock"
      End
   End
   Begin VB.Menu mhelp 
      Caption         =   "&Help"
      Begin VB.Menu mwinhelp 
         Caption         =   "&Windows Help"
      End
      Begin VB.Menu mtechnical 
         Caption         =   "&Technical Assistance"
      End
      Begin VB.Menu mnull1 
         Caption         =   "-"
      End
      Begin VB.Menu mabout 
         Caption         =   "&About"
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "Options"
      End
      Begin VB.Menu mnuUserOpt 
         Caption         =   "User Options"
      End
      Begin VB.Menu mnuWeb 
         Caption         =   "Go To Custom Software Website"
      End
   End
   Begin VB.Menu mnuTest 
      Caption         =   "Test"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuTestRange 
      Caption         =   "NewRange"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Option Compare Text
'15/Jul/2004

Private Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private LoadAx As Boolean

Private Const SW_SHOWNORMAL = 1



Private Sub Check_Ax()

          Dim ax As Control

10        On Error GoTo Check_Ax_Error

20        If SysOptDeptBio(0) = False Then grdUrg.ColWidth(2) = 0
30        If SysOptDeptHaem(0) = False Then grdUrg.ColWidth(1) = 0
40        If SysOptDeptCoag(0) = False Then grdUrg.ColWidth(3) = 0
50        If SysOptDeptEnd(0) = False Then grdUrg.ColWidth(4) = 0
60        If SysOptDeptBga(0) = False Then grdUrg.ColWidth(5) = 0
70        If SysOptDeptImm(0) = False Then grdUrg.ColWidth(6) = 0

80        lblImmEnd(0).Visible = SysOptDeptEnd(0)
90        lstEndNotPrinted.Visible = SysOptDeptEnd(0)
100       lstEndNoResults.Visible = SysOptDeptEnd(0)

110       lblImmEnd(1).Visible = SysOptDeptImm(0)
120       lstImmNotPrinted.Visible = SysOptDeptImm(0)
130       lstImmNoResults.Visible = SysOptDeptImm(0)
140       frmUrg.Visible = SysOptUrgent(0)

150       mnuBad.Visible = SysOptBadRes(0)

160       For Each ax In Me.Controls
170           If TypeOf ax Is Menu Then
180               If InStr(ax.Caption, "Biochem") Then
190                   ax.Enabled = SysOptDeptBio(0)
200               ElseIf InStr(ax.Caption, "Haem") Then
210                   ax.Enabled = SysOptDeptHaem(0)
220               ElseIf InStr(ax.Caption, "Coag") Then
230                   ax.Enabled = SysOptDeptCoag(0)
240               ElseIf InStr(ax.Caption, "Extern") Then
250                   ax.Enabled = SysOptDeptExt(0)
260               ElseIf InStr(ax.Caption, "Blood Gas") Then
270                   ax.Enabled = SysOptDeptBga(0)
280               ElseIf InStr(ax.Caption, "Imm") Then
290                   ax.Enabled = SysOptDeptImm(0)
300               ElseIf InStr(ax.Caption, "Endo") Then
310                   ax.Enabled = SysOptDeptEnd(0)
320               ElseIf InStr(ax.Caption, "Micro") Then
330                   ax.Enabled = SysOptDeptMicro(0)
340               ElseIf InStr(ax.Caption, "Semen") Then
350                   ax.Enabled = SysOptDeptSemen(0)
360               ElseIf InStr(ax.Caption, "Histology") Then
370                   ax.Enabled = SysOptDeptHisto(0)
380               ElseIf InStr(ax.Caption, "Cyto") Then
390                   ax.Enabled = SysOptDeptCyto(0)
400               End If
410           End If
420       Next

430       LoadAx = True

440       Exit Sub

Check_Ax_Error:

          Dim strES As String
          Dim intEL As Integer

450       intEL = Erl
460       strES = Err.Description
470       LogError "frmMain", "Check_Ax", intEL, strES

End Sub


Private Sub CheckPanelTypeUpdate(ByVal TableName As String)

          Dim sql As String

10        On Error GoTo CheckPanelTypeUpdate_Error

20        sql = "UPDATE " & TableName & " " & _
                "SET PanelType = 'S' WHERE PanelType IS NULL"
30        Cnxn(0).Execute sql

40        Exit Sub

CheckPanelTypeUpdate_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmMain", "CheckPanelTypeUpdate", intEL, strES, sql

End Sub

Private Sub chkAutoRefresh_Click()

          Dim tb As New Recordset
          Dim sql As String

10        On Error GoTo chkAutoRefresh_Click_Error

20        If Trim(UserName) = "" Then Exit Sub

30        SysOptAutoRef(0) = Val(chkAutoRefresh.Value)

40        sql = "IF EXISTS (SELECT * FROM Options " & _
                "           WHERE Description = 'AUTOREF' " & _
                "           AND UserName = '" & UserName & "') " & _
                "  UPDATE Options " & _
                "  SET Contents = '" & chkAutoRefresh.Value & "' " & _
                "  WHERE Description = 'AUTOREF' " & _
                "  AND UserName = '" & UserName & "' " & _
                "ELSE " & _
                "  INSERT INTO Options " & _
                "  (Description, Contents, UserName) " & _
                "  VALUES " & _
                "  ('AutoRef', " & _
                "   '" & chkAutoRefresh.Value & "', " & _
                "   '" & AddTicks(UserName) & "')"
50        Cnxn(0).Execute sql

60        Exit Sub

chkAutoRefresh_Click_Error:

          Dim strES As String
          Dim intEL As Integer

70        intEL = Erl
80        strES = Err.Description
90        LogError "frmMain", "chkAutoRefresh_Click", intEL, strES

End Sub

Private Sub cmdMicroSurveillanceSearches_Click()
10        On Error GoTo cmdMicroSurveillanceSearches_Click_Error

20        frmMicroSurveillanceSearches.Show 1

30        Exit Sub

cmdMicroSurveillanceSearches_Click_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmMicroSurveillanceSearches", "cmdMicroSurveillanceSearches_Click", intEL, strES

End Sub

Private Sub cmdUnvalidated_Click()

10        On Error GoTo cmdUnvalidated_Click_Error

20        frmNotValidatedPrinted.Show 1

30        Exit Sub

cmdUnvalidated_Click_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmMain", "cmdUnvalidated_Click", intEL, strES

End Sub

Private Sub Command1_Click()
'dim tb as new recordset
'Dim sql As String
'Dim ForeName As String
'Dim s As SexName
'Dim SurName As String
'Dim n as long
'
'sql = "SELECT * from demographics"
'Set tb = New Recordset
'RecOpenServer 0, tb, sql
'
'Do While Not tb.EOF
'  If Trim(tb!PatName) & "" <> "" Then
'       For n = 1 To Len(Trim(tb!PatName))
'        If Mid(Trim(tb!PatName), n, 1) = " " Then
'          SurName = Mid(tb!PatName, n + 1, Len(Trim(tb!PatName)) - n + 1)
'          Exit For
'        Else
'          ForeName = ForeName & Mid(Trim(tb!PatName), n, 1)
'        End If
'       Next
'      Set s = colSexNames(ForeName)
'      If Not s Is Nothing Then
'        If s.sex <> "D" Then
'          tb!PatName = SurName & " " & ForeName
'          tb!sex = s.sex
'          tb.UPDATE
'        End If
'      End If
'  End If
'  ForeName = ""
'  SurName = ""
'  tb.MoveNext
'Loop

End Sub

Private Sub EnableMenus(ByVal Enable As Boolean)

10    On Error GoTo EnableMenus_Error

20    frmUrg.Enabled = Enable
30    mLogOn.Enabled = Not Enable
40    mLogOff.Enabled = Enable
50    mResetLastUsed.Enabled = Enable
60    mshowerror.Enabled = Enable
70    mEdit.Enabled = Enable
80    mOrd.Enabled = Enable
90    morder.Enabled = Enable
100   msearch.Enabled = Enable
110   mlists.Enabled = False
120   mreports.Enabled = Enable
130   mstock.Enabled = Enable
140   mPrint.Enabled = Enable
150   mqc.Enabled = Enable
160   mS.Enabled = False
170   mComments.Enabled = False
180   mDefaults.Enabled = False
190   mS.Enabled = False
200   mGeneralLists.Enabled = False
210   mnuDocumentControl.Enabled = False
220   mPrinters.Enabled = False
230   mnuBatch.Enabled = Enable
240   mnuStoCon.Enabled = False
250   mnuMain.Enabled = False
260   mshowerror.Visible = False
270   cmdConsultantList.Visible = False
280   cmdMicroSurveillanceSearches.Visible = False

290   If Not Enable Then
300       UserName = ""
310       UserCode = ""
320       UserMemberOf = ""
330   End If

340   If InStr(UserMemberOf, "Managers") > 0 Then
350       mlists.Enabled = True
360       mComments.Enabled = True
370       mDefaults.Enabled = True
380       mS.Enabled = True
390       mGeneralLists.Enabled = True
400       mnuDocumentControl.Enabled = True
410       mPrinters.Enabled = True
420       mnuStoCon.Enabled = True
430       mshowerror.Visible = True
440       cmdConsultantList.Visible = True
450       If ISITEMINLIST(UserName, "SSUsers") = True Then
460           cmdMicroSurveillanceSearches.Visible = True
470       Else
480           cmdMicroSurveillanceSearches.Visible = False
490       End If
500   End If

510   If UserMemberOf = "Managers" Then
520       mnuArc.Enabled = True
530       mnuMain.Enabled = True
540   End If

550   If UserMemberOf = "Secretarys" And SysOptGpClin(0) Then
560       mlists.Enabled = True
570   End If

580   If UserMemberOf <> "Administrators" Then
590       If SysOptExtDefault(0) Then
600           mnuExtDef.Enabled = True
610           mDefaults.Enabled = True
620       End If
630   End If

640   mnuEndoDef.Enabled = SysOptDeptEnd(0)

650   mnuEditMicrobiology.Enabled = SysOptDeptMicro(0)

660   If UserMemberOf <> "Administrators" Then
670       If SysOptExtDefault(0) Then
680           mnuExtDef.Enabled = True
690           mDefaults.Enabled = True
700           mnuEndoDef.Enabled = False
710           mDefaultsBio.Enabled = False
720           mDefaultsCoag.Enabled = False
730           mDefaultsHaem.Enabled = False
740           mnuImm.Enabled = False
750       Else
760           mnuEndoDef.Enabled = SysOptDeptEnd(0)
770           mDefaultsBio.Enabled = SysOptDeptBio(0)
780           mDefaultsCoag.Enabled = SysOptDeptCoag(0)
790           mDefaultsHaem.Enabled = SysOptDeptHaem(0)
800           mnuImm.Enabled = SysOptDeptImm(0)
810           mnuBatMicro.Enabled = SysOptDeptMicro(0)
820           mnuMicroOrder.Enabled = SysOptDeptMicro(0)
830       End If
840   End If

850   If UserMemberOf = "HistoLookUp" Then
860       mEdit.Enabled = False
870       mOrd.Enabled = False
880       mnuBatch.Enabled = False
890       mreports.Enabled = False
900       mqc.Enabled = False
910       mPrint.Enabled = False
920   End If

930   StatusBar1.Panels(3).Text = UserName

940   Exit Sub

EnableMenus_Error:

      Dim strES As String
      Dim intEL As Integer

950   intEL = Erl
960   strES = Err.Description
970   LogError "frmMain", "EnableMenus", intEL, strES

End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdConsultantList_Click
' Author    : XPMUser
' Date      : 2/27/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmdConsultantList_Click()
10        On Error GoTo cmdConsultantList_Click_Error


20        frmConsultantListView.Show 1    ' Masood 27-02-2014


30        Exit Sub


cmdConsultantList_Click_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmMain", "cmdConsultantList_Click", intEL, strES
End Sub

Private Sub exitmenu_Click()

10        Unload Me

End Sub

Private Sub Form_Activate()

          Dim Path As String
          Dim strVersion As String

10        On Error GoTo Form_Activate_Error

20        If Not IsIDE Then
30            If SysOptChange(0) = True Then
40                Path = CheckNewEXE("NetAcquire")    '<---Change this to your prog Name
50                If Path <> "" Then
60                    Shell App.Path & "\CustomStart.exe NetAcquire"    '<---Change this to your prog Name
70                    End
80                    Exit Sub
90                End If
100           End If
110       End If

120       If LoadAx = False Then Check_Ax

130       If UserName <> "" Then
140           If SysOptAutoRef(0) = False Then
150               chkAutoRefresh.Value = 0
160           Else
170               chkAutoRefresh.Value = 1
180           End If
190       End If
          'Added Revision in panel Zya, 12-22-23
200       strVersion = App.Major & "." & App.Minor & "." & App.Revision
          'Zyam
210       Me.Caption = "NetAcquire - Laboratory Information System. Version " & strVersion

220       tmrNotPrinted.Enabled = True
230       tmrUrgent.Enabled = True
240       TimerBar.Enabled = True

250       If chkAutoRefresh.Value = 1 Then
260           timerChk.Enabled = True
270           tmrRefresh.Enabled = True
280       End If

290       If TestSys = True Then Me.Caption = Me.Caption & " - TEST SYSTEM"
          'VisibiltyofConsultantCmd
          'frmUrg.Visible = (UCase(UserMemberOf) <> "HISTOLOOKUP")

300       Exit Sub

Form_Activate_Error:

          Dim strES As String
          Dim intEL As Integer

310       intEL = Erl
320       strES = Err.Description
330       LogError "frmMain", "Form_Activate", intEL, strES

End Sub


'---------------------------------------------------------------------------------------
' Procedure : VisibiltyofConsultantCmd
' Author    : XPMUser
' Date      : 3/14/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Sub VisibiltyofConsultantCmd()
      'Exit Sub
          Dim sql As String
          Dim tb As New ADODB.Recordset

10        On Error GoTo VisibiltyofConsultantCmd_Error

20        cmdConsultantList.Visible = False
30        sql = "select * from USERS where name ='" & UserName & "' And Upper(MemberOf) = Upper('Managers')"
40        Set tb = New Recordset
50        RecOpenClient 0, tb, sql
60        If Not tb.EOF Then
70            cmdConsultantList.Visible = True
80        End If
90        tb.Close

100       Exit Sub


VisibiltyofConsultantCmd_Error:

          Dim strES As String
          Dim intEL As Integer

110       intEL = Erl
120       strES = Err.Description
130       LogError "frmMain", "VisibiltyofConsultantCmd", intEL, strES, sql
End Sub

Private Sub Form_DblClick()

10        If IsIDE Then
20            frmDelta.Discipline = "Bio"
30            frmDelta.Show 1
40        End If

End Sub


Private Sub Form_Deactivate()

10        On Error GoTo Form_Deactivate_Error

20        tmrNotPrinted.Enabled = False
30        timerChk.Enabled = False
40        tmrRefresh.Enabled = False
50        tmrUrgent.Enabled = False
60        TimerBar.Enabled = False

70        Exit Sub

Form_Deactivate_Error:

          Dim strES As String
          Dim intEL As Integer



80        intEL = Erl
90        strES = Err.Description
100       LogError "frmMain", "Form_Deactivate", intEL, strES


End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

10        On Error GoTo Form_KeyPress_Error

20        pb = 0
30        pbCounter = 0

40        Exit Sub

Form_KeyPress_Error:

          Dim strES As String
          Dim intEL As Integer



50        intEL = Erl
60        strES = Err.Description
70        LogError "frmMain", "Form_KeyPress", intEL, strES


End Sub

Private Sub Form_Load()

ReDim Cnxn(0 To 2) As Connection

On Error GoTo Form_Load_Error

If App.PrevInstance Then End

SetFormStyle Me

ConnectToDatabase

FillInterpTable
GetControlNames
LoadOptions

If SysOptShowIQ200(0) = False Then
    mnuIQ200Worklist.Visible = False
End If

If UCase(HospName(0)) = "MULLINGAR" Then
   mnuStatsCol.Visible = True
End If

'CheckLoggedOnUsersInDb
'CheckIQ200RepeatsInDb
'CheckFaecesResultsInDb
'CheckFaecesResultsArcInDb
'
'PopulateFaeces

'80    EnsureColumnExists "SemenResults", "Motility", "nvarchar(50)"
'EnsureColumnExists "CoagRequests", "Analyser", "nvarchar(50)"
'EnsureColumnExists "CoagRequests", "Programmed", "tinyint"
'EnsureColumnExists "SemenResults", "GradeA", "nvarchar(50)"
'EnsureColumnExists "SemenResults", "GradeB", "nvarchar(50)"
'EnsureColumnExists "SemenResults", "GradeC", "nvarchar(50)"
'EnsureColumnExists "SemenResults", "GradeD", "nvarchar(50)"
'EnsureColumnExists "SemenResults", "pH", "nvarchar(50)"
'EnsureColumnExists "SemenResults", "SpecimenType", "nvarchar(50)"
'EnsureColumnExists "SemenResults", "Morphology", "nvarchar(50)"
'EnsureColumnExists "SemenResults", "MotilitySlow", "nvarchar(50)"
'
'
'EnsureColumnExists "SemenResultsArc", "GradeA", "nvarchar(50)"
'EnsureColumnExists "SemenResultsArc", "GradeB", "nvarchar(50)"
'EnsureColumnExists "SemenResultsArc", "GradeC", "nvarchar(50)"
'EnsureColumnExists "SemenResultsArc", "GradeD", "nvarchar(50)"
'EnsureColumnExists "SemenResultsArc", "pH", "nvarchar(50)"
'EnsureColumnExists "SemenResultsArc", "SpecimenType", "nvarchar(50)"
'EnsureColumnExists "SemenResultsArc", "Morphology", "nvarchar(50)"
'EnsureColumnExists "SemenResultsArc", "MotilitySlow", "nvarchar(50)"
'
'EnsureColumnExists "UrineIdent", "ZN", "nvarchar(50)"
'EnsureColumnExists "UrineIdent", "Indole", "nvarchar(50)"
'EnsureColumnExists "UrineIdentArc", "ZN", "nvarchar(50)"
'EnsureColumnExists "UrineIdentArc", "Indole", "nvarchar(50)"
'
'EnsureColumnExists "Isolates", "NonReportable", "tinyint"
'EnsureColumnExists "IsolatesArc", "NonReportable", "tinyint"
'EnsureColumnExists "GenericResults", "TestDateTime", "datetime NULL"
'EnsureColumnExists "GenericResults", "DateTimeOfRecord", "datetime NULL"
'EnsureColumnExists "GenericResults", "Valid", "tinyint NULL"
'EnsureColumnExists "GenericResults", "Printed", "tinyint NULL"
'EnsureColumnExists "GenericResults", "Counter", "int IDENTITY(1,1) NOT NULL"
'EnsureColumnExists "GenericResultsArc", "TestDateTime", "datetime NULL"
'EnsureColumnExists "GenericResultsArc", "DateTimeOfRecord", "datetime NULL"
'EnsureColumnExists "GenericResultsArc", "Valid", "tinyint NULL"
'EnsureColumnExists "GenericResultsArc", "Printed", "tinyint NULL"
'EnsureColumnExists "GenericResultsArc", "Counter", "int IDENTITY(1,1) NOT NULL"
'EnsureColumnExists "Antibiotics", "ReportName", "nvarchar(50)"
'
'EnsureColumnExists "PrintPending", "NoOfCopies", "int DEFAULT 1"
'EnsureColumnExists "PrintPending", "FinalInterim", "char(1) DEFAULT 'F'"
'EnsureColumnExists "BioTestDefinitions", "CheckTime", "int DEFAULT 1"
'EnsureColumnExists "EndTestDefinitions", "CheckTime", "int DEFAULT 1"
'EnsureColumnExists "CoagTestDefinitions", "CheckTime", "int DEFAULT 1"
'EnsureColumnExists "ImmTestDefinitions", "CheckTime", "int DEFAULT 1"
'EnsureColumnExists "HaemTestDefinitions", "CheckTime", "int DEFAULT 1"
'
'90    EnsureColumnExists "EndPanels", "PanelType", "nvarchar(50) DEFAULT 'S'"
'100   CheckPanelTypeUpdate "EndPanels"
'
'110   EnsureColumnExists "IPanels", "PanelType", "nvarchar(50) DEFAULT 'S'"
'120   CheckPanelTypeUpdate "IPanels"
'
'130   EnsureColumnExists "ExtTests", "Code", "nvarchar(50)"
'140   EnsureIndexExists "ExtTests", "Code", "IDX_ExtTest_Code"
'150   EnsureColumnExists "ExtTests", "Department", "nvarchar(50)"
'
'160   EnsureColumnExists "ExtPanels", "TestName", "nvarchar(50)"
'170   EnsureColumnExists "ExtPanels", "Department", "nvarchar(50)"
'
'180   EnsureColumnExists "FaecalRequests", "RedSub", "bit"
'
'190   EnsureColumnExists "Demographics", "PenicillinAllergy", "bit"
'200   EnsureColumnExists "ArcDemographics", "PenicillinAllergy", "bit"
'
'220   CheckMicroExternalsInDb
'230   EnsureColumnExists "MicroExternals", "OrderCSFGlu", "bit DEFAULT 0 NOT NULL"
'240   EnsureColumnExists "MicroExternals", "OrderCSFTP", "bit DEFAULT 0 NOT NULL"
'
'250   CheckMicroExternalResultsInDb
'260   CheckMicroExternalResultsArcInDb
'
'270   CheckUrineRequestsInDb
'
'280   EnsureColumnExists "Organisms", "Site", "nvarchar(50)"
'
'290   CheckMicroExtLabNameInDb
'
'300   CheckPrintValidLogInDb
'310   CheckLockStatusInDb
'320   CheckGenericResultsInDb
'330   CheckFaecesWorksheetInDb
'340   CheckIsolatesRepeatsInDb
'350   CheckIsolatesArcInDb
'360   CheckSensitivitiesRepeatsInDb
'370   CheckSensitivitiesArcInDb
'
'380   EnsureColumnExists "FaecalRequests", "cS", "nvarchar(50)"
'390   EnsureColumnExists "FaecalRequests", "ssScreen", "nvarchar(50)"
'400   EnsureColumnExists "FaecalRequests", "Campylobacter", "nvarchar(50)"
'410   EnsureColumnExists "FaecalRequests", "Coli0157", "nvarchar(50)"
'420   EnsureColumnExists "FaecalRequests", "Cryptosporidium", "nvarchar(50)"
'430   EnsureColumnExists "FaecalRequests", "Rota", "nvarchar(50)"
'440   EnsureColumnExists "FaecalRequests", "Adeno", "nvarchar(50)"
'450   EnsureColumnExists "FaecalRequests", "OB0", "nvarchar(50)"
'460   EnsureColumnExists "FaecalRequests", "OB1", "nvarchar(50)"
'470   EnsureColumnExists "FaecalRequests", "OB2", "nvarchar(50)"
'480   EnsureColumnExists "FaecalRequests", "OP", "nvarchar(50)"
'490   EnsureColumnExists "FaecalRequests", "ToxinAB", "nvarchar(50)"
'500   EnsureColumnExists "FaecalRequests", "HPylori", "nvarchar(50)"
'
'EnsureColumnExists "Faeces", "CDiffCulture", "nvarchar(50)"
'520   EnsureColumnExists "Faeces", "OB1", "nvarchar(1)"
'530   EnsureColumnExists "Faeces", "OB2", "nvarchar(1)"
'540   EnsureColumnExists "Faeces", "ToxinAB", "nvarchar(1)"
'550   EnsureColumnExists "Faeces", "Cryptosporidium", "nvarchar(50)"
'560   EnsureColumnExists "Faeces", "OP", "nvarchar(50)"
'570   EnsureColumnExists "Faeces", "HPylori", "nvarchar(50)"
'580   EnsureColumnExists "Faeces", "UserName", "nvarchar(50)"
'590   EnsureColumnExists "Faeces", "DateTimeOfRecord", "datetime default getdate()"
'600   EnsureColumnExists "Faeces", "Healthlink", "tinyint"
'610   EnsureColumnExists "FaecesArc", "UserName", "nvarchar(50)"
'620   EnsureColumnExists "FaecesArc", "DateTimeOfRecord", "datetime default getdate()"
'630   EnsureColumnExists "FaecesArc", "Healthlink", "tinyint"
'
'640   EnsureColumnExists "Isolates", "UserName", "nvarchar(50)"
'650   EnsureColumnExists "Isolates", "Healthlink", "tinyint"
'660   EnsureColumnExists "IsolatesArc", "UserName", "nvarchar(50)"
'670   EnsureColumnExists "IsolatesArc", "Healthlink", "tinyint"
'
'680   EnsureColumnExists "MicroSiteDetails", "UserName", "nvarchar(50)"
'690   EnsureColumnExists "MicroSiteDetailsArc", "UserName", "nvarchar(50)"
'
'700   UpdateExternals
'
'710   EnsureColumnExists "GenericResults", "UserName", "nvarchar(50)"
'720   EnsureColumnExists "GenericResults", "Healthlink", "tinyint"
'730   EnsureColumnExists "GenericResultsArc", "Healthlink", "tinyint"
'  EnsureColumnExists "GenericResults", "DateTimeOfRecord", "datetime DEFAULT getdate()"
'
'740   CheckFaecesArcInDb
'750   CheckGenericResultsArcInDb
'760   CheckMicroSiteDetailsArcInDb
'770   CheckSemenResultsArcInDb
'780   CheckUrineArcInDb
'790   CheckUrineIdentArcInDb
'
'800   EnsureColumnExists "Urine", "UserName", "nvarchar(50)"
'810   EnsureColumnExists "UrineArc", "UserName", "nvarchar(50)"
'
'820   EnsureColumnExists "SemenResults", "UserName", "nvarchar(50)"
'830   EnsureColumnExists "SemenResults", "DateTimeOfRecord", "datetime default getdate()"
'840   EnsureColumnExists "SemenResultsArc", "UserName", "nvarchar(50)"
'850   EnsureColumnExists "SemenResultsArc", "DateTimeOfRecord", "datetime default getdate()"
'
'860   EnsureColumnExists "UrineIdent", "UserName", "nvarchar(50)"
'870   EnsureColumnExists "UrineRequests", "UserName", "nvarchar(50)"
'
'880   CheckUrineRequestsArcInDb
'
'890   EnsureColumnExists "FaecalRequests", "UserName", "nvarchar(50)"
'900   EnsureColumnExists "FaecalRequests", "DateTimeOfRecord", "datetime default getdate()"
'
'910   CheckFaecalRequestsArcInDb
'
'920   EnsureColumnExists "ImmTestDefinitions", "Method", "nvarchar(50)"
'930   EnsureColumnExists "ImmTestDefinitionsArc", "Method", "nvarchar(50)"
'940   EnsureColumnExists "ImmTestDefinitions", "IsAllergy", "tinyint DEFAULT 0"
'950   EnsureColumnExists "ImmTestDefinitionsArc", "IsAllergy", "tinyint DEFAULT 0"
'960   EnsureColumnExists "ImmRequests", "Method", "nvarchar(50)"
EnsureColumnExists "BioTestDefinitionsArc", "ShowLessThan", "tinyint DEFAULT 0"
EnsureColumnExists "BioTestDefinitionsArc", "ShowMoreThan", "tinyint DEFAULT 0"
EnsureColumnExists "BioTestDefinitions", "ShowLessThan", "tinyint DEFAULT 0"
EnsureColumnExists "BioTestDefinitions", "ShowMoreThan", "tinyint DEFAULT 0"
EnsureColumnExists "ImmTestDefinitionsArc", "ShowLessThan", "tinyint DEFAULT 0"
EnsureColumnExists "ImmTestDefinitionsArc", "ShowMoreThan", "tinyint DEFAULT 0"
EnsureColumnExists "ImmTestDefinitions", "ShowLessThan", "tinyint DEFAULT 0"
EnsureColumnExists "ImmTestDefinitions", "ShowMoreThan", "tinyint DEFAULT 0"

EnsureColumnExists "EndTestDefinitionsArc", "ShowLessThan", "tinyint DEFAULT 0"
EnsureColumnExists "EndTestDefinitionsArc", "ShowMoreThan", "tinyint DEFAULT 0"
EnsureColumnExists "EndTestDefinitions", "ShowLessThan", "tinyint DEFAULT 0"
EnsureColumnExists "EndTestDefinitions", "ShowMoreThan", "tinyint DEFAULT 0"
EnsureColumnExists "ExternalDefinitions", "InUse", "bit DEFAULT 1 NULL"
EnsureColumnExists "Options", "Details", "nvarchar(1000) NULL"
EnsureColumnExists "Options", "optCategory", "nvarchar(100) NULL"
EnsureColumnExists "Options", "OptionName", "nvarchar(300) NULL"
EnsureColumnExists "FaecesWorksheet", "Day141", "nvarchar(50)"
EnsureColumnExists "FaecesWorksheet", "Day142", "nvarchar(50)"
EnsureColumnExists "FaecesWorksheet", "Day143", "nvarchar(50)"
EnsureColumnExists "FaecesWorksheet", "Day251", "nvarchar(50)"
EnsureColumnExists "FaecesWorksheet", "Day252", "nvarchar(50)"
EnsureColumnExists "FaecesWorksheet", "Day253", "nvarchar(50)"
EnsureColumnExists "Faeces", "GDH", "nvarchar(50)"
EnsureColumnExists "Faeces", "PCR", "nvarchar(100)"
EnsureColumnExists "Faeces", "GiardiaLambila", "nvarchar(50)"
EnsureColumnExists "FaecesAudit", "GDH", "nvarchar(50)"
EnsureColumnExists "FaecesAudit", "PCR", "nvarchar(100)"
EnsureColumnExists "FaecesAudit", "GiardiaLambila", "nvarchar(50)"
EnsureColumnExists "AutoComments", "CommentType", "tinyint NULL Default 0"
EnsureColumnExists "PhoneLog", "Title", "nvarchar(10) NULL"
EnsureColumnExists "PhoneLog", "PersonName", "nvarchar(50) NULL"
EnsureColumnExists "Faeces", "GDHDetail", "nvarchar(100)"
EnsureColumnExists "Faeces", "PCRDetail", "nvarchar(100)"
EnsureColumnExists "FaecesAudit", "GDHDetail", "nvarchar(100)"
EnsureColumnExists "FaecesAudit", "PCRDetail", "nvarchar(100)"
EnsureColumnExists "FaecalRequests", "GDH", "bit NULL"
EnsureColumnExists "FaecalRequests", "PCR", "bit NULL"
EnsureColumnExists "FaecalRequestsAudit", "GDH", "bit NULL"
EnsureColumnExists "FaecalRequestsAudit", "PCR", "bit NULL"
'Masood 19_Feb_2013
EnsureColumnExists "ConsultantList", "Department", "nvarchar(100)"
EnsureColumnExists "ConsultantList", "Status", "nvarchar(100)"
EnsureColumnExists "ConsultantList", "Username", "nvarchar(100)"
EnsureColumnExists "ConsultantList", "DateTimeOfRecord", "datetime default getdate()"
EnsureColumnExists "PrintPending", "PrintAction", "nvarchar(100)"
EnsureColumnExists "EndResults", "DefIndex", "[numeric](18, 0)"
EnsureColumnExists "EndRepeats", "DefIndex", "[numeric](18, 0)"
EnsureColumnExists "EndResultsAudit", "DefIndex", "[numeric](18, 0)"
EnsureColumnExists "ImmResults", "DefIndex", "[numeric](18, 0)"
EnsureColumnExists "ImmRepeats", "DefIndex", "[numeric](18, 0)"
EnsureColumnExists "ImmResultsAudit", "DefIndex", "[numeric](18, 0)"
'Masood 19_Feb_2013
'**************************Ensure Option Exists
'====Farhan 21/10/2014====
EnsureColumnExists "MediBridgeResults", "Source", "nvarchar(100)"
EnsureColumnExists "MediBridgeResults", "Department", "nvarchar(100)"
EnsureColumnExists "MediBridgeResults", "TestName", "nvarchar(100)"
EnsureColumnExists "BioRequests", "Hospital", "nvarchar(50)"
EnsureColumnExists "EndRequests", "Hospital", "nvarchar(50)"
EnsureColumnExists "ImmRequests", "Hospital", "nvarchar(50)"
' ---------farhan---------
'Note: Other E.coli serotypes may cause EHEC/HUS. If there is a strong clinical suspicion despite the above negative test, please contact the microbiology lab which will send the specimen to a reference laboratory for further testing.
EnsureOptionExists "MicrobiologyEColi0157Comment", ""
'Please do not repeat unless change in risk factors for C.difficile associated diarrhoea
EnsureOptionExists "MicrobiologyCDiffPCRNegativeComment", ""
'Please do not repeat unless change in risk factors for C.difficile associated diarrhoea.
EnsureOptionExists "MicrobiologyCDiffGDHNegativeComment", ""

EnsureColumnExists "FaecalRequests", "GL", "bit NULL"
EnsureColumnExists "Reports", "status", "nvarchar(100)"
EnsureColumnExists "EndTestDefinitions", "HealthlinkPanel", "nvarchar(50)"
EnsureColumnExists "ConsultantList", "ACK", "bit NULL"
EnsureColumnExists "ConsultantList", "ConAck", "bit NULL"

EnsureColumnExists "ExternalDefinitions", "BiomnisCode", "nvarchar(50)"

'CheckPhoresisRequestsInDb
CheckPrintingRulesInDb
CheckUnauthorisedReportsInDb  'Masood 19_Feb_2013
CheckBioDefIndexInDb
CheckEndDefIndexInDb
CheckImmDefIndexInDb
CheckLabLinkCommunicationInDb
CheckLabLinkConnectionConfigInDb
CheckLabLinkMappingInDb
'Trevor 13th November 2015
'*************************
CheckDisablePrintingInDb
CheckBiomnisRequestsInDb
'*************************
Exit Sub

Form_Load_Error:

Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "frmMain", "Form_Load", intEL, strES

End Sub

Private Sub UpdateLoggedOnUser()

          Dim tb As Recordset
          Dim sql As String
          Dim MachineName As String

10        On Error GoTo UpdateLoggedOnUser_Error

20        MachineName = UCase$(vbGetComputerName())

30        sql = "IF EXISTS (SELECT * FROM LoggedOnUsers WHERE " & _
                "           MachineName = '" & MachineName & "' " & _
                "           AND AppName = 'NetAcquire') " & _
                "  UPDATE LoggedOnUsers " & _
                "  SET UserName = '" & UserName & "' " & _
                "  WHERE  MachineName = '" & MachineName & "' " & _
                "  AND AppName = 'NetAcquire'" & _
                "ELSE " & _
                "  INSERT INTO LoggedOnUsers " & _
                "  (MachineName, UserName, AppName) " & _
                "  VALUES " & _
                "  ('" & MachineName & "', " & _
                "   '" & UserName & "', " & _
                "   'NetAcquire')"
40        Cnxn(0).Execute sql

50        Exit Sub

UpdateLoggedOnUser_Error:

          Dim strES As String
          Dim intEL As Integer

60        intEL = Erl
70        strES = Err.Description
80        LogError "frmMain", "UpdateLoggedOnUser", intEL, strES, sql

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

          Static oButton As Integer
          Static oShift As Integer
          Static Ox As Single
          Static oY As Single

10        On Error GoTo Form_MouseMove_Error

20        If oButton <> Button Or oShift <> Shift Or Ox <> x Or oY <> Y Then
30            pb = 0
40            pbCounter = 0
50        End If

60        oButton = Button
70        oShift = Shift
80        Ox = x
90        oY = Y

100       Exit Sub

Form_MouseMove_Error:

          Dim strES As String
          Dim intEL As Integer

110       intEL = Erl
120       strES = Err.Description
130       LogError "frmMain", "Form_MouseMove", intEL, strES

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
          Dim n As Long

10        On Error GoTo Form_QueryUnload_Error

20        For n = 0 To intOtherHospitalsInGroup
30            Cnxn(n).Close
40        Next

50        Exit Sub

Form_QueryUnload_Error:

          Dim strES As String
          Dim intEL As Integer



60        intEL = Erl
70        strES = Err.Description
80        LogError "frmMain", "Form_QueryUnload", intEL, strES

End Sub


Private Sub Form_Unload(Cancel As Integer)

10    On Error GoTo Form_Unload_Error

20    End

30    Exit Sub

Form_Unload_Error:

       Dim strES As String
       Dim intEL As Integer

40     intEL = Erl
50     strES = Err.Description
60     LogError "frmMain", "Form_Unload", intEL, strES
          
End Sub

Private Sub gBioNoResults_Click()

          Static SortOrder As Boolean

10        On Error GoTo gBioNoResults_Click_Error

20        With gBioNoResults
30            If .MouseRow = 0 Then
40                If SortOrder Then
50                    .Sort = flexSortGenericAscending
60                Else
70                    .Sort = flexSortGenericDescending
80                End If
90                SortOrder = Not SortOrder
100           End If
110       End With

120       Exit Sub

gBioNoResults_Click_Error:

          Dim strES As String
          Dim intEL As Integer



130       intEL = Erl
140       strES = Err.Description
150       LogError "frmMain", "gBioNoResults_Click", intEL, strES


End Sub

Private Sub gBioNotPrinted_Click()

          Static SortOrder As Boolean

10        On Error GoTo gBioNotPrinted_Click_Error

20        With gBioNotPrinted
30            If .MouseRow = 0 Then
40                If SortOrder Then
50                    .Sort = flexSortGenericAscending
60                Else
70                    .Sort = flexSortGenericDescending
80                End If
90                SortOrder = Not SortOrder
100           End If
110       End With

120       Exit Sub

gBioNotPrinted_Click_Error:

          Dim strES As String
          Dim intEL As Integer



130       intEL = Erl
140       strES = Err.Description
150       LogError "frmMain", "gBioNotPrinted_Click", intEL, strES


End Sub

Public Sub GetLogOn()

10        On Error GoTo GetLogOn_Error

20        With frmManager
30            .LookUp = True
40            .Operator = True
50            .Administrator = True
60            .Manager = True
70            .Secretary = True
80            .HistoLookUp = True
90            .Show 1
100       End With

110       pb = 0
120       pbCounter = 0
130       If LogOffDelayMin > 0 Then
140           pb.Max = LogOffDelaySecs
150       Else
160           pb.Max = 10
170       End If

180       Exit Sub

GetLogOn_Error:

          Dim strES As String
          Dim intEL As Integer



190       intEL = Erl
200       strES = Err.Description
210       LogError "frmMain", "GetLogOn", intEL, strES


End Sub

Private Sub grdUrg_Click()

          Dim TempTab


10        On Error GoTo grdUrg_Click_Error

20        If grdUrg.RowSel = 0 Then Exit Sub

30        If grdUrg.TextMatrix(grdUrg.RowSel, 0) = "" Then
40            Exit Sub
50        End If

60        If UCase(UserMemberOf) = "HISTOLOOKUP" Then
70            iMsg "You are not allowed to view urgent samples"
80            Exit Sub
90        End If

100       With frmEditAll
110           .txtSampleID = grdUrg.TextMatrix(grdUrg.RowSel, 0)
120           .txtSampleID_LostFocus
              '  .ClearDemographics
              '  .LoadDemographics
              '  .LoadBiochemistry
              '  .LoadCoagulation
              '  .LoadEndocrinology
              '  .LoadImmunology
              '  .LoadHaematology
              '  .LoadBloodGas
130           TempTab = SysOptDefaultTab(0)
140           SysOptDefaultTab(0) = Val(grdUrg.ColSel)
150           .Show 1
160           SysOptDefaultTab(0) = TempTab
170       End With

180       Exit Sub

grdUrg_Click_Error:

          Dim strES As String
          Dim intEL As Integer



190       intEL = Erl
200       strES = Err.Description
210       LogError "frmMain", "grdUrg_Click", intEL, strES


End Sub

Private Sub Image1_Click()
          Dim s As String

10        On Error GoTo Image1_Click_Error

20        s = UCase$(iBOX("Password?", , , True))
30        If s <> SysOptPrintAll(0) Or s = "" Then Exit Sub

40        frmUpdatePrinted.Show 1

50        Exit Sub

Image1_Click_Error:

          Dim strES As String
          Dim intEL As Integer



60        intEL = Erl
70        strES = Err.Description
80        LogError "frmMain", "Image1_Click", intEL, strES


End Sub







Private Sub Label6_Click()

End Sub

Private Sub lstImmNotPrinted_Click()
          Dim n As Integer

10        On Error GoTo lstImmNotPrinted_Click_Error

20        If iMsg("Do You want to print this!", vbYesNo) = vbYes Then
30            For n = 0 To lstImmNotPrinted.ListCount
40                Printer.Print lstImmNotPrinted.List(n)
50            Next
60            Printer.EndDoc
70        End If


80        Exit Sub

lstImmNotPrinted_Click_Error:

          Dim strES As String
          Dim intEL As Integer



90        intEL = Erl
100       strES = Err.Description
110       LogError "frmMain", "lstImmNotPrinted_Click", intEL, strES


End Sub

Private Sub mabout_Click()

10        frmAbout.Show 1

End Sub

Private Sub mAddCoagTest_Click()

10        frmAddCoagTest.Show 1

End Sub

Private Sub mAddCode_Click()

10        frmAddNewTest.Department = "Biochemistry"
20        frmAddNewTest.Show 1

End Sub

Private Sub mbarcode_Click()

10        frmBarCodes.Show 1

End Sub

Private Sub mBarCodesH_Click()

10        frmBarCodes.Show 1

End Sub

Private Sub mbatch_Click()

10        frmPrintOptions.Show 1

End Sub

Private Sub mBioListSplits_Click()

10        On Error GoTo mBioListSplits_Click_Error

20        With frmBioSplitList
30            .Disp = "Bio"
40            .Show 1
50        End With

60        Exit Sub

mBioListSplits_Click_Error:

          Dim strES As String
          Dim intEL As Integer



70        intEL = Erl
80        strES = Err.Description
90        LogError "frmMain", "mBioListSplits_Click", intEL, strES

End Sub

Private Sub mBioPlausible_Click()

10        frmBioPlausible.Show 1

End Sub

Private Sub mClinicians_Click()

10        frmClinicians.Show 1

End Sub

Private Sub mCoagDefinitions_Click()

10        frmCoagDefinitions.Show 1

End Sub



Private Sub mCreatClear_Click()

10        frmCreatClear.Show 1

End Sub

Private Sub mdelta_Click()

'fdelta.Show 1

End Sub


Private Sub mEditAll_Click()

10        frmEditAll.Show 1

End Sub

Private Sub mEditSemen_Click()

10        frmEditSemen.Show 1

End Sub

Private Sub mFasting_Click()

10        frmFastings.Show 1

End Sub

Private Sub mGeneralLists_Click()

10        frmLists.Show 1

End Sub

Private Sub mGluByName_Click()

10        frmGluByName.Show 1

End Sub

Private Sub mglucose_Click()

10        frmGlucose.Show 1

End Sub

Private Sub mGPs_Click()

10        frmGps.Show 1

End Sub

Private Sub mHaemDefinitions_Click()

10        frmHaemDefinitions.Show 1

End Sub

Private Sub mListHospitals_Click()

10        frmHospital.Show 1

End Sub

Private Sub mLogOff_Click()

10        On Error GoTo mLogOff_Click_Error

20        UserName = ""
30        UserCode = ""

40        EnableMenus False

50        UpdateLoggedOnUser

60        Exit Sub

mLogOff_Click_Error:

          Dim strES As String
          Dim intEL As Integer



70        intEL = Erl
80        strES = Err.Description
90        LogError "frmMain", "mLogOff_Click", intEL, strES


End Sub

Private Sub mLogOn_Click()

10        On Error GoTo mLogOn_Click_Error


20        GetLogOn

30        If UserMemberOf <> "Administrators" Then
40            If UserCode <> "" Then EnableMenus UserCode <> ""
50        End If

60        UpdateLoggedOnUser

70        Exit Sub

mLogOn_Click_Error:

          Dim strES As String
          Dim intEL As Integer

80        intEL = Erl
90        strES = Err.Description
100       LogError "frmMain", "mLogOn_Click", intEL, strES

End Sub

Private Sub mmeans_Click()

10        frmViewRM.Show 1

End Sub

Private Sub mneBioContChart_Click()

10        On Error GoTo mneBioContChart_Click_Error

20        With frmBioChart
30            .Caption = .Caption & " Biochemistry"
40            .lblType = "QCB"
50            .Show 1
60        End With

70        Exit Sub

mneBioContChart_Click_Error:

          Dim strES As String
          Dim intEL As Integer



80        intEL = Erl
90        strES = Err.Description
100       LogError "frmMain", "mneBioContChart_Click", intEL, strES


End Sub

Private Sub mnormal_Click()

10        frmBioDefaults.Show 1

End Sub

Private Sub mnu24Urn_Click()

10        frm24hrUrine.Show 1

End Sub

Private Sub mnuAbBio_Click()

10        frmAbnormals.Show 1

End Sub

Private Sub mnuAddBgaTest_Click()

10        frmAddBgaCode.Show 1

End Sub

Private Sub mnuAddETest_Click()

10        frmAddNewTest.Department = "Endocrinology"

20        frmAddNewTest.Show 1

End Sub

Private Sub mnuAddExtTest_Click()

10        frmAddress.Show 1

End Sub

Private Sub mnuAddTest_Click()


10        With frmExtTests
20            .Department = "General"
30            .Show 1
40        End With
          '10    frmExternalTests.Show 1

End Sub

Private Sub mnuAdReg_Click()

10        frmReagentSet.Show 1

End Sub

Private Sub mnuAllergyMethods_Click()

10        frmAllergyMethods.Show 1

End Sub

Private Sub mnuAllergyPanels_Click()

10        frmAllergyPanels.Show 1

End Sub

Private Sub mnuArchive_Click()

10        frmArchive.Show 1

End Sub

Private Sub mnuArchiveMicro_Click()

10        frmAuditMicro.Show 1

End Sub

Private Sub mnuAutoComment_Click(Index As Integer)

10        frmAutoGenerateComments.Discipline = Mid$(mnuAutoComment(Index).Caption, 2)    'remove the "&"
20        frmAutoGenerateComments.Show 1

End Sub

Private Sub mnuAxsymResults_Click()
10        With frmListsGeneric
20            .ListType = "AxsymResults"
30            .ListTypeName = "Axsym Result"
40            .ListTypeNames = "Axsym Results"
50            .Show 1
60        End With
End Sub

Private Sub mnuBad_Click()

10        frmBadRes.Show 1

End Sub

Private Sub mnuBatchExt_Click()

10        frmExternalBatch.Show 1

End Sub

Private Sub mnuBatchOccult_Click()
          Dim strStatus As String

10        strStatus = getBatchEntryOpenStatus("OPTBATCHENTRYOCCULTBLOOD")

20        If strStatus = "" Then
30            frmBatchOccult.Show 1
40        Else
50            iMsg "User - " & strStatus & " - has the Occult Blood Batch Entry screen already opened." & vbCrLf & "Only one user can open the batch entry screen at a time."
60        End If

End Sub

Private Sub mnuBatchPrinting_Click()
10        frmBatchPrinting.Show 1
End Sub

Private Sub mnuBatERA_Click()

          Dim strStatus As String

10        strStatus = getBatchEntryOpenStatus("OPTBATCHENTRYERA")

20        If strStatus = "" Then
30            frmBatchERA.Show 1
40        Else
50            iMsg "User - " & strStatus & " - has the Batch Entry screen already opened." & vbCrLf & "Only one user can open the batch entry screen at a time."
60        End If


End Sub

Private Sub mnuBatFCul_Click()

10        frmBatECPS.Show 1

End Sub

Private Sub mnuBatHaem_Click()

10        frmAsot.Show 1

End Sub

Private Sub mnuBatOva_Click()

          Dim strStatus As String

10        strStatus = getBatchEntryOpenStatus("OPTBATCHENTRYOVA")

20        If strStatus = "" Then
30            frmBatchOva.Show 1
40        Else
50            iMsg "User - " & strStatus & " - has the Batch Entry screen already opened." & vbCrLf & "Only one user can open the batch entry screen at a time."
60        End If

End Sub

Private Sub mnuBgaRanges_Click()

10        frmBgaDefaults.Show 1

End Sub


Private Sub mnuBioAnalysers_Click(Index As Integer)

10        On Error GoTo mnuBioAnalysers_Click_Error

20        Select Case Index
          Case 0:
30            With frmListsGeneric
40                .ListType = "BioAnalysers"
50                .ListTypeName = "Biochemistry Analyser"
60                .ListTypeNames = "Biochemistry Analysers"
70                .Show 1
80            End With
90        Case 1:
100           frmTestCodeMapping.Discipline = "BIO"
110           frmTestCodeMapping.Show 1

120       End Select

130       Exit Sub

mnuBioAnalysers_Click_Error:

          Dim strES As String
          Dim intEL As Integer

140       intEL = Erl
150       strES = Err.Description
160       LogError "frmMain", "mnuBioAnalysers_Click", intEL, strES

End Sub

Private Sub mnuBioUsa_Click()

10        'frmStats.Show 1

End Sub







Private Sub mnuCatImm_Click()

10        frmAddCat.Show 1

End Sub

Private Sub mnuCoagContChart_Click()

10        On Error GoTo mnuCoagContChart_Click_Error

20        With frmBioChart
30            .Caption = .Caption & " Coagulation"
40            .lblType = "QCC"
50            .Show 1
60        End With

70        Exit Sub

mnuCoagContChart_Click_Error:

          Dim strES As String
          Dim intEL As Integer



80        intEL = Erl
90        strES = Err.Description
100       LogError "frmMain", "mnuCoagContChart_Click", intEL, strES

End Sub

Private Sub mnuCoagLimits_Click()

10        frmCoagLimits.Show 1

End Sub


Private Sub mnuCoagView_Click()

10        On Error GoTo mnuCoagView_Click_Error

20        frmQCparent.optCoag.Value = True
30        frmQCparent.Show

40        Exit Sub

mnuCoagView_Click_Error:

          Dim strES As String
          Dim intEL As Integer



50        intEL = Erl
60        strES = Err.Description
70        LogError "frmMain", "mnuCoagView_Click", intEL, strES

End Sub

Private Sub mnuCommentList_Click(Index As Integer)

10        On Error GoTo mnuCommentList_Click_Error

20        Select Case Index
          Case 1
30        Case 2
40        Case 3
50        Case 4
60        End Select

70        With frmComments
80            .optType(Index) = True
90            .Show 1
100       End With

110       Exit Sub

mnuCommentList_Click_Error:

          Dim strES As String
          Dim intEL As Integer



120       intEL = Erl
130       strES = Err.Description
140       LogError "frmMain", "mnuCommentList_Click", intEL, strES

End Sub

Private Sub mnuCoSoTo_Click()

10        frmCoagSourceTotals.Show 1

End Sub

Private Sub mnuCoToTe_Click()

10        On Error GoTo mnuCoToTe_Click_Error

20        frmCoagTotalTests.Show 1

30        Exit Sub

mnuCoToTe_Click_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmMain", "mnuCoToTe_Click", intEL, strES

End Sub



Private Sub mnuCytoStat_Click()

10        frmCytoStats.Show 1

End Sub

Private Sub mnuDocumentControl_Click()

10        On Error GoTo mnuDocumentControl_Click_Error

20        With frmOptions
30            .EditDescription = False
40            .SelectedType = "User"
50            .SelectedCategory = "Document Control Numbers"
60            .fraFilter.Visible = False

70            .Show 1
80        End With

90        Exit Sub

mnuDocumentControl_Click_Error:

          Dim strES As String
          Dim intEL As Integer

100       intEL = Erl
110       strES = Err.Description
120       LogError "frmMain", "mnuDocumentControl_Click", intEL, strES

End Sub

Private Sub mnuEditMicrobiology_Click()

10        frmEditMicrobiologyNew.Show 1

End Sub

Private Sub mnuEndAb_Click()

10        'frmEndAbnormalsNew.Show 1

End Sub


Private Sub mnuEndoPlausible_Click()

10        frmEndPlausible.Show 1

End Sub

Private Sub mnuEndoSplits_Click()

10        On Error GoTo mnuEndoSplits_Click_Error

20        With frmBioSplitList
30            .Disp = "End"
40            .Show 1
50        End With

60        Exit Sub

mnuEndoSplits_Click_Error:

          Dim strES As String
          Dim intEL As Integer



70        intEL = Erl
80        strES = Err.Description
90        LogError "frmMain", "mnuEndoSplits_Click", intEL, strES

End Sub

Private Sub mnuEndPanels_Click()

10        frmPanels.Department = "Endocrinology"
20        frmPanels.Show 1

End Sub


Private Sub mnueodRpt_Click(Index As Integer)

    Select Case Index
    Case 0:
        frmDayEndCommon.Show 1
    Case 1:
        frmBGSummary.Show 1
    Case 2:
        frmCoagSummary.Show 1
    Case 3:
        frmEndDayEndCommon.Show 1
    Case 4:
        frmHaemSummary.Show 1
    Case 5:
        frmEndDayImmCommon.Show 1
    Case 6:
        frmEndDayExtCommon.Show 1
    End Select

End Sub



Private Sub mnuExtAddPanel_Click()

10        With frmExtPanels
20            .Department = "General"
30            .Show 1
40        End With

End Sub


Private Sub mnuExtSou_Click()

10        frmExternalStats.Show 1

End Sub

Private Sub mnuExtStats_Click()

10        frmExtStats.Show 1

End Sub

Private Sub mnuFaecesLogIn_Click()

10        frmBatchLogInFaeces.Show 1

End Sub

Private Sub mnuFrozen_Click()

10        frmFrozen.Show 1

End Sub

Private Sub mnuGPClinWard_Click()

10        frmGPClinWard.Show 1

End Sub



Private Sub mnuHaeAnalysers_Click(Index As Integer)

   On Error GoTo mnuHaeAnalysers_Click_Error

20        Select Case Index
          Case 0:
30            With frmListsGeneric
40                .ListType = "HaemAnalysers"
50                .ListTypeName = "Haematology Analyser"
60                .ListTypeNames = "Haematology Analysers"
70                .Show 1
80            End With
90        Case 1:
100           frmTestCodeMapping.Discipline = "HAEM"
110           frmTestCodeMapping.Show 1

120       End Select


   Exit Sub

mnuHaeAnalysers_Click_Error:

    Dim strES As String
    Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "frmMain", "mnuHaeAnalysers_Click", intEL, strES


End Sub



Private Sub mnuHaemFime_Click()

10        frmHaemBSummary.Show 1

End Sub

Private Sub mnuHaemImm_Click()

10        frmHaemImm.Show 1

End Sub

Private Sub mnuHisto_Click()

10        frmHistList.Show 1

End Sub

Private Sub mnuHistoCytoYear_Click()
10        frmHistoCytoReport.Show 1
End Sub

Private Sub mnuHistoStat_Click()

10        frmHistoStats.Show 1

End Sub

Private Sub MnuHistoWk_Click()

10        frmHistoWork.Show 1

End Sub


Private Sub mnuHospital_Click(Index As Integer)

10        On Error GoTo mnuHospital_Click_Error

20        mLogOff_Click

30        Set Cnxn(0) = New Connection
40        Cnxn(0).Open mnuHospital(0).Tag
50        HospName(0) = mnuHospital(0).Caption

60        Exit Sub

mnuHospital_Click_Error:

          Dim strES As String
          Dim intEL As Integer



70        intEL = Erl
80        strES = Err.Description
90        LogError "frmMain", "mnuHospital_Click", intEL, strES


End Sub

Private Sub mnuImmAbn_Click()

10        frmImmAbnormals.Show 1

End Sub

Private Sub mnuEndAnalysers_Click(Index As Integer)

10       On Error GoTo mnuEndAnalysers_Click_Error

20        Select Case Index
          Case 0:
30            With frmListsGeneric
40                .ListType = "EndAnalysers"
50                .ListTypeName = "Endocrinology Analyser"
60                .ListTypeNames = "Endocrinology Analysers"
70                .Show 1
80            End With
90        Case 1:
100           frmTestCodeMapping.Discipline = "END"
110           frmTestCodeMapping.Show 1
120       End Select

130      Exit Sub

mnuEndAnalysers_Click_Error:

    Dim strES As String
    Dim intEL As Integer

140   intEL = Erl
150   strES = Err.Description
160   LogError "frmMain", "mnuEndAnalysers_Click", intEL, strES

End Sub

Private Sub mnuImmAuto_Click()

10        frmImmBatch.Show 1

End Sub


Private Sub mnuImmCat_Click()

10        frmAddCat.Show 1

End Sub



Private Sub mnuImmHiv_Click()

10        frmBatchHiv.Show 1

End Sub

Private Sub mnuImmNorm_Click()

10        frmImmDefaults.Show 1

End Sub

Private Sub mnuImmPanel_Click()

10        frmPanels.Department = "Immunology"
20        frmPanels.Show 1

End Sub

Private Sub mnuImmSplit_Click()

10        On Error GoTo mnuImmSplit_Click_Error

20        With frmBioSplitList
30            .Disp = "Imm"
40            .Show 1
50        End With

60        Exit Sub

mnuImmSplit_Click_Error:

          Dim strES As String
          Dim intEL As Integer



70        intEL = Erl
80        strES = Err.Description
90        LogError "frmMain", "mnuImmSplit_Click", intEL, strES


End Sub

Private Sub mnuImmTest_Click()

10        frmAddNewTest.Department = "Immunology"
20        frmAddNewTest.Show 1

End Sub

Private Sub mnuImmTestCodes_Click()

10        With frmListsGeneric
20            .ListType = "IC"
30            .ListTypeName = "Immulogy Test Code"
40            .ListTypeNames = "Immunology Test Codes"
50            .Show 1
60        End With

End Sub

Private Sub mnuIQ200Worklist_Click()
10        frmIQ200Worklist.Show 1
End Sub

Private Sub mnuIsoRep_Click()

10        frmIsolateReport.Show 1

End Sub

Private Sub mnuMain_Click()

10        On Error GoTo mnuMain_Click_Error

20        If UCase(iBOX("Enter Password!", , , True)) = SysOptOptPass(0) Then
30            frmMaintenance.Show 1
40        Else
50            iMsg "Maintenance not available at this time!"
60        End If

70        Exit Sub

mnuMain_Click_Error:

          Dim strES As String
          Dim intEL As Integer



80        intEL = Erl
90        strES = Err.Description
100       LogError "frmMain", "mnuMain_Click", intEL, strES


End Sub


Private Sub mnuMicroAddPanel_Click()

10        With frmExtPanels
20            .Department = "Micro"
30            .Show 1
40        End With

          '    frmMicroExtPanels.Show 1

End Sub

Private Sub mnuMicroAddTest_Click()

10        With frmExtTests
20            .Department = "Micro"
30            .Show 1
40        End With

          'frmMicroExtTests.Show 1

End Sub

Private Sub mnuMicrobiologySub_Click(Index As Integer)

10        On Error GoTo mnuMicrobiologySub_Click_Error

20        If Index = 7 Then
30            frmMicroSetUp.Show 1
40        End If

50        Exit Sub

mnuMicrobiologySub_Click_Error:

          Dim strES As String
          Dim intEL As Integer

60        intEL = Erl
70        strES = Err.Description
80        LogError "frmMain", "mnuMicrobiologySub_Click", intEL, strES

End Sub

Private Sub mnuMicroCandSSub_Click(Index As Integer)

10        On Error GoTo mnuMicroCandSSub_Click_Error

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

mnuMicroCandSSub_Click_Error:

          Dim strES As String
          Dim intEL As Integer

210       intEL = Erl
220       strES = Err.Description
230       LogError "frmMain", "mnuMicroCandSSub_Click", intEL, strES

End Sub

Private Sub mnuMicroExtAddress_Click()

10        frmAddress.Show 1

End Sub

Private Sub mnuMicroFaecesSub_Click(Index As Integer)

10        On Error GoTo mnuMicroFaecesSub_Click_Error

20        Select Case Index
          Case 0:    'XLD
30            With frmListsGeneric
40                .ListType = "FaecesXLD"
50                .ListTypeName = "XLD Entry"
60                .ListTypeNames = "XLD Entries"
70                .Show 1
80            End With

90        Case 1:    'DCA
100           With frmListsGeneric
110               .ListType = "FaecesDCA"
120               .ListTypeName = "DCA Entry"
130               .ListTypeNames = "DCA Entries"
140               .Show 1
150           End With

160       Case 2:    'SMAC
170           With frmListsGeneric
180               .ListType = "FaecesSMAC"
190               .ListTypeName = "SMAC Entry"
200               .ListTypeNames = "SMAC Entries"
210               .Show 1
220           End With

230       Case 3:    'CROMO
240           With frmListsGeneric
250               .ListType = "FaecesCROMO"
260               .ListTypeName = "CROMO Entry"
270               .ListTypeNames = "CROMO Entries"
280               .Show 1
290           End With

300       Case 4:    'cAMP
310           With frmListsGeneric
320               .ListType = "FaecesCAMP"
330               .ListTypeName = "CAMP Entry"
340               .ListTypeNames = "CAMP Entries"
350               .Show 1
360           End With
370       Case 5:    'STEC1
380           With frmListsGeneric
390               .ListType = "FaecesSTEC1"
400               .ListTypeName = "Day1 STEC Entry"
410               .ListTypeNames = "Day1 STEC Entries"
420               .Show 1
430           End With
440       Case 6:    'STEC2
450           With frmListsGeneric
460               .ListType = "FaecesSTEC2"
470               .ListTypeName = "Day2 STEC Entry"
480               .ListTypeNames = "Day2 STEC Entries"
490               .Show 1
500           End With


510       End Select


520       Exit Sub

mnuMicroFaecesSub_Click_Error:

          Dim strES As String
          Dim intEL As Integer

530       intEL = Erl
540       strES = Err.Description
550       LogError "frmMain", "mnuMicroFaecesSub_Click", intEL, strES

End Sub

Private Sub mnuMicroFluidsSub_Click(Index As Integer)

10        On Error GoTo mnuMicroFluidsSub_Click_Error

20        Select Case Index
          Case 0:    'Appearance
30            With frmListsGeneric
40                .ListType = "FA"
50                .ListTypeName = "Fluid Appearance"
60                .ListTypeNames = "Fluid Appearances"
70                .Show 1
80            End With

90        Case 1:    'Cell Count
100           With frmListsGeneric
110               .ListType = "CC"
120               .ListTypeName = "Cell Count"
130               .ListTypeNames = "Cell Counts"
140               .Show 1
150           End With

160       Case 2:    'Gram Stain
170           With frmListsGeneric
180               .ListType = "CG"
190               .ListTypeName = "Gram Stain Result"
200               .ListTypeNames = "Gram Stain Results"
210               .Show 1
220           End With

230       Case 3:    'ZN Stains
240           With frmListsGeneric
250               .ListType = "FluidZN"
260               .ListTypeName = "ZN Stain Result"
270               .ListTypeNames = "ZN Stain Results"
280               .Show 1
290           End With

300       Case 4:    'Leishman Stain
310           With frmListsGeneric
320               .ListType = "CL"
330               .ListTypeName = "Leishman's Stain Result"
340               .ListTypeNames = "Leishman's Stain Results"
350               .Show 1
360           End With

370       Case 5:    'Wet Prep
380           With frmListsGeneric
390               .ListType = "FW"
400               .ListTypeName = "Wet Prep"
410               .ListTypeNames = "Wet Preps"
420               .Show 1
430           End With

440       Case 6:    'Crystals
450           With frmListsGeneric
460               .ListType = "FC"
470               .ListTypeName = "Crystal"
480               .ListTypeNames = "Crystals"
490               .Show 1
500           End With

510       Case 7:    'Sites
520           frmMicroFluidSites.Show 1

530       End Select

540       Exit Sub

mnuMicroFluidsSub_Click_Error:

          Dim strES As String
          Dim intEL As Integer

550       intEL = Erl
560       strES = Err.Description
570       LogError "frmMain", "mnuMicroFluidsSub_Click", intEL, strES

End Sub

Private Sub mnuMicroGeneral_Click()

10        frmMicroGeneral.Show 1

End Sub

Private Sub mnuMicroIdentificationSub_Click(Index As Integer)

10        On Error GoTo mnuMicroIdentificationSub_Click_Error

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

mnuMicroIdentificationSub_Click_Error:

          Dim strES As String
          Dim intEL As Integer

110       intEL = Erl
120       strES = Err.Description
130       LogError "frmMain", "mnuMicroIdentificationSub_Click", intEL, strES

End Sub

Private Sub mnuMicroOrder_Click()

10        frmMicroOrders.Show 1

End Sub

Private Sub mnuMicroStatistics_Click()

10        On Error GoTo mnuMicroStatistics_Click_Error

20        frmMicroStatsGeneral.Show 1

30        Exit Sub

mnuMicroStatistics_Click_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmMain", "mnuMicroStatistics_Click", intEL, strES

End Sub

Private Sub mnuMicroTitlesSub_Click(Index As Integer)

10        On Error GoTo mnuMicroTitlesSub_Click_Error

20        Select Case Index
          Case 0:    'FOB
30            With frmListsGenericColour
40                .ListType = "OccultBlood"
50                .ListTypeName = "Occult Blood Entry"
60                .ListTypeNames = "Occult Blood Entries"
70                .Show 1
80            End With

90        Case 1:    'H. Pylori
100           With frmListsGenericColour
110               .ListType = "HPylori"
120               .ListTypeName = "H. Pylori Entry"
130               .ListTypeNames = "H. Pylori Entries"
140               .Show 1
150           End With

160       Case 2:    'C.Diff Culture
170           With frmListsGenericColour
180               .ListType = "CDiffCulture"
190               .ListTypeName = "C. Diff Culture Entry"
200               .ListTypeNames = "C. Diff Culture Entries"
210               .Show 1
220           End With

230       Case 3:    'C.Diff Toxin/AB
240           With frmListsGenericColour
250               .ListType = "CDiffToxinAB"
260               .ListTypeName = "C. Diff Toxin A/B Entry"
270               .ListTypeNames = "C. Diff Toxin A/B Entries"
280               .Show 1
290           End With

300       Case 4:    'C.Diff GDH
310           With frmListsGenericColour
320               .ListType = "CDiffGDH"
330               .ListTypeName = "C. Diff GDH Entry"
340               .ListTypeNames = "C. Diff GDH Entries"
350               .Show 1
360           End With
370       Case 5:    'C.Diff GDH Detail
380           With frmListsGeneric
390               .ListType = "CDiffGDHDetail"
400               .ListTypeName = "C. Diff GDH Detail Entry"
410               .ListTypeNames = "C. Diff GDH Detail Entries"
420               .Show 1
430           End With

440       Case 6:    'C.Diff PCR
450           With frmListsGenericColour
460               .ListType = "CDiffPCR"
470               .ListTypeName = "C. Diff PCR Entry"
480               .ListTypeNames = "C. Diff PCR Entries"
490               .Show 1
500           End With
510       Case 7:    'C.Diff PCR Detail
520           With frmListsGeneric
530               .ListType = "CDiffPCRDetail"
540               .ListTypeName = "C. Diff PCR Detail Entry"
550               .ListTypeNames = "C. Diff PCR Detail Entries"
560               .Show 1
570           End With

580       Case 8:    'Rota
590           With frmListsGenericColour
600               .ListType = "Rota"
610               .ListTypeName = "Rota Virus Entry"
620               .ListTypeNames = "Rota Virus Entries"
630               .Show 1
640           End With

650       Case 9:    'Adeno
660           With frmListsGenericColour
670               .ListType = "Adeno"
680               .ListTypeName = "Adeno Virus Entry"
690               .ListTypeNames = "Adeno Virus Entries"
700               .Show 1
710           End With

720       Case 10:    'RSV
730           With frmListsGenericColour
740               .ListType = "RSV"
750               .ListTypeName = "RSV Entry"
760               .ListTypeNames = "RSV Entries"
770               .Show 1
780           End With

790       Case 11:    'Cryptosporidium
800           With frmListsGenericColour
810               .ListType = "Crypto"
820               .ListTypeName = "Cryptosporidium Entry"
830               .ListTypeNames = "Cryptosporidium Entries"
840               .Show 1
850           End With

860       Case 12:    'OP Comments
870       Case 13:
880           With frmListsGenericColour
890               .ListType = "Giardia"
900               .ListTypeName = "Giardia Lamblia Entry"
910               .ListTypeNames = "Giardia Lamblia Entries"
920               .Show 1
930           End With

940       End Select

950       Exit Sub

mnuMicroTitlesSub_Click_Error:

          Dim strES As String
          Dim intEL As Integer

960       intEL = Erl
970       strES = Err.Description
980       LogError "frmMain", "mnuMicroTitlesSub_Click", intEL, strES

End Sub

Private Sub mnuMicroUrineSub_Click(Index As Integer)

10        On Error GoTo mnuMicroUrineSub_Click_Error

20        Select Case Index
          Case 0:    'Bacteria
30            With frmListsGeneric
40                .ListType = "BB"
50                .ListTypeName = "Bacteria Entry"
60                .ListTypeNames = "Bacteria Entries"
70                .Show 1
80            End With
90        Case 1:    'WCC
100           With frmListsGenericColour
110               .ListType = "WW"
120               .ListTypeName = "WCC Entry"
130               .ListTypeNames = "WCC Entries"
140               .Show 1
150           End With

160       Case 2:    'RCC
170           With frmListsGenericColour
180               .ListType = "RR"
190               .ListTypeName = "RCC Entry"
200               .ListTypeNames = "RCC Entries"
210               .Show 1
220           End With

230       Case 3:    'Crystals
240           With frmMicroLists
250               .optList(6).Value = True
260               .Show 1
270           End With
280       Case 4:    'Casts
290           With frmMicroLists
300               .optList(5).Value = True
310               .Show 1
320           End With
330       Case 5:    'Misc
340           With frmMicroLists
350               .optList(7).Value = True
360               .Show 1
370           End With
380       Case 7:    'Pregnancy
390           With frmListsGeneric
400               .ListType = "PG"
410               .ListTypeName = "Pregnancy Entry"
420               .ListTypeNames = "Pregnancy Entries"
430               .Show 1
440           End With
450       End Select

460       Exit Sub

mnuMicroUrineSub_Click_Error:

          Dim strES As String
          Dim intEL As Integer

470       intEL = Erl
480       strES = Err.Description
490       LogError "frmMain", "mnuMicroUrineSub_Click", intEL, strES

End Sub

Private Sub mnuMisc_Click()

10        frmLists.Show 1

End Sub


Private Sub mnuMissing_Click()

10        frmMissing.Show 1

End Sub

Private Sub mnuNCRI_Click()

10        frmNcri.Show 1

End Sub

Private Sub mnuNormRanges_Click()

10        frmEndDefaults.Show 1

End Sub

Private Sub mnuOptions_Click()

10        On Error GoTo mnuOptions_Click_Error

20        If UCase(iBOX("Enter Password!", , , True)) = SysOptOptPass(0) Then
30            frmOption.Show 1
40            Check_Ax
50        Else
60            iMsg "Incorrect password"
70        End If

80        Exit Sub

mnuOptions_Click_Error:

          Dim strES As String
          Dim intEL As Integer

90        intEL = Erl
100       strES = Err.Description
110       LogError "frmMain", "mnuOptions_Click", intEL, strES

End Sub

Private Sub mnuOther_Click()

10        frmEditHisto.Show 1

End Sub

Private Sub mnuOut_Click()

10        frmExtOut.Show 1

End Sub

Private Sub mnuOutStand_Click()

10        frmUnfinished.Show 1

End Sub

Private Sub mnuPhoneLog_Click()

10        On Error GoTo mnuPhoneLog_Click_Error

20        frmPhoneLogHistory.Show 1

30        Exit Sub

mnuPhoneLog_Click_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmMain", "mnuPhoneLog_Click", intEL, strES

End Sub

Private Sub mnuPlausiImm_Click()

10        frmImmPlausible.Show 1

End Sub

Private Sub mnuPregnancy_Click()

10        frmBatchPregnancy.Show 1

End Sub

Private Sub mnuPrintPriorities_Click()
10        frmPrintPriorities.Show 1
End Sub



Private Sub mnuSemenConsistencyList_Click()
10        With frmListsGeneric
20            .ListType = "SemenConsistency"
30            .ListTypeName = "Consistency"
40            .ListTypeNames = "Consistencies"
50            .Show 1
60        End With

End Sub

Private Sub mnuSemenCountList_Click()
10        With frmListsGeneric
20            .ListType = "SemenCount"
30            .ListTypeName = "Semen Count"
40            .ListTypeNames = "Semen Counts"
50            .Show 1
60        End With
End Sub

Private Sub mnuSemenTypeList_Click()
10        With frmListsGeneric
20            .ListType = "SemenSpecimenType"
30            .ListTypeName = "Specimen Type"
40            .ListTypeNames = "Specimen Types"
50            .Show 1
60        End With
End Sub

Private Sub mnuSemenVolumeList_Click()
10        With frmListsGeneric
20            .ListType = "SemenVolume"
30            .ListTypeName = "Semen Volume"
40            .ListTypeNames = "Semen Volumes"
50            .Show 1
60        End With
End Sub

Private Sub mnuSiteCount_Click()

10        frmSiteCount.Show 1

End Sub

Private Sub mnuStatFea_Click()

10        frmMicroTotals.Show 1

End Sub

Private Sub mnuStatsCol_Click()

'If UCase(HospName(0)) = "MULLINGAR" Then
    frmStatisticsCollection.Show 1
'End If

End Sub

Private Sub mnuStatsFluids_Click()

10        frmMicroFluidStats.Show 1

End Sub

Private Sub mnuStock_Click()

10        frmReagentLevel.Show 1

End Sub

Private Sub mnuStTotCoag_Click()

10        On Error GoTo mnuStTotCoag_Click_Error

20        With frmTotals
30            .Caption = "Coagulation - Totals"
40            .TotDept = "Coag"
50            .Show 1
60        End With

70        Exit Sub

mnuStTotCoag_Click_Error:

          Dim strES As String
          Dim intEL As Integer

80        intEL = Erl
90        strES = Err.Description
100       LogError "frmMain", "mnuStTotCoag_Click", intEL, strES

End Sub

Private Sub mnuStTotEn_Click()

10        frmEndTotalTests.Show 1

End Sub

Private Sub mnuStTotEndo_Click()

10        On Error GoTo mnuStTotEndo_Click_Error

20        With frmTotals
30            .Caption = "Endocrinology - Totals"
40            .TotDept = "End"
50            .Show 1
60        End With

70        Exit Sub

mnuStTotEndo_Click_Error:

          Dim strES As String
          Dim intEL As Integer

80        intEL = Erl
90        strES = Err.Description
100       LogError "frmMain", "mnuStTotEndo_Click", intEL, strES

End Sub

Private Sub mnuStTotIm_Click()

10        frmImmTotalTests.Show 1

End Sub

Private Sub mnuStTotImm_Click()

10        On Error GoTo mnuStTotImm_Click_Error

20        With frmTotals
30            .Caption = "Immunology - Totals"
40            .TotDept = "Imm"
50            .Show 1
60        End With

70        Exit Sub

mnuStTotImm_Click_Error:

          Dim strES As String
          Dim intEL As Integer



80        intEL = Erl
90        strES = Err.Description
100       LogError "frmMain", "mnuStTotImm_Click", intEL, strES


End Sub

Private Sub mnuSuperStats_Click()

10        frmSuperStats.Show 1

End Sub

Private Sub mnuTestRange_Click()

10        frmBioRanges.Show 1

End Sub

Private Sub mnuTotBio_Click()

10        frmBioTotalTests.Show 1

End Sub

Private Sub mnuUrBat_Click()

10        frmBatchUrine.Show 1

End Sub

Private Sub mnuUrLog_Click()

10        frmBatchLogInUrine.Show 1

End Sub

Private Sub mnuUrnStats_Click()

10        frmUrineStats.Show 1

End Sub

'Private Sub mnuUserForm_Click()
'
'10    On Error GoTo mnuUserForm_Click_Error
'
'20    If Username = "" Then Exit Sub
'
'30    frmFormOpt.Show 1
'
'40    Exit Sub
'
'mnuUserForm_Click_Error:
'
'      Dim strES As String
'      Dim intEL As Integer
'
'
'
'50    intEL = Erl
'60    strES = Err.Description
'70    LogError "frmMain", "mnuUserForm_Click", intEL, strES
'
'
'End Sub

Private Sub mnuUserOpt_Click()

10        On Error GoTo mnuUserOpt_Click_Error

          'If Username = "" Then Exit Sub
20        frmPrintingRules.Show 1


30        Exit Sub

mnuUserOpt_Click_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmMain", "mnuUserOpt_Click", intEL, strES


End Sub

Private Sub mnuViewWards_Click()

10        frmViewWards.Show 1

End Sub

Private Sub mnuWeb_Click()
          Dim sProgramName As String * 255
          Dim retFind&, retShell&, iFileNum%, sRandFileName$

10        On Error GoTo mnuWeb_Click_Error

20        sRandFileName = "x27z42j.html"      ' pick a name that is unlikely to be on disk

30        iFileNum = FreeFile                 ' write the temp file to disk
40        Open App.Path & "\" & sRandFileName For Binary As iFileNum
50        Put iFileNum, , vbNullString
60        Close iFileNum

70        retFind = FindExecutable(sRandFileName, App.Path, sProgramName)

80        Kill App.Path & "\" & sRandFileName    ' get rid of it now that we are done testing...

90        Select Case retFind
          Case 0
100           MsgBox "Sorry out of memory, please close some programs and try again."
110           Exit Sub

120       Case 1 To 30
              ' error (you could have some generic error message here as well.)
130           MsgBox "Error Number: " & retFind
140           Exit Sub

150       Case 31
160           MsgBox "Unable to find Web Browser."
170           Exit Sub

180       Case Is > 31

190           DoEvents    ' without this the label may be white or blank
200           retShell = ShellExecute(frmMain.hWnd, "Open", "http://www.customsoftware.ie", "", "", SW_SHOWNORMAL)


210       End Select

220       Exit Sub

mnuWeb_Click_Error:

          Dim strES As String
          Dim intEL As Integer



230       intEL = Erl
240       strES = Err.Description
250       LogError "frmMain", "mnuWeb_Click", intEL, strES


End Sub


Private Sub morder_Click()

10        On Error GoTo morder_Click_Error

20        With frmNewOrder
30            .FromEdit = False
40            .Show 1
50        End With

60        Exit Sub

morder_Click_Error:

          Dim strES As String
          Dim intEL As Integer



70        intEL = Erl
80        strES = Err.Description
90        LogError "frmMain", "morder_Click", intEL, strES


End Sub

Private Sub mPanelBarCodes_Click()

10        frmPanelsBarcodeSelection.Show 1

End Sub

Private Sub mpanels_Click()

10        frmPanels.Department = "Biochemistry"
20        frmPanels.Show 1

End Sub

Private Sub mpf_Click()

'fdp.Show 1

End Sub

Private Sub mPrinters_Click()

10        frmPrinters.Show 1

End Sub

Private Sub mpseq_Click()

'fpseq.Show 1

End Sub

Private Sub mqclimits_Click()

10        frmLimits.Show 1

End Sub

Private Sub mqcview_Click()

10        On Error GoTo mqcview_Click_Error

20        frmQCparent.optBio.Value = True
30        frmQCparent.Show

40        Exit Sub

mqcview_Click_Error:

          Dim strES As String
          Dim intEL As Integer



50        intEL = Erl
60        strES = Err.Description
70        LogError "frmMain", "mqcview_Click", intEL, strES


End Sub

Private Sub mResetLastUsed_Click()

          Dim LU As String
          Dim NewLU As String


10        On Error GoTo mResetLastUsed_Click_Error

20        LU = GetSetting("NetAcquire", "StartUp", "LastUsed", "1")

30        NewLU = iBOX("Enter new 'Last Used' Number.", , LU)

40        If Val(NewLU) > 0 Then
50            SaveSetting "NetAcquire", "StartUp", "LastUsed", Format$(Val(NewLU))
60            iMsg "'Last Used' Number changed to " & Format$(Val(NewLU)), vbInformation
70        Else
80            iMsg "'Last Used' Number not changed!", vbExclamation
90        End If



100       Exit Sub

mResetLastUsed_Click_Error:

          Dim strES As String
          Dim intEL As Integer



110       intEL = Erl
120       strES = Err.Description
130       LogError "frmMain", "mResetLastUsed_Click", intEL, strES


End Sub

Private Sub msearchmore_Click(Index As Integer)

10        On Error GoTo msearchmore_Click_Error

20        With frmPatHistoryNew
30            .oFor(Index) = True
40            .FromEdit = False
50            .Show 1
60        End With

70        Exit Sub

msearchmore_Click_Error:

          Dim strES As String
          Dim intEL As Integer



80        intEL = Erl
90        strES = Err.Description
100       LogError "frmMain", "msearchmore_Click", intEL, strES


End Sub

Private Sub msearchmore1_Click(Index As Integer)


10        On Error GoTo msearchmore1_Click_Error

20        With frmPatHistoryNew
30            .oFor(Index) = True
40            .FromEdit = False
50            .Show 1
60        End With

70        Exit Sub

msearchmore1_Click_Error:

          Dim strES As String
          Dim intEL As Integer



80        intEL = Erl
90        strES = Err.Description
100       LogError "frmMain", "msearchmore1_Click", intEL, strES


End Sub

Private Sub mSetSourceNames_Click()

10        frmSetSources.Show 1

End Sub

Private Sub mshowerror_Click()

10        frmSystemErrorLog.Show 1

End Sub

Private Sub mstock_Click()

'fstock.Show 1

End Sub

Private Sub mtechnical_Click()

          Dim sProgramName As String * 255
          Dim retFind&, retShell&, iFileNum%, sRandFileName$

10        On Error GoTo mtechnical_Click_Error

20        sRandFileName = "x27z42j.html"      ' pick a name that is unlikely to be on disk

30        iFileNum = FreeFile                 ' write the temp file to disk
40        Open App.Path & "\" & sRandFileName For Binary As iFileNum
50        Put iFileNum, , vbNullString
60        Close iFileNum

70        retFind = FindExecutable(sRandFileName, App.Path, sProgramName)

80        Kill App.Path & "\" & sRandFileName    ' get rid of it now that we are done testing...

90        Select Case retFind
          Case 0
100           MsgBox "Sorry out of memory, please close some programs and try again."
110           Exit Sub

120       Case 1 To 30
              ' error (you could have some generic error message here as well.)
130           MsgBox "Error Number: " & retFind
140           Exit Sub

150       Case 31
160           MsgBox "Unable to find Web Browser."
170           Exit Sub

180       Case Is > 31

190           DoEvents    ' without this the label may be white or blank
200           retShell = ShellExecute(frmMain.hWnd, "Open", "http://www.customsoftware.ie/contact.html", "", "", SW_SHOWNORMAL)


210       End Select
220       Exit Sub


230       Exit Sub

mtechnical_Click_Error:

          Dim strES As String
          Dim intEL As Integer



240       intEL = Erl
250       strES = Err.Description
260       LogError "frmMain", "mtechnical_Click", intEL, strES


End Sub



Private Sub mtotbio_Click()

10        On Error GoTo mtotbio_Click_Error

20        With frmTotals
30            .Caption = "Biochemistry - Totals"
40            .TotDept = "Bio"
50            .Show 1
60        End With

70        Exit Sub

mtotbio_Click_Error:

          Dim strES As String
          Dim intEL As Integer



80        intEL = Erl
90        strES = Err.Description
100       LogError "frmMain", "mtotbio_Click", intEL, strES


End Sub

Private Sub mtothaem_Click()

10        frmTotHaem.Show 1

End Sub

Private Sub mViewStats_Click()

10        frmStatSources.Show 1

End Sub

Private Sub mWards_Click()

10        frmWardList.Show 1

End Sub

Private Sub mworklist_Click()

10        frmDaily.Show 1

End Sub


Private Sub TimerBar_Timer()


10        On Error GoTo TimerBar_Timer_Error

20        If mLogOff.Enabled Then
30            StatusBar1.Panels(3).Text = UserName
40            pbCounter = pbCounter + 1
50            If pbCounter < pb.Max Then
60                pb = pbCounter
70            Else
80                EnableMenus False
90                pbCounter = 0
100               pb = 0
110           End If
120       Else
130           StatusBar1.Panels(3).Text = ""
140       End If


150       Exit Sub

TimerBar_Timer_Error:

          Dim strES As String
          Dim intEL As Integer



160       intEL = Erl
170       strES = Err.Description
180       LogError "frmMain", "TimerBar_Timer", intEL, strES


End Sub

Private Sub timerChk_Timer()

          Static Counter As Long

          'Timer interval = 10 secs

10        Counter = Counter + 1
20        If Counter > 360 Then    ' = 1 hour
30            Counter = 0
40            LoadOptions
50        End If

End Sub

Private Sub tmrNotPrinted_Timer()

          Dim tb As New Recordset
          Dim sql As String
          Dim s As String
          Static h As Long
          Dim Found As Boolean

10        On Error GoTo tmrNotPrinted_Timer_Error

          'If IsIDE Then Exit Sub

20        Found = False

          'sql = "DELETE from haemresults WHERE sampleid > 9990000"
          'Cnxn(0).Execute sql

30        frmMainCounter = frmMainCounter + 1
40        frmMainImageCounter = frmMainImageCounter + 1
50        If frmMainImageCounter > 4 Then
60            frmMainImageCounter = 1

70            If Format$("04/05/2001", "dd/mmm/yyyy") <> "04/May/2001" Then
80                MsgBox "Date/Time Format in" & vbCrLf & _
                         "International Settings" & vbCrLf & _
                         "is not set correctly." & vbCrLf & vbCrLf & _
                         "Cannot proceed!", vbCritical
90                End
100           End If
110       End If

120       Image1.Picture = ImageList1.ListImages(frmMainImageCounter).Picture

          'If frmMainCounter < 5 Then
          '    Exit Sub
          'End If
130       frmMainCounter = 0

140       If chkAutoRefresh <> 1 Then Exit Sub

150       h = h + 1
160       If h = 6 Then h = 1

170       If h = 1 Then
180           With gBioNoResults
190               .Rows = 2
200               .AddItem ""
210               .RemoveItem 1
220               .Visible = False
230           End With
240           sql = "SELECT distinct SampleID, AnalyserID from BioRequests Where DateTime > getdate() - " & Val(cmbRefreshDays)
250           Set tb = New Recordset
260           Set tb = Cnxn(0).Execute(sql)
270           Do While Not tb.EOF
280               Select Case Trim$(tb!AnalyserID & "")
                  Case "4": s = "Immuno"
290               Case "A": s = "Bio (A)"
300               Case "B": s = "Bio (B)"
310               Case Else: s = "General"
320               End Select
330               gBioNoResults.AddItem s & vbTab & tb!SampleID & ""
340               tb.MoveNext
350           Loop
360           With gBioNoResults
370               If .Rows > 2 Then
380                   .RemoveItem 1
390               End If
400           End With

410           gBioNoResults.Visible = True

420           With gBioNotPrinted
430               .Rows = 2
440               .AddItem ""
450               .RemoveItem 1
460               .Visible = False
470           End With
480           sql = "SELECT DISTINCT TOP 100 R.SampleID, R.Analyser FROM BioResults R " & _
                    "INNER JOIN BioTestDefinitions T ON R.Code = T.Code " & _
                    "WHERE COALESCE(Printed, 0) = 0 And RunTime > getdate() - 14 And COALESCE(T.Printable, 0) = 1"

              '    sql = "SELECT TOP 100 T.* " & _
                   '          "FROM (SELECT DISTINCT SampleID, Analyser " & _
                   '          "      FROM BioResults WHERE " & _
                   '          "      COALESCE(Printed, 0) = 0 And RunTime > getdate() - " & Val(cmbRefreshDays) & " ) T"
490           Set tb = New Recordset
500           RecOpenServer 0, tb, sql
510           Do While Not tb.EOF
520               Select Case Trim$(tb!Analyser & "")
                  Case "4": s = "Immuno"
530               Case "A": s = "Bio (A)"
540               Case "B": s = "Bio (B)"
550               Case "P1": s = SysOptBioN1(0)
560               Case "P2": s = SysOptBioN2(0)
570               Case Else: s = "General"
580               End Select
590               gBioNotPrinted.AddItem s & vbTab & tb!SampleID & ""
600               tb.MoveNext
610           Loop
620           With gBioNotPrinted
630               If .Rows > 2 Then
640                   .RemoveItem 1
650               End If
660               .Visible = True
670           End With

680       ElseIf h = 2 Then
690           sql = "SELECT distinct Top 100 SampleID, RunDateTime from HaemResults WHERE " & _
                    "COALESCE(Printed, 0) = 0  And RunDateTime > getdate() - " & Val(cmbRefreshDays) & " " & _
                    "Order By RunDateTime Desc"
700           Set tb = New Recordset
710           Set tb = Cnxn(0).Execute(sql)
720           lstHaemNotPrinted.Clear
730           Do While Not tb.EOF
740               lstHaemNotPrinted.AddItem tb!SampleID & ""
750               tb.MoveNext
760           Loop
770           tb.Close

780       ElseIf h = 3 Then
790           sql = "SELECT distinct top 100 SampleID from CoagRequests Order By SampleID Desc"
800           Set tb = New Recordset
810           Set tb = Cnxn(0).Execute(sql)
820           lstCoagNoResults.Clear
830           Do While Not tb.EOF
840               lstCoagNoResults.AddItem tb!SampleID & ""
850               tb.MoveNext
860           Loop
870           tb.Close

              '    sql = "SELECT DISTINCT Top 100 SampleID, Max(RunTime) MS FROM CoagResults WHERE " & _
                   '          "SampleID NOT IN " & _
                   '          "  (SELECT DISTINCT SampleID FROM CoagResults WHERE Printed = 1 AND RunTime > getdate() - " & Val(cmbRefreshDays) & ") " & _
                   '          "AND SampleID > '20000' AND RunTime > getdate() - " & Val(cmbRefreshDays) & " " & _
                   '          "Group By SampleID " & _
                   '          "Order By MS Desc"
880           sql = "SELECT  SampleID, MAX(RunTime) MS FROM CoagResults " & _
                    "WHERE SampleID > '20000' AND RunTime > getdate() - " & Val(cmbRefreshDays) & " " & _
                    "GROUP BY SampleID " & _
                    "HAVING SUM(COALESCE(Printed, 0)) = 0 " & _
                    "ORDER BY MS DESC"

890           Set tb = New Recordset
900           Set tb = Cnxn(0).Execute(sql)
910           lstCoagNotPrinted.Clear
920           Do While Not tb.EOF
930               lstCoagNotPrinted.AddItem tb!SampleID & ""
940               tb.MoveNext
950           Loop
960       End If

970       If SysOptDeptImm(0) = True And h = 4 Then
980           sql = "SELECT distinct top 100 SampleID, DateTime from immRequests Order By DateTime Desc"
990           Set tb = New Recordset
1000          Set tb = Cnxn(0).Execute(sql)
1010          lstImmNoResults.Clear
1020          Do While Not tb.EOF
1030              lstImmNoResults.AddItem tb!SampleID & ""
1040              tb.MoveNext
1050          Loop
1060          tb.Close

1070          sql = "SELECT distinct Top 100 SampleID, RunTime from immResults WHERE " & _
                    "COALESCE(Printed, 0) = 0 AND RunTime > getdate() - " & Val(cmbRefreshDays) & " " & _
                    "Order By RunTime Desc"
1080          Set tb = New Recordset
1090          Set tb = Cnxn(0).Execute(sql)
1100          lstImmNotPrinted.Clear
1110          Do While Not tb.EOF
1120              lstImmNotPrinted.AddItem tb!SampleID & ""
1130              tb.MoveNext
1140          Loop
1150      End If

1160      If SysOptDeptEnd(0) = True And h = 5 Then
1170          sql = "SELECT distinct Top 100 SampleID, DateTime from EndRequests Order By DateTime Desc"
1180          Set tb = New Recordset
1190          Set tb = Cnxn(0).Execute(sql)
1200          lstEndNoResults.Clear
1210          Do While Not tb.EOF
1220              lstEndNoResults.AddItem tb!SampleID & ""
1230              tb.MoveNext
1240          Loop
1250          tb.Close

1260          sql = "SELECT distinct Top 100 SampleID, RunTime from EndResults WHERE " & _
                    "COALESCE(Printed, 0) = 0 AND RunTime > getdate() - " & Val(cmbRefreshDays) & " " & _
                    "Order By RunTime Desc"
1270          Set tb = New Recordset
1280          Set tb = Cnxn(0).Execute(sql)
1290          lstEndNotPrinted.Clear
1300          Do While Not tb.EOF
1310              lstEndNotPrinted.AddItem tb!SampleID & ""
1320              tb.MoveNext
1330          Loop
1340      End If

1350      Exit Sub

tmrNotPrinted_Timer_Error:

          Dim strES As String
          Dim intEL As Integer

1360      intEL = Erl
1370      strES = Err.Description
1380      LogError "frmMain", "tmrNotPrinted_Timer", intEL, strES, sql

End Sub

Private Sub tmrRefresh_Timer()

          Static Counter As Long
          Dim sql As String

          'tmrRefresh.Interval set to 30 seconds

10        On Error GoTo tmrRefresh_Timer_Error

20        Counter = Counter + 1

30        If Counter = 30 Then    '15 minutes
              'CheckPAS
40        ElseIf Counter >= 60 Then    '30 minutes
50            Counter = 0
              'CheckPAS
              ' colHaemTestDefinitions.Refresh
              
              'If UCase$(HospName(0)) = "PORTLAOISE" Then
                  'sql = "DELETE from haemresults WHERE sampleid > 9990036"
                  'Cnxn(0).Execute sql
              'End If
60        End If

70        Exit Sub

tmrRefresh_Timer_Error:

          Dim strES As String
          Dim intEL As Integer



80        intEL = Erl
90        strES = Err.Description
100       LogError "frmMain", "tmrRefresh_Timer", intEL, strES, sql

End Sub


Private Sub tmrUrgent_Timer()

          Dim tb As New Recordset
          Dim sql As String
          Dim rs As Recordset
          Dim Found As Boolean
          Dim n As Long

10        On Error GoTo tmrUrgent_Timer_Error

20        DoEvents

30        If SysOptUrgent(0) Then
40            ClearFGrid grdUrg
50            sql = "UPDATE Demographics SET Urgent = 0 WHERE " & _
                    "RecDate < '" & Format(Now - SysOptUrgentRef(0), "dd/MMM/yyyy hh:mm:ss") & "' " & _
                    "AND Urgent = 1"
60            Cnxn(0).Execute sql
70            sql = "SELECT * FROM Demographics WHERE Urgent = 1 " & _
                    "AND (RecDate > '" & Format(Now - SysOptUrgentRef(0), "dd/MMM/yyyy hh:mm:ss") & "' " & _
                    "     OR RecDate IS NULL " & _
                    "     OR RecDate <> '') " & _
                    "ORDER BY RunDate desc"
80            Set tb = New Recordset
90            Set tb = Cnxn(0).Execute(sql)
100           Do While Not tb.EOF
110               grdUrg.AddItem tb!SampleID
120               grdUrg.Row = grdUrg.Rows - 1
130               sql = "SELECT valid from haemresults WHERE sampleid = '" & tb!SampleID & "'"
140               Set rs = New Recordset
150               RecOpenServer 0, rs, sql
160               If rs.EOF Then
170                   sql = "SELECT sampleid as valid from haemrequests WHERE sampleid = '" & tb!SampleID & "'"
180                   Set rs = New Recordset
190                   RecOpenServer 0, rs, sql
200               End If
210               If Not rs.EOF Then
220                   If rs!Valid <> True Then
230                       grdUrg.Col = 1
240                       grdUrg.CellBackColor = vbRed
250                   End If
260               Else
270                   grdUrg.Col = 1
280                   grdUrg.CellBackColor = vbWhite
290               End If
300               sql = "SELECT valid from bioresults WHERE sampleid = '" & tb!SampleID & "'"
310               Set rs = New Recordset
320               RecOpenServer 0, rs, sql
330               If rs.EOF Then
340                   sql = "SELECT sampleid as valid from biorequests WHERE sampleid = '" & tb!SampleID & "'"
350                   Set rs = New Recordset
360                   RecOpenServer 0, rs, sql
370               End If
380               If Not rs.EOF Then
390                   Do While Not rs.EOF
400                       If rs!Valid <> 1 Then
410                           grdUrg.Col = 2
420                           grdUrg.CellBackColor = vbRed
430                       End If
440                       rs.MoveNext
450                   Loop
460               Else
470                   grdUrg.Col = 2
480                   grdUrg.CellBackColor = vbWhite
490               End If
500               sql = "SELECT valid from coagresults WHERE sampleid = '" & tb!SampleID & "'"
510               Set rs = New Recordset
520               RecOpenServer 0, rs, sql
530               If rs.EOF Then
540                   sql = "SELECT sampleid as valid from coagrequests WHERE sampleid = '" & tb!SampleID & "'"
550                   Set rs = New Recordset
560                   RecOpenServer 0, rs, sql
570               End If
580               If Not rs.EOF Then
590                   Do While Not rs.EOF
600                       If rs!Valid <> True Then
610                           grdUrg.Col = 3
620                           grdUrg.CellBackColor = vbRed
630                       End If
640                       rs.MoveNext
650                   Loop
660               Else
670                   grdUrg.Col = 3
680                   grdUrg.CellBackColor = vbWhite
690               End If
700               If SysOptDeptEnd(0) Then
710                   sql = "SELECT valid from endresults WHERE sampleid = '" & tb!SampleID & "'"
720                   Set rs = New Recordset
730                   RecOpenServer 0, rs, sql
740                   If rs.EOF Then
750                       sql = "SELECT sampleid as valid from endrequests WHERE sampleid = '" & tb!SampleID & "'"
760                       Set rs = New Recordset
770                       RecOpenServer 0, rs, sql
780                   End If
790                   If Not rs.EOF Then
800                       Do While Not rs.EOF
810                           If rs!Valid <> 1 Then
820                               grdUrg.Col = 4
830                               grdUrg.CellBackColor = vbRed
840                           End If
850                           rs.MoveNext
860                       Loop
870                   Else
880                       grdUrg.Col = 4
890                       grdUrg.CellBackColor = vbWhite
900                   End If
910               End If

920               If SysOptDeptBga(0) Then
930                   sql = "SELECT valid from bgaresults WHERE sampleid = '" & tb!SampleID & "'"
940                   Set rs = New Recordset
950                   RecOpenServer 0, rs, sql
960                   If Not rs.EOF Then
970                       Do While Not rs.EOF
980                           If rs!Valid <> True Then
990                               grdUrg.Col = 5
1000                              grdUrg.CellBackColor = vbRed
1010                          End If
1020                          rs.MoveNext
1030                      Loop
1040                  Else
1050                      grdUrg.Col = 5
1060                      grdUrg.CellBackColor = vbWhite
1070                  End If
1080              End If
1090              If SysOptDeptImm(0) Then
1100                  sql = "SELECT valid from immresults WHERE sampleid = '" & tb!SampleID & "'"
1110                  Set rs = New Recordset
1120                  RecOpenServer 0, rs, sql
1130                  If rs.EOF Then
1140                      sql = "SELECT sampleid as valid from IMmrequests WHERE sampleid = '" & tb!SampleID & "'"
1150                      Set rs = New Recordset
1160                      RecOpenServer 0, rs, sql
1170                  End If
1180                  If Not rs.EOF Then
1190                      Do While Not rs.EOF
1200                          If rs!Valid <> 1 Then
1210                              grdUrg.Col = 6
1220                              grdUrg.CellBackColor = vbRed
1230                          End If
1240                          rs.MoveNext
1250                      Loop
1260                  Else
1270                      grdUrg.Col = 6
1280                      grdUrg.CellBackColor = vbWhite
1290                  End If
1300              End If

1310              For n = 1 To 6
1320                  grdUrg.Col = n
1330                  If grdUrg.CellBackColor = vbRed Then
1340                      Found = True
1350                  End If
1360              Next
1370              If Found = False Then grdUrg.RemoveItem grdUrg.Rows - 1
1380              Found = False
1390              tb.MoveNext
1400          Loop

1410          FixG grdUrg
1420      End If

1430      Exit Sub

tmrUrgent_Timer_Error:

          Dim strES As String
          Dim intEL As Integer



1440      intEL = Erl
1450      strES = Err.Description
1460      LogError "frmMain", "tmrUrgent_Timer", intEL, strES, sql

End Sub
