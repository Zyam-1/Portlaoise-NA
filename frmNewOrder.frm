VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmNewOrder 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Test Order"
   ClientHeight    =   10590
   ClientLeft      =   135
   ClientTop       =   465
   ClientWidth     =   14610
   Icon            =   "frmNewOrder.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10590
   ScaleWidth      =   14610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox LstHaePanel 
      Height          =   2205
      Left            =   13140
      MultiSelect     =   1  'Simple
      TabIndex        =   64
      Top             =   2070
      Width           =   1305
   End
   Begin VB.ListBox lEndoPanel 
      Height          =   1425
      Left            =   10395
      MultiSelect     =   1  'Simple
      TabIndex        =   43
      Top             =   1935
      Width           =   1275
   End
   Begin VB.Frame Frame81 
      Caption         =   "Endocrinology Comments"
      Height          =   1455
      Index           =   0
      Left            =   9090
      TabIndex        =   61
      Top             =   8955
      Width           =   3975
      Begin VB.TextBox txtImmComment 
         BackColor       =   &H80000018&
         Height          =   1140
         Index           =   0
         Left            =   90
         MaxLength       =   320
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   62
         Top             =   270
         Width           =   3765
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Biochemistry Comments"
      Height          =   1530
      Left            =   1665
      TabIndex        =   59
      Top             =   8865
      Width           =   4470
      Begin VB.TextBox txtBioComment 
         BackColor       =   &H80000018&
         Height          =   1095
         Left            =   45
         MaxLength       =   560
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   60
         Tag             =   "Biochemistry Comment"
         ToolTipText     =   "Only 360 Characters"
         Top             =   315
         Width           =   4335
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   120
      TabIndex        =   51
      Top             =   1320
      Width           =   6375
      Begin VB.CheckBox chkGBottle 
         Caption         =   "Glucose bottle is in use"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2340
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   15
         Value           =   1  'Checked
         Width           =   3855
      End
      Begin VB.OptionButton optPlasma 
         Alignment       =   1  'Right Justify
         Caption         =   "Plasma"
         Height          =   195
         Left            =   120
         TabIndex        =   53
         Top             =   135
         Width           =   915
      End
      Begin VB.OptionButton optSerum 
         Caption         =   "Serum"
         Height          =   195
         Left            =   1080
         TabIndex        =   52
         Top             =   135
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.CheckBox chkADM 
      Caption         =   "ADM"
      Height          =   420
      Left            =   11580
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   420
      Visible         =   0   'False
      Width           =   2865
   End
   Begin VB.ListBox lstFluid 
      Height          =   3270
      IntegralHeight  =   0   'False
      Left            =   7785
      MultiSelect     =   1  'Simple
      TabIndex        =   45
      Top             =   5580
      Width           =   1230
   End
   Begin VB.ListBox lImmunoPanel 
      Height          =   1425
      Left            =   9090
      MultiSelect     =   1  'Simple
      TabIndex        =   42
      Top             =   1935
      Width           =   1230
   End
   Begin VB.CheckBox chkUrgent 
      BackColor       =   &H000000FF&
      Caption         =   "Urgent"
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
      Left            =   2385
      TabIndex        =   40
      Top             =   765
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Frame Frame1 
      Caption         =   "Test Names"
      Height          =   705
      Left            =   8820
      TabIndex        =   34
      Top             =   120
      Width           =   1155
      Begin VB.OptionButton optLong 
         Caption         =   "Long"
         Height          =   195
         Left            =   270
         TabIndex        =   36
         Top             =   240
         Value           =   -1  'True
         Width           =   705
      End
      Begin VB.OptionButton optShort 
         Caption         =   "Short"
         Height          =   195
         Left            =   270
         TabIndex        =   35
         Top             =   450
         Width           =   675
      End
   End
   Begin VB.ListBox lstHaem 
      Height          =   3765
      ItemData        =   "frmNewOrder.frx":030A
      Left            =   13140
      List            =   "frmNewOrder.frx":0311
      MultiSelect     =   1  'Simple
      TabIndex        =   32
      Top             =   5040
      Width           =   1350
   End
   Begin VB.ListBox lstImmunoTests 
      Height          =   5070
      IntegralHeight  =   0   'False
      Left            =   9090
      MultiSelect     =   1  'Simple
      TabIndex        =   26
      Top             =   3765
      Width           =   1230
   End
   Begin VB.ListBox lstCoag 
      Height          =   6420
      IntegralHeight  =   0   'False
      Left            =   11730
      MultiSelect     =   1  'Simple
      TabIndex        =   24
      Top             =   2400
      Width           =   1350
   End
   Begin VB.Timer TimerBar 
      Interval        =   1000
      Left            =   12660
      Top             =   930
   End
   Begin VB.OptionButton oSorF 
      Caption         =   "Fasting"
      Height          =   225
      Index           =   1
      Left            =   1350
      TabIndex        =   22
      Top             =   780
      Width           =   975
   End
   Begin VB.OptionButton oSorF 
      Alignment       =   1  'Right Justify
      Caption         =   "Random"
      Height          =   225
      Index           =   0
      Left            =   420
      TabIndex        =   21
      Top             =   780
      Value           =   -1  'True
      Width           =   915
   End
   Begin VB.ListBox lCSFTests 
      Height          =   2160
      IntegralHeight  =   0   'False
      Left            =   7770
      MultiSelect     =   1  'Simple
      TabIndex        =   9
      Top             =   2760
      Width           =   1230
   End
   Begin VB.ListBox lUrinePanel 
      Height          =   840
      Left            =   6150
      MultiSelect     =   1  'Simple
      TabIndex        =   7
      Top             =   2760
      Width           =   1545
   End
   Begin VB.ListBox lUrineTests 
      Height          =   4980
      IntegralHeight  =   0   'False
      Left            =   6150
      MultiSelect     =   1  'Simple
      TabIndex        =   8
      Top             =   3870
      Width           =   1545
   End
   Begin VB.CommandButton bsave 
      Appearance      =   0  'Flat
      Caption         =   "&Save"
      Height          =   720
      Left            =   4185
      Picture         =   "frmNewOrder.frx":031E
      Style           =   1  'Graphical
      TabIndex        =   2
      Tag             =   "save"
      Top             =   0
      Width           =   1305
   End
   Begin VB.TextBox tSampleID 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   60
      MaxLength       =   8
      TabIndex        =   0
      Top             =   270
      Width           =   1485
   End
   Begin VB.TextBox tinput 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2700
      TabIndex        =   1
      Top             =   270
      Width           =   1365
   End
   Begin VB.CommandButton bcancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   720
      Left            =   6705
      Picture         =   "frmNewOrder.frx":0628
      Style           =   1  'Graphical
      TabIndex        =   4
      Tag             =   "cancel"
      Top             =   0
      Width           =   1065
   End
   Begin VB.CommandButton bclear 
      Appearance      =   0  'Flat
      Caption         =   "Cle&ar"
      Height          =   720
      Left            =   5565
      Picture         =   "frmNewOrder.frx":0932
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Width           =   1065
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   135
      Left            =   45
      TabIndex        =   23
      Top             =   10440
      Width           =   14520
      _ExtentX        =   25612
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   1
      Max             =   600
      Scrolling       =   1
   End
   Begin ComCtl2.UpDown UpDown1 
      Height          =   465
      Left            =   1575
      TabIndex        =   48
      Top             =   270
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   820
      _Version        =   327681
      Value           =   1
      BuddyControl    =   "tSampleID"
      BuddyDispid     =   196635
      OrigLeft        =   1920
      OrigTop         =   540
      OrigRight       =   2160
      OrigBottom      =   1020
      Max             =   99999999
      Min             =   1
      Orientation     =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   0
      Enabled         =   -1  'True
   End
   Begin VB.ListBox lSerumPanel 
      Height          =   6360
      IntegralHeight  =   0   'False
      Left            =   90
      MultiSelect     =   1  'Simple
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   2490
      Width           =   1485
   End
   Begin VB.ListBox lSerumTests 
      BackColor       =   &H00FFFFFF&
      Columns         =   3
      Height          =   6315
      IntegralHeight  =   0   'False
      Left            =   1665
      MultiSelect     =   1  'Simple
      TabIndex        =   6
      Top             =   2475
      Width           =   4470
   End
   Begin VB.ListBox lstEndoTests 
      Height          =   5070
      IntegralHeight  =   0   'False
      Left            =   10395
      MultiSelect     =   1  'Simple
      TabIndex        =   37
      Top             =   3765
      Width           =   1275
   End
   Begin VB.ListBox lEndoPanelPlasma 
      BackColor       =   &H00C0FFFF&
      Height          =   1425
      Left            =   10395
      MultiSelect     =   1  'Simple
      TabIndex        =   57
      Top             =   1935
      Width           =   1275
   End
   Begin VB.ListBox lEndoTestsPlasma 
      BackColor       =   &H00C0FFFF&
      Height          =   5070
      IntegralHeight  =   0   'False
      Left            =   10395
      MultiSelect     =   1  'Simple
      TabIndex        =   58
      Top             =   3765
      Width           =   1275
   End
   Begin VB.ListBox lPlasmaTests 
      BackColor       =   &H00C0FFFF&
      Columns         =   3
      Height          =   6315
      IntegralHeight  =   0   'False
      Left            =   1665
      MultiSelect     =   1  'Simple
      TabIndex        =   55
      Top             =   2490
      Width           =   4470
   End
   Begin VB.ListBox lPlasmaPanel 
      BackColor       =   &H00C0FFFF&
      Height          =   6360
      IntegralHeight  =   0   'False
      Left            =   90
      MultiSelect     =   1  'Simple
      Sorted          =   -1  'True
      TabIndex        =   56
      Top             =   2490
      Width           =   1485
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Haematology"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   300
      Index           =   2
      Left            =   13140
      TabIndex        =   66
      Top             =   1320
      Width           =   1350
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Extended IPU"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   315
      Left            =   13140
      TabIndex        =   65
      Top             =   1680
      Width           =   1350
   End
   Begin VB.Label ACLTop500 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ACL Top 500"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   11760
      TabIndex        =   63
      Top             =   1680
      Width           =   1350
   End
   Begin VB.Label lblCoagAnalyserName 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Maureen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   420
      Left            =   10020
      TabIndex        =   50
      Top             =   420
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fluids"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   300
      Index           =   3
      Left            =   7800
      TabIndex        =   47
      Top             =   5010
      Width           =   1215
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tests"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   5
      Left            =   7800
      TabIndex        =   46
      Top             =   5310
      Width           =   1215
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Panels"
      ForeColor       =   &H00004080&
      Height          =   255
      Left            =   10395
      TabIndex        =   44
      Top             =   1665
      Width           =   1290
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Panels"
      ForeColor       =   &H00004080&
      Height          =   255
      Left            =   9090
      TabIndex        =   41
      Top             =   1665
      Width           =   1260
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Endo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   300
      Index           =   2
      Left            =   10395
      TabIndex        =   39
      Top             =   1320
      Width           =   1275
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tests"
      ForeColor       =   &H00004080&
      Height          =   255
      Index           =   4
      Left            =   10395
      TabIndex        =   38
      Top             =   3435
      Width           =   1275
   End
   Begin VB.Label lAnalyserID 
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1215
      Left            =   7890
      TabIndex        =   33
      Top             =   -180
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tests"
      ForeColor       =   &H000040C0&
      Height          =   255
      Index           =   3
      Left            =   13140
      TabIndex        =   31
      Top             =   4710
      Width           =   1350
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Haematology"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   300
      Index           =   1
      Left            =   13140
      TabIndex        =   30
      Top             =   4410
      Width           =   1350
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tests"
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   2
      Left            =   11730
      TabIndex        =   29
      Top             =   2070
      Width           =   1350
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tests"
      ForeColor       =   &H00004080&
      Height          =   255
      Index           =   1
      Left            =   9090
      TabIndex        =   28
      Top             =   3435
      Width           =   1230
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Immuno"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   300
      Index           =   1
      Left            =   9090
      TabIndex        =   27
      Top             =   1320
      Width           =   1230
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Coagulation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   300
      Index           =   0
      Left            =   11730
      TabIndex        =   25
      Top             =   1320
      Width           =   1350
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tests"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   0
      Left            =   7785
      TabIndex        =   10
      Top             =   2490
      Width           =   1215
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tests"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   6150
      TabIndex        =   11
      Top             =   3600
      Width           =   1560
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Panels"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   6150
      TabIndex        =   20
      Top             =   2490
      Width           =   1560
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tests"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1650
      TabIndex        =   19
      Top             =   2190
      Width           =   4485
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Panels"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   90
      TabIndex        =   18
      Top             =   2190
      Width           =   1470
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CSF"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   300
      Index           =   0
      Left            =   7785
      TabIndex        =   17
      Top             =   2190
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Urine"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   6150
      TabIndex        =   16
      Top             =   2190
      Width           =   1560
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Biochemistry"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   90
      TabIndex        =   15
      Top             =   1860
      Width           =   8880
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Sample Number"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   90
      TabIndex        =   14
      Top             =   60
      Width           =   1335
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Test Code"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   2700
      TabIndex        =   13
      Top             =   60
      Width           =   735
   End
   Begin VB.Label lstatus 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4185
      TabIndex        =   12
      Top             =   780
      Width           =   3615
   End
End
Attribute VB_Name = "frmNewOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mFromEdit As Boolean

Private mSampleID As String

Private CoagChanged As Boolean
Private HaemChanged As Boolean
Private BioChanged As Boolean
Private ImmunoChanged As Boolean
Private EndoChanged As Boolean
Private HaeChanged As Boolean

Private AnalyserID As String

Private Activated As Boolean

Private Type udtBarCode
    BarCodeType As String
    Name As String
    Code As String
End Type
Private BarCodes() As udtBarCode

Private Type udtQBNames
    Short As String
    Long As String
End Type
Private QuickBioNames() As udtQBNames

Private Type udtQINames
    Short As String
    Long As String
End Type
Private QuickImmNames() As udtQINames

Private Type udtQENames
    Short As String
    Long As String
End Type
Private QuickEndNames() As udtQENames

Private CoagAnalyserName() As String






Private Sub LstHaePanel_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

      Dim n As Long
      Dim Found As Long

10    On Error GoTo LstHaePanel_MouseUp_Error

20    HaeChanged = True

30    Found = 0

40    For n = 0 To LstHaePanel.ListCount - 1
50        If LstHaePanel.Selected(n) Then Found = 1
60    Next

70    Exit Sub

LstHaePanel_MouseUp_Error:
      Dim strES As String
      Dim intEL As Integer

80    intEL = Erl
90    strES = Err.Description
100   LogError "frmNewOrder", "LstHaePanel_MouseUp", intEL, strES

End Sub

Private Sub tSampleID_Change()
10    If tSampleID = "" Then ClearRequests
End Sub

Private Sub txtBioComment_Change()

10    On Error GoTo txtBioComment_Change_Error

      'If bValidateBio.Caption = "VALID" Then Exit Sub

      'cmdSaveBio.Enabled = True

20    Exit Sub

txtBioComment_Change_Error:

      Dim strES As String
      Dim intEL As Integer

30    intEL = Erl
40    strES = Err.Description
50    LogError "frmNewOrder", "txtBioComment_Change", intEL, strES

End Sub

Private Sub txtBioComment_KeyDown(KeyCode As Integer, Shift As Integer)
      Dim sql As String
      Dim tb As New Recordset
      Dim s As Variant
      Dim n As Long
      Dim z As Integer

10    On Error GoTo txtBioComment_KeyDown_Error

      'If bValidateBio.Caption = "VALID" Then Exit Sub

20    If KeyCode = vbKeyF2 Then
30        If Trim(txtBioComment) = "" Then Exit Sub    '
40        n = txtBioComment.SelStart
50        If n < 3 Then Exit Sub
60        z = 1
70        s = Mid(txtBioComment, (n - z), z + 1)
80        z = 2

90        If ListText("BI", s) <> "" Then
100           s = ListText("BI", s)
110       Else
120           s = ""
130       End If

140       If s = "" And Len(txtBioComment) > 2 Then
150           z = 2
160           s = Mid(txtBioComment, (n - z), z + 1)
170           z = 3

180           If ListText("BI", s) <> "" Then
190               s = ListText("BI", s)
200           Else
210               s = ""
220           End If
230       End If

240       If s = "" Then
250           z = 1
260           s = Mid(txtBioComment, n, z + 1)

270           If ListText("BI", s) <> "" Then
280               s = ListText("BI", s)
290           End If
300       End If

310       txtBioComment = Left(txtBioComment, (n - (z)))
320       txtBioComment = txtBioComment & s

330       txtBioComment.SelStart = Len(txtBioComment)

340   ElseIf KeyCode = 114 Then

350       sql = "SELECT * from lists WHERE listtype = 'BI' order by listorder"
360       Set tb = New Recordset
370       RecOpenServer 0, tb, sql
380       Do While Not tb.EOF
390           s = Trim(tb!Text)
400           frmMessages.lstComm.AddItem s
410           tb.MoveNext
420       Loop

430       Set frmMessages.f = Me
440       Set frmMessages.T = txtBioComment
450       frmMessages.Show 1

460   End If

      'cmdSaveBio.Enabled = True

470   Exit Sub

txtBioComment_KeyDown_Error:

      Dim strES As String
      Dim intEL As Integer

480   intEL = Erl
490   strES = Err.Description
500   LogError "frmNewOrder", "txtBioComment_KeyDown", intEL, strES, sql

End Sub

Private Sub txtBioComment_KeyPress(KeyAscii As Integer)

10    On Error GoTo txtBioComment_KeyPress_Error


20    KeyAscii = VI(KeyAscii, AlphaNumeric_NoApos)


30    Exit Sub

txtBioComment_KeyPress_Error:

      Dim strES As String
      Dim intEL As Integer

40    intEL = Erl
50    strES = Err.Description
60    LogError "frmNewOrder", "txtBioComment_KeyPress", intEL, strES

End Sub





Private Function CheckCodes() As Boolean

          Dim n As Long
          Dim Y As Long



10        On Error GoTo CheckCodes_Error

20        CheckCodes = False

30        For n = 0 To UBound(BarCodes)
40            With BarCodes(n)
50                If .Code = tinput Then
60                    If .BarCodeType = "Control" Then
70                        Select Case .Name
                          Case "CTLCANCEL": Unload Me: Exit Function
80                        Case "CTLSAVE":
90                            bsave = True
100                           If mFromEdit Then
110                               mFromEdit = False
120                               Unload Me
130                               Exit Function
140                           End If
150                       Case "CTLCLEAR": ClearRequests
160                       Case "CTLRANDOM": oSorF(0) = True
170                       Case "CTLFASTING": oSorF(1) = True
180                       Case "CTLA": lAnalyserID = "A": AnalyserID = "A"
190                       Case "CTLB": lAnalyserID = "B": AnalyserID = "B"
200                       Case "CTLFBC": lstHaem.Selected(0) = Not lstHaem.Selected(0): HaemChanged = True
210                       Case "CTLESR": lstHaem.Selected(1) = Not lstHaem.Selected(1): HaemChanged = True
220                       Case "CTLRETICS": lstHaem.Selected(2) = Not lstHaem.Selected(2): HaemChanged = True
230                       Case "CTLMONOSPOT": lstHaem.Selected(3) = Not lstHaem.Selected(3): HaemChanged = True
240                       Case "CTLMALARIA": lstHaem.Selected(4) = Not lstHaem.Selected(4): HaemChanged = True
250                       Case "CTLSICKLEDEX": lstHaem.Selected(5) = Not lstHaem.Selected(5): HaemChanged = True
260                       Case "CTLASOT": lstHaem.Selected(6) = Not lstHaem.Selected(6): HaemChanged = True
270                       Case "CTLADM": chkADM.Value = IIf(chkADM.Value = 1, 0, 1)
280                       Case "CTLGLUCOSE": chkGBottle.Value = IIf(chkGBottle.Value = 1, 0, 1)
290                       End Select
300                       tinput = ""
310                       tinput.SetFocus
320                       CheckCodes = True
330                       Exit Function

340                   ElseIf .BarCodeType = "Coag" Then

350                       For Y = 0 To lstCoag.ListCount - 1
360                           If .Name = UCase$(lstCoag.List(Y)) Then
370                               lstCoag.Selected(Y) = Not lstCoag.Selected(Y)
380                               CoagChanged = True
390                               CheckCodes = True
400                               Exit Function
410                           End If
420                       Next

430                   ElseIf .BarCodeType = "Immuno" Then

440                       For Y = 0 To lstImmunoTests.ListCount - 1
450                           If .Name = UCase$(lstImmunoTests.List(Y)) Then
460                               lstImmunoTests.Selected(Y) = Not lstImmunoTests.Selected(Y)
470                               ImmunoChanged = True
480                               CheckCodes = True
490                               Exit Function
500                           End If
510                       Next

520                   ElseIf .BarCodeType = "Endo" Then

530                       For Y = 0 To lstEndoTests.ListCount - 1
540                           If .Name = UCase$(lstEndoTests.List(Y)) Then
550                               lstEndoTests.Selected(Y) = Not lstEndoTests.Selected(Y)
560                               EndoChanged = True
570                               CheckCodes = True
580                               Exit Function
590                           End If
600                       Next


610                   End If
620               End If
630           End With
640       Next




650       Exit Function

CheckCodes_Error:

          Dim strES As String
          Dim intEL As Integer



660       intEL = Erl
670       strES = Err.Description
680       LogError "frmNewOrder", "CheckCodes", intEL, strES


End Function

Private Sub FillCoagAnalyser()

      Dim sql As String
      Dim tb As Recordset
      Dim n As Integer

10    On Error GoTo FillCoagAnalyser_Error

20    sql = "SELECT Contents FROM Options WHERE " & _
            "Description LIKE 'CoagAnalyserName%' " & _
            "AND COALESCE(Contents, '') <> ''"
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql
50    If tb.EOF Then
60        lblCoagAnalyserName = ""
70        lblCoagAnalyserName.Visible = False
80    Else
          'lblCoagAnalyserName.Visible = True
90        ReDim CoagAnalyserName(0 To 0)
100       n = 0
110       Do While Not tb.EOF
120           ReDim Preserve CoagAnalyserName(0 To n)
130           CoagAnalyserName(n) = tb!Contents
140           n = n + 1
150           tb.MoveNext
160       Loop
170       sql = "SELECT Contents FROM Options WHERE " & _
                "Description = 'CoagAnalyserDefault' " & _
                "AND COALESCE(Contents, '') <> ''"
180       Set tb = New Recordset
190       RecOpenServer 0, tb, sql
200       If tb.EOF Then
210           lblCoagAnalyserName = CoagAnalyserName(0)
220       Else
230           lblCoagAnalyserName = tb!Contents
240       End If
250   End If

260   Exit Sub

FillCoagAnalyser_Error:

      Dim strES As String
      Dim intEL As Integer

270   intEL = Erl
280   strES = Err.Description
290   LogError "frmNewOrder", "FillCoagAnalyser", intEL, strES, sql

End Sub

Private Sub FillKnownHaeOrders()

          Dim tb As Recordset
          Dim sql As String
          Dim n As Integer
          Dim TestName As String
          Dim LongOrShort As String


10        On Error GoTo FillKnownHaeOrders_Error

20        LongOrShort = IIf(optLong, "Long", "Short")

30        With LstHaePanel
40            For n = 0 To .ListCount - 1
50                .Selected(n) = False
60            Next

70            sql = "SELECT T." & LongOrShort & "Name Name, R.Code " & _
                    "FROM HaeRequests R " & _
                    "JOIN HaemTestDefinitions T " & _
                    "ON R.Code = T.Code " & _
                    "WHERE SampleID = '" & Val(tSampleID) & "' " & _
                    "AND InUse = 1 "



80            sql = "SELECT Code as PanelName   FROM HaeRequests WHERE SampleID = '" & Val(tSampleID) & "' "



90            Set tb = New Recordset
100           RecOpenClient 0, tb, sql
110           Do While Not tb.EOF
120               TestName = tb!PanelName    ' CoagNameFor(tb!Code & "")
130               For n = 0 To .ListCount - 1
140                   If .List(n) = TestName Then
150                       .Selected(n) = True
160                       Exit For
170                   End If
180               Next
190               tb.MoveNext
200           Loop
210       End With


220       Exit Sub


FillKnownHaeOrders_Error:

          Dim strES As String
          Dim intEL As Integer

230       intEL = Erl
240       strES = Err.Description
250       LogError "frmNewOrder", "FillKnownHaeOrders", intEL, strES, sql

End Sub

Private Sub FillKnownBioOrders()

      Dim tb As New Recordset
      Dim sql As String
      Dim n As Long
      Dim Found As Boolean
      Dim LongOrShort As String

10    On Error GoTo FillKnownBioOrders_Error

20    Found = False

30    LongOrShort = IIf(optLong, "Long", "Short")

40    oSorF(0) = True    'Random
50    sql = "SELECT Fasting from Demographics WHERE " & _
            "SampleID = '" & Val(tSampleID) & "'"
60    Set tb = New Recordset
70    RecOpenServer 0, tb, sql
80    If Not tb.EOF Then
90        If Not IsNull(tb!Fasting) Then
100           If tb!Fasting Then oSorF(1) = True
110       End If
120   End If


130   sql = "SELECT BR.*, BT." & LongOrShort & "Name as Name " & _
            "from BioRequests as BR, BioTestDefinitions as BT WHERE " & _
            "SampleID = '" & Val(tSampleID) & "' " & _
            "and BR.Code = BT.Code " & _
            "and BR.SampleType = '" & SysOptBioST(0) & " '"
140   Set tb = New Recordset
150   RecOpenServer 0, tb, sql
160   Do While Not tb.EOF
170       For n = 0 To lSerumTests.ListCount - 1
180           If tb!Name = lSerumTests.List(n) Then
190               lSerumTests.Selected(n) = True
200               Found = True
210               Exit For
220           End If
230       Next
240       tb.MoveNext
250   Loop

260   If Not Found Then
270       sql = "SELECT BR.*, BT." & LongOrShort & "Name as Name " & _
                "from BioRequests as BR, BioTestDefinitions as BT WHERE " & _
                "SampleID = '" & Val(tSampleID) & "' " & _
                "and BR.Code = BT.Code " & _
                "and BR.SampleType = 'PL'"
280       Set tb = New Recordset
290       RecOpenServer 0, tb, sql
300       Do While Not tb.EOF
310           For n = 0 To lPlasmaTests.ListCount - 1
320               If tb!Name = lPlasmaTests.List(n) Then
330                   lPlasmaTests.Selected(n) = True
340                   Found = True
350                   Exit For
360               End If
370           Next
380           tb.MoveNext
390       Loop
400   End If

410   If Not Found Then
420       sql = "SELECT BR.*, BT." & LongOrShort & "Name as Name " & _
                "from BioRequests as BR, BioTestDefinitions as BT WHERE " & _
                "SampleID = '" & Val(tSampleID) & "' " & _
                "and BR.Code = BT.Code " & _
                "and BR.SampleType = 'U'"
430       Set tb = New Recordset
440       RecOpenServer 0, tb, sql
450       Do While Not tb.EOF
460           For n = 0 To lUrineTests.ListCount - 1
470               If tb!Name = lUrineTests.List(n) Then
480                   lUrineTests.Selected(n) = True
490                   Found = True
500                   Exit For
510               End If
520           Next
530           tb.MoveNext
540       Loop
550   End If

560   If Not Found Then
570       sql = "SELECT BR.*, BT." & LongOrShort & "Name as Name " & _
                "from BioRequests as BR, BioTestDefinitions as BT WHERE " & _
                "SampleID = '" & Val(tSampleID) & "' " & _
                "and BR.Code = BT.Code " & _
                "and BR.SampleType = 'FL'"
580       Set tb = New Recordset
590       RecOpenServer 0, tb, sql
600       Do While Not tb.EOF
610           For n = 0 To lstFluid.ListCount - 1
620               If tb!Name = lstFluid.List(n) Then
630                   lstFluid.Selected(n) = True
640                   Found = True
650                   Exit For
660               End If
670           Next
680           tb.MoveNext
690       Loop
700   End If

710   If Not Found Then
720       sql = "SELECT BR.*, BT." & LongOrShort & "Name as Name " & _
                "from BioRequests as BR, BioTestDefinitions as BT WHERE " & _
                "SampleID = '" & Val(tSampleID) & "' " & _
                "and BR.Code = BT.Code " & _
                "and BR.SampleType = 'C'"
730       Set tb = New Recordset
740       RecOpenServer 0, tb, sql
750       Do While Not tb.EOF
760           For n = 0 To lCSFTests.ListCount - 1
770               If tb!Name = lCSFTests.List(n) Then
780                   lCSFTests.Selected(n) = True
790                   Found = True
800                   Exit For
810               End If
820           Next
830           tb.MoveNext
840       Loop
850   End If

860   If Not Found Then
870       sql = "SELECT BR.*, BT." & LongOrShort & "Name as Name " & _
                "from BioRequests as BR, BioTestDefinitions as BT WHERE " & _
                "SampleID = '" & Val(tSampleID) & "' " & _
                "and BR.Code = BT.Code "
880       Set tb = New Recordset
890       RecOpenServer 0, tb, sql
900       Do While Not tb.EOF
910           For n = 0 To lstImmunoTests.ListCount - 1
920               If tb!Name = lstImmunoTests.List(n) Then
930                   lstImmunoTests.Selected(n) = True
940                   Exit For
950               End If
960           Next
970           tb.MoveNext
980       Loop
990   End If

1000  If SysOptDeptImm(0) Then
1010      sql = "SELECT BR.*, BT." & LongOrShort & "Name as Name " & _
                "from ImmRequests as BR, ImmTestDefinitions as BT WHERE " & _
                "SampleID = '" & Val(tSampleID) & "' " & _
                "and BR.Code = BT.Code "
1020      Set tb = New Recordset
1030      RecOpenServer 0, tb, sql
1040      Do While Not tb.EOF
1050          For n = 0 To lstImmunoTests.ListCount - 1
1060              If tb!Name = lstImmunoTests.List(n) Then
1070                  lstImmunoTests.Selected(n) = True
1080                  Exit For
1090              End If
1100          Next
1110          tb.MoveNext
1120      Loop
1130  End If

1140  If SysOptDeptEnd(0) Then
1150      sql = "SELECT BR.*, BT." & LongOrShort & "Name as Name " & _
                "from endRequests as BR, endTestDefinitions as BT WHERE " & _
                "SampleID = '" & Val(tSampleID) & "' " & _
                "and BR.Code = BT.Code " & _
                "AND BT.SampleType <> 'PL'"
1160      Set tb = New Recordset
1170      RecOpenServer 0, tb, sql
1180      Do While Not tb.EOF
1190          For n = 0 To lstEndoTests.ListCount - 1
1200              If tb!Name = lstEndoTests.List(n) Then
1210                  lstEndoTests.Selected(n) = True
1220                  Exit For
1230              End If
1240          Next
1250          tb.MoveNext
1260      Loop

1270      sql = "SELECT BR.*, BT." & LongOrShort & "Name as Name " & _
                "from endRequests as BR, endTestDefinitions as BT WHERE " & _
                "SampleID = '" & Val(tSampleID) & "' " & _
                "and BR.Code = BT.Code " & _
                "AND BT.SampleType = 'PL'"
1280      Set tb = New Recordset
1290      RecOpenServer 0, tb, sql
1300      Do While Not tb.EOF
1310          For n = 0 To lEndoTestsPlasma.ListCount - 1
1320              If tb!Name = lEndoTestsPlasma.List(n) Then
1330                  lEndoTestsPlasma.Selected(n) = True
1340                  Exit For
1350              End If
1360          Next
1370          tb.MoveNext
1380      Loop
1390  End If

1400  Exit Sub

FillKnownBioOrders_Error:

      Dim strES As String
      Dim intEL As Integer

1410  intEL = Erl
1420  strES = Err.Description
1430  LogError "frmNewOrder", "FillKnownBioOrders", intEL, strES, sql

End Sub

Private Sub LoadBarCodes()

      Dim sql As String
      Dim tb As New Recordset
      Dim intCurrentUpper As Long

10    On Error GoTo LoadBarCodes_Error

20    ReDim BarCodes(0 To 0) As udtBarCode
      Dim CodeAdded As Boolean
      Dim LongOrShort As String

30    LongOrShort = IIf(optLong, "Long", "Short")

40    CodeAdded = False
50    sql = "SELECT * from BarCodeControl"
60    Set tb = New Recordset
70    RecOpenServer 0, tb, sql
80    With tb
90        Do While Not .EOF
100           If Trim$(!Text) <> "" And Trim$(!Code) <> "" Then
110               If CodeAdded Then
120                   intCurrentUpper = UBound(BarCodes)
130                   ReDim Preserve BarCodes(0 To intCurrentUpper + 1)
140                   intCurrentUpper = intCurrentUpper + 1
150               End If
160               BarCodes(intCurrentUpper).Name = Trim$(UCase$(!Text))
170               BarCodes(intCurrentUpper).Code = Trim$(UCase$(!Code))
180               BarCodes(intCurrentUpper).BarCodeType = "Control"
190               CodeAdded = True
200           End If
210           .MoveNext
220       Loop
230   End With

240   sql = "SELECT Distinct TestName, Code from CoagTestDefinitions"
250   Set tb = New Recordset
260   RecOpenServer 0, tb, sql
270   With tb
280       Do While Not .EOF
290           If Trim$(!TestName & "") <> "" And Trim$(!Code & "") <> "" Then
300               If CodeAdded Then
310                   intCurrentUpper = UBound(BarCodes)
320                   ReDim Preserve BarCodes(0 To intCurrentUpper + 1)
330                   intCurrentUpper = intCurrentUpper + 1
340               End If
350               BarCodes(intCurrentUpper).Name = Trim$(UCase$(!TestName))
360               BarCodes(intCurrentUpper).Code = Trim$(UCase$(!Code))
370               BarCodes(intCurrentUpper).BarCodeType = "Coag"
380               CodeAdded = True
390           End If
400           .MoveNext
410       Loop
420   End With

430   sql = "SELECT Distinct " & LongOrShort & "Name as name, BarCode from BioTestDefinitions WHERE " & _
            "Analyser = '4' " & _
            "and Hospital = '" & HospName(0) & "'"
440   Set tb = New Recordset
450   RecOpenServer 0, tb, sql
460   With tb
470       Do While Not .EOF
480           If Trim$(!Name) <> "" And Trim$(!BarCode) <> "" Then
490               If CodeAdded Then
500                   intCurrentUpper = UBound(BarCodes)
510                   ReDim Preserve BarCodes(0 To intCurrentUpper + 1)
520                   intCurrentUpper = intCurrentUpper + 1
530               End If
540               BarCodes(intCurrentUpper).Name = Trim$(UCase$(!Name))
550               BarCodes(intCurrentUpper).Code = Trim$(UCase$(!BarCode))
560               BarCodes(intCurrentUpper).BarCodeType = "Immuno"
570               CodeAdded = True
580           End If
590           .MoveNext
600       Loop
610   End With

620   Exit Sub

LoadBarCodes_Error:

      Dim strES As String
      Dim intEL As Integer

630   intEL = Erl
640   strES = Err.Description
650   LogError "frmNewOrder", "LoadBarCodes", intEL, strES, sql

End Sub

Public Property Let SampleID(ByVal sNewValue As String)

10    On Error GoTo SampleID_Error

20    mSampleID = sNewValue

30    Exit Property

SampleID_Error:

      Dim strES As String
      Dim intEL As Integer

40    intEL = Erl
50    strES = Err.Description
60    LogError "frmNewOrder", "SampleID", intEL, strES

End Property

Private Sub FillCoagList()

      Dim sql As String
      Dim tb As New Recordset

10    On Error GoTo FillCoagList_Error

20    lstCoag.Clear

30    sql = "SELECT DISTINCT TestName, PrintPriority FROM CoagTestDefinitions WHERE " & _
            "Hospital = '" & HospName(0) & "' " & _
            "AND InUse = 1 " & _
            "ORDER BY PrintPriority"
40    Set tb = New Recordset
50    RecOpenServer 0, tb, sql
60    Do While Not tb.EOF
70        If Trim$(UCase$(tb!TestName & "")) <> "FIB" Then
80            lstCoag.AddItem tb!TestName
90        End If
100       tb.MoveNext
110   Loop

120   Exit Sub

FillCoagList_Error:

      Dim strES As String
      Dim intEL As Integer

130   intEL = Erl
140   strES = Err.Description
150   LogError "frmNewOrder", "FillCoagList", intEL, strES, sql

End Sub

Public Property Let FromEdit(ByVal bFromEdit As Boolean)

10    On Error GoTo FromEdit_Error

20    mFromEdit = bFromEdit

30    Exit Property

FromEdit_Error:

      Dim strES As String
      Dim intEL As Integer

40    intEL = Erl
50    strES = Err.Description
60    LogError "frmNewOrder", "FromEdit", intEL, strES

End Property

Function CheckCSF() As Boolean

      Dim Y As Long
      Dim tb As New Recordset
      Dim sql As String
      Dim LongOrShort As String


10    On Error GoTo CheckCSF_Error

20    LongOrShort = IIf(optLong, "Long", "Short")

30    CheckCSF = False
40    sql = "SELECT " & LongOrShort & "Name as Name from BioTestDefinitions WHERE " & _
            "SampleType = 'C' " & _
            "and BarCode =  '" & tinput & "'"
50    Set tb = New Recordset
60    RecOpenServer 0, tb, sql
70    If Not tb.EOF Then
80        CheckCSF = True
90        For Y = 0 To lCSFTests.ListCount - 1
100           If lCSFTests.List(Y) = tb!Name Then
110               lCSFTests.Selected(Y) = Not lCSFTests.Selected(Y)
120               BioChanged = True
130               Exit For
140           End If
150       Next
160   End If




170   Exit Function

CheckCSF_Error:

      Dim strES As String
      Dim intEL As Integer



180   intEL = Erl
190   strES = Err.Description
200   LogError "frmNewOrder", "CheckCSF", intEL, strES, sql


End Function

Function CheckPlasma() As Boolean

      Dim Y As Long
      Dim tb As New Recordset
      Dim sql As String
      Dim LongOrShort As String


10    On Error GoTo CheckPlasma_Error

20    LongOrShort = IIf(optLong, "Long", "Short")

30    CheckPlasma = False
40    sql = "SELECT " & LongOrShort & "Name as Name from BioTestDefinitions WHERE " & _
            "SampleType = 'PL' " & _
            "and BarCode = '" & tinput & "' and knowntoanalyser = '1'"
50    Set tb = New Recordset
60    RecOpenServer 0, tb, sql
70    If Not tb.EOF Then
80        CheckPlasma = True
90        For Y = 0 To lPlasmaTests.ListCount - 1
100           If lPlasmaTests.List(Y) = tb!Name Then
110               lPlasmaTests.Selected(Y) = Not lPlasmaTests.Selected(Y)
120               Exit For
130           End If
140       Next
150       BioChanged = True
160   End If




170   Exit Function

CheckPlasma_Error:

      Dim strES As String
      Dim intEL As Integer



180   intEL = Erl
190   strES = Err.Description
200   LogError "frmNewOrder", "CheckPlasma", intEL, strES, sql


End Function

Function CheckSerum() As Boolean

      Dim Y As Long
      Dim tb As New Recordset
      Dim sql As String
      Dim LongOrShort As String


10    On Error GoTo CheckSerum_Error

20    LongOrShort = IIf(optLong, "Long", "Short")

30    CheckSerum = False
40    sql = "SELECT " & LongOrShort & "Name as Name from BioTestDefinitions WHERE " & _
            "SampleType = '" & SysOptBioST(0) & "' " & _
            "and BarCode = '" & tinput & "' and knowntoanalyser = '1'"
50    Set tb = New Recordset
60    RecOpenServer 0, tb, sql
70    If Not tb.EOF Then
80        CheckSerum = True
90        For Y = 0 To lSerumTests.ListCount - 1
100           If lSerumTests.List(Y) = tb!Name Then
110               lSerumTests.Selected(Y) = Not lSerumTests.Selected(Y)
120               Exit For
130           End If
140       Next
150       BioChanged = True
160   End If




170   Exit Function

CheckSerum_Error:

      Dim strES As String
      Dim intEL As Integer



180   intEL = Erl
190   strES = Err.Description
200   LogError "frmNewOrder", "CheckSerum", intEL, strES, sql


End Function




Function CheckPlasmaPanel() As Boolean

      Dim Y As Long
      Dim tb As New Recordset
      Dim sql As String


10    On Error GoTo CheckPlasmaPanel_Error

20    CheckPlasmaPanel = False

30    sql = "SELECT distinct(panelname) from Panels WHERE " & _
            "BarCode = '" & tinput & "' " & _
            "and PanelType = 'PL' "
40    Set tb = New Recordset
50    RecOpenServer 0, tb, sql

60    Do While Not tb.EOF
70        CheckPlasmaPanel = True
80        For Y = 0 To lPlasmaPanel.ListCount - 1
90            If lPlasmaPanel.List(Y) = tb!PanelName Then
100               lPlasmaPanel.Selected(Y) = Not lPlasmaPanel.Selected(Y)
110               Exit For
120           End If
130       Next
140       tb.MoveNext
150   Loop

160   If CheckPlasmaPanel Then
170       BioChanged = True
180   End If



190   Exit Function

CheckPlasmaPanel_Error:

      Dim strES As String
      Dim intEL As Integer



200   intEL = Erl
210   strES = Err.Description
220   LogError "frmNewOrder", "CheckPlasmaPanel", intEL, strES, sql


End Function
Function CheckEPlasmaPanel() As Boolean

      Dim Y As Long
      Dim tb As New Recordset
      Dim sql As String


10    On Error GoTo CheckEPlasmaPanel_Error

20    CheckEPlasmaPanel = False

30    sql = "SELECT distinct(panelname) from EndPanels WHERE " & _
            "BarCode = '" & tinput & "' " & _
            "and PanelType = 'PL' "
40    Set tb = New Recordset
50    RecOpenServer 0, tb, sql

60    Do While Not tb.EOF
70        CheckEPlasmaPanel = True
80        For Y = 0 To lEndoPanelPlasma.ListCount - 1
90            If lEndoPanelPlasma.List(Y) = tb!PanelName Then
100               lEndoPanelPlasma.Selected(Y) = Not lEndoPanelPlasma.Selected(Y)
110               Exit For
120           End If
130       Next
140       tb.MoveNext
150   Loop

160   If CheckEPlasmaPanel Then
170       BioChanged = True
180   End If



190   Exit Function

CheckEPlasmaPanel_Error:

      Dim strES As String
      Dim intEL As Integer



200   intEL = Erl
210   strES = Err.Description
220   LogError "frmNewOrder", "CheckEPlasmaPanel", intEL, strES, sql


End Function

Function CheckSerumPanel() As Boolean

      Dim Y As Long
      Dim tb As New Recordset
      Dim sql As String


10    On Error GoTo CheckSerumPanel_Error

20    CheckSerumPanel = False

30    sql = "SELECT distinct(panelname) from Panels WHERE " & _
            "BarCode = '" & tinput & "' " & _
            "and PanelType = '" & SysOptBioST(0) & "' "
40    Set tb = New Recordset
50    RecOpenServer 0, tb, sql

60    Do While Not tb.EOF
70        CheckSerumPanel = True
80        For Y = 0 To lSerumPanel.ListCount - 1
90            If lSerumPanel.List(Y) = tb!PanelName Then
100               lSerumPanel.Selected(Y) = Not lSerumPanel.Selected(Y)
110               Exit For
120           End If
130       Next
140       tb.MoveNext
150   Loop

160   If CheckSerumPanel Then
170       BioChanged = True
180   End If



190   Exit Function

CheckSerumPanel_Error:

      Dim strES As String
      Dim intEL As Integer



200   intEL = Erl
210   strES = Err.Description
220   LogError "frmNewOrder", "CheckSerumPanel", intEL, strES, sql


End Function
Function CheckESerumPanel() As Boolean

      Dim Y As Long
      Dim tb As New Recordset
      Dim sql As String


10    On Error GoTo CheckESerumPanel_Error

20    CheckESerumPanel = False

30    sql = "SELECT distinct(panelname) from EndPanels WHERE " & _
            "BarCode = '" & tinput & "' " & _
            "and PanelType = '" & SysOptBioST(0) & "' "
40    Set tb = New Recordset
50    RecOpenServer 0, tb, sql

60    Do While Not tb.EOF
70        CheckESerumPanel = True
80        For Y = 0 To lEndoPanel.ListCount - 1
90            If lEndoPanel.List(Y) = tb!PanelName Then
100               lEndoPanel.Selected(Y) = Not lEndoPanel.Selected(Y)
110               Exit For
120           End If
130       Next
140       tb.MoveNext
150   Loop


160   BioChanged = True




170   Exit Function

CheckESerumPanel_Error:

      Dim strES As String
      Dim intEL As Integer



180   intEL = Erl
190   strES = Err.Description
200   LogError "frmNewOrder", "CheckESerumPanel", intEL, strES, sql


End Function
Function CheckImmSerumPanel() As Boolean

      Dim Y As Long
      Dim tb As New Recordset
      Dim sql As String


10    On Error GoTo CheckIMMSerumPanel_Error

20    CheckImmSerumPanel = False

30    sql = "SELECT distinct(panelname) from IPanels WHERE " & _
            "BarCode = '" & tinput & "' " & _
            "and PanelType = '" & SysOptBioST(0) & "' "
40    Set tb = New Recordset
50    RecOpenServer 0, tb, sql

60    Do While Not tb.EOF
70        CheckImmSerumPanel = True
80        For Y = 0 To lImmunoPanel.ListCount - 1
90            If lImmunoPanel.List(Y) = tb!PanelName Then
100               lImmunoPanel.Selected(Y) = Not lImmunoPanel.Selected(Y)
110               Exit For
120           End If
130       Next
140       tb.MoveNext
150   Loop


160   BioChanged = True




170   Exit Function

CheckIMMSerumPanel_Error:

      Dim strES As String
      Dim intEL As Integer



180   intEL = Erl
190   strES = Err.Description
200   LogError "frmNewOrder", "CheckIMMSerumPanel", intEL, strES, sql


End Function
Function CheckUrine() As Boolean

      Dim tb As New Recordset
      Dim sql As String
      Dim Y As Long
      Dim LongOrShort As String


10    On Error GoTo CheckUrine_Error

20    LongOrShort = IIf(optLong, "Long", "Short")

30    CheckUrine = False
40    sql = "SELECT " & LongOrShort & "Name as Name from BioTestDefinitions WHERE " & _
            "SampleType = 'U' " & _
            "and BarCode = '" & tinput & "'"
50    Set tb = New Recordset
60    RecOpenServer 0, tb, sql
70    If Not tb.EOF Then
80        CheckUrine = True
90        For Y = 0 To lUrineTests.ListCount - 1
100           If lUrineTests.List(Y) = tb!Name Then
110               lUrineTests.Selected(Y) = Not lUrineTests.Selected(Y)
120               Exit For
130           End If
140       Next
150       BioChanged = True
160   End If



170   Exit Function

CheckUrine_Error:

      Dim strES As String
      Dim intEL As Integer



180   intEL = Erl
190   strES = Err.Description
200   LogError "frmNewOrder", "CheckUrine", intEL, strES, sql


End Function

Function CheckUrinePanel() As Boolean

      Dim Y As Long
      Dim tb As New Recordset
      Dim sql As String


10    On Error GoTo CheckUrinePanel_Error

20    CheckUrinePanel = False
30    sql = "SELECT * from Panels WHERE " & _
            "BarCode = '" & tinput & "' " & _
            "and PanelType = 'U' "
40    Set tb = New Recordset
50    RecOpenServer 0, tb, sql
60    Do While Not tb.EOF
70        CheckUrinePanel = True
80        For Y = 0 To lUrinePanel.ListCount - 1
90            If lUrinePanel.List(Y) = tb!PanelName Then
100               lUrinePanel.Selected(Y) = Not lUrinePanel.Selected(Y)
110               Exit For
120           End If
130       Next

140       tb.MoveNext
150   Loop

160   BioChanged = True



170   Exit Function

CheckUrinePanel_Error:

      Dim strES As String
      Dim intEL As Integer



180   intEL = Erl
190   strES = Err.Description
200   LogError "frmNewOrder", "CheckUrinePanel", intEL, strES, sql


End Function
Sub ClearRequests()

      Dim n As Long


10    On Error GoTo ClearRequests_Error

20    chkUrgent.Value = 0

30    For n = 0 To lPlasmaPanel.ListCount - 1
40        If lPlasmaPanel.Selected(n) = True Then lPlasmaPanel.Selected(n) = False
50    Next

60    For n = 0 To lPlasmaTests.ListCount - 1
70        If lPlasmaTests.Selected(n) = True Then lPlasmaTests.Selected(n) = False
80    Next

90    For n = 0 To lSerumPanel.ListCount - 1
100       If lSerumPanel.Selected(n) = True Then lSerumPanel.Selected(n) = False
110   Next

120   For n = 0 To lSerumTests.ListCount - 1
130       If lSerumTests.Selected(n) = True Then lSerumTests.Selected(n) = False
140   Next

150   For n = 0 To lUrinePanel.ListCount - 1
160       If lUrinePanel.Selected(n) = True Then lUrinePanel.Selected(n) = False
170   Next

180   For n = 0 To lUrineTests.ListCount - 1
190       lUrineTests.Selected(n) = False
200   Next

210   For n = 0 To lCSFTests.ListCount - 1
220       lCSFTests.Selected(n) = False
230   Next

240   For n = 0 To lImmunoPanel.ListCount - 1
250       lImmunoPanel.Selected(n) = False
260   Next

270   For n = 0 To lstImmunoTests.ListCount - 1
280       lstImmunoTests.Selected(n) = False
290   Next

300   For n = 0 To lEndoPanel.ListCount - 1
310       lEndoPanel.Selected(n) = False
320   Next

330   For n = 0 To lstEndoTests.ListCount - 1
340       lstEndoTests.Selected(n) = False
350   Next

360   For n = 0 To lstCoag.ListCount - 1
370       lstCoag.Selected(n) = False
380   Next

390   For n = 0 To lstFluid.ListCount - 1
400       lstFluid.Selected(n) = False
410   Next

420   For n = 0 To lstHaem.ListCount - 1
430       lstHaem.Selected(n) = False
440   Next

450   For n = 0 To LstHaePanel.ListCount - 1
460       LstHaePanel.Selected(n) = False
470   Next

480   For n = 0 To lEndoPanelPlasma.ListCount - 1
490       lEndoPanelPlasma.Selected(n) = False
500   Next

510   For n = 0 To lEndoTestsPlasma.ListCount - 1
520       lEndoTestsPlasma.Selected(n) = False
530   Next
540   txtBioComment = ""
550   txtImmComment(0) = ""

560   chkUrgent.Value = 0




570   Exit Sub

ClearRequests_Error:

      Dim strES As String
      Dim intEL As Integer



580   intEL = Erl
590   strES = Err.Description
600   LogError "frmNewOrder", "ClearRequests", intEL, strES


End Sub

Private Sub FillLists()

Dim tb As New Recordset
Dim sql As String
Dim LongOrShort As String
Dim n As Long
Dim Found As Boolean

On Error GoTo FillLists_Error

FillCoagList

LongOrShort = IIf(optLong, "Long", "Short")

lSerumPanel.Clear
lSerumTests.Clear
lUrinePanel.Clear
lUrineTests.Clear
lCSFTests.Clear
lstImmunoTests.Clear
lstEndoTests.Clear
lImmunoPanel.Clear
lEndoPanel.Clear
lstFluid.Clear
lPlasmaPanel.Clear
lPlasmaTests.Clear
lEndoPanelPlasma.Clear
lEndoTestsPlasma.Clear

With lstHaem
    .Clear
    .AddItem "FBC"
    .AddItem "ESR"
    .AddItem "Retics"
    .AddItem "Monospot"
    .AddItem "Malaria"
    .AddItem "Sickledex"
    .AddItem "Asot"
    If SysOptBadRes(0) Then .AddItem "Bad"
    If SysOptAlwaysRequestFBC(0) Then
        .Selected(0) = True
    End If
End With

With LstHaePanel
    .Clear
    sql = "Select distinct PanelName, ListOrder from HaePanels where " & _
          " Hospital = '" & HospName(0) & "' " & _
          "Order by ListOrder"
    Set tb = New Recordset
    RecOpenServer 0, tb, sql

    Do While Not tb.EOF
        .AddItem tb!PanelName
        tb.MoveNext
    Loop
End With


sql = "SELECT distinct PanelName, ListOrder, PanelType from Panels WHERE " & _
      "Hospital = '" & HospName(0) & "' " & _
      "Order by ListOrder"
Set tb = New Recordset
RecOpenServer 0, tb, sql
Do While Not tb.EOF
    Select Case tb!PanelType
        Case SysOptBioST(0):
            lSerumPanel.AddItem tb!PanelName
        Case "U"
            lUrinePanel.AddItem tb!PanelName
        Case "PL"
            lPlasmaPanel.AddItem tb!PanelName
    End Select

    tb.MoveNext
Loop


'sql = "SELECT distinct PanelName, ListOrder from Panels WHERE " & _
 '      "PanelType = '" & SysOptBioST(0) & "' " & _
 '      "and Hospital = '" & HospName(0) & "' " & _
 '      "Order by ListOrder"
'Set tb = New Recordset
'RecOpenServer 0, tb, sql
'Do While Not tb.EOF
'    lSerumPanel.AddItem tb!PanelName
'    tb.MoveNext
'Loop

sql = "SELECT distinct PanelName, ListOrder from IPanels WHERE " & _
      "Hospital = '" & HospName(0) & "' " & _
      "Order by ListOrder"
Set tb = New Recordset
RecOpenServer 0, tb, sql
Do While Not tb.EOF
    lImmunoPanel.AddItem tb!PanelName
    tb.MoveNext
Loop

sql = "SELECT distinct PanelName, ListOrder, PanelType from EndPanels WHERE " & _
      "Hospital = '" & HospName(0) & "' " & _
      "Order by ListOrder"
Set tb = New Recordset
RecOpenServer 0, tb, sql
Do While Not tb.EOF
    Select Case tb!PanelType
        Case "S"
            lEndoPanel.AddItem tb!PanelName
        Case "PL"
            lEndoPanelPlasma.AddItem tb!PanelName
    End Select


    tb.MoveNext
Loop

'sql = "SELECT distinct PanelName, ListOrder from Panels WHERE " & _
 '      "PanelType = 'U' " & _
 '      "and Hospital = '" & HospName(0) & "' " & _
 '      "Order by ListOrder"
'Set tb = New Recordset
'RecOpenServer 0, tb, sql
'Do While Not tb.EOF
'    lUrinePanel.AddItem tb!PanelName
'    tb.MoveNext
'Loop

sql = "SELECT distinct " & LongOrShort & "Name as Name, " & _
      "PrintPriority, SampleType from BioTestDefinitions WHERE " & _
      "Analyser <> '4' " & _
      "and KnownToAnalyser = '1' and inuse = '1'" & _
      "order by PrintPriority"
Set tb = Cnxn(0).Execute(sql)
Do While Not tb.EOF
    Select Case tb!SampleType
        Case SysOptBioST(0):
            lSerumTests.AddItem tb!Name
        Case "PL":
            lPlasmaTests.AddItem tb!Name
        Case "U":
            lUrineTests.AddItem tb!Name
        Case "FL"
            lstFluid.AddItem tb!Name
        Case "C"
            lCSFTests.AddItem tb!Name
    End Select
    tb.MoveNext
Loop

'sql = "SELECT distinct " & LongOrShort & "Name as Name, " & _
 '      "PrintPriority from BioTestDefinitions WHERE " & _
 '      "Hospital = '" & HospName(0) & "' " & _
 '      "and Analyser <> '4' " & _
 '      "and SampleType = 'U' " & _
 '      "and KnownToAnalyser = '1' and inuse = '1' " & _
 '      "order by PrintPriority"
'Set tb = Cnxn(0).Execute(sql)
'Do While Not tb.EOF
'    lUrineTests.AddItem tb!Name
'    tb.MoveNext
'Loop

'sql = "SELECT distinct " & LongOrShort & "Name as Name, " & _
 '      "PrintPriority from BioTestDefinitions WHERE " & _
 '      "Hospital = '" & HospName(0) & "' " & _
 '      "and Analyser <> '4' " & _
 '      "and SampleType = 'FL' " & _
 '      "and KnownToAnalyser = '1' and inuse = '1' " & _
 '      "order by PrintPriority"
'Set tb = Cnxn(0).Execute(sql)
'Do While Not tb.EOF
'    lstFluid.AddItem tb!Name
'    tb.MoveNext
'Loop

'sql = "SELECT distinct " & LongOrShort & "Name as Name, " & _
 '      "PrintPriority from BioTestDefinitions WHERE " & _
 '      "Hospital = '" & HospName(0) & "' " & _
 '      "and Analyser <> '4' " & _
 '      "and SampleType = 'C' " & _
 '      "and KnownToAnalyser = '1' and inuse = '1' " & _
 '      "order by PrintPriority"
'Set tb = Cnxn(0).Execute(sql)
'Do While Not tb.EOF
'    lCSFTests.AddItem tb!Name
'    tb.MoveNext
'Loop

If SysOptDeptEnd(0) Then
    '        sql = "SELECT distinct " & LongOrShort & "Name as Name, printpriority, SampleType " & _
             '              " from endTestDefinitions WHERE " & _
             '              "Hospital = '" & HospName(0) & "' and inuse = '1' " & _
             '              "order by PrintPriority"
    sql = "SELECT distinct " & LongOrShort & "Name as Name, printpriority, SampleType " & _
          " from endTestDefinitions WHERE " & _
          " inuse = '1' " & _
          "order by PrintPriority"
    Set tb = Cnxn(0).Execute(sql)
    Do While Not tb.EOF
        For n = 0 To lstEndoTests.ListCount - 1
            If lstEndoTests.List(n) = tb!Name Then
                Found = True
                Exit For
            End If
        Next
        If Found = False Then
            Select Case tb!SampleType
                Case "PL"
                    lEndoTestsPlasma.AddItem tb!Name
                Case Else
                    lstEndoTests.AddItem tb!Name
            End Select

        End If
        Found = False
        tb.MoveNext
    Loop
End If

If SysOptDeptImm(0) Then
    '        sql = "SELECT distinct " & LongOrShort & "Name as Name, printpriority " & _
             '              " from ImmTestDefinitions WHERE " & _
             '              "Hospital = '" & HospName(0) & "' and inuse = '1' " & _
             '              "order by PrintPriority"
    sql = "SELECT distinct " & LongOrShort & "Name as Name, printpriority " & _
          " from ImmTestDefinitions WHERE " & _
          " inuse = '1' " & _
          "order by PrintPriority"
    Set tb = Cnxn(0).Execute(sql)
    Do While Not tb.EOF
        For n = 0 To lstImmunoTests.ListCount - 1
            If lstImmunoTests.List(n) = tb!Name Then
                Found = True
                Exit For
            End If
        Next
        If Found = False Then
            lstImmunoTests.AddItem tb!Name
        End If
        Found = False
        tb.MoveNext
    Loop
Else
    sql = "SELECT distinct " & LongOrShort & "Name as Name, " & _
          "PrintPriority from BioTestDefinitions WHERE " & _
          "Hospital = '" & HospName(0) & "' " & _
          "and Analyser = '4' " & _
          "and KnownToAnalyser = '1' and inuse = '1' " & _
          "order by PrintPriority"
    Set tb = Cnxn(0).Execute(sql)
    Do While Not tb.EOF
        lstImmunoTests.AddItem tb!Name
        tb.MoveNext
    Loop
End If

Exit Sub

FillLists_Error:

Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "frmNewOrder", "FillLists", intEL, strES, sql

End Sub

Private Sub UpDateRequestsHae(ByVal Discipline As String, _
                              ByVal TestName As String, _
                              ShortNameOrLongName As String)

          Dim sql As String

10        On Error GoTo UpDateRequestsHae_Error


20        sql = "INSERT INTO " & Discipline & "Requests " & vbNewLine
30        sql = sql & "(SampleId, Code, DateTimeOfRecord, SampleType, Units,  UserName,Analyser,Programmed) " & vbNewLine
40        sql = sql & " VALUES ( " & vbNewLine
50        sql = sql & " '" & tSampleID & "', " & vbNewLine
60        sql = sql & "        '" & TestName & "' , getdate(), " & vbNewLine
70        sql = sql & "        'Blood EDTA' , '',  " & vbNewLine
80        sql = sql & "       '" & UserName & "','IPU',0 " & vbNewLine
90        sql = sql & " )"

100       Cnxn(0).Execute sql




110       Exit Sub


UpDateRequestsHae_Error:

          Dim strES As String
          Dim intEL As Integer

120       intEL = Erl
130       strES = Err.Description
140       LogError "frmNewOrder", "UpDateRequestsHae", intEL, strES, sql

End Sub

Private Sub DeleteOldEnteries(Discipline As String)
10        On Error GoTo DeleteOldEnteries_Error

          Dim sql As String

20        sql = "Delete from  " & Discipline & "Requests "
30        sql = sql & " WHERE SAMPLEID ='" & tSampleID & "'"

40        Cnxn(0).Execute sql


50        Exit Sub


DeleteOldEnteries_Error:

          Dim strES As String
          Dim intEL As Integer

60        intEL = Erl
70        strES = Err.Description
80        LogError "frmNewOrder", "DeleteOldEnteries", intEL, strES
End Sub


Private Sub SaveHae()
10        On Error GoTo SaveHae_Error

          Dim n As Integer
          Dim TestName As String
          Dim sql As String
20        With LstHaePanel
30            DeleteOldEnteries ("Hae")
40            For n = 0 To .ListCount - 1
50                If .Selected(n) Then
60                    TestName = .List(n)
70                    UpDateRequestsHae "Hae", TestName, IIf((optLong.Value = True), "Long", "Short")
80                End If
90            Next
100       End With


        Dim tb As New ADODB.Recordset
110       sql = "SELECT * FROM demographics WHERE " & _
                "SampleID = '" & tSampleID & "'"
120       Set tb = New Recordset
130       RecOpenClient 0, tb, sql
140       If tb.EOF Then
150           tb.AddNew
160           tb!Rundate = Format$(Now, "dd/mmm/yyyy")
170           tb!SampleID = tSampleID
180           tb!Faxed = 0
190           tb!RooH = 0
200       End If
210       If chkUrgent.Value = 1 Then
220           tb!Urgent = 1
230       Else
240           tb!Urgent = 0
250       End If

260       tb.Update


270       Exit Sub


SaveHae_Error:

          Dim strES As String
          Dim intEL As Integer

280       intEL = Erl
290       strES = Err.Description
300       LogError "frmNewOrder", "SaveHae", intEL, strES
End Sub

Private Sub SaveBio()

      Dim n As Long
      Dim Code As String
      Dim sql As String
      Dim tb As New Recordset
      Dim LongOrShort As String
      Dim FBio As Boolean
      Dim FImm As Boolean
      Dim FEnd As Boolean
      Dim Analyser As String
      Dim Method As String
      Dim GlucoseCode As String
      Dim strAnalyser As String
      Dim strPatChart As String
      Dim strPatName As String
      Dim strPatDoB As String
      Dim strPatSex As String

10    On Error GoTo SaveBio_Error


20    On Error GoTo SaveBio_Error

      'GlucoseCode = GetOptionSetting("GlucoseCode", "996")
      'GlucoseCode = GetOptionSetting("GlucoseCode1", "996")


30    If GetOptionSetting("RemisolInterfaceInUse", 0) Then
          'Save Patient details to be used in Remisol tess ordering
40        sql = "SELECT Chart, PatName, Dob, sex  FROM Demographics WHERE " & _
                "SampleID = '" & tSampleID & "'"
50        Set tb = New Recordset
60        RecOpenServer 0, tb, sql
70        If tb.EOF Then
80            strPatChart = ""
90            strPatName = ""
100           strPatDoB = ""
110           strPatSex = "U"
120       Else
130           strPatChart = Trim(tb!Chart & "")
140           strPatName = Trim(tb!PatName & "")
150           strPatDoB = Format(tb!Dob, "dd/mmm/yyyy")
160           strPatSex = tb!sex & ""
170       End If

180       sql = "If NOT Exists(Select * From RemisolPatientDetails Where SampleID = '" & tSampleID & "' ) " & _
                "Begin " & _
                "Insert Into RemisolPatientDetails (SampleID, Chart, PatName, Dob, sex, username) Values "
190       If strPatDoB = "" Then
200           sql = sql & "('" & tSampleID & "', '" & strPatChart & "', '" & AddTicks(strPatName) & "', null, '" & strPatSex & "', '" & AddTicks(UserName) & "') " & _
                    "End"
210       Else
220           sql = sql & "('" & tSampleID & "', '" & strPatChart & "', '" & AddTicks(strPatName) & "', '" & strPatDoB & "', '" & strPatSex & "', '" & AddTicks(UserName) & "') " & _
                    "End"
230       End If

240       Cnxn(0).Execute sql
250   End If


260   LongOrShort = IIf(optLong, "Long", "Short")


270   Cnxn(0).Execute ("DELETE from BioRequests WHERE " & _
                       "SampleID = '" & tSampleID & "' " & _
                       "and Programmed = 0")

280   If optSerum Then
290       For n = 0 To lSerumTests.ListCount - 1
300           If lSerumTests.Selected(n) Then
                  '60            If chkADM.Value = 1 Then
                  '70              Code = lSerumTests.List(n)
                  '80              lAnalyserID = "Centralink"
                  '90            Else
310               If LongOrShort = "Long" Then
320                   Code = CodeForLongName(lSerumTests.List(n))
330               Else
340                   Code = CodeForShortName(lSerumTests.List(n))
350               End If
360               strAnalyser = BioAnalyserIDForCode(Code)

370               sql = "SELECT Contents FROM Options WHERE Description Like 'GlucoseCode%'"
380               Set tb = New Recordset
390               RecOpenServer 0, tb, sql
400               If tb.EOF Then
410                   GlucoseCode = ""
420               Else
430                   While Not tb.EOF
440                       If Code = tb!Contents & "" Then
450                           GlucoseCode = tb!Contents
460                       End If
470                       tb.MoveNext
480                   Wend
490               End If

500               sql = "INSERT into BioRequests " & _
                        "(SampleID, Code, DateTime, SampleType, Programmed, AnalyserID, GBottle, Hospital) VALUES " & _
                        "('" & tSampleID & "', " & _
                        "'" & Code & "', " & _
                        "'" & Format$(Now, "dd/mmm/yyyy hh:mm") & "', " & _
                        "'" & SysOptBioST(0) & "', " & _
                        "'0', " & _
                        "'" & strAnalyser & "', " & _
                        IIf(Code = GlucoseCode, chkGBottle.Value, 0) & ", " & _
                        "'" & GetHospitalName(Code, "Bio") & "')"
510               Cnxn(0).Execute sql
520               FBio = True
530               strAnalyser = ""
540           End If
550       Next
560   ElseIf optPlasma Then
570       For n = 0 To lPlasmaTests.ListCount - 1
580           If lPlasmaTests.Selected(n) Then

590               If LongOrShort = "Long" Then
600                   Code = CodeForLongName(lPlasmaTests.List(n))
610               Else
620                   Code = CodeForShortName(lPlasmaTests.List(n))
630               End If
640               strAnalyser = BioAnalyserIDForCode(Code)
                  
650               sql = "SELECT Contents FROM Options WHERE Description Like 'GlucoseCode%'"
660               Set tb = New Recordset
670               RecOpenServer 0, tb, sql
680               If tb.EOF Then
690                   GlucoseCode = ""
700               Else
710                   While Not tb.EOF
720                       If Code = tb!Contents & "" Then
730                           GlucoseCode = tb!Contents
740                       End If
750                       tb.MoveNext
760                   Wend
770               End If

780               sql = "INSERT into BioRequests " & _
                        "(SampleID, Code, DateTime, SampleType, Programmed, AnalyserID, GBottle, Hospital) VALUES " & _
                        "('" & tSampleID & "', " & _
                        "'" & Code & "', " & _
                        "'" & Format$(Now, "dd/mmm/yyyy hh:mm") & "', " & _
                        "'PL', " & _
                        "'0', " & _
                        "'" & strAnalyser & "', " & _
                        IIf(Code = GlucoseCode, chkGBottle.Value, 0) & ", " & _
                        "'" & GetHospitalName(Code, "Bio") & "')"
790               Cnxn(0).Execute sql
800               FBio = True
810               strAnalyser = ""
820           End If
830       Next
840   End If
850   For n = 0 To lUrineTests.ListCount - 1
860       If lUrineTests.Selected(n) Then
              '230           If chkADM.Value = 1 Then
              '240             Code = lSerumTests.List(n)
              '250             lAnalyserID = "Centralink"
              '260           Else
870           If LongOrShort = "Long" Then
880               Code = CodeForLongName(lUrineTests.List(n))
890           Else
900               Code = CodeForShortName(lUrineTests.List(n))
910           End If
920           strAnalyser = BioAnalyserIDForCode(Code)
              '320           End If
930           sql = "INSERT into BioRequests " & _
                    "(SampleID, Code, DateTime, SampleType, Programmed, AnalyserID, GBottle, Hospital) VALUES " & _
                    "('" & tSampleID & "', " & _
                    "'" & Code & "', " & _
                    "'" & Format$(Now, "dd/mmm/yyyy hh:mm") & "', " & _
                    "'U', " & _
                    "'0', " & _
                    "'" & strAnalyser & "', " & _
                    "0, " & _
                    "'" & GetHospitalName(Code, "Bio") & "')"
940           Cnxn(0).Execute sql
950           FBio = True
960           strAnalyser = ""
970       End If
980   Next

990   For n = 0 To lCSFTests.ListCount - 1
1000      If lCSFTests.Selected(n) Then
1010          If LongOrShort = "Long" Then
1020              Code = CodeForLongName(lCSFTests.List(n))
1030          Else
1040              Code = CodeForShortName(lCSFTests.List(n))
1050          End If
1060          strAnalyser = BioAnalyserIDForCode(Code)
1070          sql = "INSERT into BioRequests " & _
                    "(SampleID, Code, DateTime, SampleType, Programmed, AnalyserID, GBottle, Hospital) VALUES " & _
                    "('" & tSampleID & "', " & _
                    "'" & Code & "', " & _
                    "'" & Format$(Now, "dd/mmm/yyyy hh:mm") & "', " & _
                    "'C', " & _
                    "'0', " & _
                    "'" & strAnalyser & "', " & _
                    "0, " & _
                    "'" & GetHospitalName(Code, "Bio") & "')"
1080          Cnxn(0).Execute sql
1090          FBio = True
1100          strAnalyser = ""
1110      End If
1120  Next

1130  For n = 0 To lstFluid.ListCount - 1
1140      If lstFluid.Selected(n) Then
1150          If LongOrShort = "Long" Then
1160              Code = CodeForLongName(lstFluid.List(n))
1170          Else
1180              Code = CodeForShortName(lstFluid.List(n))
1190          End If
1200          strAnalyser = BioAnalyserIDForCode(Code)
1210          sql = "INSERT into BioRequests " & _
                    "(SampleID, Code, DateTime, SampleType, Programmed, AnalyserID, GBottle, Hospital) VALUES " & _
                    "('" & tSampleID & "', " & _
                    "'" & Code & "', " & _
                    "'" & Format$(Now, "dd/mmm/yyyy hh:mm") & "', " & _
                    "'FL', " & _
                    "'0', " & _
                    "'" & strAnalyser & "', " & _
                    "0, " & _
                    "'" & GetHospitalName(Code, "Bio") & "')"
1220          Cnxn(0).Execute sql
1230          FBio = True
1240          strAnalyser = ""
1250      End If
1260  Next



1270  If SysOptDeptEnd(0) = True Then
1280      Cnxn(0).Execute ("DELETE from endRequests WHERE " & _
                           "SampleID = '" & tSampleID & "' " & _
                           "and Programmed = 0")
1290      If optSerum Then
1300          For n = 0 To lstEndoTests.ListCount - 1
1310              If lstEndoTests.Selected(n) Then

1320                  If LongOrShort = "Long" Then
1330                      Code = eCodeForLongName(lstEndoTests.List(n))
1340                  Else
1350                      Code = eCodeForShortName(lstEndoTests.List(n))
1360                  End If

1370                  Analyser = eAnylForCode(Code)
1380                  If chkADM.Value = 1 Then
1390                      Analyser = "ADM"
1400                  End If

                      'Created on 01/02/2011 16:00:09
                      'Autogenerated by SQL Scripting
1410                  sql = "If Exists(Select 1 From EndRequests " & _
                            "Where SampleID = @SampleID0 " & _
                            "And Code = '@Code1' " & _
                            "And AnalyserID = '@AnalyserID5' ) " & _
                            "Begin " & _
                            "Update EndRequests Set " & _
                            "Code = '@Code1', " & _
                            "DateTime = '@DateTime3', " & _
                            "SampleType = '@SampleType4', " & _
                            "AnalyserID = '@AnalyserID5' " & _
                            "Where SampleID = @SampleID0 " & _
                            "And Code = '@Code1' " & _
                            "And AnalyserID = '@AnalyserID5'  " & _
                            "And Hospital = '@Hospital'  " & _
                            "End  " & _
                            "Else " & _
                            "Begin  " & _
                            "Insert Into EndRequests (SampleID, Code, Programmed, DateTime, SampleType, AnalyserID,Hospital) Values " & _
                            "(@SampleID0, '@Code1', @Programmed2, '@DateTime3', '@SampleType4', '@AnalyserID5','@Hospital') " & _
                            "End"


1420                  sql = Replace(sql, "@SampleID0", tSampleID)
1430                  sql = Replace(sql, "@Code1", Code)
1440                  sql = Replace(sql, "@Programmed2", 0)
1450                  sql = Replace(sql, "@DateTime3", Format$(Now, "dd/mmm/yyyy hh:mm"))
1460                  sql = Replace(sql, "@SampleType4", "S")
1470                  sql = Replace(sql, "@AnalyserID5", Analyser)
1480                  sql = Replace(sql, "@Hospital", GetHospitalName(Code, "End"))


1490                  Cnxn(0).Execute sql
                      '            sql = "select * from endrequests where " & _
                                   '                  "sampleid = '" & tSampleID & "' " & _
                                   '                  "and code = '" & Code & "' " & _
                                   '                  "and analyserid = '" & Analyser & "'"
                      '            Set tb = New Recordset
                      '            RecOpenServer 0, tb, sql
                      '            If tb.EOF Then
                      '                tb.AddNew
                      '                tb!SampleID = tSampleID
                      '                tb!Programmed = 0
                      '            End If
                      '            tb!SampleType = "S"
                      '            tb!AnalyserID = Analyser
                      '            tb!Code = Code
                      '            tb!Datetime = Format$(Now, "dd/mmm/yyyy hh:mm")
                      '            tb.Update
1500                  FEnd = True
1510              End If
1520          Next
1530      ElseIf optPlasma Then
1540          For n = 0 To lEndoTestsPlasma.ListCount - 1
1550              If lEndoTestsPlasma.Selected(n) Then

1560                  If LongOrShort = "Long" Then
1570                      Code = eCodeForLongName(lEndoTestsPlasma.List(n))
1580                  Else
1590                      Code = eCodeForShortName(lEndoTestsPlasma.List(n))
1600                  End If

1610                  Analyser = eAnylForCode(Code)
1620                  If chkADM.Value = 1 Then
1630                      Analyser = "ADM"
1640                  End If

                      'Created on 01/02/2011 16:00:09
                      'Autogenerated by SQL Scripting
1650                  sql = "If Exists(Select 1 From EndRequests " & _
                            "Where SampleID = @SampleID0 " & _
                            "And Code = '@Code1' " & _
                            "And AnalyserID = '@AnalyserID5' ) " & _
                            "Begin " & _
                            "Update EndRequests Set " & _
                            "Code = '@Code1', " & _
                            "DateTime = '@DateTime3', " & _
                            "SampleType = '@SampleType4', " & _
                            "AnalyserID = '@AnalyserID5' " & _
                            "Where SampleID = @SampleID0 " & _
                            "And Code = '@Code1' " & _
                            "And AnalyserID = '@AnalyserID5'  " & _
                            "And Hospital ='@Hospital'" & _
                            "End  " & _
                            "Else " & _
                            "Begin  " & _
                            "Insert Into EndRequests (SampleID, Code, Programmed, DateTime, SampleType, AnalyserID,Hospital) Values " & _
                            "(@SampleID0, '@Code1', @Programmed2, '@DateTime3', '@SampleType4', '@AnalyserID5','@Hospital') " & _
                            "End"


1660                  sql = Replace(sql, "@SampleID0", tSampleID)
1670                  sql = Replace(sql, "@Code1", Code)
1680                  sql = Replace(sql, "@Programmed2", 0)
1690                  sql = Replace(sql, "@DateTime3", Format$(Now, "dd/mmm/yyyy hh:mm"))
1700                  sql = Replace(sql, "@SampleType4", "PL")
1710                  sql = Replace(sql, "@AnalyserID5", Analyser)
1720                  sql = Replace(sql, "@Hospital", GetHospitalName(Code, "End"))

1730                  Cnxn(0).Execute sql

1740                  FEnd = True
1750              End If
1760          Next
1770      End If

1780  End If

1790  If SysOptDeptImm(0) = True Then
1800      Cnxn(0).Execute ("DELETE FROM ImmRequests WHERE " & _
                           "SampleID = '" & tSampleID & "' " & _
                           "AND Programmed = 0")
1810      For n = 0 To lstImmunoTests.ListCount - 1
1820          If lstImmunoTests.Selected(n) Then
                  '970         If chkADM.Value = 1 Then
                  '980           Code = lstImmunoTests.List(n)
                  '990           Analyser = "Centralink"
                  '1000        Else
1830              If LongOrShort = "Long" Then
1840                  Code = iCodeForLongName(lstImmunoTests.List(n))
1850              Else
1860                  Code = ICodeForShortName(lstImmunoTests.List(n))
1870              End If
                  '1060        End If
1880              sql = "SELECT Analyser, Method FROM ImmTestDefinitions WHERE " & _
                        "Code = '" & Code & "'"
1890              Set tb = New Recordset
1900              RecOpenServer 0, tb, sql
1910              If Not tb.EOF Then
1920                  Analyser = Trim$(tb!Analyser & "")
1930                  Method = tb!Method & ""
1940              End If

                  'Created on 02/02/2011 10:49:02
                  'Autogenerated by SQL Scripting

1950              sql = "If Exists(Select 1 From ImmRequests " & _
                        "Where SampleID = @SampleID0 " & _
                        "And Code = '@Code1' " & _
                        "And AnalyserID = '@AnalyserID5' ) " & _
                        "Begin " & _
                        "Update ImmRequests Set " & _
                        "Code = '@Code1', " & _
                        "DateTime = '@DateTime3', " & _
                        "SampleType = '@SampleType4', " & _
                        "AnalyserID = '@AnalyserID5', " & _
                        "Method = '@Method6', " & _
                        "Hospital = '@Hospital7' " & _
                        "Where SampleID = @SampleID0 " & _
                        "And Code = '@Code1' " & _
                        "And AnalyserID = '@AnalyserID5'  " & _
                        "End  " & _
                        "Else " & _
                        "Begin  " & _
                        "Insert Into ImmRequests (SampleID, Code, Programmed, DateTime, SampleType, AnalyserID, Method, Hospital) Values " & _
                        "(@SampleID0, '@Code1', @Programmed2, '@DateTime3', '@SampleType4', '@AnalyserID5', '@Method6', '@Hospital7') " & _
                        "End"

1960              sql = Replace(sql, "@SampleID0", tSampleID)
1970              sql = Replace(sql, "@Code1", Code)
1980              sql = Replace(sql, "@Programmed2", 0)
1990              sql = Replace(sql, "@DateTime3", Format$(Now, "dd/mmm/yyyy hh:mm"))
2000              sql = Replace(sql, "@SampleType4", "S")
2010              sql = Replace(sql, "@AnalyserID5", Trim(Analyser) & "")
2020              sql = Replace(sql, "@Method6", Method)
2030              sql = Replace(sql, "@Hospital7", GetHospitalName(Code, "Imm"))

2040              Cnxn(0).Execute sql


                  '            sql = "SELECT * FROM ImmRequests WHERE " & _
                               '                  "SampleID = '" & tSampleID & "' " & _
                               '                  "AND Code = '" & Code & "' " & _
                               '                  "AND AnalyserID = '" & Analyser & "'"
                  '            Set tb = New Recordset
                  '            RecOpenServer 0, tb, sql
                  '            If tb.EOF Then
                  '                tb.AddNew
                  '                tb!SampleID = tSampleID
                  '                tb!Programmed = 0
                  '            End If
                  '            tb!SampleType = "S"
                  '            tb!AnalyserID = Trim(Analyser) & ""
                  '            tb!Method = Method
                  '            tb!Code = Code
                  '            tb!Datetime = Format$(Now, "dd/mmm/yyyy hh:mm")
                  '            tb.Update
2050              FImm = True
2060          End If
2070      Next
2080  End If

2090  If SysOptDeptEnd(0) = False And SysOptDeptImm(0) = False Then
2100      For n = 0 To lstImmunoTests.ListCount - 1
2110          If lstImmunoTests.Selected(n) Then
2120              If LongOrShort = "Long" Then
2130                  Code = CodeForLongName(lstImmunoTests.List(n))
2140              Else
2150                  Code = CodeForShortName(lstImmunoTests.List(n))
2160              End If
2170              sql = "INSERT into BioRequests " & _
                        "(SampleID, Code, DateTime, SampleType, Programmed, AnalyserID) VALUES " & _
                        "('" & tSampleID & "', " & _
                        "'" & Code & "', " & _
                        "'" & Format$(Now, "dd/mmm/yyyy hh:mm") & "', " & _
                        "'S', " & _
                        "'0', " & _
                        "'4')"
2180              Cnxn(0).Execute sql
2190              FBio = True
2200          End If
2210      Next
2220  End If

      'Created on 02/02/2011 11:39:04
      'Autogenerated by SQL Scripting

2230  sql = "If Exists(Select 1 From Demographics " & _
            "Where SampleID = @SampleID44 ) " & _
            "Begin " & _
            "Update Demographics Set " & _
            "Fasting = @Fasting19, " & _
            "RunDate = '@RunDate42', " & _
            "Urgent = @Urgent48, " & _
            "username = '@username49' " & _
            "Where SampleID = @SampleID44  " & _
            "End  " & _
            "Else " & _
            "Begin  " & _
            "Insert Into Demographics (Fasting, RunDate, SampleID, Urgent, username) Values " & _
            "(@Fasting19, '@RunDate42', @SampleID44, @Urgent48, '@username49') " & _
            "End"

2240  sql = Replace(sql, "@Fasting19", IIf(oSorF(1), 1, 0))
2250  sql = Replace(sql, "@RunDate42", Format$(Now, "dd/mmm/yyyy"))
2260  sql = Replace(sql, "@SampleID44", tSampleID)
2270  sql = Replace(sql, "@Urgent48", IIf(chkUrgent.Value = 1, 1, 0))
2280  sql = Replace(sql, "@username49", UserName)

2290  Cnxn(0).Execute sql

      'sql = "SELECT * FROM Demographics WHERE " & _
       '      "SampleID = '" & tSampleID & "'"
      'Set tb = New Recordset
      'RecOpenServer 0, tb, sql
      'If tb.EOF Then
      '    tb.AddNew
      '    tb!Rundate = Format$(Now, "dd/mmm/yyyy")
      '    tb!SampleID = tSampleID
      'End If
      '
      'If chkUrgent.Value = 1 Then tb!Urgent = 1 Else tb!Urgent = 0
      'tb!Fasting = IIf(oSorF(1), 1, 0)
      'tb!UserName = UserName
      'tb.Update

2300  Exit Sub

SaveBio_Error:

      Dim strES As String
      Dim intEL As Integer

2310  intEL = Erl
2320  strES = Err.Description
2330  LogError "frmNewOrder", "SaveBio", intEL, strES, sql


2340  Exit Sub


End Sub

Private Sub SaveCoag()

      Dim sql As String
      Dim n As Long
      Dim TestCode As String
      Dim tb As Recordset
      Dim Unit As String
      Dim Analyser As String

10    On Error GoTo SaveCoag_Error

20    sql = "DELETE from CoagRequests WHERE " & _
            "SampleID = '" & mSampleID & "' " & _
            "AND Analyser = '" & lblCoagAnalyserName & "'"
30    Cnxn(0).Execute sql

40    For n = 0 To lstCoag.ListCount - 1
50        If lstCoag.Selected(n) Then
60            If chkADM.Value = 1 Then
70                TestCode = CoagCodeFor(lstCoag.List(n))
80                Analyser = "ADM"
90            Else
100               TestCode = CoagCodeFor(lstCoag.List(n))
110               Analyser = ""
120           End If
130           Unit = CoagUnitsFor(TestCode)
140           sql = "INSERT INTO CoagRequests " & _
                    "(SampleID, Code,AnalyserID, Units, Analyser, Programmed) VALUES " & _
                    "('" & mSampleID & "', " & _
                    "'" & TestCode & "', " & _
                    "'" & Analyser & "', " & _
                    "'" & Unit & "', " & _
                    "'" & lblCoagAnalyserName & "', 0 )"
150           Cnxn(0).Execute sql
160       End If
170   Next

180   sql = "SELECT * FROM Demographics WHERE " & _
            "SampleID = '" & tSampleID & "'"

190   Set tb = New Recordset
200   RecOpenClient 0, tb, sql
210   If tb.EOF Then
220       tb.AddNew
230       tb!Rundate = Format$(Now, "dd/mmm/yyyy")
240       tb!SampleID = tSampleID
250   End If

260   If chkUrgent.Value = 1 Then tb!Urgent = 1 Else tb!Urgent = 0
270   tb.Update

280   For n = 0 To lstCoag.ListCount - 1
290       lstCoag.Selected(n) = False
300   Next

310   Exit Sub

SaveCoag_Error:

      Dim strES As String
      Dim intEL As Integer

320   intEL = Erl
330   strES = Err.Description
340   LogError "frmNewOrder", "SaveCoag", intEL, strES, sql

End Sub

Private Sub SetDisplay(ByVal LongOrShort As String)

      Dim sql As String
      Dim tb As New Recordset

10    On Error GoTo SetDisplay_Error

20    If SysOptLongOrShortBioNames(0) <> LongOrShort Then
30        If iMsg("Do you want to reset the Default Display to " & LongOrShort & " Names?", vbQuestion + vbYesNo) = vbYes Then
40            sql = "SELECT * from Options WHERE " & _
                    "Description = 'LongOrShortBioNames'"
50            Set tb = New Recordset
60            RecOpenServer 0, tb, sql
70            If tb.EOF Then
80                tb.AddNew
90                tb!Description = "LongOrShortBioNames"
100           End If
110           tb!Contents = LongOrShort
120           tb.Update
130           SysOptLongOrShortBioNames(0) = LongOrShort
140       End If
150   End If

160   Exit Sub

SetDisplay_Error:

      Dim strES As String
      Dim intEL As Integer

170   intEL = Erl
180   strES = Err.Description
190   LogError "frmNewOrder", "SetDisplay", intEL, strES, sql

End Sub

Private Sub bcancel_Click()

10    Unload Me

End Sub

Private Sub bClear_Click()

10    On Error GoTo bClear_Click_Error

20    pBar = 0
30    ClearRequests

40    Exit Sub

bClear_Click_Error:

      Dim strES As String
      Dim intEL As Integer



50    intEL = Erl
60    strES = Err.Description
70    LogError "frmNewOrder", "bClear_Click", intEL, strES


End Sub

Private Sub bSave_Click()
      Dim sql As String

10    On Error GoTo bSave_Click_Error

20    pBar = 0

30    If Trim$(tSampleID) = "" Then
40        iMsg "Sample Number Required.", vbCritical
50        Exit Sub
60    End If

70    If SysOptUrgent(0) Then
80        If chkUrgent.Value = 1 Then
90            sql = "Update demographics set urgent = 1 where sampleid = " & tSampleID & ""
100           Cnxn(0).Execute sql
110       End If
120   End If

130   If CoagChanged Then SaveCoag
140   If HaemChanged Then SaveHaem
150   If BioChanged Then SaveBio
160   If HaeChanged Then SaveHae

170   SaveComments


180   ClearRequests

190   tSampleID = Format$(Val(tSampleID) + 1)
200   tSampleID.SelStart = 0
210   tSampleID.SelLength = Len(tSampleID)

220   If mFromEdit Then
230       mFromEdit = False
240       Unload Me
250       Exit Sub
260   End If

270   tSampleID.SetFocus

280   Exit Sub

bSave_Click_Error:

      Dim strES As String
      Dim intEL As Integer

290   intEL = Erl
300   strES = Err.Description
310   LogError "frmNewOrder", "bSave_Click", intEL, strES, sql

End Sub
Private Sub SaveComments()

      Dim Obs As New Observations

10    On Error GoTo SaveComments_Error

20    tSampleID = Format(Val(mSampleID))
30    If Val(tSampleID) = 0 Then Exit Sub



40    Obs.Save tSampleID, True, _
               "Biochemistry", Trim$(txtBioComment), _
               "Endocrinology", Trim$(txtImmComment(0))

50    Exit Sub

SaveComments_Error:

      Dim strES As String
      Dim intEL As Integer

60    intEL = Erl
70    strES = Err.Description
80    LogError "frmNewOrder", "SaveComments", intEL, strES

End Sub

Private Sub SaveHaem()

          Dim tb As New Recordset
          Dim sql As String
          Dim n As Long
          Dim Found As Boolean
          Dim ESRFound As Boolean
          Dim ReticsFound As Boolean
          Dim MonoSpotFound As Boolean
          Dim MalariaFound As Boolean
          Dim SickleFound As Boolean
          Dim FBCFound As Boolean
          Dim BadFound As Boolean
          Dim AsotFound As Boolean
          Dim StrOrder As String
          Dim Analyser As String
          Dim CodeToOrder As String


10        On Error GoTo SaveHaem_Error

20        StrOrder = ""
30        Found = False
40        ESRFound = False
50        ReticsFound = False
60        MonoSpotFound = False
70        MalariaFound = False
80        SickleFound = False
90        FBCFound = False
100       BadFound = False
110       AsotFound = False

120       sql = "DELETE FROM HaemRequests WHERE " & _
                "SampleID = '" & mSampleID & "'"
130       Cnxn(0).Execute sql

140       If chkADM.Value = 1 Then
150           CodeToOrder = ""
160           Analyser = "ADM"

170       End If



180       sql = "If Exists(Select 1 From HaemResults " & _
                "Where sampleid = @sampleid120 ) " & _
                "Begin " & _
                "Update HaemResults Set " & _
                "casot = @casot10, " & _
                "cesr = @cesr20, " & _
                "cmalaria = @cmalaria23, " & _
                "cmonospot = @cmonospot24, " & _
                "cretics = @cretics26, " & _
                "csickledex = @csickledex27 " & _
                "Where sampleid = @sampleid120  " & _
                "End  "

190       sql = Replace(sql, "@casot10", 0)
200       sql = Replace(sql, "@cesr20", 0)
210       sql = Replace(sql, "@cmalaria23", 0)
220       sql = Replace(sql, "@cmonospot24", 0)
230       sql = Replace(sql, "@cretics26", 0)
240       sql = Replace(sql, "@csickledex27", 0)
250       sql = Replace(sql, "@sampleid120", mSampleID)

260       Cnxn(0).Execute sql

270       For n = 0 To 5
280           If lstHaem.Selected(n) Then
290               Select Case n
                  Case 0    'FBC
300                   Found = True
310                   FBCFound = True
320               Case 1    'ESR
330                   Found = True
340                   ESRFound = True
350               Case 2    'Retics
360                   Found = True
370                   ReticsFound = True
380               Case 3    'MonoSpot
390                   Found = True
400                   MonoSpotFound = True
410               Case 4    'Malaria
420                   Found = True
430                   MalariaFound = True
440               Case 5    'Sickle
450                   Found = True
460                   SickleFound = True
470               Case 6    'ASot
480                   Found = True
490                   AsotFound = True
500               End Select
510           End If
520       Next

530       If SysOptBadRes(0) Then
540           If lstHaem.Selected(7) Then
550               Found = True
560               BadFound = True
570           End If
580       End If
590       If Found Then
600           sql = "SELECT * from  demographics WHERE " & _
                    "sampleid = '" & tSampleID & "'"
610           Set tb = New Recordset
620           RecOpenServer 0, tb, sql
630           If tb.EOF Then
640               tb.AddNew
650               tb!SampleID = tSampleID
660               tb!Rundate = Format(Now, "dd/MMM/yyyy")
670           End If
680           If chkUrgent.Value = 1 Then tb!Urgent = 1 Else tb!Urgent = 0
690           tb.Update

700           If SysOptHaemAn1(0) = "ADVIA" Then
710               sql = "DELETE from haemrequests WHERE sampleid = '" & mSampleID & "'"
720               Cnxn(0).Execute sql
730               If FBCFound = True And ReticsFound = True Then
740                   StrOrder = "B"
750                   CodeToOrder = "^^^HAEHGB\^^^HAENEUTA\^^^HAERETA"
760               ElseIf FBCFound = True Then
770                   StrOrder = "F"
780                   CodeToOrder = "^^^HAEHGB\^^^HAENEUTA"
790               ElseIf ReticsFound = True Then
800                   StrOrder = "R"
810                   CodeToOrder = "^^^HAERETA"
820               End If
830               If StrOrder <> "" Then

840                   sql = "SELECT * from HaemRequests WHERE " & _
                            "SampleID = '" & tSampleID & "'"
850                   Set tb = New Recordset
860                   RecOpenServer 0, tb, sql
870                   If tb.EOF Then
880                       tb.AddNew
890                   End If
900                   tb!SampleID = tSampleID
910                   tb!Anl1 = 0
920                   tb!Anl2 = 0
930                   tb!Orders = StrOrder
940                   tb!Code = CodeToOrder
950                   tb!AnalyserID = Analyser
960                   tb.Update

970               End If


980           End If
990       End If

1000      If ESRFound Or ReticsFound Or MonoSpotFound Or MalariaFound Or SickleFound Or BadFound Then
1010          sql = "SELECT * from HaemResults WHERE " & _
                    "SampleID = '" & tSampleID & "'"
1020          Set tb = New Recordset
1030          RecOpenClient 0, tb, sql
1040          If tb.EOF Then
1050              tb.AddNew
1060              tb!Rundate = Format$(Now, "dd/mmm/yyyy")
1070              tb!RunDateTime = Format$(Now, "dd/mmm/yyyy hh:mm")
1080              tb!SampleID = tSampleID
1090              tb!Faxed = 0
1100              tb!Printed = 0
1110              tb!ccoag = 0
1120              tb!Valid = 0
1130              tb!Printed = 0
1140          End If
              '**************Trevor 30/01/2020*****************
1150          tb!cESR = IIf(ESRFound, 1, 0)
1160          tb!cRetics = IIf(ReticsFound, 1, 0)
1170          tb!cMonospot = IIf(MonoSpotFound, 1, 0)
1180          tb!cmalaria = IIf(MalariaFound, 1, 0)
1190          tb!csickledex = IIf(SickleFound, 1, 0)
1200          tb!cASot = IIf(AsotFound, 1, 0)
              
1210          If ESRFound Then
1220              tb!cESR = 1
1230              If (tb!esr & "") = "" Then tb!esr = "?"
1240          Else
1250              tb!cESR = 0
1260              If tb!esr = "?" Then tb!esr = ""
1270          End If

1280          If ReticsFound Then
1290              tb!cRetics = 1
1300              If (tb!reta & "") = "" Then tb!reta = "?"
1310              If (tb!RetP & "") = "" Then tb!RetP = "?"
1320          Else
1330              tb!cRetics = 0
1340              If (tb!reta & "?") = "" Then tb!reta = ""
1350              If (tb!RetP & "?") = "" Then tb!RetP = ""
1360          End If
      '
      '        If MonoSpotFound Then
      '            tb!cMonospot = 1
      '            If (tb!MonoSpot & "") = "" Then tb!MonoSpot = "?"
      '        Else
      '            tb!cMonospot = 0
      '            If tb!MonoSpot = "?" Then tb!MonoSpot = ""
      '        End If
      '
      '        If MalariaFound Then
      '            tb!cMalaria = 1
      '            If (tb!Malaria & "") = "" Then tb!Malaria = "?"
      '        Else
      '            tb!cMalaria = 0
      '            If tb!Malaria = "?" Then tb!Malaria = ""
      '        End If
      '
      '        If SickleFound Then
      '            tb!csickledex = 1
      '            If (tb!sickledex & "") = "" Then tb!sickledex = "?"
      '        Else
      '            tb!csickledex = 0
      '            If tb!sickledex = "?" Then tb!sickledex = ""
      '        End If

              '***********************************************

1370          If SysOptBadRes(0) Then
1380              tb!cbad = IIf(BadFound, 1, 0)
1390          End If
1400          tb.Update
1410      End If

1420      Exit Sub

SaveHaem_Error:

          Dim strES As String
          Dim intEL As Integer

1430      intEL = Erl
1440      strES = Err.Description
1450      LogError "frmNewOrder", "SaveHaem", intEL, strES, sql

End Sub






Private Sub chkGBottle_Click()
10    If chkGBottle.Value = 0 Then
20        chkGBottle.Caption = "Glucose bottle is NOT in use"
30        chkGBottle.BackColor = vbRed
40    ElseIf chkGBottle.Value = 1 Then
50        chkGBottle.Caption = "Glucose bottle is in use"
60        chkGBottle.BackColor = &H8000000F
70    End If
End Sub

Private Sub Form_Activate()

10    On Error GoTo Form_Activate_Error

      '20    chkADM.Value = 0
      '30    If IsIDE Then
      '40      chkADM.Value = 1
      '50    End If

20    If Activated Then Exit Sub
30    Activated = True

40    Set_Font Me

50    tSampleID = mSampleID

60    tSampleID.SetFocus

70    If SysOptUrgent(0) Then chkUrgent.Visible = True

80    If mSampleID <> "" Then
90        FillKnownCoagOrders
100       FillKnownHaemOrders
110       FillKnownHaeOrders
120       FillKnownBioOrders
130       tinput.SetFocus
140   End If

150   If SysOptAlwaysRequestFBC(0) Then
160       lstHaem.Selected(0) = True
170   End If

180   If SysOptLongOrShortBioNames(0) = "Long" Then
190       optLong = True
200   Else
210       optShort = True
220   End If

230   CoagChanged = False
240   HaemChanged = False: HaeChanged = False
250   BioChanged = False
260   ImmunoChanged = False
270   EndoChanged = False

280   If GetOptionSetting("DisableGBottleDetection", 0) = 1 Then
290       chkGBottle.Value = 0
300   Else
310       chkGBottle.Value = 1
320   End If


330   Exit Sub

Form_Activate_Error:

      Dim strES As String
      Dim intEL As Integer

340   intEL = Erl
350   strES = Err.Description
360   LogError "frmNewOrder", "Form_Activate", intEL, strES

End Sub

Private Sub Form_Load()

10    On Error GoTo Form_Load_Error

20    If GetOptionSetting("EnableADM", 0) = 0 Then
30        chkADM.Value = 0
40        chkADM.Visible = False
50    Else
60        chkADM.Value = 0
70        chkADM.Visible = True
80    End If

90    If optSerum.Value = True Then
100       lSerumPanel.Visible = True
110       lSerumTests.Visible = True
120       lPlasmaPanel.Visible = False
130       lPlasmaTests.Visible = False
140       lEndoPanelPlasma.Visible = False
150       lEndoTestsPlasma.Visible = False
160   ElseIf optPlasma.Visible = True Then
170       lSerumPanel.Visible = False
180       lSerumTests.Visible = False
190       lPlasmaPanel.Visible = True
200       lPlasmaTests.Visible = True
210       lEndoPanelPlasma.Visible = True
220       lEndoTestsPlasma.Visible = True
230   End If




240   FillLists
250   LoadBarCodes
260   FillQuickBioNames
270   FillQuickImmNames
280   FillQuickEndNames

290   FillCoagAnalyser

300   Activated = False

310   Exit Sub

Form_Load_Error:

      Dim strES As String
      Dim intEL As Integer

320   intEL = Erl
330   strES = Err.Description
340   LogError "frmNewOrder", "Form_Load", intEL, strES

End Sub

Private Sub FillKnownCoagOrders()

      Dim tb As New Recordset
      Dim sql As String
      Dim n As Long
      Dim TestName As String

10    On Error GoTo FillKnownCoagOrders_Error

20    For n = 0 To lstCoag.ListCount - 1
30        lstCoag.Selected(n) = False
40    Next

50    If Val(mSampleID) = 0 Then Exit Sub

60    sql = "SELECT * from CoagRequests WHERE " & _
            "SampleID = '" & mSampleID & "' AND Analyser = '" & lblCoagAnalyserName & "'"
70    Set tb = New Recordset
80    RecOpenServer 0, tb, sql
90    Do While Not tb.EOF
100       TestName = CoagNameFor(tb!Code & "")
110       For n = 0 To lstCoag.ListCount - 1
120           If Trim(lstCoag.List(n)) = Trim(TestName) Then
130               lstCoag.Selected(n) = True
140               Exit For
150           End If
160       Next
170       tb.MoveNext
180   Loop

190   Exit Sub

FillKnownCoagOrders_Error:

      Dim strES As String
      Dim intEL As Integer

200   intEL = Erl
210   strES = Err.Description
220   LogError "frmNewOrder", "FillKnownCoagOrders", intEL, strES, sql

End Sub

Private Sub FillKnownHaemOrders()

      Dim tb As New Recordset
      Dim sql As String
      Dim n As Long

10    On Error GoTo FillKnownHaemOrders_Error

20    For n = 0 To 5
30        lstHaem.Selected(n) = False
40    Next

50    sql = "SELECT * from HaemResults WHERE " & _
            "SampleID = '" & mSampleID & "'"
60    Set tb = New Recordset
70    RecOpenServer 0, tb, sql
80    If Not tb.EOF Then
90        If (tb!cESR <> 0 And tb!esr & "" = "") _
             Or (tb!cESR <> 0 And tb!esr & "" = "?") _
             Or (tb!cESR <> 0 And IsNull(tb!esr)) Then
100           lstHaem.Selected(1) = True
110       End If
120       If tb!cRetics <> 0 And tb!reta = "" Or tb!cRetics <> 0 And tb!reta = "?" Or tb!cRetics <> 0 And IsNull(tb!reta) Then lstHaem.Selected(2) = True
130       If tb!cMonospot <> 0 And tb!Monospot = "" Or tb!cMonospot <> 0 And IsNull(tb!Monospot) Then lstHaem.Selected(3) = True
140       If tb!cmalaria <> 0 Then lstHaem.Selected(4) = True
150       If tb!csickledex <> 0 Then lstHaem.Selected(5) = True
160       If tb!cASot <> 0 And tb!tASOt = "" Or tb!cASot <> 0 And IsNull(tb!tASOt) Then lstHaem.Selected(6) = True
170       If SysOptBadRes(0) Then
180           If tb!cbad <> 0 Then lstHaem.Selected(7) = True
190       End If
200   End If

210   sql = "SELECT * from HaemRequests WHERE " & _
            "SampleID = '" & mSampleID & "'"
220   Set tb = New Recordset
230   RecOpenServer 0, tb, sql
240   If Not tb.EOF Then
250       If tb!Orders = "F" Or tb!Orders = "B" Then lstHaem.Selected(0) = True
260   End If

270   Exit Sub

FillKnownHaemOrders_Error:

      Dim strES As String
      Dim intEL As Integer

280   intEL = Erl
290   strES = Err.Description
300   LogError "frmNewOrder", "FillKnownHaemOrders", intEL, strES, sql

End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

10    On Error GoTo Form_MouseMove_Error

20    pBar = 0

30    Exit Sub

Form_MouseMove_Error:

      Dim strES As String
      Dim intEL As Integer



40    intEL = Erl
50    strES = Err.Description
60    LogError "frmNewOrder", "Form_MouseMove", intEL, strES


End Sub


Private Sub Form_Paint()

10    On Error GoTo Form_Paint_Error

20    If SysOptAlwaysRequestFBC(0) Then
30        lstHaem.Selected(0) = True
40    End If

50    Exit Sub

Form_Paint_Error:

      Dim strES As String
      Dim intEL As Integer



60    intEL = Erl
70    strES = Err.Description
80    LogError "frmNewOrder", "Form_Paint", intEL, strES


End Sub

Private Sub Form_Unload(Cancel As Integer)

10    On Error GoTo Form_Unload_Error

20    Activated = False

30    mSampleID = ""

40    SaveOptionSetting "CoagAnalyserDefault", lblCoagAnalyserName

50    Exit Sub

Form_Unload_Error:

      Dim strES As String
      Dim intEL As Integer

60    intEL = Erl
70    strES = Err.Description
80    LogError "frmNewOrder", "Form_Unload", intEL, strES

End Sub

Private Sub lAnalyserID_Click()

10    On Error GoTo lAnalyserID_Click_Error

20    lAnalyserID = IIf(lAnalyserID = "A", "B", "A")
30    AnalyserID = lAnalyserID

40    Exit Sub

lAnalyserID_Click_Error:

      Dim strES As String
      Dim intEL As Integer



50    intEL = Erl
60    strES = Err.Description
70    LogError "frmNewOrder", "lAnalyserID_Click", intEL, strES


End Sub

Private Sub lblCoagAnalyserName_Click()

      Dim n As Integer
      Dim c As Integer

10    On Error GoTo lblCoagAnalyserName_Click_Error

20    n = UBound(CoagAnalyserName)
30    For c = 0 To n
40        If CoagAnalyserName(c) = lblCoagAnalyserName Then
50            If c = n Then
60                lblCoagAnalyserName = CoagAnalyserName(0)
70            Else
80                lblCoagAnalyserName = CoagAnalyserName(c + 1)
90            End If
100           Exit For
110       End If
120   Next
130   FillKnownCoagOrders

140   Exit Sub

lblCoagAnalyserName_Click_Error:

      Dim strES As String
      Dim intEL As Integer

150   intEL = Erl
160   strES = Err.Description
170   LogError "frmNewOrder", "lblCoagAnalyserName_Click", intEL, strES

End Sub

Private Sub lCSfTests_Click()

10    On Error GoTo lCSfTests_Click_Error

20    pBar = 0

30    Exit Sub

lCSfTests_Click_Error:

      Dim strES As String
      Dim intEL As Integer



40    intEL = Erl
50    strES = Err.Description
60    LogError "frmNewOrder", "lCSfTests_Click", intEL, strES


End Sub

Private Sub lCSfTests_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

10    On Error GoTo lCSfTests_MouseUp_Error

20    BioChanged = True

30    Exit Sub

lCSfTests_MouseUp_Error:

      Dim strES As String
      Dim intEL As Integer



40    intEL = Erl
50    strES = Err.Description
60    LogError "frmNewOrder", "lCSfTests_MouseUp", intEL, strES


End Sub


Private Sub lEndoPanelPlasma_Click()

      Dim sql As String
      Dim tb As Recordset
      Dim T As Integer

10    On Error GoTo lEndoPanelPlasma_Click_Error



20    pBar = 0

30    sql = "SELECT * from EndPanels WHERE " & _
            "PanelName = '" & lEndoPanelPlasma.Text & "'" & _
            " and PanelType = 'PL' "
40    Set tb = New Recordset
50    RecOpenServer 0, tb, sql
60    Do While Not tb.EOF
70        For T = 0 To lEndoTestsPlasma.ListCount - 1
80            If optShort.Value = True Then
90                If UCase(EndLongNameFor(lEndoTestsPlasma.List(T))) = UCase((tb!Content)) Then
100                   lEndoTestsPlasma.Selected(T) = True
110                   Exit For
120               End If
130           Else
140               If UCase((lEndoTestsPlasma.List(T))) = UCase((tb!Content)) Then
150                   lEndoTestsPlasma.Selected(T) = True
160                   Exit For
170               End If
180           End If
190       Next
200       tb.MoveNext
210   Loop


220   Exit Sub

lEndoPanelPlasma_Click_Error:

      Dim strES As String
      Dim intEL As Integer

230   intEL = Erl
240   strES = Err.Description
250   LogError "frmNewOrder", "lEndoPanelPlasma_Click", intEL, strES, sql

End Sub

Private Sub lEndoPanelPlasma_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

10    On Error GoTo lEndoPanelPlasma_MouseUp_Error

20    BioChanged = True

30    Exit Sub

lEndoPanelPlasma_MouseUp_Error:

      Dim strES As String
      Dim intEL As Integer

40    intEL = Erl
50    strES = Err.Description
60    LogError "frmNewOrder", "lEndoPanelPlasma_MouseUp", intEL, strES

End Sub

Private Sub lEndoTestsPlasma_Click()

10    On Error GoTo lEndoTestsPlasma_Click_Error

20    pBar = 0

30    Exit Sub

lEndoTestsPlasma_Click_Error:

      Dim strES As String
      Dim intEL As Integer

40    intEL = Erl
50    strES = Err.Description
60    LogError "frmNewOrder", "lEndoTestsPlasma_Click", intEL, strES

End Sub

Private Sub lEndoTestsPlasma_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

10    On Error GoTo lEndoTestsPlasma_MouseUp_Error

20    BioChanged = True

30    Exit Sub

lEndoTestsPlasma_MouseUp_Error:

      Dim strES As String
      Dim intEL As Integer

40    intEL = Erl
50    strES = Err.Description
60    LogError "frmNewOrder", "lEndoTestsPlasma_MouseUp", intEL, strES

End Sub

Private Sub lImmunoPanel_Click()
      Dim T As Long
      Dim tb As New Recordset
      Dim sql As String

10    On Error GoTo lImmunoPanel_Click_Error

20    pBar = 0

30    sql = "SELECT * from IPanels WHERE " & _
            "PanelName = '" & lImmunoPanel.Text & "'"
40    Set tb = New Recordset
50    RecOpenServer 0, tb, sql
60    Do While Not tb.EOF
70        For T = 0 To lstImmunoTests.ListCount - 1
80            If UCase(QuickImmLongNameFor(lstImmunoTests.List(T))) = UCase(QuickImmLongNameFor(tb!Content)) Then
90                lstImmunoTests.Selected(T) = True
100               Exit For
110           End If
120       Next
130       tb.MoveNext
140   Loop

150   BioChanged = True

160   Exit Sub

lImmunoPanel_Click_Error:

      Dim strES As String
      Dim intEL As Integer

170   intEL = Erl
180   strES = Err.Description
190   LogError "frmNewOrder", "lImmunoPanel_Click", intEL, strES, sql

End Sub

Private Sub lImmunoPanel_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

10    On Error GoTo lImmunoPanel_MouseUp_Error

20    BioChanged = True

30    Exit Sub

lImmunoPanel_MouseUp_Error:

      Dim strES As String
      Dim intEL As Integer



40    intEL = Erl
50    strES = Err.Description
60    LogError "frmNewOrder", "lImmunoPanel_MouseUp", intEL, strES


End Sub

Private Sub lPlasmaPanel_Click()

      Dim sql As String
      Dim tb As Recordset
      Dim T As Integer

10    On Error GoTo lPlasmaPanel_Click_Error

20    pBar = 0

30    sql = "SELECT * from Panels WHERE " & _
            "PanelName = '" & lPlasmaPanel.Text & "'" & _
            " and PanelType = 'PL' "
40    Set tb = New Recordset
50    RecOpenServer 0, tb, sql
60    Do While Not tb.EOF
70        For T = 0 To lPlasmaTests.ListCount - 1
80            If optShort.Value = True Then
90                If UCase(BioLongNameFor(lPlasmaTests.List(T))) = UCase((tb!Content)) Then
100                   lPlasmaTests.Selected(T) = True
110                   Exit For
120               End If
130           Else
140               If UCase((lPlasmaTests.List(T))) = UCase((tb!Content)) Then
150                   lPlasmaTests.Selected(T) = True
160                   Exit For
170               End If
180           End If
190       Next
200       tb.MoveNext
210   Loop

220   Exit Sub

lPlasmaPanel_Click_Error:

      Dim strES As String
      Dim intEL As Integer

230   intEL = Erl
240   strES = Err.Description
250   LogError "frmNewOrder", "lPlasmaPanel_Click", intEL, strES, sql

End Sub

Private Sub lPlasmaPanel_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

10    On Error GoTo lPlasmaPanel_MouseUp_Error

20    BioChanged = True

30    Exit Sub

lPlasmaPanel_MouseUp_Error:

      Dim strES As String
      Dim intEL As Integer

40    intEL = Erl
50    strES = Err.Description
60    LogError "frmNewOrder", "lPlasmaPanel_MouseUp", intEL, strES

End Sub

Private Sub lPlasmaTests_Click()

10    On Error GoTo lPlasmaTests_Click_Error

20    pBar = 0

30    Exit Sub

lPlasmaTests_Click_Error:

      Dim strES As String
      Dim intEL As Integer

40    intEL = Erl
50    strES = Err.Description
60    LogError "frmNewOrder", "lPlasmaTests_Click", intEL, strES

End Sub

Private Sub lPlasmaTests_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

10    On Error GoTo lPlasmaTests_MouseUp_Error

20    BioChanged = True

30    Exit Sub

lPlasmaTests_MouseUp_Error:

      Dim strES As String
      Dim intEL As Integer

40    intEL = Erl
50    strES = Err.Description
60    LogError "frmNewOrder", "lPlasmaTests_MouseUp", intEL, strES

End Sub

Private Sub lSerumPanel_Click()

      Dim T As Long
      Dim tb As New Recordset
      Dim sql As String

10    On Error GoTo lSerumPanel_Click_Error

20    pBar = 0

30    sql = "SELECT * from Panels WHERE " & _
            "PanelName = '" & lSerumPanel.Text & "'" & _
            " and PanelType = '" & SysOptBioST(0) & "' "
40    Set tb = New Recordset
50    RecOpenServer 0, tb, sql
60    Do While Not tb.EOF
70        For T = 0 To lSerumTests.ListCount - 1
80            If optShort.Value = True Then
90                If UCase(BioLongNameFor(lSerumTests.List(T))) = UCase((tb!Content)) Then
100                   lSerumTests.Selected(T) = True
110                   Exit For
120               End If
130           Else
140               If UCase((lSerumTests.List(T))) = UCase((tb!Content)) Then
150                   lSerumTests.Selected(T) = True
160                   Exit For
170               End If
180           End If
190       Next
200       tb.MoveNext
210   Loop




220   Exit Sub

lSerumPanel_Click_Error:

      Dim strES As String
      Dim intEL As Integer



230   intEL = Erl
240   strES = Err.Description
250   LogError "frmNewOrder", "lSerumPanel_Click", intEL, strES, sql


End Sub

Private Sub lSerumPanel_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

10    On Error GoTo lSerumPanel_MouseUp_Error

20    BioChanged = True

30    Exit Sub

lSerumPanel_MouseUp_Error:

      Dim strES As String
      Dim intEL As Integer



40    intEL = Erl
50    strES = Err.Description
60    LogError "frmNewOrder", "lSerumPanel_MouseUp", intEL, strES


End Sub


Private Sub lSerumTests_Click()

10    On Error GoTo lSerumTests_Click_Error

20    pBar = 0

30    Exit Sub

lSerumTests_Click_Error:

      Dim strES As String
      Dim intEL As Integer



40    intEL = Erl
50    strES = Err.Description
60    LogError "frmNewOrder", "lSerumTests_Click", intEL, strES


End Sub


Private Sub lSerumTests_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

10    On Error GoTo lSerumTests_MouseUp_Error

20    BioChanged = True

30    Exit Sub

lSerumTests_MouseUp_Error:

      Dim strES As String
      Dim intEL As Integer



40    intEL = Erl
50    strES = Err.Description
60    LogError "frmNewOrder", "lSerumTests_MouseUp", intEL, strES


End Sub


Private Sub lstCoag_Click()

10    On Error GoTo lstCoag_Click_Error

20    pBar = 0

30    Exit Sub

lstCoag_Click_Error:

      Dim strES As String
      Dim intEL As Integer



40    intEL = Erl
50    strES = Err.Description
60    LogError "frmNewOrder", "lstCoag_Click", intEL, strES


End Sub

Private Sub lstCoag_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

10    On Error GoTo lstCoag_MouseUp_Error

20    CoagChanged = True

30    Exit Sub

lstCoag_MouseUp_Error:

      Dim strES As String
      Dim intEL As Integer



40    intEL = Erl
50    strES = Err.Description
60    LogError "frmNewOrder", "lstCoag_MouseUp", intEL, strES


End Sub




Private Sub lstEndoTests_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

10    On Error GoTo lstEndoTests_MouseUp_Error

20    BioChanged = True

30    Exit Sub

lstEndoTests_MouseUp_Error:

      Dim strES As String
      Dim intEL As Integer



40    intEL = Erl
50    strES = Err.Description
60    LogError "frmNewOrder", "lstEndoTests_MouseUp", intEL, strES


End Sub



Private Sub lstFluid_Click()

10    On Error GoTo lstFluid_Click_Error

20    pBar = 0

30    Exit Sub

lstFluid_Click_Error:

      Dim strES As String
      Dim intEL As Integer



40    intEL = Erl
50    strES = Err.Description
60    LogError "frmNewOrder", "lstFluid_Click", intEL, strES


End Sub

Private Sub lstFluid_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

10    On Error GoTo lstFluid_MouseUp_Error

20    BioChanged = True

30    Exit Sub

lstFluid_MouseUp_Error:

      Dim strES As String
      Dim intEL As Integer



40    intEL = Erl
50    strES = Err.Description
60    LogError "frmNewOrder", "lstFluid_MouseUp", intEL, strES


End Sub

Private Sub lstHaem_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
      Dim n As Long
      Dim Found As Long

10    On Error GoTo lstHaem_MouseUp_Error

20    HaemChanged = True

30    Found = 0

40    For n = 0 To lstHaem.ListCount - 1
50        If lstHaem.Selected(n) Then Found = 1
60    Next


70    Exit Sub

lstHaem_MouseUp_Error:

      Dim strES As String
      Dim intEL As Integer



80    intEL = Erl
90    strES = Err.Description
100   LogError "frmNewOrder", "lstHaem_MouseUp", intEL, strES


End Sub


Private Sub lstImmunoTests_Click()

10    On Error GoTo lstImmunoTests_Click_Error

20    pBar = 0

30    Exit Sub

lstImmunoTests_Click_Error:

      Dim strES As String
      Dim intEL As Integer



40    intEL = Erl
50    strES = Err.Description
60    LogError "frmNewOrder", "lstImmunoTests_Click", intEL, strES


End Sub

Private Sub lstImmunoTests_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

10    On Error GoTo lstImmunoTests_MouseUp_Error

20    BioChanged = True

30    Exit Sub

lstImmunoTests_MouseUp_Error:

      Dim strES As String
      Dim intEL As Integer



40    intEL = Erl
50    strES = Err.Description
60    LogError "frmNewOrder", "lstImmunoTests_MouseUp", intEL, strES


End Sub


Private Sub lUrinePanel_Click()

      Dim n As Long
      Dim tb As New Recordset
      Dim sql As String


10    On Error GoTo lUrinePanel_Click_Error

20    pBar = 0

      '30        For n = 0 To lUrineTests.ListCount - 1
      '40            lUrineTests.Selected(n) = False
      '50        Next

30    sql = "SELECT * from Panels WHERE " & _
            "PanelName = '" & lUrinePanel.Text & "'" & _
            "and PanelType = 'U' "
40    Set tb = New Recordset
50    RecOpenServer 0, tb, sql
60    Do While Not tb.EOF
70        For n = 0 To lUrineTests.ListCount - 1
80            If optShort.Value = True Then
90                If UCase(BioLongNameFor(lUrineTests.List(n))) = UCase((tb!Content)) Then
100                   lUrineTests.Selected(n) = True
110                   Exit For
120               End If
130           Else
140               If UCase((lUrineTests.List(n))) = UCase((tb!Content)) Then
150                   lUrineTests.Selected(n) = True
160                   Exit For
170               End If
180           End If
190       Next
200       tb.MoveNext
210   Loop

220   BioChanged = True



230   Exit Sub

lUrinePanel_Click_Error:

      Dim strES As String
      Dim intEL As Integer



240   intEL = Erl
250   strES = Err.Description
260   LogError "frmNewOrder", "lUrinePanel_Click", intEL, strES, sql


End Sub

Private Sub lUrinePanel_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

10    On Error GoTo lUrinePanel_MouseUp_Error

20    BioChanged = True

30    bsave.Enabled = True

40    Exit Sub

lUrinePanel_MouseUp_Error:

      Dim strES As String
      Dim intEL As Integer



50    intEL = Erl
60    strES = Err.Description
70    LogError "frmNewOrder", "lUrinePanel_MouseUp", intEL, strES


End Sub


Private Sub lUrinETests_Click()

10    On Error GoTo lUrinETests_Click_Error

20    pBar = 0

30    Exit Sub

lUrinETests_Click_Error:

      Dim strES As String
      Dim intEL As Integer



40    intEL = Erl
50    strES = Err.Description
60    LogError "frmNewOrder", "lUrinETests_Click", intEL, strES


End Sub


Private Sub lUrinETests_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

10    On Error GoTo lUrinETests_MouseUp_Error

20    BioChanged = True

30    Exit Sub

lUrinETests_MouseUp_Error:

      Dim strES As String
      Dim intEL As Integer



40    intEL = Erl
50    strES = Err.Description
60    LogError "frmNewOrder", "lUrinETests_MouseUp", intEL, strES


End Sub


Private Sub optLong_Click()

10    On Error GoTo optLong_Click_Error

20    SetDisplay "Long"

30    FillLists
40    FillKnownBioOrders

50    Exit Sub

optLong_Click_Error:

      Dim strES As String
      Dim intEL As Integer



60    intEL = Erl
70    strES = Err.Description
80    LogError "frmNewOrder", "optLong_Click", intEL, strES


End Sub

Private Sub optPlasma_Click()

10    On Error GoTo optPlasma_Click_Error

20    lSerumPanel.Visible = False
30    lSerumTests.Visible = False
40    lPlasmaPanel.Visible = True
50    lPlasmaTests.Visible = True
60    lEndoPanelPlasma.Visible = True
70    lEndoTestsPlasma.Visible = True
80    lEndoPanel.Visible = False
90    lstEndoTests.Visible = False
100   chkGBottle.Visible = True
110   Exit Sub

optPlasma_Click_Error:

      Dim strES As String
      Dim intEL As Integer

120   intEL = Erl
130   strES = Err.Description
140   LogError "frmNewOrder", "optPlasma_Click", intEL, strES


End Sub

Private Sub optSerum_Click()

10    On Error GoTo optSerum_Click_Error

20    lSerumPanel.Visible = True
30    lSerumTests.Visible = True
40    lPlasmaPanel.Visible = False
50    lPlasmaTests.Visible = False
60    lEndoPanel.Visible = True
70    lstEndoTests.Visible = True
80    lEndoPanelPlasma.Visible = False
90    lEndoTestsPlasma.Visible = False

100   chkGBottle.Visible = True


110   Exit Sub

optSerum_Click_Error:

      Dim strES As String
      Dim intEL As Integer

120   intEL = Erl
130   strES = Err.Description
140   LogError "frmNewOrder", "optSerum_Click", intEL, strES

End Sub

Private Sub optShort_Click()

10    On Error GoTo optShort_Click_Error

20    SetDisplay "Short"

30    FillLists
40    FillKnownBioOrders

50    Exit Sub

optShort_Click_Error:

      Dim strES As String
      Dim intEL As Integer



60    intEL = Erl
70    strES = Err.Description
80    LogError "frmNewOrder", "optShort_Click", intEL, strES


End Sub


Private Sub TimerBar_Timer()

10    On Error GoTo TimerBar_Timer_Error

20    If pBar >= (pBar.Max - 1) Then
30        Unload Me
40        Exit Sub
50    Else
60        pBar = pBar + 1
70    End If

80    Exit Sub

TimerBar_Timer_Error:

      Dim strES As String
      Dim intEL As Integer



90    intEL = Erl
100   strES = Err.Description
110   LogError "frmNewOrder", "TimerBar_Timer", intEL, strES


End Sub

Private Sub tinput_KeyPress(KeyAscii As Integer)

10    On Error GoTo tinput_KeyPress_Error

20    If KeyAscii = 13 Then
30        KeyAscii = 0
40        tinput_LostFocus
50    End If

60    Exit Sub

tinput_KeyPress_Error:

      Dim strES As String
      Dim intEL As Integer



70    intEL = Erl
80    strES = Err.Description
90    LogError "frmNewOrder", "tinput_KeyPress", intEL, strES


End Sub


Private Sub tinput_LostFocus()


10    On Error GoTo tinput_LostFocus_Error

20    If Trim$(tinput) = "" Then Exit Sub

30    tinput = UCase$(Trim$(tinput))

40    If Not CheckCodes() Then
50        If Not CheckSerumPanel() Then
60            If Not CheckPlasmaPanel() Then
70                If Not CheckESerumPanel() Then
80                    If Not CheckEPlasmaPanel() Then
90                        If Not CheckUrinePanel() Then
100                           If Not CheckImmSerumPanel() Then
110                               If Not CheckSerum() Then
120                                   If Not CheckPlasma() Then
130                                       If Not CheckESerum() Then
140                                           If Not CheckUrine() Then
150                                               If CheckCSF() Then
160                                                   BioChanged = True
170                                               End If
180                                           End If
190                                       End If
200                                   End If
210                               End If
220                           End If
230                       End If
240                   End If
250               End If
260           End If
270       End If
280   End If
290   tinput = ""
300   If tinput.Visible Then
310       tinput.SetFocus
320   End If




330   Exit Sub

tinput_LostFocus_Error:

      Dim strES As String
      Dim intEL As Integer



340   intEL = Erl
350   strES = Err.Description
360   LogError "frmNewOrder", "tinput_LostFocus", intEL, strES


End Sub

Private Sub tsampleid_GotFocus()

10    On Error GoTo tsampleid_GotFocus_Error

20    If tSampleID = "" Then ClearRequests
30    If SysOptAlwaysRequestFBC(0) Then
40        lstHaem.Selected(0) = True
50        lstHaem.Refresh
60    End If

70    Exit Sub

tsampleid_GotFocus_Error:

      Dim strES As String
      Dim intEL As Integer



80    intEL = Erl
90    strES = Err.Description
100   LogError "frmNewOrder", "tsampleid_GotFocus", intEL, strES


End Sub


Private Sub tsampleid_KeyPress(KeyAscii As Integer)

10    On Error GoTo tsampleid_KeyPress_Error

20    If KeyAscii = 13 Then
30        KeyAscii = 0
40        tinput.SetFocus
50    End If

60    Exit Sub

tsampleid_KeyPress_Error:

      Dim strES As String
      Dim intEL As Integer



70    intEL = Erl
80    strES = Err.Description
90    LogError "frmNewOrder", "tsampleid_KeyPress", intEL, strES


End Sub


Private Sub tsampleid_LostFocus()

      Dim sql As String
      Dim tb As Recordset

10    On Error GoTo tsampleid_LostFocus_Error

20    If Trim$(tSampleID) = "" Then Exit Sub

30    If Not IsNumeric(tSampleID) Then
40        iMsg "Sample Id must be Numeric!"
50        tSampleID = ""
60        tSampleID.SetFocus
70        Exit Sub
80    End If
90    tSampleID = Val(tSampleID)

100   tSampleID = Trim$(tSampleID)

110   If SysOptUrgent(0) Then
120       sql = "SELECT * FROM Demographics WHERE " & _
                "SampleID = '" & tSampleID & "'"

130       Set tb = New Recordset
140       RecOpenServer 0, tb, sql
150       If Not tb.EOF Then
160           If tb!Urgent = 1 Then chkUrgent.Value = 1
170       End If
180   End If

190   mSampleID = tSampleID

200   FillKnownBioOrders
210   FillKnownCoagOrders
220   FillKnownHaemOrders
230   FillKnownHaeOrders
240   LoadComments

250   CoagChanged = False
260   HaemChanged = False: HaemChanged = False
270   BioChanged = False
280   ImmunoChanged = False

290   Exit Sub

tsampleid_LostFocus_Error:

      Dim strES As String
      Dim intEL As Integer



300   intEL = Erl
310   strES = Err.Description
320   LogError "frmNewOrder", "tsampleid_LostFocus", intEL, strES, sql

End Sub

Private Function FillQuickBioNames() As String

      Dim tb As Recordset
      Dim sql As String
      Dim UB As Long

10    On Error GoTo FillQuickBioNames_Error

20    ReDim Preserve QuickBioNames(0 To 0)

30    sql = "SELECT LongName, ShortName from BioTestDefinitions"
40    Set tb = New Recordset
50    RecOpenServer 0, tb, sql
60    Do While Not tb.EOF
70        UB = UBound(QuickBioNames) + 1
80        ReDim Preserve QuickBioNames(0 To UB)
90        QuickBioNames(UB).Short = tb!ShortName & ""
100       QuickBioNames(UB).Long = tb!LongName & ""
110       tb.MoveNext
120   Loop

130   Exit Function

FillQuickBioNames_Error:

      Dim strES As String
      Dim intEL As Integer

140   intEL = Erl
150   strES = Err.Description
160   LogError "frmNewOrder", "FillQuickBioNames", intEL, strES, sql

End Function



Private Function FillQuickImmNames() As String

      Dim tb As Recordset
      Dim sql As String
      Dim UB As Long

10    On Error GoTo FillQuickImmNames_Error

20    ReDim Preserve QuickImmNames(0 To 0)

30    sql = "SELECT LongName, ShortName from IMMTestDefinitions"
40    Set tb = New Recordset
50    RecOpenServer 0, tb, sql
60    Do While Not tb.EOF
70        UB = UBound(QuickImmNames) + 1
80        ReDim Preserve QuickImmNames(0 To UB)
90        QuickImmNames(UB).Short = tb!ShortName & ""
100       QuickImmNames(UB).Long = tb!LongName & ""
110       tb.MoveNext
120   Loop

130   Exit Function

FillQuickImmNames_Error:

      Dim strES As String
      Dim intEL As Integer

140   intEL = Erl
150   strES = Err.Description
160   LogError "frmNewOrder", "FillQuickImmNames", intEL, strES, sql

End Function

Private Function QuickImmLongNameFor(ByVal LongOrShortName As String) As String

      Dim n As Long

10    On Error GoTo QuickImmLongNameFor_Error

20    For n = 1 To UBound(QuickImmNames)
30        If LongOrShortName = QuickImmNames(n).Long Or LongOrShortName = QuickImmNames(n).Short Then
40            QuickImmLongNameFor = QuickImmNames(n).Long
50            Exit For
60        End If
70    Next

80    Exit Function

QuickImmLongNameFor_Error:

      Dim strES As String
      Dim intEL As Integer

90    intEL = Erl
100   strES = Err.Description
110   LogError "frmNewOrder", "QuickImmLongNameFor", intEL, strES

End Function



Private Sub lEndoPanel_Click()
      Dim T As Long
      Dim tb As New Recordset
      Dim sql As String

10    On Error GoTo lEndoPanel_Click_Error

20    pBar = 0

30    sql = "SELECT * from EndPanels WHERE " & _
            "PanelName = '" & lEndoPanel.Text & "'"
40    Set tb = New Recordset
50    RecOpenServer 0, tb, sql
60    Do While Not tb.EOF
70        For T = 0 To lstEndoTests.ListCount - 1
80            If UCase(QuickEndLongNameFor(lstEndoTests.List(T))) = UCase(QuickEndLongNameFor(tb!Content)) Then
90                lstEndoTests.Selected(T) = True
100               Exit For
110           End If
120       Next
130       tb.MoveNext
140   Loop

150   BioChanged = True


160   Exit Sub

lEndoPanel_Click_Error:

      Dim strES As String
      Dim intEL As Integer



170   intEL = Erl
180   strES = Err.Description
190   LogError "frmNewOrder", "lEndoPanel_Click", intEL, strES, sql


End Sub

Private Sub lEndoPanel_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

10    On Error GoTo lEndoPanel_MouseUp_Error

20    BioChanged = True

30    Exit Sub

lEndoPanel_MouseUp_Error:

      Dim strES As String
      Dim intEL As Integer



40    intEL = Erl
50    strES = Err.Description
60    LogError "frmNewOrder", "lEndoPanel_MouseUp", intEL, strES


End Sub


Private Function FillQuickEndNames() As String

      Dim tb As Recordset
      Dim sql As String
      Dim UB As Long

10    On Error GoTo FillQuickEndNames_Error

20    ReDim Preserve QuickEndNames(0 To 0)

30    sql = "SELECT LongName, ShortName from EndTestDefinitions"
40    Set tb = New Recordset
50    RecOpenServer 0, tb, sql
60    Do While Not tb.EOF
70        UB = UBound(QuickEndNames) + 1
80        ReDim Preserve QuickEndNames(0 To UB)
90        QuickEndNames(UB).Short = tb!ShortName & ""
100       QuickEndNames(UB).Long = tb!LongName & ""
110       tb.MoveNext
120   Loop

130   Exit Function

FillQuickEndNames_Error:

      Dim strES As String
      Dim intEL As Integer

140   intEL = Erl
150   strES = Err.Description
160   LogError "frmNewOrder", "FillQuickEndNames", intEL, strES, sql

End Function

Private Function QuickEndLongNameFor(ByVal LongOrShortName As String) As String

      Dim n As Long

10    On Error GoTo QuickEndLongNameFor_Error

20    For n = 1 To UBound(QuickEndNames)
30        If LongOrShortName = QuickEndNames(n).Long Or LongOrShortName = QuickEndNames(n).Short Then
40            QuickEndLongNameFor = QuickEndNames(n).Long
50            Exit For
60        End If
70    Next

80    Exit Function

QuickEndLongNameFor_Error:

      Dim strES As String
      Dim intEL As Integer

90    intEL = Erl
100   strES = Err.Description
110   LogError "frmNewOrder", "QuickEndLongNameFor", intEL, strES

End Function

Function CheckESerum() As Boolean

      Dim Y As Long
      Dim tb As New Recordset
      Dim sql As String
      Dim LongOrShort As String

10    On Error GoTo CheckESerum_Error

20    LongOrShort = IIf(optLong, "Long", "Short")

30    CheckESerum = False
40    sql = "SELECT " & LongOrShort & "Name as Name from EndTestDefinitions WHERE " & _
            "SampleType = '" & SysOptBioST(0) & "' " & _
            "and BarCode = '" & tinput & "' and knowntoanalyser = '1'"
50    Set tb = New Recordset
60    RecOpenServer 0, tb, sql
70    If Not tb.EOF Then
80        CheckESerum = True
90        For Y = 0 To lstEndoTests.ListCount - 1
100           If lstEndoTests.List(Y) = tb!Name Then
110               lstEndoTests.Selected(Y) = Not lstEndoTests.Selected(Y)
120               Exit For
130           End If
140       Next
150       BioChanged = True
160   End If


170   Exit Function

CheckESerum_Error:

      Dim strES As String
      Dim intEL As Integer



180   intEL = Erl
190   strES = Err.Description
200   LogError "frmNewOrder", "CheckESerum", intEL, strES, sql


End Function


Private Sub txtImmComment_Change(Index As Integer)
10    On Error GoTo txtImmComment_Change_Error

      ' If bValidateImm(Index).Caption = "VALID" Then Exit Sub

      '30        If Index = 0 Then
      '40            cmdSaveImm(0).Enabled = True
      '50        Else
      '60            cmdSaveImm(1).Enabled = True
      '70        End If

20    Exit Sub

txtImmComment_Change_Error:

      Dim strES As String
      Dim intEL As Integer

30    intEL = Erl
40    strES = Err.Description
50    LogError "frmNewOrder", "txtImmComment_Change", intEL, strES
End Sub
Private Sub txtImmComment_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

      Dim sql As String
      Dim tb As New Recordset
      Dim s As Variant
      Dim n As Long
      Dim z As Long

10    On Error GoTo txtImmComment_KeyDown_Error

      'If bValidateImm(Index).Caption = "VALID" Then Exit Sub

20    If Index = 0 Then

30        If KeyCode = vbKeyF2 Then
40            If Len(Trim(txtImmComment(0))) < 3 Then Exit Sub
50            n = txtImmComment(0).SelStart
60            z = 3
70            s = Mid(txtImmComment(0), (n - z) + 1, z + 1)
80            z = 3
90            If ListText("EN", s) <> "" Then
100               s = ListText("EN", s)
110           Else
120               s = ""
130           End If

140           If s = "" Then
150               z = 1
160               s = Mid(txtImmComment(0), n - z, z + 1)
170               z = 2
180               If ListText("EN", s) <> "" Then
190                   s = ListText("EN", s)
200               Else
210                   s = ""
220               End If
230           End If

240           If s = "" Then
250               z = 1
260               s = Mid(txtImmComment(0), n, z)

270               If ListText("EN", s) <> "" Then
280                   s = ListText("EN", s)
290               End If
300           End If
310           txtImmComment(0) = Left(txtImmComment(0), (n - (z)))
320           txtImmComment(0) = txtImmComment(0) & s

330           txtImmComment(0).SelStart = Len(txtImmComment(0))

340       ElseIf KeyCode = vbKeyF3 Then

350           sql = "SELECT * from lists WHERE listtype = 'EN' order by listorder"
360           Set tb = New Recordset
370           RecOpenServer 0, tb, sql
380           Do While Not tb.EOF
390               s = Trim(tb!Text)
400               frmMessages.lstComm.AddItem s
410               tb.MoveNext
420           Loop

430           Set frmMessages.f = Me
440           Set frmMessages.T = txtImmComment(0)
450           frmMessages.Show 1

460       End If

          'cmdSaveImm(0).Enabled = True
470   Else

480       If KeyCode = vbKeyF2 Then
490           If Len(Trim(txtImmComment(1))) < 2 Then Exit Sub
500           n = txtImmComment(1).SelStart
510           If n < 3 Then Exit Sub
520           s = UCase(Mid(txtImmComment(1), (n - 2), 3))
530           If ListText("IM", s) <> "" Then
540               s = ListText("IM", s)
550           End If
560           txtImmComment(1) = Left(txtImmComment(1), (n) - 3)
570           txtImmComment(1) = txtImmComment(1) & s
580           txtImmComment(1).SelStart = Len(txtImmComment(1))
590       ElseIf KeyCode = vbKeyF3 Then
600           sql = "SELECT * from lists WHERE listtype = 'IM'"
610           Set tb = New Recordset
620           RecOpenServer 0, tb, sql
630           Do While Not tb.EOF
640               s = Trim(tb!Text)
650               frmMessages.lstComm.AddItem s
660               tb.MoveNext
670           Loop
680           Set frmMessages.f = Me
690           Set frmMessages.T = txtImmComment(1)
700           frmMessages.Show 1

710       End If
          'cmdSaveImm(1).Enabled = True
720   End If

730   Exit Sub

txtImmComment_KeyDown_Error:

      Dim strES As String
      Dim intEL As Integer

740   intEL = Erl
750   strES = Err.Description
760   LogError "frmNewOrder", "txtImmComment_KeyDown", intEL, strES, sql

End Sub
Private Sub txtImmComment_KeyPress(Index As Integer, KeyAscii As Integer)

10    On Error GoTo txtImmComment_KeyPress_Error

20    KeyAscii = VI(KeyAscii, AlphaNumeric_NoApos)

30    Exit Sub

txtImmComment_KeyPress_Error:

      Dim strES As String
      Dim intEL As Integer

40    intEL = Erl
50    strES = Err.Description
60    LogError "frmNewOrder", "txtImmComment_KeyPress", intEL, strES


End Sub

Private Sub UpDown1_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

      Dim sql As String
      Dim tb As Recordset

10    On Error GoTo UpDown1_MouseUp_Error

20    ClearRequests

30    If SysOptUrgent(0) Then
40        sql = "SELECT * FROM Demographics WHERE " & _
                "SampleID = '" & tSampleID & "'"

50        Set tb = New Recordset
60        RecOpenServer 0, tb, sql
70        If Not tb.EOF Then
80            If tb!Urgent = 1 Then chkUrgent.Value = 1
90        End If
100   End If

110   mSampleID = tSampleID

120   FillKnownBioOrders
130   FillKnownCoagOrders
140   FillKnownHaemOrders
150   FillKnownHaeOrders
160   LoadComments

170   CoagChanged = False
180   HaemChanged = False: HaeChanged = False
190   BioChanged = False
200   ImmunoChanged = False

210   Exit Sub

UpDown1_MouseUp_Error:

      Dim strES As String
      Dim intEL As Integer



220   intEL = Erl
230   strES = Err.Description
240   LogError "frmNewOrder", "UpDown1_MouseUp", intEL, strES, sql

End Sub
Private Sub LoadComments()

      Dim Ob As Observation
      Dim Obs As Observations

10    On Error GoTo LoadComments_Error

20    txtBioComment = ""
30    txtImmComment(0) = ""


40    If Trim$(mSampleID) = "" Then Exit Sub

50    Set Obs = New Observations
60    Set Obs = Obs.Load(mSampleID, "Biochemistry", "Demographic", "Haematology", "Coagulation", _
                         "Immunology", "Endocrinology", "BloodGas")
70    If Not Obs Is Nothing Then
80        For Each Ob In Obs
90            Select Case UCase$(Ob.Discipline)
                  Case "BIOCHEMISTRY": txtBioComment = Split_Comm(Ob.Comment)
100               Case "ENDOCRINOLOGY": txtImmComment(0) = Split_Comm(Ob.Comment)
110           End Select
120       Next
130   End If

140   Exit Sub

LoadComments_Error:

      Dim strES As String
      Dim intEL As Integer

150   intEL = Erl
160   strES = Err.Description
170   LogError "frmnewOrder", "LoadComments", intEL, strES

End Sub

