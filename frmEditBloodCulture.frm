VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEditBloodCulture 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Microbiology"
   ClientHeight    =   9960
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   16650
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9960
   ScaleWidth      =   16650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame6 
      Caption         =   "Copies"
      Height          =   855
      Left            =   11340
      TabIndex        =   111
      Top             =   8670
      Width           =   1395
      Begin VB.TextBox txtNoCopies 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   60
         Locked          =   -1  'True
         TabIndex        =   112
         Text            =   "1"
         Top             =   300
         Width           =   330
      End
      Begin ComCtl2.UpDown udNoCopies 
         Height          =   360
         Left            =   390
         TabIndex        =   113
         Top             =   270
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   635
         _Version        =   327681
         Value           =   1
         BuddyControl    =   "txtNoCopies"
         BuddyDispid     =   196610
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
      Begin VB.Label lblFinal 
         Alignment       =   2  'Center
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "F"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   690
         TabIndex        =   115
         ToolTipText     =   "Print Final Report"
         Top             =   300
         Width           =   300
      End
      Begin VB.Label lblInterim 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "I"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   990
         TabIndex        =   114
         ToolTipText     =   "Print Interim Report"
         Top             =   300
         Width           =   300
      End
   End
   Begin VB.Frame fraBC 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1785
      Index           =   1
      Left            =   30
      TabIndex        =   104
      Top             =   7770
      Width           =   10965
      Begin VB.TextBox txtConC 
         Height          =   1005
         Left            =   5850
         MultiLine       =   -1  'True
         TabIndex        =   108
         Top             =   630
         Width           =   5250
      End
      Begin VB.TextBox txtMSC 
         Height          =   1005
         Left            =   420
         MultiLine       =   -1  'True
         TabIndex        =   107
         Top             =   630
         Width           =   4935
      End
      Begin VB.ComboBox cmbMSC 
         Height          =   315
         Left            =   2430
         TabIndex        =   106
         Text            =   "cmbMSC"
         Top             =   300
         Width           =   2925
      End
      Begin VB.ComboBox cmbConC 
         Height          =   315
         Left            =   7410
         TabIndex        =   105
         Text            =   "cmbConC"
         Top             =   300
         Width           =   3435
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Medical Scientist Comments"
         Height          =   195
         Index           =   0
         Left            =   420
         TabIndex        =   110
         Top             =   390
         Width           =   1980
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Consultant Comments"
         Height          =   195
         Index           =   1
         Left            =   5850
         TabIndex        =   109
         Top             =   390
         Width           =   1530
      End
   End
   Begin VB.Frame fraBC 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   6435
      Index           =   0
      Left            =   0
      TabIndex        =   31
      Top             =   1260
      Width           =   16605
      Begin VB.ComboBox cmbOrgGroup 
         BackColor       =   &H0080FFFF&
         Height          =   315
         Index           =   3
         Left            =   5880
         TabIndex        =   79
         Text            =   "cmbOrgGroup"
         Top             =   480
         Width           =   2205
      End
      Begin VB.ComboBox cmbOrgGroup 
         BackColor       =   &H0080FFFF&
         Height          =   315
         Index           =   2
         Left            =   3150
         TabIndex        =   78
         Text            =   "cmbOrgGroup"
         Top             =   480
         Width           =   2205
      End
      Begin VB.ComboBox cmbOrgGroup 
         BackColor       =   &H0080FFFF&
         Height          =   315
         Index           =   1
         Left            =   420
         TabIndex        =   77
         Text            =   "cmbOrgGroup"
         Top             =   480
         Width           =   2205
      End
      Begin VB.ComboBox cmbOrgGroup 
         BackColor       =   &H0080FFFF&
         Height          =   315
         Index           =   4
         Left            =   8640
         TabIndex        =   76
         Text            =   "cmbOrgGroup"
         Top             =   480
         Width           =   2205
      End
      Begin VB.ComboBox cmbOrgName 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Index           =   1
         Left            =   420
         TabIndex        =   75
         Text            =   "cmbOrgName"
         Top             =   810
         Width           =   2205
      End
      Begin VB.ComboBox cmbOrgName 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Index           =   2
         Left            =   3150
         TabIndex        =   74
         Text            =   "cmbOrgName"
         Top             =   810
         Width           =   2205
      End
      Begin VB.ComboBox cmbOrgName 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Index           =   3
         Left            =   5880
         TabIndex        =   73
         Text            =   "cmbOrgName"
         Top             =   810
         Width           =   2205
      End
      Begin VB.ComboBox cmbOrgName 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Index           =   4
         Left            =   8640
         TabIndex        =   72
         Text            =   "cmbOrgName"
         Top             =   810
         Width           =   2205
      End
      Begin VB.ComboBox cmbABSelect 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Index           =   1
         IntegralHeight  =   0   'False
         Left            =   420
         TabIndex        =   71
         Text            =   "cmbABSelect"
         Top             =   5970
         Width           =   2205
      End
      Begin VB.ComboBox cmbABSelect 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Index           =   2
         IntegralHeight  =   0   'False
         Left            =   3150
         TabIndex        =   70
         Text            =   "cmbABSelect"
         Top             =   5970
         Width           =   2205
      End
      Begin VB.ComboBox cmbABSelect 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Index           =   3
         IntegralHeight  =   0   'False
         Left            =   5880
         TabIndex        =   69
         Text            =   "cmbABSelect"
         Top             =   5970
         Width           =   2205
      End
      Begin VB.ComboBox cmbABSelect 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Index           =   4
         IntegralHeight  =   0   'False
         Left            =   8640
         TabIndex        =   68
         Text            =   "cmbABSelect"
         Top             =   5970
         Width           =   2205
      End
      Begin VB.CommandButton cmdRemoveSecondary 
         Height          =   525
         Index           =   1
         Left            =   60
         Picture         =   "frmEditBloodCulture.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   67
         ToolTipText     =   "Remove Secondary Lists"
         Top             =   3240
         Width           =   315
      End
      Begin VB.CommandButton cmdUseSecondary 
         Height          =   525
         Index           =   1
         Left            =   60
         Picture         =   "frmEditBloodCulture.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   66
         ToolTipText     =   "Use Secondary Lists"
         Top             =   3780
         Width           =   315
      End
      Begin VB.CommandButton cmdRemoveSecondary 
         Height          =   525
         Index           =   2
         Left            =   2790
         Picture         =   "frmEditBloodCulture.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   65
         ToolTipText     =   "Remove Secondary Lists"
         Top             =   3240
         Width           =   315
      End
      Begin VB.CommandButton cmdUseSecondary 
         Height          =   525
         Index           =   2
         Left            =   2790
         Picture         =   "frmEditBloodCulture.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   64
         ToolTipText     =   "Use Secondary Lists"
         Top             =   3780
         Width           =   315
      End
      Begin VB.CommandButton cmdRemoveSecondary 
         Height          =   525
         Index           =   3
         Left            =   5520
         Picture         =   "frmEditBloodCulture.frx":0C28
         Style           =   1  'Graphical
         TabIndex        =   63
         ToolTipText     =   "Remove Secondary Lists"
         Top             =   3240
         Width           =   315
      End
      Begin VB.CommandButton cmdUseSecondary 
         Height          =   525
         Index           =   3
         Left            =   5520
         Picture         =   "frmEditBloodCulture.frx":0F32
         Style           =   1  'Graphical
         TabIndex        =   62
         ToolTipText     =   "Use Secondary Lists"
         Top             =   3780
         Width           =   315
      End
      Begin VB.CommandButton cmdRemoveSecondary 
         Height          =   525
         Index           =   4
         Left            =   8280
         Picture         =   "frmEditBloodCulture.frx":123C
         Style           =   1  'Graphical
         TabIndex        =   61
         ToolTipText     =   "Remove Secondary Lists"
         Top             =   3240
         Width           =   315
      End
      Begin VB.CommandButton cmdUseSecondary 
         Height          =   525
         Index           =   4
         Left            =   8280
         Picture         =   "frmEditBloodCulture.frx":1546
         Style           =   1  'Graphical
         TabIndex        =   60
         ToolTipText     =   "Use Secondary Lists"
         Top             =   3780
         Width           =   315
      End
      Begin VB.CheckBox chkNonReportable 
         DownPicture     =   "frmEditBloodCulture.frx":1850
         Height          =   345
         Index           =   1
         Left            =   120
         Picture         =   "frmEditBloodCulture.frx":1DDA
         Style           =   1  'Graphical
         TabIndex        =   59
         ToolTipText     =   "Check to make culture non-reportable"
         Top             =   810
         Width           =   285
      End
      Begin VB.CheckBox chkNonReportable 
         DownPicture     =   "frmEditBloodCulture.frx":2364
         Height          =   345
         Index           =   2
         Left            =   2850
         Picture         =   "frmEditBloodCulture.frx":28EE
         Style           =   1  'Graphical
         TabIndex        =   58
         ToolTipText     =   "Check to make culture non-reportable"
         Top             =   810
         Width           =   285
      End
      Begin VB.CheckBox chkNonReportable 
         DownPicture     =   "frmEditBloodCulture.frx":2E78
         Height          =   345
         Index           =   3
         Left            =   5580
         Picture         =   "frmEditBloodCulture.frx":3402
         Style           =   1  'Graphical
         TabIndex        =   57
         ToolTipText     =   "Check to make culture non-reportable"
         Top             =   810
         Width           =   285
      End
      Begin VB.CheckBox chkNonReportable 
         DownPicture     =   "frmEditBloodCulture.frx":398C
         Height          =   345
         Index           =   4
         Left            =   8340
         Picture         =   "frmEditBloodCulture.frx":3F16
         Style           =   1  'Graphical
         TabIndex        =   56
         ToolTipText     =   "Check to make culture non-reportable"
         Top             =   810
         Width           =   285
      End
      Begin VB.ComboBox cmbOrgGroup 
         BackColor       =   &H0080FFFF&
         Height          =   315
         Index           =   5
         Left            =   11370
         TabIndex        =   55
         Text            =   "cmbOrgGroup"
         Top             =   480
         Width           =   2205
      End
      Begin VB.ComboBox cmbOrgName 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Index           =   5
         Left            =   11370
         TabIndex        =   54
         Text            =   "cmbOrgName"
         Top             =   810
         Width           =   2205
      End
      Begin VB.ComboBox cmbABSelect 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Index           =   5
         IntegralHeight  =   0   'False
         Left            =   11370
         TabIndex        =   53
         Text            =   "cmbABSelect"
         Top             =   5940
         Width           =   2205
      End
      Begin VB.CommandButton cmdRemoveSecondary 
         Height          =   525
         Index           =   5
         Left            =   11010
         Picture         =   "frmEditBloodCulture.frx":44A0
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "Remove Secondary Lists"
         Top             =   3270
         Width           =   315
      End
      Begin VB.CommandButton cmdUseSecondary 
         Height          =   525
         Index           =   5
         Left            =   11010
         Picture         =   "frmEditBloodCulture.frx":47AA
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Use Secondary Lists"
         Top             =   3810
         Width           =   315
      End
      Begin VB.CheckBox chkNonReportable 
         DownPicture     =   "frmEditBloodCulture.frx":4AB4
         Height          =   345
         Index           =   5
         Left            =   11070
         Picture         =   "frmEditBloodCulture.frx":503E
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Check to make culture non-reportable"
         Top             =   810
         Width           =   285
      End
      Begin VB.ComboBox cmbOrgGroup 
         BackColor       =   &H0080FFFF&
         Height          =   315
         Index           =   6
         Left            =   14070
         TabIndex        =   49
         Text            =   "cmbOrgGroup"
         Top             =   480
         Width           =   2205
      End
      Begin VB.ComboBox cmbOrgName 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Index           =   6
         Left            =   14070
         TabIndex        =   48
         Text            =   "cmbOrgName"
         Top             =   810
         Width           =   2205
      End
      Begin VB.ComboBox cmbABSelect 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Index           =   6
         IntegralHeight  =   0   'False
         Left            =   14070
         TabIndex        =   47
         Text            =   "cmbABSelect"
         Top             =   5910
         Width           =   2205
      End
      Begin VB.CommandButton cmdRemoveSecondary 
         Height          =   525
         Index           =   6
         Left            =   13710
         Picture         =   "frmEditBloodCulture.frx":55C8
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Remove Secondary Lists"
         Top             =   3300
         Width           =   315
      End
      Begin VB.CommandButton cmdUseSecondary 
         Height          =   525
         Index           =   6
         Left            =   13710
         Picture         =   "frmEditBloodCulture.frx":58D2
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Use Secondary Lists"
         Top             =   3840
         Width           =   315
      End
      Begin VB.CheckBox chkNonReportable 
         DownPicture     =   "frmEditBloodCulture.frx":5BDC
         Height          =   345
         Index           =   6
         Left            =   13770
         Picture         =   "frmEditBloodCulture.frx":6166
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Check to make culture non-reportable"
         Top             =   810
         Width           =   285
      End
      Begin VB.CommandButton cmdReportAll 
         Height          =   285
         Index           =   1
         Left            =   2610
         Picture         =   "frmEditBloodCulture.frx":66F0
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Make All Reportable"
         Top             =   1740
         Width           =   285
      End
      Begin VB.CommandButton cmdReportNone 
         Height          =   285
         Index           =   1
         Left            =   2610
         Picture         =   "frmEditBloodCulture.frx":69C6
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Make All Non-Reportable"
         Top             =   2040
         Width           =   285
      End
      Begin VB.CommandButton cmdReportAll 
         Height          =   285
         Index           =   2
         Left            =   5370
         Picture         =   "frmEditBloodCulture.frx":6C9C
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Make All Reportable"
         Top             =   1740
         Width           =   285
      End
      Begin VB.CommandButton cmdReportNone 
         Height          =   285
         Index           =   2
         Left            =   5370
         Picture         =   "frmEditBloodCulture.frx":6F72
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Make All Non-Reportable"
         Top             =   2040
         Width           =   285
      End
      Begin VB.CommandButton cmdReportAll 
         Height          =   285
         Index           =   3
         Left            =   8100
         Picture         =   "frmEditBloodCulture.frx":7248
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Make All Reportable"
         Top             =   1740
         Width           =   285
      End
      Begin VB.CommandButton cmdReportNone 
         Height          =   285
         Index           =   3
         Left            =   8100
         Picture         =   "frmEditBloodCulture.frx":751E
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Make All Non-Reportable"
         Top             =   2040
         Width           =   285
      End
      Begin VB.CommandButton cmdReportAll 
         Height          =   285
         Index           =   4
         Left            =   10830
         Picture         =   "frmEditBloodCulture.frx":77F4
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Make All Reportable"
         Top             =   1740
         Width           =   285
      End
      Begin VB.CommandButton cmdReportNone 
         Height          =   285
         Index           =   4
         Left            =   10830
         Picture         =   "frmEditBloodCulture.frx":7ACA
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Make All Non-Reportable"
         Top             =   2040
         Width           =   285
      End
      Begin VB.CommandButton cmdReportAll 
         Height          =   285
         Index           =   5
         Left            =   13560
         Picture         =   "frmEditBloodCulture.frx":7DA0
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Make All Reportable"
         Top             =   1740
         Width           =   285
      End
      Begin VB.CommandButton cmdReportNone 
         Height          =   285
         Index           =   5
         Left            =   13560
         Picture         =   "frmEditBloodCulture.frx":8076
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Make All Non-Reportable"
         Top             =   2040
         Width           =   285
      End
      Begin VB.CommandButton cmdReportAll 
         Height          =   285
         Index           =   6
         Left            =   16260
         Picture         =   "frmEditBloodCulture.frx":834C
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Make All Reportable"
         Top             =   1740
         Width           =   285
      End
      Begin VB.CommandButton cmdReportNone 
         Height          =   285
         Index           =   6
         Left            =   16260
         Picture         =   "frmEditBloodCulture.frx":8622
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Make All Non-Reportable"
         Top             =   2040
         Width           =   285
      End
      Begin MSFlexGridLib.MSFlexGrid grdAB 
         Height          =   4515
         Index           =   3
         Left            =   5880
         TabIndex        =   80
         Top             =   1440
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   7964
         _Version        =   393216
         Cols            =   7
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
         FormatString    =   "<Antibiotic             |^S/R|>Rprt|<Result  |<Date/Time            |<Operator"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid grdAB 
         Height          =   4515
         Index           =   2
         Left            =   3150
         TabIndex        =   81
         Top             =   1440
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   7964
         _Version        =   393216
         Cols            =   6
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
         FormatString    =   "<Antibiotic             |^S/R|>Rprt|<Result  |<Date/Time            |<Operator"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid grdAB 
         Height          =   4515
         Index           =   1
         Left            =   420
         TabIndex        =   82
         Top             =   1440
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   7964
         _Version        =   393216
         Cols            =   6
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
         FormatString    =   "<Antibiotic             |^S/R|>Rprt|<Result  |<Date/Time            |<Operator"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid grdAB 
         Height          =   4515
         Index           =   4
         Left            =   8640
         TabIndex        =   83
         Top             =   1440
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   7964
         _Version        =   393216
         Cols            =   7
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
         FormatString    =   "<Antibiotic             |^S/R|>Rprt|<Result  |<Date/Time            |<Operator"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid grdAB 
         Height          =   4485
         Index           =   5
         Left            =   11370
         TabIndex        =   84
         Top             =   1440
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   7911
         _Version        =   393216
         Cols            =   7
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
         FormatString    =   "<Antibiotic             |^S/R|>Rprt|<Result  |<Date/Time            |<Operator"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid grdAB 
         Height          =   4455
         Index           =   6
         Left            =   14070
         TabIndex        =   85
         Top             =   1440
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   7858
         _Version        =   393216
         Cols            =   7
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
         FormatString    =   "<Antibiotic             |^S/R|>Rprt|<Result  |<Date/Time            |<Operator"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblSetAllR 
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   120
         TabIndex        =   103
         ToolTipText     =   "Set All Resistant"
         Top             =   4320
         Width           =   270
      End
      Begin VB.Label lblSetAllS 
         AutoSize        =   -1  'True
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   120
         TabIndex        =   102
         ToolTipText     =   "Set All Sensitive"
         Top             =   4710
         Width           =   255
      End
      Begin VB.Label lblSetAllS 
         AutoSize        =   -1  'True
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   2850
         TabIndex        =   101
         ToolTipText     =   "Set All Sensitive"
         Top             =   4710
         Width           =   255
      End
      Begin VB.Label lblSetAllR 
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   2850
         TabIndex        =   100
         ToolTipText     =   "Set All Resistant"
         Top             =   4320
         Width           =   270
      End
      Begin VB.Label lblSetAllS 
         AutoSize        =   -1  'True
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   5580
         TabIndex        =   99
         ToolTipText     =   "Set All Sensitive"
         Top             =   4710
         Width           =   255
      End
      Begin VB.Label lblSetAllR 
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   5580
         TabIndex        =   98
         ToolTipText     =   "Set All Resistant"
         Top             =   4320
         Width           =   270
      End
      Begin VB.Label lblSetAllS 
         AutoSize        =   -1  'True
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   8340
         TabIndex        =   97
         ToolTipText     =   "Set All Sensitive"
         Top             =   4710
         Width           =   255
      End
      Begin VB.Label lblSetAllR 
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   8340
         TabIndex        =   96
         ToolTipText     =   "Set All Resistant"
         Top             =   4320
         Width           =   270
      End
      Begin VB.Image imgSquareCross 
         Height          =   225
         Left            =   270
         Picture         =   "frmEditBloodCulture.frx":88F8
         Top             =   6210
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Image imgSquareTick 
         Height          =   225
         Left            =   60
         Picture         =   "frmEditBloodCulture.frx":8BCE
         Top             =   6210
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Label lblSetAllS 
         AutoSize        =   -1  'True
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         Left            =   11070
         TabIndex        =   95
         ToolTipText     =   "Set All Sensitive"
         Top             =   4740
         Width           =   255
      End
      Begin VB.Label lblSetAllR 
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         Left            =   11070
         TabIndex        =   94
         ToolTipText     =   "Set All Resistant"
         Top             =   4350
         Width           =   270
      End
      Begin VB.Label lblSetAllS 
         AutoSize        =   -1  'True
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   6
         Left            =   13770
         TabIndex        =   93
         ToolTipText     =   "Set All Sensitive"
         Top             =   4770
         Width           =   255
      End
      Begin VB.Label lblSetAllR 
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   6
         Left            =   13770
         TabIndex        =   92
         ToolTipText     =   "Set All Resistant"
         Top             =   4380
         Width           =   270
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "-- Aerobic --"
         Height          =   285
         Index           =   0
         Left            =   420
         TabIndex        =   91
         Top             =   180
         Width           =   4935
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "-- Anaerobic --"
         Height          =   285
         Index           =   1
         Left            =   5880
         TabIndex        =   90
         Top             =   180
         Width           =   4965
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "-- FAN --"
         Height          =   285
         Index           =   2
         Left            =   11370
         TabIndex        =   89
         Top             =   180
         Width           =   4905
      End
      Begin VB.Label lblBC 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF80&
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
         Height          =   285
         Index           =   1
         Left            =   420
         TabIndex        =   88
         Top             =   1140
         Width           =   4935
      End
      Begin VB.Label lblBC 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
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
         Height          =   285
         Index           =   3
         Left            =   5880
         TabIndex        =   87
         Top             =   1140
         Width           =   4965
      End
      Begin VB.Label lblBC 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
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
         Height          =   285
         Index           =   5
         Left            =   11370
         TabIndex        =   86
         Top             =   1140
         Width           =   4905
      End
   End
   Begin VB.Frame fraSampleID 
      Height          =   1095
      Left            =   510
      TabIndex        =   9
      Top             =   150
      Width           =   14265
      Begin VB.Label lblGP 
         BackColor       =   &H80000018&
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
         Height          =   255
         Left            =   8910
         TabIndex        =   30
         Top             =   660
         Width           =   2385
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "GP"
         Height          =   195
         Left            =   8640
         TabIndex        =   29
         Top             =   690
         Width           =   225
      End
      Begin VB.Label lblClinician 
         BackColor       =   &H80000018&
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
         Height          =   255
         Left            =   6000
         TabIndex        =   28
         Top             =   690
         Width           =   2505
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Clinician"
         Height          =   195
         Left            =   5310
         TabIndex        =   27
         Top             =   720
         Width           =   585
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Ward"
         Height          =   195
         Left            =   690
         TabIndex        =   26
         Top             =   720
         Width           =   420
      End
      Begin VB.Label lblABsInUse 
         BorderStyle     =   1  'Fixed Single
         Height          =   645
         Left            =   11580
         TabIndex        =   25
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label lblWard 
         BackColor       =   &H80000018&
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
         Height          =   255
         Left            =   1170
         TabIndex        =   24
         Top             =   690
         Width           =   4035
      End
      Begin VB.Label lblSex 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
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
         Height          =   255
         Left            =   10560
         TabIndex        =   23
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblAge 
         BackColor       =   &H80000018&
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
         Height          =   255
         Left            =   9690
         TabIndex        =   22
         Top             =   240
         Width           =   795
      End
      Begin VB.Label lblDoB 
         BackColor       =   &H80000018&
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
         Height          =   255
         Left            =   8550
         TabIndex        =   21
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblAandE 
         BackColor       =   &H80000018&
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
         Height          =   255
         Left            =   6900
         TabIndex        =   20
         Top             =   240
         Width           =   1605
      End
      Begin VB.Label lblChart 
         BackColor       =   &H80000018&
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
         Height          =   255
         Left            =   5250
         TabIndex        =   19
         Top             =   240
         Width           =   1605
      End
      Begin VB.Label lblName 
         BackColor       =   &H80000018&
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
         Height          =   255
         Left            =   1170
         TabIndex        =   18
         Top             =   240
         Width           =   4035
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Sex"
         Height          =   195
         Index           =   90
         Left            =   10770
         TabIndex        =   17
         Top             =   -30
         Width           =   270
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Age"
         Height          =   195
         Index           =   89
         Left            =   9930
         TabIndex        =   16
         Top             =   -30
         Width           =   285
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "D.o.B"
         Height          =   195
         Index           =   88
         Left            =   8910
         TabIndex        =   15
         Top             =   0
         Width           =   405
      End
      Begin VB.Label lblNameTitle 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         Height          =   195
         Left            =   2700
         TabIndex        =   14
         Top             =   0
         Width           =   420
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "A and E"
         Height          =   195
         Index           =   0
         Left            =   7350
         TabIndex        =   13
         Top             =   -30
         Width           =   570
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Chart"
         Height          =   195
         Left            =   5760
         TabIndex        =   12
         Top             =   0
         Width           =   375
      End
      Begin VB.Label lblSampleID 
         BackColor       =   &H80000018&
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
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1005
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Sample ID"
         Height          =   195
         Index           =   92
         Left            =   510
         TabIndex        =   10
         Top             =   0
         Width           =   735
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   9585
      Width           =   16650
      _ExtentX        =   29369
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   5
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "User Name"
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            TextSave        =   "31/05/2024"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Todays Date"
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   4410
            MinWidth        =   4410
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Demographic Check"
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
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
   Begin VB.CommandButton cmdViewReports 
      Caption         =   "Reports"
      Height          =   800
      Left            =   12900
      Picture         =   "frmEditBloodCulture.frx":8EA4
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "View Printed && Faxed Reports"
      Top             =   7830
      Width           =   900
   End
   Begin VB.CommandButton cmdValidateMicro 
      Caption         =   "&Validate"
      Height          =   800
      Left            =   14220
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmEditBloodCulture.frx":91AE
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8730
      Width           =   900
   End
   Begin VB.CommandButton cmdSaveMicro 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   800
      Left            =   14220
      Picture         =   "frmEditBloodCulture.frx":95F0
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7830
      Width           =   900
   End
   Begin VB.CommandButton cmdSetPrinter 
      Height          =   800
      Left            =   11580
      Picture         =   "frmEditBloodCulture.frx":9C5A
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Back Colour Normal:Automatic Printer Selection.Back Colour Red:-Forced"
      Top             =   7830
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Timer TimerBar 
      Interval        =   1000
      Left            =   12360
      Top             =   0
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   135
      Left            =   570
      TabIndex        =   3
      Top             =   0
      Width           =   14205
      _ExtentX        =   25056
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   1
      Max             =   600
      Scrolling       =   1
   End
   Begin VB.CommandButton bHistory 
      Caption         =   "&History"
      Height          =   800
      Left            =   15420
      Picture         =   "frmEditBloodCulture.frx":9F64
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   7830
      Width           =   900
   End
   Begin VB.CommandButton bPrint 
      Caption         =   "&Print"
      Height          =   800
      Left            =   12900
      Picture         =   "frmEditBloodCulture.frx":A3A6
      Style           =   1  'Graphical
      TabIndex        =   1
      Tag             =   "bprint"
      Top             =   8730
      Width           =   900
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   800
      Left            =   15420
      Picture         =   "frmEditBloodCulture.frx":AA10
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8730
      Width           =   900
   End
   Begin VB.Menu mnuLists 
      Caption         =   "&Lists"
      Begin VB.Menu mnuMSC 
         Caption         =   "&Medical Scientist Comment"
      End
      Begin VB.Menu mnuConsultantComment 
         Caption         =   "&Consultant Comment"
      End
   End
End
Attribute VB_Name = "frmEditBloodCulture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pPrintToPrinter As String

Private SampleIDWithOffset As Double

Private ForceSaveability As Boolean

Private BacTek3DInUse As Boolean
Private ObservaInUse As Boolean

Private Sub AdjustNegative(ByVal Index As Integer)

      Dim sql As String
      Dim tb As Recordset

10    On Error GoTo AdjustNegative_Error

20    sql = "SELECT * FROM Isolates WHERE " & _
            "SampleID = '" & SampleIDWithOffset & "' " & _
            "AND IsolateNumber = '" & Index & "'"
30    Set tb = New Recordset
40    RecOpenClient 0, tb, sql
50    If tb.EOF Then
60        sql = "INSERT INTO Isolates " & _
                "([SampleID], [IsolateNumber], [OrganismGroup], [OrganismName]) " & _
                "VALUES " & _
                "('" & SampleIDWithOffset & "', " & _
                " '" & Index & "', " & _
                " 'Negative Results', " & _
                " 'Sterile after 5 days') "
70        Cnxn(0).Execute sql

80        cmbOrgGroup(Index) = "Negative Results"
90        cmbOrgName(Index) = "Sterile after 5 days"
100   End If

110   Exit Sub

AdjustNegative_Error:

      Dim strES As String
      Dim intEL As Integer

120   intEL = Erl
130   strES = Err.Description
140   LogError "frmEditBloodCulture", "AdjustNegative", intEL, strES, sql

End Sub

Private Sub AdjustOrganism()

      Dim sql As String
      Dim tb As Recordset


150   On Error GoTo AdjustOrganism_Error

      'QMS Ref #818120

160   sql = "SELECT I.SampleID, I.IsolateNumber FROM Isolates I " & _
            "Inner Join Sensitivities S " & _
            "On I.SampleID = S.SampleID " & _
            "And I.IsolateNumber = S.IsolateNumber " & _
            "WHERE I.SampleID = '" & SampleIDWithOffset & "' " & _
            "AND I.OrganismName = 'Staphylococcus aureus' " & _
            "AND S.AntibioticCode = 'OXA' And S.RSI = 'R'"

170   Set tb = New Recordset
180   RecOpenClient 0, tb, sql
190   If Not tb.EOF Then
200       While Not tb.EOF
210           sql = "Update Isolates Set " & _
                    "OrganismName = 'Staphylococcus aureus (MRSA)' " & _
                    "Where SampleID = '" & tb!SampleID & "' " & _
                    "AND IsolateNumber = " & tb!IsolateNumber
220           Cnxn(0).Execute sql
230           cmbOrgName(tb!IsolateNumber) = "Staphylococcus aureus (MRSA)"
240           tb.MoveNext
250       Wend
260   End If



270   Exit Sub

AdjustOrganism_Error:

      Dim strES As String
      Dim intEL As Integer

280   intEL = Erl
290   strES = Err.Description
300   LogError "frmEditBloodCulture", "AdjustOrganism", intEL, strES, sql

End Sub


Private Sub LoadListMSComment()

      Dim tb As Recordset
      Dim sql As String

310   On Error GoTo LoadListMSComment_Error

320   cmbMSC.Clear

330   sql = "Select * from Lists where " & _
            "ListType = 'MSComment' " & _
            "ORDER BY ListOrder"
340   Set tb = New Recordset
350   RecOpenServer 0, tb, sql
360   Do While Not tb.EOF
370       cmbMSC.AddItem tb!Text & ""
380       tb.MoveNext
390   Loop

400   Exit Sub

LoadListMSComment_Error:

      Dim strES As String
      Dim intEL As Integer

410   intEL = Erl
420   strES = Err.Description
430   LogError "frmEditBloodCulture", "LoadListMSComment", intEL, strES, sql

End Sub

Private Sub LoadListConsultantComment()

      Dim tb As Recordset
      Dim sql As String

440   On Error GoTo LoadListConsultantComment_Error

450   cmbConC.Clear

460   sql = "Select * from Lists where " & _
            "ListType = 'ConsComment' " & _
            "ORDER BY ListOrder"
470   Set tb = New Recordset
480   RecOpenServer 0, tb, sql
490   Do While Not tb.EOF
500       cmbConC.AddItem tb!Text & ""
510       tb.MoveNext
520   Loop

530   Exit Sub

LoadListConsultantComment_Error:

      Dim strES As String
      Dim intEL As Integer

540   intEL = Erl
550   strES = Err.Description
560   LogError "frmEditBloodCulture", "LoadListConsultantComment", intEL, strES, sql

End Sub

Private Sub SetComboWidths()

      Dim n As Integer

570   On Error GoTo SetComboWidths_Error

580   For n = 1 To 6
590       SetComboDropDownWidth cmbOrgGroup(n)
600       SetComboDropDownWidth cmbOrgName(n)
610   Next

620   Exit Sub

SetComboWidths_Error:

      Dim strES As String
      Dim intEL As Integer

630   intEL = Erl
640   strES = Err.Description
650   LogError "frmEditBloodCulture", "SetComboWidths", intEL, strES

End Sub
Private Sub chkNonReportable_Click(Index As Integer)

660   cmdSaveMicro.Enabled = True

End Sub

Private Sub cmbMSC_Click()

670   If txtMSC <> "" Then
680       txtMSC = txtMSC & " "
690   End If
700   txtMSC = txtMSC & cmbMSC

710   cmdSaveMicro.Enabled = True

End Sub


Private Sub FillABSelect(ByVal Index As Integer)

      Dim tb As Recordset
      Dim sql As String
      Dim n As Integer
      Dim ExcludeList As String
      Dim T As Single

720   On Error GoTo FillABSelect_Error

730   cmbABSelect(Index).Clear

740   ExcludeList = ""
750   For n = 1 To grdAB(Index).Rows - 1
760       ExcludeList = ExcludeList & _
                        "AntibioticName <> '" & grdAB(Index).TextMatrix(n, 0) & "' and "
770   Next
780   ExcludeList = Left$(ExcludeList, Len(ExcludeList) - 4)

790   sql = "SELECT DISTINCT RTRIM(AntibioticName) AS AntibioticName, ListOrder " & _
            "FROM Antibiotics WHERE " & _
            ExcludeList & _
            "ORDER BY ListOrder"

800   Set tb = New Recordset
810   RecOpenServer 0, tb, sql
820   T = Timer
830   With cmbABSelect(Index)
840       Do While Not tb.EOF
850           .AddItem tb!AntibioticName & ""
860           tb.MoveNext
870       Loop
880   End With

890   Exit Sub

FillABSelect_Error:

      Dim strES As String
      Dim intEL As Integer

900   intEL = Erl
910   strES = Err.Description
920   LogError "frmEditBloodCulture", "FillABSelect", intEL, strES, sql

End Sub

Private Sub FillMSandConsultantComment()

930   LoadListMSComment
940   LoadListConsultantComment

End Sub

Private Sub FillOrgNames(ByVal Index As Integer)

      Dim tb As Recordset
      Dim sql As String

950   On Error GoTo FillOrgNames_Error

960   cmbOrgName(Index).Clear

970   If cmbOrgGroup(Index).Text = "Negative Results" Then
980       sql = "Select * from Organisms where " & _
                "GroupName = '" & cmbOrgGroup(Index).Text & "' " & _
                "AND Site = 'Blood Culture' " & _
                "order by ListOrder"
990   Else
1000      sql = "Select Distinct Name, ListOrder from Organisms where " & _
                "GroupName = '" & cmbOrgGroup(Index).Text & "' " & _
                "order by ListOrder"
1010  End If
1020  Set tb = New Recordset
1030  RecOpenClient 0, tb, sql
1040  Do While Not tb.EOF
1050      cmbOrgName(Index).AddItem tb!Name & ""
1060      tb.MoveNext
1070  Loop

1080  SetComboWidths

1090  Exit Sub

FillOrgNames_Error:

      Dim strES As String
      Dim intEL As Integer

1100  intEL = Erl
1110  strES = Err.Description
1120  LogError "frmEditBloodCulture", "FillOrgNames", intEL, strES, sql

End Sub


Private Sub GetSampleIDWithOffset()

1130  SampleIDWithOffset = Val(lblSampleID) + SysOptMicroOffset(0)

End Sub


Private Sub LoadSensitivitiesForced(ByVal Index As Integer)

      Dim tb As Recordset
      Dim sql As String
      Dim Report As Boolean

1140  On Error GoTo LoadSensitivitiesForced_Error

1150  sql = "SELECT LTRIM(RTRIM(A.AntibioticName)) AS AntibioticName, " & _
            "S.Report, S.RSI, S.CPOFlag, S.Result, S.RunDateTime, S.UserName " & _
            "FROM Sensitivities S, Antibiotics A " & _
            "WHERE SampleID = '" & SampleIDWithOffset & "' " & _
            "AND IsolateNumber = '" & Index & "' " & _
            "AND S.AntibioticCode = A.Code " & _
            "AND S.Forced = 1"
1160  Set tb = New Recordset
1170  RecOpenServer 0, tb, sql
1180  Do While Not tb.EOF
1190      With grdAB(Index)
1200          .AddItem tb!AntibioticName & vbTab & _
                       tb!RSI & vbTab & _
                       tb!CPOFlag & vbTab & _
                       tb!Result & vbTab & _
                       Format(tb!RunDateTime, "dd/mm/yy hh:mm") & _
                       tb!UserName & ""
1210          .Row = .Rows - 1
1220          .Col = 2
1230          If IsNull(tb!Report) Then
1240              Set .CellPicture = Me.Picture
1250          Else
1260              Set .CellPicture = IIf(tb!Report, imgSquareTick.Picture, imgSquareCross.Picture)
1270          End If

1280          .Col = 0
1290          .CellBackColor = &HFFFFC0

1300          tb.MoveNext
1310      End With
1320  Loop

1330  Exit Sub

LoadSensitivitiesForced_Error:

      Dim strES As String
      Dim intEL As Integer

1340  intEL = Erl
1350  strES = Err.Description
1360  LogError "frmEditBloodCulture", "LoadSensitivitiesForced", intEL, strES, sql

End Sub

Private Sub LoadSensitivitiesSecondary(ByVal Index As Integer)

      Dim tb As Recordset
      Dim sql As String
      Dim Report As Boolean

1370  On Error GoTo LoadSensitivitiesSecondary_Error

1380  sql = "SELECT LTRIM(RTRIM(A.AntibioticName)) AS AntibioticName, " & _
            "S.Report, S.RSI, S.CPOFlag, S.Result, S.RunDateTime, S.UserName " & _
            "FROM Sensitivities S, Antibiotics A " & _
            "WHERE SampleID = '" & SampleIDWithOffset & "' " & _
            "AND IsolateNumber = '" & Index & "' " & _
            "AND S.AntibioticCode = A.Code " & _
            "AND S.Secondary = 1"
1390  Set tb = New Recordset
1400  RecOpenServer 0, tb, sql
1410  Do While Not tb.EOF
1420      With grdAB(Index)
1430          .AddItem tb!AntibioticName & vbTab & _
                       tb!RSI & vbTab & _
                       tb!CPOFlag & vbTab & _
                       tb!Result & vbTab & _
                       Format(tb!RunDateTime, "dd/mm/yy hh:mm") & _
                       tb!UserName & ""
1440          .Row = .Rows - 1
1450          .Col = 2
1460          If IsNull(tb!Report) Then
1470              Set .CellPicture = Me.Picture
1480          Else
1490              Set .CellPicture = IIf(tb!Report, imgSquareTick.Picture, imgSquareCross.Picture)
1500          End If

1510          .Col = 0
1520          .CellFontBold = True

1530          tb.MoveNext
1540      End With
1550  Loop

1560  Exit Sub

LoadSensitivitiesSecondary_Error:

      Dim strES As String
      Dim intEL As Integer

1570  intEL = Erl
1580  strES = Err.Description
1590  LogError "frmEditBloodCulture", "LoadSensitivitiesSecondary", intEL, strES, sql

End Sub

Private Function LoadIsolates() As Boolean
      'returns true if loaded
      Dim tb As Recordset
      Dim sql As String
      Dim intIsolate As Integer

1600  On Error GoTo LoadIsolates_Error

1610  LoadIsolates = False

1620  For intIsolate = 1 To 6
1630      cmbOrgGroup(intIsolate) = ""
1640      cmbOrgName(intIsolate) = ""
1650      chkNonReportable(intIsolate).Value = 0
1660  Next

1670  sql = "Select * from Isolates where " & _
            "SampleID = '" & SampleIDWithOffset & "'"
1680  Set tb = New Recordset
1690  RecOpenClient 0, tb, sql
1700  Do While Not tb.EOF
1710      LoadIsolates = True
1720      cmbOrgGroup(tb!IsolateNumber) = tb!OrganismGroup & ""
1730      cmbOrgName(tb!IsolateNumber) = tb!OrganismName & ""
1740      chkNonReportable(tb!IsolateNumber) = IIf(IsNull(tb!NonReportable), 0, tb!NonReportable)
1750      tb.MoveNext
1760  Loop

1770  Exit Function

LoadIsolates_Error:

      Dim strES As String
      Dim intEL As Integer

1780  intEL = Erl
1790  strES = Err.Description
1800  LogError "frmEditBloodCulture", "LoadIsolates", intEL, strES, sql

End Function


Private Sub PrintThis()

      Dim tb As Recordset
      Dim sql As String
      Dim FinalOrInterim As String

1810  On Error GoTo PrintThis_Error

1820  pBar = 0
1830  GetSampleIDWithOffset

1840  sql = "SELECT * FROM PrintPending WHERE " & _
            "Department = 'N' " & _
            "AND SampleID = '" & Val(lblSampleID) + SysOptMicroOffset(0) & "'"
1850  Set tb = New Recordset
1860  RecOpenClient 0, tb, sql
1870  If tb.EOF Then
1880      tb.AddNew
1890  End If
1900  tb!SampleID = Val(lblSampleID) + SysOptMicroOffset(0)
1910  tb!Ward = lblWard
1920  tb!Clinician = lblClinician
1930  tb!GP = lblGP
1940  tb!Department = "N"
1950  tb!Initiator = UserName
1960  tb!UsePrinter = pPrintToPrinter
1970  tb!NoOfCopies = Val(txtNoCopies)
1980  FinalOrInterim = "F"
1990  If lblInterim.BackColor = vbGreen Then
2000      FinalOrInterim = "I"
2010  End If
2020  tb!FinalInterim = FinalOrInterim
2030  tb.Update

2040  Exit Sub

PrintThis_Error:

      Dim strES As String
      Dim intEL As Integer

2050  intEL = Erl
2060  strES = Err.Description
2070  LogError "frmEditBloodCulture", "PrintThis", intEL, strES, sql

End Sub

Private Sub SaveComments()

      Dim Obs As New Observations

2080  On Error GoTo SaveComments_Error

2090  If txtMSC = "Medical Scientist Comments" Or Trim$(txtMSC) = "" Then
2100      Obs.Save SampleIDWithOffset, True, "MicroCS", ""
2110  Else
2120      Obs.Save SampleIDWithOffset, True, "MicroCS", Trim$(txtMSC)
2130  End If

2140  If txtConC = "Consultant Comments" Or Trim$(txtConC) = "" Then
2150      Obs.Save SampleIDWithOffset, True, "MicroConsultant", ""
2160  Else
2170      Obs.Save SampleIDWithOffset, True, "MicroConsultant", Trim$(txtConC)
2180  End If

2190  Exit Sub

SaveComments_Error:

      Dim strES As String
      Dim intEL As Integer

2200  intEL = Erl
2210  strES = Err.Description
2220  LogError "frmEditBloodCulture", "SaveComments", intEL, strES

End Sub

Private Sub SaveIsolates()

      Dim sql As String
      Dim intIsolate As Integer

2230  On Error GoTo SaveIsolates_Error

2240  For intIsolate = 1 To 6
2250      If cmbOrgGroup(intIsolate) <> "" Then

2260          sql = "IF EXISTS(select SampleID from Isolates " & _
                    "Where SampleID = '@SampleID' and IsolateNumber = @IsolateNumber) " & _
                    "BEGIN " & _
                    "Update Isolates SET " & _
                    "OrganismGroup = '@OrganismGroup', " & _
                    "OrganismName = '@OrganismName', " & _
                    "UserName = '@UserName', " & _
                    "NonReportable = @NonReportable " & _
                    "Where " & _
                    "SampleID = '@SampleID' and " & _
                    "IsolateNumber = @IsolateNumber " & _
                    "End " & _
                    "Else " & _
                    "BEGIN " & _
                    "INSERT INTO Isolates( " & _
                    "SampleID, " & _
                    "IsolateNumber, " & _
                    "OrganismGroup, " & _
                    "OrganismName, " & _
                    "UserName, " & _
                    "NonReportable) "
2270          sql = sql & _
                    "VALUES( " & _
                    "'@SampleID', " & _
                    "@IsolateNumber, " & _
                    "'@OrganismGroup', " & _
                    "'@OrganismName', " & _
                    "'@UserName', " & _
                    "@NonReportable) " & _
                    "End "
2280          sql = Replace(sql, "@SampleID", SampleIDWithOffset)
2290          sql = Replace(sql, "@IsolateNumber", intIsolate)
2300          sql = Replace(sql, "@OrganismGroup", cmbOrgGroup(intIsolate))
2310          sql = Replace(sql, "@OrganismName", cmbOrgName(intIsolate))
2320          sql = Replace(sql, "@UserName", UserName)
2330          sql = Replace(sql, "@NonReportable", chkNonReportable(intIsolate).Value)
2340          Cnxn(0).Execute sql
              '        sql = "SELECT * FROM Isolates WHERE " & _
                       '            "SampleID = '" & SampleIDWithOffset & "' " & _
                       '            "AND IsolateNumber = '" & intIsolate & "'"
              '        Set tb = New Recordset
              '        RecOpenServer 0, tb, sql
              '        If tb.EOF Then
              '            tb.AddNew
              '            tb!SampleID = SampleIDWithOffset
              '            tb!IsolateNumber = intIsolate
              '        End If
              '        tb!OrganismGroup = cmbOrgGroup(intIsolate)
              '        tb!OrganismName = cmbOrgName(intIsolate)
              '        tb!UserName = UserName
              '        tb!NonReportable = chkNonReportable(intIsolate).Value
              '        tb.Update
2350      Else

2360          sql = "DELETE FROM Isolates WHERE " & _
                    "SampleID = '" & SampleIDWithOffset & "' " & _
                    "AND IsolateNumber = '" & intIsolate & "'"
2370          Cnxn(0).Execute sql

2380          sql = "DELETE FROM Sensitivities WHERE " & _
                    "SampleID = '" & SampleIDWithOffset & "' " & _
                    "AND IsolateNumber = '" & intIsolate & "'"
2390          Cnxn(0).Execute sql

2400      End If
2410  Next

2420  Exit Sub

SaveIsolates_Error:

      Dim strES As String
      Dim intEL As Integer

2430  intEL = Erl
2440  strES = Err.Description
2450  LogError "frmEditBloodCulture", "SaveIsolates", intEL, strES, sql

End Sub

Private Sub SetAsForced(ByVal intIndex As Integer, _
                        ByVal strABName As String, _
                        ByVal blnReport As Boolean)

      Dim tb As Recordset
      Dim sql As String

2460  On Error GoTo SetAsForced_Error

      'Created on 08/10/2010 12:18:36
      'Autogenerated by SQL Scripting

2470  sql = "If Exists(Select 1 From ForcedABReport " & _
            "Where SampleID = @SampleID0 And ABName = '@ABName1' And [Index] = @Index3 ) " & _
            "Begin " & _
            "Update ForcedABReport Set " & _
            "SampleID = @SampleID0, ABName = '@ABName1', Report = @Report2, [Index] = @Index3 " & _
            "Where SampleID = @SampleID0 And ABName = '@ABName1' And [Index] = @Index3  " & _
            "End  " & _
            "Else " & _
            "Begin  " & _
            "Insert Into ForcedABReport (SampleID, ABName, Report, [Index]) Values (@SampleID0, '@ABName1', @Report2, @Index3) " & _
            "End"

2480  sql = Replace(sql, "@SampleID0", SysOptMicroOffset(0) + Val(lblSampleID))
2490  sql = Replace(sql, "@ABName1", strABName)
2500  sql = Replace(sql, "@Report2", IIf(blnReport, 1, 0))
2510  sql = Replace(sql, "@Index3", intIndex)


2520  Cnxn(0).Execute sql


2530  sql = "Select * from ForcedABReport where " & _
            "ABName = '" & strABName & "' " & _
            "and [Index] = " & intIndex & " " & _
            "and SampleID = " & SysOptMicroOffset(0) + Val(lblSampleID)
2540  Set tb = New Recordset
2550  RecOpenServer 0, tb, sql
2560  If tb.EOF Then
2570      tb.AddNew
2580  End If
2590  tb!SampleID = SysOptMicroOffset(0) + Val(lblSampleID)
2600  tb!ABName = strABName
2610  tb!Report = blnReport
2620  tb!Index = intIndex
2630  tb.Update

2640  Exit Sub

SetAsForced_Error:

      Dim strES As String
      Dim intEL As Integer

2650  intEL = Erl
2660  strES = Err.Description
2670  LogError "frmEditBloodCulture", "SetAsForced", intEL, strES, sql

End Sub

Private Sub cmbABSelect_Click(Index As Integer)

      Dim sql As String
      Dim tb As Recordset
      Dim Y As Integer

2680  grdAB(Index).AddItem cmbABSelect(Index).Text
2690  grdAB(Index).Row = grdAB(Index).Rows - 1
2700  grdAB(Index).Col = 0
2710  grdAB(Index).CellBackColor = &HFFFFC0
2720  grdAB(Index).Col = 2
2730  Set grdAB(Index).CellPicture = Me.Picture

2740  sql = "Select distinct * from Sensitivities as S, Antibiotics as A where " & _
            "SampleID = '" & SampleIDWithOffset & "' " & _
            "and IsolateNumber = '" & Index & "' " & _
            "and S.AntibioticCode = A.Code " & _
            "and AntibioticName = '" & cmbABSelect(Index).Text & "'"
2750  Set tb = New Recordset
2760  RecOpenClient 0, tb, sql
2770  If Not tb.EOF Then

2780      With grdAB(Index)
2790          Y = .Rows - 1
2800          .Row = Y
2810          .TextMatrix(Y, 1) = tb!RSI & ""
2820          .TextMatrix(Y, 2) = tb!CPOFlag & ""
2830          .TextMatrix(Y, 3) = tb!Result & ""
2840          .TextMatrix(Y, 4) = Format(tb!RunDateTime, "dd/mm/yy hh:mm")
2850          .TextMatrix(Y, 5) = tb!UserName & ""
2860          .Col = 2
2870          If IsNull(tb!Report) Then
2880              Set .CellPicture = Me.Picture
2890          Else
2900              Set .CellPicture = IIf(tb!Report, imgSquareTick.Picture, imgSquareCross.Picture)
2910          End If
2920      End With

2930  End If

2940  cmbABSelect(Index) = ""

2950  FillABSelect Index

2960  cmdSaveMicro.Enabled = True

End Sub

Private Sub cmbABSelect_KeyPress(Index As Integer, KeyAscii As Integer)

2970  KeyAscii = 0

End Sub

Private Sub cmbConC_Click()

2980  If txtConC <> "" Then
2990      txtConC = txtConC & " "
3000  End If
3010  txtConC = txtConC & cmbConC

3020  cmdSaveMicro.Enabled = True

End Sub


Private Sub cmbOrgGroup_Click(Index As Integer)

3030  FillAbGrid Index
3040  FillABSelect Index
3050  FillOrgNames Index

3060  cmdSaveMicro.Enabled = True
3070  grdAB(Index).Visible = True

End Sub

Private Sub cmbOrgGroup_LostFocus(Index As Integer)

      Dim tb As Recordset
      Dim sql As String

3080  sql = "Select * from Lists where " & _
            "ListType = 'OR' " & _
            "and Code = '" & cmbOrgGroup(Index) & "'"
3090  Set tb = New Recordset
3100  RecOpenServer 0, tb, sql
3110  If Not tb.EOF Then
3120      cmbOrgGroup(Index) = tb!Text & ""
3130  End If

End Sub


Private Sub cmbOrgName_Click(Index As Integer)

3140  cmdSaveMicro.Enabled = True

End Sub

Private Sub cmbOrgName_LostFocus(Index As Integer)

      Dim tb As Recordset
      Dim sql As String

3150  sql = "SELECT Name FROM Organisms WHERE " & _
            "Code = '" & AddTicks(cmbOrgName(Index)) & "'"
3160  Set tb = New Recordset
3170  RecOpenServer 0, tb, sql
3180  If Not tb.EOF Then
3190      cmbOrgName(Index) = tb!Name & ""
3200  End If

End Sub


Private Sub cmdRemoveSecondary_Click(Index As Integer)

      Dim n As Integer

3210  grdAB(Index).Col = 0
3220  For n = grdAB(Index).Rows - 1 To 1 Step -1
3230      grdAB(Index).Row = n
3240      If grdAB(Index).CellFontBold = True Then
3250          DeleteSensitivity Index, grdAB(Index).TextMatrix(n, 0)
3260          If n = 1 Then
3270              grdAB(Index).AddItem ""
3280          End If
3290          grdAB(Index).RemoveItem n
3300      End If
3310  Next

3320  FillABSelect Index

3330  Exit Sub

End Sub

Private Sub cmdReportAll_Click(Index As Integer)

      Dim n As Integer

3340  With grdAB(Index)
3350      .Col = 2
3360      For n = 1 To .Rows - 1
3370          If .TextMatrix(n, 0) <> "" Then
3380              .Row = n
3390              Set .CellPicture = imgSquareTick.Picture
3400          End If
3410      Next
3420  End With

End Sub

Private Sub cmdReportNone_Click(Index As Integer)

      Dim n As Integer

3430  With grdAB(Index)
3440      .Col = 2
3450      For n = 1 To .Rows - 1
3460          If .TextMatrix(n, 0) <> "" Then
3470              .Row = n
3480              Set .CellPicture = imgSquareCross.Picture
3490          End If
3500      Next
3510  End With

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

3520  On Error GoTo cmdUseSecondary_Click_Error

3530  sql = "Select Distinct A.AntibioticName, D.ListOrder, " & _
            "A.AllowIfPregnant, A.AllowIfOutPatient, A.AllowIfChild " & _
            "from ABDefinitions as D, Antibiotics as A where " & _
            "D.OrganismGroup = '" & cmbOrgGroup(Index) & "' " & _
            "and D.Site = 'Blood Culture' " & _
            "and D.PriSec = 'S' " & _
            "and D.AntibioticName = A.AntibioticName " & _
            "order by D.ListOrder"
3540  Set tb = New Recordset
3550  RecOpenServer 0, tb, sql
3560  If tb.EOF Then
3570      sql = "Select Distinct A.AntibioticName, D.ListOrder, " & _
                "A.AllowIfPregnant, A.AllowIfOutPatient, A.AllowIfChild " & _
                "from ABDefinitions as D, Antibiotics as A where " & _
                "D.OrganismGroup = '" & cmbOrgGroup(Index) & "' " & _
                "and (D.Site = 'Generic' or D.Site is Null ) and D.PriSec = 'S' " & _
                "and D.AntibioticName = A.AntibioticName " & _
                "order by D.ListOrder"
3580      Set tb = New Recordset
3590      RecOpenServer 0, tb, sql
3600      If tb.EOF Then
3610          Exit Sub
3620      End If
3630  End If
3640  Do While Not tb.EOF

3650      Found = False
3660      ABName = Trim$(tb!AntibioticName & "")
3670      ABCode = AntibioticCodeFor(ABName)
3680      sql = "Select * from Sensitivities where " & _
                "SampleID = '" & SysOptMicroOffset(0) + lblSampleID & "' " & _
                "and IsolateNumber = '" & Index & "' " & _
                "and AntibioticCode = '" & ABCode & "'"
3690      Set tbC = New Recordset
3700      RecOpenServer 0, tbC, sql
3710      If Not tbC.EOF Then
3720          RSI = tbC!RSI & ""
3730          Res = tbC!Result & ""
3740          RunDateTime = Format(tbC!RunDateTime, "dd/mm/yy hh:mm")
3750          Operator = tbC!UserName & ""
3760      Else
3770          RSI = ""
3780          Res = ""
3790          RunDateTime = ""
3800          Operator = ""
3810      End If

3820      For n = 1 To grdAB(Index).Rows - 1
3830          If Trim$(grdAB(Index).TextMatrix(n, 0)) = ABName Then
3840              Found = True
3850              Exit For
3860          End If
3870      Next

3880      If Not Found Then
3890          grdAB(Index).AddItem ABName & vbTab & _
                                   RSI & vbTab & _
                                   vbTab & _
                                   Res & vbTab & _
                                   RunDateTime & vbTab & Operator
3900          grdAB(Index).Row = grdAB(Index).Rows - 1
3910          grdAB(Index).Col = 0
3920          grdAB(Index).CellFontBold = True
3930          grdAB(Index).Col = 2
3940          Set grdAB(Index).CellPicture = imgSquareCross.Picture
3950      End If

3960      tb.MoveNext
3970  Loop

3980  FillABSelect Index

3990  cmdSaveMicro.Enabled = True

4000  Exit Sub

cmdUseSecondary_Click_Error:

      Dim strES As String
      Dim intEL As Integer

4010  intEL = Erl
4020  strES = Err.Description
4030  LogError "frmEditBloodCulture", "cmdUseSecondary_Click", intEL, strES, sql

End Sub

Private Sub cmdViewReports_Click()

4040  frmRFT.SampleID = Val(lblSampleID) + SysOptMicroOffset(0)
4050  frmRFT.Dept = "N"
4060  frmRFT.Show 1

End Sub

Private Sub grdAB_Click(Index As Integer)

      Dim s As String
      Dim RSI As Boolean

4070  On Error GoTo grdAB_Click_Error

4080  cmdSaveMicro.Enabled = True

4090  With grdAB(Index)
4100      If .MouseRow = 0 Then Exit Sub

4110      If .CellBackColor = &HFFFFC0 Then
4120          .Enabled = False
4130          If iMsg("Remove " & Trim$(.Text) & " from List?", vbQuestion + vbYesNo) = vbYes Then
4140              DeleteSensitivity Index, .TextMatrix(.Row, 0)
4150              .RemoveItem .Row
4160              FillABSelect Index
4170          End If
4180          .Enabled = True
4190      ElseIf .Col = 1 Then
4200          s = Trim$(.TextMatrix(.Row, 1))
4210          Select Case s
                  Case "": s = "R": RSI = True
4220              Case "R": s = "S": RSI = True
4230              Case "S": s = "I": RSI = True
4240              Case "I": s = "": RSI = False
4250              Case Else: s = "": RSI = False
4260          End Select
4270          .TextMatrix(.Row, 1) = s
4280          If cmbOrgName(Index) = "Staphylococcus aureus" And UCase(.TextMatrix(.Row, 0)) = "OXACILLIN" And s = "R" Then
4290              cmbOrgName(Index) = "Staphylococcus aureus (MRSA)"
4300          ElseIf cmbOrgName(Index) = "Staphylococcus aureus (MRSA)" And UCase(.TextMatrix(.Row, 0)) = "OXACILLIN" And s <> "R" Then
4310              cmbOrgName(Index) = "Staphylococcus aureus"
4320          End If
4330          .Col = 2
4340          If RSI Then
4350              Set .CellPicture = imgSquareCross.Picture
4360          Else
4370              Set .CellPicture = Nothing
4380          End If
4390      ElseIf .Col = 2 Then
4400          If .CellPicture = imgSquareTick.Picture Then
4410              Set .CellPicture = imgSquareCross.Picture
4420              SetAsForced Index, .TextMatrix(.Row, 0), False
4430          Else
4440              If .TextMatrix(.Row, 2) = "C" Then
4450                  If MsgBox("Report " & .TextMatrix(.Row, 0) & " on a Child?", vbQuestion + vbYesNo) = vbNo Then
4460                      Exit Sub
4470                  End If
4480              ElseIf .TextMatrix(.Row, 2) = "P" Then
4490                  If MsgBox("Report " & .TextMatrix(.Row, 0) & " for Pregnant Patient?", vbQuestion + vbYesNo) = vbNo Then
4500                      Exit Sub
4510                  End If
4520              ElseIf .TextMatrix(.Row, 2) = "O" Then
4530                  If MsgBox("Report " & .TextMatrix(.Row, 0) & " for an Out-Patient?", vbQuestion + vbYesNo) = vbNo Then
4540                      Exit Sub
4550                  End If
4560              End If
4570              Set .CellPicture = imgSquareTick.Picture
4580              SetAsForced Index, .TextMatrix(.Row, 0), True
4590          End If
4600      End If

4610      .LeftCol = 0

4620  End With

4630  Exit Sub

grdAB_Click_Error:

      Dim strES As String
      Dim intEL As Integer

4640  intEL = Erl
4650  strES = Err.Description
4660  LogError "frmEditBloodCulture", "grdAB_Click", intEL, strES

End Sub
Private Sub DeleteSensitivity(ByVal Index As Integer, ByVal ABName As String)

      Dim ABCode As String
      Dim sql As String

4670  On Error GoTo DeleteSensitivity_Error

4680  ABCode = AntibioticCodeFor(ABName)
4690  sql = "DELETE FROM Sensitivities WHERE " & _
            "SampleID = '" & Val(lblSampleID) + SysOptMicroOffset(0) & "' " & _
            "AND IsolateNumber = '" & Index & "' " & _
            "AND AntibioticCode = '" & ABCode & "'"
4700  Cnxn(0).Execute sql

4710  Exit Sub

DeleteSensitivity_Error:

      Dim strES As String
      Dim intEL As Integer

4720  intEL = Erl
4730  strES = Err.Description
4740  LogError "frmEditBloodCulture", "DeleteSensitivity", intEL, strES, sql

End Sub

Private Sub grdAB_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)

4750  pBar = 0

End Sub

Private Sub lblFinal_Click()

4760  With lblFinal
4770      .BackColor = vbGreen
4780      .FontBold = True
4790  End With

4800  With lblInterim
4810      .BackColor = &H8000000F
4820      .FontBold = False
4830  End With

End Sub

Private Sub lblInterim_Click()

4840  With lblInterim
4850      .BackColor = vbGreen
4860      .FontBold = True
4870  End With

4880  With lblFinal
4890      .BackColor = &H8000000F
4900      .FontBold = False
4910  End With


End Sub


Private Sub lblSetAllR_Click(Index As Integer)

      Dim Y As Integer

4920  With grdAB(Index)
4930      .Col = 2
4940      For Y = 1 To .Rows - 1
4950          If .TextMatrix(Y, 0) <> "" Then
4960              .TextMatrix(Y, 1) = "R"
4970              .Row = Y
4980              Set .CellPicture = imgSquareCross.Picture
4990          End If
5000      Next
5010  End With

5020  cmdSaveMicro.Enabled = True

End Sub

Private Sub lblSetAllS_Click(Index As Integer)

      Dim Y As Integer

5030  With grdAB(Index)
5040      .Col = 2
5050      For Y = 1 To .Rows - 1
5060          If .TextMatrix(Y, 0) <> "" Then
5070              .TextMatrix(Y, 1) = "S"
5080              .Row = Y
5090              Set .CellPicture = imgSquareCross.Picture
5100          End If
5110      Next
5120  End With

5130  cmdSaveMicro.Enabled = True

End Sub


Private Sub bHistory_Click()

5140  pBar = 0

5150  With frmMicroReport
5160      .PatChart = lblChart
5170      .PatName = lblName
5180      .PatDoB = lblDoB
5190      .PatSex = Trim$(Left$(lblSex & " ", 1))
5200      .Show 1
5210  End With

End Sub




'---------------------------------------------------------------------------------------
' Procedure : bprint_Click
' Author    : Masood
' Date      : 21/Sep/2016
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub bprint_Click()



10    On Error GoTo bprint_Click_Error

      Dim pSampleID As String

20    pSampleID = SysOptMicroOffset(0) + Val(lblSampleID)
30    If lblFinal.BackColor = vbGreen Then
40        If SampleRelasedtoConsultant(pSampleID, "Micro") Then
50            iMsg "This report is being reviewed by consultant and cannot be released to the ward as a final report.", vbInformation
60            Exit Sub
70        End If
80    End If

90    If frmEditMicrobiologyNew.cmdDemoVal.Caption = "&Validate" Then
100       If iMsg("Demographics are not validated. Do you want to validate now?", vbQuestion + vbYesNo) = vbYes Then
110           If Not EntriesOK(lblSampleID, frmEditMicrobiologyNew.txtName, frmEditMicrobiologyNew.txtSex, frmEditMicrobiologyNew.cmbWard.Text, frmEditMicrobiologyNew.cmbGP.Text) Then
120               Exit Sub
130           Else
140               frmEditMicrobiologyNew.ValidateDemographics True
150           End If
160       Else
170           Exit Sub
180       End If
190   End If

200   SaveMicro

210   frmValidateAll.SampleIDToValidate = SampleIDWithOffset
220   frmValidateAll.Show 1


230   PrintThis


240   GetSampleIDWithOffset
250   LoadAllDetails

260   Exit Sub


bprint_Click_Error:

      Dim strES As String
      Dim intEL As Integer

270   intEL = Erl
280   strES = Err.Description
290   LogError "frmEditBloodCulture", "bprint_Click", intEL, strES

End Sub

Private Sub cmdSaveMicro_Click()

10    On Error GoTo cmdSaveMicro_Click_Error

20    pBar = 0

30    SaveMicro
40    GetSampleIDWithOffset
50    cmdSaveMicro.Enabled = False
60    LoadAllDetails

70    Exit Sub

cmdSaveMicro_Click_Error:
      Dim strES As String
      Dim intEL As Integer

80    intEL = Erl
90    strES = Err.Description
100   LogError "frmEditBloodCulture", "cmdSaveMicro_Click", intEL, strES

End Sub

Private Function SaveMicro()

10    On Error GoTo SaveMicro_Error

20    pBar = 0

30    GetSampleIDWithOffset
40    SaveIsolates
50    SaveSensitivities gNOCHANGE
60    SaveComments


70    Exit Function

SaveMicro_Error:
      Dim strES As String
      Dim intEL As Integer

80    intEL = Erl
90    strES = Err.Description
100   LogError "frmEditBloodCulture", "SaveMicro", intEL, strES

End Function

Private Sub SaveSensitivities(ByVal Validate As Integer)

      Dim sql As String
      Dim intOrg As Integer
      Dim n As Integer
      Dim ABCode As String
      Dim ReportCounter As Integer

10    On Error GoTo SaveSensitivities_Error

20    ReportCounter = 0

30    For intOrg = 1 To 6

40        With grdAB(intOrg)

50            For n = 1 To .Rows - 1
60                If .TextMatrix(n, 0) <> "" Then
70                    ABCode = AntibioticCodeFor(.TextMatrix(n, 0))

                      'Created on 07/10/2010 15:18:57
                      'Autogenerated by SQL Scripting

80                    sql = "If Exists(Select 1 From Sensitivities " & _
                            "Where SampleID = @SampleID And IsolateNumber = @IsolateNumber And AntibioticCode = '@AntibioticCode' ) " & _
                            "Begin " & _
                            "Update Sensitivities Set " & _
                            "SampleID = @SampleID, IsolateNumber = @IsolateNumber, AntibioticCode = '@AntibioticCode', Organism = '@Organism', Antibiotic = '@Antibiotic', Result = '@Result', Report = @Report, CPOFlag = '@CPOFlag', RSI = '@RSI', UserName = '@UserName', Forced = @Forced, Secondary = @Secondary " & _
                            "Where SampleID = @SampleID And IsolateNumber = @IsolateNumber And AntibioticCode = '@AntibioticCode'  " & _
                            "End  " & _
                            "Else " & _
                            "Begin  " & _
                            "Insert Into Sensitivities (SampleID, IsolateNumber, AntibioticCode, Organism, Antibiotic, Result, Report, CPOFlag, RunDate, RunDateTime, RSI, UserName, Forced, Secondary) Values (@SampleID, @IsolateNumber, '@AntibioticCode', '@Organism', '@Antibiotic', '@Result', @Report, '@CPOFlag', '@RunDate', '@RunDateTime', '@RSI', '@UserName', @Forced, @Secondary) " & _
                            "End"

90                    sql = Replace(sql, "@SampleID", SampleIDWithOffset)
100                   sql = Replace(sql, "@IsolateNumber", intOrg)
110                   sql = Replace(sql, "@AntibioticCode", ABCode)
120                   sql = Replace(sql, "@OrgIndex", intOrg)
130                   sql = Replace(sql, "@Organism", cmbOrgName(intOrg))
140                   sql = Replace(sql, "@Antibiotic", .TextMatrix(n, 0))
150                   sql = Replace(sql, "@Result", .TextMatrix(n, 3))
160                   .Row = n
170                   .Col = 2
180                   If .CellPicture = imgSquareTick.Picture Then
190                       sql = Replace(sql, "@Report", 1)
200                   ElseIf .CellPicture = imgSquareCross.Picture Then
210                       sql = Replace(sql, "@Report", 0)
220                   Else
230                       sql = Replace(sql, "@Report", "Null")
240                   End If

250                   sql = Replace(sql, "@CPOFlag", .TextMatrix(n, 2))
260                   sql = Replace(sql, "@RunDateTime", Format(Now, "dd/mmm/yyyy hh:mm"))
270                   sql = Replace(sql, "@RunDate", Format(Now, "dd/mmm/yyyy"))
280                   sql = Replace(sql, "@RSI", .TextMatrix(n, 1))
290                   sql = Replace(sql, "@UserName", UserName)
                      
300                   .Row = n
310                   .Col = 0
320                   sql = Replace(sql, "@Forced", IIf(.CellBackColor = &HFFFFC0, 1, 0))
330                   sql = Replace(sql, "@Secondary", IIf(.CellFontBold = True, 1, 0))

340                   Cnxn(0).Execute sql

350               End If
360           Next
370       End With

380   Next

390   If Validate = gYES Then
400       sql = "UPDATE Sensitivities " & _
                "SET Valid = 1, " & _
                "AuthoriserCode = '" & UserName & "' " & _
                "WHERE SampleID = '" & SampleIDWithOffset & "'"
410       Cnxn(0).Execute sql
420   ElseIf Validate = gNO Then
430       sql = "UPDATE Sensitivities " & _
                "SET Valid = 0, " & _
                "AuthoriserCode = NULL " & _
                "WHERE SampleID = '" & SampleIDWithOffset & "'"
440       Cnxn(0).Execute sql
450   End If

460   Exit Sub

SaveSensitivities_Error:

      Dim strES As String
      Dim intEL As Integer

470   intEL = Erl
480   strES = Err.Description
490   LogError "frmEditBloodCulture", "SaveSensitivities", intEL, strES, sql

End Sub


Private Function LoadSensitivities() As Integer
      'Returns number of isolates

      Dim tb As Recordset
      Dim sql As String
      Dim intIsolate As Integer
      Dim Rows As Integer

5810  On Error GoTo LoadSensitivities_Error

5820  LoadSensitivities = 0

5830  For intIsolate = 1 To 6

5840      sql = "IF NOT EXISTS(SELECT * FROM ABDefinitions " & _
                "              WHERE Site = 'Blood Culture' " & _
                "              AND OrganismGroup = '" & cmbOrgGroup(intIsolate) & "')" & _
                "  INSERT INTO ABDefinitions " & _
                "  SELECT AntibioticName, OrganismGroup, 'Blood Culture' Site, ListOrder, PriSec, AutoReport, AutoReportIf, AutoPriority " & _
                "  FROM ABDefinitions " & _
                "  WHERE Site = 'Generic' " & _
                "  AND OrganismGroup = '" & cmbOrgGroup(intIsolate) & "'"
5850      Cnxn(0).Execute sql

5860      With grdAB(intIsolate)
5870          .Rows = 2
5880          .AddItem ""
5890          .RemoveItem 1

5900          sql = "SELECT B.AntibioticName, S.Report, S.CPOFlag, S.RSI, S.RunDateTime, S.UserName, S.Result, " & _
                    "B.AutoReport, B.AutoReportIf , B.AutoPriority, B.ListOrder " & _
                    "FROM " & _
                    "(SELECT AntibioticName, Listorder, COALESCE(AutoReport, 0) AutoReport, " & _
                    "COALESCE(AutoReportIf,'') AutoReportIf, COALESCE(AutoPriority,0) AutoPriority " & _
                    "FROM ABDefinitions WHERE Site = 'Blood Culture' " & _
                    "AND OrganismGroup = '" & cmbOrgGroup(intIsolate) & "' AND PriSec = 'P') B " & _
                    "LEFT OUTER JOIN " & _
                    "(Select * from Sensitivities WHERE SampleID = '" & SampleIDWithOffset & "' " & _
                    "AND IsolateNumber = '" & intIsolate & "' " & _
                    "AND COALESCE(Forced, 0) = 0 AND COALESCE(Secondary, 0) = 0) S ON B.AntibioticName = S.Antibiotic " & _
                    "ORDER BY B.ListOrder"

              '        "If Exists (Select 1 From ABDefinitions Where Site = 'Blood Culture') " & _
                       '            "Begin " & _
                       '                "SELECT S.Antibiotic, S.Report, S.CPOFlag, S.RSI, " & _
                       '                "S.RunDateTime, S.UserName, S.Result " & _
                       '                "FROM Sensitivities S LEFT OUTER JOIN Antibiotics A ON S.AntibioticCode = A.Code " & _
                       '                "Inner Join (Select AntibioticName, Listorder from ABDefinitions Where Site = 'Blood Culture' And OrganismGroup = '" & cmbOrgGroup(intIsolate) & "') B on S.Antibiotic = B.AntibioticName " & _
                       '                "WHERE SampleID = '" & SampleIDWithOffset & "' " & _
                       '                "AND IsolateNumber = '" & intIsolate & "' and coalesce(forced,0) =0 and coalesce(secondary,0) =0 " & _
                       '                "ORDER BY B.ListOrder " & _
                       '            "End " & _
                       '            "Else " & _
                       '                "SELECT S.Antibiotic, S.Report, S.CPOFlag, S.RSI, " & _
                       '                "S.RunDateTime, S.UserName, S.Result " & _
                       '                "FROM Sensitivities S LEFT OUTER JOIN Antibiotics A ON S.AntibioticCode = A.Code " & _
                       '                "Inner Join (Select AntibioticName, Listorder from ABDefinitions Where Site = 'Generic' And OrganismGroup = '" & cmbOrgGroup(intIsolate) & "') B on S.Antibiotic = B.AntibioticName " & _
                       '                "WHERE SampleID = '" & SampleIDWithOffset & "' " & _
                       '                "AND IsolateNumber = '" & intIsolate & "' and coalesce(forced,0) =0 and coalesce(secondary,0) =0 " & _
                       '                "ORDER BY B.ListOrder "

5910          Set tb = New Recordset
5920          RecOpenServer 0, tb, sql
5930          If Not tb.EOF Then
5940              LoadSensitivities = intIsolate
5950          End If
5960          Do While Not tb.EOF
5970              .AddItem tb!AntibioticName & vbTab & _
                           tb!RSI & vbTab & _
                           tb!CPOFlag & vbTab & _
                           tb!Result & vbTab & _
                           Format(tb!RunDateTime, "dd/mm/yy hh:mm") & _
                           tb!UserName & ""
5980              .Row = .Rows - 1
5990              .Col = 2
6000              If IsNull(tb!Report) Then
6010                  Set .CellPicture = Me.Picture
6020              Else
6030                  Set .CellPicture = IIf(tb!Report, imgSquareTick.Picture, imgSquareCross.Picture)
6040              End If
6050              tb.MoveNext
6060          Loop
6070          If .Rows > 2 Then
6080              .RemoveItem 1
6090          End If
6100      End With

6110      LoadSensitivitiesForced intIsolate
6120      LoadSensitivitiesSecondary intIsolate

6130      FillABSelect intIsolate
6140  Next

6150  Exit Function

LoadSensitivities_Error:

      Dim strES As String
      Dim intEL As Integer

6160  intEL = Erl
6170  strES = Err.Description
6180  LogError "frmEditBloodCulture", "LoadSensitivities", intEL, strES, sql

End Function

Private Sub cmdValidateMicro_Click()

      Dim tb As Recordset
      Dim sql As String
      Dim Validate As Boolean

6190  On Error GoTo cmdValidateMicro_Click_Error

6200  GetSampleIDWithOffset

6210  Validate = cmdValidateMicro.Caption = "&Validate"

6220  If Validate Then
6230      cmdValidateMicro.Caption = "Un&Validate"
6240      UpdatePrintValidLog SampleIDWithOffset, "BLOODCULTURE", 1, 2
6250  Else
6260      sql = "SELECT Password FROM Users WHERE " & _
                "Name = '" & AddTicks(UserName) & "'"
6270      Set tb = New Recordset
6280      RecOpenServer 0, tb, sql
6290      If Not tb.EOF Then
6300          If UCase$(iBOX("Password Required", , , True)) = UCase$(tb!PassWord & "") Then
6310              cmdValidateMicro.Caption = "&Validate"
6320              UpdatePrintValidLog SampleIDWithOffset, "BLOODCULTURE", 0, 2
6330          Else
6340              Exit Sub
6350          End If
6360      Else
6370          Exit Sub
6380      End If
6390  End If

6400  SaveIsolates
6410  SaveSensitivities gYES

6420  SaveComments
6430  cmdSaveMicro.Enabled = False

6440  GetSampleIDWithOffset
6450  LoadAllDetails

6460  cmdSaveMicro.Enabled = False

6470  Exit Sub

cmdValidateMicro_Click_Error:

      Dim strES As String
      Dim intEL As Integer

6480  intEL = Erl
6490  strES = Err.Description
6500  LogError "frmEditBloodCulture", "cmdValidateMicro_Click", intEL, strES, sql

End Sub




Private Sub LoadAllDetails()

6510  ForceSaveability = False

6520  AdjustOrganism
6530  LoadBloodCulture

6540  If LoadIsolates() Then
6550      SetComboWidths
6560  End If

6570  LoadSensitivities

6580  LoadComments

6590  fraBC(0).Enabled = True
6600  fraBC(1).Enabled = True
6610  If CheckValidStatus() Then
6620      fraBC(0).Enabled = False
6630      fraBC(1).Enabled = False
6640  End If

End Sub
Private Sub LoadComments()

      Dim Ob As Observation
      Dim Obs As Observations

6650  On Error GoTo LoadComments_Error

6660  txtMSC = "Medical Scientist Comments"
6670  txtConC = "Consultant Comments"

6680  If Trim$(lblSampleID) = "" Then Exit Sub

6690  Set Obs = New Observations
6700  Set Obs = Obs.Load(Val(lblSampleID) + SysOptMicroOffset(0), "MicroCS", "MicroConsultant")
6710  If Not Obs Is Nothing Then
6720      For Each Ob In Obs
6730          Select Case UCase$(Ob.Discipline)
                  Case "MICROCS": txtMSC = Split_Comm(Ob.Comment)
6740              Case "MICROCONSULTANT": txtConC = Split_Comm(Ob.Comment)
6750          End Select
6760      Next
6770  End If

6780  Exit Sub

LoadComments_Error:

      Dim strES As String
      Dim intEL As Integer

6790  intEL = Erl
6800  strES = Err.Description
6810  LogError "frmEditBloodCulture", "LoadComments", intEL, strES

End Sub


Private Sub LoadBloodCulture()

      Dim tb As Recordset
      Dim sql As String
      Dim s As String
      Dim Index As Integer

6820  On Error GoTo LoadBloodCulture_Error

6830  For Index = 1 To 6
6840      grdAB(Index).Rows = 2
6850      grdAB(Index).AddItem ""
6860      grdAB(Index).RemoveItem 1
6870  Next

6880  lblBC(1) = ""
6890  lblBC(3) = ""
6900  lblBC(5) = ""

6910  If Val(lblSampleID) = 0 Then Exit Sub

6920  sql = "SELECT * FROM BloodCultureResults WHERE " & _
            "SampleID = '" & lblSampleID + SysOptMicroOffset(0) & "' " & _
            "ORDER BY RunDateTime DESC"
6930  Set tb = New Recordset
6940  RecOpenServer 0, tb, sql
6950  Do While Not tb.EOF

6960      Select Case tb!TypeOfTest & ""
              Case GetOptionSetting("BcAerobicBottle", "BSA"): Index = 1
6970          Case GetOptionSetting("BcAnarobicBottle", "BSN"): Index = 3
6980          Case GetOptionSetting("BcFanBottle", "BFA"): Index = 5
6990          Case Else: Index = 5
7000      End Select

7010      Select Case tb!Result & ""
              Case "+": s = "Positive"
7020          Case "-": s = "Negative": AdjustNegative Index
7030          Case "*": s = "Neg to date"
7040          Case Else: s = "Unknown"
7050      End Select

7060      lblBC(Index) = tb!TypeOfTest & " " & s

7070      tb.MoveNext

7080  Loop


7090  Exit Sub

LoadBloodCulture_Error:

      Dim strES As String
      Dim intEL As Integer

7100  intEL = Erl
7110  strES = Err.Description
7120  LogError "frmEditBloodCulture", "LoadBloodCulture", intEL, strES, sql

End Sub

Private Sub cmdCancel_Click()

7130  pBar = 0

7140  Unload Me

End Sub


Private Sub cmdSetPrinter_Click()

7150  On Error GoTo cmdSetPrinter_Click_Error

7160  Set frmForcePrinter.f = frmEditBloodCulture
7170  frmForcePrinter.Show 1

7180  If pPrintToPrinter = "Automatic Selection" Then
7190      pPrintToPrinter = ""
7200  End If

7210  If pPrintToPrinter <> "" Then
7220      cmdSetPrinter.BackColor = vbRed
7230      cmdSetPrinter.ToolTipText = "Print Forced to " & pPrintToPrinter
7240  Else
7250      cmdSetPrinter.BackColor = vbButtonFace
7260      pPrintToPrinter = ""
7270      cmdSetPrinter.ToolTipText = "Printer Selected Automatically"
7280  End If

7290  Exit Sub

cmdSetPrinter_Click_Error:

      Dim strES As String
      Dim intEL As Integer

7300  intEL = Erl
7310  strES = Err.Description
7320  LogError "frmEditBloodCulture", "cmdSetPrinter_Click", intEL, strES

End Sub

Private Sub Form_Activate()

7330  TimerBar.Enabled = True
7340  pBar = 0

7350  ObservaInUse = IIf(GetOptionSetting("ObservaInUse", "0") = "0", False, True)

7360  BacTek3DInUse = IIf(GetOptionSetting("Bactek3DInUse", "0") = "0", False, True)

7370  FillOrganismGroups
7380  FillMSandConsultantComment

7390  GetSampleIDWithOffset
7400  LoadAllDetails

7410  cmdSaveMicro.Enabled = False
7420  cmdValidateMicro.Enabled = True

7430  With lblFinal
7440      .BackColor = vbGreen
7450      .FontBold = True
7460  End With

7470  With lblInterim
7480      .BackColor = &H8000000F
7490      .FontBold = False
7500  End With

End Sub

Private Sub FillOrganismGroups()

      Dim n As Integer
      Dim tb As Recordset
      Dim sql As String
      Dim temp As String

7510  On Error GoTo FillOrganismGroups_Error

7520  sql = "Select * from Lists where " & _
            "ListType = 'OR' " & _
            "order by ListOrder"
7530  Set tb = New Recordset
7540  RecOpenServer 0, tb, sql

7550  For n = 1 To 6
7560      cmbOrgGroup(n).Clear
7570      cmbOrgName(n).Clear
7580  Next

7590  Do While Not tb.EOF
7600      temp = tb!Text & ""
7610      For n = 1 To 6
7620          cmbOrgGroup(n).AddItem temp
7630      Next
7640      tb.MoveNext
7650  Loop

7660  SetComboWidths

7670  Exit Sub

FillOrganismGroups_Error:

      Dim strES As String
      Dim intEL As Integer

7680  intEL = Erl
7690  strES = Err.Description
7700  LogError "frmEditBloodCulture", "FillOrganismGroups", intEL, strES, sql

End Sub

Private Sub FillAbGrid(ByVal Index As Integer)

      Dim tb As Recordset
      Dim sql As String
      Dim ReportCounter As Integer

7710  On Error GoTo FillAbGrid_Error

7720  With grdAB(Index)
7730      .Visible = False
7740      .Rows = 2
7750      .AddItem ""
7760      .RemoveItem 1
7770  End With

7780  ReportCounter = 0

7790  sql = "Select Distinct A.AntibioticName, D.ListOrder, " & _
            "A.AllowIfPregnant, A.AllowIfOutPatient, A.AllowIfChild " & _
            "from ABDefinitions as D, Antibiotics as A where " & _
            "D.OrganismGroup = '" & cmbOrgGroup(Index) & "' " & _
            "and D.Site = 'Blood Culture' " & _
            "and D.PriSec = 'P' " & _
            "and D.AntibioticName = A.AntibioticName " & _
            "order by D.ListOrder"
7800  Set tb = New Recordset
7810  RecOpenClient 0, tb, sql
7820  If tb.EOF Then
7830      sql = "Select Distinct A.AntibioticName, D.ListOrder, " & _
                "A.AllowIfPregnant, A.AllowIfOutPatient, A.AllowIfChild " & _
                "from ABDefinitions as D, Antibiotics as A where " & _
                "D.OrganismGroup = '" & cmbOrgGroup(Index) & "' " & _
                "and Site = 'Generic' " & _
                "and D.PriSec = 'P' " & _
                "and D.AntibioticName = A.AntibioticName " & _
                "order by D.ListOrder"
7840      Set tb = New Recordset
7850      RecOpenClient 0, tb, sql
7860      If tb.EOF Then
              ' iMsg "Site/Organism not defined.", vbCritical
7870          grdAB(Index).Visible = True
7880          Exit Sub
7890      End If
7900  End If

7910  Do While Not tb.EOF
7920      grdAB(Index).AddItem Trim$(tb!AntibioticName)
7930      grdAB(Index).Row = grdAB(Index).Rows - 1
7940      grdAB(Index).Col = 2
7950      Set grdAB(Index).CellPicture = Me.Picture
7960      tb.MoveNext
7970  Loop

7980  If grdAB(Index).Rows > 2 Then
7990      grdAB(Index).RemoveItem 1
8000  End If
8010  grdAB(Index).Visible = True

8020  Exit Sub

FillAbGrid_Error:

      Dim strES As String
      Dim intEL As Integer

8030  intEL = Erl
8040  strES = Err.Description
8050  LogError "frmEditBloodCulture", "FillAbGrid", intEL, strES, sql

End Sub

Private Sub Form_Deactivate()

8060  pBar = 0
8070  TimerBar.Enabled = False

End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)

8080  pBar = 0

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

8090  pBar = 0

End Sub

Private Sub Form_Unload(Cancel As Integer)

8100  pPrintToPrinter = ""

End Sub

Private Sub fraSampleID_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

8110  pBar = 0

End Sub


Private Sub mnuConsultantComment_Click()

8120  With frmListsGeneric
8130      .ListType = "ConsComment"
8140      .ListTypeName = "Consultants Comment"
8150      .ListTypeNames = "Consultants Comments"
8160      .Show 1
8170  End With

8180  LoadListConsultantComment

End Sub

Private Sub mnuMSC_Click()

8190  With frmListsGeneric
8200      .ListType = "MSComment"
8210      .ListTypeName = "Medical Scientist Comment"
8220      .ListTypeNames = "Medical Scientist Comments"
8230      .Show 1
8240  End With

8250  LoadListMSComment

End Sub

Private Sub txtConC_GotFocus()

8260  If txtConC.Text = "Consultant Comments" Then
8270      txtConC.Text = ""
8280  End If

End Sub


Private Sub txtConC_KeyUp(KeyCode As Integer, Shift As Integer)

8290  cmdSaveMicro.Enabled = True

End Sub


Private Sub txtConC_LostFocus()

8300  If Trim$(txtConC) = "" Then
8310      txtConC = "Consultant Comments"
8320  End If

End Sub


Private Sub TimerBar_Timer()

8330  pBar = pBar + 1

8340  If pBar = pBar.Max Then
8350      Unload Me
8360      Exit Sub
8370  End If

End Sub


Private Sub txtMSC_GotFocus()

8380  If txtMSC.Text = "Medical Scientist Comments" Then
8390      txtMSC.Text = ""
8400  End If

End Sub


Private Sub txtMSC_KeyUp(KeyCode As Integer, Shift As Integer)

8410  cmdSaveMicro.Enabled = True

End Sub

Private Sub txtMSC_LostFocus()

8420  If Trim$(txtMSC) = "" Then
8430      txtMSC = "Medical Scientist Comments"
8440  End If

End Sub
Public Property Let PrintToPrinter(ByVal strNewValue As String)

8450  pPrintToPrinter = strNewValue

End Property
Public Property Get PrintToPrinter() As String

8460  PrintToPrinter = pPrintToPrinter

End Property

Private Function CheckValidStatus() As Boolean
      'Pass the tab number: Returns true if Valid

      Dim tb As Recordset
      Dim sql As String
      Dim Dept As String

8470  On Error GoTo CheckValidStatus_Error

8480  CheckValidStatus = False
8490  cmdValidateMicro.Caption = "&Validate"

8500  Dept = "B"

8510  sql = "SELECT Valid, Printed FROM PrintValidLog WHERE " & _
            "SampleID = '" & Val(lblSampleID) + SysOptMicroOffset(0) & "' " & _
            "AND Department = '" & Dept & "'"
8520  Set tb = New Recordset
8530  RecOpenServer 0, tb, sql
8540  If Not tb.EOF Then
8550      If tb!Valid = 1 Then
8560          CheckValidStatus = True
8570          cmdValidateMicro.Caption = "Un&Validate"
8580      End If

8590  End If

8600  Exit Function

CheckValidStatus_Error:

      Dim strES As String
      Dim intEL As Integer

8610  intEL = Erl
8620  strES = Err.Description
8630  LogError "frmEditBloodCulture", "CheckValidStatus", intEL, strES, sql

End Function

