VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEditSemen 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Semen Analysis"
   ClientHeight    =   8925
   ClientLeft      =   30
   ClientTop       =   375
   ClientWidth     =   14835
   Icon            =   "frmEditSemen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8925
   ScaleWidth      =   14835
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraDate 
      Caption         =   "Sample Date"
      Height          =   1545
      Left            =   5940
      TabIndex        =   94
      Top             =   1860
      Width           =   7275
      Begin VB.Frame Frame5 
         Height          =   495
         Left            =   4860
         TabIndex        =   103
         Top             =   1050
         Width           =   2355
         Begin VB.OptionButton cRooH 
            Caption         =   "Out of Hours"
            Height          =   195
            Index           =   1
            Left            =   1020
            TabIndex        =   105
            Top             =   180
            Width           =   1215
         End
         Begin VB.OptionButton cRooH 
            Alignment       =   1  'Right Justify
            Caption         =   "Routine"
            Height          =   195
            Index           =   0
            Left            =   90
            TabIndex        =   104
            Top             =   180
            Width           =   885
         End
      End
      Begin MSComCtl2.DTPicker dtRunDate 
         Height          =   315
         Left            =   5520
         TabIndex        =   95
         Top             =   390
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         Format          =   273022977
         CurrentDate     =   36942
      End
      Begin MSComCtl2.DTPicker dtSampleDate 
         Height          =   315
         Left            =   180
         TabIndex        =   96
         Top             =   405
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         Format          =   273022977
         CurrentDate     =   36942
      End
      Begin MSMask.MaskEdBox tSampleTime 
         Height          =   315
         Left            =   1560
         TabIndex        =   97
         ToolTipText     =   "Time of Sample"
         Top             =   405
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
         Left            =   2820
         TabIndex        =   98
         Top             =   405
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         Format          =   273022977
         CurrentDate     =   36942
      End
      Begin MSMask.MaskEdBox tRecTime 
         Height          =   315
         Left            =   4200
         TabIndex        =   99
         ToolTipText     =   "Time of Sample"
         Top             =   405
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
      Begin VB.Image iRunDate 
         Height          =   330
         Index           =   0
         Left            =   5520
         Picture         =   "frmEditSemen.frx":030A
         Stretch         =   -1  'True
         ToolTipText     =   "Previous Day"
         Top             =   720
         Width           =   480
      End
      Begin VB.Image iRunDate 
         Height          =   330
         Index           =   1
         Left            =   6420
         Picture         =   "frmEditSemen.frx":074C
         Stretch         =   -1  'True
         ToolTipText     =   "Next Day"
         Top             =   720
         Width           =   480
      End
      Begin VB.Image iSampleDate 
         Height          =   330
         Index           =   0
         Left            =   210
         Picture         =   "frmEditSemen.frx":0B8E
         Stretch         =   -1  'True
         ToolTipText     =   "Previous Day"
         Top             =   720
         Width           =   480
      End
      Begin VB.Image iSampleDate 
         Height          =   330
         Index           =   1
         Left            =   1050
         Picture         =   "frmEditSemen.frx":0FD0
         Stretch         =   -1  'True
         ToolTipText     =   "Next Day"
         Top             =   720
         Width           =   480
      End
      Begin VB.Image iToday 
         Height          =   330
         Index           =   0
         Left            =   6030
         Picture         =   "frmEditSemen.frx":1412
         Stretch         =   -1  'True
         ToolTipText     =   "Set to Today"
         Top             =   720
         Width           =   360
      End
      Begin VB.Image iToday 
         Height          =   330
         Index           =   1
         Left            =   690
         Picture         =   "frmEditSemen.frx":1854
         Stretch         =   -1  'True
         ToolTipText     =   "Set to Today"
         Top             =   720
         Width           =   360
      End
      Begin VB.Image iToday 
         Height          =   330
         Index           =   2
         Left            =   3300
         Picture         =   "frmEditSemen.frx":1C96
         Stretch         =   -1  'True
         ToolTipText     =   "Set to Today"
         Top             =   720
         Width           =   360
      End
      Begin VB.Image iRecDate 
         Height          =   330
         Index           =   1
         Left            =   3690
         Picture         =   "frmEditSemen.frx":20D8
         Stretch         =   -1  'True
         ToolTipText     =   "Next Day"
         Top             =   720
         Width           =   480
      End
      Begin VB.Image iRecDate 
         Height          =   330
         Index           =   0
         Left            =   2820
         Picture         =   "frmEditSemen.frx":251A
         Stretch         =   -1  'True
         ToolTipText     =   "Previous Day"
         Top             =   720
         Width           =   480
      End
      Begin VB.Label lblDate 
         Alignment       =   2  'Center
         Caption         =   "Received in Lab"
         Height          =   255
         Index           =   0
         Left            =   2670
         TabIndex        =   102
         Top             =   0
         Width           =   1500
      End
      Begin VB.Label lblDate 
         Alignment       =   2  'Center
         Caption         =   "Run Date"
         Height          =   225
         Index           =   1
         Left            =   5580
         TabIndex        =   101
         Top             =   0
         Width           =   930
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
         Height          =   315
         Left            =   2280
         TabIndex        =   100
         Top             =   1170
         Visible         =   0   'False
         Width           =   2385
      End
   End
   Begin VB.CommandButton cmdArchive 
      BackColor       =   &H0000FFFF&
      Caption         =   "Archived Entries"
      Height          =   675
      Left            =   13350
      Picture         =   "frmEditSemen.frx":295C
      Style           =   1  'Graphical
      TabIndex        =   93
      Top             =   1050
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.CommandButton cmdViewMicroRep 
      Caption         =   "View Reports"
      Height          =   780
      Left            =   360
      Picture         =   "frmEditSemen.frx":2C66
      Style           =   1  'Graphical
      TabIndex        =   84
      Top             =   7560
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.CommandButton cmdValidate 
      Caption         =   "&Validate"
      Height          =   645
      Left            =   13320
      Picture         =   "frmEditSemen.frx":2F70
      Style           =   1  'Graphical
      TabIndex        =   67
      Top             =   6300
      Width           =   1290
   End
   Begin VB.Frame fraSpecimen 
      Caption         =   "Specimen"
      Height          =   4755
      Left            =   5940
      TabIndex        =   58
      Top             =   3555
      Width           =   7275
      Begin VB.Frame Frame4 
         Caption         =   "Type"
         Height          =   855
         Left            =   120
         TabIndex        =   110
         Top             =   420
         Width           =   2175
         Begin VB.ComboBox cmbSpecimenType 
            Height          =   315
            Left            =   120
            TabIndex        =   111
            Text            =   "cmbSpecimenType"
            Top             =   360
            Width           =   1905
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Morphology"
         Height          =   1425
         Left            =   4800
         TabIndex        =   90
         Top             =   1560
         Width           =   2265
         Begin VB.TextBox txtMorphology 
            Height          =   285
            Left            =   150
            TabIndex        =   91
            Text            =   "15"
            Top             =   600
            Width           =   555
         End
         Begin VB.Label lblSemenMorph 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "% Abnormal Forms"
            Height          =   285
            Left            =   750
            TabIndex        =   92
            Top             =   600
            Width           =   1425
         End
      End
      Begin VB.ComboBox cmbSemenComments 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   80
         Top             =   3240
         Width           =   6645
      End
      Begin VB.Frame Frame8 
         Caption         =   "Comments"
         Height          =   1605
         Left            =   120
         TabIndex        =   78
         Top             =   3000
         Width           =   6975
         Begin VB.TextBox txtSemenComment 
            BackColor       =   &H80000018&
            Height          =   945
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   79
            Top             =   570
            Width           =   6645
         End
      End
      Begin VB.Frame fraMotility 
         Caption         =   "Motility"
         Height          =   1425
         Left            =   120
         TabIndex        =   68
         Top             =   1560
         Width           =   4635
         Begin VB.TextBox txtMotility 
            Height          =   285
            Index           =   4
            Left            =   2220
            TabIndex        =   107
            Top             =   600
            Width           =   600
         End
         Begin VB.TextBox txtMotility 
            Height          =   285
            Index           =   3
            Left            =   750
            TabIndex        =   87
            Top             =   210
            Width           =   690
         End
         Begin VB.TextBox txtMotility 
            Height          =   285
            Index           =   0
            Left            =   2220
            TabIndex        =   71
            Top             =   210
            Width           =   600
         End
         Begin VB.TextBox txtMotility 
            Height          =   285
            Index           =   1
            Left            =   2220
            TabIndex        =   70
            Top             =   1005
            Width           =   600
         End
         Begin VB.TextBox txtMotility 
            Height          =   285
            Index           =   2
            Left            =   750
            TabIndex        =   69
            Top             =   1005
            Width           =   690
         End
         Begin ComCtl2.UpDown udMotility 
            Height          =   375
            Index           =   2
            Left            =   1440
            TabIndex        =   72
            Top             =   960
            Width           =   240
            _ExtentX        =   450
            _ExtentY        =   661
            _Version        =   327681
            BuddyControl    =   "txtMotility(2)"
            BuddyDispid     =   196636
            BuddyIndex      =   2
            OrigLeft        =   810
            OrigTop         =   1140
            OrigRight       =   1050
            OrigBottom      =   1485
            Increment       =   5
            Max             =   100
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin ComCtl2.UpDown udMotility 
            Height          =   375
            Index           =   1
            Left            =   2820
            TabIndex        =   73
            Top             =   960
            Width           =   240
            _ExtentX        =   450
            _ExtentY        =   661
            _Version        =   327681
            BuddyControl    =   "txtMotility(1)"
            BuddyDispid     =   196636
            BuddyIndex      =   1
            OrigLeft        =   810
            OrigTop         =   780
            OrigRight       =   1050
            OrigBottom      =   1065
            Increment       =   5
            Max             =   100
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin ComCtl2.UpDown udMotility 
            Height          =   375
            Index           =   0
            Left            =   2820
            TabIndex        =   74
            Top             =   165
            Width           =   240
            _ExtentX        =   450
            _ExtentY        =   661
            _Version        =   327681
            BuddyControl    =   "txtMotility(0)"
            BuddyDispid     =   196636
            BuddyIndex      =   0
            OrigLeft        =   810
            OrigTop         =   270
            OrigRight       =   1050
            OrigBottom      =   555
            Increment       =   5
            Max             =   100
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin ComCtl2.UpDown udMotility 
            Height          =   375
            Index           =   3
            Left            =   1440
            TabIndex        =   88
            Top             =   165
            Width           =   240
            _ExtentX        =   450
            _ExtentY        =   661
            _Version        =   327681
            BuddyControl    =   "txtMotility(3)"
            BuddyDispid     =   196636
            BuddyIndex      =   3
            OrigLeft        =   810
            OrigTop         =   1140
            OrigRight       =   1050
            OrigBottom      =   1485
            Increment       =   5
            Max             =   100
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin ComCtl2.UpDown udMotility 
            Height          =   375
            Index           =   4
            Left            =   2820
            TabIndex        =   108
            Top             =   570
            Width           =   240
            _ExtentX        =   450
            _ExtentY        =   661
            _Version        =   327681
            BuddyControl    =   "txtMotility(4)"
            BuddyDispid     =   196636
            BuddyIndex      =   4
            OrigLeft        =   810
            OrigTop         =   780
            OrigRight       =   1050
            OrigBottom      =   1065
            Increment       =   5
            Max             =   100
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "% Slow Progressive"
            Height          =   195
            Index           =   3
            Left            =   3060
            TabIndex        =   109
            Top             =   645
            Width           =   1380
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "% Motile"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   89
            Top             =   255
            Width           =   585
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "% Progressive"
            Height          =   195
            Index           =   0
            Left            =   3060
            TabIndex        =   77
            Top             =   255
            Width           =   990
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "% Non-Progressive"
            Height          =   195
            Index           =   1
            Left            =   3060
            TabIndex        =   76
            Top             =   1050
            Width           =   1335
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Caption         =   "% Non Motile"
            Height          =   435
            Index           =   4
            Left            =   150
            TabIndex        =   75
            Top             =   930
            Width           =   600
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Volume"
         Height          =   645
         Left            =   4740
         TabIndex        =   64
         Top             =   150
         Width           =   2355
         Begin VB.ComboBox cmbVolume 
            Height          =   315
            Left            =   150
            TabIndex        =   65
            Text            =   "cmbVolume"
            Top             =   240
            Width           =   1365
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            Caption         =   "mL"
            Height          =   195
            Index           =   1
            Left            =   1650
            TabIndex        =   66
            Top             =   300
            Width           =   210
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Count"
         Height          =   675
         Index           =   0
         Left            =   2460
         TabIndex        =   61
         Top             =   810
         Width           =   4635
         Begin VB.ComboBox cmbCount 
            Height          =   315
            Left            =   210
            TabIndex        =   62
            Text            =   "cmbCount"
            Top             =   240
            Width           =   2775
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            Caption         =   "Million per mL"
            Height          =   195
            Index           =   0
            Left            =   3030
            TabIndex        =   63
            Top             =   300
            Width           =   960
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Consistency"
         Height          =   645
         Left            =   2460
         TabIndex        =   59
         Top             =   150
         Width           =   2085
         Begin VB.ComboBox cmbConsistency 
            Height          =   315
            Left            =   210
            TabIndex        =   60
            Text            =   "cmbConsistency"
            Top             =   240
            Width           =   1545
         End
      End
   End
   Begin VB.CommandButton cmdSaveHold 
      Caption         =   "Save && &Hold"
      Enabled         =   0   'False
      Height          =   735
      Left            =   13350
      Picture         =   "frmEditSemen.frx":327A
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   5520
      Width           =   1275
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   735
      Left            =   13350
      Picture         =   "frmEditSemen.frx":3584
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4740
      Width           =   1275
   End
   Begin VB.CommandButton cmdOrderSemen 
      Caption         =   "Order Analysis"
      Height          =   735
      Left            =   4545
      Picture         =   "frmEditSemen.frx":388E
      Style           =   1  'Graphical
      TabIndex        =   56
      Tag             =   "bOrder"
      Top             =   7560
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.Frame fraDemographics 
      Height          =   5655
      Left            =   390
      TabIndex        =   37
      Top             =   1860
      Width           =   5445
      Begin VB.CommandButton cmdCopyTo 
         BackColor       =   &H008080FF&
         Caption         =   "++ cc ++"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   4980
         Style           =   1  'Graphical
         TabIndex        =   106
         ToolTipText     =   "Copy To"
         Top             =   2520
         Width           =   375
      End
      Begin VB.ComboBox cmbHospital 
         Height          =   315
         Left            =   1050
         TabIndex        =   7
         Text            =   "cmbHospital"
         Top             =   2520
         Width           =   3915
      End
      Begin VB.ComboBox cmbDemogComments 
         Height          =   315
         Left            =   1050
         Style           =   2  'Dropdown List
         TabIndex        =   81
         Top             =   3840
         Width           =   3915
      End
      Begin VB.TextBox txtDemographicComment 
         Height          =   915
         Left            =   1050
         MaxLength       =   160
         MultiLine       =   -1  'True
         TabIndex        =   39
         Top             =   4140
         Width           =   3885
      End
      Begin VB.ComboBox cmbClinDetails 
         Height          =   315
         Left            =   1050
         Sorted          =   -1  'True
         TabIndex        =   38
         Top             =   5070
         Width           =   3915
      End
      Begin VB.ComboBox cmbWard 
         Height          =   315
         Left            =   1050
         TabIndex        =   8
         Text            =   "cmbWard"
         Top             =   2850
         Width           =   3915
      End
      Begin VB.TextBox tAddress 
         Height          =   285
         Index           =   0
         Left            =   750
         MaxLength       =   30
         TabIndex        =   5
         Top             =   1830
         Width           =   4215
      End
      Begin VB.TextBox tAddress 
         Height          =   285
         Index           =   1
         Left            =   750
         MaxLength       =   30
         TabIndex        =   6
         Top             =   2100
         Width           =   4215
      End
      Begin VB.ComboBox cmbClinician 
         Height          =   315
         Left            =   1050
         TabIndex        =   9
         Text            =   "cmbClinician"
         Top             =   3180
         Width           =   3915
      End
      Begin VB.ComboBox cmbGP 
         Height          =   315
         Left            =   1050
         TabIndex        =   10
         Text            =   "cmbGP"
         Top             =   3510
         Width           =   3915
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         Caption         =   "Hospital"
         Height          =   195
         Index           =   1
         Left            =   420
         TabIndex        =   82
         Top             =   2580
         Width           =   570
      End
      Begin VB.Label lblSex 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4260
         TabIndex        =   55
         Top             =   1200
         Width           =   705
      End
      Begin VB.Label lblAge 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3180
         TabIndex        =   54
         Top             =   1200
         Width           =   585
      End
      Begin VB.Label lblDoB 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   750
         TabIndex        =   53
         Top             =   1170
         Width           =   1515
      End
      Begin VB.Label lblName 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   750
         TabIndex        =   52
         Top             =   780
         Width           =   4215
      End
      Begin VB.Label lblChart 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   750
         TabIndex        =   51
         Top             =   300
         Width           =   1485
      End
      Begin VB.Label Label36 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Cl Details"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   330
         TabIndex        =   50
         Top             =   5130
         Width           =   660
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Chart #"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   49
         Top             =   330
         Width           =   525
      End
      Begin VB.Label Label17 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Name"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   210
         TabIndex        =   48
         Top             =   810
         Width           =   420
      End
      Begin VB.Label Label21 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "D.o.B"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   210
         TabIndex        =   47
         Top             =   1230
         Width           =   405
      End
      Begin VB.Label Label22 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Age"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   2760
         TabIndex        =   46
         Top             =   1200
         Width           =   285
      End
      Begin VB.Label Label23 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Sex"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   3930
         TabIndex        =   45
         Top             =   1200
         Width           =   270
      End
      Begin VB.Label Label24 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Ward"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   600
         TabIndex        =   44
         Top             =   2910
         Width           =   390
      End
      Begin VB.Label Label25 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Address"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   780
         TabIndex        =   43
         Top             =   1620
         Width           =   570
      End
      Begin VB.Label Label30 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Comments"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   255
         TabIndex        =   42
         Top             =   3870
         Width           =   735
      End
      Begin VB.Label Label31 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Clinician"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   405
         TabIndex        =   41
         Top             =   3240
         Width           =   585
      End
      Begin VB.Label Label32 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "GP"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   765
         TabIndex        =   40
         Top             =   3540
         Width           =   225
      End
   End
   Begin VB.CommandButton cmdSemenHistory 
      Caption         =   "&History"
      Height          =   660
      Left            =   13350
      Picture         =   "frmEditSemen.frx":3B98
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   6990
      Width           =   1275
   End
   Begin VB.CommandButton cmdSetPrinter 
      Height          =   615
      Left            =   13350
      Picture         =   "frmEditSemen.frx":3FDA
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Back Colour Normal:Automatic Printer Selection.Back Colour Red:-Forced"
      Top             =   1770
      Width           =   1275
   End
   Begin VB.Timer TimerBar 
      Interval        =   1000
      Left            =   11520
      Top             =   0
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   135
      Left            =   390
      TabIndex        =   29
      Top             =   120
      Width           =   12825
      _ExtentX        =   22622
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   1
      Max             =   600
      Scrolling       =   1
   End
   Begin VB.CommandButton bPrintHold 
      Caption         =   "Print && Hold"
      Height          =   720
      Left            =   13350
      Picture         =   "frmEditSemen.frx":42E4
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   2415
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1395
      Left            =   390
      TabIndex        =   18
      Top             =   330
      Width           =   12825
      Begin VB.CommandButton cmdDartViewer 
         Height          =   390
         Left            =   7290
         Picture         =   "frmEditSemen.frx":45EE
         Style           =   1  'Graphical
         TabIndex        =   86
         Top             =   120
         Width           =   375
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
         Left            =   4050
         TabIndex        =   1
         Tag             =   "A and E Number"
         ToolTipText     =   "A & E Number"
         Top             =   540
         Width           =   1635
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
         Left            =   2490
         MaxLength       =   8
         TabIndex        =   12
         Top             =   540
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
         Left            =   5700
         MaxLength       =   30
         TabIndex        =   2
         Tag             =   "tName"
         Top             =   540
         Width           =   4455
      End
      Begin VB.TextBox txtDoB 
         Height          =   285
         Left            =   10890
         MaxLength       =   10
         TabIndex        =   3
         Top             =   270
         Width           =   1125
      End
      Begin VB.TextBox txtAge 
         Height          =   285
         Left            =   10890
         MaxLength       =   4
         TabIndex        =   13
         Top             =   630
         Width           =   1125
      End
      Begin VB.TextBox txtSex 
         Height          =   285
         Left            =   10890
         MaxLength       =   6
         TabIndex        =   4
         Text            =   "Male"
         Top             =   990
         Width           =   1125
      End
      Begin VB.Frame Frame6 
         Height          =   1395
         Left            =   0
         TabIndex        =   21
         Top             =   0
         Width           =   2355
         Begin VB.ComboBox cMRU 
            Height          =   315
            Left            =   570
            TabIndex        =   30
            Text            =   "cMRU"
            Top             =   1020
            Width           =   1605
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
            Left            =   150
            MaxLength       =   12
            TabIndex        =   0
            Top             =   510
            Width           =   1785
         End
         Begin ComCtl2.UpDown UpDown1 
            Height          =   480
            Left            =   1920
            TabIndex        =   22
            Top             =   510
            Width           =   240
            _ExtentX        =   450
            _ExtentY        =   847
            _Version        =   327681
            Value           =   1
            BuddyControl    =   "txtSampleID"
            BuddyDispid     =   196690
            OrigLeft        =   1920
            OrigTop         =   540
            OrigRight       =   2160
            OrigBottom      =   1020
            Max             =   9999999
            Min             =   1
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "MRU"
            Height          =   195
            Left            =   150
            TabIndex        =   31
            Top             =   1080
            Width           =   375
         End
         Begin VB.Image iRelevant 
            Height          =   480
            Index           =   1
            Left            =   1500
            Picture         =   "frmEditSemen.frx":4EB8
            Top             =   120
            Width           =   480
         End
         Begin VB.Image iRelevant 
            Appearance      =   0  'Flat
            Height          =   480
            Index           =   0
            Left            =   120
            Picture         =   "frmEditSemen.frx":51C2
            Top             =   120
            Width           =   480
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            Caption         =   "Sample ID"
            Height          =   195
            Left            =   720
            TabIndex        =   23
            Top             =   30
            Width           =   735
         End
      End
      Begin VB.CommandButton bsearch 
         Appearance      =   0  'Flat
         Caption         =   "Se&arch"
         Height          =   345
         Left            =   9480
         TabIndex        =   20
         Top             =   180
         Width           =   675
      End
      Begin VB.CommandButton bDoB 
         Appearance      =   0  'Flat
         Caption         =   "S&earch"
         Height          =   285
         Left            =   12060
         TabIndex        =   19
         Top             =   270
         Width           =   705
      End
      Begin VB.Label lblAandE 
         Caption         =   "A and E"
         Height          =   225
         Left            =   4185
         TabIndex        =   85
         Top             =   330
         Width           =   615
      End
      Begin VB.Label lblChartNumber 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Monaghan Chart #"
         Height          =   285
         Left            =   2505
         TabIndex        =   34
         ToolTipText     =   "Click to change Location"
         Top             =   270
         Width           =   1545
      End
      Begin VB.Label lblAddWardGP 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2490
         TabIndex        =   33
         Top             =   1050
         Width           =   7665
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
         Left            =   7710
         TabIndex        =   32
         Top             =   210
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         Height          =   195
         Index           =   0
         Left            =   5940
         TabIndex        =   27
         Top             =   330
         Width           =   420
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         Caption         =   "D.o.B"
         Height          =   195
         Index           =   0
         Left            =   10440
         TabIndex        =   26
         Top             =   300
         Width           =   405
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "Age"
         Height          =   195
         Index           =   0
         Left            =   10530
         TabIndex        =   25
         Top             =   660
         Width           =   285
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         Caption         =   "Sex"
         Height          =   195
         Left            =   10560
         TabIndex        =   24
         Top             =   1020
         Width           =   270
      End
   End
   Begin VB.CommandButton bViewBB 
      Caption         =   "Transfusion Details"
      Height          =   885
      Left            =   13350
      Picture         =   "frmEditSemen.frx":54CC
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   90
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.CommandButton bPrint 
      Caption         =   "&Print"
      Height          =   750
      Left            =   13350
      Picture         =   "frmEditSemen.frx":5D96
      Style           =   1  'Graphical
      TabIndex        =   15
      Tag             =   "bprint"
      Top             =   3180
      Width           =   1275
   End
   Begin VB.CommandButton bFAX 
      Caption         =   "FAX"
      Height          =   750
      Index           =   0
      Left            =   13350
      Picture         =   "frmEditSemen.frx":60A0
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3960
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.CommandButton bCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   615
      Left            =   13350
      Picture         =   "frmEditSemen.frx":63AA
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   7710
      Width           =   1275
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   83
      Top             =   8640
      Width           =   14835
      _ExtentX        =   26167
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
End
Attribute VB_Name = "frmEditSemen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mNewRecord As Boolean

Private Activated As Boolean

Private pPrintToPrinter As String

Private SampleIDWithOffset As Double


Private Sub LoadSemenMorphology()

          Dim sql As String
          Dim tb As Recordset

10        On Error GoTo LoadSemenMorphology_Error

20        txtMorphology = ""
30        lblSemenMorph.Caption = GetOptionSetting("SemenMorphologyTitle", "% Abnormal Forms")

40        sql = "SELECT Result FROM GenericResults WHERE " & _
                "SampleID = '" & Val(txtSampleID) + SysOptSemenOffset(0) & "' " & _
                "AND TestName = 'SemenMorphResult'"
50        Set tb = New Recordset
60        RecOpenServer 0, tb, sql
70        If Not tb.EOF Then
80            txtMorphology = tb!Result & ""
90        End If

100       sql = "SELECT * FROM GenericResults WHERE " & _
                "SampleID = '" & Val(txtSampleID) + SysOptSemenOffset(0) & "' " & _
                "AND TestName = 'SemenMorphDescription'"
110       Set tb = New Recordset
120       RecOpenServer 0, tb, sql
130       If Not tb.EOF Then
140           lblSemenMorph = tb!Result & ""
150       End If

160       Exit Sub

LoadSemenMorphology_Error:

          Dim strES As String
          Dim intEL As Integer

170       intEL = Erl
180       strES = Err.Description
190       LogError "frmEditSemen", "LoadSemenMorphology", intEL, strES, sql

End Sub

Private Sub SaveSemenMorphology()

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo SaveSemenMorphology_Error

20        If Trim$(txtMorphology) = "" Then

30            sql = "DELETE FROM GenericResults WHERE " & _
                    "SampleID = '" & Val(txtSampleID) + SysOptSemenOffset(0) & "' " & _
                    "AND ( TestName = 'SemenMorphResult' " & _
                    "   OR TestName = 'SemenMorphDescription' )"
40            Cnxn(0).Execute sql

50        Else

60            sql = "SELECT * FROM GenericResults WHERE " & _
                    "SampleID = '" & Val(txtSampleID) + SysOptSemenOffset(0) & "' " & _
                    "AND TestName = 'SemenMorphResult'"
70            Set tb = New Recordset
80            RecOpenServer 0, tb, sql
90            If tb.EOF Then
100               tb.AddNew
110               tb!SampleID = Val(txtSampleID) + SysOptSemenOffset(0)
120           End If
130           tb!TestName = "SemenMorphResult"
140           tb!Result = Left$(txtMorphology, 50)
150           tb!UserName = UserName
160           tb.Update

170           sql = "SELECT * FROM GenericResults WHERE " & _
                    "SampleID = '" & Val(txtSampleID) + SysOptSemenOffset(0) & "' " & _
                    "AND TestName = 'SemenMorphDescription'"
180           Set tb = New Recordset
190           RecOpenServer 0, tb, sql
200           If tb.EOF Then
210               tb.AddNew
220               tb!SampleID = Val(txtSampleID) + SysOptSemenOffset(0)
230           End If
240           tb!TestName = "SemenMorphDescription"
250           tb!Result = lblSemenMorph
260           tb!UserName = UserName
270           tb.Update

280       End If

290       Exit Sub

SaveSemenMorphology_Error:

          Dim strES As String
          Dim intEL As Integer

300       intEL = Erl
310       strES = Err.Description
320       LogError "frmEditSemen", "SaveSemenMorphology", intEL, strES, sql

End Sub

Private Sub cmbCount_Click()

10        cmdSave.Enabled = True
20        cmdSaveHold.Enabled = True

30        fraMotility.Visible = True
40        If cmbCount = "Aspermatazoa/None Seen" Then
50            fraMotility.Visible = False
60            txtMotility(0) = ""
70            txtMotility(1) = ""
80            txtMotility(2) = ""
90            txtMotility(3) = ""
100           txtMotility(4) = ""
110       End If

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
100       LogError "frmEditSemen", "cmbHospital_LostFocus", intEL, strES

End Sub



Private Sub cmbSpecimenType_Click()
10        cmdSave.Enabled = True
20        cmdSaveHold.Enabled = True

End Sub

Private Sub cmdArchive_Click()

10        With frmAuditMicro
20            .SampleID = Val(txtSampleID)
30            .Show 1
40        End With

End Sub

Private Sub cmdCopyTo_Click()

          Dim s As String

10        s = cmbWard & " " & cmbClinician
20        s = Trim$(s) & " " & cmbGP
30        s = Trim$(s)

40        frmCopyTo.EditScreen = Me
50        frmCopyTo.lblOriginal = s
60        frmCopyTo.lblSampleID = txtSampleID + SysOptSemenOffset(0)
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
                "SampleID = '" & SysOptSemenOffset(0) + Val(txtSampleID) & "'"
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
170       LogError "frmEditSemen", "CheckCC", intEL, strES, sql

End Sub

Private Sub cmdDartViewer_Click()

10        On Error GoTo cmdDartViewer_Click_Error

20        Shell "C:\Program Files\The PlumTree Group\Dartviewer\Dartviewer.exe " & txtSampleID, vbNormalFocus

30        Exit Sub

cmdDartViewer_Click_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmEditSemen", "cmdDartViewer_Click", intEL, strES

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
70        LogError "frmEditSemen", "SetWardClinGP", intEL, strES

End Sub
Private Sub bcancel_Click()

10        pBar = 0

20        Unload Me

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
190       LogError "frmEditSemen", "bDoB_Click", intEL, strES

End Sub

Private Sub bFAX_Click(Index As Integer)

10        pBar = 0

End Sub

Private Sub bprint_Click()

          Dim tb As New Recordset
          Dim sql As String

10        On Error GoTo bprint_Click_Error

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

160       If Trim$(cmbWard) = "GP" Then
170           If Len(cmbGP) = 0 Then
180               iMsg "Must have Ward or GP entry.", vbCritical
190               Exit Sub
200           End If
210       End If

220       GetSampleIDWithOffset
230       SaveDemographics gNOCHANGE

240       LogTimeOfPrinting SampleIDWithOffset, "S"

250       sql = "SELECT * FROM PrintPending WHERE " & _
                "Department = 'Z' " & _
                "AND SampleID = '" & SampleIDWithOffset & "'"
260       Set tb = New Recordset
270       RecOpenClient 0, tb, sql
280       If tb.EOF Then
290           tb.AddNew
300       End If
310       tb!SampleID = SampleIDWithOffset
320       tb!Department = "Z"
330       tb!Ward = cmbWard
340       tb!Clinician = cmbClinician
350       tb!GP = cmbGP
360       tb!Initiator = UserName
370       tb!UsePrinter = pPrintToPrinter
380       tb.Update

390       txtSampleID = Format$(Val(txtSampleID) + 1)
400       GetSampleIDWithOffset

410       LoadAllDetails

420       Exit Sub

bprint_Click_Error:

          Dim strES As String
          Dim intEL As Integer

430       intEL = Erl
440       strES = Err.Description
450       LogError "frmEditSemen", "bPrint_Click", intEL, strES, sql

End Sub

Private Sub bPrintHold_Click()

          Dim tb As New Recordset
          Dim sql As String

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

220       GetSampleIDWithOffset
230       SaveDemographics gNOCHANGE

240       SaveSemen 1

250       LogTimeOfPrinting SampleIDWithOffset, "S"
260       sql = "SELECT * FROM PrintPending WHERE " & _
                "Department = 'Z' " & _
                "AND SampleID = '" & SampleIDWithOffset & "'"
270       Set tb = New Recordset
280       RecOpenClient 0, tb, sql
290       If tb.EOF Then
300           tb.AddNew
310       End If
320       tb!SampleID = SampleIDWithOffset
330       tb!Ward = cmbWard
340       tb!Clinician = cmbClinician
350       tb!GP = cmbGP
360       tb!Department = "Z"
370       tb!Initiator = UserName
380       tb!UsePrinter = pPrintToPrinter
390       tb.Update

400       Exit Sub

bPrintHold_Click_Error:

          Dim strES As String
          Dim intEL As Integer

410       intEL = Erl
420       strES = Err.Description
430       LogError "frmEditSemen", "bPrintHold_Click", intEL, strES, sql

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
130               FlashNoPrevious lNoPrevious
140           End If
150       End With

160       Exit Sub

bsearch_Click_Error:

          Dim strES As String
          Dim intEL As Integer

170       intEL = Erl
180       strES = Err.Description
190       LogError "frmEditSemen", "bsearch_Click", intEL, strES

End Sub

Private Sub bViewBB_Click()

10        On Error GoTo bViewBB_Click_Error

20        pBar = 0

30        If Trim$(txtChart) <> "" Then
40            frmViewBB.lChart = txtChart
50            frmViewBB.Show 1
60        End If

70        Exit Sub

bViewBB_Click_Error:

          Dim strES As String
          Dim intEL As Integer

80        intEL = Erl
90        strES = Err.Description
100       LogError "frmEditSemen", "bViewBB_Click", intEL, strES

End Sub

Private Sub CheckPrevious()

          Dim sql As String
          Dim tb As New Recordset

10        On Error GoTo CheckPrevious_Error

20        If Trim$(txtName) <> "" Then
30            sql = "SELECT D.SampleID from Demographics as D, SemenResults as S WHERE " & _
                    "PatName = '" & AddTicks(txtName) & "' and " & _
                    "D.SampleID < " & SampleIDWithOffset & " " & _
                    "and D.SampleID = S.SampleID"
40            Set tb = New Recordset
50            RecOpenServer 0, tb, sql
60            If Not tb.EOF Then
70                cmdSemenHistory.Visible = True
80            Else
90                cmdSemenHistory.Visible = False
100           End If
110       Else
120           cmdSemenHistory.Visible = False
130       End If

140       Exit Sub

CheckPrevious_Error:

          Dim strES As String
          Dim intEL As Integer

150       intEL = Erl
160       strES = Err.Description
170       LogError "frmEditSemen", "CheckPrevious", intEL, strES, sql

End Sub

Private Sub ClearSemen()

          Dim n As Long

10        On Error GoTo ClearSemen_Error

20        cmbVolume = ""
30        cmbCount = ""
40        cmbConsistency = ""
50        cmbSpecimenType = "Semen"
60        For n = 0 To 4
70            txtMotility(n) = ""
80        Next

90        Exit Sub

ClearSemen_Error:

          Dim strES As String
          Dim intEL As Integer

100       intEL = Erl
110       strES = Err.Description
120       LogError "frmEditSemen", "ClearSemen", intEL, strES

End Sub

Private Sub cmbClinDetails_Click()

10        On Error GoTo cmbClinDetails_Click_Error

20        cmdSaveHold.Enabled = True
30        cmdSave.Enabled = True

40        Exit Sub

cmbClinDetails_Click_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmEditSemen", "cmbClinDetails_Click", intEL, strES

End Sub

Private Sub cmbClinDetails_LostFocus()

          Dim NewVal As String

10        On Error GoTo cmbClinDetails_LostFocus_Error

20        pBar = 0

30        If Trim$(cmbClinDetails) = "" Then Exit Sub

40        NewVal = ListText("CD", cmbClinDetails)
50        If NewVal <> "" Then
60            cmbClinDetails = NewVal
70        End If

80        Exit Sub

cmbClinDetails_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

90        intEL = Erl
100       strES = Err.Description
110       LogError "frmEditSemen", "cmbClinDetails_LostFocus", intEL, strES

End Sub

Private Sub cmbClinician_Change()

10        On Error GoTo cmbClinician_Change_Error

20        SetWardClinGP

30        Exit Sub

cmbClinician_Change_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmEditSemen", "cmbClinician_Change", intEL, strES

End Sub

Private Sub cmbClinician_Click()

10        On Error GoTo cmbClinician_Click_Error

20        cmdSaveHold.Enabled = True
30        cmdSave.Enabled = True

40        Exit Sub

cmbClinician_Click_Error:

          Dim strES As String
          Dim intEL As Integer



50        intEL = Erl
60        strES = Err.Description
70        LogError "frmEditSemen", "cmbClinician_Click", intEL, strES


End Sub

Private Sub cmbClinician_KeyPress(KeyAscii As Integer)

10        On Error GoTo cmbClinician_KeyPress_Error

20        cmdSaveHold.Enabled = True
30        cmdSave.Enabled = True

40        Exit Sub

cmbClinician_KeyPress_Error:

          Dim strES As String
          Dim intEL As Integer



50        intEL = Erl
60        strES = Err.Description
70        LogError "frmEditSemen", "cmbClinician_KeyPress", intEL, strES


End Sub

Private Sub cmbClinician_LostFocus()

10        On Error GoTo cmbClinician_LostFocus_Error

20        pBar = 0
30        cmbClinician = QueryKnown("Clin", cmbClinician, cmbHospital)

40        Exit Sub

cmbClinician_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmEditSemen", "cmbClinician_LostFocus", intEL, strES

End Sub

Private Sub cmbConsistency_Click()

10        cmdSave.Enabled = True
20        cmdSaveHold.Enabled = True

End Sub

Private Sub cmbConsistency_KeyPress(KeyAscii As Integer)

10        On Error GoTo cmbConsistency_KeyPress_Error

20        Select Case chr$(KeyAscii)
          Case "W", "w": cmbConsistency = "Watery"
30        Case "M", "m": cmbConsistency = "Mucoid"
40        Case "P", "p": cmbConsistency = "Purulent"
50        Case Else: cmbConsistency = ""
60        End Select

70        KeyAscii = 0

80        cmdSave.Enabled = True
90        cmdSaveHold.Enabled = True

100       Exit Sub

cmbConsistency_KeyPress_Error:

          Dim strES As String
          Dim intEL As Integer



110       intEL = Erl
120       strES = Err.Description
130       LogError "frmEditSemen", "cmbConsistency_KeyPress", intEL, strES


End Sub

Private Sub cmbDemogComments_Click()

10        On Error GoTo cmbDemogComments_Click_Error

20        txtDemographicComment = txtDemographicComment & cmbDemogComments & " "
30        cmbDemogComments.ListIndex = -1

40        Exit Sub

cmbDemogComments_Click_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmEditSemen", "cmbDemogComments_Click", intEL, strES

End Sub

Private Sub cmbGP_Change()

10        On Error GoTo cmbGP_Change_Error

20        SetWardClinGP
30        cmbWard = "GP"

40        Exit Sub

cmbGP_Change_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmEditSemen", "cmbGP_Change", intEL, strES

End Sub

Private Sub cmbGP_Click()

10        On Error GoTo cmbGP_Click_Error

20        pBar = 0

30        cmbWard = "GP"

40        SetWardClinGP

50        cmdSave.Enabled = True
60        cmdSaveHold.Enabled = True

70        Exit Sub

cmbGP_Click_Error:

          Dim strES As String
          Dim intEL As Integer



80        intEL = Erl
90        strES = Err.Description
100       LogError "frmEditSemen", "cmbGP_Click", intEL, strES

End Sub


Private Sub cmbGP_KeyPress(KeyAscii As Integer)

10        On Error GoTo cmbGP_KeyPress_Error

20        cmdSave.Enabled = True
30        cmdSaveHold.Enabled = True
40        cmbWard = "GP"

50        Exit Sub

cmbGP_KeyPress_Error:

          Dim strES As String
          Dim intEL As Integer



60        intEL = Erl
70        strES = Err.Description
80        LogError "frmEditSemen", "cmbGP_KeyPress", intEL, strES


End Sub


Private Sub cmbGP_LostFocus()

10        On Error GoTo cmbGP_LostFocus_Error

20        cmbGP = QueryKnown("GP", cmbGP, cmbHospital)

30        Exit Sub

cmbGP_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmEditSemen", "cmbGP_LostFocus", intEL, strES

End Sub
Private Sub cmbHospital_Click()

10        On Error GoTo cmbHospital_Click_Error

20        FillGPsClinWard Me, cmbHospital

30        cmdSaveHold.Enabled = True
40        cmdSave.Enabled = True

50        Exit Sub

cmbHospital_Click_Error:

          Dim strES As String
          Dim intEL As Integer

60        intEL = Erl
70        strES = Err.Description
80        LogError "frmEditSemen", "cmbHospital_Click", intEL, strES

End Sub

Private Sub cmbSemenComments_Click()

10        On Error GoTo cmbSemenComments_Click_Error

20        txtSemenComment = txtSemenComment & cmbSemenComments & " "
30        cmbSemenComments.ListIndex = -1

40        Exit Sub

cmbSemenComments_Click_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmEditSemen", "cmbSemenComments_Click", intEL, strES

End Sub

Private Sub cmbVolume_Click()

10        cmdSave.Enabled = True
20        cmdSaveHold.Enabled = True

End Sub

Private Sub cmbWard_Change()

10        On Error GoTo cmbWard_Change_Error

20        SetWardClinGP

30        Exit Sub

cmbWard_Change_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmEditSemen", "cmbWard_Change", intEL, strES

End Sub

Private Sub cmbWard_Click()

10        On Error GoTo cmbWard_Click_Error

20        SetWardClinGP

30        cmdSaveHold.Enabled = True
40        cmdSave.Enabled = True

50        Exit Sub

cmbWard_Click_Error:

          Dim strES As String
          Dim intEL As Integer



60        intEL = Erl
70        strES = Err.Description
80        LogError "frmEditSemen", "cmbWard_Click", intEL, strES

End Sub

Private Sub cmbWard_KeyPress(KeyAscii As Integer)

10        On Error GoTo cmbWard_KeyPress_Error

20        cmdSaveHold.Enabled = True
30        cmdSave.Enabled = True

40        Exit Sub

cmbWard_KeyPress_Error:

          Dim strES As String
          Dim intEL As Integer



50        intEL = Erl
60        strES = Err.Description
70        LogError "frmEditSemen", "cmbWard_KeyPress", intEL, strES


End Sub

Private Sub cmbWard_LostFocus()

          Dim tb As New Recordset
          Dim sql As String

10        On Error GoTo cmbWard_LostFocus_Error

20        If Trim$(cmbWard) = "" Then
30            cmbWard = "GP"
40            Exit Sub
50        End If

60        sql = "SELECT Text FROM Wards WHERE " & _
                "Text = '" & cmbWard & "' " & _
                "OR Code = '" & cmbWard & "'  And InUse = '1'"
70        Set tb = New Recordset
80        RecOpenServer 0, tb, sql
90        If Not tb.EOF Then
100           cmbWard = Trim(tb!Text)
110       Else
120           cmbWard = "GP"
130       End If

140       Exit Sub

cmbWard_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

150       intEL = Erl
160       strES = Err.Description
170       LogError "frmEditSemen", "cmbWard_LostFocus", intEL, strES, sql

End Sub

Private Sub cmdOrderSemen_Click()

          Dim tb As New Recordset
          Dim sql As String

10        On Error GoTo cmdOrderSemen_Click_Error

20        GetSampleIDWithOffset

30        sql = "SELECT * from SemenResults WHERE " & _
                "SampleID = '" & SampleIDWithOffset & "'"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        If tb.EOF Then
70            tb.AddNew
80            tb!SampleID = SampleIDWithOffset
90            tb.Update
100       End If

110       pBar = 0

120       Exit Sub

cmdOrderSemen_Click_Error:

          Dim strES As String
          Dim intEL As Integer

130       intEL = Erl
140       strES = Err.Description
150       LogError "frmEditSemen", "cmdOrderSemen_Click", intEL, strES, sql

End Sub

Private Sub cmdSave_Click()

10        On Error GoTo cmdSave_Click_Error

20        pBar = 0

30        GetSampleIDWithOffset


40        If Trim$(txtSampleID) = "" Then
50            iMsg "Must have Lab Number.", vbCritical
60            Exit Sub
70        End If

80        If Trim$(txtName) <> "" Then
90            If Trim$(cmbWard) = "" Then
100               iMsg "Must have Ward entry.", vbCritical
110               Exit Sub
120           End If

130           If Trim$(cmbWard) = "GP" Then
140               If Trim$(cmbGP) = "" Then
150                   iMsg "Must have GP entry.", vbCritical
160                   Exit Sub
170               End If
180           End If
190       End If

200       If lblChartNumber.BackColor = vbRed Then
210           If iMsg("Confirm this Patient has" & vbCrLf & _
                      lblChartNumber.Caption, vbYesNo + vbQuestion, , vbRed, 12) = vbNo Then
220               Exit Sub
230           End If
240       End If

250       If dtRunDate < dtSampleDate Then
260           iMsg "Sample Date After Run Date. Please Amend!"
270           Exit Sub
280       End If

290       cmdSaveHold.Caption = "Saving"

300       GetSampleIDWithOffset

310       SaveDemographics gNOCHANGE
320       SaveSemen 0
330       SaveComments
340       UPDATEMRU txtSampleID, cMRU
350       cmdSaveHold.Caption = "Save && &Hold"
360       cmdSaveHold.Enabled = False
370       cmdSave.Enabled = False

380       SaveSetting "NetAcquire", "StartUp", "LastUsedSemen", txtSampleID

390       txtSampleID = Format$(Val(txtSampleID) + 1)
400       GetSampleIDWithOffset

410       LoadAllDetails

420       cmdSave.Enabled = True
430       cmdSaveHold.Enabled = True
440       txtSampleID.SelStart = 0
450       txtSampleID.SelLength = 999
460       txtSampleID.SetFocus

470       Exit Sub

cmdSave_Click_Error:

          Dim strES As String
          Dim intEL As Integer

480       intEL = Erl
490       strES = Err.Description
500       LogError "frmEditSemen", "cmdsave_Click", intEL, strES

End Sub

Private Sub cmdSaveHold_Click()

10        On Error GoTo cmdSaveHold_Click_Error

20        pBar = 0

30        If Trim$(txtSampleID) = "" Then
40            iMsg "Must have Lab Number.", vbCritical
50            Exit Sub
60        End If

70        If Trim$(txtName) <> "" Then
80            If Trim$(cmbWard) = "" Then
90                iMsg "Must have Ward entry.", vbCritical
100               Exit Sub
110           End If

120           If Trim$(cmbWard) = "GP" Then
130               If Trim$(cmbGP) = "" Then
140                   iMsg "Must have GP entry.", vbCritical
150                   Exit Sub
160               End If
170           End If
180       End If

190       If dtRunDate < dtSampleDate Then
200           iMsg "Sample Date After Run Date. Please Amend!"
210           Exit Sub
220       End If

230       If lblChartNumber.BackColor = vbRed Then
240           If iMsg("Confirm this Patient has" & vbCrLf & _
                      lblChartNumber.Caption, vbYesNo + vbQuestion, , vbRed, 12) = vbNo Then
250               Exit Sub
260           End If
270       End If

280       cmdSaveHold.Caption = "Saving"

290       GetSampleIDWithOffset

300       SaveDemographics gNOCHANGE
310       SaveSemen 0
320       SaveComments
330       UPDATEMRU txtSampleID, cMRU

340       cmdSaveHold.Caption = "Save && &Hold"
350       cmdSaveHold.Enabled = False
360       cmdSave.Enabled = False

370       Exit Sub

cmdSaveHold_Click_Error:

          Dim strES As String
          Dim intEL As Integer

380       intEL = Erl
390       strES = Err.Description
400       LogError "frmEditSemen", "cmdSaveHold_Click", intEL, strES

End Sub

Private Sub cmdSemenHistory_Click()

10        On Error GoTo cmdSemenHistory_Click_Error

20        With frmSemenHistory
30            .lblName = txtName
40            .Show 1
50        End With

60        Exit Sub

cmdSemenHistory_Click_Error:

          Dim strES As String
          Dim intEL As Integer

70        intEL = Erl
80        strES = Err.Description
90        LogError "frmEditSemen", "cmdSemenHistory_Click", intEL, strES

End Sub

Private Sub cmdSetPrinter_Click()

10        On Error GoTo cmdSetPrinter_Click_Error

20        Set frmForcePrinter.f = frmEditSemen
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
180       LogError "frmEditSemen", "cmdSetPrinter_Click", intEL, strES

End Sub

Private Sub cmdvalidate_Click()

10        On Error GoTo cmdvalidate_Click_Error

20        pBar = 0

30        If Trim$(txtSampleID) = "" Then
40            iMsg "Must have Lab Number.", vbCritical
50            Exit Sub
60        End If

70        If Trim$(txtName) <> "" Then
80            If Trim$(cmbWard) = "" Then
90                iMsg "Must have Ward entry.", vbCritical
100               Exit Sub
110           End If

120           If Trim$(cmbWard) = "GP" Then
130               If Trim$(cmbGP) = "" Then
140                   iMsg "Must have GP entry.", vbCritical
150                   Exit Sub
160               End If
170           End If
180       End If

190       If lblChartNumber.BackColor = vbRed Then
200           If iMsg("Confirm this Patient has" & vbCrLf & _
                      lblChartNumber.Caption, vbYesNo + vbQuestion, , vbRed, 12) = vbNo Then
210               Exit Sub
220           End If
230       End If

240       GetSampleIDWithOffset

250       If cmdValidate.Caption = "&VALID" Then
260           If UCase(iBOX("Unvalidate ! Enter Password", , , True)) = UCase(UserPass) Then
270               SaveDemographics 0
280               SaveSemen 0
290               SaveComments
300               LoadAllDetails
310               Me.Refresh
320               Exit Sub
330           End If
340       Else
350           cmdValidate.Caption = "Validating"
360           SaveDemographics 1
370           SaveSemen 1
380           SaveComments
390       End If

400       UPDATEMRU txtSampleID, cMRU

410       GetSampleIDWithOffset
420       LoadAllDetails

430       Exit Sub

cmdvalidate_Click_Error:

          Dim strES As String
          Dim intEL As Integer

440       intEL = Erl
450       strES = Err.Description
460       LogError "frmEditSemen", "cmdvalidate_Click", intEL, strES

End Sub

Private Sub cmdViewMicroRep_Click()

10        On Error GoTo cmdViewMicroRep_Click_Error

20        frmRFT.SampleID = Val(txtSampleID) + SysOptSemenOffset(0)
30        frmRFT.Dept = "Z"
40        frmRFT.Show 1

50        Exit Sub

cmdViewMicroRep_Click_Error:

          Dim strES As String
          Dim intEL As Integer

60        intEL = Erl
70        strES = Err.Description
80        LogError "frmEditSemen", "cmdViewMicroRep_Click", intEL, strES

End Sub

Private Sub cMRU_Click()

10        On Error GoTo cMRU_Click_Error

20        txtSampleID = cMRU
30        GetSampleIDWithOffset

40        LoadAllDetails

50        cmdSaveHold.Enabled = False
60        cmdSave.Enabled = False

70        Exit Sub

cMRU_Click_Error:

          Dim strES As String
          Dim intEL As Integer

80        intEL = Erl
90        strES = Err.Description
100       LogError "frmEditSemen", "cMRU_Click", intEL, strES

End Sub

Private Sub cMRU_KeyPress(KeyAscii As Integer)

10        KeyAscii = 0

End Sub

Private Sub cRooH_Click(Index As Integer)

10        On Error GoTo cRooH_Click_Error

20        cmdSaveHold.Enabled = True
30        cmdSave.Enabled = True

40        Exit Sub

cRooH_Click_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmEditSemen", "cRooH_Click", intEL, strES

End Sub

Private Sub dtRecDate_CloseUp()

10        On Error GoTo dtRecDate_CloseUp_Error

20        pBar = 0

30        cmdSaveHold.Enabled = True
40        cmdSave.Enabled = True

50        Exit Sub

dtRecDate_CloseUp_Error:

          Dim strES As String
          Dim intEL As Integer

60        intEL = Erl
70        strES = Err.Description
80        LogError "frmEditSemen", "dtRecDate_CloseUp", intEL, strES

End Sub

Private Sub dtRecDate_LostFocus()

10        SetDatesColour Me

End Sub


Private Sub dtRunDate_CloseUp()

10        On Error GoTo dtRunDate_CloseUp_Error

20        pBar = 0

30        cmdSaveHold.Enabled = True
40        cmdSave.Enabled = True

50        Exit Sub

dtRunDate_CloseUp_Error:

          Dim strES As String
          Dim intEL As Integer

60        intEL = Erl
70        strES = Err.Description
80        LogError "frmEditSemen", "dtRunDate_CloseUp", intEL, strES

End Sub

Private Sub dtRunDate_LostFocus()

10        SetDatesColour Me

End Sub


Private Sub dtSampleDate_CloseUp()

10        On Error GoTo dtSampleDate_CloseUp_Error

20        pBar = 0

30        cmdSaveHold.Enabled = True
40        cmdSave.Enabled = True

50        Exit Sub

dtSampleDate_CloseUp_Error:

          Dim strES As String
          Dim intEL As Integer

60        intEL = Erl
70        strES = Err.Description
80        LogError "frmEditSemen", "dtSampleDate_CloseUp", intEL, strES

End Sub

Private Sub FillLists()

          Dim sql As String

10        On Error GoTo FillLists_Error

20        FillGPsClinWard Me, HospName(0)

          'cmbHospital.Clear

30        LoadListGeneric cmbHospital, "HO"
40        LoadListGeneric cmbConsistency, "SemenConsistency"
50        LoadListGeneric cmbVolume, "SemenVolume"
60        LoadListGeneric cmbCount, "SemenCount"
70        LoadListGeneric cmbSpecimenType, "SemenSpecimenType"
80        LoadListGeneric cmbSemenComments, "SE"
90        LoadListGeneric cmbDemogComments, "DE"
100       LoadListGeneric cmbClinDetails, "CD"

110       FixComboWidth cmbHospital
120       FixComboWidth cmbConsistency
130       FixComboWidth cmbVolume
140       FixComboWidth cmbCount
150       FixComboWidth cmbSpecimenType
160       FixComboWidth cmbSemenComments
170       FixComboWidth cmbDemogComments
180       FixComboWidth cmbClinDetails

          'sql = "SELECT * from lists WHERE listtype = 'HO'"
          'Set tb = New Recordset
          'RecOpenServer 0, tb, sql
          'Do While Not tb.EOF
          '      cmbHospital.AddItem Trim(tb!Text)
          '  tb.MoveNext
          'Loop
          '
          'With cmbConsistency
          '  .Clear
          '  .AddItem ""
          '  .AddItem "Watery"
          '  .AddItem "Mucoid"
          '  .AddItem "Purulent"
          'End With
          '
          'With cmbVolume
          '  .Clear
          '  .AddItem ""
          '  For sngVol = 10 To 0.5 Step -0.5
          '    .AddItem Format$(sngVol, "0.0")
          '  Next
          'End With
          '
          'With cmbCount
          '  .Clear
          '  .AddItem "Aspermatazoa/None Seen"
          '  .AddItem "< 1"
          '  .AddItem "2"
          '  .AddItem "5"
          '  .AddItem "10"
          '  .AddItem "20"
          '  .AddItem "40"
          '  .AddItem "60"
          '  .AddItem "> 60"
          'End With

          'cmbSemenComments.Clear
          'cmbDemogComments.Clear
          'cmbClinDetails.Clear
          '
          'sql = "SELECT * from lists"
          'Set tb = New Recordset
          'RecOpenServer 0, tb, sql
          'Do While Not tb.EOF
          '  If Trim(tb!ListType) = "SE" Then
          '    cmbSemenComments.AddItem Trim(tb!Text)
          '  ElseIf Trim(tb!ListType) = "DE" Then
          '    cmbDemogComments.AddItem Trim(tb!Text)
          '  ElseIf Trim(tb!ListType) = "CD" Then
          '    cmbClinDetails.AddItem Trim(tb!Text)
          '  End If
          '  tb.MoveNext
          'Loop

190       Exit Sub

FillLists_Error:

          Dim strES As String
          Dim intEL As Integer

200       intEL = Erl
210       strES = Err.Description
220       LogError "frmEditSemen", "FillLists", intEL, strES, sql

End Sub

Private Sub dtSampleDate_LostFocus()

10        SetDatesColour Me

End Sub

Private Sub Form_Activate()

10        TimerBar.Enabled = True
20        pBar = 0

End Sub

Private Sub Form_Deactivate()

10        pBar = 0
20        TimerBar.Enabled = False

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

10        pBar = 0

End Sub

Private Sub Form_Load()

10        On Error GoTo Form_Load_Error

20        lblSemenMorph.Caption = GetOptionSetting("SemenMorphologyTitle", "% Abnormal Forms")

30        FillLists

40        FillMRU cMRU

50        cmdViewMicroRep.Visible = SysOptRTFView(0)

60        With lblChartNumber
70            .BackColor = &H8000000F
80            .ForeColor = vbBlack
90        End With

100       dtRunDate = Format$(Now, "dd/mm/yyyy")
110       dtRecDate = Format$(Now, "dd/mm/yyyy")
120       dtSampleDate = Format$(Now, "dd/mm/yyyy")

130       UpDown1.Max = 99999999

140       txtSampleID = GetSetting("NetAcquire", "StartUp", "LastUsedSemen", "1")
150       GetSampleIDWithOffset
160       LoadAllDetails

170       cmdSaveHold.Enabled = False
180       cmdSave.Enabled = False

190       StatusBar1.Panels(1).Text = UserName

200       txtSampleID.SelStart = 0
210       txtSampleID.SelLength = 999

220       Activated = False

230       Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

240       intEL = Erl
250       strES = Err.Description
260       LogError "frmEditSemen", "Form_Load", intEL, strES

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        pBar = 0

End Sub

Private Sub Form_Unload(Cancel As Integer)

10        On Error GoTo Form_Unload_Error

20        If Val(txtSampleID) > Val(GetSetting("NetAcquire", "StartUp", "LastUsedSemen", "1")) Then
30            SaveSetting "NetAcquire", "StartUp", "LastUsedSemen", txtSampleID
40        End If

50        pPrintToPrinter = ""

60        Activated = False

70        Exit Sub

Form_Unload_Error:

          Dim strES As String
          Dim intEL As Integer



80        intEL = Erl
90        strES = Err.Description
100       LogError "frmEditSemen", "Form_Unload", intEL, strES


End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        pBar = 0

End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        pBar = 0

End Sub

Private Sub Frame5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        pBar = 0

End Sub

Private Sub Frame6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        pBar = 0

End Sub

Private Sub Frame8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        pBar = 0

End Sub

Private Sub GetSampleIDWithOffset()

10        On Error GoTo GetSampleIDWithOffset_Error

20        SampleIDWithOffset = Val(txtSampleID) + Val(SysOptSemenOffset(0))

30        Exit Sub

GetSampleIDWithOffset_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmEditSemen", "GetSampleIDWithOffset", intEL, strES


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

100       cmdSave.Enabled = True
110       cmdSaveHold.Enabled = True

120       Exit Sub

iRecDate_Click_Error:

          Dim strES As String
          Dim intEL As Integer

130       intEL = Erl
140       strES = Err.Description
150       LogError "frmEditSemen", "iRecDate_Click", intEL, strES

End Sub

Private Sub irelevant_Click(Index As Integer)

          Dim sql As String
          Dim tb As New Recordset
          Dim strDirection As String

10        On Error GoTo irelevant_Click_Error

20        strDirection = IIf(Index = 0, "<", ">")
30        GetSampleIDWithOffset

40        sql = "SELECT top 1 SampleID from SemenResults WHERE " & _
                "cast(SampleID as numeric) " & strDirection & " '" & SampleIDWithOffset & "' " & _
                "Order by SampleID " & IIf(strDirection = "<", "Desc", "Asc")

50        Set tb = New Recordset
60        RecOpenClient 0, tb, sql
70        If Not tb.EOF Then
80            txtSampleID = Val(tb!SampleID & "") - SysOptSemenOffset(0)
90        End If

100       GetSampleIDWithOffset
110       LoadAllDetails

120       cmdSaveHold.Enabled = False
130       cmdSave.Enabled = False

140       Exit Sub

irelevant_Click_Error:

          Dim strES As String
          Dim intEL As Integer

150       intEL = Erl
160       strES = Err.Description
170       LogError "frmEditSemen", "irelevant_Click", intEL, strES, sql

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

90        SetDatesColour Me

100       cmdSave.Enabled = True
110       cmdSaveHold.Enabled = True

120       Exit Sub

iRunDate_Click_Error:

          Dim strES As String
          Dim intEL As Integer

130       intEL = Erl
140       strES = Err.Description
150       LogError "frmEditSemen", "iRunDate_Click", intEL, strES

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

90        SetDatesColour Me

100       cmdSave.Enabled = True
110       cmdSaveHold.Enabled = True

120       Exit Sub

iSampleDate_Click_Error:

          Dim strES As String
          Dim intEL As Integer

130       intEL = Erl
140       strES = Err.Description
150       LogError "frmEditSemen", "iSampleDate_Click", intEL, strES

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

170       SetDatesColour Me

180       cmdSave.Enabled = True
190       cmdSaveHold.Enabled = True

200       Exit Sub

iToday_Click_Error:

          Dim strES As String
          Dim intEL As Integer

210       intEL = Erl
220       strES = Err.Description
230       LogError "frmEditSemen", "iToday_Click", intEL, strES


End Sub

Private Sub lblChartNumber_Click()

10        On Error GoTo lblChartNumber_Click_Error

20        With lblChartNumber
30            .BackColor = &H8000000F
40            .ForeColor = vbBlack
50        End With

60        If Trim$(txtChart) <> "" Then
70            LoadPatientFromChart Me, mNewRecord
80            cmdSaveHold.Enabled = True
90            cmdSave.Enabled = True
100       End If

110       Exit Sub

lblChartNumber_Click_Error:

          Dim strES As String
          Dim intEL As Integer

120       intEL = Erl
130       strES = Err.Description
140       LogError "frmEditSemen", "lblChartNumber_Click", intEL, strES

End Sub

Private Sub LoadAllDetails()

10        On Error GoTo LoadAllDetails_Error

20        LoadDemographics
30        LoadSemen
40        LoadSemenMorphology
50        LoadComments

60        CheckCC

70        Exit Sub

LoadAllDetails_Error:

          Dim strES As String
          Dim intEL As Integer

80        intEL = Erl
90        strES = Err.Description
100       LogError "frmEditSemen", "LoadAllDetails", intEL, strES

End Sub

Private Sub LoadComments()

          Dim Ob As Observation
          Dim Obs As Observations

10        On Error GoTo LoadComments_Error

20        txtSemenComment = ""
30        txtDemographicComment = ""

40        If Trim$(txtSampleID) = "" Then Exit Sub

50        Set Obs = New Observations
60        Set Obs = Obs.Load(SampleIDWithOffset, "Demographic", "Semen")
70        If Not Obs Is Nothing Then
80            For Each Ob In Obs
90                Select Case UCase$(Ob.Discipline)
                  Case "DEMOGRAPHIC": txtDemographicComment = Ob.Comment
100               Case "SEMEN": txtSemenComment = Ob.Comment
110               End Select
120           Next
130       End If

140       Exit Sub

LoadComments_Error:

          Dim strES As String
          Dim intEL As Integer

150       intEL = Erl
160       strES = Err.Description
170       LogError "frmEditSemen", "LoadComments", intEL, strES

End Sub

Private Sub LoadDemographics()

          Dim sql As String
          Dim tb As New Recordset
          Dim SampleDate As String
          Dim RooH As Boolean

10        On Error GoTo LoadDemographics_Error

20        RooH = IsRoutine()
30        cRooH(0) = RooH
40        cRooH(1) = Not RooH
50        bViewBB.Enabled = False

60        If Trim$(txtSampleID) = "" Then Exit Sub

70        sql = "SELECT * FROM Demographics WHERE " & _
                "SampleID = '" & SampleIDWithOffset & "'"

80        Set tb = New Recordset
90        RecOpenServer 0, tb, sql
100       If tb.EOF Then
110           mNewRecord = True
120           dtRunDate = Format$(Now, "dd/MM/yyyy")
130           dtRecDate = Format$(Now, "dd/MM/yyyy")
140           dtSampleDate = Format$(Now, "dd/MM/yyyy")
150           txtChart = ""
160           txtAandE = ""
170           txtName = ""
180           taddress(0) = ""
190           taddress(1) = ""
200           txtSex = ""
210           txtDoB = ""
220           txtAge = ""
230           cmbWard = "GP"
240           cmbClinician = ""
250           cmbGP = ""
260           cmbHospital = initial2upper(HospName(0))
270           txtDemographicComment = ""
280           tSampleTime.Mask = ""
290           tSampleTime.Text = ""
300           tSampleTime.Mask = "##:##"
310           lblChartNumber.Caption = HospName(0) & " Chart #"
320           lblChartNumber.BackColor = &H8000000F
330           lblChartNumber.ForeColor = vbBlack
340           cmbClinDetails = ""
350       Else
360           If Trim$(tb!Hospital & "") <> "" Then
370               cmbHospital = Trim$(tb!Hospital)
380               lblChartNumber = Trim$(tb!Hospital) & " Chart #"
390               If UCase(tb!Hospital) = UCase(HospName(0)) Then
400                   lblChartNumber.BackColor = &H8000000F
410                   lblChartNumber.ForeColor = vbBlack
420               Else
430                   lblChartNumber.BackColor = vbRed
440                   lblChartNumber.ForeColor = vbYellow
450               End If
460           Else
470               cmbHospital = initial2upper(HospName(0))
480               lblChartNumber.Caption = initial2upper(HospName(0)) & " Chart #"
490               lblChartNumber.BackColor = &H8000000F
500               lblChartNumber.ForeColor = vbBlack
510           End If
520           If IsDate(tb!Rundate) Then
530               dtRunDate = Format$(tb!Rundate, "dd/mm/yyyy")
540           Else
550               dtRunDate = Format$(Now, "dd/mm/yyyy")
560           End If
570           If IsDate(tb!RecDate) Then
580               dtRecDate = Format$(tb!RecDate, "dd/mm/yyyy")
590           Else
600               dtRecDate = dtRunDate
610           End If
620           StatusBar1.Panels(4).Text = dtRunDate
630           mNewRecord = False
640           If tb!RooH & "" <> "" Then cRooH(0) = tb!RooH
650           If tb!RooH & "" <> "" Then cRooH(1) = Not tb!RooH
660           txtChart = tb!Chart & ""
670           txtAandE = Trim$(tb!AandE & "")
680           txtName = tb!PatName & ""
690           taddress(0) = tb!Addr0 & ""
700           taddress(1) = tb!Addr1 & ""
710           Select Case Left$(Trim$(UCase$(tb!sex & "")), 1)
              Case "M": txtSex = "Male"
720           Case "F": txtSex = "Female"
730           Case Else: txtSex = ""
740           End Select
750           txtDoB = Format$(tb!Dob, "dd/mm/yyyy")
760           txtAge = tb!Age & ""
770           cmbClinician = tb!Clinician & ""
780           cmbGP = tb!GP & ""
790           cmbWard = tb!Ward & ""
800           cmbClinDetails = tb!ClDetails & ""
810           If IsDate(tb!SampleDate) Then
820               dtSampleDate = Format$(tb!SampleDate, "dd/mm/yyyy")
830               If Format$(tb!SampleDate, "hh:mm") <> "00:00" Then
840                   tSampleTime = Format$(tb!SampleDate, "hh:mm")
850               Else
860                   tSampleTime.Mask = ""
870                   tSampleTime.Text = ""
880                   tSampleTime.Mask = "##:##"
890               End If
900           Else
910               dtSampleDate = Format$(Now, "dd/mm/yyyy")
920               tSampleTime.Mask = ""
930               tSampleTime.Text = ""
940               tSampleTime.Mask = "##:##"
950           End If
960       End If
970       cmdSaveHold.Enabled = False
980       cmdSave.Enabled = False

990       CheckPrevious

1000      Exit Sub

LoadDemographics_Error:

          Dim strES As String
          Dim intEL As Integer

1010      intEL = Erl
1020      strES = Err.Description
1030      LogError "frmEditSemen", "LoadDemographics", intEL, strES, sql

End Sub

Private Sub LoadSemen()

          Dim tb As New Recordset
          Dim sql As String

10        On Error GoTo LoadSemen_Error

20        sql = "SELECT * from SemenResults WHERE " & _
                "SampleID = '" & SampleIDWithOffset & "'"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql

50        cmdValidate.Caption = "&Validate"
60        fraSpecimen.Enabled = True
70        fraDemographics.Enabled = True
80        fraDate.Enabled = True

90        If tb.EOF Then
100           ClearSemen
110       Else
120           cmbCount = Trim$(tb!SemenCount & "")
130           cmbVolume = Trim$(tb!Volume & "")
140           cmbConsistency = Trim$(tb!Consistency & "")
150           cmbSpecimenType = Trim$(tb!SpecimenType & "")
160           txtMotility(0) = Trim$(tb!MotilityPro & "")
170           txtMotility(1) = Trim$(tb!MotilityNonPro & "")
180           txtMotility(2) = Trim$(tb!MotilityNonMotile & "")
190           txtMotility(3) = Trim$(tb!Motility & "")
200           txtMotility(4) = Trim$(tb!MotilitySlow & "")
210           If Not IsNull(tb!Valid) Then
220               If tb!Valid = 1 Then
230                   cmdValidate.Caption = "&VALID"
240                   fraSpecimen.Enabled = False
250                   fraDemographics.Enabled = False
260                   fraDate.Enabled = False
270               End If
280           End If
290       End If

300       fraMotility.Visible = True
310       If cmbCount = "Aspermatazoa/None Seen" Then
320           fraMotility.Visible = False
330           txtMotility(0) = ""
340           txtMotility(1) = ""
350           txtMotility(2) = ""
360           txtMotility(3) = ""
370           txtMotility(4) = ""
380       End If

390       cmdArchive.Visible = IsAnyRecordPresent("SemenResultsAudit", SampleIDWithOffset) Or _
                               IsAnyRecordPresent("DemographicsAudit", SampleIDWithOffset)

400       Exit Sub

LoadSemen_Error:

          Dim strES As String
          Dim intEL As Integer

410       intEL = Erl
420       strES = Err.Description
430       LogError "frmEditSemen", "LoadSemen", intEL, strES, sql

End Sub

Public Property Let PrintToPrinter(ByVal strNewValue As String)
Attribute PrintToPrinter.VB_HelpID = 2402

10        On Error GoTo PrintToPrinter_Error

20        pPrintToPrinter = strNewValue

30        Exit Property

PrintToPrinter_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmEditSemen", "PrintToPrinter", intEL, strES


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
60        LogError "frmEditSemen", "PrintToPrinter", intEL, strES


End Property

Private Sub SaveComments()

          Dim Obs As New Observations

10        On Error GoTo SaveComments_Error

20        If Trim$(txtSampleID) = "" Then Exit Sub

30        Obs.Save SampleIDWithOffset, True, _
                   "Semen", Trim$(txtSemenComment), _
                   "Demographic", Trim$(txtDemographicComment)

40        Exit Sub

SaveComments_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmEditSemen", "SaveComments", intEL, strES

End Sub

Private Sub SaveDemographics(ByVal Validate As Integer)

          Dim sql As String
          Dim tb As New Recordset

10        On Error GoTo SaveDemographics_Error

20        SaveComments

30        If Trim$(tSampleTime) <> "__:__" Then
40            If Not IsDate(tSampleTime) Then
50                iMsg "Invalid Time", vbExclamation
60                Exit Sub
70            End If
80        End If

90        sql = "SELECT * FROM Demographics WHERE " & _
                "SampleID = '" & SampleIDWithOffset & "'"
100       Set tb = New Recordset
110       RecOpenClient 0, tb, sql
120       If tb.EOF Then
130           tb.AddNew
140           tb!Fasting = 0
150           tb!Faxed = 0
160       End If

170       tb!RooH = cRooH(0)

180       tb!Rundate = Format$(dtRunDate, "dd/mmm/yyyy")
190       If IsDate(tSampleTime) Then
200           tb!SampleDate = Format$(dtSampleDate, "dd/mmm/yyyy") & " " & Format$(tSampleTime, "hh:mm")
210       Else
220           tb!SampleDate = Format$(dtSampleDate, "dd/mmm/yyyy")
230       End If

240       If IsDate(tRecTime) Then
250           tb!RecDate = Format$(dtRecDate, "dd/MMM/yyyy") & " " & Format$(tRecTime, "HH:nn")
260       Else
270           tb!RecDate = Format$(dtRecDate, "dd/MMM/yyyy")
280       End If

290       tb!SampleID = SampleIDWithOffset
300       tb!Chart = txtChart
310       tb!AandE = txtAandE
320       tb!PatName = Trim$(txtName)
330       If IsDate(txtDoB) Then
340           tb!Dob = Format$(txtDoB, "dd/mmm/yyyy")
350       Else
360           tb!Dob = Null
370       End If
380       tb!Age = txtAge
390       tb!sex = Left$(txtSex, 1)
400       tb!Addr0 = taddress(0)
410       tb!Addr1 = taddress(1)
420       tb!Ward = Left$(cmbWard, 50)
430       tb!Clinician = Left$(cmbClinician, 50)
440       tb!GP = Left$(cmbGP, 50)
450       tb!ClDetails = Left$(cmbClinDetails, 30)
460       tb!Hospital = HospName(0)
470       If Validate = 0 Or Validate = 1 Then
480           tb!Valid = Validate
490       End If

500       tb.Update

510       LogTimeOfPrinting SampleIDWithOffset, "D"



520       Exit Sub

SaveDemographics_Error:

          Dim strES As String
          Dim intEL As Integer



530       intEL = Erl
540       strES = Err.Description
550       LogError "frmEditSemen", "SaveDemographics", intEL, strES

End Sub

Private Sub SaveSemen(Optional ByVal Valid As Integer = 0)

          Dim sql As String

10        On Error GoTo SaveSemen_Error

          'Created on 08/10/2010 10:41:52
          'Autogenerated by SQL Scripting

20        sql = "If Exists(Select 1 From SemenResults " & _
                "          Where SampleID = @SampleID ) " & _
                "    Update SemenResults " & _
                "    SET Volume = '@Volume', " & _
                "    SemenCount = '@SemenCount', " & _
                "    MotilityPro = '@MotilityPro', " & _
                "    MotilityNonPro = '@MotilityNonPro', " & _
                "    MotilityNonMotile = '@MotilityNonMotile', " & _
                "    Consistency = '@Consistency', " & _
                "    Valid = @Valid, " & _
                "    SpecimenType = '@SpecimenType', " & _
                "    UserName = '@UserName', " & _
                "    MotilitySlow = '@MotilitySlow', " & _
                "    Motility = '@Motility' " & _
                "    WHERE SampleID = @SampleID  " & _
                "ELSE " & _
                "    INSERT INTO SemenResults (SampleID, Volume, SemenCount, MotilityPro, " & _
                "    MotilityNonPro, MotilityNonMotile, Consistency, Valid, SpecimenType, " & _
                "    UserName, MotilitySlow, Motility) VALUES " & _
                "    (@SampleID, '@Volume', '@SemenCount', '@MotilityPro', '@MotilityNonPro', " & _
                "    '@MotilityNonMotile', '@Consistency', @Valid, '@SpecimenType', " & _
                "    '@UserName', '@MotilitySlow', '@Motility') "

30        sql = Replace(sql, "@SampleID", SampleIDWithOffset)
40        sql = Replace(sql, "@Volume", Left$(cmbVolume, 5))
50        sql = Replace(sql, "@SemenCount", Left$(cmbCount, 50))
60        sql = Replace(sql, "@MotilityPro", Left$(txtMotility(0), 5))
70        sql = Replace(sql, "@MotilityNonPro", Left$(txtMotility(1), 5))
80        sql = Replace(sql, "@MotilityNonMotile", Left$(txtMotility(2), 5))
90        sql = Replace(sql, "@MotilitySlow", Left$(txtMotility(4), 5))
100       sql = Replace(sql, "@Consistency", cmbConsistency)
110       sql = Replace(sql, "@Valid", IIf(Valid, 1, 0))
120       sql = Replace(sql, "@SpecimenType", cmbSpecimenType)
130       sql = Replace(sql, "@Motility", Left$(txtMotility(3), 5))
140       sql = Replace(sql, "@UserName", UserName)

150       Cnxn(0).Execute sql

          'sql = "SELECT * from SemenResults WHERE " & _
           '      "SampleID = '" & SampleIDWithOffset & "'"
          'Set tb = New Recordset
          'RecOpenServer 0, tb, sql
          'If tb.EOF Then
          '    tb.AddNew
          '    tb!SampleID = SampleIDWithOffset
          'End If
          'tb!Volume = Left$(cmbVolume, 5)
          'tb!MotilityPro = Left$(txtMotility(0), 5)
          'tb!MotilityNonPro = Left$(txtMotility(1), 5)
          'tb!MotilityNonMotile = Left$(txtMotility(2), 5)
          'tb!Motility = Left$(txtMotility(3), 5)
          'tb!MotilitySlow = Left$(txtMotility(4), 5)
          'tb!Consistency = cmbConsistency
          'tb!SpecimenType = cmbSpecimenType
          'tb!SemenCount = Left$(cmbCount, 50)
          'tb!UserName = UserName
          'If Valid = 1 Then
          '    tb!Valid = 1
          '    UpdatePrintValidLog SampleIDWithOffset, "SEMEN", 1, 0
          'Else
          '    UpdatePrintValidLog SampleIDWithOffset, "SEMEN", 0, 0
          '    tb!Valid = 0
          'End If

          'tb.Update

160       SaveSemenMorphology

170       Exit Sub

SaveSemen_Error:

          Dim strES As String
          Dim intEL As Integer

180       intEL = Erl
190       strES = Err.Description
200       LogError "frmEditSemen", "SaveSemen", intEL, strES, sql

End Sub

Private Sub lblSemenMorph_Click()

10        If lblSemenMorph = "% Abnormal Forms" Then
20            lblSemenMorph = "% Normal Forms"
30        Else
40            lblSemenMorph = "% Abnormal Forms"
50        End If

60        SaveOptionSetting "SemenMorphologyTitle", lblSemenMorph

70        cmdSave.Enabled = True
80        cmdSaveHold.Enabled = True

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

80        intEL = Erl
90        strES = Err.Description
100       LogError "frmEditSemen", "TimerBar_Timer", intEL, strES

End Sub

Private Sub tRecTime_LostFocus()

10        SetDatesColour Me

End Sub


Private Sub tSampleTime_KeyPress(KeyAscii As Integer)

10        cmdSaveHold.Enabled = True
20        cmdSave.Enabled = True

End Sub

Private Sub taddress_Change(Index As Integer)

10        On Error GoTo taddress_Change_Error

20        SetWardClinGP

30        Exit Sub

taddress_Change_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmEditSemen", "taddress_Change", intEL, strES

End Sub

Private Sub taddress_KeyPress(Index As Integer, KeyAscii As Integer)

10        On Error GoTo taddress_KeyPress_Error

20        cmdSaveHold.Enabled = True
30        cmdSave.Enabled = True

40        Exit Sub

taddress_KeyPress_Error:

          Dim strES As String
          Dim intEL As Integer



50        intEL = Erl
60        strES = Err.Description
70        LogError "frmEditSemen", "taddress_KeyPress", intEL, strES


End Sub

Private Sub taddress_LostFocus(Index As Integer)

10        On Error GoTo taddress_LostFocus_Error

20        taddress(Index) = initial2upper(taddress(Index))

30        Exit Sub

taddress_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmEditSemen", "taddress_LostFocus", intEL, strES


End Sub

Private Sub tSampleTime_LostFocus()

10        SetDatesColour Me

End Sub

Private Sub txtAandE_LostFocus()

10        If UCase(HospName(0)) = "MULLINGAR" Then
20            LoadPatientFromAandE Me, True
30        End If


40        If Trim(txtName) = "" Then
50            LoadDemo txtAandE
60        End If

70        txtAandE = UCase(txtAandE)

80        cmdSave.Enabled = True
90        cmdSaveHold.Enabled = True

End Sub


Private Sub LoadDemo(ByVal IDNumber As String)

          Dim tb As New Recordset
          Dim sql As String
          Dim IDType As String
          Dim n As Long

10        IDType = CheckDemographics(IDNumber)
20        If IDType = "" Then
              'clearpatient
30            Exit Sub
40        End If

          'Rem Code Change 16/01/2006
50        sql = "SELECT * from patientifs WHERE " & _
                IDType & " = '" & AddTicks(IDNumber) & "' "

60        Set tb = New Recordset
70        RecOpenServer 0, tb, sql
80        If tb.EOF = True Then
              '   clearpatient
90        Else
100           If Trim(tb!Chart & "") = "" Then txtChart = tb!Mrn & "" Else txtChart = tb!Chart & ""
110           txtAandE = tb!AandE & ""
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
230           txtAge = CalcAge(tb!Dob & "", dtSampleDate)
240           Select Case tb!sex & ""
              Case "M": txtSex = "Male"
250           Case "F": txtSex = "Female"
260           Case Else: txtSex = ""
270           End Select
280           n = InStr(tb!Address0 & "", "''")
290           If n <> 0 Then
300               tb!Address0 = Left$(tb!Address0, n) & Mid$(tb!Address0, n + 2)
310               tb.Update
320           End If

330           taddress(0) = initial2upper(Trim(tb!Address0 & ""))
340           taddress(1) = initial2upper(Trim(tb!Address1 & ""))
350           cmbWard.Text = initial2upper(tb!Ward & "")
360           cmbClinician.Text = initial2upper(tb!Clinician & "")
370       End If
380       tb.Close
End Sub

Private Sub txtage_Change()

10        lblAge = txtAge

End Sub

Private Sub txtAge_KeyPress(KeyAscii As Integer)

10        On Error GoTo txtAge_KeyPress_Error

20        cmdSaveHold.Enabled = True
30        cmdSave.Enabled = True

40        Exit Sub

txtAge_KeyPress_Error:

          Dim strES As String
          Dim intEL As Integer



50        intEL = Erl
60        strES = Err.Description
70        LogError "frmEditSemen", "txtAge_KeyPress", intEL, strES


End Sub

Private Sub txtchart_Change()

10        lblChart = txtChart

End Sub

Private Sub txtChart_KeyPress(KeyAscii As Integer)

10        On Error GoTo txtChart_KeyPress_Error

20        cmdSaveHold.Enabled = True
30        cmdSave.Enabled = True

40        Exit Sub

txtChart_KeyPress_Error:

          Dim strES As String
          Dim intEL As Integer



50        intEL = Erl
60        strES = Err.Description
70        LogError "frmEditSemen", "txtChart_KeyPress", intEL, strES


End Sub

Private Sub txtchart_LostFocus()

10        On Error GoTo txtchart_LostFocus_Error

20        If Trim$(txtChart) = "" Then Exit Sub
30        If Trim$(txtName) <> "" Then Exit Sub

40        LoadPatientFromChart Me, mNewRecord

50        Exit Sub

txtchart_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer



60        intEL = Erl
70        strES = Err.Description
80        LogError "frmEditSemen", "txtchart_LostFocus", intEL, strES


End Sub

Private Sub txtDemographicComment_KeyPress(KeyAscii As Integer)

10        On Error GoTo txtDemographicComment_KeyPress_Error

20        cmdSaveHold.Enabled = True
30        cmdSave.Enabled = True

40        Exit Sub

txtDemographicComment_KeyPress_Error:

          Dim strES As String
          Dim intEL As Integer



50        intEL = Erl
60        strES = Err.Description
70        LogError "frmEditSemen", "txtDemographicComment_KeyPress", intEL, strES


End Sub

Private Sub txtDemographicComment_KeyDown(KeyCode As Integer, Shift As Integer)

          Dim s As Variant
          Dim n As Long
          Dim sql As String
          Dim tb As New Recordset

10        On Error GoTo txtDemographicComment_KeyDown_Error

20        If KeyCode = 113 Then

30            n = txtDemographicComment.SelStart

40            s = Mid(txtDemographicComment, n - 1, 2)

50            If ListText("DE", s) <> "" Then
60                s = ListText("DE", s)
70            End If

80            txtDemographicComment = Left(txtDemographicComment, n - 2)
90            txtDemographicComment = txtDemographicComment & s

100           txtDemographicComment.SelStart = Len(txtDemographicComment)

110       ElseIf KeyCode = 114 Then

120           sql = "SELECT * from lists WHERE listtype = 'DE'"
130           Set tb = New Recordset
140           RecOpenServer 0, tb, sql
150           Do While Not tb.EOF
160               s = Trim(tb!Text)
170               frmMessages.lstComm.AddItem s
180               tb.MoveNext
190           Loop

200           frmMessages.f = Me
210           frmMessages.T = txtDemographicComment
220           frmMessages.Show 1

230           iMsg s

240       End If

250       Exit Sub

txtDemographicComment_KeyDown_Error:

          Dim strES As String
          Dim intEL As Integer

260       intEL = Erl
270       strES = Err.Description
280       LogError "frmEditSemen", "txtDemographicComment_KeyDown", intEL, strES

End Sub
Private Sub txtDoB_Change()

10        lblDoB = txtDoB

End Sub

Private Sub txtDoB_KeyPress(KeyAscii As Integer)

10        On Error GoTo txtDoB_KeyPress_Error

20        cmdSaveHold.Enabled = True
30        cmdSave.Enabled = True

40        Exit Sub

txtDoB_KeyPress_Error:

          Dim strES As String
          Dim intEL As Integer



50        intEL = Erl
60        strES = Err.Description
70        LogError "frmEditSemen", "txtDoB_KeyPress", intEL, strES


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

Private Sub txtMorphology_KeyUp(KeyCode As Integer, Shift As Integer)

10        cmdSave.Enabled = True
20        cmdSaveHold.Enabled = True

End Sub


Private Sub txtMotility_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

10        cmdSave.Enabled = True
20        cmdSaveHold.Enabled = True

End Sub


Private Sub txtName_Change()

10        lblName = txtName

End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)

10        On Error GoTo txtName_KeyPress_Error

20        cmdSaveHold.Enabled = True
30        cmdSave.Enabled = True

40        Exit Sub

txtName_KeyPress_Error:

          Dim strES As String
          Dim intEL As Integer



50        intEL = Erl
60        strES = Err.Description
70        LogError "frmEditSemen", "txtName_KeyPress", intEL, strES


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
100       LogError "frmEditSemen", "txtname_LostFocus", intEL, strES


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
60        LogError "frmEditSemen", "txtSampleID_KeyPress", intEL, strES

End Sub

Private Sub txtSampleID_LostFocus()

10        On Error GoTo txtSampleID_LostFocus_Error

20        If Val(txtSampleID) < 1 Or Trim$(txtSampleID) = "" Or Val(txtSampleID) > (2 ^ 31) - 1 Then
30            txtSampleID = ""
40            txtSampleID.SetFocus
50            Exit Sub
60        End If

70        txtSampleID = Format$(Val(txtSampleID))
80        txtSampleID = Int(txtSampleID)
90        GetSampleIDWithOffset

100       LoadAllDetails

110       cmdSaveHold.Enabled = False
120       cmdSave.Enabled = False

130       Exit Sub

txtSampleID_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer



140       intEL = Erl
150       strES = Err.Description
160       LogError "frmEditSemen", "txtSampleID_LostFocus", intEL, strES


End Sub

Private Sub txtSemenComment_KeyDown(KeyCode As Integer, Shift As Integer)

          Dim sql As String
          Dim tb As New Recordset
          Dim s As String
          Dim n As Integer
          Dim Replacement As String

10        On Error GoTo txtSemenComment_KeyDown_Error

20        If KeyCode = vbKeyF2 Then

30            s = GetLastWord(txtSemenComment)
40            n = Len(s)
50            If n >= 2 Then
60                Replacement = ListText("SE", s)
70                If Replacement <> "" Then
80                    txtSemenComment = Left(txtSemenComment, Len(txtSemenComment) - n) & Replacement
90                End If
100           End If

110       ElseIf KeyCode = vbKeyF3 Then

120           sql = "SELECT * from lists WHERE listtype = 'SE'"
130           Set tb = New Recordset
140           RecOpenServer 0, tb, sql
150           Do While Not tb.EOF
160               s = Trim(tb!Text)
170               frmMessages.lstComm.AddItem s
180               tb.MoveNext
190           Loop

200           Set frmMessages.f = Me
210           Set frmMessages.T = txtSemenComment
220           frmMessages.Show 1

230       End If

240       Exit Sub

txtSemenComment_KeyDown_Error:

          Dim strES As String
          Dim intEL As Integer

250       intEL = Erl
260       strES = Err.Description
270       LogError "frmEditSemen", "txtSemenComment_KeyDown", intEL, strES, sql

End Sub

Private Sub txtSemenComment_KeyPress(KeyAscii As Integer)

10        On Error GoTo txtSemenComment_KeyPress_Error

20        cmdSave.Enabled = True
30        cmdSaveHold.Enabled = True

40        Exit Sub

txtSemenComment_KeyPress_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmEditSemen", "txtSemenComment_KeyPress", intEL, strES

End Sub



Private Sub txtSex_Change()

10        lblSex = txtSex

End Sub

Private Sub txtsex_Click()

10        On Error GoTo txtsex_Click_Error

20        Select Case Trim$(txtSex)
          Case "": txtSex = "Male"
30        Case "Male": txtSex = "Female"
40        Case "Female": txtSex = ""
50        Case Else: txtSex = ""
60        End Select

70        cmdSaveHold.Enabled = True
80        cmdSave.Enabled = True

90        Exit Sub

txtsex_Click_Error:

          Dim strES As String
          Dim intEL As Integer



100       intEL = Erl
110       strES = Err.Description
120       LogError "frmEditSemen", "txtsex_Click", intEL, strES


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
70        LogError "frmEditSemen", "txtsex_KeyPress", intEL, strES


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
60        LogError "frmEditSemen", "txtSex_LostFocus", intEL, strES


End Sub




Private Sub udMotility_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
10        If txtMotility(Index).Enabled = True Then txtMotility(Index).SetFocus
End Sub

Private Sub udMotility_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

10        cmdSave.Enabled = True
20        cmdSaveHold.Enabled = True

End Sub

Private Sub UpDown1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        On Error GoTo UpDown1_MouseUp_Error

20        pBar = 0

30        GetSampleIDWithOffset

40        LoadAllDetails

50        cmdSaveHold.Enabled = False
60        cmdSave.Enabled = False

70        Exit Sub

UpDown1_MouseUp_Error:

          Dim strES As String
          Dim intEL As Integer



80        intEL = Erl
90        strES = Err.Description
100       LogError "frmEditSemen", "UpDown1_MouseUp", intEL, strES


End Sub
