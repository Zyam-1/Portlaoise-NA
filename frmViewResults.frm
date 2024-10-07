VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmViewResults 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire"
   ClientHeight    =   10695
   ClientLeft      =   15
   ClientTop       =   315
   ClientWidth     =   14820
   Icon            =   "frmViewResults.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmViewResults.frx":030A
   ScaleHeight     =   10695
   ScaleWidth      =   14820
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdExcel 
      Caption         =   "Export"
      Enabled         =   0   'False
      Height          =   885
      Index           =   5
      Left            =   12375
      Picture         =   "frmViewResults.frx":280C
      Style           =   1  'Graphical
      TabIndex        =   142
      Top             =   9480
      Width           =   615
   End
   Begin VB.CommandButton cmdExcel 
      Caption         =   "Export"
      Enabled         =   0   'False
      Height          =   885
      Index           =   3
      Left            =   8520
      Picture         =   "frmViewResults.frx":842A
      Style           =   1  'Graphical
      TabIndex        =   141
      Top             =   9480
      Width           =   615
   End
   Begin VB.CommandButton cmdExcel 
      Caption         =   "Export"
      Enabled         =   0   'False
      Height          =   885
      Index           =   4
      Left            =   12495
      Picture         =   "frmViewResults.frx":E048
      Style           =   1  'Graphical
      TabIndex        =   140
      Top             =   4665
      Width           =   615
   End
   Begin VB.CommandButton cmdExcel 
      Caption         =   "Export"
      Enabled         =   0   'False
      Height          =   885
      Index           =   2
      Left            =   8670
      Picture         =   "frmViewResults.frx":13C66
      Style           =   1  'Graphical
      TabIndex        =   139
      Top             =   4665
      Width           =   615
   End
   Begin VB.CommandButton cmdExcel 
      Caption         =   "Export"
      Enabled         =   0   'False
      Height          =   885
      Index           =   0
      Left            =   1770
      Picture         =   "frmViewResults.frx":19884
      Style           =   1  'Graphical
      TabIndex        =   137
      Top             =   9480
      Width           =   615
   End
   Begin VB.TextBox lblEndComment 
      Height          =   1275
      Left            =   6840
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   127
      Top             =   8055
      Width           =   3705
   End
   Begin VB.TextBox lblCoagComment 
      Height          =   915
      Left            =   6795
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   126
      Top             =   3645
      Width           =   3795
   End
   Begin VB.TextBox lblImmComment 
      Height          =   1275
      Left            =   10665
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   125
      Top             =   8055
      Width           =   3660
   End
   Begin VB.TextBox lblBgaComment 
      Height          =   825
      Left            =   10800
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   124
      Top             =   3690
      Width           =   3480
   End
   Begin VB.TextBox lblHaemComment 
      Height          =   1275
      Left            =   3285
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   123
      Top             =   8100
      Width           =   3390
   End
   Begin VB.TextBox lblBioComment 
      Height          =   1275
      Left            =   90
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   122
      Top             =   8100
      Width           =   3030
   End
   Begin VB.CommandButton cmdFax 
      Appearance      =   0  'Flat
      Caption         =   "&Fax"
      Enabled         =   0   'False
      Height          =   885
      Index           =   5
      Left            =   11760
      Picture         =   "frmViewResults.frx":1F4A2
      Style           =   1  'Graphical
      TabIndex        =   114
      Top             =   9480
      Width           =   615
   End
   Begin VB.CommandButton cmdCum 
      Caption         =   "Immunology Cumulative"
      Height          =   885
      Index           =   5
      Left            =   10725
      Picture         =   "frmViewResults.frx":1F7AC
      Style           =   1  'Graphical
      TabIndex        =   113
      Top             =   9480
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdPrint 
      Appearance      =   0  'Flat
      Caption         =   "&Print"
      Enabled         =   0   'False
      Height          =   885
      Index           =   5
      Left            =   12990
      Picture         =   "frmViewResults.frx":1FAB6
      Style           =   1  'Graphical
      TabIndex        =   112
      Top             =   9480
      Width           =   615
   End
   Begin VB.CommandButton cmdFax 
      Appearance      =   0  'Flat
      Caption         =   "&Fax"
      Enabled         =   0   'False
      Height          =   885
      Index           =   4
      Left            =   11880
      Picture         =   "frmViewResults.frx":1FDC0
      Style           =   1  'Graphical
      TabIndex        =   107
      Top             =   4665
      Width           =   615
   End
   Begin VB.CommandButton cmdCum 
      Caption         =   "Blood Gas Cumulative"
      Height          =   885
      Index           =   4
      Left            =   10785
      Picture         =   "frmViewResults.frx":200CA
      Style           =   1  'Graphical
      TabIndex        =   106
      Top             =   4665
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdPrint 
      Appearance      =   0  'Flat
      Caption         =   "&Print"
      Enabled         =   0   'False
      Height          =   885
      Index           =   4
      Left            =   13110
      Picture         =   "frmViewResults.frx":203D4
      Style           =   1  'Graphical
      TabIndex        =   105
      Top             =   4665
      Width           =   615
   End
   Begin VB.CommandButton cmdPrint 
      Appearance      =   0  'Flat
      Caption         =   "&Print"
      Enabled         =   0   'False
      Height          =   885
      Index           =   3
      Left            =   9135
      Picture         =   "frmViewResults.frx":206DE
      Style           =   1  'Graphical
      TabIndex        =   100
      Top             =   9480
      Width           =   615
   End
   Begin VB.CommandButton cmdCum 
      Caption         =   "Endocinology Cumulative"
      Height          =   885
      Index           =   3
      Left            =   6810
      Picture         =   "frmViewResults.frx":209E8
      Style           =   1  'Graphical
      TabIndex        =   99
      Top             =   9480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdFax 
      Appearance      =   0  'Flat
      Caption         =   "&Fax"
      Enabled         =   0   'False
      Height          =   885
      Index           =   3
      Left            =   7905
      Picture         =   "frmViewResults.frx":20CF2
      Style           =   1  'Graphical
      TabIndex        =   98
      Top             =   9480
      Width           =   615
   End
   Begin VB.CommandButton cmdFax 
      Appearance      =   0  'Flat
      Caption         =   "&Fax"
      Enabled         =   0   'False
      Height          =   885
      Index           =   2
      Left            =   8055
      Picture         =   "frmViewResults.frx":20FFC
      Style           =   1  'Graphical
      TabIndex        =   76
      Top             =   4665
      Width           =   615
   End
   Begin VB.CommandButton cmdFax 
      Appearance      =   0  'Flat
      Caption         =   "&Fax"
      Enabled         =   0   'False
      Height          =   885
      Index           =   1
      Left            =   4305
      Picture         =   "frmViewResults.frx":21306
      Style           =   1  'Graphical
      TabIndex        =   75
      Top             =   9480
      Width           =   615
   End
   Begin VB.CommandButton cmdFax 
      Appearance      =   0  'Flat
      Caption         =   "&Fax"
      Enabled         =   0   'False
      Height          =   885
      Index           =   0
      Left            =   1155
      Picture         =   "frmViewResults.frx":21610
      Style           =   1  'Graphical
      TabIndex        =   74
      Top             =   9480
      Width           =   615
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   885
      Left            =   13995
      Picture         =   "frmViewResults.frx":2191A
      Style           =   1  'Graphical
      TabIndex        =   73
      Top             =   9480
      Width           =   750
   End
   Begin VB.CommandButton cmdCum 
      Caption         =   "Biochemistry Cumulative"
      Height          =   885
      Index           =   0
      Left            =   120
      Picture         =   "frmViewResults.frx":21C24
      Style           =   1  'Graphical
      TabIndex        =   72
      Top             =   9480
      Width           =   1035
   End
   Begin VB.CommandButton cmdCum 
      Caption         =   "Haematology Cumulative"
      Height          =   885
      Index           =   1
      Left            =   3270
      Picture         =   "frmViewResults.frx":21F2E
      Style           =   1  'Graphical
      TabIndex        =   71
      Top             =   9480
      Width           =   1035
   End
   Begin VB.CommandButton cmdCum 
      Caption         =   "Coagulation Cumulative"
      Height          =   885
      Index           =   2
      Left            =   6810
      Picture         =   "frmViewResults.frx":22238
      Style           =   1  'Graphical
      TabIndex        =   70
      Top             =   4665
      Width           =   1245
   End
   Begin VB.CommandButton cmdPrint 
      Appearance      =   0  'Flat
      Caption         =   "&Print"
      Enabled         =   0   'False
      Height          =   885
      Index           =   0
      Left            =   2385
      Picture         =   "frmViewResults.frx":22542
      Style           =   1  'Graphical
      TabIndex        =   69
      Top             =   9480
      Width           =   615
   End
   Begin VB.CommandButton cmdPrint 
      Appearance      =   0  'Flat
      Caption         =   "&Print"
      Enabled         =   0   'False
      Height          =   885
      Index           =   1
      Left            =   4920
      Picture         =   "frmViewResults.frx":2284C
      Style           =   1  'Graphical
      TabIndex        =   68
      Top             =   9480
      Width           =   615
   End
   Begin VB.CommandButton cmdPrint 
      Appearance      =   0  'Flat
      Caption         =   "&Print"
      Enabled         =   0   'False
      Height          =   885
      Index           =   2
      Left            =   9285
      Picture         =   "frmViewResults.frx":22B56
      Style           =   1  'Graphical
      TabIndex        =   67
      Top             =   4665
      Width           =   615
   End
   Begin VB.Frame Frame2 
      Caption         =   "Haematology"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   6345
      Left            =   3240
      TabIndex        =   18
      Top             =   1440
      Width           =   3435
      Begin VB.TextBox tLucP 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2520
         MaxLength       =   5
         TabIndex        =   77
         Top             =   3390
         Width           =   825
      End
      Begin VB.TextBox tBasP 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2520
         MaxLength       =   5
         TabIndex        =   26
         Top             =   3120
         Width           =   825
      End
      Begin VB.TextBox tEosP 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2520
         MaxLength       =   5
         TabIndex        =   20
         Top             =   2850
         Width           =   825
      End
      Begin VB.TextBox tMonoP 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2520
         MaxLength       =   5
         TabIndex        =   25
         Top             =   2580
         Width           =   825
      End
      Begin VB.TextBox tLymP 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2520
         MaxLength       =   5
         TabIndex        =   29
         Top             =   2310
         Width           =   825
      End
      Begin VB.TextBox tNeutP 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2520
         MaxLength       =   5
         TabIndex        =   24
         Top             =   2040
         Width           =   825
      End
      Begin VB.TextBox tLucA 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   570
         TabIndex        =   78
         Top             =   3420
         Width           =   825
      End
      Begin VB.TextBox tBasA 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   570
         TabIndex        =   19
         Top             =   3120
         Width           =   825
      End
      Begin VB.TextBox tEosA 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   570
         MaxLength       =   5
         TabIndex        =   23
         Top             =   2850
         Width           =   825
      End
      Begin VB.TextBox tMonoA 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   570
         MaxLength       =   5
         TabIndex        =   28
         Top             =   2580
         Width           =   825
      End
      Begin VB.TextBox tLymA 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   570
         MaxLength       =   5
         TabIndex        =   27
         Top             =   2310
         Width           =   825
      End
      Begin VB.TextBox tNeutA 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   570
         MaxLength       =   5
         TabIndex        =   21
         Top             =   2040
         Width           =   825
      End
      Begin VB.TextBox tMalaria 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2520
         TabIndex        =   134
         Top             =   5010
         Width           =   825
      End
      Begin VB.TextBox tSickledex 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2520
         TabIndex        =   133
         Top             =   5310
         Width           =   825
      End
      Begin VB.TextBox tHypo 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2520
         TabIndex        =   131
         Top             =   4710
         Width           =   825
      End
      Begin VB.TextBox tNrbc 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2520
         MaxLength       =   5
         TabIndex        =   97
         Top             =   1725
         Width           =   825
      End
      Begin VB.TextBox tHDW 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2520
         MaxLength       =   5
         TabIndex        =   90
         Top             =   1410
         Width           =   825
      End
      Begin VB.TextBox tMpXi 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   570
         TabIndex        =   81
         Top             =   4965
         Width           =   825
      End
      Begin VB.TextBox tLI 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2520
         TabIndex        =   80
         Top             =   4410
         Width           =   825
      End
      Begin VB.TextBox tPlt 
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
         Left            =   570
         MaxLength       =   5
         TabIndex        =   41
         Top             =   4230
         Width           =   825
      End
      Begin VB.TextBox tPdw 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2520
         TabIndex        =   40
         Top             =   4110
         Width           =   825
      End
      Begin VB.TextBox tMPV 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   570
         TabIndex        =   39
         Top             =   4665
         Width           =   825
      End
      Begin VB.TextBox tPLCR 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2520
         TabIndex        =   38
         Top             =   3810
         Width           =   825
      End
      Begin VB.TextBox tMCV 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   570
         MaxLength       =   5
         TabIndex        =   37
         Top             =   1005
         Width           =   825
      End
      Begin VB.TextBox tRDWSD 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2520
         MaxLength       =   5
         TabIndex        =   36
         Top             =   555
         Width           =   825
      End
      Begin VB.TextBox tRDWCV 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2520
         MaxLength       =   5
         TabIndex        =   35
         Top             =   270
         Width           =   825
      End
      Begin VB.TextBox tMCH 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2520
         MaxLength       =   5
         TabIndex        =   34
         Top             =   840
         Width           =   825
      End
      Begin VB.TextBox tHct 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   570
         MaxLength       =   5
         TabIndex        =   33
         Top             =   1305
         Width           =   825
      End
      Begin VB.TextBox tHgb 
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
         Left            =   570
         MaxLength       =   5
         TabIndex        =   32
         Top             =   585
         Width           =   825
      End
      Begin VB.TextBox tRBC 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   570
         MaxLength       =   5
         TabIndex        =   31
         Top             =   270
         Width           =   825
      End
      Begin VB.TextBox tMCHC 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2520
         MaxLength       =   5
         TabIndex        =   30
         Top             =   1125
         Width           =   825
      End
      Begin VB.TextBox tWBC 
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
         Left            =   570
         MaxLength       =   5
         TabIndex        =   22
         Top             =   1590
         Width           =   825
      End
      Begin VB.Label Label36 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "PLCR"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1732
         TabIndex        =   136
         Top             =   3855
         Width           =   420
      End
      Begin VB.Label Label35 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "PDW"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1747
         TabIndex        =   135
         Top             =   4155
         Width           =   390
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "% Hypo"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1710
         TabIndex        =   132
         Top             =   4770
         Width           =   810
      End
      Begin VB.Label Label32 
         Caption         =   "%"
         Height          =   240
         Left            =   1575
         TabIndex        =   129
         Top             =   5715
         Width           =   150
      End
      Begin VB.Label lRetP 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1080
         TabIndex        =   128
         Top             =   5685
         Width           =   465
      End
      Begin VB.Label Label28 
         Caption         =   "nRBC"
         Height          =   285
         Left            =   2070
         TabIndex        =   96
         Top             =   1755
         Width           =   555
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "HDW"
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   2
         Left            =   2070
         TabIndex        =   91
         Top             =   1455
         Width           =   405
      End
      Begin VB.Label lMan 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Manual Differential"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   1770
         TabIndex        =   89
         Top             =   60
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.Label lASOT 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   570
         TabIndex        =   87
         Top             =   5985
         Width           =   825
      End
      Begin VB.Label Label25 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "ASOT"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   86
         Top             =   6075
         Width           =   435
      End
      Begin VB.Label lRa 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2520
         TabIndex        =   85
         Top             =   5940
         Width           =   825
      End
      Begin VB.Label Label23 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "RA"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2250
         TabIndex        =   84
         Top             =   5985
         Width           =   225
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "MPXI"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   135
         TabIndex        =   83
         Top             =   5025
         Width           =   390
      End
      Begin VB.Label Label21 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "LI"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2370
         TabIndex        =   82
         Top             =   4470
         Width           =   135
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "#     Luc     %"
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   1440
         TabIndex        =   79
         Top             =   3435
         Width           =   990
      End
      Begin VB.Image imgHaemGraphs 
         Height          =   345
         Left            =   1755
         Picture         =   "frmViewResults.frx":22E60
         Stretch         =   -1  'True
         ToolTipText     =   "Graphs for this Sample"
         Top             =   5940
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "ESR"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   150
         TabIndex        =   65
         Top             =   5445
         Width           =   330
      End
      Begin VB.Label Label17 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Retics"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   90
         TabIndex        =   64
         Top             =   5715
         Width           =   450
      End
      Begin VB.Label Label20 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Monospot"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1800
         TabIndex        =   63
         Top             =   5670
         Width           =   705
      End
      Begin VB.Label lesr 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   570
         TabIndex        =   62
         Top             =   5385
         Width           =   825
      End
      Begin VB.Label lmonospot 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2520
         TabIndex        =   61
         Top             =   5625
         Width           =   810
      End
      Begin VB.Label lretics 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   570
         TabIndex        =   60
         Top             =   5685
         Width           =   465
      End
      Begin VB.Label Label18 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Sickle Screen"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1515
         TabIndex        =   59
         Top             =   5370
         Width           =   990
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "MPV"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   195
         TabIndex        =   58
         Top             =   4725
         Width           =   345
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Index           =   1
         Left            =   195
         TabIndex        =   57
         Top             =   4320
         Width           =   285
      End
      Begin VB.Label Label26 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Malaria"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1995
         TabIndex        =   56
         Top             =   5070
         Width           =   510
      End
      Begin VB.Label Label19 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "RDW CV"
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   1800
         TabIndex        =   55
         Top             =   315
         Width           =   660
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "MCHC"
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   0
         Left            =   1995
         TabIndex        =   54
         Top             =   1170
         Width           =   465
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "MCH"
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   0
         Left            =   2100
         TabIndex        =   53
         Top             =   885
         Width           =   360
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Hct"
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   52
         Top             =   1350
         Width           =   255
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "MCV"
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   51
         Top             =   1035
         Width           =   345
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "RBC"
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   50
         Top             =   300
         Width           =   330
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "RDW SD"
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   0
         Left            =   1785
         TabIndex        =   49
         Top             =   600
         Width           =   675
      End
      Begin VB.Label Label10 
         Caption         =   "Hgb"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   90
         TabIndex        =   48
         Top             =   630
         Width           =   435
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "#    Neut     %"
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   1440
         TabIndex        =   47
         Top             =   2085
         Width           =   990
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "#   Mono    %"
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   1440
         TabIndex        =   46
         Top             =   2625
         Width           =   990
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "#   Lymph   %"
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   1440
         TabIndex        =   45
         Top             =   2355
         Width           =   990
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000007&
         Height          =   240
         Index           =   0
         Left            =   45
         TabIndex        =   44
         Top             =   1695
         Width           =   525
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "#     Eos     %"
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   1440
         TabIndex        =   43
         Top             =   2895
         Width           =   990
      End
      Begin VB.Label Label34 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "#     Bas     %"
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   1440
         TabIndex        =   42
         Top             =   3165
         Width           =   990
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   14400
      Top             =   1980
   End
   Begin VB.Frame Frame1 
      Height          =   1350
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14775
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         Caption         =   "GP"
         Height          =   195
         Left            =   11040
         TabIndex        =   144
         Top             =   660
         Width           =   225
      End
      Begin VB.Label lblGP 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   11340
         TabIndex        =   143
         Top             =   630
         Width           =   1455
      End
      Begin VB.Label lblRunDate 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   11475
         TabIndex        =   121
         Top             =   960
         Width           =   1275
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Run Date"
         Height          =   195
         Index           =   0
         Left            =   10710
         TabIndex        =   120
         Top             =   990
         Width           =   690
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Time Taken"
         Height          =   195
         Left            =   8460
         TabIndex        =   119
         Top             =   990
         Width           =   855
      End
      Begin VB.Label lblTimeTaken 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Not Specified"
         Height          =   255
         Left            =   9360
         TabIndex        =   118
         Top             =   960
         Width           =   1290
      End
      Begin VB.Image imgLatest 
         Height          =   285
         Left            =   14175
         Picture         =   "frmViewResults.frx":232A2
         Stretch         =   -1  'True
         ToolTipText     =   "View Most Recent Record"
         Top             =   915
         Width           =   435
      End
      Begin VB.Image imgNext 
         Height          =   285
         Left            =   13725
         Picture         =   "frmViewResults.frx":235AC
         Stretch         =   -1  'True
         ToolTipText     =   "View Next Record"
         Top             =   915
         Width           =   435
      End
      Begin VB.Image imgPrevious 
         Height          =   285
         Left            =   13275
         Picture         =   "frmViewResults.frx":238B6
         Stretch         =   -1  'True
         ToolTipText     =   "View Previous Record"
         Top             =   915
         Width           =   435
      End
      Begin VB.Image imgEarliest 
         Height          =   285
         Left            =   12825
         Picture         =   "frmViewResults.frx":23BC0
         Stretch         =   -1  'True
         ToolTipText     =   "View Earliest Record"
         Top             =   915
         Width           =   435
      End
      Begin VB.Label lblRecordInfo 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Record 8888 of 8888"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   12825
         TabIndex        =   117
         Top             =   645
         Width           =   1890
      End
      Begin VB.Label lblHosp 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   9810
         TabIndex        =   110
         Top             =   630
         Width           =   1155
      End
      Begin VB.Label Label31 
         Caption         =   "Hospital"
         Height          =   255
         Left            =   9180
         TabIndex        =   111
         Top             =   630
         Width           =   705
      End
      Begin VB.Label lblWard 
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   10980
         TabIndex        =   109
         Top             =   210
         Width           =   3735
      End
      Begin VB.Label Label30 
         Caption         =   "Ward"
         Height          =   225
         Left            =   10500
         TabIndex        =   108
         Top             =   270
         Width           =   465
      End
      Begin VB.Label lblDemogComment 
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
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   150
         TabIndex        =   66
         Top             =   960
         Width           =   8265
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Sex"
         Height          =   195
         Index           =   1
         Left            =   2145
         TabIndex        =   14
         Top             =   660
         Width           =   270
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Age"
         Height          =   195
         Index           =   1
         Left            =   9270
         TabIndex        =   13
         Top             =   285
         Width           =   285
      End
      Begin VB.Label lblSex 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2475
         TabIndex        =   12
         Top             =   630
         Width           =   840
      End
      Begin VB.Label lblAge 
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   9585
         TabIndex        =   11
         Top             =   210
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Address"
         Height          =   195
         Index           =   1
         Left            =   3375
         TabIndex        =   10
         Top             =   660
         Width           =   570
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Date of Birth"
         Height          =   195
         Index           =   1
         Left            =   6780
         TabIndex        =   9
         Top             =   285
         Width           =   885
      End
      Begin VB.Label lblChartTitle 
         AutoSize        =   -1  'True
         Caption         =   "Chart"
         Height          =   195
         Left            =   510
         TabIndex        =   8
         Top             =   660
         Width           =   525
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         Height          =   195
         Index           =   1
         Left            =   2325
         TabIndex        =   7
         Top             =   285
         Width           =   420
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Sample ID"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   6
         Top             =   285
         Width           =   735
      End
      Begin VB.Label lblAddress 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4035
         TabIndex        =   5
         Top             =   630
         Width           =   5070
      End
      Begin VB.Label lblDoB 
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   7695
         TabIndex        =   4
         Top             =   210
         Width           =   1230
      End
      Begin VB.Label lblChart 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1065
         TabIndex        =   3
         Top             =   630
         Width           =   1020
      End
      Begin VB.Label lblName 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2775
         TabIndex        =   2
         Top             =   210
         Width           =   3900
      End
      Begin VB.Label lblSampleID 
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
         Height          =   345
         Left            =   975
         TabIndex        =   1
         Top             =   210
         Width           =   1290
      End
      Begin VB.Label lblSampleDate 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   9360
         TabIndex        =   145
         Top             =   960
         Width           =   1290
      End
   End
   Begin MSFlexGridLib.MSFlexGrid gBio 
      Height          =   5925
      Left            =   180
      TabIndex        =   15
      Top             =   1830
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   10451
      _Version        =   393216
      Cols            =   3
      BackColor       =   16777215
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      BackColorBkg    =   12632256
      GridColorFixed  =   12632256
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      Appearance      =   0
      FormatString    =   "<Parameter          |<Result       |V"
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
   Begin MSComctlLib.ProgressBar PBar 
      Height          =   195
      Left            =   120
      TabIndex        =   16
      Top             =   10440
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSFlexGridLib.MSFlexGrid gCoag 
      Height          =   1860
      Left            =   6780
      TabIndex        =   17
      Top             =   1770
      Width           =   3780
      _ExtentX        =   6668
      _ExtentY        =   3281
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      BackColorBkg    =   12632256
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      Appearance      =   0
      FormatString    =   "<Parameter  |<Result  |<Units      |^   |^      "
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
   Begin MSFlexGridLib.MSFlexGrid gDiff 
      Height          =   1815
      Left            =   3540
      TabIndex        =   88
      Top             =   10890
      Width           =   4785
      _ExtentX        =   8440
      _ExtentY        =   3201
      _Version        =   393216
      Cols            =   3
   End
   Begin MSFlexGridLib.MSFlexGrid gImm 
      Height          =   1665
      Left            =   10665
      TabIndex        =   101
      Top             =   6165
      Width           =   3675
      _ExtentX        =   6482
      _ExtentY        =   2937
      _Version        =   393216
      Cols            =   4
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      Appearance      =   0
      FormatString    =   "<Parameter          |<Result       |V |Comm"
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
      Height          =   1875
      Left            =   10800
      TabIndex        =   103
      Top             =   1770
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   3307
      _Version        =   393216
      Cols            =   3
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      BackColorBkg    =   12632256
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      Appearance      =   0
      FormatString    =   "<Parameter          |<Result       |V"
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
   Begin MSFlexGridLib.MSFlexGrid gEnd 
      Height          =   1665
      Left            =   6840
      TabIndex        =   115
      Top             =   6165
      Width           =   3675
      _ExtentX        =   6482
      _ExtentY        =   2937
      _Version        =   393216
      Cols            =   3
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      Appearance      =   0
      FormatString    =   "<Parameter          |<Result       |V"
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
   Begin VB.Label lblExcelInfo 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exporting..."
      Height          =   285
      Left            =   13500
      TabIndex        =   138
      Top             =   1320
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Label lblnopas 
      Height          =   285
      Left            =   9045
      TabIndex        =   130
      Top             =   135
      Width           =   780
   End
   Begin VB.Label Label27 
      Caption         =   "Immunology"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   315
      Index           =   3
      Left            =   10665
      TabIndex        =   116
      Top             =   5850
      Width           =   1785
   End
   Begin VB.Label Label27 
      Caption         =   "Blood Gas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   315
      Index           =   2
      Left            =   10890
      TabIndex        =   104
      Top             =   1440
      Width           =   1785
   End
   Begin VB.Label Label27 
      Caption         =   "Endocrinology"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   315
      Index           =   1
      Left            =   6720
      TabIndex        =   102
      Top             =   5820
      Width           =   1785
   End
   Begin VB.Label lblFasting 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FASTING"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   690
      TabIndex        =   95
      Top             =   7800
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblHaemValid 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "VALID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   4470
      TabIndex        =   94
      Top             =   7800
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label27 
      Caption         =   "Biochemistry"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   315
      Index           =   0
      Left            =   180
      TabIndex        =   93
      Top             =   1440
      Width           =   1785
   End
   Begin VB.Label Label24 
      Caption         =   "Coagulation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   315
      Left            =   6780
      TabIndex        =   92
      Top             =   1410
      Width           =   1785
   End
End
Attribute VB_Name = "frmViewResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Activated As Boolean

Dim Tn As Long
Dim Nrc As Long

Private tbDem As Recordset
Private pPrintToPrinter As String

Private CurrentRecordNumber As Long

Private Sub CheckCumulative()

          Dim tb As New Recordset
          Dim sql As String

10        On Error GoTo CheckCumulative_Error

20        On Error GoTo CheckCumulative_Error

30        sql = "SELECT count(distinct(demographics.sampleid)) as tot from Demographics, bioresults WHERE demographics.sampleid = bioresults.sampleid "
40        If Trim(lblChart) <> "" Then
50            sql = sql & "and ( demographics.Chart = '" & lblChart & "') "
60        End If
70        sql = sql & "and (demographics.PatName = '" & AddTicks(lblName) & "' and demographics.dob = '" & Format(lblDoB, "dd/MMM/yyyy") & "') "

80        Set tb = New Recordset
90        RecOpenServer Tn, tb, sql
100       cmdCum(0).Visible = tb!Tot > 1

110       sql = "SELECT count(distinct(demographics.sampleid)) as tot from Demographics, haemresults WHERE demographics.sampleid = haemresults.sampleid "
120       If Trim(lblChart) <> "" Then
130           sql = sql & "and ( demographics.Chart = '" & lblChart & "') "
140       End If
150       sql = sql & "and (demographics.PatName = '" & AddTicks(lblName) & "' and demographics.dob = '" & Format(lblDoB, "dd/MMM/yyyy") & "') "
160       Set tb = New Recordset
170       RecOpenServer Tn, tb, sql
180       cmdCum(1).Visible = tb!Tot > 1

190       sql = "SELECT count(distinct(demographics.sampleid)) as tot from Demographics, coagresults WHERE demographics.sampleid = coagresults.sampleid "
200       If Trim(lblChart) <> "" Then
210           sql = sql & "and ( demographics.Chart = '" & lblChart & "') "
220       End If
230       sql = sql & "and (demographics.PatName = '" & AddTicks(lblName) & "' and demographics.dob = '" & Format(lblDoB, "dd/MMM/yyyy") & "') "
240       Set tb = New Recordset
250       RecOpenServer Tn, tb, sql
260       cmdCum(2).Visible = tb!Tot > 1

270       If SysOptDeptImm(0) Then
280           sql = "SELECT count(distinct(demographics.sampleid)) as tot from Demographics, immresults WHERE demographics.sampleid = immresults.sampleid "
290           If Trim(lblChart) <> "" Then
300               sql = sql & "and ( demographics.Chart = '" & lblChart & "') "
310           End If
320           sql = sql & "and (demographics.PatName = '" & AddTicks(lblName) & "' and demographics.dob = '" & Format(lblDoB, "dd/MMM/yyyy") & "') "
330           Set tb = New Recordset
340           RecOpenServer Tn, tb, sql
350           cmdCum(5).Visible = tb!Tot > 1
360       End If

370       If SysOptDeptEnd(0) Then
380           sql = "SELECT count(distinct(demographics.sampleid)) as tot from Demographics, endresults WHERE demographics.sampleid = endresults.sampleid "
390           If Trim(lblChart) <> "" Then
400               sql = sql & "and ( demographics.Chart = '" & lblChart & "') "
410           End If
420           sql = sql & "and (demographics.PatName = '" & AddTicks(lblName) & "' and demographics.dob = '" & Format(lblDoB, "dd/MMM/yyyy") & "') "
430           Set tb = New Recordset
440           RecOpenServer Tn, tb, sql
450           cmdCum(3).Visible = tb!Tot > 1
460       End If

470       sql = "SELECT count(distinct(demographics.sampleid)) as tot from Demographics, bgaresults WHERE demographics.sampleid = bgaresults.sampleid "
480       If Trim(lblChart) <> "" Then
490           sql = sql & "and ( demographics.Chart = '" & lblChart & "') "
500       End If
510       sql = sql & "and (demographics.PatName = '" & AddTicks(lblName) & "' and demographics.dob = '" & Format(lblDoB, "dd/MMM/yyyy") & "') "
520       Set tb = New Recordset
530       RecOpenServer Tn, tb, sql
540       cmdCum(4).Visible = tb!Tot > 1





550       Exit Sub

CheckCumulative_Error:

          Dim strES As String
          Dim intEL As Integer


560       intEL = Erl
570       strES = Err.Description
580       LogError "frmViewResults", "CheckCumulative", intEL, strES, sql


End Sub

Private Sub FillDemographics()


10        On Error GoTo FillDemographics_Error

20        lblSampleID = tbDem!SampleID & ""
30        lblName = tbDem!PatName & ""
40        If Not IsNull(tbDem!Dob) Then
50            lblDoB = Format(tbDem!Dob, "dd/MMM/yyyy")
60        Else
70            lblDoB = ""
80        End If
90        lblAge = tbDem!Age & ""
100       If lblAge = "" Then lblAge = CalcOldAge(lblDoB, tbDem!Rundate)
110       lblRundate = Format(tbDem!Rundate, "dd/MMM/yyyy")
120       If IsDate(tbDem!SampleDate) Then
130           If Format(tbDem!SampleDate, "hh:mm") <> "00:00" Then
140               lblTimeTaken = Format(tbDem!SampleDate, "dd/MM/yy hh:mm")
150           Else
160               lblTimeTaken = "Not Specified"
170           End If
180       Else
190           lblTimeTaken = "Not Specified"
200       End If
210       lblWard = tbDem!Ward & ""
220       lblGP = tbDem!GP & ""
230       lblHosp = tbDem!Hospital & ""
240       Select Case Left$(UCase$(tbDem!sex & ""), 1)
          Case "M": lblSex = "Male"
250       Case "F": lblSex = "Female"
260       Case Else: lblSex = ""
270       End Select
280       lblAddress = tbDem!Addr0 & " " & tbDem!Addr1 & ""

290       If tbDem!Fasting = True Then lblFasting.Visible = True Else lblFasting.Visible = False

300       If CurrentRecordNumber = Nrc Then
310           lblRecordInfo = "Most Recent Record."
320       ElseIf CurrentRecordNumber = 1 Then
330           lblRecordInfo = "Earliest Record."
340       Else
350           lblRecordInfo = "Record " & CurrentRecordNumber & " of " & Nrc
360       End If

370       imgNext.Visible = False
380       imgLatest.Visible = False
390       If Nrc = 1 Then
400           imgEarliest.Visible = False
410           imgPrevious.Visible = False
420       Else
430           imgEarliest.Visible = True
440           imgPrevious.Visible = True
450       End If


460       Exit Sub

FillDemographics_Error:

          Dim strES As String
          Dim intEL As Integer


470       intEL = Erl
480       strES = Err.Description
490       LogError "frmViewResults", "FillDemographics", intEL, strES


End Sub

Private Sub LoadAllResults()

10        On Error GoTo LoadAllResults_Error

20        lblSampleID = Format$(Val(lblSampleID))
30        If lblSampleID = "0" Then Exit Sub

40        If SysOptDeptBio(0) Then LoadBiochemistry
50        If SysOptDeptCoag(0) Then LoadCoag
60        If SysOptDeptHaem(0) Then LoadHaem
70        If SysOptDeptImm(0) Then LoadImmunology
80        If SysOptDeptBga(0) Then LoadBloodGas
90        LoadComments
100       If SysOptDeptEnd(0) Then LoadEndocrinology

110       Exit Sub

LoadAllResults_Error:

          Dim strES As String
          Dim intEL As Integer


120       intEL = Erl
130       strES = Err.Description
140       LogError "frmViewResults", "LoadAllResults", intEL, strES


End Sub

Private Sub LoadComments()

          Dim Ob As Observation
          Dim Obs As Observations

10        On Error GoTo LoadComments_Error

20        lblDemogComment = ""
30        lblCoagComment = ""
40        lblBioComment = ""
50        lblHaemComment = ""
60        lblImmComment = ""
70        lblEndComment = ""

80        If Trim$(lblSampleID) = "" Then Exit Sub

90        Set Obs = New Observations
100       Set Obs = Obs.Load(lblSampleID, "Biochemistry", "Demographic", "Haematology", "Coagulation", _
                             "Immunology", "Endocrinology")
110       If Not Obs Is Nothing Then
120           For Each Ob In Obs
130               Select Case UCase$(Ob.Discipline)
                  Case "BIOCHEMISTRY": lblBioComment = Ob.Comment
140               Case "HAEMATOLOGY": lblHaemComment = Ob.Comment
150               Case "DEMOGRAPHIC": lblDemogComment = Ob.Comment
160               Case "COAGULATION": lblCoagComment = Ob.Comment
170               Case "IMMUNOLOGY": lblImmComment = Ob.Comment
180               Case "ENDOCRINOLOGY": lblEndComment = Ob.Comment
190               End Select
200           Next
210       End If

220       Exit Sub

LoadComments_Error:

          Dim strES As String
          Dim intEL As Integer

230       intEL = Erl
240       strES = Err.Description
250       LogError "frmViewResults", "LoadComments", intEL, strES

End Sub

Private Sub LoadInitialDemographics()

          Dim sql As String
          Dim Asql As String
          Dim tsql As String
          Dim tb As Recordset


10        On Error GoTo LoadInitialDemographics_Error

20        sql = "SELECT * from Demographics WHERE "
30        Asql = "SELECT count(sampleid) as tot from Demographics WHERE "

40        If Trim(lblChart) <> "" Then
50            tsql = "Chart = '" & Trim(lblChart) & "' " & _
                     "and PatName = '" & AddTicks(lblName) & "' " & _
                     "and dob = '" & Format(lblDoB, "dd/MMM/yyyy") & "' "
60        Else
70            tsql = "PatName = '" & AddTicks(lblName) & "' " & _
                     "and dob = '" & Format(lblDoB, "dd/MMM/yyyy") & "' "
80        End If
90        Set tbDem = New Recordset
100       RecOpenClient Tn, tbDem, sql & tsql & "Order by SampleDate asc, SampleID asc"
110       Nrc = 0

120       If Not tbDem.EOF Then
130           tbDem.MoveLast
140           Set tb = New Recordset
150           RecOpenServer Tn, tb, Asql & tsql
160           Nrc = tb!Tot

170           CurrentRecordNumber = Nrc
180           Do While Not tbDem.BOF
190               If lblSampleID <> tbDem!SampleID & "" Then
200                   tbDem.MovePrevious
210                   CurrentRecordNumber = CurrentRecordNumber - 1
220               Else
230                   lblName = tbDem!PatName & ""
240                   lblWard = tbDem!Ward & ""
250                   lblGP = tbDem!GP & ""
260                   lblHosp = tbDem!Hospital & ""
270                   If Not IsNull(tbDem!Dob) Then
280                       lblDoB = Format(tbDem!Dob, "dd/MMM/yyyy")
290                   Else
300                       lblDoB = ""
310                   End If
320                   lblAge = tbDem!Age & ""
330                   lblRundate = Format(tbDem!Rundate, "dd/MMM/yyyy")
340                   If lblAge = "" Then lblAge = CalcOldAge(lblDoB, lblRundate)
350                   If IsDate(tbDem!SampleDate) Then
360                       If Format(tbDem!SampleDate, "hh:mm") <> "00:00" Then
370                           lblTimeTaken = Format(tbDem!SampleDate, "dd/MM/yy hh:mm")
                           
380                       Else
390                           lblTimeTaken = "Not Specified"
400                           lblSampleDate = Format(tbDem!SampleDate, "dd/MM/yyyy")
410                       End If
420                   Else
430                       lblTimeTaken = "Not Specified"
440                   End If
450                   Select Case Left$(UCase$(tbDem!sex & ""), 1)
                      Case "M": lblSex = "Male"
460                   Case "F": lblSex = "Female"
470                   Case Else: lblSex = ""
480                   End Select
490                   lblAddress = tbDem!Addr0 & " " & tbDem!Addr1 & ""
500                   If tbDem!Fasting = True Then lblFasting.Visible = True Else lblFasting.Visible = False
510                   Exit Do
520               End If
530           Loop

540           If CurrentRecordNumber = Nrc Then
550               lblRecordInfo = "Most Recent Record."
560               imgNext.Visible = False
570               imgLatest.Visible = False
580           ElseIf CurrentRecordNumber = 1 Then
590               lblRecordInfo = "Earliest Record."
600               imgEarliest.Visible = False
610               imgPrevious.Visible = False
620               imgNext.Visible = True
630               imgLatest.Visible = True
640           Else
650               lblRecordInfo = "Record " & CurrentRecordNumber & " of " & Nrc
660           End If

670           If Nrc = 1 Then
680               imgEarliest.Visible = False
690               imgPrevious.Visible = False
700           Else
710           End If
720       End If




730       Exit Sub

LoadInitialDemographics_Error:

          Dim strES As String
          Dim intEL As Integer


740       intEL = Erl
750       strES = Err.Description
760       LogError "frmViewResults", "LoadInitialDemographics", intEL, strES, sql


End Sub

Private Sub cmdCancel_Click()

10        Unload Me

End Sub

Private Sub cmdCum_Click(Index As Integer)

10        On Error GoTo cmdCum_Click_Error

20        Select Case Index
          Case 0:
30            With frmFullBio
40                .lblChart = lblChart
50                .lblDoB = lblDoB
60                .lblName = lblName
70                .lblSex = lblSex
80                .Show 1
90            End With
100       Case 1:
110           With frmFullHaem
120               .lblChart = lblChart
130               .lblDoB = lblDoB
140               .lblName = lblName
150               .lblSex = lblSex
160               .Tn = Tn
170               .Show 1
180           End With
190       Case 2:
200           With frmFullCoag
210               .lblChart = lblChart
220               .lblDoB = lblDoB
230               .lblName = lblName
240               .Tn = Tn
250               .Show 1
260           End With
270       Case 3:
280           With frmFullEnd
290               .lblChart = lblChart
300               .lblDoB = lblDoB
310               .lblName = lblName
320               .Tn = Tn
330               .Show 1
340           End With
350       Case 4:
360           With frmFullBga
370               .lblChart = lblChart
380               .lblDoB = lblDoB
390               .lblName = lblName
400               .Tn = Tn
410               .Show 1
420           End With
430       Case 5:
440           With frmFullImm
450               .lblSex = lblSex
460               .lblChart = lblChart
470               .lblDoB = lblDoB
480               .lblName = lblName
490               .Tn = Tn
500               .Show 1
510           End With
520       End Select

530       Exit Sub

cmdCum_Click_Error:

          Dim strES As String
          Dim intEL As Integer

540       intEL = Erl
550       strES = Err.Description
560       LogError "frmViewResults", "cmdCum_Click", intEL, strES

End Sub

Private Sub cmdExcel_Click(Index As Integer)

          Dim strHeading As String

10        On Error GoTo cmdExcel_Click_Error

20        strHeading = Choose(Index + 1, "Biochemistry", "", "Coagulation", "Blood Gas", "Endocrinology", "Immunology") & " History" & vbCr
30        strHeading = strHeading & "Patient Name: " & lblName & vbCr
40        strHeading = strHeading & vbCr

50        Select Case Index
          Case 0: ExportFlexGrid gBio, Me, strHeading
60        Case 2: ExportFlexGrid gCoag, Me, strHeading
70        Case 3: ExportFlexGrid gEnd, Me, strHeading
80        Case 4: ExportFlexGrid gBga, Me, strHeading
90        Case 5: ExportFlexGrid gImm, Me, strHeading
100       End Select
110       Exit Sub

cmdExcel_Click_Error:

          Dim strES As String
          Dim intEL As Integer

120       intEL = Erl
130       strES = Err.Description
140       LogError "frmViewResults", "cmdExcel_Click", intEL, strES


End Sub

Private Sub cmdFax_Click(Index As Integer)

          Dim sql As String
          Dim tb As New Recordset
          Dim FaxNumber As String
          Dim strWard As String
          Dim strGp As String
          Dim strClin As String
          Dim Dept As String

10        On Error GoTo cmdFax_Click_Error

20        Dept = ""

30        If lblGP <> "" Then
40            sql = "SELECT * from GPS WHERE text = '" & lblGP & "' and hospitalcode = '" & ListCodeFor("HO", lblHosp) & "'"
50            Set tb = New Recordset
60            RecOpenServer 0, tb, sql
70            If Not tb.EOF Then
80                FaxNumber = Trim$(tb!FAX & "")
90            End If
100       ElseIf lblWard <> "" Then
110           sql = "SELECT * from wards WHERE text = '" & lblWard & "'  and hospitalcode = '" & ListCodeFor("HO", lblHosp) & "'"
120           Set tb = New Recordset
130           RecOpenServer 0, tb, sql
140           If Not tb.EOF Then
150               FaxNumber = tb!FAX
160           End If
170       End If




180       sql = "SELECT * FROM Demographics WHERE " & _
                "SampleID = '" & lblSampleID & "'"

190       Set tb = New Recordset
200       RecOpenServer 0, tb, sql
210       If Not tb.EOF Then
220           strWard = tb!Ward & ""
230           strClin = tb!Clinician & ""
240           strGp = tb!GP & ""
250       End If

260       FaxNumber = iBOX("Faxnumber ", , FaxNumber)

270       If FaxNumber = "" Then
280           iMsg "Fax Cancelled!"
290           Exit Sub
300       End If

310       If SysOptFaxCom(0) Then
320           sql = "SELECT * FROM PrintPending WHERE " & _
                    "Department = 'M' " & _
                    "AND SampleID = '" & lblSampleID & "'"
330           Set tb = New Recordset
340           RecOpenClient Tn, tb, sql
350           If tb.EOF Then
360               tb.AddNew
370           End If
380           tb!SampleID = lblSampleID
390           tb!Department = "M"
400           tb!Initiator = UserName
410           tb!UsePrinter = "FAX"
420           tb!pTime = Now
430           tb!FaxNumber = FaxNumber
440           tb!Ward = strWard
450           tb!Clinician = strClin
460           tb!GP = strGp
470           tb.Update
480       Else
490           Select Case Index
              Case 0
500               sql = "UPDATE BIORESULTS SET PRINTED = 0 WHERE SAMPLEID = " & lblSampleID & ""
510               Cnxn(0).Execute sql
520           Case 2
530               sql = "UPDATE COAGRESULTS SET PRINTED = 0 WHERE SAMPLEID = " & lblSampleID & ""
540               Cnxn(0).Execute sql
550           Case 3
560               sql = "UPDATE ENDRESULTS SET PRINTED = 0 WHERE SAMPLEID = " & lblSampleID & ""
570               Cnxn(0).Execute sql
580           Case 5
590               sql = "UPDATE IMMRESULTS SET PRINTED = 0 WHERE SAMPLEID = " & lblSampleID & ""
600               Cnxn(0).Execute sql
610           End Select
620           Select Case Index
              Case 0: Dept = "B"
630           Case 1: Dept = "H"
640           Case 2: Dept = "C"
650           Case 3: Dept = "E"
660           Case 5: Dept = "J"
670           End Select
680           If Dept = "" Then Exit Sub
690           sql = "SELECT * FROM PrintPending WHERE " & _
                    "Department = '" & Dept & "' " & _
                    "AND SampleID = '" & lblSampleID & "' " & _
                    "AND (FaxNumber <> '' OR FaxNumber IS NOT NULL)"
700           Set tb = New Recordset
710           RecOpenClient Tn, tb, sql
720           If tb.EOF Then
730               tb.AddNew
740           End If
750           tb!SampleID = lblSampleID
760           tb!Department = Dept
770           tb!Initiator = UserName
780           tb!UsePrinter = "Fax"
790           tb!FaxNumber = FaxNumber
800           tb!Ward = strWard
810           tb!Clinician = strClin
820           tb!GP = strGp
830           tb.Update
840       End If

850       cmdPrint(Index).Enabled = False

860       Exit Sub

cmdFax_Click_Error:

          Dim strES As String
          Dim intEL As Integer

870       intEL = Erl
880       strES = Err.Description
890       LogError "frmViewResults", "cmdFax_Click", intEL, strES, sql

End Sub

Private Sub cmdPrint_Click(Index As Integer)

          Dim sql As String
          Dim tb As New Recordset
          Dim strWard As String
          Dim strGp As String
          Dim strClin As String

10        On Error GoTo cmdPrint_Click_Error

20        strWard = ""
30        strGp = ""
40        strClin = ""
50        If Index = 0 Then    'Biochemistry - Reprint
60            sql = "UPDATE BioResults " & _
                    "Set Printed = '0' WHERE " & _
                    "SampleID = '" & lblSampleID & "'"
70            Cnxn(0).Execute sql
80        ElseIf Index = 3 Then
90            sql = "UPDATE EndResults " & _
                    "Set Printed = '0' WHERE " & _
                    "SampleID = '" & lblSampleID & "'"
100           Cnxn(0).Execute sql
110       ElseIf Index = 5 Then
120           sql = "UPDATE immResults " & _
                    "Set Printed = '0' WHERE " & _
                    "SampleID = '" & lblSampleID & "'"
130           Cnxn(0).Execute sql
140       End If

          'If Index = 0 Then    'Biochemistry - Reprint
          '    sql = "UPDATE BioResults " & _
               '          "Set Printed = '0', Valid = 1, Operator = '" & AddTicks(UserCode) & "' WHERE " & _
               '          "SampleID = '" & lblSampleID & "'"
          '    Cnxn(0).Execute sql
          'ElseIf Index = 3 Then
          '    sql = "UPDATE EndResults " & _
               '          "Set Printed = '0', Valid = 1, Operator = '" & AddTicks(UserCode) & "'  WHERE " & _
               '          "SampleID = '" & lblSampleID & "'"
          '    Cnxn(0).Execute sql
          'ElseIf Index = 5 Then
          '    sql = "UPDATE immResults " & _
               '          "Set Printed = '0', Valid = 1, Operator = '" & AddTicks(UserCode) & "'  WHERE " & _
               '          "SampleID = '" & lblSampleID & "'"
          '    Cnxn(0).Execute sql
          'End If
150       sql = "SELECT * FROM Demographics WHERE " & _
                "SampleID = '" & lblSampleID & "'"

160       Set tb = New Recordset
170       RecOpenServer 0, tb, sql
180       If Not tb.EOF Then
190           strWard = tb!Ward & ""
200           strClin = tb!Clinician & ""
210           strGp = tb!GP & ""
220       End If

230       sql = "SELECT * FROM PrintPending WHERE " & _
                "Department = '" & Choose(Index + 1, "B", "H", "C", "E", "G", "I") & "' " & _
                "AND SampleID = '" & lblSampleID & "'"
240       Set tb = New Recordset
250       RecOpenClient 0, tb, sql
260       If tb.EOF Then
270           tb.AddNew
280       End If
290       tb!SampleID = lblSampleID
300       tb!Department = Choose(Index + 1, "B", "H", "C", "E", "G", "I")
310       If SysOptRealImm(0) And tb!Department = "I" Then tb!Department = "J"
320       tb!Initiator = UserName
330       tb!UsePrinter = pPrintToPrinter
340       tb!UseConnection = Tn
350       tb!Ward = strWard
360       tb!Clinician = strClin
370       tb!GP = strGp
380       tb.Update

390       cmdPrint(Index).Enabled = False

400       Exit Sub

cmdPrint_Click_Error:

          Dim strES As String
          Dim intEL As Integer

410       intEL = Erl
420       strES = Err.Description
430       LogError "frmViewResults", "cmdPrint_Click", intEL, strES, sql

End Sub

Private Sub cmdPrint_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)



10        On Error GoTo cmdPrint_MouseDown_Error

20        If Button = 2 Then
30            Set frmForcePrinter.f = frmViewResults
40            frmForcePrinter.Show 1
50        End If




60        Exit Sub

cmdPrint_MouseDown_Error:

          Dim strES As String
          Dim intEL As Integer


70        intEL = Erl
80        strES = Err.Description
90        LogError "frmViewResults", "cmdPrint_MouseDown", intEL, strES


End Sub
Public Property Let PrintToPrinter(ByVal strNewValue As String)
Attribute PrintToPrinter.VB_HelpID = 1355

10        On Error GoTo PrintToPrinter_Error

20        pPrintToPrinter = strNewValue

30        Exit Property

PrintToPrinter_Error:

          Dim strES As String
          Dim intEL As Integer


40        intEL = Erl
50        strES = Err.Description
60        LogError "frmViewResults", "PrintToPrinter", intEL, strES


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
60        LogError "frmViewResults", "PrintToPrinter", intEL, strES


End Property
Private Sub Form_Activate()
          Dim n As Long


10        On Error GoTo Form_Activate_Error

20        pBar.Max = LogOffDelaySecs
30        pBar = 0

40        Me.Refresh

50        Timer1.Enabled = True

60        If Activated Then Exit Sub
70        Activated = True

80        For n = 0 To Cn
90            If HospName(n) = lblHosp Then
100               Tn = n
110           End If
120       Next

130       If TestSys = True Then Me.Caption = Me.Caption & " - TEST SYSTEM"

140       LoadInitialDemographics
150       CheckCumulative
160       LoadAllResults





170       Exit Sub

Form_Activate_Error:

          Dim strES As String
          Dim intEL As Integer


180       intEL = Erl
190       strES = Err.Description
200       LogError "frmViewResults", "Form_Activate", intEL, strES


End Sub
Private Sub LoadHaem()
      'Get Manual Diff to Load over Analyser Diff (FIX)

          Dim tb As New Recordset
          Dim sql As String
          Dim tbd As Recordset
          Dim lym As String
          Dim neut As String
          Dim DiffFound As Boolean
          Dim P As String
          Dim A As String
          Dim w As String
          Dim n As Long
          Dim Plt As String

10        On Error GoTo LoadHaem_Error

20        lblSampleID = Format$(Val(lblSampleID))

30        If lblSampleID = "0" Then Exit Sub
40        LockHaem False

50        ClearHaem
60        imgHaemGraphs.Visible = False

70        cmdPrint(1).Enabled = False
80        cmdFax(1).Enabled = False
90        sql = "SELECT * FROM Haemresults WHERE " & _
                "SampleID = '" & lblSampleID & "'"

100       If UserMemberOf = "Secretarys" Or UCase(UserMemberOf) = "HISTOLOOKUP" Then
110           sql = sql & " AND Valid = '1'"
120       End If
130       Set tb = New Recordset
140       RecOpenServer Tn, tb, sql
150       If Not tb.EOF Then
160           cmdPrint(1).Enabled = True
170           cmdFax(1).Enabled = True
              '  If Not tb!Valid Or IsNull(tb!Valid) Then
              '    lNotValid.Visible = True
              '    imgHaemGraphs.Visible = False
              '  Else
180           If Not IsNull(tb!gwb1) Or Not IsNull(tb!gwb2) Or Not IsNull(tb!gRbc) Or Not IsNull(tb!gplt) Then
190               imgHaemGraphs.Visible = True
200           End If

210           If Not IsNull(tb!rbc) Then
220               Colourise "RBC", tRBC, Trim(tb!rbc & ""), lblSex, lblDoB, lblSampleDate
230           End If

240           If Not IsNull(tb!Hgb) Then
250               Colourise "Hgb", tHgb, Trim(tb!Hgb & ""), lblSex, lblDoB, lblSampleDate
260           End If

270           If Not IsNull(tb!MCV) Then
280               Colourise "MCV", tMCV, Trim(tb!MCV & ""), lblSex, lblDoB, lblSampleDate
290           End If

300           If Not IsNull(tb!hct) Then
310               Colourise "Hct", tHct, Trim(tb!hct & ""), lblSex, lblDoB, lblSampleDate
320           End If

330           If Not IsNull(tb!RDWCV) And Val(tb!RDWCV & "") <> 0 Then
340               Colourise "RDWCV", tRDWCV, Trim(tb!RDWCV & ""), lblSex, lblDoB, lblSampleDate
350           End If

360           If Not IsNull(tb!rdwsd) And Val(tb!rdwsd & "") <> 0 Then
370               Colourise "RDWSD", tRDWSD, Trim(tb!rdwsd & ""), lblSex, lblDoB, lblSampleDate
380           End If

390           If Not IsNull(tb!mch) Then
400               Colourise "MCH", tMCH, Trim(tb!mch & ""), lblSex, lblDoB, lblSampleDate
410           End If

420           If Not IsNull(tb!mchc) Then
430               Colourise "MCHC", tMCHC, Trim(tb!mchc & ""), lblSex, lblDoB, lblSampleDate
440           End If

450           If Not IsNull(tb!HDW) Then
460               Colourise "HDW", tHDW, Trim(tb!HDW & ""), lblSex, lblDoB, lblSampleDate
470           End If

480           tHypo = Trim(tb!hyp & "")
490           If Not IsNull(tb!Plt) Then
500               Plt = Trim(tb!Plt)
                  '      If InStr(Plt, ">") Then
                  '        x = InStr(Plt, ">")
                  '        Plt = Mid(Plt, x + 1)
                  '      End If
510               Colourise "plt", tPlt, Trim(Plt & ""), lblSex, lblDoB, lblSampleDate
520           End If

530           If Not IsNull(tb!mpv) Then
540               Colourise "MPV", tMPV, Trim(tb!mpv & ""), lblSex, lblDoB, lblSampleDate
550           End If

560           If Not IsNull(tb!plcr) Then
570               Colourise "PLCR", tPLCR, Trim(tb!plcr & ""), lblSex, lblDoB, lblSampleDate
580           End If

590           If Not IsNull(tb!pdw) Then
600               Colourise "Pdw", tPdw, Trim(tb!pdw & ""), lblSex, lblDoB, lblSampleDate
610           End If

620           If Not IsNull(tb!wbc) Then
630               Colourise "WBC", tWBC, Trim(tb!wbc & ""), lblSex, lblDoB, lblSampleDate
640           End If


650           tMalaria = tb!Malaria & ""
660           tSickledex = tb!Sickledex & ""
670           lesr = tb!esr & ""
680           lretics = tb!reta & ""
690           lRetP = tb!RetP & ""
700           If tb!Monospot & "" <> "" Then
710               If tb!Monospot & "" = "N" Then
720                   lmonospot = "Negative"
730               ElseIf tb!Monospot = "P" Then
740                   lmonospot = "Positive"
750               ElseIf tb!Monospot = "I" Then
760                   lmonospot = "Inconclusive"
770               End If
780           End If
790           If SysOptHaemAn1(0) = "ADVIA" Then
800               tLI = Trim(tb!Li & "")
810               tMpXi = Trim(tb!mpxi & "")
820               tNrbc = Trim(tb!nrbcp & "")
830           End If
840           lASOT = Trim(tb!tASOt & "")
850           lRa = Trim(tb!tRa & "")
860           lblHaemValid.Visible = True
870           If tb!Valid = 1 Then
880               lblHaemValid = "VALID"
890           Else
900               lblHaemValid = "NOT VALID"
910           End If
              'diff
920           sql = "SELECT * from differentials WHERE runnumber = '" & lblSampleID & "' and prndiff = 1  "
930           Set tbd = New Recordset
940           RecOpenServer Tn, tbd, sql
950           If Not tbd.EOF Then
960               DiffFound = True
970               For n = 0 To 14
980                   w = Trim(tbd("Wording" & Format(n)) & "")
990                   P = IIf(Val(tbd("P" & Format(n)) & "") = Tn, "", tbd("P" & Format(n)))
1000                  A = IIf(Val(tbd("A" & Format(n)) & "") = Tn, "", tbd("A" & Format(n)))
1010                  gDiff.AddItem w & vbTab & P & vbTab & A
1020              Next
1030          End If

1040          lMan.Visible = DiffFound

1050          If DiffFound = True Then
1060              For n = 1 To gDiff.Rows - 1
1070                  If InStr(UCase(gDiff.TextMatrix(n, 0)), "LYMP") > 0 Then
1080                      neut = gDiff.TextMatrix(n, 2)
1090                      lym = gDiff.TextMatrix(n, 1)
1100                      Exit For
1110                  End If
1120              Next
1130          Else
1140              neut = Trim(tb!LymA & "")
1150              lym = Trim(tb!LymP & "")
1160          End If

1170          If Not IsNull(neut) Then
1180              Colourise "LymA", tLymA, neut, lblSex, lblDoB, lblSampleDate
1190          End If

1200          If Not IsNull(lym) Then
1210              Colourise "LymP", tLymP, lym, lblSex, lblDoB, lblSampleDate
1220          End If

1230          If DiffFound = True Then
1240              For n = 1 To gDiff.Rows - 1
1250                  If InStr(UCase(gDiff.TextMatrix(n, 0)), "MONO") > 0 Then
1260                      neut = gDiff.TextMatrix(n, 2)
1270                      lym = gDiff.TextMatrix(n, 1)
1280                      Exit For
1290                  End If
1300              Next
1310          Else
1320              neut = Trim(tb!MonoA & "")
1330              lym = Trim(tb!MonoP & "")
1340          End If


1350          If Not IsNull(neut) Then
1360              Colourise "MonoA", tMonoA, neut, lblSex, lblDoB, lblSampleDate
1370          End If

1380          If Not IsNull(lym) Then
1390              Colourise "MonoP", tMonoP, lym, lblSex, lblDoB, lblSampleDate
1400          End If

1410          neut = ""
1420          lym = ""
1430          If DiffFound = True Then
1440              For n = 1 To gDiff.Rows - 1
1450                  If InStr(UCase(gDiff.TextMatrix(n, 0)), "NEUT") > 0 Then
1460                      neut = gDiff.TextMatrix(n, 2)
1470                      lym = gDiff.TextMatrix(n, 1)
1480                      Exit For
1490                  End If
1500              Next
1510          Else
1520              neut = Trim(tb!NeutA & "")
1530              lym = Trim(tb!NeutP & "")
1540          End If

1550          If Not IsNull(neut) Then
1560              Colourise "NeutA", tNeutA, neut, lblSex, lblDoB, lblSampleDate
1570          End If

1580          If Not IsNull(lym) Then
1590              Colourise "NeutP", tNeutP, lym, lblSex, lblDoB, lblSampleDate
1600          End If

1610          If DiffFound = True Then
1620              For n = 1 To gDiff.Rows - 1
1630                  If InStr(UCase(gDiff.TextMatrix(n, 0)), "EOS") > 0 Then
1640                      neut = gDiff.TextMatrix(n, 2)
1650                      lym = gDiff.TextMatrix(n, 1)
1660                      Exit For
1670                  End If
1680              Next
1690          Else
1700              neut = Trim(tb!EosA & "")
1710              lym = Trim(tb!EosP & "")
1720          End If
1730          If Not IsNull(neut) Then
1740              Colourise "EosA", tEosA, neut, lblSex, lblDoB, lblSampleDate
1750          End If

1760          If Not IsNull(lym) Then
1770              Colourise "EosP", tEosP, lym, lblSex, lblDoB, lblSampleDate
1780          End If

1790          If DiffFound = True Then
1800              For n = 1 To gDiff.Rows - 1
1810                  If InStr(UCase(gDiff.TextMatrix(n, 0)), "BAS") > 0 Then
1820                      neut = gDiff.TextMatrix(n, 2)
1830                      lym = gDiff.TextMatrix(n, 1)
1840                      Exit For
1850                  End If
1860              Next
1870          Else
1880              neut = Trim(tb!BasA & "")
1890              lym = Trim(tb!BasP & "")
1900          End If
1910          If Not IsNull(neut) Then
1920              Colourise "BasA", tBasA, neut, lblSex, lblDoB, lblSampleDate
1930          End If

1940          If Not IsNull(lym) Then
1950              Colourise "BasP", tBasP, lym, lblSex, lblDoB, lblSampleDate
1960          End If

1970          If DiffFound = True Then
1980              neut = ""
1990              lym = ""
2000              For n = 1 To gDiff.Rows - 1
2010                  If InStr(UCase(gDiff.TextMatrix(n, 0)), "LUC") > 0 Then
2020                      neut = gDiff.TextMatrix(n, 2)
2030                      lym = gDiff.TextMatrix(n, 1)
2040                      Exit For
2050                  End If
2060              Next
2070          Else
2080              neut = Trim(tb!luca & "")
2090              lym = Trim(tb!lucp & "")
2100          End If

2110          If Not IsNull(neut) Then
2120              Colourise "LucA", tLucA, neut, lblSex, lblDoB, lblSampleDate
2130          End If

2140          If Not IsNull(lym) Then
2150              Colourise "LucP", tLucP, lym, lblSex, lblDoB, lblSampleDate
2160          End If


2170          If SysOptHaemAn1(0) = "ADVIA" And tb!Image & "" <> "" Then
2180              imgHaemGraphs.Visible = True
2190          End If

2200      End If

2210      LockHaem True


2220      Exit Sub

LoadHaem_Error:

          Dim strES As String
          Dim intEL As Integer



2230      intEL = Erl
2240      strES = Err.Description
2250      LogError "frmViewResults", "LoadHaem", intEL, strES, sql

End Sub
Private Sub Colourise(ByVal Analyte As String, _
                      ByVal Destination As TextBox, _
                      ByVal strValue As String, _
                      ByVal sex As String, _
                      ByVal Dob As String, _
                      ByVal Rundate As String)

          Dim Value As Single
          Dim x As Long


10        On Error GoTo Colourise_Error

20        Value = Val(strValue)

30        If InStr(strValue, ">") Then
40            x = InStr(strValue, ">")
50            Value = Mid(strValue, x + 1)
60        End If

70        Destination.Text = strValue
80        If Trim$(strValue) = "" Then
90            Destination.BackColor = &HFFFFFF
100           Destination.ForeColor = &H0&
110           Exit Sub
120       End If



130       Select Case InterpH(Value, Analyte, sex, Dob, Tn, Rundate)
          Case "X":
140           Destination.BackColor = vbBlack
150           Destination.ForeColor = vbWhite
160       Case "H":
170           Destination.BackColor = vbRed
180           Destination.ForeColor = vbYellow
190       Case "L"
200           Destination.BackColor = vbBlue
210           Destination.ForeColor = vbYellow
220       Case Else
230           Destination.BackColor = &HFFFFFF
240           Destination.ForeColor = &H0&
250       End Select



260       Exit Sub

Colourise_Error:

          Dim strES As String
          Dim intEL As Integer



270       intEL = Erl
280       strES = Err.Description
290       LogError "frmViewResults", "Colourise", intEL, strES


End Sub

Private Sub ClearHaem()


10        On Error GoTo ClearHaem_Error

20        tWBC = ""
30        tWBC.BackColor = &HFFFFFF
40        tWBC.ForeColor = &H0&

50        tRBC = ""
60        tRBC.BackColor = &HFFFFFF
70        tRBC.ForeColor = &H0&
80        tHypo = ""

90        tHgb = ""
100       tHgb.BackColor = &HFFFFFF
110       tHgb.ForeColor = &H0&

120       tMCV = ""
130       tMCV.BackColor = &HFFFFFF
140       tMCV.ForeColor = &H0&

150       tHct = ""
160       tHct.BackColor = &HFFFFFF
170       tHct.ForeColor = &H0&

180       tRDWCV = ""
190       tRDWCV.BackColor = &HFFFFFF
200       tRDWCV.ForeColor = &H0&

210       tRDWSD = ""
220       tRDWSD.BackColor = &HFFFFFF
230       tRDWSD.ForeColor = &H0&

240       tMCH = ""
250       tMCH.BackColor = &HFFFFFF
260       tMCH.ForeColor = &H0&

270       tMCHC = ""
280       tMCHC.BackColor = &HFFFFFF
290       tMCHC.ForeColor = &H0&

300       tPlt = ""
310       tPlt.BackColor = &HFFFFFF
320       tPlt.ForeColor = &H0&

330       tMPV = ""
340       tMPV.BackColor = &HFFFFFF
350       tMPV.ForeColor = &H0&

360       tPLCR = ""
370       tPLCR.BackColor = &HFFFFFF
380       tPLCR.ForeColor = &H0&

390       tPdw = ""
400       tPdw.BackColor = &HFFFFFF
410       tPdw.ForeColor = &H0&

420       tLymA = ""
430       tLymA.BackColor = &HFFFFFF
440       tLymA.ForeColor = &H0&

450       tLymP = ""
460       tLymP.BackColor = &HFFFFFF
470       tLymP.ForeColor = &H0&

480       tMonoA = ""
490       tMonoA.BackColor = &HFFFFFF
500       tMonoA.ForeColor = &H0&

510       tMonoP = ""
520       tMonoP.BackColor = &HFFFFFF
530       tMonoP.ForeColor = &H0&

540       tNeutA = ""
550       tNeutA.BackColor = &HFFFFFF
560       tNeutA.ForeColor = &H0&

570       tNeutP = ""
580       tNeutP.BackColor = &HFFFFFF
590       tNeutP.ForeColor = &H0&

600       tEosA = ""
610       tEosA.BackColor = &HFFFFFF
620       tEosA.ForeColor = &H0&

630       tEosP = ""
640       tEosP.BackColor = &HFFFFFF
650       tEosP.ForeColor = &H0&

660       tBasA = ""
670       tBasA.BackColor = &HFFFFFF
680       tBasA.ForeColor = &H0&

690       tBasP = ""
700       tBasP.BackColor = &HFFFFFF
710       tBasP.ForeColor = &H0&

720       tWBC = ""
730       tWBC.BackColor = &HFFFFFF
740       tWBC.ForeColor = &H0&

750       tLucP = ""
760       tLucP.BackColor = &HFFFFFF
770       tLucP.ForeColor = &H0&

780       tLucA = ""
790       tLucA.BackColor = &HFFFFFF
800       tLucA.ForeColor = &H0&

810       tLI = ""
820       tLI.BackColor = &HFFFFFF
830       tLI.ForeColor = &H0&

840       tMpXi = ""
850       tMpXi.BackColor = &HFFFFFF
860       tMpXi.ForeColor = &H0&

870       tHDW = ""
880       tHDW.BackColor = &HFFFFFF
890       tHDW.ForeColor = &H0&

900       tMalaria = ""
910       tPdw.BackColor = &HFFFFFF
920       tPdw.ForeColor = &H0&

930       tSickledex = ""
940       tPdw.BackColor = &HFFFFFF
950       tPdw.ForeColor = &H0&

960       lesr = ""
970       lretics = ""
980       lmonospot = ""
990       lASOT = ""
1000      lRa = ""
1010      tNrbc = ""




1020      Exit Sub

ClearHaem_Error:

          Dim strES As String
          Dim intEL As Integer



1030      intEL = Erl
1040      strES = Err.Description
1050      LogError "frmViewResults", "ClearHaem", intEL, strES


End Sub

Private Sub LoadCoag()

          Dim Cxs As New CoagResults
          Dim Cx As CoagResult
          Dim s As String
          Dim FormatStr As String
          Dim Low As Single
          Dim High As Single
          Dim n As Long
          Dim x As Long
          Dim sql As String
          Dim tb As New Recordset
          Dim DaysOld As Long


10        On Error GoTo LoadCoag_Error

20        gCoag.Rows = 2
30        gCoag.AddItem ""
40        gCoag.RemoveItem 1

50        cmdPrint(2).Enabled = False
60        cmdFax(2).Enabled = False
70        cmdExcel(2).Enabled = False

80        If lblDoB <> "" And Len(lblDoB) > 9 And lblRundate <> "" And IsDate(lblRundate) Then
90            DaysOld = Abs(DateDiff("d", lblRundate, lblDoB))
100       Else
110           DaysOld = 12783
120       End If
130       If DaysOld = 0 Then DaysOld = 1


140       If UserMemberOf = "Secretarys" Or UCase(UserMemberOf) = "HISTOLOOKUP" Then
150           Set Cxs = Cxs.Load(lblSampleID, gVALID, gDONTCARE, Trim(SysOptExp(Tn)), Tn)
160       Else
170           Set Cxs = Cxs.Load(lblSampleID, gDONTCARE, gDONTCARE, Trim(SysOptExp(Tn)), Tn)
180       End If
190       If Cxs.Count <> 0 Then
200           cmdPrint(2).Enabled = True
210           cmdFax(2).Enabled = True
220           cmdExcel(2).Enabled = True
230           For Each Cx In Cxs
                  '        sql = "SELECT * FROM CoagTestDefinitions WHERE " & _
                           '              "Code = '" & Cx.Code & "'"
240               sql = "SELECT * from coagtestdefinitions WHERE " & _
                        "(code = '" & Trim(Cx.Code) & "' OR TestName = '" & Trim$(Cx.Code) & "') " & _
                        "and agefromdays <= '" & DaysOld & "' and agetodays >= '" & DaysOld & "'"
250               Set tb = New Recordset
260               RecOpenServer Tn, tb, sql
270               If Not tb.EOF Then
280                   'If Trim(Cx.Units) = "INR" Then
290                       's = "INR" & vbTab
300                   'Else
310                       s = tb!TestName & vbTab
320                   'End If
330                   Select Case Trim(tb!DP)
                      Case 0: FormatStr = "###0"
340                   Case 1: FormatStr = "##0.0"
350                   Case 2: FormatStr = "#0.00"
360                   Case 3: FormatStr = "0.000"
370                   End Select
380                   s = s & Format(Cx.Result, FormatStr) & vbTab
390                   If Trim(Cx.Units) = "INR" Then
400                       s = s & vbTab
410                   ElseIf Trim(Cx.Units) = "ÆG/ML" Then
420                       s = s & "ug/ml" & vbTab
430                   Else
440                       s = s & Cx.Units & vbTab
450                   End If
460                   Select Case UCase(Left(lblSex, 1))
                      Case "M": Low = tb!MaleLow: High = tb!MaleHigh
470                   Case "F": Low = tb!FemaleLow: High = tb!FemaleHigh
480                   Case Else: Low = tb!FemaleLow: High = tb!MaleHigh
490                   End Select
                      'Zyam 5-07-24
                      If InStr(1, Cx.Result, ">") Then
                        Cx.Result = Right(Cx.Result, Len(Cx.Result) - 1)
                      End If
                      'Zyam
                      If Cx.Code <> "1" And Cx.Code <> "13" And Cx.Code <> "14" And Cx.Code <> "27" And Cx.Code <> "94" And Cx.Code <> "95" Then
500                     If Val(Cx.Result) <> 0 Then
510                         If Cx.Result < Low Then
520                             s = s & "L"
530                         ElseIf Cx.Result > High Then
540                             s = s & "H"
550                         End If
560                     End If
                      End If
570                   s = s & vbTab & _
                          IIf(Cx.Valid, "V", "") & _
                          IIf(Cx.Printed, "P", "")
580                   gCoag.AddItem s
590               End If
600           Next
610       End If

620       If gCoag.Rows > 2 Then
630           gCoag.RemoveItem 1
640       End If

650       For n = 1 To gCoag.Rows - 1
660           If gCoag.TextMatrix(n, 3) = "L" Then
670               For x = 0 To 4
680                   gCoag.Row = n
690                   gCoag.Col = x
700                   gCoag.CellBackColor = vbBlue
710                   gCoag.CellForeColor = vbYellow
720               Next
730           ElseIf gCoag.TextMatrix(n, 3) = "H" Then
740               For x = 0 To 4
750                   gCoag.Row = n
760                   gCoag.Col = x
770                   gCoag.CellBackColor = vbRed
780                   gCoag.CellForeColor = vbYellow
790               Next
800           Else
810               For x = 0 To 4
820                   gCoag.Row = n
830                   gCoag.Col = x
840                   gCoag.CellBackColor = vbWhite
850                   gCoag.CellForeColor = vbBlack
860               Next
870           End If
880       Next



890       Exit Sub

LoadCoag_Error:

          Dim strES As String
          Dim intEL As Integer



900       intEL = Erl
910       strES = Err.Description
920       LogError "frmViewResults", "LoadCoag", intEL, strES


End Sub

Private Sub LoadBiochemistry()

          Dim s As String
          Dim Value As Single
          Dim valu As String
          Dim BRs As New BIEResults
          Dim br As BIEResult
          Dim CodeGLU As String
          Dim CodeCHO As String
          Dim CodeTRI As String
          Dim T As String

10        On Error GoTo LoadBiochemistry_Error

20        If UserMemberOf = "Secretarys" Or UCase(UserMemberOf) = "HISTOLOOKUP" Then
30            Set BRs = BRs.Load("Bio", lblSampleID, "Results", gVALID, gDONTCARE, Tn, "", lblRundate)
40        Else
50            Set BRs = BRs.Load("Bio", lblSampleID, "Results", gDONTCARE, gDONTCARE, Tn, "", lblRundate)
60        End If

70        gBio.Visible = False
80        gBio.Rows = 2
90        gBio.AddItem ""
100       gBio.RemoveItem 1
110       cmdPrint(0).Enabled = False
120       cmdFax(0).Enabled = False
130       cmdExcel(0).Enabled = False
140       CodeGLU = "996"
150       CodeCHO = "1"
160       CodeTRI = "62"

170       For Each br In BRs
180           cmdPrint(0).Enabled = True
190           cmdFax(0).Enabled = True
200           cmdExcel(0).Enabled = True
210           valu = ""
220           If UCase(br.Result) = "POS" Or UCase(br.Result) = "NEG" Then valu = br.Result Else Value = Val(br.Result)
230           If Not IsNumeric(br.Result) Then valu = br.Result
240           If TestAffected(br) = True And UserMemberOf = "Secretarys" Then
250               s = br.LongName & vbTab & "*****"
260           ElseIf TestAffected(br) = True And SysOptBioMask(0) Then
270               s = br.LongName & vbTab & "*****"

280           Else
290               If valu = "" Then
300                   Select Case br.Printformat
                      Case 0: valu = Format(Value, "0")
310                   Case 1: valu = Format(Value, "0.0")
320                   Case 2: valu = Format(Value, "0.00")
330                   Case 3: valu = Format(Value, "0.000")
340                   Case Else: valu = Format(Value, "0.000")
350                   End Select
360               End If
370               T = QuickInterpBio(br)
380               If T = "***" Then valu = T
390               s = br.LongName & vbTab & valu & vbTab & IIf(br.Valid, "V", "")
400           End If
410           gBio.AddItem s
420           Select Case Trim$(T)
              Case "Low":
430               gBio.Row = gBio.Rows - 1
440               gBio.Col = 1
450               gBio.CellBackColor = vbBlue
460               gBio.CellForeColor = vbYellow
470           Case "High":
480               gBio.Row = gBio.Rows - 1
490               gBio.Col = 1
500               gBio.CellBackColor = vbRed
510               gBio.CellForeColor = vbYellow
520           Case Else:
530               gBio.Row = gBio.Rows - 1
540               gBio.Col = 1
550               gBio.CellBackColor = 0
560           End Select
570       Next

580       If gBio.Rows > 2 Then
590           gBio.RemoveItem 1
600       End If

610       gBio.Visible = True

620       Exit Sub

LoadBiochemistry_Error:

          Dim strES As String
          Dim intEL As Integer

630       intEL = Erl
640       strES = Err.Description
650       LogError "frmViewResults", "LoadBiochemistry", intEL, strES

End Sub

Private Sub Form_Deactivate()

10        On Error GoTo Form_Deactivate_Error

20        Timer1.Enabled = False

30        Exit Sub

Form_Deactivate_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmViewResults", "Form_Deactivate", intEL, strES


End Sub


Private Sub Form_Load()

10        On Error GoTo Form_Load_Error

20        Activated = False

30        gBio.Font.Bold = True

40        pBar.Max = LogOffDelaySecs
50        pBar = 0

60        Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

70        intEL = Erl
80        strES = Err.Description
90        LogError "frmViewResults", "Form_Load", intEL, strES

End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

10        On Error GoTo Form_MouseMove_Error

20        pBar = 0

30        Exit Sub

Form_MouseMove_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmViewResults", "Form_MouseMove", intEL, strES


End Sub

Private Sub Form_Unload(Cancel As Integer)

10        On Error GoTo Form_Unload_Error

20        Activated = False

30        Exit Sub

Form_Unload_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmViewResults", "Form_Unload", intEL, strES


End Sub




Private Sub gImm_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

10        On Error GoTo gImm_MouseMove_Error

20        Y = gImm.MouseCol
30        x = gImm.MouseRow
40        gImm.ToolTipText = "Immunology Results"

50        If gImm.MouseCol = 3 Or gImm.MouseCol = 1 Then
60            If Trim(gImm.TextMatrix(x, Y)) <> "" Then gImm.ToolTipText = gImm.TextMatrix(x, Y)
70        End If


80        Exit Sub

gImm_MouseMove_Error:

          Dim strES As String
          Dim intEL As Integer



90        intEL = Erl
100       strES = Err.Description
110       LogError "frmViewResults", "gImm_MouseMove", intEL, strES


End Sub

Private Sub imgEarliest_Click()


10        On Error GoTo imgEarliest_Click_Error

20        tbDem.MoveFirst
30        CurrentRecordNumber = 1

40        FillDemographics
50        LoadAllResults

60        imgEarliest.Visible = False
70        imgPrevious.Visible = False
80        If Nrc > 1 Then
90            imgNext.Visible = True
100           imgLatest.Visible = True
110       End If

120       pBar = 0




130       Exit Sub

imgEarliest_Click_Error:

          Dim strES As String
          Dim intEL As Integer



140       intEL = Erl
150       strES = Err.Description
160       LogError "frmViewResults", "imgEarliest_Click", intEL, strES


End Sub

Private Sub imgHaemGraphs_Click()

10        On Error GoTo imgHaemGraphs_Click_Error

20        frmHaemGraphs.SampleID = lblSampleID
30        frmHaemGraphs.Show 1

40        Exit Sub

imgHaemGraphs_Click_Error:

          Dim strES As String
          Dim intEL As Integer



50        intEL = Erl
60        strES = Err.Description
70        LogError "frmViewResults", "imgHaemGraphs_Click", intEL, strES


End Sub

Private Sub imgLatest_Click()


10        On Error GoTo imgLatest_Click_Error

20        tbDem.MoveLast
30        CurrentRecordNumber = Nrc

40        FillDemographics
50        LoadAllResults

60        imgNext.Visible = False
70        imgLatest.Visible = False

80        If Nrc > 1 Then
90            imgPrevious.Visible = True
100           imgEarliest.Visible = True
110       End If

120       pBar = 0

130       Exit Sub

imgLatest_Click_Error:

          Dim strES As String
          Dim intEL As Integer



140       intEL = Erl
150       strES = Err.Description
160       LogError "frmViewResults", "imgLatest_Click", intEL, strES


End Sub

Private Sub imgNext_Click()


10        On Error GoTo imgNext_Click_Error

20        tbDem.MoveNext
30        CurrentRecordNumber = CurrentRecordNumber + 1

40        FillDemographics
50        LoadAllResults

60        If CurrentRecordNumber < Nrc Then
70            imgNext.Visible = True
80            imgLatest.Visible = True
90        Else
100           imgNext.Visible = False
110           imgLatest.Visible = False
120       End If

130       If Nrc > 1 Then
140           imgPrevious.Visible = True
150           imgEarliest.Visible = True
160       End If

170       pBar = 0



180       Exit Sub

imgNext_Click_Error:

          Dim strES As String
          Dim intEL As Integer



190       intEL = Erl
200       strES = Err.Description
210       LogError "frmViewResults", "imgNext_Click", intEL, strES


End Sub

Private Sub imgPrevious_Click()


10        On Error GoTo imgPrevious_Click_Error

20        tbDem.MovePrevious
30        CurrentRecordNumber = CurrentRecordNumber - 1

40        FillDemographics
50        LoadAllResults

60        If CurrentRecordNumber > 1 Then
70            imgEarliest.Visible = True
80            imgPrevious.Visible = True
90        Else
100           imgEarliest.Visible = False
110           imgPrevious.Visible = False
120       End If

130       If Nrc > 1 Then
140           imgNext.Visible = True
150           imgLatest.Visible = True
160       End If

170       pBar = 0



180       Exit Sub

imgPrevious_Click_Error:

          Dim strES As String
          Dim intEL As Integer



190       intEL = Erl
200       strES = Err.Description
210       LogError "frmViewResults", "imgPrevious_Click", intEL, strES


End Sub




Private Sub Timer1_Timer()

      'tmrRefresh.Interval set to 1000
10        On Error GoTo Timer1_Timer_Error

20        pBar = pBar + 1

30        If pBar = pBar.Max Then
40            Unload Me
50        End If

60        Exit Sub

Timer1_Timer_Error:

          Dim strES As String
          Dim intEL As Integer



70        intEL = Erl
80        strES = Err.Description
90        LogError "frmViewResults", "Timer1_Timer", intEL, strES


End Sub


Private Sub LoadImmunology()

          Dim s As String
          Dim Value As Single
          Dim valu As String
          Dim BRs As New BIEResults
          Dim br As BIEResult
          Dim T As String



10        On Error GoTo LoadImmunology_Error

20        If UserMemberOf = "Secretarys" Or UCase(UserMemberOf) = "HISTOLOOKUP" Then
30            Set BRs = BRs.Load("Imm", lblSampleID, "Results", gVALID, gDONTCARE, Tn, "Default", lblRundate)
40        Else
50            Set BRs = BRs.Load("Imm", lblSampleID, "Results", gDONTCARE, gDONTCARE, Tn, "Default", lblRundate)
60        End If

70        gImm.Visible = False
80        gImm.Rows = 2
90        gImm.AddItem ""
100       gImm.RemoveItem 1
110       cmdPrint(5).Enabled = False
120       cmdFax(5).Enabled = False
130       cmdExcel(5).Enabled = False



140       For Each br In BRs
150           cmdPrint(5).Enabled = True
160           cmdFax(5).Enabled = True
170           cmdExcel(5).Enabled = True
180           valu = ""
190           If ImmTestAffected(br) = True And UserMemberOf = "Secretarys" Then
200               s = br.LongName & vbTab & "*****"
210           Else
220               If br.Result = "POS" Or br.Result = "NEG" Or Not IsNumeric(br.Result) Then valu = br.Result Else Value = Val(br.Result)
230               If valu = "" Then
240                   Select Case br.Printformat
                      Case 0: valu = Format(Value, "0")
250                   Case 1: valu = Format(Value, "0.0")
260                   Case 2: valu = Format(Value, "0.00")
270                   Case 3: valu = Format(Value, "0.000")
280                   Case Else: valu = Format(Value, "0.000")
290                   End Select
300               End If
310               T = QuickInterpImm(br)
320               If T = "***" Then valu = T
330               s = br.LongName & vbTab & valu & vbTab & IIf(br.Valid, "V", "") & vbTab & br.Comment
340           End If
350           gImm.AddItem s
360           Select Case Trim$(T)
              Case "Low":
370               gImm.Row = gImm.Rows - 1
380               gImm.Col = 1
390               gImm.CellBackColor = vbBlue
400               gImm.CellForeColor = vbYellow
410           Case "High":
420               gImm.Row = gImm.Rows - 1
430               gImm.Col = 1
440               gImm.CellBackColor = vbRed
450               gImm.CellForeColor = vbYellow
460           Case Else:
470               gImm.Row = gImm.Rows - 1
480               gImm.Col = 1
490               gImm.CellBackColor = 0
500           End Select
510       Next

520       If gImm.Rows > 2 Then
530           gImm.RemoveItem 1
540       End If

550       gImm.Visible = True



560       Exit Sub

LoadImmunology_Error:

          Dim strES As String
          Dim intEL As Integer



570       intEL = Erl
580       strES = Err.Description
590       LogError "frmViewResults", "LoadImmunology", intEL, strES


End Sub

Private Sub LoadBloodGas()

          Dim s As String
          Dim Value As Single
          Dim valu As String
          Dim BRs As New BIEResults
          Dim BRres As BIEResults
          Dim br As BIEResult
          Dim T As String
          Dim n As Long

10        On Error GoTo LoadBloodGas_Error

20        If lblSampleID = "" Then Exit Sub

30        ClearFGrid gBga

40        Set BRres = BRs.Load("Bga", lblSampleID, "Results", gDONTCARE, gDONTCARE, 0, "", lblRundate)

50        With gBga
60            .Rows = 2
70            .AddItem ""
80            .RemoveItem 1
90        End With

100       cmdPrint(4).Enabled = False
110       cmdFax(4).Enabled = False
120       cmdExcel(4).Enabled = False

130       If Not BRres Is Nothing Then
140           cmdExcel(4).Enabled = True
150           For Each br In BRres
160               s = br.ShortName & vbTab
170               lblRundate = br.RunTime
180               If IsNumeric(br.Result) Then
190                   Value = Val(br.Result)
200                   Select Case br.Printformat
                      Case 0: valu = Format$(Value, "0")
210                   Case 1: valu = Format$(Value, "0.0")
220                   Case 2: valu = Format$(Value, "0.00")
230                   Case 3: valu = Format$(Value, "0.000")
240                   Case Else: valu = Format$(Value, "0.000")
250                   End Select
260               Else
270                   valu = br.Result
280               End If
290               s = s & valu
300               T = ""
310               If IsNumeric(br.Result) Then
320                   If Value > Val(br.PlausibleHigh) Then
330                       s = s & "X"
340                   ElseIf Value < Val(br.PlausibleLow) Then
350                       s = s & "X"
360                   Else
370                       If Value < Val(br.FlagLow) Then
380                           T = "FL"
390                       ElseIf Value > Val(br.FlagHigh) Then
400                           T = "FH"
410                       End If
420                       If Value < Val(br.Low) Then
430                           T = "L"
440                       ElseIf Value > Val(br.High) Then
450                           T = "H"
460                       End If
470                   End If
480               End If
490               s = s & vbTab & _
                      IIf(br.Valid, "V", " ") & vbTab & _
                      IIf(br.Printed, "P", " ") & vbTab
500               gBga.AddItem s

510               If T <> "" Then
520                   gBga.Row = gBga.Rows - 1
530                   gBga.Col = 1
540                   Select Case T
                      Case "H":
550                       For n = 0 To 7
560                           gBga.Col = n
570                           gBga.CellBackColor = SysOptHighBack(0)
580                           gBga.CellForeColor = SysOptHighFore(0)
590                       Next
600                   Case "L":
610                       For n = 0 To 7
620                           gBga.Col = n
630                           gBga.CellBackColor = SysOptLowBack(0)
640                           gBga.CellForeColor = SysOptLowFore(0)
650                       Next
660                   Case "X":
670                       For n = 0 To 7
680                           gBga.Col = n
690                           gBga.CellBackColor = SysOptPlasBack(0)
700                           gBga.CellForeColor = SysOptPlasFore(0)
710                       Next
720                   End Select
730               End If

740           Next
750       End If

760       FixG gBga

770       Exit Sub

LoadBloodGas_Error:

          Dim strES As String
          Dim intEL As Integer

780       intEL = Erl
790       strES = Err.Description
800       LogError "frmViewResults", "LoadBloodGas", intEL, strES

End Sub



Private Sub LoadEndocrinology()

          Dim s As String
          Dim Value As Single
          Dim valu As String
          Dim BRs As New BIEResults
          Dim br As BIEResult
          Dim T As String



10        On Error GoTo LoadEndocrinology_Error

20        If UserMemberOf = "Secretarys" Or UCase(UserMemberOf) = "HISTOLOOKUP" Then
30            Set BRs = BRs.Load("End", lblSampleID, "Results", gVALID, gDONTCARE, Tn, "Default", lblRundate)
40        Else
50            Set BRs = BRs.Load("End", lblSampleID, "Results", gDONTCARE, gDONTCARE, Tn, "Default", lblRundate)
60        End If

70        gEnd.Visible = False
80        gEnd.Rows = 2
90        gEnd.AddItem ""
100       gEnd.RemoveItem 1
110       cmdPrint(3).Enabled = False
120       cmdFax(3).Enabled = False
130       cmdExcel(3).Enabled = False



140       For Each br In BRs
150           cmdPrint(3).Enabled = True
160           cmdFax(3).Enabled = True
170           cmdExcel(3).Enabled = True
180           valu = ""
190           If UCase(br.Analyser) = "VIROLOGY" Then
                  'if AxSym virology then translate result here
200               valu = TranslateEndResultVirology(br.Code, br.Result)
                  'now result is non numeric. so it won't generate any flags or apply any rules
210           ElseIf br.Result = "POS" Or br.Result = "NEG" Or Not IsNumeric(br.Result) Then
220               valu = br.Result
230               If Left(br.Result, 1) = "<" Then
240                   T = "Low"
250               ElseIf Left(br.Result, 1) = ">" Then
260                   T = "High"
270               End If

280           Else
290               Value = Val(br.Result)
300               Select Case br.Printformat
                  Case 0: valu = Format(Value, "0")
310               Case 1: valu = Format(Value, "0.0")
320               Case 2: valu = Format(Value, "0.00")
330               Case 3: valu = Format(Value, "0.000")
340               Case Else: valu = Format(Value, "0.000")
350               End Select
360               T = QuickInterpEnd(br)
370           End If

380           If T = "***" Then valu = T
390           s = IIf(UCase(br.Analyser) = "VIROLOGY", br.ShortName, br.LongName) & vbTab & valu & vbTab & IIf(br.Valid, "V", "")
400           gEnd.AddItem s
410           If UCase(br.Analyser) <> "VIROLOGY" Then
420               Select Case Trim$(T)
                  Case "Low":
430                   gEnd.Row = gEnd.Rows - 1
440                   gEnd.Col = 1
450                   gEnd.CellBackColor = vbBlue
460                   gEnd.CellForeColor = vbYellow
470               Case "High":
480                   gEnd.Row = gEnd.Rows - 1
490                   gEnd.Col = 1
500                   gEnd.CellBackColor = vbRed
510                   gEnd.CellForeColor = vbYellow
520               Case Else:
530                   gEnd.Row = gEnd.Rows - 1
540                   gEnd.Col = 1
550                   gEnd.CellBackColor = 0
560               End Select
570           End If
580       Next

590       If gEnd.Rows > 2 Then
600           gEnd.RemoveItem 1
610       End If

620       gEnd.Visible = True




630       Exit Sub

LoadEndocrinology_Error:

          Dim strES As String
          Dim intEL As Integer



640       intEL = Erl
650       strES = Err.Description
660       LogError "frmViewResults", "LoadEndocrinology", intEL, strES


End Sub

Private Sub LockHaem(ByVal iLock As Boolean)

10        On Error GoTo LockHaem_Error

20        tHgb.Locked = iLock
30        tRBC.Locked = iLock
40        tWBC.Locked = iLock
50        tHct.Locked = iLock
60        tMCV.Locked = iLock
70        tMCHC.Locked = iLock
80        tNrbc.Locked = iLock
90        tHDW.Locked = iLock
100       tNeutA.Locked = iLock
110       tNeutP.Locked = iLock
120       tMonoA.Locked = iLock
130       tMonoP.Locked = iLock
140       tEosA.Locked = iLock
150       tEosP.Locked = iLock
160       tLymA.Locked = iLock
170       tLymP.Locked = iLock
180       tBasA.Locked = iLock
190       tBasP.Locked = iLock
200       tPlt.Locked = iLock
210       tLucA.Locked = iLock
220       tLucP.Locked = iLock
230       tMPV.Locked = iLock
240       tPdw.Locked = iLock
250       tMpXi.Locked = iLock
260       tLI.Locked = iLock
270       tPLCR.Locked = iLock
280       tMalaria.Locked = iLock
290       tSickledex.Locked = iLock

300       Exit Sub

LockHaem_Error:

          Dim strES As String
          Dim intEL As Integer



310       intEL = Erl
320       strES = Err.Description
330       LogError "frmViewResults", "LockHaem", intEL, strES

End Sub

