VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.Ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmConsultantListView 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   10290
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   18405
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10290
   ScaleWidth      =   18405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraDartViewer 
      Caption         =   "Enter Sample ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   15480
      TabIndex        =   47
      Top             =   660
      Width           =   2535
      Begin VB.TextBox txtDartSampleID 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   120
         TabIndex        =   49
         Top             =   300
         Width           =   1755
      End
      Begin VB.CommandButton cmdDartViewer 
         Height          =   405
         Left            =   1980
         Picture         =   "frmConsultantListView.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   300
         Width           =   375
      End
   End
   Begin VB.Frame frmeRefreshing 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00008000&
      Height          =   2400
      Left            =   5000
      TabIndex        =   44
      Top             =   4000
      Width           =   3825
      Begin VB.CommandButton btnCancel 
         Caption         =   "Cancel"
         Height          =   465
         Left            =   2760
         TabIndex        =   46
         Top             =   1860
         Width           =   1005
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Please wait while report is refreshing"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   1665
         Left            =   60
         TabIndex        =   45
         Top             =   60
         Width           =   3690
      End
   End
   Begin VB.CommandButton bCancel 
      Appearance      =   0  'Flat
      Caption         =   "Exit"
      Height          =   900
      Left            =   16980
      Picture         =   "frmConsultantListView.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   40
      Tag             =   "101"
      ToolTipText     =   "Exit Screen"
      Top             =   9060
      Width           =   1000
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   10065
      Left            =   150
      TabIndex        =   0
      Top             =   120
      Width           =   18135
      _ExtentX        =   31988
      _ExtentY        =   17754
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "HOSPITAL ONE"
      TabPicture(0)   =   "frmConsultantListView.frx":0BD4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraMicro"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "HOSPITAL TWO"
      TabPicture(1)   =   "frmConsultantListView.frx":0BF0
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "HOSPITAL THREE"
      TabPicture(2)   =   "frmConsultantListView.frx":0C0C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame2"
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   2955
         Left            =   13400
         TabIndex        =   22
         Top             =   6885
         Width           =   4515
         Begin VB.CommandButton cmdRefresh 
            Appearance      =   0  'Flat
            Caption         =   "Refresh"
            Height          =   900
            Index           =   0
            Left            =   1080
            Picture         =   "frmConsultantListView.frx":0C28
            Style           =   1  'Graphical
            TabIndex        =   41
            Tag             =   "101"
            ToolTipText     =   "Exit Screen"
            Top             =   2040
            Width           =   1000
         End
         Begin VB.CommandButton cmdSaveC 
            Caption         =   "Save"
            Height          =   900
            Left            =   0
            Picture         =   "frmConsultantListView.frx":14F2
            Style           =   1  'Graphical
            TabIndex        =   37
            ToolTipText     =   "Save Changes"
            Top             =   2040
            Width           =   1000
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
            Height          =   315
            Left            =   4050
            TabIndex        =   25
            ToolTipText     =   "Choose a comment from a list"
            Top             =   1680
            Width           =   435
         End
         Begin VB.ComboBox cmbConC 
            Height          =   315
            Left            =   60
            TabIndex        =   23
            Text            =   "cmbConC"
            Top             =   1680
            Visible         =   0   'False
            Width           =   3975
         End
         Begin VB.TextBox txtConC 
            Height          =   2000
            Left            =   0
            MultiLine       =   -1  'True
            TabIndex        =   24
            Top             =   0
            Width           =   4515
         End
      End
      Begin VB.Frame Frame2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   9615
         Left            =   -74940
         TabIndex        =   15
         Top             =   360
         Width           =   18000
         Begin VB.CommandButton cmdAck2 
            Caption         =   "Acknowledge"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1200
            Left            =   14970
            Picture         =   "frmConsultantListView.frx":17FC
            Style           =   1  'Graphical
            TabIndex        =   67
            Top             =   4740
            Width           =   1200
         End
         Begin VB.Frame Frame7 
            Height          =   3612
            Left            =   13200
            TabIndex        =   61
            Top             =   1080
            Width           =   4752
            Begin TabDlg.SSTab SSTab3 
               Height          =   3372
               Left            =   60
               TabIndex        =   62
               Top             =   180
               Width           =   4668
               _ExtentX        =   8229
               _ExtentY        =   5953
               _Version        =   393216
               Tabs            =   2
               TabsPerRow      =   2
               TabHeight       =   420
               TabCaption(0)   =   "Ready for Consultant"
               TabPicture(0)   =   "frmConsultantListView.frx":20C6
               Tab(0).ControlEnabled=   -1  'True
               Tab(0).Control(0)=   "grdSid3"
               Tab(0).Control(0).Enabled=   0   'False
               Tab(0).ControlCount=   1
               TabCaption(1)   =   "Acknowledged Reports"
               TabPicture(1)   =   "frmConsultantListView.frx":20E2
               Tab(1).ControlEnabled=   0   'False
               Tab(1).Control(0)=   "AckGrdSid3"
               Tab(1).ControlCount=   1
               Begin MSFlexGridLib.MSFlexGrid grdSid3 
                  Height          =   3012
                  Left            =   60
                  TabIndex        =   63
                  Top             =   300
                  Width           =   4572
                  _ExtentX        =   8070
                  _ExtentY        =   5318
                  _Version        =   393216
                  Cols            =   9
                  BackColor       =   -2147483624
                  ForeColor       =   -2147483635
                  BackColorFixed  =   -2147483647
                  ForeColorFixed  =   -2147483624
                  FocusRect       =   0
                  HighLight       =   0
                  GridLines       =   3
                  GridLinesFixed  =   3
                  FormatString    =   $"frmConsultantListView.frx":20FE
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
               Begin MSFlexGridLib.MSFlexGrid AckGrdSid3 
                  Height          =   3012
                  Left            =   -74940
                  TabIndex        =   64
                  Top             =   300
                  Width           =   4572
                  _ExtentX        =   8070
                  _ExtentY        =   5318
                  _Version        =   393216
                  Cols            =   9
                  BackColor       =   -2147483624
                  ForeColor       =   -2147483635
                  BackColorFixed  =   -2147483647
                  ForeColorFixed  =   -2147483624
                  FocusRect       =   0
                  HighLight       =   0
                  GridLines       =   3
                  GridLinesFixed  =   3
                  FormatString    =   $"frmConsultantListView.frx":21C9
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
         End
         Begin VB.CommandButton cmdRevertToLab3 
            Caption         =   "Revert to Lab"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1200
            Left            =   16320
            Picture         =   "frmConsultantListView.frx":2294
            Style           =   1  'Graphical
            TabIndex        =   52
            Top             =   4740
            Width           =   1200
         End
         Begin VB.CommandButton cmdReleaseReport3 
            Caption         =   "Authorise Report"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1200
            Left            =   13620
            Picture         =   "frmConsultantListView.frx":297E
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   4740
            Width           =   1200
         End
         Begin VB.Frame Frame5 
            BorderStyle     =   0  'None
            Caption         =   "Frame3"
            Height          =   2955
            Left            =   13335
            TabIndex        =   30
            Top             =   6525
            Width           =   4515
            Begin VB.CommandButton cmdConC3 
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
               Height          =   315
               Left            =   4050
               TabIndex        =   31
               ToolTipText     =   "Choose a comment from a list"
               Top             =   1680
               Width           =   435
            End
            Begin VB.CommandButton cmdRefresh 
               Appearance      =   0  'Flat
               Caption         =   "Refresh"
               Height          =   900
               Index           =   2
               Left            =   1080
               Picture         =   "frmConsultantListView.frx":3068
               Style           =   1  'Graphical
               TabIndex        =   43
               Tag             =   "101"
               ToolTipText     =   "Exit Screen"
               Top             =   2040
               Width           =   1000
            End
            Begin VB.CommandButton cmdSaveC3 
               Caption         =   "Save"
               Height          =   900
               Left            =   0
               Picture         =   "frmConsultantListView.frx":3932
               Style           =   1  'Graphical
               TabIndex        =   39
               ToolTipText     =   "Save Changes"
               Top             =   2040
               Width           =   1000
            End
            Begin VB.ComboBox cmbConC3 
               Height          =   315
               Left            =   60
               TabIndex        =   33
               Text            =   "cmbConC"
               Top             =   1680
               Visible         =   0   'False
               Width           =   3975
            End
            Begin VB.TextBox txtConC3 
               Height          =   2000
               Left            =   0
               MultiLine       =   -1  'True
               TabIndex        =   32
               Top             =   0
               Width           =   4515
            End
         End
         Begin VB.TextBox txtPages3 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000011&
            Height          =   425
            Left            =   14850
            Locked          =   -1  'True
            TabIndex        =   21
            Top             =   6030
            Width           =   1455
         End
         Begin VB.CommandButton cmdMove3 
            Caption         =   ">"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   425
            Index           =   0
            Left            =   16365
            TabIndex        =   20
            Top             =   6030
            Width           =   524
         End
         Begin VB.CommandButton cmdMove3 
            Caption         =   ">>"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   425
            Index           =   1
            Left            =   16950
            TabIndex        =   19
            Top             =   6030
            Width           =   524
         End
         Begin VB.CommandButton cmdMove3 
            Caption         =   "<"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   425
            Index           =   2
            Left            =   14265
            TabIndex        =   18
            Top             =   6030
            Width           =   524
         End
         Begin VB.CommandButton cmdMove3 
            Caption         =   "<<"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   425
            Index           =   3
            Left            =   13680
            TabIndex        =   17
            Top             =   6030
            Width           =   524
         End
         Begin RichTextLib.RichTextBox txtReport3 
            Height          =   9195
            Left            =   210
            TabIndex        =   16
            Top             =   300
            Width           =   13005
            _ExtentX        =   22913
            _ExtentY        =   16219
            _Version        =   393217
            ReadOnly        =   -1  'True
            ScrollBars      =   2
            TextRTF         =   $"frmConsultantListView.frx":3C3C
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   9615
         Left            =   -74940
         TabIndex        =   8
         Top             =   360
         Width           =   18000
         Begin VB.CommandButton cmdAck1 
            Caption         =   "Acknowledge"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1200
            Left            =   14982
            Picture         =   "frmConsultantListView.frx":3CBC
            Style           =   1  'Graphical
            TabIndex        =   66
            Top             =   4740
            Width           =   1200
         End
         Begin VB.Frame Frame6 
            Height          =   3612
            Left            =   13200
            TabIndex        =   57
            Top             =   1080
            Width           =   4752
            Begin TabDlg.SSTab SSTab2 
               Height          =   3372
               Left            =   60
               TabIndex        =   58
               Top             =   180
               Width           =   4668
               _ExtentX        =   8229
               _ExtentY        =   5953
               _Version        =   393216
               Tabs            =   2
               TabsPerRow      =   2
               TabHeight       =   420
               TabCaption(0)   =   "Ready for Consultant"
               TabPicture(0)   =   "frmConsultantListView.frx":4586
               Tab(0).ControlEnabled=   -1  'True
               Tab(0).Control(0)=   "grdSid2"
               Tab(0).Control(0).Enabled=   0   'False
               Tab(0).ControlCount=   1
               TabCaption(1)   =   "Acknowledged Reports"
               TabPicture(1)   =   "frmConsultantListView.frx":45A2
               Tab(1).ControlEnabled=   0   'False
               Tab(1).Control(0)=   "AckGrdSid2"
               Tab(1).ControlCount=   1
               Begin MSFlexGridLib.MSFlexGrid grdSid2 
                  Height          =   3012
                  Left            =   60
                  TabIndex        =   59
                  Top             =   300
                  Width           =   4572
                  _ExtentX        =   8070
                  _ExtentY        =   5318
                  _Version        =   393216
                  Cols            =   9
                  BackColor       =   -2147483624
                  ForeColor       =   -2147483635
                  BackColorFixed  =   -2147483647
                  ForeColorFixed  =   -2147483624
                  FocusRect       =   0
                  HighLight       =   0
                  GridLines       =   3
                  GridLinesFixed  =   3
                  FormatString    =   $"frmConsultantListView.frx":45BE
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
               Begin MSFlexGridLib.MSFlexGrid AckGrdSid2 
                  Height          =   3012
                  Left            =   -74940
                  TabIndex        =   60
                  Top             =   300
                  Width           =   4572
                  _ExtentX        =   8070
                  _ExtentY        =   5318
                  _Version        =   393216
                  Cols            =   9
                  BackColor       =   -2147483624
                  ForeColor       =   -2147483635
                  BackColorFixed  =   -2147483647
                  ForeColorFixed  =   -2147483624
                  FocusRect       =   0
                  HighLight       =   0
                  GridLines       =   3
                  GridLinesFixed  =   3
                  FormatString    =   $"frmConsultantListView.frx":4689
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
         End
         Begin VB.CommandButton cmdRevertToLab2 
            Caption         =   "Revert to Lab"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1200
            Left            =   16344
            Picture         =   "frmConsultantListView.frx":4754
            Style           =   1  'Graphical
            TabIndex        =   51
            Top             =   4740
            Width           =   1200
         End
         Begin VB.CommandButton cmdReleaseReport2 
            Caption         =   "Authorise Report"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1200
            Left            =   13620
            Picture         =   "frmConsultantListView.frx":4E3E
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   4740
            Width           =   1200
         End
         Begin VB.Frame Frame4 
            BorderStyle     =   0  'None
            Caption         =   "Frame3"
            Height          =   2955
            Left            =   13335
            TabIndex        =   26
            Top             =   6525
            Width           =   4545
            Begin VB.CommandButton cmdConC2 
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
               Height          =   315
               Left            =   4050
               TabIndex        =   27
               ToolTipText     =   "Choose a comment from a list"
               Top             =   1680
               Width           =   435
            End
            Begin VB.CommandButton cmdRefresh 
               Appearance      =   0  'Flat
               Caption         =   "Refresh"
               Height          =   900
               Index           =   1
               Left            =   1080
               Picture         =   "frmConsultantListView.frx":5528
               Style           =   1  'Graphical
               TabIndex        =   42
               Tag             =   "101"
               ToolTipText     =   "Exit Screen"
               Top             =   2040
               Width           =   1000
            End
            Begin VB.CommandButton cmdSaveC2 
               Caption         =   "Save"
               Height          =   900
               Left            =   0
               Picture         =   "frmConsultantListView.frx":5DF2
               Style           =   1  'Graphical
               TabIndex        =   38
               ToolTipText     =   "Save Changes"
               Top             =   2040
               Width           =   1000
            End
            Begin VB.ComboBox cmbConC2 
               Height          =   315
               Left            =   60
               TabIndex        =   29
               Text            =   "cmbConC"
               Top             =   1680
               Visible         =   0   'False
               Width           =   3975
            End
            Begin VB.TextBox txtConC2 
               Height          =   2000
               Left            =   0
               MultiLine       =   -1  'True
               TabIndex        =   28
               Top             =   0
               Width           =   4515
            End
         End
         Begin RichTextLib.RichTextBox txtReport2 
            Height          =   9195
            Left            =   210
            TabIndex        =   14
            Top             =   300
            Width           =   13005
            _ExtentX        =   22913
            _ExtentY        =   16219
            _Version        =   393217
            ReadOnly        =   -1  'True
            ScrollBars      =   2
            TextRTF         =   $"frmConsultantListView.frx":60FC
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.CommandButton cmdMove2 
            Caption         =   "<<"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   425
            Index           =   3
            Left            =   13680
            TabIndex        =   13
            Top             =   6030
            Width           =   524
         End
         Begin VB.CommandButton cmdMove2 
            Caption         =   "<"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   425
            Index           =   2
            Left            =   14265
            TabIndex        =   12
            Top             =   6030
            Width           =   524
         End
         Begin VB.CommandButton cmdMove2 
            Caption         =   ">>"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   425
            Index           =   1
            Left            =   16950
            TabIndex        =   11
            Top             =   6030
            Width           =   524
         End
         Begin VB.CommandButton cmdMove2 
            Caption         =   ">"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   425
            Index           =   0
            Left            =   16365
            TabIndex        =   10
            Top             =   6030
            Width           =   524
         End
         Begin VB.TextBox txtPages2 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000011&
            Height          =   425
            Left            =   14850
            Locked          =   -1  'True
            TabIndex        =   9
            Top             =   6030
            Width           =   1455
         End
      End
      Begin VB.Frame fraMicro 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   9615
         Left            =   60
         TabIndex        =   1
         Top             =   360
         Width           =   18000
         Begin VB.CommandButton cmdAck 
            Caption         =   "Acknowledge"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1200
            Left            =   14970
            Picture         =   "frmConsultantListView.frx":617C
            Style           =   1  'Graphical
            TabIndex        =   65
            Top             =   4740
            Width           =   1200
         End
         Begin VB.Frame FrmList 
            Height          =   3612
            Left            =   13200
            TabIndex        =   53
            Top             =   1080
            Width           =   4752
            Begin TabDlg.SSTab SSTab 
               Height          =   3372
               Left            =   60
               TabIndex        =   54
               Top             =   180
               Width           =   4668
               _ExtentX        =   8229
               _ExtentY        =   5953
               _Version        =   393216
               Tabs            =   2
               TabsPerRow      =   2
               TabHeight       =   420
               TabCaption(0)   =   "Ready for Consultant"
               TabPicture(0)   =   "frmConsultantListView.frx":6A46
               Tab(0).ControlEnabled=   -1  'True
               Tab(0).Control(0)=   "grdSID"
               Tab(0).Control(0).Enabled=   0   'False
               Tab(0).ControlCount=   1
               TabCaption(1)   =   "Acknowledged by Lab"
               TabPicture(1)   =   "frmConsultantListView.frx":6A62
               Tab(1).ControlEnabled=   0   'False
               Tab(1).Control(0)=   "AckGrdSid"
               Tab(1).ControlCount=   1
               Begin MSFlexGridLib.MSFlexGrid grdSID 
                  Height          =   3012
                  Left            =   60
                  TabIndex        =   55
                  Top             =   288
                  Width           =   4572
                  _ExtentX        =   8070
                  _ExtentY        =   5318
                  _Version        =   393216
                  Cols            =   9
                  BackColor       =   -2147483624
                  ForeColor       =   -2147483635
                  BackColorFixed  =   -2147483647
                  ForeColorFixed  =   -2147483624
                  FocusRect       =   0
                  HighLight       =   0
                  GridLines       =   3
                  GridLinesFixed  =   3
                  FormatString    =   $"frmConsultantListView.frx":6A7E
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
               Begin MSFlexGridLib.MSFlexGrid AckGrdSid 
                  Height          =   3012
                  Left            =   -74940
                  TabIndex        =   56
                  Top             =   300
                  Width           =   4572
                  _ExtentX        =   8070
                  _ExtentY        =   5318
                  _Version        =   393216
                  Cols            =   9
                  BackColor       =   -2147483624
                  ForeColor       =   -2147483635
                  BackColorFixed  =   -2147483647
                  ForeColorFixed  =   -2147483624
                  FocusRect       =   0
                  HighLight       =   0
                  GridLines       =   3
                  GridLinesFixed  =   3
                  FormatString    =   $"frmConsultantListView.frx":6B49
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
         End
         Begin VB.CommandButton cmdRevertToLab 
            Caption         =   "Revert to Lab"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1200
            Left            =   16320
            Picture         =   "frmConsultantListView.frx":6C14
            Style           =   1  'Graphical
            TabIndex        =   50
            Top             =   4740
            Width           =   1200
         End
         Begin VB.CommandButton cmdReleaseReport 
            Caption         =   "Authorise Report"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1200
            Left            =   13620
            Picture         =   "frmConsultantListView.frx":72FE
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   4740
            Width           =   1200
         End
         Begin VB.TextBox txtPages 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000011&
            Height          =   425
            Left            =   14850
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   6030
            Width           =   1455
         End
         Begin VB.CommandButton cmdMove 
            Caption         =   ">"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   425
            Index           =   0
            Left            =   16365
            TabIndex        =   5
            Top             =   6030
            Width           =   524
         End
         Begin VB.CommandButton cmdMove 
            Caption         =   ">>"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   425
            Index           =   1
            Left            =   16950
            TabIndex        =   4
            Top             =   6030
            Width           =   524
         End
         Begin VB.CommandButton cmdMove 
            Caption         =   "<"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   425
            Index           =   2
            Left            =   14265
            TabIndex        =   3
            Top             =   6030
            Width           =   524
         End
         Begin VB.CommandButton cmdMove 
            Caption         =   "<<"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   425
            Index           =   3
            Left            =   13680
            TabIndex        =   2
            Top             =   6030
            Width           =   524
         End
         Begin RichTextLib.RichTextBox txtReport 
            Height          =   9195
            Left            =   210
            TabIndex        =   7
            Top             =   300
            Width           =   13005
            _ExtentX        =   22913
            _ExtentY        =   16219
            _Version        =   393217
            ReadOnly        =   -1  'True
            ScrollBars      =   2
            TextRTF         =   $"frmConsultantListView.frx":79E8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
   End
   Begin VB.Image imgRedCross 
      Height          =   225
      Left            =   0
      Picture         =   "frmConsultantListView.frx":7A68
      Top             =   0
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image imgGreenTick 
      Height          =   225
      Left            =   0
      Picture         =   "frmConsultantListView.frx":7D3E
      Top             =   210
      Visible         =   0   'False
      Width           =   210
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   8910
      Top             =   9600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmConsultantListView.frx":8014
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmConsultantListView.frx":8366
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmConsultantListView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private CurrentGridRow As Integer
Private PreviousGridRow As Integer
Dim tbReports As Recordset



'---------------------------------------------------------------------------------------
' Procedure : cmbConC_KeyPress
' Author    : Babar Shahzad
' Date      : 10/10/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmbConC_KeyPress(KeyAscii As Integer)

10    On Error GoTo cmbConC_KeyPress_Error

20    KeyAscii = AutoComplete(cmbConC, KeyAscii, False)

30    Exit Sub

cmbConC_KeyPress_Error:

       Dim strES As String
       Dim intEL As Integer

40     intEL = Erl
50     strES = Err.Description
60     LogError "frmConsultantListView", "cmbConC_KeyPress", intEL, strES
          
End Sub

Private Sub cmbConC2_KeyPress(KeyAscii As Integer)

10    On Error GoTo cmbConC2_KeyPress_Error

20    KeyAscii = AutoComplete(cmbConC2, KeyAscii, False)

30    Exit Sub

cmbConC2_KeyPress_Error:

       Dim strES As String
       Dim intEL As Integer

40     intEL = Erl
50     strES = Err.Description
60     LogError "frmConsultantListView", "cmbConC2_KeyPress", intEL, strES
          
End Sub

Private Sub cmbConC3_KeyPress(KeyAscii As Integer)

10    On Error GoTo cmbConC3_KeyPress_Error

20    KeyAscii = AutoComplete(cmbConC3, KeyAscii, False)

30    Exit Sub

cmbConC3_KeyPress_Error:

       Dim strES As String
       Dim intEL As Integer

40     intEL = Erl
50     strES = Err.Description
60     LogError "frmConsultantListView", "cmbConC3_KeyPress", intEL, strES
          
End Sub

Private Sub cmdAck_Click()

10    On Error GoTo cmdAck_Click_Error

20    ConAck 0, grdSID
30    FillGrid 0, grdSID
40    FillAckGrid 0, AckGrdSid
50    Exit Sub

cmdAck_Click_Error:

      Dim strES As String
      Dim intEL As Integer

60    intEL = Erl
70    strES = Err.Description
80    LogError "frmConsultantListView", "cmdAck_Click", intEL, strES

End Sub

Private Sub cmdAck1_Click()
10    On Error GoTo cmdAck1_Click_Error

20    ConAck 1, grdSID
30    FillGrid 1, grdSID
40    FillAckGrid 1, AckGrdSid
50    Exit Sub

cmdAck1_Click_Error:

      Dim strES As String
      Dim intEL As Integer

60    intEL = Erl
70    strES = Err.Description
80    LogError "frmConsultantListView", "cmdAck1_Click", intEL, strES
End Sub

Private Sub cmdAck2_Click()
10    On Error GoTo cmdAck2_Click_Error

20    ConAck 2, grdSID
30    FillGrid 2, grdSID
40    FillAckGrid 2, AckGrdSid
50    Exit Sub

cmdAck2_Click_Error:

      Dim strES As String
      Dim intEL As Integer

60    intEL = Erl
70    strES = Err.Description
80    LogError "frmConsultantListView", "cmdAck2_Click", intEL, strES
End Sub

Private Sub cmdDartViewer_Click()

10    On Error GoTo cmdDartViewer_Click_Error

20    If txtDartSampleID = "" Then Exit Sub
30    If Not IsNumeric(txtDartSampleID.Text) Then Exit Sub
40    Shell "C:\Program Files\The PlumTree Group\Dartviewer\Dartviewer.exe " & Format(txtDartSampleID, "000000"), vbNormalFocus

50    Exit Sub

cmdDartViewer_Click_Error:

      Dim strES As String
      Dim intEL As Integer

60    intEL = Erl
70    strES = Err.Description
80    LogError "frmEditMicrobiologyNew", "cmdDartViewer_Click", intEL, strES

End Sub

Sub GetReport(ByVal GridRecordNo As MSFlexGrid, ByVal TextBoxToFill As RichTextBox, ByVal TextPages As TextBox, ConIndex As Integer, txtCommentBox As TextBox)
      Static SortOrder As Boolean
      Dim X As Integer
      Dim Y As Integer
      Dim SampleIDWithOffset As String
      Dim sql As String
      Dim Obs As New Observations



10    On Error GoTo GetReport_Error


20    TextBoxToFill.Text = ""
30    TextBoxToFill.TextRTF = ""
40    TextBoxToFill.SelText = ""
50    txtCommentBox = ""
60    SampleIDWithOffset = Val(GridRecordNo.TextMatrix(GridRecordNo.row, 0)) + SysOptMicroOffset(0)

      'If GridRecordNo.MouseRow = 0 Then
      '    If SortOrder Then
      '        GridRecordNo.Sort = flexSortGenericAscending
      '    Else
      '        GridRecordNo.Sort = flexSortGenericDescending
      '    End If
      '    SortOrder = Not SortOrder
      '    Exit Sub
      'End If

      'For Y = 1 To GridRecordNo.Rows - 1
      '    GridRecordNo.Row = Y
      '    For X = 1 To GridRecordNo.Cols - 3
      '        GridRecordNo.Col = X
      '        GridRecordNo.CellBackColor = 0
      '    Next
      'Next

      'GridRecordNo.Row = GridRecordNo.MouseRow
      'For X = 1 To GridRecordNo.Cols - 3
      '    GridRecordNo.Col = X
      '    GridRecordNo.CellBackColor = vbYellow
      'Next
70    If GridRecordNo.TextMatrix(GridRecordNo.row, 0) <> "" Then
          '    MicroSamp = SampleIDWithOffset
80        sql = "SELECT PageNumber, Report FROM UnauthorisedReports WHERE SampleID = '" & SampleIDWithOffset & "' " & _
                "AND SUBSTRING(RepNo,2,LEN(RepNo)) = " & _
                "(SELECT TOP 1 SUBSTRING(RepNo,2,LEN(RepNo)) as RepNo FROM UnauthorisedReports " & _
                "WHERE  Sampleid = '" & SampleIDWithOffset & "' ORDER BY PrintTime DESC)" & _
                "ORDER BY PageNumber"
90        Set tbReports = New Recordset
100       RecOpenClient Val(ConIndex), tbReports, sql
110       If Not tbReports.EOF Then
120           TextPages = "Page " & tbReports.AbsolutePosition & " of " & tbReports.recordCount
130           TextBoxToFill.SelText = tbReports!Report
140           TextBoxToFill.Tag = SampleIDWithOffset
150       End If
160   End If

170   Set Obs = Obs.Load(TextBoxToFill.Tag, "MicroConsultant")
180   If Not Obs Is Nothing Then
190       If Obs.Count > 0 Then
200           txtCommentBox = Obs.Item(1).Comment
210       End If
220   End If

230   TextBoxToFill.SelStart = Len(TextBoxToFill.Text)

240   Exit Sub


GetReport_Error:

      Dim strES As String
      Dim intEL As Integer

250   intEL = Erl
260   strES = Err.Description
270   LogError "frmConsultantListView", "GetReport", intEL, strES, sql
End Sub

Private Function SaveComents(ByVal SampleID As String, ByVal OverWrite As Boolean, Discipline As String, ConIndex As Integer, Comment As String)

      Dim sql As String
      Dim n As Integer

10    On Error GoTo Save_Error

20    If SampleID = "" Then
30        Exit Function
40    End If

50    Comment = RemoveLeadingCrLf(Comment)
60    If Comment = "" Then
70        sql = "DELETE FROM Observations " & _
                "WHERE SampleID = '" & SampleID & "' " & _
                "AND Discipline = '" & Discipline & "'"
80    Else
90        sql = "IF EXISTS (SELECT * FROM Observations " & _
                "WHERE SampleID = '" & SampleID & "' " & _
                "AND Discipline = '" & Discipline & "') " & _
                "  UPDATE Observations "
100       If OverWrite Then
110           sql = sql & "  SET Comment = '" & Comment & "' "
120       Else
130           sql = sql & "  SET Comment = Comment + ' " & Comment & "' "
140       End If
150       sql = sql & "  WHERE SampleID = '" & SampleID & "' " & _
                "  AND Discipline = '" & Discipline & "' " & _
                "ELSE " & _
                "  INSERT INTO Observations " & _
                "  (SampleID, Discipline, Comment, UserName, DateTimeOfRecord ) " & _
                "  VALUES " & _
                "  ('" & SampleID & "', " & _
                "   '" & Discipline & "', " & _
                "   '" & Comment & "', " & _
                "   '" & AddTicks(UserName) & "', " & _
                "   '" & Format$(Now, "yyyy/MM/dd HH:nn:ss") & "')"
160   End If
170   Cnxn(ConIndex).Execute sql

180   Exit Function

Save_Error:

      Dim strES As String
      Dim intEL As Integer

190   intEL = Erl
200   strES = Err.Description
210   LogError "Observations", "Save", intEL, strES, sql


End Function

Private Sub PrintThis(PrintAction As String, SampleID As String, ConIndex As Integer)

      Dim tb As Recordset
      Dim sql As String
      Dim FinalOrInterim As String
      Dim Ward As String
      Dim GP As String
      Dim Clin As String

10    On Error GoTo PrintThis_Error


20    sql = "SELECT * FROM Demographics WHERE SampleID = " & SampleID
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql
50    If Not tb.EOF Then
60        Ward = tb!Ward & ""
70        GP = tb!GP & ""
80        Clin = tb!Clinician & ""
90    End If
      'pBar = 0
      'GetSampleIDWithOffset
      'If Not EntriesOK(txtReport.Tag, txtName, txtSex, cmbWard.Text, cmbGP.Text) Then
      '    Exit Sub
      'End If

      'If Not CheckTimes() Then Exit Sub

      'SaveDemographics

100   sql = "Select * from PrintPending where " & _
            "Department = 'N' " & _
            "and PrintAction = '" & PrintAction & "' " & _
            "and SampleID = '" & SampleID & "'"
110   Set tb = New Recordset
120   RecOpenClient ConIndex, tb, sql
130   If tb.EOF Then
140       tb.AddNew
150   End If
160   tb!SampleID = SampleID
170   tb!Ward = Ward
180   tb!Clinician = Clin
190   tb!GP = GP
200   tb!Department = "N"
210   tb!Initiator = UserName
220   tb!UsePrinter = ""
230   tb!NoOfCopies = Val(1)
240   FinalOrInterim = "F"
250   If PrintAction = "SaveTemp" Then
260       FinalOrInterim = "I"
270   End If
280   tb!FinalInterim = FinalOrInterim
290   tb!pTime = Now
300   tb!PrintAction = PrintAction
310   tb.Update

320   Exit Sub

PrintThis_Error:

      Dim strES As String
      Dim intEL As Integer

330   intEL = Erl
340   strES = Err.Description
350   LogError "frmEditMicrobiologyNew", "PrintThis", intEL, strES, sql

End Sub

Private Sub FillGrid(ConIndex As String, ByVal GridtoFill As MSFlexGrid)

      Dim sql As String
      Dim tb As Recordset
      Dim s As String
      Dim PatInfo As String
      Dim strType As String
      Dim ShowFaecesWardEnq As Boolean
      Dim ShowUrineWardEnq As Boolean
      Dim ShowBloodCultureWardEnq As Boolean
      Dim ShowSwabWardEnq As Boolean

10    On Error GoTo FillGrid_Error
20    Select Case ConIndex
      Case 0:
30        txtReport = ""
40    Case 1:
50        txtReport2 = ""
60    Case 2:
70        txtReport3 = ""
80    End Select
      'grdSID
90    With GridtoFill
100       .Rows = 2
110       .AddItem ""
120       .RemoveItem 1
          '.FormatString = "<SampleID   |<Run Date     |<Sample Date         |<Pat Name                               |<DOB             |<Age   |<Sex  |<Address                                                      |<   |<   "
130       .ColWidth(0) = 1250
140       .ColWidth(1) = 1200
150       .ColWidth(2) = 1600
160       .ColWidth(3) = 0
170       .ColWidth(4) = 0
180       .ColWidth(5) = 0
190       .ColWidth(6) = 0
200       .ColWidth(7) = 0
210       .ColWidth(8) = 0

220   End With

      'sql = "Select * from Demographics "
      'sql = sql & " Order by RunDate desc"

230   sql = "SELECT     D.SampleID,D.SampleDate,D.Rundate ,D.PatName,D.DoB,D.Age,D.Sex,D.Addr0 "
240   sql = sql & " FROM ConsultantList as C INNER JOIN Demographics as D ON D.SampleID = C.SampleID"
250   sql = sql & " Where ISNULL(C.Status,0)=0 "
260   sql = sql & " Order by C.DateTimeOfRecord "

270   Set tb = New Recordset
280   RecOpenClient Val(ConIndex), tb, sql

290   Do While Not tb.EOF

300       PatInfo = tb!PatName & vbTab & tb!Dob & vbTab & tb!Age & vbTab & tb!sex & vbTab & tb!Addr0 & vbTab

310       If Val(tb!SampleID & "") > SysOptMicroOffset(0) Then
320           s = Format$(Val(tb!SampleID) - SysOptMicroOffset(0)) & vbTab & _
                  tb!Rundate & vbTab
330       Else
340           s = Val(tb!SampleID) & vbTab & _
                  tb!Rundate & vbTab
350       End If

360       If IsDate(tb!SampleDate) Then
370           If Format(tb!SampleDate, "hh:mm") <> "00:00" Then
380               s = s & Format(tb!SampleDate, "dd/MM/yyyy hh:mm")
390           Else
400               s = s & Format(tb!SampleDate, "dd/MM/yyyy")
410           End If
420       Else
430           s = s & "Not Specified"
440       End If
450       s = s & vbTab & PatInfo

460       strType = LoadOutstandingMicro(tb!SampleID, Val(ConIndex))
          '        s = s & strType


470       GridtoFill.AddItem s

480       tb.MoveNext
490   Loop

500   If GridtoFill.Rows > 2 Then
510       GridtoFill.RemoveItem 1
520       Call ShowSignals(GridtoFill, ConIndex)
530       GridtoFill.row = 1
540       CurrentGridRow = GridtoFill.row
550       PreviousGridRow = GridtoFill.row
560       MarkFlexGridRow GridtoFill, CurrentGridRow, 1, 7, vbYellow

570       txtReport.Tag = GridtoFill.TextMatrix(GridtoFill.row, 0)
580       Call GetReport(GridtoFill, txtReport, txtPages, CInt(ConIndex), txtConC)
590   End If


600   Exit Sub

FillGrid_Error:

      Dim strES As String
      Dim intEL As Integer

610   intEL = Erl
620   strES = Err.Description
630   LogError "frmViewResults", "FillGrid", intEL, strES, sql


End Sub
Private Sub FillAckGrid(ConIndex As String, ByVal GridtoFill As MSFlexGrid)

      Dim sql As String
      Dim tb As Recordset
      Dim s As String
      Dim PatInfo As String
      Dim strType As String
      Dim ShowFaecesWardEnq As Boolean
      Dim ShowUrineWardEnq As Boolean
      Dim ShowBloodCultureWardEnq As Boolean
      Dim ShowSwabWardEnq As Boolean

10    On Error GoTo FillGrid_Error
      'Select Case ConIndex
      'Case 0:
      '    txtReport = ""
      'Case 1:
      '    txtReport2 = ""
      'Case 2:
      '    txtReport3 = ""
      'End Select
      'grdSID
20    With GridtoFill
30        .Rows = 2
40        .AddItem ""
50        .RemoveItem 1
          '.FormatString = "<SampleID   |<Run Date     |<Sample Date         |<Pat Name                               |<DOB             |<Age   |<Sex  |<Address                                                      |<   |<   "
60        .ColWidth(0) = 1250
70        .ColWidth(1) = 1200
80        .ColWidth(2) = 1600
90        .ColWidth(3) = 0
100       .ColWidth(4) = 0
110       .ColWidth(5) = 0
120       .ColWidth(6) = 0
130       .ColWidth(7) = 0
140       .ColWidth(8) = 0

150   End With

      'sql = "Select * from Demographics "
      'sql = sql & " Order by RunDate desc"

160   sql = "SELECT     D.SampleID,D.SampleDate,D.Rundate ,D.PatName,D.DoB,D.Age,D.Sex,D.Addr0 "
170   sql = sql & " FROM ConsultantList as C INNER JOIN Demographics as D ON D.SampleID = C.SampleID"
180   sql = sql & " Where ISNULL(C.Status,0)=2 and isnull(c.ack,0)=1 "
190   sql = sql & " Order by C.DateTimeOfRecord "

200   Set tb = New Recordset
210   RecOpenClient Val(ConIndex), tb, sql

220   Do While Not tb.EOF

230       PatInfo = tb!PatName & vbTab & tb!Dob & vbTab & tb!Age & vbTab & tb!sex & vbTab & tb!Addr0 & vbTab

240       If Val(tb!SampleID & "") > SysOptMicroOffset(0) Then
250           s = Format$(Val(tb!SampleID) - SysOptMicroOffset(0)) & vbTab & _
                  tb!Rundate & vbTab
260       Else
270           s = Val(tb!SampleID) & vbTab & _
                  tb!Rundate & vbTab
280       End If

290       If IsDate(tb!SampleDate) Then
300           If Format(tb!SampleDate, "hh:mm") <> "00:00" Then
310               s = s & Format(tb!SampleDate, "dd/MM/yyyy hh:mm")
320           Else
330               s = s & Format(tb!SampleDate, "dd/MM/yyyy")
340           End If
350       Else
360           s = s & "Not Specified"
370       End If
380       s = s & vbTab & PatInfo

390       strType = LoadOutstandingMicro(tb!SampleID, Val(ConIndex))
          '        s = s & strType


400       GridtoFill.AddItem s

410       tb.MoveNext
420   Loop

430   If GridtoFill.Rows > 2 Then
440       GridtoFill.RemoveItem 1
450       Call ShowSignals(GridtoFill, ConIndex)
460       GridtoFill.row = 1
470       CurrentGridRow = GridtoFill.row
480       PreviousGridRow = GridtoFill.row
490       MarkFlexGridRow GridtoFill, CurrentGridRow, 1, 7, vbYellow

      '    txtReport.Tag = GridtoFill.TextMatrix(GridtoFill.Row, 0)
      '    Call GetReport(GridtoFill, txtReport, txtPages, CInt(ConIndex), txtConC)
500   End If


510   Exit Sub

FillGrid_Error:

      Dim strES As String
      Dim intEL As Integer

520   intEL = Erl
530   strES = Err.Description
540   LogError "frmViewResults", "FillGrid", intEL, strES, sql


End Sub
'---------------------------------------------------------------------------------------
' Procedure : AuthenticateConsultantList
' Author    : XPMUser
' Date      : 3/12/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'

Private Sub RevertToLab(ConIndex As Integer, grid As MSFlexGrid)
      Dim i As Integer
      Dim sql As String
      Dim SampleIDWithOffset As String
10    On Error GoTo RevertToLab_Error

20    With grid
30        SampleIDWithOffset = Val(.TextMatrix(.row, 0)) + SysOptMicroOffset(0)
40        sql = "update ConsultantList set Status = 2 where sampleid ='" & SampleIDWithOffset & "'"
50        Cnxn(ConIndex).Execute sql
60        RemoveReport Val(ConIndex), SampleIDWithOffset, "N", 0
70    End With
80    Exit Sub

RevertToLab_Error:

      Dim strES As String
      Dim intEL As Integer

90    intEL = Erl
100   strES = Err.Description
110   LogError "frmConsultantListView", "RevertToLab", intEL, strES, sql

End Sub
Private Sub ConAck(ConIndex As Integer, grid As MSFlexGrid)
      Dim i As Integer
      Dim sql As String
      Dim SampleIDWithOffset As String
10    On Error GoTo ConAck_Error

20    With grid
30        SampleIDWithOffset = Val(.TextMatrix(.row, 0)) + SysOptMicroOffset(0)
40        sql = "update ConsultantList set ConAck = 1 where sampleid ='" & SampleIDWithOffset & "'"
50        Cnxn(ConIndex).Execute sql
          'RemoveReport Val(ConIndex), SampleIDWithOffset, "N", 0
60    End With
70    Exit Sub

ConAck_Error:

      Dim strES As String
      Dim intEL As Integer

80    intEL = Erl
90    strES = Err.Description
100   LogError "frmConsultantListView", "ConAck", intEL, strES, sql

End Sub
Private Function GetConAck(ConIndex As Integer, grid As MSFlexGrid) As Boolean
      Dim i As Integer
      Dim sql As String
      Dim tb As ADODB.Recordset
      Dim SampleIDWithOffset As String
10    On Error GoTo getConAck_Error
20    GetConAck = False
30    With grid
40        SampleIDWithOffset = Val(.TextMatrix(.row, 0)) + SysOptMicroOffset(0)
50        sql = "Select ConAck from ConsultantList where sampleid ='" & SampleIDWithOffset & "'"
60        Set tb = New Recordset
70        RecOpenServer 0, tb, sql
          'RemoveReport Val(ConIndex), SampleIDWithOffset, "N", 0
80    End With
90    If tb.EOF Or IsNull(tb!ConAck) Then
100   Else
110      GetConAck = tb!ConAck
120   End If
130   Exit Function

getConAck_Error:

      Dim strES As String
      Dim intEL As Integer

140   intEL = Erl
150   strES = Err.Description
160   LogError "frmConsultantListView", "getConAck", intEL, strES, sql

End Function
Private Sub AuthenticateConsultantList(ConIndex As Integer, grid As MSFlexGrid)

      Dim i As Integer
      Dim sql As String
      Dim SampleIDWithOffset As String

10    On Error GoTo AuthenticateConsultantList_Error
20    If grid.TextMatrix(grid.row, 0) = "" Then Exit Sub
30    With grid
40        SampleIDWithOffset = Val(.TextMatrix(.row, 0)) + SysOptMicroOffset(0)
50        sql = "update ConsultantList set Status = 1, Username = '" & UserName & "' where sampleid ='" & SampleIDWithOffset & "'"
60        Cnxn(ConIndex).Execute sql
70        RemoveReport Val(ConIndex), SampleIDWithOffset, "N", 0
80        Call PrintThis("PrintSaveFinal", SampleIDWithOffset, ConIndex)
          '        For i = 1 To .Rows - 1
          '                 .Row = i
          '                 .Col = 9
          '                If .CellPicture = imgGreenTick Then
          '                    SampleIDWithOffset = Val(.TextMatrix(i, 0)) + SysOptMicroOffset(0)
          '                    sql = "update ConsultantList set Status = 1 where sampleid ='" & SampleIDWithOffset & "'"
          '                    Cnxn(ConIndex).Execute sql
          '                    RemoveReport SampleIDWithOffset, "N", 0
          '                    Call PrintThis("PrintSaveFinal", SampleIDWithOffset)
          '                End If
          '            Next i
90    End With

100   Exit Sub

AuthenticateConsultantList_Error:

      Dim strES As String
      Dim intEL As Integer

110   intEL = Erl
120   strES = Err.Description
130   LogError "frmConsultantListView", "AuthenticateConsultantList", intEL, strES, sql
End Sub

'---------------------------------------------------------------------------------------
' Procedure : IsReportReadytoShow
' Author    : XPMUser
' Date      : 3/6/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub ShowSignals(ByVal GridtoFill As MSFlexGrid, ConIndex As String)

10    On Error GoTo IsReportReadytoShow_Error
      Dim SampleID As String
      Dim SampleIDWithOffset As String
      Dim i As Integer

20    With GridtoFill
30        For i = 1 To .Rows - 1
40            .row = i
50            .Col = 8
60            SampleIDWithOffset = Val(.TextMatrix(i, 0)) + SysOptMicroOffset(0)
70            If IsReportReady(SampleIDWithOffset, ConIndex) = True Then
80                .CellBackColor = vbGreen
                  'Set .CellPicture = ImageList1.ListImages(1).Picture

                  '    .CellAlignment = flexAlignCenterCenter
90            Else
100               .CellBackColor = vbRed
                  '                    Set .CellPicture = ImageList1.ListImages(2).Picture
                  '                        .CellAlignment = flexAlignCenterCenter
110           End If




120       Next i

130   End With


140   Exit Sub


IsReportReadytoShow_Error:

      Dim strES As String
      Dim intEL As Integer

150   intEL = Erl
160   strES = Err.Description
170   LogError "frmConsultantListView", "IsReportReadytoShow", intEL, strES
End Sub

'---------------------------------------------------------------------------------------
' Procedure : IsReportReady
' Author    : XPMUser
' Date      : 3/6/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Function IsReportReady(SampleID As String, ConIndex As String) As Boolean

10    On Error GoTo IsReportReady_Error
      Dim sql As String
      Dim tb As ADODB.Recordset
20    IsReportReady = False

30    sql = "select * from UnauthorisedReports where sampleid ='" & SampleID & "'"
40    Set tb = New Recordset
50    RecOpenClient ConIndex, tb, sql
60    If Not tb.EOF Then
70        IsReportReady = True
80    End If


90    Exit Function


IsReportReady_Error:

      Dim strES As String
      Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "frmConsultantListView", "IsReportReady", intEL, strES
End Function

Private Function LoadOutstandingMicro(ByVal SampleIDWithOffset As String, ConIndex As String) As String

      Dim tb As Recordset
      Dim sql As String
      Dim s As String
      Dim n As Integer

10    On Error GoTo LoadOutstandingMicro_Error

20    sql = "Select * from MicroSiteDetails where " & _
            "SampleID = '" & SampleIDWithOffset & "' "
30    Set tb = New Recordset
40    RecOpenServer Val(ConIndex), tb, sql

50    If Not tb.EOF Then

60        If UCase(Trim$(tb!Site & "")) = "FAECES" Then
70            LoadOutstandingMicro = UCase(Trim$(tb!Site & ""))
80            Exit Function
90        End If

100       s = tb!Site & " " & tb!SiteDetails & " "
110       If tb!Site & "" = "Urine" Or tb!Site & "" = "Faeces" Then
120           sql = "Select * from MicroRequests where " & _
                    "SampleID = '" & SampleIDWithOffset & "'"
130           Set tb = New Recordset
140           RecOpenServer Val(ConIndex), tb, sql

150           If Not tb.EOF Then

160               For n = 0 To 2
170                   If tb!Faecal And 2 ^ n Then
180                       s = s & Choose(n + 1, "C & S ", "C. Difficile ", "O/P ")
190                   End If
200               Next

210               For n = 3 To 5
220                   If tb!Faecal And 2 ^ n Then
230                       s = s & "Occult Blood "
240                       Exit For
250                   End If
260               Next

270               If tb!Faecal And 2 ^ 6 Then
280                   s = s & "Rota/Adeno "
290               End If

300               For n = 7 To 10
310                   If tb!Faecal And 2 ^ n Then
320                       s = s & Choose(n + 1, "Toxin A ", "Coli 0157 ", _
                                         "E/P Coli ", "S/S Screen ")
330                   End If
340               Next

350               For n = 0 To 5
360                   If tb!Urine And 2 ^ n Then
370                       s = s & Choose(n + 1, "C & S", "Pregnancy ", "Fat Globules ", _
                                         "Bence Jones ", "SG ", "HCG ")
380                   End If
390               Next
400           End If
410       End If
420   End If
430   LoadOutstandingMicro = Trim(s)

440   Exit Function

LoadOutstandingMicro_Error:

      Dim strES As String
      Dim intEL As Integer

450   intEL = Erl
460   strES = Err.Description
470   LogError "frmViewResults", "LoadOutstandingMicro", intEL, strES, sql


End Function

Private Sub FillMSandConsultantComment()

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo FillMSandConsultantComment_Error

20    cmbConC.Clear

30    sql = "Select * from Lists where " & _
            "ListType = 'ConsComment' " & _
            "ORDER BY ListOrder"
40    Set tb = New Recordset
50    RecOpenServer 0, tb, sql
60    Do While Not tb.EOF
70        cmbConC.AddItem tb!Text & ""
80        cmbConC2.AddItem tb!Text & ""
90        cmbConC3.AddItem tb!Text & ""
100       tb.MoveNext
110   Loop

120   Exit Sub

FillMSandConsultantComment_Error:

      Dim strES As String
      Dim intEL As Integer

130   intEL = Erl
140   strES = Err.Description
150   LogError "frmEditMicrobiologyNew", "FillMSandConsultantComment", intEL, strES, sql


End Sub

Private Sub bcancel_Click()

10    On Error GoTo bCancel_Click_Error

20    Unload Me

30    Exit Sub

bCancel_Click_Error:

      Dim strES As String
      Dim intEL As Integer

40    intEL = Erl
50    strES = Err.Description
60    LogError "frmConsultantListView", "bCancel_Click", intEL, strES

End Sub

Private Sub btnCancel_Click()
10    btnCancel.Tag = "C"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdMove_Click
' Author    : XPMUser
' Date      : 3/5/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmdMove_Click(Index As Integer)


10    On Error GoTo cmdMove_Click_Error



20    With tbReports
30        If .State = adStateClosed Then Exit Sub

40        Select Case Index
          Case 0:
50            .MoveNext
60            If .EOF Then
70                .MoveLast
80                Exit Sub
90            End If
100           txtReport.Text = ""
110           txtPages = "Page " & tbReports.AbsolutePosition & " of " & tbReports.recordCount
120           txtReport.SelText = tbReports!Report
130       Case 1:
140           If .AbsolutePosition = .recordCount Then Exit Sub
150           .MoveLast
160           txtReport.Text = ""
170           txtPages = "Page " & tbReports.AbsolutePosition & " of " & tbReports.recordCount
180           txtReport.SelText = tbReports!Report
190       Case 2:
200           .MovePrevious
210           If .BOF Then
220               .MoveFirst
230               Exit Sub
240           End If
250           txtReport.Text = ""
260           txtPages = "Page " & tbReports.AbsolutePosition & " of " & tbReports.recordCount
270           txtReport.SelText = tbReports!Report
280       Case 3:
290           If .AbsolutePosition = 1 Then Exit Sub
300           .MoveFirst
310           txtReport.Text = ""
320           txtPages = "Page " & tbReports.AbsolutePosition & " of " & tbReports.recordCount
330           txtReport.SelText = tbReports!Report

340       End Select
350   End With
360   Exit Sub


370   Exit Sub


cmdMove_Click_Error:

      Dim strES As String
      Dim intEL As Integer

380   intEL = Erl
390   strES = Err.Description
400   LogError "frmConsultantListView", "cmdMove_Click", intEL, strES
End Sub

Private Sub cmdMove2_Click(Index As Integer)

10    On Error GoTo cmdMove2_Click_Error


20    With tbReports
30        If .State = adStateClosed Then Exit Sub

40        Select Case Index
          Case 0:
50            .MoveNext
60            If .EOF Then
70                .MoveLast
80                Exit Sub
90            End If
100           txtReport2.Text = ""
110           txtPages2 = "Page " & tbReports.AbsolutePosition & " of " & tbReports.recordCount
120           txtReport2.SelText = tbReports!Report
130       Case 1:
140           If .AbsolutePosition = .recordCount Then Exit Sub
150           .MoveLast
160           txtReport2.Text = ""
170           txtPages2 = "Page " & tbReports.AbsolutePosition & " of " & tbReports.recordCount
180           txtReport2.SelText = tbReports!Report
190       Case 2:
200           .MovePrevious
210           If .BOF Then
220               .MoveFirst
230               Exit Sub
240           End If
250           txtReport2.Text = ""
260           txtPages2 = "Page " & tbReports.AbsolutePosition & " of " & tbReports.recordCount
270           txtReport2.SelText = tbReports!Report
280       Case 3:
290           If .AbsolutePosition = 1 Then Exit Sub
300           .MoveFirst
310           txtReport2.Text = ""
320           txtPages2 = "Page " & tbReports.AbsolutePosition & " of " & tbReports.recordCount
330           txtReport2.SelText = tbReports!Report

340       End Select
350   End With

360   Exit Sub

cmdMove2_Click_Error:

      Dim strES As String
      Dim intEL As Integer

370   intEL = Erl
380   strES = Err.Description
390   LogError "frmConsultantListView", "cmdMove2_Click", intEL, strES

End Sub

Private Sub cmdMove3_Click(Index As Integer)

10    On Error GoTo cmdMove3_Click_Error


20    With tbReports
30        If .State = adStateClosed Then Exit Sub

40        Select Case Index
          Case 0:
50            .MoveNext
60            If .EOF Then
70                .MoveLast
80                Exit Sub
90            End If
100           txtReport3.Text = ""
110           txtPages3 = "Page " & tbReports.AbsolutePosition & " of " & tbReports.recordCount
120           txtReport3.SelText = tbReports!Report
130       Case 1:
140           If .AbsolutePosition = .recordCount Then Exit Sub
150           .MoveLast
160           txtReport3.Text = ""
170           txtPages3 = "Page " & tbReports.AbsolutePosition & " of " & tbReports.recordCount
180           txtReport3.SelText = tbReports!Report
190       Case 2:
200           .MovePrevious
210           If .BOF Then
220               .MoveFirst
230               Exit Sub
240           End If
250           txtReport3.Text = ""
260           txtPages3 = "Page " & tbReports.AbsolutePosition & " of " & tbReports.recordCount
270           txtReport3.SelText = tbReports!Report
280       Case 3:
290           If .AbsolutePosition = 1 Then Exit Sub
300           .MoveFirst
310           txtReport3.Text = ""
320           txtPages3 = "Page " & tbReports.AbsolutePosition & " of " & tbReports.recordCount
330           txtReport3.SelText = tbReports!Report

340       End Select
350   End With


360   Exit Sub

cmdMove3_Click_Error:

      Dim strES As String
      Dim intEL As Integer

370   intEL = Erl
380   strES = Err.Description
390   LogError "frmConsultantListView", "cmdMove3_Click", intEL, strES

End Sub

Private Sub cmdRefresh_Click(Index As Integer)


10    Select Case Index
      Case 0:
20        FillGrid 0, grdSID
30        FillAckGrid 0, AckGrdSid
40    Case 1:
50        FillGrid 1, grdSid2
60        FillAckGrid 1, AckGrdSid2
70    Case 2:
80        FillGrid 2, grdSid3
90        FillAckGrid 2, AckGrdSid3
100   End Select

End Sub


Private Sub cmdReleaseReport2_Click()

10    On Error GoTo cmdReleaseReport2_Click_Error

20    Call AuthenticateConsultantList(1, grdSid2)
30    Call FillGrid(1, grdSid2)

40    Exit Sub

cmdReleaseReport2_Click_Error:

      Dim strES As String
      Dim intEL As Integer

50    intEL = Erl
60    strES = Err.Description
70    LogError "frmConsultantListView", "cmdReleaseReport2_Click", intEL, strES

End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdReleaseReport_Click
' Author    : XPMUser
' Date      : 3/12/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmdReleaseReport_Click()

      Dim SID As String

10    On Error GoTo cmdReleaseReport_Click_Error

20    Call AuthenticateConsultantList(0, grdSID)
30    SID = Format$(Val(Val(grdSID.TextMatrix(grdSID.row, 0))) + SysOptMicroOffset(0))
40    ReleaseMicro SID, 1
50    Call FillGrid(0, grdSID)




60    Exit Sub


cmdReleaseReport_Click_Error:

      Dim strES As String
      Dim intEL As Integer

70    intEL = Erl
80    strES = Err.Description
90    LogError "frmConsultantListView", "cmdReleaseReport_Click", intEL, strES
End Sub

Private Sub cmbConC_LostFocus()

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo cmbConC_LostFocus_Error

20    cmbConC.Text = QueryCombo(cmbConC)
30    If cmbConC <> "" Then
40        sql = "SELECT * FROM Lists WHERE " & _
                "ListType = 'ConsComment' " & _
                "AND Code = '" & cmbConC & "'"
50        Set tb = New Recordset
60        RecOpenServer 0, tb, sql
70        If Not tb.EOF Then
80            cmbConC = tb!Text & ""
90        End If
100   End If

110   If txtConC() = "" Then
120       txtConC = cmbConC
130   Else
140       txtConC = txtConC & cmbConC
150   End If

160   cmbConC.Visible = False
170   cmbConC = ""

180   Exit Sub

cmbConC_LostFocus_Error:

      Dim strES As String
      Dim intEL As Integer

190   intEL = Erl
200   strES = Err.Description
210   LogError "frmEditMicrobiologyNew", "cmbConC_LostFocus", intEL, strES, sql


End Sub

Private Sub cmbConC2_LostFocus()

      Dim sql As String
      Dim tb As Recordset

10    On Error GoTo cmbConC2_LostFocus_Error

20    cmbConC2.Text = QueryCombo(cmbConC2)
30    If cmbConC2 <> "" Then
40        sql = "SELECT * FROM Lists WHERE " & _
                "ListType = 'ConsComment' " & _
                "AND Code = '" & cmbConC2 & "'"
50        Set tb = New Recordset
60        RecOpenServer 0, tb, sql
70        If Not tb.EOF Then
80            cmbConC2 = tb!Text & ""
90        End If
100   End If

110   If txtConC2 = "" Then
120       txtConC2 = cmbConC2
130   Else
140       txtConC2 = txtConC2 & cmbConC2
150   End If

160   cmbConC2.Visible = False
170   cmbConC2 = ""


180   Exit Sub

cmbConC2_LostFocus_Error:

      Dim strES As String
      Dim intEL As Integer

190   intEL = Erl
200   strES = Err.Description
210   LogError "frmConsultantListView", "cmbConC2_LostFocus", intEL, strES

End Sub

Private Sub cmbConC3_LostFocus()

      Dim sql As String
      Dim tb As Recordset

10    On Error GoTo cmbConC3_LostFocus_Error

20    cmbConC3.Text = QueryCombo(cmbConC3)
30    If cmbConC3 <> "" Then
40        sql = "SELECT * FROM Lists WHERE " & _
                "ListType = 'ConsComment' " & _
                "AND Code = '" & cmbConC3 & "'"
50        Set tb = New Recordset
60        RecOpenServer 0, tb, sql
70        If Not tb.EOF Then
80            cmbConC3 = tb!Text & ""
90        End If
100   End If

110   If txtConC3 = "" Then
120       txtConC3 = cmbConC3
130   Else
140       txtConC3 = txtConC3 & cmbConC3
150   End If

160   cmbConC3.Visible = False
170   cmbConC3 = ""


180   Exit Sub

cmbConC3_LostFocus_Error:

      Dim strES As String
      Dim intEL As Integer

190   intEL = Erl
200   strES = Err.Description
210   LogError "frmConsultantListView", "cmbConC3_LostFocus", intEL, strES

End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdConC_Click
' Author    : XPMUser
' Date      : 3/6/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmdConC_Click()
10    On Error GoTo cmdConC_Click_Error


20    cmbConC.Visible = True
30    cmbConC.SetFocus


40    Exit Sub


cmdConC_Click_Error:

      Dim strES As String
      Dim intEL As Integer

50    intEL = Erl
60    strES = Err.Description
70    LogError "frmConsultantListView", "cmdConC_Click", intEL, strES
End Sub

Private Sub cmdConC2_Click()

10    On Error GoTo cmdConC2_Click_Error

20    cmbConC2.Visible = True
30    cmbConC2.SetFocus

40    Exit Sub

cmdConC2_Click_Error:

      Dim strES As String
      Dim intEL As Integer

50    intEL = Erl
60    strES = Err.Description
70    LogError "frmConsultantListView", "cmdConC2_Click", intEL, strES

End Sub

Private Sub cmdConC3_Click()

10    On Error GoTo cmdConC3_Click_Error

20    cmbConC3.Visible = True
30    cmbConC3.SetFocus

40    Exit Sub

cmdConC3_Click_Error:

      Dim strES As String
      Dim intEL As Integer

50    intEL = Erl
60    strES = Err.Description
70    LogError "frmConsultantListView", "cmdConC3_Click", intEL, strES

End Sub

Private Sub cmdRevertToLab_Click()

10    On Error GoTo cmdRevertToLab_Click_Error

20    RevertToLab 0, grdSID
30    FillGrid 0, grdSID
40    FillAckGrid 0, AckGrdSid
50    Exit Sub

cmdRevertToLab_Click_Error:

      Dim strES As String
      Dim intEL As Integer

60    intEL = Erl
70    strES = Err.Description
80    LogError "frmConsultantListView", "cmdRevertToLab_Click", intEL, strES

End Sub

Private Sub cmdRevertToLab2_Click()

10    On Error GoTo cmdRevertToLab2_Click_Error

20    RevertToLab 1, grdSID
30    FillGrid 1, grdSID
40    FillAckGrid 1, AckGrdSid
50    Exit Sub

cmdRevertToLab2_Click_Error:

      Dim strES As String
      Dim intEL As Integer

60    intEL = Erl
70    strES = Err.Description
80    LogError "frmConsultantListView", "cmdRevertToLab2_Click", intEL, strES

End Sub

Private Sub cmdRevertToLab3_Click()

10    On Error GoTo cmdRevertToLab3_Click_Error

20    RevertToLab 2, grdSID
30    FillGrid 2, grdSID
40    FillAckGrid 2, AckGrdSid
50    Exit Sub

cmdRevertToLab3_Click_Error:

      Dim strES As String
      Dim intEL As Integer

60    intEL = Erl
70    strES = Err.Description
80    LogError "frmConsultantListView", "cmdRevertToLab3_Click", intEL, strES

End Sub

Private Sub cmdSaveC_Click()
10    On Error GoTo cmdSaveC_Click_Error
20    btnCancel.Tag = ""
30    Call SaveComents(txtReport.Tag, True, "MicroConsultant", 0, Trim$(txtConC))
40    RemoveReport 0, txtReport.Tag, "N", 0

50    Call PrintThis("SaveTemp", txtReport.Tag, 0)
      '50        FillGrid 0, grdSID   ' Coment By Masood 21-05-2014
60    txtReport.Text = ""
70    txtConC.Text = ""
      '80        iMsg "Comments are saved" & vbCrLf & "Please press refresh to refresh list", vbInformation
80    Call RefereshReport(0, txtReport.Tag, grdSID)
90    Exit Sub

cmdSaveC_Click_Error:

      Dim strES As String
      Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "frmConsultantListView", "cmdSaveC_Click", intEL, strES

End Sub


Private Function RefereshReport(ConIndex As String, SampleID As String, ByVal GridtoFill As MSFlexGrid) As Boolean    ' Masood 21-05-2014

      Dim sql As String
      Dim tb As New ADODB.Recordset

10    On Error GoTo CheckRefereshed_Error
20    frmeRefreshing.Visible = True
30    MarkFlexGridRow GridtoFill, CurrentGridRow, 8, 8, vbRed
WaitForPrinter:
40    sql = "SELECT * FROM UnauthorisedReports WHERE " & _
            "SampleID =  '" & SampleID & "'"
50    Set tb = New Recordset
60    RecOpenServer 0, tb, sql
70    If tb.EOF = True Then
80        tb.Close
          ' Me.Caption = "Please Wait while Refreshing Data"
90        If btnCancel.Tag <> "" Then
100           frmeRefreshing.Visible = False
110           btnCancel.Tag = ""
120           SSTab1.Enabled = True
130           Exit Function
140       End If
150       DoEvents

160       SSTab1.Enabled = False
170       GoTo WaitForPrinter
180   Else
          'Me.Caption = "Refreshed"
190       MarkFlexGridRow GridtoFill, CurrentGridRow, 1, 7, vbGreen
200       Call FillGrid(ConIndex, GridtoFill)
210       SSTab1.Enabled = True
220       frmeRefreshing.Visible = False
230   End If


240   Exit Function


CheckRefereshed_Error:

      Dim strES As String
      Dim intEL As Integer

250   intEL = Erl
260   strES = Err.Description
270   LogError "frmConsultantListView", "CheckRefereshed", intEL, strES, sql
End Function


Private Sub cmdSaveC2_Click()
10    On Error GoTo cmdSaveC2_Click_Error
20    If txtReport2.Tag = "" Then
30        Exit Sub
40    End If

50    Call SaveComents(txtReport2.Tag, True, "MicroConsultant", 1, Trim$(txtConC2))
60    RemoveReport 1, txtReport2.Tag, "N", 0
70    Call PrintThis("SaveTemp", txtReport2.Tag, 1)
80    FillGrid 1, grdSid2
90    txtReport2.Text = ""
100   txtConC2.Text = ""
      'iMsg "Comments are saved" & vbCrLf & "Please press refresh to refresh list", vbInformation

110   Call RefereshReport(1, txtReport2.Tag, grdSid2)


120   Exit Sub

cmdSaveC2_Click_Error:

      Dim strES As String
      Dim intEL As Integer

130   intEL = Erl
140   strES = Err.Description
150   LogError "frmConsultantListView", "cmdSaveC2_Click", intEL, strES


End Sub

Private Sub cmdSaveC3_Click()
10    On Error GoTo cmdSaveC3_Click_Error
20    If txtReport3.Tag = "" Then
30        Exit Sub
40    End If
50    Call SaveComents(txtReport3.Tag, True, "MicroConsultant", 2, Trim$(txtConC3))
60    RemoveReport 2, txtReport3.Tag, "N", 0
70    Call PrintThis("SaveTemp", txtReport3.Tag, 2)
      'FillGrid 2, grdSID2
80    txtConC3.Text = ""
90    txtReport3.Text = ""
      'iMsg "Comments are saved" & vbCrLf & "Please press refresh to refresh list", vbInformation
100   Call RefereshReport(2, txtReport3.Tag, grdSid3)
110   Exit Sub

cmdSaveC3_Click_Error:

      Dim strES As String
      Dim intEL As Integer

120   intEL = Erl
130   strES = Err.Description
140   LogError "frmConsultantListView", "cmdSaveC3_Click", intEL, strES


End Sub





'---------------------------------------------------------------------------------------
' Procedure : cmdReleaseReport3_Click
' Author    : XPMUser
' Date      : 3/13/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmdReleaseReport3_Click()
10    On Error GoTo cmdReleaseReport3_Click_Error


20    Call AuthenticateConsultantList(2, grdSid3)
30    Call FillGrid(2, grdSid3)


40    Exit Sub


cmdReleaseReport3_Click_Error:

      Dim strES As String
      Dim intEL As Integer

50    intEL = Erl
60    strES = Err.Description
70    LogError "frmConsultantListView", "cmdReleaseReport3_Click", intEL, strES

End Sub

Private Sub Command1_Click()

End Sub

'---------------------------------------------------------------------------------------
' Procedure : grdSID_Click
' Author    : XPMUser
' Date      : 3/4/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub grdSID_Click()

      Static SortOrder As Boolean
10    On Error GoTo grdSID_Click_Error

20    If grdSID.MouseRow = 0 Then
30        MarkFlexGridRow grdSID, CurrentGridRow, 1, 7, 0
40        grdSID.Col = grdSID.MouseCol
50        If SortOrder Then
60            grdSID.Sort = flexSortGenericAscending
70        Else
80            grdSID.Sort = flexSortGenericDescending
90        End If
100       SortOrder = Not SortOrder
110       MarkFlexGridRow grdSID, grdSID.row, 1, 7, vbYellow
120       Exit Sub
130   End If

140   PreviousGridRow = CurrentGridRow
150   CurrentGridRow = grdSID.row
160   If PreviousGridRow > 0 Then
170       MarkFlexGridRow grdSID, PreviousGridRow, 1, 7, 0
180   End If
190   MarkFlexGridRow grdSID, CurrentGridRow, 1, 7, vbYellow
200   txtReport.Tag = grdSID.TextMatrix(grdSID.row, 0)
210   cmdAck.Enabled = Not (GetConAck(0, grdSID))

220   Call GetReport(grdSID, txtReport, txtPages, 0, txtConC)
230   If txtReport.Text <> "" And CurrentGridRow > 0 Then
240       MarkFlexGridRow grdSID, CurrentGridRow, 8, 8, vbGreen
250   End If
260   Exit Sub


grdSID_Click_Error:

      Dim strES As String
      Dim intEL As Integer

270   intEL = Erl
280   strES = Err.Description
290   LogError "frmConsultantListView", "grdSID_Click", intEL, strES
End Sub

'---------------------------------------------------------------------------------------
' Procedure : grdSID2_Click
' Author    : XPMUser
' Date      : 3/4/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub grdSID2_Click()

      Static SortOrder As Boolean

10    On Error GoTo grdSID2_Click_Error

20    If grdSid2.MouseRow = 0 Then
30        MarkFlexGridRow grdSid2, CurrentGridRow, 1, 7, 0
40        grdSid2.Col = grdSid2.MouseCol
50        If SortOrder Then
60            grdSid2.Sort = flexSortGenericAscending
70        Else
80            grdSid2.Sort = flexSortGenericDescending
90        End If
100       SortOrder = Not SortOrder
110       MarkFlexGridRow grdSid2, grdSid2.row, 1, 7, vbYellow
120       Exit Sub
130   End If

140   PreviousGridRow = CurrentGridRow
150   CurrentGridRow = grdSid2.row
160   If PreviousGridRow > 0 Then
170       MarkFlexGridRow grdSid2, PreviousGridRow, 1, 7, 0
180   End If
190   MarkFlexGridRow grdSid2, CurrentGridRow, 1, 7, vbYellow
200   txtReport2.Tag = grdSid2.TextMatrix(grdSid2.row, 0)
210   txtReport2 = ""
220   Call GetReport(grdSid2, txtReport2, txtPages2, 1, txtConC2)


230   If txtReport2.Text <> "" And CurrentGridRow > 0 Then
240       MarkFlexGridRow grdSid2, CurrentGridRow, 8, 8, vbGreen
250   End If

260   Exit Sub


grdSID2_Click_Error:

      Dim strES As String
      Dim intEL As Integer

270   intEL = Erl
280   strES = Err.Description
290   LogError "frmConsultantListView", "grdSID2_Click", intEL, strES
End Sub

'---------------------------------------------------------------------------------------
' Procedure : grdSID3_Click
' Author    : XPMUser
' Date      : 3/5/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub grdSID3_Click()

      Static SortOrder As Boolean

10    On Error GoTo grdSID3_Click_Error

20    If grdSid3.MouseRow = 0 Then
30        MarkFlexGridRow grdSid3, CurrentGridRow, 1, 7, 0
40        grdSid3.Col = grdSid3.MouseCol
50        If SortOrder Then
60            grdSid3.Sort = flexSortGenericAscending
70        Else
80            grdSid3.Sort = flexSortGenericDescending
90        End If
100       SortOrder = Not SortOrder
110       MarkFlexGridRow grdSid3, grdSid3.row, 1, 7, vbYellow
120       Exit Sub
130   End If


140   PreviousGridRow = CurrentGridRow
150   CurrentGridRow = grdSid3.row
160   If PreviousGridRow > 0 Then
170       MarkFlexGridRow grdSid3, PreviousGridRow, 1, 7, 0
180   End If
190   MarkFlexGridRow grdSid3, CurrentGridRow, 1, 7, vbYellow
200   txtReport3.Tag = grdSid3.TextMatrix(grdSid3.row, 0)
210   txtReport3 = ""
220   Call GetReport(grdSid3, txtReport3, txtPages3, 1, txtConC3)

230   If txtReport3.Text <> "" And CurrentGridRow > 0 Then
240       MarkFlexGridRow grdSid3, CurrentGridRow, 8, 8, vbGreen
250   End If
260   Exit Sub


grdSID3_Click_Error:

      Dim strES As String
      Dim intEL As Integer

270   intEL = Erl
280   strES = Err.Description
290   LogError "frmConsultantListView", "grdSID3_Click", intEL, strES

End Sub

Private Sub Form_Load()
10    Me.Caption = "NetAcquire - Consultant List "

20    With SSTab1
30        .TabCaption(0) = HospName(0)
40        .TabCaption(1) = HospName(1)
50        .TabCaption(2) = HospName(2)

60        If HospName(0) <> "" Then
70            Call FillGrid(0, grdSID)
80            Call FillAckGrid(0, AckGrdSid)
90        Else
100           .TabCaption(0) = "Connection Failed"
110           .TabEnabled(0) = False
120       End If
130       If HospName(1) <> "" Then
140           Call FillGrid(1, grdSid2)
150           Call FillAckGrid(1, AckGrdSid2)
160       Else
170           .TabCaption(1) = "Connection Failed"
180           .TabEnabled(1) = False
190       End If
200       If HospName(2) <> "" Then
210           Call FillGrid(2, grdSid3)
220           Call FillAckGrid(2, AckGrdSid3)

230       Else
240           .TabCaption(2) = "Connection Failed"
250           .TabEnabled(2) = False
260       End If
270   End With
280   FillMSandConsultantComment
      ' Masood 21-05-2014
290   frmeRefreshing.Visible = False
300   SSTab1.Enabled = True
      ' Masood 21-05-2014
End Sub




Private Sub MSFlexGrid4_Click()

End Sub

