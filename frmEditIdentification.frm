VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEditIdentification 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NetAcquire - Blood Culture Identification"
   ClientHeight    =   9585
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15975
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9585
   ScaleWidth      =   15975
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer TimerBar 
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   800
      Left            =   14925
      Picture         =   "frmEditIdentification.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   133
      Top             =   8580
      Width           =   900
   End
   Begin VB.CommandButton cmdSaveMicro 
      Caption         =   "&Save"
      Height          =   800
      Left            =   13860
      Picture         =   "frmEditIdentification.frx":066A
      Style           =   1  'Graphical
      TabIndex        =   132
      Top             =   8580
      Width           =   900
   End
   Begin VB.CommandButton cmdViewReports 
      Caption         =   "Reports"
      Height          =   800
      Left            =   12780
      Picture         =   "frmEditIdentification.frx":0CD4
      Style           =   1  'Graphical
      TabIndex        =   131
      ToolTipText     =   "View Printed && Faxed Reports"
      Top             =   8580
      Width           =   900
   End
   Begin VB.Frame fraSampleID 
      Height          =   1275
      Left            =   120
      TabIndex        =   108
      Top             =   450
      Width           =   15765
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Sample ID"
         Height          =   195
         Index           =   92
         Left            =   510
         TabIndex        =   129
         Top             =   0
         Width           =   735
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
         TabIndex        =   128
         Top             =   240
         Width           =   1005
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Chart"
         Height          =   195
         Left            =   5760
         TabIndex        =   127
         Top             =   0
         Width           =   375
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "A and E"
         Height          =   195
         Index           =   0
         Left            =   7350
         TabIndex        =   126
         Top             =   -30
         Width           =   570
      End
      Begin VB.Label lblNameTitle 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         Height          =   195
         Left            =   2700
         TabIndex        =   125
         Top             =   0
         Width           =   420
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "D.o.B"
         Height          =   195
         Index           =   88
         Left            =   8910
         TabIndex        =   124
         Top             =   0
         Width           =   405
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Age"
         Height          =   195
         Index           =   89
         Left            =   9930
         TabIndex        =   123
         Top             =   -30
         Width           =   285
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Sex"
         Height          =   195
         Index           =   90
         Left            =   10770
         TabIndex        =   122
         Top             =   -30
         Width           =   270
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
         TabIndex        =   121
         Top             =   240
         Width           =   4035
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
         TabIndex        =   120
         Top             =   240
         Width           =   1605
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
         TabIndex        =   119
         Top             =   240
         Width           =   1605
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
         TabIndex        =   118
         Top             =   240
         Width           =   1095
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
         TabIndex        =   117
         Top             =   240
         Width           =   795
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
         TabIndex        =   116
         Top             =   240
         Width           =   735
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
         TabIndex        =   115
         Top             =   690
         Width           =   4035
      End
      Begin VB.Label lblABsInUse 
         BorderStyle     =   1  'Fixed Single
         Height          =   645
         Left            =   11580
         TabIndex        =   114
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Ward"
         Height          =   195
         Left            =   690
         TabIndex        =   113
         Top             =   720
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Clinician"
         Height          =   195
         Index           =   1
         Left            =   5310
         TabIndex        =   112
         Top             =   720
         Width           =   585
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
         TabIndex        =   111
         Top             =   690
         Width           =   2505
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "GP"
         Height          =   195
         Index           =   1
         Left            =   8640
         TabIndex        =   110
         Top             =   690
         Width           =   225
      End
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
         TabIndex        =   109
         Top             =   660
         Width           =   2385
      End
   End
   Begin VB.Frame FrameExtras 
      Caption         =   "Organism 6"
      ForeColor       =   &H00C000C0&
      Height          =   6495
      Index           =   6
      Left            =   13320
      TabIndex        =   90
      Top             =   1860
      Width           =   2505
      Begin VB.TextBox txtZN 
         Height          =   285
         Index           =   6
         Left            =   900
         TabIndex        =   99
         Top             =   600
         Width           =   1515
      End
      Begin VB.TextBox txtIndole 
         Height          =   285
         Index           =   6
         Left            =   900
         TabIndex        =   98
         Top             =   1230
         Width           =   1515
      End
      Begin VB.TextBox txtNotes 
         Height          =   3645
         Index           =   6
         Left            =   90
         MultiLine       =   -1  'True
         TabIndex        =   97
         Top             =   2730
         Width           =   2355
      End
      Begin VB.ComboBox cmbWetPrep 
         Height          =   315
         Index           =   6
         Left            =   900
         TabIndex        =   96
         Top             =   900
         Width           =   1515
      End
      Begin VB.ComboBox cmbGram 
         Height          =   315
         Index           =   6
         Left            =   900
         Sorted          =   -1  'True
         TabIndex        =   95
         Top             =   270
         Width           =   1515
      End
      Begin VB.TextBox txtReincubation 
         Height          =   285
         Index           =   6
         Left            =   900
         TabIndex        =   94
         Tag             =   "Rei"
         Top             =   2430
         Width           =   1515
      End
      Begin VB.TextBox txtOxidase 
         Height          =   285
         Index           =   6
         Left            =   900
         TabIndex        =   93
         Tag             =   "Oxi"
         Top             =   2130
         Width           =   1515
      End
      Begin VB.TextBox txtCatalase 
         Height          =   285
         Index           =   6
         Left            =   900
         TabIndex        =   92
         Tag             =   "Cat"
         Top             =   1830
         Width           =   1515
      End
      Begin VB.TextBox txtCoagulase 
         Height          =   285
         Index           =   6
         Left            =   900
         TabIndex        =   91
         Tag             =   "Coa"
         Top             =   1530
         Width           =   1515
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "ZN Stain"
         Height          =   195
         Index           =   6
         Left            =   225
         TabIndex        =   107
         Top             =   630
         Width           =   630
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Indole"
         Height          =   195
         Index           =   6
         Left            =   435
         TabIndex        =   106
         Top             =   1260
         Width           =   435
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Reinc"
         Height          =   195
         Index           =   18
         Left            =   450
         TabIndex        =   105
         Top             =   2430
         Width           =   420
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Wet Prep"
         Height          =   195
         Index           =   17
         Left            =   195
         TabIndex        =   104
         Top             =   930
         Width           =   675
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Oxidase"
         Height          =   195
         Index           =   16
         Left            =   300
         TabIndex        =   103
         Top             =   2160
         Width           =   570
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Catalase"
         Height          =   195
         Index           =   15
         Left            =   255
         TabIndex        =   102
         Top             =   1860
         Width           =   615
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Coagulase"
         Height          =   195
         Index           =   14
         Left            =   120
         TabIndex        =   101
         Top             =   1560
         Width           =   750
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Gram Stain"
         Height          =   195
         Index           =   13
         Left            =   90
         TabIndex        =   100
         Top             =   330
         Width           =   780
      End
   End
   Begin VB.Frame FrameExtras 
      Caption         =   "Organism 5"
      ForeColor       =   &H00C000C0&
      Height          =   6495
      Index           =   5
      Left            =   10680
      TabIndex        =   72
      Top             =   1860
      Width           =   2505
      Begin VB.TextBox txtCoagulase 
         Height          =   285
         Index           =   5
         Left            =   900
         TabIndex        =   81
         Tag             =   "Coa"
         Top             =   1530
         Width           =   1515
      End
      Begin VB.TextBox txtCatalase 
         Height          =   285
         Index           =   5
         Left            =   900
         TabIndex        =   80
         Tag             =   "Cat"
         Top             =   1830
         Width           =   1515
      End
      Begin VB.TextBox txtOxidase 
         Height          =   285
         Index           =   5
         Left            =   900
         TabIndex        =   79
         Tag             =   "Oxi"
         Top             =   2130
         Width           =   1515
      End
      Begin VB.TextBox txtReincubation 
         Height          =   285
         Index           =   5
         Left            =   900
         TabIndex        =   78
         Tag             =   "Rei"
         Top             =   2430
         Width           =   1515
      End
      Begin VB.ComboBox cmbGram 
         Height          =   315
         Index           =   5
         Left            =   900
         Sorted          =   -1  'True
         TabIndex        =   77
         Top             =   270
         Width           =   1515
      End
      Begin VB.ComboBox cmbWetPrep 
         Height          =   315
         Index           =   5
         Left            =   900
         TabIndex        =   76
         Top             =   900
         Width           =   1515
      End
      Begin VB.TextBox txtNotes 
         Height          =   3645
         Index           =   5
         Left            =   90
         MultiLine       =   -1  'True
         TabIndex        =   75
         Top             =   2730
         Width           =   2355
      End
      Begin VB.TextBox txtIndole 
         Height          =   285
         Index           =   5
         Left            =   900
         TabIndex        =   74
         Top             =   1230
         Width           =   1515
      End
      Begin VB.TextBox txtZN 
         Height          =   285
         Index           =   5
         Left            =   900
         TabIndex        =   73
         Top             =   600
         Width           =   1515
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Gram Stain"
         Height          =   195
         Index           =   12
         Left            =   90
         TabIndex        =   89
         Top             =   330
         Width           =   780
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Coagulase"
         Height          =   195
         Index           =   11
         Left            =   120
         TabIndex        =   88
         Top             =   1560
         Width           =   750
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Catalase"
         Height          =   195
         Index           =   10
         Left            =   255
         TabIndex        =   87
         Top             =   1860
         Width           =   615
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Oxidase"
         Height          =   195
         Index           =   9
         Left            =   300
         TabIndex        =   86
         Top             =   2160
         Width           =   570
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Wet Prep"
         Height          =   195
         Index           =   8
         Left            =   195
         TabIndex        =   85
         Top             =   930
         Width           =   675
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Reinc"
         Height          =   195
         Index           =   7
         Left            =   450
         TabIndex        =   84
         Top             =   2430
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Indole"
         Height          =   195
         Index           =   5
         Left            =   435
         TabIndex        =   83
         Top             =   1260
         Width           =   435
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "ZN Stain"
         Height          =   195
         Index           =   5
         Left            =   225
         TabIndex        =   82
         Top             =   630
         Width           =   630
      End
   End
   Begin VB.Frame FrameExtras 
      Caption         =   "Organism 1"
      ForeColor       =   &H00C000C0&
      Height          =   6495
      Index           =   1
      Left            =   120
      TabIndex        =   54
      Top             =   1860
      Width           =   2505
      Begin VB.TextBox txtCoagulase 
         Height          =   285
         Index           =   1
         Left            =   900
         TabIndex        =   63
         Tag             =   "Coa"
         Top             =   1530
         Width           =   1515
      End
      Begin VB.TextBox txtCatalase 
         Height          =   285
         Index           =   1
         Left            =   900
         TabIndex        =   62
         Tag             =   "Cat"
         Top             =   1830
         Width           =   1515
      End
      Begin VB.TextBox txtOxidase 
         Height          =   285
         Index           =   1
         Left            =   900
         TabIndex        =   61
         Tag             =   "Oxi"
         Top             =   2130
         Width           =   1515
      End
      Begin VB.TextBox txtReincubation 
         Height          =   285
         Index           =   1
         Left            =   900
         TabIndex        =   60
         Tag             =   "Rei"
         Top             =   2430
         Width           =   1515
      End
      Begin VB.ComboBox cmbGram 
         Height          =   315
         Index           =   1
         Left            =   900
         Sorted          =   -1  'True
         TabIndex        =   59
         Top             =   270
         Width           =   1515
      End
      Begin VB.ComboBox cmbWetPrep 
         Height          =   315
         Index           =   1
         Left            =   900
         TabIndex        =   58
         Top             =   900
         Width           =   1515
      End
      Begin VB.TextBox txtNotes 
         Height          =   3645
         Index           =   1
         Left            =   90
         MultiLine       =   -1  'True
         TabIndex        =   57
         Top             =   2730
         Width           =   2355
      End
      Begin VB.TextBox txtIndole 
         Height          =   285
         Index           =   1
         Left            =   900
         TabIndex        =   56
         Top             =   1230
         Width           =   1515
      End
      Begin VB.TextBox txtZN 
         Height          =   285
         Index           =   1
         Left            =   900
         TabIndex        =   55
         Top             =   600
         Width           =   1515
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Gram Stain"
         Height          =   195
         Index           =   3
         Left            =   90
         TabIndex        =   71
         Top             =   330
         Width           =   780
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Coagulase"
         Height          =   195
         Index           =   35
         Left            =   120
         TabIndex        =   70
         Top             =   1560
         Width           =   750
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Catalase"
         Height          =   195
         Index           =   36
         Left            =   255
         TabIndex        =   69
         Top             =   1860
         Width           =   615
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Oxidase"
         Height          =   195
         Index           =   37
         Left            =   300
         TabIndex        =   68
         Top             =   2160
         Width           =   570
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Wet Prep"
         Height          =   195
         Index           =   34
         Left            =   195
         TabIndex        =   67
         Top             =   930
         Width           =   675
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Reinc"
         Height          =   195
         Index           =   38
         Left            =   450
         TabIndex        =   66
         Top             =   2430
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Indole"
         Height          =   195
         Index           =   0
         Left            =   435
         TabIndex        =   65
         Top             =   1260
         Width           =   435
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "ZN Stain"
         Height          =   195
         Index           =   0
         Left            =   225
         TabIndex        =   64
         Top             =   630
         Width           =   630
      End
   End
   Begin VB.Frame FrameExtras 
      Caption         =   "Organism 2"
      ForeColor       =   &H00C000C0&
      Height          =   6495
      Index           =   2
      Left            =   2760
      TabIndex        =   36
      Top             =   1860
      Width           =   2505
      Begin VB.TextBox txtZN 
         Height          =   285
         Index           =   2
         Left            =   900
         TabIndex        =   45
         Top             =   600
         Width           =   1515
      End
      Begin VB.TextBox txtIndole 
         Height          =   285
         Index           =   2
         Left            =   900
         TabIndex        =   44
         Top             =   1230
         Width           =   1515
      End
      Begin VB.TextBox txtNotes 
         Height          =   3645
         Index           =   2
         Left            =   90
         MultiLine       =   -1  'True
         TabIndex        =   43
         Top             =   2730
         Width           =   2355
      End
      Begin VB.ComboBox cmbWetPrep 
         Height          =   315
         Index           =   2
         Left            =   900
         TabIndex        =   42
         Top             =   900
         Width           =   1515
      End
      Begin VB.ComboBox cmbGram 
         Height          =   315
         Index           =   2
         Left            =   900
         Sorted          =   -1  'True
         TabIndex        =   41
         Top             =   270
         Width           =   1515
      End
      Begin VB.TextBox txtReincubation 
         Height          =   285
         Index           =   2
         Left            =   900
         TabIndex        =   40
         Tag             =   "Rei"
         Top             =   2430
         Width           =   1515
      End
      Begin VB.TextBox txtOxidase 
         Height          =   285
         Index           =   2
         Left            =   900
         TabIndex        =   39
         Tag             =   "Oxi"
         Top             =   2130
         Width           =   1515
      End
      Begin VB.TextBox txtCatalase 
         Height          =   285
         Index           =   2
         Left            =   900
         TabIndex        =   38
         Tag             =   "Cat"
         Top             =   1830
         Width           =   1515
      End
      Begin VB.TextBox txtCoagulase 
         Height          =   285
         Index           =   2
         Left            =   900
         TabIndex        =   37
         Tag             =   "Coa"
         Top             =   1530
         Width           =   1515
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "ZN Stain"
         Height          =   195
         Index           =   2
         Left            =   255
         TabIndex        =   53
         Top             =   630
         Width           =   630
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Indole"
         Height          =   195
         Index           =   2
         Left            =   435
         TabIndex        =   52
         Top             =   1260
         Width           =   435
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Reinc"
         Height          =   195
         Index           =   42
         Left            =   450
         TabIndex        =   51
         Top             =   2430
         Width           =   420
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Wet Prep"
         Height          =   195
         Index           =   43
         Left            =   195
         TabIndex        =   50
         Top             =   930
         Width           =   675
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Oxidase"
         Height          =   195
         Index           =   44
         Left            =   300
         TabIndex        =   49
         Top             =   2160
         Width           =   570
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Catalase"
         Height          =   195
         Index           =   45
         Left            =   255
         TabIndex        =   48
         Top             =   1860
         Width           =   615
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Coagulase"
         Height          =   195
         Index           =   46
         Left            =   120
         TabIndex        =   47
         Top             =   1560
         Width           =   750
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Gram Stain"
         Height          =   195
         Index           =   47
         Left            =   90
         TabIndex        =   46
         Top             =   330
         Width           =   780
      End
   End
   Begin VB.Frame FrameExtras 
      Caption         =   "Organism 3"
      ForeColor       =   &H00C000C0&
      Height          =   6495
      Index           =   3
      Left            =   5400
      TabIndex        =   18
      Top             =   1860
      Width           =   2505
      Begin VB.TextBox txtZN 
         Height          =   285
         Index           =   3
         Left            =   900
         TabIndex        =   27
         Top             =   600
         Width           =   1515
      End
      Begin VB.TextBox txtIndole 
         Height          =   285
         Index           =   3
         Left            =   900
         TabIndex        =   26
         Top             =   1230
         Width           =   1515
      End
      Begin VB.TextBox txtNotes 
         Height          =   3645
         Index           =   3
         Left            =   90
         MultiLine       =   -1  'True
         TabIndex        =   25
         Top             =   2730
         Width           =   2355
      End
      Begin VB.ComboBox cmbWetPrep 
         Height          =   315
         Index           =   3
         Left            =   900
         TabIndex        =   24
         Top             =   900
         Width           =   1515
      End
      Begin VB.ComboBox cmbGram 
         Height          =   315
         Index           =   3
         Left            =   900
         Sorted          =   -1  'True
         TabIndex        =   23
         Top             =   270
         Width           =   1515
      End
      Begin VB.TextBox txtReincubation 
         Height          =   285
         Index           =   3
         Left            =   900
         TabIndex        =   22
         Tag             =   "Rei"
         Top             =   2430
         Width           =   1515
      End
      Begin VB.TextBox txtOxidase 
         Height          =   285
         Index           =   3
         Left            =   900
         TabIndex        =   21
         Tag             =   "Oxi"
         Top             =   2130
         Width           =   1515
      End
      Begin VB.TextBox txtCatalase 
         Height          =   285
         Index           =   3
         Left            =   900
         TabIndex        =   20
         Tag             =   "Cat"
         Top             =   1830
         Width           =   1515
      End
      Begin VB.TextBox txtCoagulase 
         Height          =   285
         Index           =   3
         Left            =   900
         TabIndex        =   19
         Tag             =   "Coa"
         Top             =   1530
         Width           =   1515
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "ZN Stain"
         Height          =   195
         Index           =   3
         Left            =   255
         TabIndex        =   35
         Top             =   630
         Width           =   630
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Indole"
         Height          =   195
         Index           =   3
         Left            =   435
         TabIndex        =   34
         Top             =   1260
         Width           =   435
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Reinc"
         Height          =   195
         Index           =   48
         Left            =   450
         TabIndex        =   33
         Top             =   2430
         Width           =   420
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Wet Prep"
         Height          =   195
         Index           =   49
         Left            =   195
         TabIndex        =   32
         Top             =   930
         Width           =   675
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Oxidase"
         Height          =   195
         Index           =   50
         Left            =   300
         TabIndex        =   31
         Top             =   2160
         Width           =   570
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Catalase"
         Height          =   195
         Index           =   51
         Left            =   255
         TabIndex        =   30
         Top             =   1860
         Width           =   615
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Coagulase"
         Height          =   195
         Index           =   52
         Left            =   120
         TabIndex        =   29
         Top             =   1560
         Width           =   750
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Gram Stain"
         Height          =   195
         Index           =   53
         Left            =   90
         TabIndex        =   28
         Top             =   330
         Width           =   780
      End
   End
   Begin VB.Frame FrameExtras 
      Caption         =   "Organism 4"
      ForeColor       =   &H00C000C0&
      Height          =   6495
      Index           =   4
      Left            =   8040
      TabIndex        =   0
      Top             =   1860
      Width           =   2505
      Begin VB.TextBox txtZN 
         Height          =   285
         Index           =   4
         Left            =   900
         TabIndex        =   9
         Top             =   600
         Width           =   1515
      End
      Begin VB.TextBox txtIndole 
         Height          =   285
         Index           =   4
         Left            =   900
         TabIndex        =   8
         Top             =   1230
         Width           =   1515
      End
      Begin VB.TextBox txtNotes 
         Height          =   3645
         Index           =   4
         Left            =   90
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   2730
         Width           =   2355
      End
      Begin VB.ComboBox cmbWetPrep 
         Height          =   315
         Index           =   4
         Left            =   900
         TabIndex        =   6
         Top             =   900
         Width           =   1515
      End
      Begin VB.ComboBox cmbGram 
         Height          =   315
         Index           =   4
         Left            =   900
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   270
         Width           =   1515
      End
      Begin VB.TextBox txtReincubation 
         Height          =   285
         Index           =   4
         Left            =   900
         TabIndex        =   4
         Tag             =   "Rei"
         Top             =   2430
         Width           =   1515
      End
      Begin VB.TextBox txtOxidase 
         Height          =   285
         Index           =   4
         Left            =   900
         TabIndex        =   3
         Tag             =   "Oxi"
         Top             =   2130
         Width           =   1515
      End
      Begin VB.TextBox txtCatalase 
         Height          =   285
         Index           =   4
         Left            =   900
         TabIndex        =   2
         Tag             =   "Cat"
         Top             =   1830
         Width           =   1515
      End
      Begin VB.TextBox txtCoagulase 
         Height          =   285
         Index           =   4
         Left            =   900
         TabIndex        =   1
         Tag             =   "Coa"
         Top             =   1530
         Width           =   1515
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "ZN Stain"
         Height          =   195
         Index           =   4
         Left            =   225
         TabIndex        =   17
         Top             =   630
         Width           =   630
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Indole"
         Height          =   195
         Index           =   4
         Left            =   435
         TabIndex        =   16
         Top             =   1260
         Width           =   435
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Reinc"
         Height          =   195
         Index           =   54
         Left            =   450
         TabIndex        =   15
         Top             =   2430
         Width           =   420
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Wet Prep"
         Height          =   195
         Index           =   55
         Left            =   195
         TabIndex        =   14
         Top             =   930
         Width           =   675
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Oxidase"
         Height          =   195
         Index           =   56
         Left            =   300
         TabIndex        =   13
         Top             =   2160
         Width           =   570
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Catalase"
         Height          =   195
         Index           =   57
         Left            =   255
         TabIndex        =   12
         Top             =   1860
         Width           =   615
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Coagulase"
         Height          =   195
         Index           =   59
         Left            =   120
         TabIndex        =   11
         Top             =   1560
         Width           =   750
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Gram Stain"
         Height          =   195
         Index           =   60
         Left            =   90
         TabIndex        =   10
         Top             =   330
         Width           =   780
      End
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   135
      Left            =   120
      TabIndex        =   130
      Top             =   120
      Width           =   15765
      _ExtentX        =   27808
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   1
      Max             =   600
      Scrolling       =   1
   End
End
Attribute VB_Name = "frmEditIdentification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private SampleIDWithOffset As Double


Private Sub cmbWetPrep_Click(Index As Integer)

10        cmdSaveMicro.Enabled = True
'20        cmdSaveHold.Enabled = True

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
110       LogError "frmEditIdentification", "cmbWetPrep_LostFocus", intEL, strES, sql


End Sub

Private Sub cmdViewReports_Click()

frmRFT.SampleID = Val(lblSampleID) + SysOptMicroOffset(0)
frmRFT.Dept = "N"
frmRFT.Show 1

End Sub
Private Sub cmdSaveMicro_Click()

pBar = 0

GetSampleIDWithOffset
SaveIdent




End Sub
Private Sub cmdCancel_Click()

pBar = 0

Unload Me

End Sub

Private Sub cmbGram_LostFocus(Index As Integer)

Dim tb As Recordset
Dim sql As String

On Error GoTo cmbGram_LostFocus_Error

sql = "Select * from Lists where " & _
      "ListType = 'GS' " & _
      "and Code = '" & cmbGram(Index) & "'"
Set tb = New Recordset
RecOpenServer 0, tb, sql
If Not tb.EOF Then
    cmbGram(Index) = tb!Text & ""
End If

Exit Sub

cmbGram_LostFocus_Error:

Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "frmEditMicrobiologyNew", "cmbGram_LostFocus", intEL, strES, sql


End Sub

Private Sub FillIdent()

Dim tb As Recordset
Dim sql As String

On Error GoTo FillIdent_Error

sql = "SELECT Text FROM Lists WHERE " & _
      "ListType = 'GS' ORDER BY ListOrder"
Set tb = New Recordset
RecOpenServer 0, tb, sql
Do While Not tb.EOF
    For n = 1 To 6
        cmbGram(n).AddItem tb!Text & ""
    Next
    tb.MoveNext
Loop

Exit Sub

FillIdent_Error:
Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "frmEditIdentification", "FillIdent", intEL, strES, sql

End Sub

Private Function IdentIsSaveable(ByVal Index As Integer) As Boolean

On Error GoTo IdentIsSaveable_Error

IdentIsSaveable = False

With frmEditIdentification

    If Trim$(.cmbGram(Index).Text & _
             .txtZN(Index).Text & _
             .cmbWetPrep(Index).Text & _
             .txtIndole(Index).Text & _
             .txtCoagulase(Index).Text & _
             .txtCatalase(Index).Text & _
             .txtOxidase(Index).Text & _
             .txtReincubation(Index) & _
             .txtNotes(Index).Text) <> "" Then

        IdentIsSaveable = True

    End If

End With

Exit Function

IdentIsSaveable_Error:

Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "basMicro", "IdentIsSaveable", intEL, strES

End Function

Private Sub SaveIdent()

Dim tb As Recordset
Dim sql As String
Dim n As Integer

On Error GoTo SaveIdent_Error

For n = 1 To 6
    If IdentIsSaveable(n) Then
        sql = "Select * from UrineIdent where " & _
              "SampleID = '" & SampleIDWithOffset & "' " & _
              "and Isolate = " & n
        Set tb = New Recordset
        RecOpenClient 0, tb, sql

        If tb.EOF Then
            tb.AddNew
        End If
        tb!SampleID = SampleIDWithOffset
        tb!Isolate = n
        tb!SampleID = SampleIDWithOffset
        tb!Gram = cmbGram(n)
        tb!ZN = txtZN(n)
        tb!WetPrep = cmbWetPrep(n)
        tb!Indole = txtIndole(n)
        tb!Coagulase = txtCoagulase(n)
        tb!Catalase = txtCatalase(n)
        tb!Oxidase = txtOxidase(n)
        tb!Reincubation = txtReincubation(n)
        tb!Notes = txtNotes(n)
        tb!UserName = UserName
        tb.Update
    Else
        sql = "Delete from UrineIdent where " & _
              "SampleID = '" & SampleIDWithOffset & "' " & _
              "and Isolate = " & n
        Cnxn(0).Execute sql
    End If

Next

Exit Sub

SaveIdent_Error:

Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "frmEditMicrobiologyNew", "SaveIdent", intEL, strES, sql

End Sub

Private Sub GetSampleIDWithOffset()

On Error GoTo GetSampleIDWithOffset_Error

SampleIDWithOffset = Val(lblSampleID) + SysOptMicroOffset(0)

Exit Sub

GetSampleIDWithOffset_Error:

Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "frmEditMicrobiologyNew", "GetSampleIDWithOffset", intEL, strES

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


For n = 1 To 6
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
'    Call LoadLockStatus(2)



Exit Function

LoadIdent_Error:

Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "frmEditMicrobiologyNew", "LoadIdent", intEL, strES, sql

End Function

Private Sub ClearIdent()

Dim Index As Integer

On Error GoTo ClearIdent_Error

For Index = 1 To 6
    cmbGram(Index) = ""
    txtZN(Index) = ""
    cmbWetPrep(Index) = ""
    txtIndole(Index) = ""
    txtCoagulase(Index) = ""
    txtCatalase(Index) = ""
    txtOxidase(Index) = ""
    txtNotes(Index) = ""
    txtReincubation(Index) = ""
Next

Exit Sub

ClearIdent_Error:

Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "frmEditMicrobiologyNew", "ClearIdent", intEL, strES


End Sub

Private Sub Form_Activate()

On Error GoTo Form_Activate_Error


GetSampleIDWithOffset
LoadIdent

Exit Sub

Form_Activate_Error:
Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "frmEditIdentification", "Form_Activate", intEL, strES

End Sub

Private Sub Form_Load()

On Error GoTo Form_Load_Error

FillIdent
FixCombos
FillWetPrep

Exit Sub

Form_Load_Error:
Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "frmEditIdentification", "Form_Load", intEL, strES

End Sub
Private Sub FillWetPrep()
    Dim n As Integer
    Dim tb As Recordset
    Dim sql As String


    sql = "SELECT Text FROM Lists WHERE " & _
          "ListType = 'WP' ORDER BY ListOrder"
    Set tb = New Recordset
    RecOpenServer 0, tb, sql

    Do While Not tb.EOF
        For n = 1 To 4
            cmbWetPrep(n).AddItem tb!Text & ""
        Next
        tb.MoveNext
    Loop
End Sub

Private Sub FixCombos()

Dim i As Integer

On Error GoTo FixCombos_Error

For i = 1 To 6
    FixComboWidth cmbGram(i)
    FixComboWidth cmbWetPrep(i)
Next
    

Exit Sub

FixCombos_Error:
Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "frmEditIdentification", "FixCombos", intEL, strES

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
'190       cmdSaveHold.Enabled = True

200       Exit Sub

ClickMe_Error:

          Dim strES As String
          Dim intEL As Integer

210       intEL = Erl
220       strES = Err.Description
230       LogError "frmEditMicrobiologyNew", "ClickMe", intEL, strES


End Sub

Private Sub txtCatalase_Click(Index As Integer)
10        ClickMe txtCatalase(Index)
End Sub

Private Sub txtCoagulase_Click(Index As Integer)
10        ClickMe txtCoagulase(Index)
End Sub

Private Sub txtIndole_Click(Index As Integer)

10        Select Case txtIndole(Index)
          Case "": txtIndole(Index) = "Pending"
20        Case "Pending": txtIndole(Index) = "Positive"
30        Case "Positive": txtIndole(Index) = "Negative"
40        Case "Negative": txtIndole(Index) = ""
50        End Select

60        cmdSaveMicro.Enabled = True
'70        cmdSaveHold.Enabled = True

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
'80        cmdSaveHold.Enabled = True

End Sub

Private Sub txtOxidase_Click(Index As Integer)
10        ClickMe txtOxidase(Index)
End Sub

Private Sub txtReincubation_Click(Index As Integer)
10        ClickMe txtReincubation(Index)
End Sub

Private Sub txtZN_Click(Index As Integer)

10        Select Case txtZN(Index)
          Case "": txtZN(Index) = "No acid fast bacilli seen"
20        Case "No acid fast bacilli seen": txtZN(Index) = "Acid fast bacilli seen"
30        Case "Acid fast bacilli seen": txtZN(Index) = ""
40        Case Else: txtZN(Index) = ""
50        End Select

60        cmdSaveMicro.Enabled = True
'70        cmdSaveHold.Enabled = True

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
'80        cmdSaveHold.Enabled = True

End Sub
