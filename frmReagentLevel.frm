VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frmReagentLevel 
   Caption         =   "NetAcquire - Reagent Levels"
   ClientHeight    =   4530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5565
   Icon            =   "frmReagentLevel.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4530
   ScaleWidth      =   5565
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   690
      Left            =   4080
      Picture         =   "frmReagentLevel.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   240
      Width           =   1245
   End
   Begin MSFlexGridLib.MSFlexGrid grdCat 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   5953
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      FormatString    =   "<Reagent Name        |<In Stock        |<Min Stock               "
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
   Begin Threed.SSPanel SSPanel1 
      Height          =   960
      Left            =   180
      TabIndex        =   1
      Top             =   60
      Width           =   3435
      _Version        =   65536
      _ExtentX        =   6059
      _ExtentY        =   1693
      _StockProps     =   15
      Caption         =   "Discipline"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   0
      Begin VB.OptionButton optDisp 
         Caption         =   "Biochemistry"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   7
         Tag             =   "Bio"
         Top             =   270
         Value           =   -1  'True
         Width           =   1275
      End
      Begin VB.OptionButton optDisp 
         Caption         =   "Coagulation"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   6
         Tag             =   "Coag"
         Top             =   495
         Width           =   1185
      End
      Begin VB.OptionButton optDisp 
         Caption         =   "Haematology"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   5
         Tag             =   "Haem"
         Top             =   720
         Width           =   1275
      End
      Begin VB.OptionButton optDisp 
         Caption         =   "Endocrinology"
         Height          =   195
         Index           =   3
         Left            =   1845
         TabIndex        =   4
         Tag             =   "End"
         Top             =   270
         Width           =   1365
      End
      Begin VB.OptionButton optDisp 
         Caption         =   "Immunology"
         Height          =   195
         Index           =   4
         Left            =   1845
         TabIndex        =   3
         Tag             =   "Imm"
         Top             =   495
         Width           =   1320
      End
      Begin VB.OptionButton optDisp 
         Caption         =   "Blood Gas"
         Height          =   195
         Index           =   5
         Left            =   1845
         TabIndex        =   2
         Tag             =   "BGA"
         Top             =   720
         Width           =   1275
      End
   End
End
Attribute VB_Name = "frmReagentLevel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

10        Unload Me

End Sub
