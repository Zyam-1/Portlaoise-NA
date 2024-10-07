VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form frmGetDoB 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Date of Birth"
   ClientHeight    =   4005
   ClientLeft      =   2910
   ClientTop       =   1455
   ClientWidth     =   3045
   ControlBox      =   0   'False
   Icon            =   "frmGetDoB.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   3045
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   720
      Left            =   1620
      Picture         =   "frmGetDoB.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3195
      Width           =   1305
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Ok"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   720
      Left            =   135
      Picture         =   "frmGetDoB.frx":1C8C
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3195
      Width           =   1305
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   2085
      Left            =   330
      TabIndex        =   2
      Top             =   690
      Width           =   2445
      Begin ComCtl2.UpDown udWithin 
         Height          =   285
         Left            =   1170
         TabIndex        =   7
         Top             =   1560
         Width           =   510
         _ExtentX        =   900
         _ExtentY        =   503
         _Version        =   327681
         Value           =   1
         BuddyControl    =   "lblWithin"
         BuddyDispid     =   196617
         OrigLeft        =   2490
         OrigTop         =   1320
         OrigRight       =   2730
         OrigBottom      =   1965
         Min             =   1
         Orientation     =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.OptionButton optFuzzy 
         Caption         =   "Use a 'Fuzzy' Search"
         Height          =   195
         Left            =   210
         TabIndex        =   4
         Top             =   720
         Value           =   -1  'True
         Width           =   1875
      End
      Begin VB.OptionButton optExact 
         Caption         =   "Use Exact Date of Birth"
         Height          =   195
         Left            =   210
         TabIndex        =   3
         Top             =   390
         Width           =   2055
      End
      Begin VB.Image Image1 
         Height          =   195
         Left            =   990
         Picture         =   "frmGetDoB.frx":360E
         Top             =   1290
         Width           =   165
      End
      Begin VB.Label lblSearch 
         AutoSize        =   -1  'True
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   330
         TabIndex        =   9
         Top             =   1290
         Width           =   615
      End
      Begin VB.Label lblWithin 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
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
         Left            =   1200
         TabIndex        =   6
         Top             =   1260
         Width           =   465
      End
      Begin VB.Label lblYears 
         AutoSize        =   -1  'True
         Caption         =   "Year"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1710
         TabIndex        =   5
         Top             =   1290
         Width           =   405
      End
   End
   Begin VB.TextBox txtDoB 
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
      Left            =   1440
      TabIndex        =   1
      ToolTipText     =   "ddmmyy or dd/mm/yyyy"
      Top             =   135
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Date of Birth"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   225
      TabIndex        =   0
      Top             =   210
      Width           =   1095
   End
End
Attribute VB_Name = "frmGetDoB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

10        On Error GoTo cmdCancel_Click_Error

20        txtDoB = ""
30        frmPatHistoryNew.oFor(0).Value = True
40        Me.Hide

50        Exit Sub

cmdCancel_Click_Error:

          Dim strES As String
          Dim intEL As Integer



60        intEL = Erl
70        strES = Err.Description
80        LogError "frmGetDoB", "cmdCancel_Click", intEL, strES


End Sub

Private Sub cmdOK_Click()

10        On Error GoTo cmdOK_Click_Error

20        Me.Hide

30        Exit Sub

cmdOK_Click_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmGetDoB", "cmdOK_Click", intEL, strES


End Sub

Private Sub optExact_Click()

10        On Error GoTo optExact_Click_Error

20        lblSearch.Enabled = False
30        lblSearch.Font.Bold = False
40        lblWithin.Enabled = False
50        lblYears.Enabled = False
60        lblYears.Font.Bold = False
70        udWithin.Enabled = False

80        Exit Sub

optExact_Click_Error:

          Dim strES As String
          Dim intEL As Integer



90        intEL = Erl
100       strES = Err.Description
110       LogError "frmGetDoB", "optExact_Click", intEL, strES


End Sub

Private Sub optFuzzy_Click()

10        On Error GoTo optFuzzy_Click_Error

20        lblSearch.Enabled = True
30        lblSearch.Font.Bold = True
40        lblWithin.Enabled = True
50        lblYears.Enabled = True
60        lblYears.Font.Bold = True
70        udWithin.Enabled = True

80        Exit Sub

optFuzzy_Click_Error:

          Dim strES As String
          Dim intEL As Integer



90        intEL = Erl
100       strES = Err.Description
110       LogError "frmGetDoB", "optFuzzy_Click", intEL, strES


End Sub

Private Sub txtDoB_Change()

10        On Error GoTo txtDoB_Change_Error

20        cmdOK.Enabled = False
30        If Len(txtDoB) = 6 Then
40            txtDoB = Convert62Date(txtDoB, BACKWARD)
50            cmdOK.Enabled = True
60        End If

70        Exit Sub

txtDoB_Change_Error:

          Dim strES As String
          Dim intEL As Integer



80        intEL = Erl
90        strES = Err.Description
100       LogError "frmGetDoB", "txtDoB_Change", intEL, strES


End Sub

Private Sub udWithin_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        On Error GoTo udWithin_MouseUp_Error

20        If lblWithin = "1" Then
30            lblYears = "Year"
40        Else
50            lblYears = "Years"
60        End If

70        Exit Sub

udWithin_MouseUp_Error:

          Dim strES As String
          Dim intEL As Integer



80        intEL = Erl
90        strES = Err.Description
100       LogError "frmGetDoB", "udWithin_MouseUp", intEL, strES


End Sub


