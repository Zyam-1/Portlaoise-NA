VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "ComCt232.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Begin VB.Form frmBioDefaults 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Biochemistry Parameter Definitions"
   ClientHeight    =   8475
   ClientLeft      =   3180
   ClientTop       =   1170
   ClientWidth     =   9735
   Icon            =   "frmBioDefaults.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8475
   ScaleWidth      =   9735
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraAnalyser 
      Caption         =   "Analyser"
      Height          =   915
      Left            =   2820
      TabIndex        =   71
      Top             =   180
      Width           =   1635
      Begin VB.ComboBox cmbAnalyser 
         Height          =   315
         Left            =   180
         Style           =   2  'Dropdown List
         TabIndex        =   72
         Top             =   300
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Enabled         =   0   'False
      Height          =   630
      Left            =   2565
      Picture         =   "frmBioDefaults.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   7785
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Caption         =   "Age Ranges"
      Height          =   2115
      Left            =   2850
      TabIndex        =   54
      Top             =   2610
      Width           =   6615
      Begin VB.CheckBox cEOD 
         Alignment       =   1  'Right Justify
         Caption         =   "End Of Day"
         Height          =   195
         Left            =   4770
         TabIndex        =   61
         Top             =   1500
         Width           =   1635
      End
      Begin VB.CheckBox cInUse 
         Alignment       =   1  'Right Justify
         Caption         =   "InUse"
         Height          =   195
         Left            =   4770
         TabIndex        =   60
         Top             =   1710
         Width           =   1635
      End
      Begin VB.CommandButton cmdAmendAgeRange 
         Caption         =   "Amend Age Range"
         Enabled         =   0   'False
         Height          =   525
         Left            =   4890
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   870
         Width           =   1155
      End
      Begin MSFlexGridLib.MSFlexGrid grdAge 
         Height          =   1725
         Left            =   570
         TabIndex        =   56
         Top             =   270
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   3043
         _Version        =   393216
         Cols            =   3
         BackColor       =   -2147483624
         ForeColor       =   -2147483635
         BackColorFixed  =   -2147483647
         ForeColorFixed  =   -2147483624
         FocusRect       =   0
         HighLight       =   0
         GridLines       =   3
         GridLinesFixed  =   3
         FormatString    =   "^Range #  |^Age From        |^Age To           "
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
      Begin VB.Label Label21 
         Caption         =   "Host Code"
         Height          =   315
         Left            =   4740
         TabIndex        =   59
         Top             =   360
         Width           =   795
      End
      Begin VB.Label lblHost 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5610
         TabIndex        =   58
         Top             =   330
         Width           =   765
      End
   End
   Begin VB.Frame fraImmulite 
      Caption         =   "Immulite Code"
      Height          =   615
      Left            =   7890
      TabIndex        =   47
      Top             =   2070
      Visible         =   0   'False
      Width           =   1575
      Begin VB.TextBox tImmuliteCode 
         Height          =   285
         Left            =   210
         TabIndex        =   48
         Top             =   240
         Width           =   945
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Analyser"
      Height          =   1005
      Left            =   9720
      TabIndex        =   43
      Top             =   1260
      Visible         =   0   'False
      Width           =   1575
      Begin VB.OptionButton optAnalyser 
         Caption         =   "None"
         Height          =   195
         Index           =   2
         Left            =   210
         TabIndex        =   46
         Top             =   720
         Width           =   795
      End
      Begin VB.OptionButton optAnalyser 
         Caption         =   "Immulite"
         Height          =   195
         Index           =   1
         Left            =   210
         TabIndex        =   45
         Top             =   480
         Width           =   885
      End
      Begin VB.OptionButton optAnalyser 
         Caption         =   "Olympus"
         Height          =   195
         Index           =   0
         Left            =   210
         TabIndex        =   44
         Top             =   240
         Value           =   -1  'True
         Width           =   915
      End
   End
   Begin VB.Frame Frame5 
      Height          =   1545
      Left            =   2850
      TabIndex        =   33
      Top             =   1140
      Width           =   5055
      Begin VB.TextBox txtCheckTime 
         Height          =   285
         Left            =   3960
         MaxLength       =   5
         TabIndex        =   66
         Top             =   1200
         Width           =   495
      End
      Begin VB.CheckBox chkPrintable 
         Alignment       =   1  'Right Justify
         Caption         =   "Printable"
         Height          =   195
         Left            =   570
         TabIndex        =   65
         Top             =   1170
         Width           =   1635
      End
      Begin VB.CheckBox oRM 
         Alignment       =   1  'Right Justify
         Caption         =   "Include in Running Mean"
         Height          =   195
         Left            =   2790
         TabIndex        =   36
         Top             =   270
         Width           =   2145
      End
      Begin VB.CheckBox oDelta 
         Alignment       =   1  'Right Justify
         Caption         =   "Do Delta"
         Height          =   225
         Left            =   3990
         TabIndex        =   35
         Top             =   540
         Width           =   945
      End
      Begin VB.CheckBox cKnown 
         Alignment       =   1  'Right Justify
         Caption         =   "Known to Analyser"
         Height          =   195
         Left            =   570
         TabIndex        =   34
         Top             =   930
         Width           =   1635
      End
      Begin VB.TextBox tdelta 
         Height          =   285
         Left            =   3960
         MaxLength       =   5
         TabIndex        =   11
         Top             =   870
         Width           =   975
      End
      Begin VB.TextBox tDP 
         Height          =   285
         Left            =   2010
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "1"
         Top             =   570
         Width           =   360
      End
      Begin VB.TextBox tPriority 
         Height          =   285
         Left            =   2010
         TabIndex        =   13
         Top             =   210
         Width           =   555
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   285
         Left            =   2370
         TabIndex        =   14
         Top             =   570
         Width           =   195
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   327681
         Value           =   1
         BuddyControl    =   "tDP"
         BuddyDispid     =   196627
         OrigLeft        =   2760
         OrigTop         =   540
         OrigRight       =   3000
         OrigBottom      =   885
         Max             =   3
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "day(s)"
         Height          =   195
         Left            =   4500
         TabIndex        =   68
         Top             =   1245
         Width           =   420
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Check time"
         Height          =   195
         Left            =   3120
         TabIndex        =   67
         Top             =   1245
         Width           =   795
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Value"
         Height          =   195
         Left            =   3510
         TabIndex        =   15
         Top             =   900
         Width           =   405
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Number of Decimal Places"
         Height          =   195
         Left            =   90
         TabIndex        =   16
         Top             =   630
         Width           =   1875
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Print Priority"
         Height          =   195
         Left            =   1110
         TabIndex        =   17
         Top             =   240
         Width           =   825
      End
   End
   Begin VB.Frame Frame4 
      Height          =   915
      Left            =   6720
      TabIndex        =   19
      Top             =   180
      Width           =   2775
      Begin VB.ComboBox cmbUnits 
         Height          =   315
         Left            =   810
         Style           =   2  'Dropdown List
         TabIndex        =   62
         Top             =   150
         Width           =   1545
      End
      Begin VB.Label lblBarCode 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   810
         TabIndex        =   57
         Top             =   480
         Width           =   1515
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Units"
         Height          =   195
         Left            =   330
         TabIndex        =   21
         Top             =   180
         Width           =   360
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Bar Code"
         Height          =   195
         Left            =   60
         TabIndex        =   20
         Top             =   540
         Width           =   660
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Sample Type"
      Height          =   915
      Left            =   4470
      TabIndex        =   18
      Top             =   180
      Width           =   2235
      Begin VB.ComboBox cmbSample 
         Height          =   315
         Left            =   225
         Style           =   2  'Dropdown List
         TabIndex        =   64
         Top             =   315
         Width           =   1860
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   630
      Left            =   6480
      Picture         =   "frmBioDefaults.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7785
      Width           =   1035
   End
   Begin VB.ListBox lstParameter 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7440
      IntegralHeight  =   0   'False
      Left            =   150
      TabIndex        =   1
      Top             =   210
      Width           =   2415
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2985
      Left            =   2850
      TabIndex        =   0
      Top             =   4740
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   5265
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
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
      TabCaption(0)   =   "Normal Range"
      TabPicture(0)   =   "frmBioDefaults.frx":091E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label6"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "tMaleHigh"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "tFemaleHigh"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "tMaleLow"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "tFemaleLow"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "chkHighGreater"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "chkLowLess"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "Flag Range"
      TabPicture(1)   =   "frmBioDefaults.frx":093A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "tfr(0)"
      Tab(1).Control(1)=   "tfr(2)"
      Tab(1).Control(2)=   "tfr(1)"
      Tab(1).Control(3)=   "tfr(3)"
      Tab(1).Control(4)=   "Label7"
      Tab(1).Control(5)=   "Label12"
      Tab(1).Control(6)=   "Label13"
      Tab(1).Control(7)=   "Label14"
      Tab(1).Control(8)=   "Label15"
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "Masks"
      TabPicture(2)   =   "frmBioDefaults.frx":0956
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label11"
      Tab(2).Control(1)=   "cH"
      Tab(2).Control(2)=   "cS"
      Tab(2).Control(3)=   "cL"
      Tab(2).Control(4)=   "cO"
      Tab(2).Control(5)=   "cG"
      Tab(2).Control(6)=   "cJ"
      Tab(2).ControlCount=   7
      TabCaption(3)   =   "Plausible"
      TabPicture(3)   =   "frmBioDefaults.frx":0972
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "tPlausibleLow"
      Tab(3).Control(1)=   "tPlausibleHigh"
      Tab(3).Control(2)=   "Label10"
      Tab(3).Control(3)=   "Label9"
      Tab(3).Control(4)=   "Label8"
      Tab(3).ControlCount=   5
      Begin VB.CheckBox chkLowLess 
         Caption         =   "Report range as ""< high"" if low range is 0.0"
         Height          =   255
         Left            =   1320
         TabIndex        =   70
         Top             =   2160
         Width           =   4755
      End
      Begin VB.CheckBox chkHighGreater 
         Caption         =   "Report range as ""> low"" if high range is 9999"
         Height          =   255
         Left            =   1320
         TabIndex        =   69
         Top             =   2460
         Width           =   4755
      End
      Begin VB.TextBox tPlausibleLow 
         Height          =   285
         Left            =   -72600
         TabIndex        =   52
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox tPlausibleHigh 
         Height          =   285
         Left            =   -72600
         TabIndex        =   51
         Top             =   990
         Width           =   1215
      End
      Begin VB.CheckBox cJ 
         Caption         =   "Icteric"
         Height          =   225
         Left            =   -72210
         TabIndex        =   42
         Top             =   1770
         Width           =   1095
      End
      Begin VB.CheckBox cG 
         Alignment       =   1  'Right Justify
         Caption         =   "Grossly Haemolysed"
         Height          =   225
         Left            =   -74100
         TabIndex        =   41
         Top             =   1770
         Width           =   1785
      End
      Begin VB.CheckBox cO 
         Caption         =   "Old"
         Height          =   225
         Left            =   -72210
         TabIndex        =   40
         Top             =   1530
         Width           =   585
      End
      Begin VB.CheckBox cL 
         Caption         =   "Lipaemic"
         Height          =   225
         Left            =   -72210
         TabIndex        =   39
         Top             =   1290
         Width           =   975
      End
      Begin VB.CheckBox cS 
         Alignment       =   1  'Right Justify
         Caption         =   "Slightly Haemolysed"
         Height          =   225
         Left            =   -74100
         TabIndex        =   38
         Top             =   1530
         Width           =   1785
      End
      Begin VB.CheckBox cH 
         Alignment       =   1  'Right Justify
         Caption         =   "Haemolysed"
         Height          =   225
         Left            =   -73500
         TabIndex        =   37
         Top             =   1290
         Width           =   1185
      End
      Begin VB.TextBox tfr 
         Height          =   315
         Index           =   0
         Left            =   -71640
         TabIndex        =   25
         Top             =   1200
         Width           =   915
      End
      Begin VB.TextBox tfr 
         Height          =   315
         Index           =   2
         Left            =   -73170
         TabIndex        =   24
         Top             =   1200
         Width           =   915
      End
      Begin VB.TextBox tfr 
         Height          =   315
         Index           =   1
         Left            =   -71640
         TabIndex        =   23
         Top             =   1800
         Width           =   915
      End
      Begin VB.TextBox tfr 
         Height          =   315
         Index           =   3
         Left            =   -73170
         TabIndex        =   22
         Top             =   1800
         Width           =   915
      End
      Begin VB.TextBox tFemaleLow 
         Height          =   315
         Left            =   1830
         MaxLength       =   5
         TabIndex        =   10
         Top             =   1800
         Width           =   915
      End
      Begin VB.TextBox tMaleLow 
         Height          =   315
         Left            =   3360
         MaxLength       =   5
         TabIndex        =   9
         Top             =   1800
         Width           =   915
      End
      Begin VB.TextBox tFemaleHigh 
         Height          =   315
         Left            =   1830
         MaxLength       =   5
         TabIndex        =   8
         Top             =   1200
         Width           =   915
      End
      Begin VB.TextBox tMaleHigh 
         Height          =   315
         Left            =   3360
         MaxLength       =   5
         TabIndex        =   7
         Top             =   1200
         Width           =   915
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Results outside this range will be marked as implausible."
         Height          =   195
         Left            =   -73740
         TabIndex        =   53
         Top             =   2430
         Width           =   3930
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Low"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   -73080
         TabIndex        =   50
         Top             =   1590
         Width           =   300
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "High"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   -73110
         TabIndex        =   49
         Top             =   1020
         Width           =   330
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Result values outside this range will be flagged as High or Low"
         Height          =   195
         Left            =   -74190
         TabIndex        =   32
         Top             =   2520
         Width           =   4410
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "These are the normal range values printed on the report forms."
         Height          =   195
         Left            =   840
         TabIndex        =   31
         Top             =   2700
         Width           =   4395
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Do not print the result if the sample is :-"
         Height          =   195
         Left            =   -74070
         TabIndex        =   30
         Top             =   750
         Width           =   2730
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Female"
         Height          =   195
         Left            =   -73020
         TabIndex        =   29
         Top             =   930
         Width           =   510
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Male"
         Height          =   195
         Left            =   -71370
         TabIndex        =   28
         Top             =   900
         Width           =   345
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "High"
         Height          =   195
         Left            =   -73950
         TabIndex        =   27
         Top             =   1290
         Width           =   330
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Low"
         Height          =   195
         Left            =   -73920
         TabIndex        =   26
         Top             =   1860
         Width           =   300
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Low"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   1290
         TabIndex        =   6
         Top             =   1860
         Width           =   300
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "High"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   1260
         TabIndex        =   5
         Top             =   1260
         Width           =   330
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Female"
         Height          =   195
         Left            =   1980
         TabIndex        =   4
         Top             =   930
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Male"
         Height          =   195
         Left            =   3630
         TabIndex        =   3
         Top             =   900
         Width           =   345
      End
   End
End
Attribute VB_Name = "frmBioDefaults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Activated As Boolean
Private loading As Boolean

Private FromDays() As Long
Private ToDays() As Long




Private Sub chkHighGreater_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
10        On Error GoTo chkHighGreater_MouseUp_Error

20        cmdUpdate.Enabled = True
30        lstParameter.Enabled = False

40        Exit Sub

chkHighGreater_MouseUp_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmBioDefaults", "chkHighGreater_MouseUp", intEL, strES


End Sub



Private Sub cEOD_Click()
          Dim sql As String



10        On Error GoTo cEOD_Click_Error

20        If lblHost = "" Then Exit Sub


30        If cEOD.Value = 1 Then
40            sql = "UPDATE biotestdefinitions set eod = 1 WHERE code = '" & lblHost & "'"
50            Cnxn(0).Execute sql
60        ElseIf cEOD.Value = 0 Then
70            sql = "UPDATE biotestdefinitions set eod = 0 WHERE code = '" & lblHost & "'"
80            Cnxn(0).Execute sql
90        End If



100       Exit Sub

cEOD_Click_Error:

          Dim strES As String
          Dim intEL As Integer

110       intEL = Erl
120       strES = Err.Description
130       LogError "frmBioDefaults", "cEOD_Click", intEL, strES, sql


End Sub

Private Sub cG_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        On Error GoTo cG_MouseUp_Error

20        cmdUpdate.Enabled = True
30        lstParameter.Enabled = False

40        Exit Sub

cG_MouseUp_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmBioDefaults", "cG_MouseUp", intEL, strES


End Sub

Private Sub cH_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        On Error GoTo cH_MouseUp_Error

20        cmdUpdate.Enabled = True
30        lstParameter.Enabled = False

40        Exit Sub

cH_MouseUp_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmBioDefaults", "cH_MouseUp", intEL, strES


End Sub

Private Sub chkLowLess_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        On Error GoTo chkLowLess_MouseUp_Error

20        cmdUpdate.Enabled = True
30        lstParameter.Enabled = False

40        Exit Sub

chkLowLess_MouseUp_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmBioDefaults", "chkLowLess_MouseUp", intEL, strES

End Sub

Private Sub chkPrintable_Click()
          Dim sql As String

10        On Error GoTo chkPrintable_Click_Error

20        If lblHost = "" Then Exit Sub

30        If chkPrintable.Value = 1 Then
40            sql = "UPDATE biotestdefinitions set Printable = 1 WHERE code = '" & lblHost & "'"
50            Cnxn(0).Execute sql
60        ElseIf chkPrintable.Value = 0 Then
70            sql = "UPDATE biotestdefinitions set printable = 0 WHERE code = '" & lblHost & "'"
80            Cnxn(0).Execute sql
90        End If

100       Exit Sub

chkPrintable_Click_Error:

          Dim strES As String
          Dim intEL As Integer

110       intEL = Erl
120       strES = Err.Description
130       LogError "frmBioDefaults", "chkPrintable_Click", intEL, strES, sql


End Sub

Private Sub cInUse_Click()
          Dim sql As String



10        On Error GoTo cInUse_Click_Error

20        If lblHost = "" Then Exit Sub


30        If cInUse.Value = 1 Then
40            sql = "UPDATE biotestdefinitions set inuse = 1 WHERE code = '" & lblHost & "'"
50            Cnxn(0).Execute sql
60        ElseIf cInUse.Value = 0 Then
70            sql = "UPDATE biotestdefinitions set inuse = 0 WHERE code = '" & lblHost & "'"
80            Cnxn(0).Execute sql
90        End If



100       Exit Sub

cInUse_Click_Error:

          Dim strES As String
          Dim intEL As Integer

110       intEL = Erl
120       strES = Err.Description
130       LogError "frmBioDefaults", "cInUse_Click", intEL, strES, sql


End Sub

Private Sub cJ_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        On Error GoTo cJ_MouseUp_Error

20        cmdUpdate.Enabled = True
30        lstParameter.Enabled = False

40        Exit Sub

cJ_MouseUp_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmBioDefaults", "cJ_MouseUp", intEL, strES


End Sub

Private Sub cKnown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        On Error GoTo cKnown_MouseUp_Error

20        cmdUpdate.Enabled = True
30        lstParameter.Enabled = False

40        Exit Sub

cKnown_MouseUp_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmBioDefaults", "cKnown_MouseUp", intEL, strES


End Sub

Private Sub cL_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        On Error GoTo cL_MouseUp_Error

20        cmdUpdate.Enabled = True
30        lstParameter.Enabled = False

40        Exit Sub

cL_MouseUp_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmBioDefaults", "cL_MouseUp", intEL, strES


End Sub

Private Sub cmbAnalyser_Click()

10    On Error GoTo cmbAnalyser_Click_Error

20    FilllParameter
30    With grdAge
40        .Rows = 2
50        .AddItem ""
60        .RemoveItem 1
70    End With

80    If cmbAnalyser.Text = "Immulite" Then
90        fraImmulite.Visible = True
100   Else
110       fraImmulite.Visible = False
120   End If

130   Exit Sub

cmbAnalyser_Click_Error:

      Dim strES As String
      Dim intEL As Integer

140   intEL = Erl
150   strES = Err.Description
160   LogError "frmBioDefaults", "cmbAnalyser_Click", intEL, strES

End Sub

Private Sub cmbSample_Click()

10        On Error GoTo cmbSample_Click_Error

20        FilllParameter

30        Exit Sub

cmbSample_Click_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmBioDefaults", "cmbSample_Click", intEL, strES


End Sub

Private Sub cmbUnits_Click()

10        On Error GoTo cmbUnits_Click_Error

20        If loading = False Then
30            cmdUpdate.Enabled = True
40            lstParameter.Enabled = False
50        End If

60        Exit Sub

cmbUnits_Click_Error:

          Dim strES As String
          Dim intEL As Integer

70        intEL = Erl
80        strES = Err.Description
90        LogError "frmBioDefaults", "cmbUnits_Click", intEL, strES


End Sub

Private Sub cmdAmendAgeRange_Click()



10        On Error GoTo cmdAmendAgeRange_Click_Error

20        If lstParameter = "" Then
30            iMsg "Please pick a Test!"
40            Exit Sub
50        End If

60        With frmAges
70            .Analyte = lstParameter
80            .SampleType = ListCodeFor("ST", cmbSample)
90            .Discipline = "Biochemistry"
100           .Show 1
110       End With

120       FillAges

130       FillAllDetails



140       Exit Sub

cmdAmendAgeRange_Click_Error:

          Dim strES As String
          Dim intEL As Integer

150       intEL = Erl
160       strES = Err.Description
170       LogError "frmBioDefaults", "cmdAmendAgeRange_Click", intEL, strES


End Sub

Private Sub cmdCancel_Click()

10        Unload Me

End Sub

Private Sub cmdUpdate_Click()

10        On Error GoTo cmdUpdate_Click_Error


20        SaveCommonDetails
30        cmdUpdate.Enabled = False
40        lstParameter.Enabled = True


50        Exit Sub

cmdUpdate_Click_Error:

          Dim strES As String
          Dim intEL As Integer

60        intEL = Erl
70        strES = Err.Description
80        LogError "frmBioDefaults", "cmdUpdate_Click", intEL, strES


End Sub

Private Sub cO_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        On Error GoTo cO_MouseUp_Error

20        cmdUpdate.Enabled = True
30        lstParameter.Enabled = False

40        Exit Sub

cO_MouseUp_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmBioDefaults", "cO_MouseUp", intEL, strES


End Sub

Private Sub cS_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        On Error GoTo cS_MouseUp_Error

20        cmdUpdate.Enabled = True
30        lstParameter.Enabled = False

40        Exit Sub

cS_MouseUp_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmBioDefaults", "cS_MouseUp", intEL, strES


End Sub

Private Sub FillAges()

          Dim tb As New Recordset
          Dim s As String
          Dim SampleType As String
          Dim n As Long
          Dim sql As String



10        On Error GoTo FillAges_Error

20        ClearFGrid grdAge



30        ReDim FromDays(0 To 0)
40        ReDim ToDays(0 To 0)

50        SampleType = ListCodeFor("ST", cmbSample)


60        sql = "SELECT * from BioTestDefinitions WHERE " & _
                "LongName = '" & lstParameter & "' " & _
                "and SampleType = '" & SampleType & "' " & _
                "Order by AgeFromDays"

70        Set tb = New Recordset
80        RecOpenServer 0, tb, sql

90        If tb.EOF Then Exit Sub
100       Do While Not tb.EOF
110           n = n + 1
120           tb.MoveNext
130       Loop
140       tb.MoveFirst
150       ReDim FromDays(0 To n - 1)
160       ReDim ToDays(0 To n - 1)
170       n = 0
180       Do While Not tb.EOF
190           If tb!AgeFromDays & "" = "" Then FromDays(n) = 0 Else FromDays(n) = tb!AgeFromDays
200           If tb!AgeToDays & "" = "" Then ToDays(n) = 43830 Else ToDays(n) = tb!AgeToDays
210           s = Format$(n) & vbTab & _
                  dmyFromCount(FromDays(n)) & vbTab & _
                  dmyFromCount(ToDays(n))
220           grdAge.AddItem s
230           n = n + 1
240           tb.MoveNext
250       Loop

260       FixG grdAge

270       grdAge.Col = 0
280       grdAge.Row = 1
290       grdAge.CellBackColor = vbYellow
300       grdAge.CellForeColor = vbBlue



310       Exit Sub

FillAges_Error:

          Dim strES As String
          Dim intEL As Integer

320       intEL = Erl
330       strES = Err.Description
340       LogError "frmBioDefaults", "FillAges", intEL, strES, sql


End Sub

Private Sub FillAllDetails()

          Dim tb As New Recordset
          Dim sql As String
          Dim SampleType As String
          Dim n As Long



10        On Error GoTo FillAllDetails_Error

20        tImmuliteCode = ""
30        tPriority = ""
40        cKnown = 0
50        oDelta = 0
60        tDP = 0
70        tdelta = ""
80        loading = True
90        cmbUnits.ListIndex = -1
100       loading = False
110       lblBarCode = ""
120       lblHost = ""
130       tPlausibleHigh = ""
140       tPlausibleLow = ""

150       tMaleHigh = ""
160       tMaleLow = ""
170       tFemaleHigh = ""
180       tFemaleLow = ""
190       tfr(0) = ""
200       tfr(1) = ""
210       tfr(2) = ""
220       tfr(3) = ""
230       cH = 0
240       cS = 0
250       cL = 0
260       cO = 0
270       cG = 0
280       cJ = 0
290       cInUse = 0
300       cEOD = 0
310       chkPrintable = 0
320       txtCheckTime = ""
330       chkLowLess.Value = 0
340       chkHighGreater.Value = 0

350       SampleType = ListCodeFor("ST", cmbSample)


360       If Trim(FromDays(0) & "") = "" Then FromDays(0) = 0
370       If Trim(ToDays(0) & "") = "" Then ToDays(0) = 43830


380       sql = "SELECT * from BioTestDefinitions WHERE " & _
                "LongName = '" & lstParameter & "' " & _
                "and AgeFromDays = '" & FromDays(0) & "' " & _
                "and AgeToDays = '" & ToDays(0) & "' " & _
                "and SampleType = '" & SampleType & "'"

390       Set tb = New Recordset
400       RecOpenServer 0, tb, sql
410       With tb

420           If Not .EOF Then
430               tImmuliteCode = Trim$(!immunocode & "")
440               tPriority = !PrintPriority
450               cKnown = IIf(!KnownToAnalyser, 1, 0)
460               If !DoDelta & "" <> "" Then oDelta = IIf(!DoDelta, 1, 0)
470               tDP = !DP
480               tdelta = !DeltaLimit & ""
490               If Trim(!Units) & "" <> "" Then
500                   For n = 0 To cmbUnits.ListCount
510                       If !Units & "" = cmbUnits.List(n) Then
520                           loading = True
530                           cmbUnits.ListIndex = n
540                           loading = False
550                       End If
560                   Next
570               End If

580               lblBarCode = !BarCode & ""

590               tPlausibleHigh = !PlausibleHigh
600               tPlausibleLow = !PlausibleLow
610               lblHost = !Code
620               tMaleHigh = !MaleHigh & ""
630               tMaleLow = !MaleLow & ""
640               tFemaleHigh = !FemaleHigh & ""
650               tFemaleLow = !FemaleLow & ""
660               tfr(0) = !FlagMaleHigh & ""
670               tfr(1) = !FlagMaleLow & ""
680               tfr(2) = !FlagFemaleHigh & ""
690               tfr(3) = !FlagFemaleLow & ""
700               cH = IIf(!h, 1, 0)
710               cS = IIf(!s, 1, 0)
720               cL = IIf(!l, 1, 0)
730               cO = IIf(!o, 1, 0)
740               cG = IIf(!g, 1, 0)
750               cJ = IIf(!J, 1, 0)
760               cInUse = IIf(!InUse, 1, 0)
770               chkPrintable = IIf(!Printable, 1, 0)
780               chkLowLess.Value = IIf(IsNull(!ShowLessThan), 0, !ShowLessThan)
790               chkHighGreater.Value = IIf(IsNull(!ShowMoreThan), 0, !ShowMoreThan)

800               If Trim(!Eod) & "" = "" Then cEOD = 0 Else cEOD = IIf(!Eod, 1, 0)
810               If Not IsNull(tb!CheckTime) Then
820                   txtCheckTime = tb!CheckTime
830               End If

840           End If
850       End With



860       Exit Sub

FillAllDetails_Error:

          Dim strES As String
          Dim intEL As Integer

870       intEL = Erl
880       strES = Err.Description
890       LogError "frmBioDefaults", "FillAllDetails", intEL, strES


End Sub

Private Sub FillDetails()
          Dim n As Long

          Dim tb As New Recordset
          Dim sql As String
          Dim SampleType As String
          Dim AgeNumber As Long



10        On Error GoTo FillDetails_Error

20        SampleType = ListCodeFor("ST", cmbSample)


30        AgeNumber = -1
40        grdAge.Col = 0
50        For AgeNumber = 1 To grdAge.Rows - 1
60            grdAge.Row = AgeNumber
70            If grdAge.CellBackColor = vbYellow Then
80                AgeNumber = AgeNumber - 1
90                Exit For
100           End If
110       Next
120       If AgeNumber = -1 Then
130           iMsg "SELECT Age Range", vbCritical
140           Exit Sub
150       End If

160       sql = "SELECT * from BioTestDefinitions WHERE " & _
                "LongName = '" & lstParameter & "' " & _
                "and AgeFromDays = '" & FromDays(AgeNumber) & "' " & _
                "and AgeToDays = '" & ToDays(AgeNumber) & "' " & _
                "and SampleType = '" & SampleType & "'"
170       Set tb = New Recordset
180       RecOpenServer 0, tb, sql
190       With tb

200           If Not .EOF Then
210               tPriority = !PrintPriority
220               cKnown = IIf(!KnownToAnalyser, 1, 0)
230               If !DoDelta & "" <> "" Then oDelta = IIf(!DoDelta, 1, 0)
240               tDP = !DP
250               tdelta = !DeltaLimit

260               If Trim(!Units) & "" <> "" Then
270                   For n = 0 To cmbUnits.ListCount
280                       If !Units & "" = cmbUnits.List(n) Then
290                           loading = True
300                           cmbUnits.ListIndex = n
310                           loading = False
320                       End If
330                   Next
340               End If
350               lblBarCode = !BarCode & ""

360               tMaleHigh = !MaleHigh
370               tMaleLow = !MaleLow
380               tFemaleHigh = !FemaleHigh
390               tFemaleLow = !FemaleLow
400               tfr(0) = !FlagMaleHigh
410               tfr(1) = !FlagMaleLow
420               tfr(2) = !FlagFemaleHigh
430               tfr(3) = !FlagFemaleLow
440               cH = IIf(!h, 1, 0)
450               cS = IIf(!s, 1, 0)
460               cL = IIf(!l, 1, 0)
470               cO = IIf(!o, 1, 0)
480               cG = IIf(!g, 1, 0)
490               cJ = IIf(!J, 1, 0)

500           End If
510       End With



520       Exit Sub

FillDetails_Error:

          Dim strES As String
          Dim intEL As Integer

530       intEL = Erl
540       strES = Err.Description
550       LogError "frmBioDefaults", "FillDetails", intEL, strES, sql


End Sub

Private Sub FilllParameter()

          Dim tb As New Recordset
          Dim sql As String
          Dim Analyser As String
          Dim SampleType As String
          Dim A2 As String

10        On Error GoTo FilllParameter_Error

20        Analyser = ListCodeFor("BioAnalysers", cmbAnalyser.Text)


30        SampleType = ListCodeFor("ST", cmbSample)

40        lstParameter.Clear

50        sql = "SELECT distinct LongName, PrintPriority from BioTestDefinitions WHERE " & _
                "SampleType = '" & SampleType & "' " & _
                "and Analyser = '" & Analyser & "' " & _
                "order by PrintPriority"

60        Set tb = New Recordset
70        RecOpenServer 0, tb, sql
80        Do While Not tb.EOF
90            lstParameter.AddItem tb!LongName
100           tb.MoveNext
110       Loop

120       lstParameter.ListIndex = -1

130       Exit Sub

FilllParameter_Error:

          Dim strES As String
          Dim intEL As Integer

140       intEL = Erl
150       strES = Err.Description
160       LogError "frmBioDefaults", "FilllParameter", intEL, strES, sql

End Sub

Private Sub FillSample()
          Dim sql As String
          Dim tb As New Recordset



10        On Error GoTo FillSample_Error

20        sql = "SELECT * from lists WHERE listtype = 'ST' order by listorder"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        Do While Not tb.EOF
60            cmbSample.AddItem Trim(tb!Text)
70            tb.MoveNext
80        Loop

90        cmbSample.ListIndex = 0




100       Exit Sub

FillSample_Error:

          Dim strES As String
          Dim intEL As Integer

110       intEL = Erl
120       strES = Err.Description
130       LogError "frmBioDefaults", "FillSample", intEL, strES, sql


End Sub

Private Sub FillUnits()
          Dim sql As String
          Dim tb As New Recordset



10        On Error GoTo FillUnits_Error

20        sql = "SELECT * from lists WHERE listtype = 'UN'"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        Do While Not tb.EOF
60            cmbUnits.AddItem Trim(tb!Text)
70            tb.MoveNext
80        Loop

90        cmbUnits.AddItem ""

100       Exit Sub



110       Exit Sub

FillUnits_Error:

          Dim strES As String
          Dim intEL As Integer

120       intEL = Erl
130       strES = Err.Description
140       LogError "frmBioDefaults", "FillUnits", intEL, strES, sql


End Sub

Private Sub Form_Activate()

10        On Error GoTo Form_Activate_Error

20        If Activated Then Exit Sub

30        Activated = True
          
40        FillSample
50        FilllParameter
60        FillUnits

70        Exit Sub

Form_Activate_Error:

          Dim strES As String
          Dim intEL As Integer

80        intEL = Erl
90        strES = Err.Description
100       LogError "frmBioDefaults", "Form_Activate", intEL, strES


End Sub

Private Sub Form_Load()

          Dim A2 As String

10        On Error GoTo Form_Load_Error

20        Activated = False
30        FillGenericList cmbAnalyser, "BioAnalysers"
40        If cmbAnalyser.ListCount > 0 Then cmbAnalyser.ListIndex = 0


50        grdAge.Font.Bold = True

60        Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

70        intEL = Erl
80        strES = Err.Description
90        LogError "frmBioDefaults", "Form_Load", intEL, strES

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
60        LogError "frmBioDefaults", "Form_Unload", intEL, strES


End Sub

Private Sub grdAge_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

          Dim n As Long
          Dim ySave As Long

10        On Error GoTo grdAge_MouseUp_Error

20        If grdAge.MouseRow = 0 Then Exit Sub

30        ySave = grdAge.Row

40        grdAge.Col = 0
50        For n = 1 To grdAge.Rows - 1
60            grdAge.Row = n
70            grdAge.CellBackColor = 0
80            grdAge.CellForeColor = 0
90        Next

100       grdAge.Row = ySave
110       grdAge.CellBackColor = vbYellow
120       grdAge.CellForeColor = vbBlue

130       FillDetails

140       Exit Sub

grdAge_MouseUp_Error:

          Dim strES As String
          Dim intEL As Integer

150       intEL = Erl
160       strES = Err.Description
170       LogError "frmBioDefaults", "grdAge_MouseUp", intEL, strES


End Sub



Private Sub lblBarCode_Click()

10        On Error GoTo lblBarCode_Click_Error

20        lblBarCode = iBOX("Scan Bar Code", , lblBarCode)

30        cmdUpdate.Enabled = True
40        lstParameter.Enabled = False

50        Exit Sub

lblBarCode_Click_Error:

          Dim strES As String
          Dim intEL As Integer

60        intEL = Erl
70        strES = Err.Description
80        LogError "frmBioDefaults", "lblBarCode_Click", intEL, strES


End Sub

Private Sub lstParameter_Click()



10        On Error GoTo lstParameter_Click_Error

20        SSTab1.Enabled = True

30        FillAges

40        FillAllDetails

50        cmdAmendAgeRange.Enabled = True

60        Exit Sub

lstParameter_Click_Error:

          Dim strES As String
          Dim intEL As Integer

70        intEL = Erl
80        strES = Err.Description
90        LogError "frmBioDefaults", "lstParameter_Click", intEL, strES


End Sub

Private Sub odelta_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        On Error GoTo odelta_MouseUp_Error

20        cmdUpdate.Enabled = True
30        lstParameter.Enabled = False

40        Exit Sub

odelta_MouseUp_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmBioDefaults", "odelta_MouseUp", intEL, strES


End Sub

Private Sub optAnalyser_Click(Index As Integer)

    On Error GoTo optAnalyser_Click_Error

    FilllParameter
    With grdAge
        .Rows = 2
        .AddItem ""
        .RemoveItem 1
    End With

    If Index = 1 Then
        If optAnalyser(1).Caption = "Immulite" Then
            fraImmulite.Visible = True
        End If
    Else
        fraImmulite.Visible = False
    End If

    Exit Sub

optAnalyser_Click_Error:

    Dim strES As String
    Dim intEL As Integer

    intEL = Erl
    strES = Err.Description
    LogError "frmBioDefaults", "optAnalyser_Click", intEL, strES


End Sub

Private Sub orm_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        On Error GoTo orm_MouseUp_Error

20        cmdUpdate.Enabled = True
30        lstParameter.Enabled = False

40        Exit Sub

orm_MouseUp_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmBioDefaults", "orm_MouseUp", intEL, strES


End Sub

Private Sub SaveCommonDetails()

          Dim sql As String
          Dim SampleType As String



10        On Error GoTo SaveCommonDetails_Error

20        SampleType = ListCodeFor("ST", cmbSample)



30        sql = "UPDATE BioTestDefinitions SET " & _
                "H = '" & IIf(cH = 1, 1, 0) & "', " & _
                "S = '" & IIf(cS = 1, 1, 0) & "', " & _
                "L = '" & IIf(cL = 1, 1, 0) & "', " & _
                "O = '" & IIf(cO = 1, 1, 0) & "', " & _
                "G = '" & IIf(cG = 1, 1, 0) & "', " & _
                "J = '" & IIf(cJ = 1, 1, 0) & "', KnownToAnalyser = '" & IIf(cKnown = 1, 1, 0) & "', " & _
                "PrintPriority = '" & Val(tPriority) & "', " & _
                "DP = '" & Val(tDP) & "', " & _
                "DoDelta = '" & oDelta & "', " & _
                "DeltaLimit ='" & Val(tdelta) & "', " & _
                "Units = '" & cmbUnits & "', " & _
                "BarCode = '" & lblBarCode & "', " & _
                "CheckTime = " & IIf(txtCheckTime = "", 1, Val(txtCheckTime)) & ", " & _
                "ShowLessThan = " & chkLowLess.Value & ", " & _
                "ShowMoreThan = " & chkHighGreater.Value & " " & _
                "WHERE LongName = '" & lstParameter & "' " & _
                "and Hospital = '" & HospName(0) & "'"

40        Cnxn(0).Execute sql



50        Exit Sub

SaveCommonDetails_Error:

          Dim strES As String
          Dim intEL As Integer

60        intEL = Erl
70        strES = Err.Description
80        LogError "frmBioDefaults", "SaveCommonDetails", intEL, strES


End Sub

Private Sub SaveNormals()

          Dim sql As String
          Dim AgeNumber As Long
          Dim SampleType As String



10        On Error GoTo SaveNormals_Error

20        AgeNumber = -1
30        grdAge.Col = 0
40        For AgeNumber = 1 To grdAge.Rows - 1
50            grdAge.Row = AgeNumber
60            If grdAge.CellBackColor = vbYellow Then
70                AgeNumber = AgeNumber - 1
80                Exit For
90            End If
100       Next
110       If AgeNumber = -1 Then
120           iMsg "SELECT Age Range", vbCritical
130           Exit Sub
140       End If

150       SampleType = ListCodeFor("ST", cmbSample)


160       sql = "UPDATE BioTestDefinitions " & _
                "Set MaleLow = '" & tMaleLow & "', " & _
                "MaleHigh = '" & tMaleHigh & "', " & _
                "FemaleLow = '" & tFemaleLow & "', " & _
                "FemaleHigh = '" & tFemaleHigh & "' WHERE " & _
                "LongName = '" & lstParameter & "' " & _
                "and AgeFromDays = '" & FromDays(AgeNumber) & "' " & _
                "and AgeToDays = '" & ToDays(AgeNumber) & "' " & _
                " " & _
                "and Hospital = '" & HospName(0) & "'"

170       Cnxn(0).Execute sql


180       Exit Sub

SaveNormals_Error:

          Dim strES As String
          Dim intEL As Integer

190       intEL = Erl
200       strES = Err.Description
210       LogError "frmBioDefaults", "SaveNormals", intEL, strES


End Sub

Private Sub tDelta_KeyUp(KeyCode As Integer, Shift As Integer)

10        On Error GoTo tDelta_KeyUp_Error

20        cmdUpdate.Enabled = True
30        lstParameter.Enabled = False

40        Exit Sub

tDelta_KeyUp_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmBioDefaults", "tDelta_KeyUp", intEL, strES


End Sub

Private Sub tDP_KeyUp(KeyCode As Integer, Shift As Integer)

10        On Error GoTo tDP_KeyUp_Error

20        cmdUpdate.Enabled = True
30        lstParameter.Enabled = False

40        Exit Sub

tDP_KeyUp_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmBioDefaults", "tDP_KeyUp", intEL, strES


End Sub

Private Sub tFemaleHigh_KeyUp(KeyCode As Integer, Shift As Integer)

10        On Error GoTo tFemaleHigh_KeyUp_Error

20        SaveNormals

30        Exit Sub

tFemaleHigh_KeyUp_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmBioDefaults", "tFemaleHigh_KeyUp", intEL, strES


End Sub

Private Sub tFemaleLow_KeyUp(KeyCode As Integer, Shift As Integer)

10        On Error GoTo tFemaleLow_KeyUp_Error

20        SaveNormals

30        Exit Sub

tFemaleLow_KeyUp_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmBioDefaults", "tFemaleLow_KeyUp", intEL, strES


End Sub



Private Sub tfr_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

          Dim sql As String
          Dim AgeNumber As Long
          Dim SampleType As String
          Dim n As Long


10        On Error GoTo tfr_KeyUp_Error

20        If Not IsNumeric(tfr(Index).Text) And tfr(Index).Text <> "" Then
30            iMsg "Wrong input"
40            Exit Sub
50        End If


60        If Index = 2 Or Index = 3 Then
70            tfr(Index - 2) = tfr(Index)
80        End If

90        AgeNumber = -1
100       grdAge.Col = 0
110       For n = 1 To grdAge.Rows - 1
120           grdAge.Row = n
130           If grdAge.CellBackColor = vbYellow Then
140               AgeNumber = grdAge.TextMatrix(n, 0)
150               Exit For
160           End If
170       Next
180       If AgeNumber = -1 Then
190           iMsg "SELECT Age Range", vbCritical
200           Exit Sub
210       End If

220       SampleType = ListCodeFor("ST", cmbSample)

230       sql = "UPDATE BioTestDefinitions " & _
                "Set FlagMaleLow = '" & tfr(1) & "', " & _
                "FlagMaleHigh = '" & tfr(0) & "', " & _
                "FlagFemaleLow = '" & tfr(3) & "', " & _
                "FlagFemaleHigh = '" & tfr(2) & "' WHERE " & _
                "LongName = '" & lstParameter & "' " & _
                "and AgeFromDays = '" & FromDays(AgeNumber) & "' " & _
                "and AgeToDays = '" & ToDays(AgeNumber) & "' " & _
                "and SampleType = '" & SampleType & "' " & _
                "and Hospital = '" & HospName(0) & "'"

240       Cnxn(0).Execute sql


250       Exit Sub

tfr_KeyUp_Error:

          Dim strES As String
          Dim intEL As Integer

260       intEL = Erl
270       strES = Err.Description
280       LogError "frmBioDefaults", "tfr_KeyUp", intEL, strES, sql


End Sub

Private Sub tImmuliteCode_KeyUp(KeyCode As Integer, Shift As Integer)

          Dim sql As String



10        On Error GoTo tImmuliteCode_KeyUp_Error

20        sql = "UPDATE BioTestDefinitions " & _
                "Set ImmunoCode = '" & Trim$(tImmuliteCode) & "' " & _
                "WHERE LongName = '" & lstParameter & "' " & _
                "and Hospital = '" & HospName(0) & "'"

30        Cnxn(0).Execute sql




40        Exit Sub

tImmuliteCode_KeyUp_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmBioDefaults", "tImmuliteCode_KeyUp", intEL, strES, sql


End Sub

Private Sub tMaleHigh_KeyUp(KeyCode As Integer, Shift As Integer)

10        On Error GoTo tMaleHigh_KeyUp_Error

20        SaveNormals

30        Exit Sub

tMaleHigh_KeyUp_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmBioDefaults", "tMaleHigh_KeyUp", intEL, strES


End Sub

Private Sub tMaleLow_KeyUp(KeyCode As Integer, Shift As Integer)

10        On Error GoTo tMaleLow_KeyUp_Error

20        SaveNormals

30        Exit Sub

tMaleLow_KeyUp_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmBioDefaults", "tMaleLow_KeyUp", intEL, strES


End Sub

Private Sub tPlausibleHigh_KeyUp(KeyCode As Integer, Shift As Integer)

          Dim sql As String



10        On Error GoTo tPlausibleHigh_KeyUp_Error

20        sql = "UPDATE BioTestDefinitions " & _
                "Set PlausibleHigh = " & Val(tPlausibleHigh) & " " & _
                "WHERE LongName = '" & lstParameter & "' " & _
                "and Hospital = '" & HospName(0) & "'"

30        Cnxn(0).Execute sql




40        Exit Sub

tPlausibleHigh_KeyUp_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmBioDefaults", "tPlausibleHigh_KeyUp", intEL, strES, sql


End Sub

Private Sub tPlausibleLow_KeyUp(KeyCode As Integer, Shift As Integer)

          Dim sql As String



10        On Error GoTo tPlausibleLow_KeyUp_Error

20        sql = "UPDATE BioTestDefinitions " & _
                "Set PlausibleLow = " & Val(tPlausibleLow) & " " & _
                "WHERE LongName = '" & lstParameter & "' " & _
                "and Hospital = '" & HospName(0) & "'"

30        Cnxn(0).Execute sql




40        Exit Sub

tPlausibleLow_KeyUp_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmBioDefaults", "tPlausibleLow_KeyUp", intEL, strES, sql


End Sub

Private Sub tPriority_KeyUp(KeyCode As Integer, Shift As Integer)

10        On Error GoTo tPriority_KeyUp_Error

20        cmdUpdate.Enabled = True
30        lstParameter.Enabled = False

40        Exit Sub

tPriority_KeyUp_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmBioDefaults", "tPriority_KeyUp", intEL, strES


End Sub

Private Sub txtCheckTime_KeyPress(KeyAscii As Integer)
10        KeyAscii = VI(KeyAscii, Numeric_Only)
End Sub

Private Sub txtCheckTime_KeyUp(KeyCode As Integer, Shift As Integer)

10        On Error GoTo txtCheckTime_KeyUp_Error

20        cmdUpdate.Enabled = True
30        lstParameter.Enabled = False

40        Exit Sub

txtCheckTime_KeyUp_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmBioDefaults", "txtCheckTime_KeyUp", intEL, strES

End Sub

Private Sub UpDown1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        On Error GoTo UpDown1_MouseUp_Error

20        cmdUpdate.Enabled = True
30        lstParameter.Enabled = False

40        Exit Sub

UpDown1_MouseUp_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmBioDefaults", "UpDown1_MouseUp", intEL, strES


End Sub
