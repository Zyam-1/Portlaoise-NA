VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmEndDefaults 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Endocrinology Parameter Definitions"
   ClientHeight    =   8250
   ClientLeft      =   225
   ClientTop       =   765
   ClientWidth     =   10605
   Icon            =   "frmEndDefaults.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8250
   ScaleWidth      =   10605
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Age Ranges"
      Height          =   2115
      Left            =   2925
      TabIndex        =   53
      Top             =   2940
      Width           =   6795
      Begin VB.Frame fraImmulite 
         Caption         =   "Immulite Code"
         Height          =   615
         Left            =   5400
         TabIndex        =   63
         Top             =   90
         Visible         =   0   'False
         Width           =   1365
         Begin VB.TextBox tImmuliteCode 
            Height          =   285
            Left            =   210
            TabIndex        =   64
            Top             =   240
            Width           =   945
         End
      End
      Begin VB.CheckBox cInUse 
         Alignment       =   1  'Right Justify
         Caption         =   "InUse"
         Height          =   195
         Left            =   4290
         TabIndex        =   61
         Top             =   1230
         Width           =   1635
      End
      Begin VB.CheckBox cEOD 
         Alignment       =   1  'Right Justify
         Caption         =   "End Of Day"
         Height          =   195
         Left            =   4290
         TabIndex        =   60
         Top             =   1020
         Width           =   1635
      End
      Begin VB.CommandButton bAmendAgeRange 
         Caption         =   "Amend Age Range"
         Enabled         =   0   'False
         Height          =   525
         Left            =   4185
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   300
         Width           =   1155
      End
      Begin MSFlexGridLib.MSFlexGrid g 
         Height          =   1725
         Left            =   120
         TabIndex        =   55
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
      Begin VB.Label lblHost 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5070
         TabIndex        =   59
         Top             =   1590
         Width           =   765
      End
      Begin VB.Label Label21 
         Caption         =   "Host Code"
         Height          =   315
         Left            =   4260
         TabIndex        =   58
         Top             =   1620
         Width           =   795
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Analyser"
      Height          =   1305
      Left            =   8100
      TabIndex        =   44
      Top             =   120
      Width           =   1905
      Begin VB.OptionButton optAnalyser 
         Caption         =   "ADM"
         Height          =   195
         Index           =   3
         Left            =   210
         TabIndex        =   71
         Top             =   240
         Value           =   -1  'True
         Width           =   1545
      End
      Begin VB.OptionButton optAnalyser 
         Caption         =   "None"
         Height          =   195
         Index           =   2
         Left            =   210
         TabIndex        =   47
         Top             =   960
         Width           =   1545
      End
      Begin VB.OptionButton optAnalyser 
         Caption         =   "Immulite"
         Height          =   195
         Index           =   1
         Left            =   210
         TabIndex        =   46
         Top             =   720
         Width           =   1545
      End
      Begin VB.OptionButton optAnalyser 
         Caption         =   "Roche"
         Height          =   195
         Index           =   0
         Left            =   210
         TabIndex        =   45
         Top             =   480
         Width           =   1545
      End
   End
   Begin VB.Frame Frame5 
      Height          =   1665
      Left            =   2925
      TabIndex        =   34
      Top             =   1200
      Width           =   5055
      Begin VB.TextBox txtCheckTime 
         Height          =   285
         Left            =   3960
         MaxLength       =   5
         TabIndex        =   66
         Top             =   1200
         Width           =   495
      End
      Begin VB.CheckBox oRM 
         Alignment       =   1  'Right Justify
         Caption         =   "Include in Running Mean"
         Height          =   195
         Left            =   2850
         TabIndex        =   37
         Top             =   270
         Width           =   2085
      End
      Begin VB.CheckBox oDelta 
         Alignment       =   1  'Right Justify
         Caption         =   "Do Delta"
         Height          =   225
         Left            =   3990
         TabIndex        =   36
         Top             =   540
         Width           =   945
      End
      Begin VB.CheckBox cKnown 
         Alignment       =   1  'Right Justify
         Caption         =   "Known to Analyser"
         Height          =   195
         Left            =   570
         TabIndex        =   35
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
         BuddyDispid     =   196626
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
         Top             =   1230
         Width           =   420
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Check Time"
         Height          =   195
         Left            =   3060
         TabIndex        =   67
         Top             =   1230
         Width           =   855
      End
      Begin VB.Label lblShortName 
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
         Left            =   30
         TabIndex        =   62
         Top             =   90
         Width           =   975
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
      Height          =   990
      Left            =   5160
      TabIndex        =   19
      Top             =   120
      Width           =   2775
      Begin VB.TextBox tUnits 
         Height          =   285
         Left            =   810
         TabIndex        =   20
         Top             =   150
         Width           =   1515
      End
      Begin VB.Label lblBarCode 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   810
         TabIndex        =   56
         Top             =   480
         Width           =   1515
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Units"
         Height          =   195
         Left            =   330
         TabIndex        =   22
         Top             =   180
         Width           =   360
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Bar Code"
         Height          =   195
         Left            =   60
         TabIndex        =   21
         Top             =   540
         Width           =   660
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Sample Type"
      Height          =   1005
      Left            =   2880
      TabIndex        =   18
      Top             =   120
      Width           =   2235
      Begin VB.ComboBox cmbSample 
         Height          =   315
         Left            =   270
         Style           =   2  'Dropdown List
         TabIndex        =   65
         Top             =   180
         Width           =   1680
      End
      Begin VB.ComboBox cCat 
         Height          =   315
         Left            =   240
         TabIndex        =   57
         Top             =   630
         Width           =   1665
      End
   End
   Begin VB.CommandButton bexit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   750
      Left            =   8235
      Picture         =   "frmEndDefaults.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1605
      Width           =   1170
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
      Height          =   7845
      IntegralHeight  =   0   'False
      Left            =   120
      TabIndex        =   1
      Top             =   210
      Width           =   2415
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2985
      Left            =   2925
      TabIndex        =   0
      Top             =   5130
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
      TabPicture(0)   =   "frmEndDefaults.frx":0614
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
      Tab(0).Control(9)=   "chkLowLess"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "chkHighGreater"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "Flag Range"
      TabPicture(1)   =   "frmEndDefaults.frx":0630
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
      TabPicture(2)   =   "frmEndDefaults.frx":064C
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
      TabPicture(3)   =   "frmEndDefaults.frx":0668
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label8"
      Tab(3).Control(1)=   "Label9"
      Tab(3).Control(2)=   "Label10"
      Tab(3).Control(3)=   "tPlausibleHigh"
      Tab(3).Control(4)=   "tPlausibleLow"
      Tab(3).ControlCount=   5
      Begin VB.CheckBox chkHighGreater 
         Caption         =   "Report range as ""> low"" if high range is 9999"
         Height          =   255
         Left            =   1080
         TabIndex        =   70
         Top             =   2640
         Width           =   4755
      End
      Begin VB.CheckBox chkLowLess 
         Caption         =   "Report range as ""< high"" if low range is 0.0"
         Height          =   255
         Left            =   1080
         TabIndex        =   69
         Top             =   2340
         Width           =   4755
      End
      Begin VB.TextBox tPlausibleLow 
         Height          =   285
         Left            =   -72600
         TabIndex        =   51
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox tPlausibleHigh 
         Height          =   285
         Left            =   -72600
         TabIndex        =   50
         Top             =   990
         Width           =   1215
      End
      Begin VB.CheckBox cJ 
         Caption         =   "Jaundiced"
         Height          =   225
         Left            =   -72210
         TabIndex        =   43
         Top             =   1770
         Width           =   1095
      End
      Begin VB.CheckBox cG 
         Alignment       =   1  'Right Justify
         Caption         =   "Grossly Haemolysed"
         Height          =   225
         Left            =   -74100
         TabIndex        =   42
         Top             =   1770
         Width           =   1785
      End
      Begin VB.CheckBox cO 
         Caption         =   "Old"
         Height          =   225
         Left            =   -72210
         TabIndex        =   41
         Top             =   1530
         Width           =   585
      End
      Begin VB.CheckBox cL 
         Caption         =   "Lipaemic"
         Height          =   225
         Left            =   -72210
         TabIndex        =   40
         Top             =   1290
         Width           =   975
      End
      Begin VB.CheckBox cS 
         Alignment       =   1  'Right Justify
         Caption         =   "Slightly Haemolysed"
         Height          =   225
         Left            =   -74100
         TabIndex        =   39
         Top             =   1530
         Width           =   1785
      End
      Begin VB.CheckBox cH 
         Alignment       =   1  'Right Justify
         Caption         =   "Haemolysed"
         Height          =   225
         Left            =   -73500
         TabIndex        =   38
         Top             =   1290
         Width           =   1185
      End
      Begin VB.TextBox tfr 
         Height          =   315
         Index           =   0
         Left            =   -71640
         TabIndex        =   26
         Top             =   1200
         Width           =   915
      End
      Begin VB.TextBox tfr 
         Height          =   315
         Index           =   2
         Left            =   -73170
         TabIndex        =   25
         Top             =   1200
         Width           =   915
      End
      Begin VB.TextBox tfr 
         Height          =   315
         Index           =   1
         Left            =   -71640
         TabIndex        =   24
         Top             =   1800
         Width           =   915
      End
      Begin VB.TextBox tfr 
         Height          =   315
         Index           =   3
         Left            =   -73170
         TabIndex        =   23
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
         TabIndex        =   52
         Top             =   2430
         Width           =   3930
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Low"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   -73080
         TabIndex        =   49
         Top             =   1590
         Width           =   300
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "High"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   -73110
         TabIndex        =   48
         Top             =   1020
         Width           =   330
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Result values outside this range will be flagged as High or Low"
         Height          =   195
         Left            =   -74190
         TabIndex        =   33
         Top             =   2520
         Width           =   4410
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "These are the normal range values printed on the report forms."
         Height          =   195
         Left            =   1020
         TabIndex        =   32
         Top             =   480
         Width           =   4395
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Do not print the result if the sample is :-"
         Height          =   195
         Left            =   -74070
         TabIndex        =   31
         Top             =   750
         Width           =   2730
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Female"
         Height          =   195
         Left            =   -73020
         TabIndex        =   30
         Top             =   930
         Width           =   510
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Male"
         Height          =   195
         Left            =   -71370
         TabIndex        =   29
         Top             =   900
         Width           =   345
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "High"
         Height          =   195
         Left            =   -73950
         TabIndex        =   28
         Top             =   1290
         Width           =   330
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Low"
         Height          =   195
         Left            =   -73920
         TabIndex        =   27
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
Attribute VB_Name = "frmEndDefaults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Activated As Boolean

Private FromDays() As Long
Private ToDays() As Long
Private Sub FillSample()
          Dim sql As String
          Dim tb As New Recordset


10        On Error GoTo FillSample_Error

20        sql = "SELECT * from lists WHERE listtype = 'ST'"
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
130       LogError "frmEndDefaults", "FillSample", intEL, strES, sql


End Sub
Private Sub bAmendAgeRange_Click()


10        On Error GoTo bAmendAgeRange_Click_Error

20        With frmAges
30            .Analyte = lstParameter
40            .SampleType = ListCodeFor("ST", cmbSample)
50            .Discipline = "Endocrinology"
60            .Cat = cCat
70            .Show 1
80        End With

90        FillAges

100       FillAllDetails


110       Exit Sub

bAmendAgeRange_Click_Error:

          Dim strES As String
          Dim intEL As Integer



120       intEL = Erl
130       strES = Err.Description
140       LogError "frmEndDefaults", "bAmendAgeRange_Click", intEL, strES


End Sub

Private Sub bexit_Click()

10        Unload Me

End Sub

Private Sub cCat_Click()


10        On Error GoTo cCat_Click_Error

20        If lstParameter.ListIndex > -1 Then
30            FillAllDetails
40        End If
50        Exit Sub

cCat_Click_Error:

          Dim strES As String
          Dim intEL As Integer



60        intEL = Erl
70        strES = Err.Description
80        LogError "frmEndDefaults", "cCat_Click", intEL, strES


End Sub

Private Sub cEOD_Click()
          Dim sql As String


10        On Error GoTo cEOD_Click_Error

20        If lblHost = "" Then Exit Sub


30        If cEOD.Value = 1 Then
40            sql = "UPDATE endtestdefinitions set eod = 1 WHERE longname = '" & lstParameter & "'"
50            Cnxn(0).Execute sql
60        ElseIf cEOD.Value = 0 Then
70            sql = "UPDATE endtestdefinitions set eod = 0 WHERE longname = '" & lstParameter & "'"
80            Cnxn(0).Execute sql
90        End If




100       Exit Sub

cEOD_Click_Error:

          Dim strES As String
          Dim intEL As Integer



110       intEL = Erl
120       strES = Err.Description
130       LogError "frmEndDefaults", "cEOD_Click", intEL, strES, sql


End Sub

Private Sub cG_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        On Error GoTo cG_MouseUp_Error

20        SaveCommonDetails

30        Exit Sub

cG_MouseUp_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmEndDefaults", "cG_MouseUp", intEL, strES


End Sub

Private Sub cH_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        On Error GoTo cH_MouseUp_Error

20        SaveCommonDetails

30        Exit Sub

cH_MouseUp_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmEndDefaults", "cH_MouseUp", intEL, strES


End Sub

Private Sub chkHighGreater_Click()

10        On Error GoTo chkHighGreater_Click_Error

20        SaveCommonDetails

30        Exit Sub

chkHighGreater_Click_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmEndDefaults", "chkHighGreater_Click", intEL, strES

End Sub

Private Sub chkLowLess_Click()

10        On Error GoTo chkLowLess_Click_Error

20        SaveCommonDetails

30        Exit Sub

chkLowLess_Click_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmEndDefaults", "chkLowLess_Click", intEL, strES

End Sub

Private Sub cInUse_Click()
          Dim sql As String


10        On Error GoTo cInUse_Click_Error

20        If lblHost = "" Then Exit Sub


30        If cInUse.Value = 1 Then
40            sql = "UPDATE endtestdefinitions set inuse = 1 WHERE code = '" & lblHost & "'"
50            Cnxn(0).Execute sql
60        ElseIf cInUse.Value = 0 Then
70            sql = "UPDATE endtestdefinitions set inuse = 0 WHERE code = '" & lblHost & "'"
80            Cnxn(0).Execute sql
90        End If



100       Exit Sub

cInUse_Click_Error:

          Dim strES As String
          Dim intEL As Integer



110       intEL = Erl
120       strES = Err.Description
130       LogError "frmEndDefaults", "cInUse_Click", intEL, strES, sql


End Sub

Private Sub cJ_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        On Error GoTo cJ_MouseUp_Error

20        SaveCommonDetails

30        Exit Sub

cJ_MouseUp_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmEndDefaults", "cJ_MouseUp", intEL, strES


End Sub

Private Sub cKnown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        On Error GoTo cKnown_MouseUp_Error

20        SaveCommonDetails

30        Exit Sub

cKnown_MouseUp_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmEndDefaults", "cKnown_MouseUp", intEL, strES


End Sub

Private Sub cL_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        On Error GoTo cL_MouseUp_Error

20        SaveCommonDetails

30        Exit Sub

cL_MouseUp_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmEndDefaults", "cL_MouseUp", intEL, strES


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
60        LogError "frmEndDefaults", "cmbSample_Click", intEL, strES


End Sub

Private Sub cO_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        On Error GoTo cO_MouseUp_Error

20        SaveCommonDetails

30        Exit Sub

cO_MouseUp_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmEndDefaults", "cO_MouseUp", intEL, strES


End Sub

Private Sub cS_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        On Error GoTo cS_MouseUp_Error

20        SaveCommonDetails

30        Exit Sub

cS_MouseUp_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmEndDefaults", "cS_MouseUp", intEL, strES


End Sub

Private Sub FillAges()

          Dim tb As New Recordset
          Dim s As String
          Dim SampleType As String
          Dim n As Long
          Dim sql As String
          Dim Cat As String



10        On Error GoTo FillAges_Error

20        If cCat = "" Then Cat = "Default" Else Cat = cCat

30        ClearFGrid g

40        SampleType = ListCodeFor("ST", cmbSample)


50        ReDim FromDays(0 To 0)
60        ReDim ToDays(0 To 0)


70        sql = "SELECT * from EndTestDefinitions WHERE " & _
                "LongName = '" & lstParameter & "' " & _
                "and SampleType = '" & SampleType & "' and category = '" & Cat & "' " & _
                "Order by AgeFromDays"
80        Set tb = New Recordset
90        RecOpenClient 0, tb, sql

100       If tb.EOF Then Exit Sub
110       ReDim FromDays(0 To tb.recordCount - 1)
120       ReDim ToDays(0 To tb.recordCount - 1)
130       n = 0
140       Do While Not tb.EOF
150           FromDays(n) = tb!AgeFromDays
160           ToDays(n) = tb!AgeToDays
170           s = Format$(n) & vbTab & _
                  dmyFromCount(FromDays(n)) & vbTab & _
                  dmyFromCount(ToDays(n))
180           g.AddItem s
190           n = n + 1
200           tb.MoveNext
210       Loop

220       FixG g

230       g.Col = 0
240       g.Row = 1
250       g.CellBackColor = vbYellow
260       g.CellForeColor = vbBlue




270       Exit Sub

FillAges_Error:

          Dim strES As String
          Dim intEL As Integer



280       intEL = Erl
290       strES = Err.Description
300       LogError "frmEndDefaults", "FillAges", intEL, strES, sql


End Sub

Private Sub FillAllDetails()

          Dim tb As New Recordset
          Dim sql As String
          Dim SampleType As String


10        On Error GoTo FillAllDetails_Error

20        tImmuliteCode = ""
30        tPriority = ""
40        cKnown = 0
50        oDelta = 0
60        tDP = 0
70        tdelta = 0
80        tUnits = ""
90        lblBarCode = ""
100       lblShortName = ""
110       tPlausibleHigh = ""
120       tPlausibleLow = ""

130       tMaleHigh = ""
140       tMaleLow = ""
150       tFemaleHigh = ""
160       tFemaleLow = ""
170       tfr(0) = ""
180       tfr(1) = ""
190       tfr(2) = ""
200       tfr(3) = ""
210       cH = 0
220       cS = 0
230       cL = 0
240       cO = 0
250       cG = 0
260       cJ = 0
270       lblHost = ""
280       cInUse = 0
290       cEOD = 0
300       txtCheckTime = ""
310       chkLowLess.Value = 0
320       chkHighGreater.Value = 0

330       If cCat = "" Then cCat = "Default"

340       SampleType = ListCodeFor("ST", cmbSample)


350       sql = "SELECT distinct * from EndTestDefinitions WHERE " & _
                "LongName = '" & lstParameter & "' " & _
                "and AgeFromDays = '" & FromDays(0) & "' " & _
                "and AgeToDays = '" & ToDays(0) & "' " & _
                "and SampleType = '" & SampleType & "' and category = 'Default'"
360       Set tb = New Recordset
370       RecOpenServer 0, tb, sql
380       With tb

390           If Not .EOF Then
400               lblShortName = tb!ShortName & ""
410               tImmuliteCode = Trim$(!immunocode & "")
420               tPriority = !PrintPriority
430               cKnown = IIf(IsNull(!KnownToAnalyser), 0, Abs(!KnownToAnalyser = True))
440               If IsNull(!DoDelta) Then
450                   oDelta.Value = 0
460               Else
470                   If !DoDelta Then
480                       oDelta.Value = 1
490                   Else
500                       oDelta.Value = 0
510                   End If
520               End If
530               tDP = !DP
540               tdelta = !DeltaLimit & ""
550               tUnits = !Units
560               lblBarCode = !BarCode & ""
570               tPlausibleHigh = !PlausibleHigh
580               tPlausibleLow = !PlausibleLow

590               tMaleHigh = !MaleHigh
600               tMaleLow = !MaleLow
610               tFemaleHigh = !FemaleHigh
620               tFemaleLow = !FemaleLow
630               tfr(0) = !FlagMaleHigh
640               tfr(1) = !FlagMaleLow
650               tfr(2) = !FlagFemaleHigh
660               tfr(3) = !FlagFemaleLow
670               cH = IIf(!h, 1, 0)
680               cS = IIf(!s, 1, 0)
690               cL = IIf(!l, 1, 0)
700               cO = IIf(!o, 1, 0)
710               cG = IIf(!g, 1, 0)
720               cJ = IIf(!J, 1, 0)
730               If IsNull(!CheckTime) Then
740                   txtCheckTime = ""
750               Else
760                   txtCheckTime = !CheckTime
770               End If

780               cInUse = IIf(!InUse, 1, 0)
790               If Trim(!Eod) & "" = "" Then cEOD = 0 Else cEOD = IIf(!Eod, 1, 0)
800               lblHost = tb!ShortName
810               chkLowLess.Value = IIf(IsNull(!ShowLessThan), 0, !ShowLessThan)
820               chkHighGreater.Value = IIf(IsNull(!ShowMoreThan), 0, !ShowMoreThan)
830           End If
840       End With

850       If cCat <> "Default" Then
860           sql = "SELECT distinct * from EndTestDefinitions WHERE " & _
                    "LongName = '" & lstParameter & "' " & _
                    "and AgeFromDays = '" & FromDays(0) & "' " & _
                    "and AgeToDays = '" & ToDays(0) & "' " & _
                    "and SampleType = '" & SampleType & "' " & _
                    "and UPPER(category) = '" & UCase(cCat) & "'"
870           Set tb = New Recordset
880           RecOpenClient 0, tb, sql
890           With tb
900               If Not .EOF Then
910                   tPlausibleHigh = !PlausibleHigh & ""
920                   tPlausibleLow = !PlausibleLow & ""
930                   tMaleHigh = !MaleHigh & ""
940                   tMaleLow = !MaleLow & ""
950                   tFemaleHigh = !FemaleHigh & ""
960                   tFemaleLow = !FemaleLow & ""
970                   tfr(0) = !FlagMaleHigh & ""
980                   tfr(1) = !FlagMaleLow & ""
990                   tfr(2) = !FlagFemaleHigh & ""
1000                  tfr(3) = !FlagFemaleLow & ""
1010                  chkLowLess.Value = IIf(IsNull(!ShowLessThan), 0, !ShowLessThan)
1020                  chkHighGreater.Value = IIf(IsNull(!ShowMoreThan), 0, !ShowMoreThan)
1030              End If
1040          End With
1050      End If

1060      Exit Sub




1070      Exit Sub

FillAllDetails_Error:

          Dim strES As String
          Dim intEL As Integer



1080      intEL = Erl
1090      strES = Err.Description
1100      LogError "frmEndDefaults", "FillAllDetails", intEL, strES, sql


End Sub

Private Sub FillCats()
          Dim sql As String
          Dim tb As New Recordset
          Dim n As Long
          Dim Found As Boolean


10        On Error GoTo FillCats_Error

20        cCat.Clear

30        sql = "SELECT * from categorys"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        Do While Not tb.EOF
70            cCat.AddItem Trim(tb!Cat)
80            tb.MoveNext
90        Loop

100       For n = 0 To cCat.ListCount
110           If UCase(cCat.List(n)) = "DEFAULT" Then Found = True
120       Next

130       If Found = False Then cCat.AddItem "Default", 0

140       If cCat.ListCount > 0 Then
150           cCat = "Default"
160       End If



170       Exit Sub

FillCats_Error:

          Dim strES As String
          Dim intEL As Integer



180       intEL = Erl
190       strES = Err.Description
200       LogError "frmEndDefaults", "FillCats", intEL, strES, sql


End Sub

Private Sub FillDetails()

          Dim tb As New Recordset
          Dim sql As String
          Dim SampleType As String
          Dim AgeNumber As Long


10        On Error GoTo FillDetails_Error

20        SampleType = ListCodeFor("ST", cmbSample)

30        AgeNumber = -1
40        g.Col = 0
50        For AgeNumber = 1 To g.Rows - 1
60            g.Row = AgeNumber
70            If g.CellBackColor = vbYellow Then
80                AgeNumber = AgeNumber - 1
90                Exit For
100           End If
110       Next
120       If AgeNumber = -1 Then
130           iMsg "SELECT Age Range", vbCritical
140           Exit Sub
150       End If

160       sql = "SELECT * from EndTestDefinitions WHERE " & _
                "LongName = '" & lstParameter & "' " & _
                "and AgeFromDays = '" & FromDays(AgeNumber) & "' " & _
                "and AgeToDays = '" & ToDays(AgeNumber) & "' " & _
                "and SampleType = '" & SampleType & "' and category = '" & cCat & "'"
170       Set tb = New Recordset
180       RecOpenServer 0, tb, sql
190       With tb

200           If Not .EOF Then
210               tPriority = !PrintPriority
220               cKnown = IIf(!KnownToAnalyser, 1, 0)
230               oDelta = IIf(!DoDelta, 1, 0)
240               tDP = !DP
250               tdelta = !DeltaLimit
260               tUnits = !Units
270               lblBarCode = !BarCode & ""
280               cCat = !Category
290               tMaleHigh = !MaleHigh
300               tMaleLow = !MaleLow
310               tFemaleHigh = !FemaleHigh
320               tFemaleLow = !FemaleLow
330               tfr(0) = !FlagMaleHigh
340               tfr(1) = !FlagMaleLow
350               tfr(2) = !FlagFemaleHigh
360               tfr(3) = !FlagFemaleLow
370               cH = IIf(!h, 1, 0)
380               cS = IIf(!s, 1, 0)
390               cL = IIf(!l, 1, 0)
400               cO = IIf(!o, 1, 0)
410               cG = IIf(!g, 1, 0)
420               cJ = IIf(!J, 1, 0)

430           End If
440       End With



450       Exit Sub

FillDetails_Error:

          Dim strES As String
          Dim intEL As Integer



460       intEL = Erl
470       strES = Err.Description
480       LogError "frmEndDefaults", "FillDetails", intEL, strES, sql


End Sub

Private Sub FilllParameter()

          Dim tb As New Recordset
          Dim sql As String
          Dim Analyser As String
          Dim SampleType As String

10        On Error GoTo FilllParameter_Error

20        If optAnalyser(0) Then
30            Select Case UCase(optAnalyser(0).Caption)
              Case "OLYMPUS": Analyser = "A"
40            Case "INTEGRA": Analyser = "I"
50            Case Else
60                Analyser = Left(optAnalyser(0).Caption, 1)
70            End Select
80        ElseIf optAnalyser(1) Then
90            Analyser = Left(optAnalyser(1).Caption, 1)
100       ElseIf optAnalyser(2) Then
          
110           Analyser = optAnalyser(2).Caption
120       ElseIf optAnalyser(3) Then
130           Analyser = optAnalyser(3).Caption
140       End If

150       SampleType = ListCodeFor("ST", cmbSample)

160       lstParameter.Clear

170       sql = "SELECT distinct LongName from endTestDefinitions WHERE " & _
                "SampleType = '" & SampleType & "' and analyser = '" & Analyser & "'"

180       Set tb = New Recordset
190       RecOpenServer 0, tb, sql
200       Do While Not tb.EOF
210           lstParameter.AddItem tb!LongName
220           tb.MoveNext
230       Loop

240       lstParameter.ListIndex = -1

250       Exit Sub

FilllParameter_Error:

          Dim strES As String
          Dim intEL As Integer

260       intEL = Erl
270       strES = Err.Description
280       LogError "frmEndDefaults", "FilllParameter", intEL, strES, sql

End Sub

Private Sub Form_Activate()

10        On Error GoTo Form_Activate_Error

20        If Activated Then Exit Sub

30        Activated = True

40        FillSample
50        FilllParameter
60        FillCats

70        Exit Sub

Form_Activate_Error:

          Dim strES As String
          Dim intEL As Integer

80        intEL = Erl
90        strES = Err.Description
100       LogError "frmEndDefaults", "Form_Activate", intEL, strES

End Sub

Private Sub Form_Load()

10        On Error GoTo Form_Load_Error

20        Activated = False

30        optAnalyser(0).Caption = GetOptionSetting("EndAn1", "")
40        optAnalyser(1).Caption = GetOptionSetting("EndAn2", "")
50        optAnalyser(2).Caption = GetOptionSetting("EndAn3", "")

60        g.Font.Bold = True

70        Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

80        intEL = Erl
90        strES = Err.Description
100       LogError "frmEndDefaults", "Form_Load", intEL, strES

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
60        LogError "frmEndDefaults", "Form_Unload", intEL, strES


End Sub

Private Sub g_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

          Dim n As Long
          Dim ySave As Long


10        On Error GoTo g_MouseUp_Error

20        If g.MouseRow = 0 Then Exit Sub

30        ySave = g.Row

40        g.Col = 0
50        For n = 1 To g.Rows - 1
60            g.Row = n
70            g.CellBackColor = 0
80            g.CellForeColor = 0
90        Next

100       g.Row = ySave
110       g.CellBackColor = vbYellow
120       g.CellForeColor = vbBlue

130       FillDetails




140       Exit Sub

g_MouseUp_Error:

          Dim strES As String
          Dim intEL As Integer



150       intEL = Erl
160       strES = Err.Description
170       LogError "frmEndDefaults", "g_MouseUp", intEL, strES


End Sub

Private Sub lblBarCode_Click()

10        On Error GoTo lblBarCode_Click_Error

20        lblBarCode = iBOX("Scan Bar Code", , lblBarCode)

30        SaveCommonDetails

40        Exit Sub

lblBarCode_Click_Error:

          Dim strES As String
          Dim intEL As Integer



50        intEL = Erl
60        strES = Err.Description
70        LogError "frmEndDefaults", "lblBarCode_Click", intEL, strES


End Sub

Private Sub lstParameter_Click()


10        On Error GoTo lstParameter_Click_Error

20        SSTab1.Enabled = True

30        FillAges

40        FillAllDetails

50        bAmendAgeRange.Enabled = True

60        Exit Sub

lstParameter_Click_Error:

          Dim strES As String
          Dim intEL As Integer



70        intEL = Erl
80        strES = Err.Description
90        LogError "frmEndDefaults", "lstParameter_Click", intEL, strES


End Sub

Private Sub odelta_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        On Error GoTo odelta_MouseUp_Error

20        SaveCommonDetails

30        Exit Sub

odelta_MouseUp_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmEndDefaults", "odelta_MouseUp", intEL, strES


End Sub

Private Sub optAnalyser_Click(Index As Integer)

10        On Error GoTo optAnalyser_Click_Error

20        FilllParameter
30        With g
40            .Rows = 2
50            .AddItem ""
60            .RemoveItem 1
70        End With

80        If Index = 1 Then
90            fraImmulite.Visible = True
100       Else
110           fraImmulite.Visible = False
120       End If

130       Exit Sub

optAnalyser_Click_Error:

          Dim strES As String
          Dim intEL As Integer

140       intEL = Erl
150       strES = Err.Description
160       LogError "frmEndDefaults", "optAnalyser_Click", intEL, strES

End Sub

Private Sub orm_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        On Error GoTo orm_MouseUp_Error

20        SaveCommonDetails

30        Exit Sub

orm_MouseUp_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmEndDefaults", "orm_MouseUp", intEL, strES


End Sub



Private Sub SaveCommonDetails()

          Dim sql As String
          Dim SampleType As String


10        On Error GoTo SaveCommonDetails_Error

20        If lstParameter.ListIndex = -1 Then Exit Sub

30        SampleType = ListCodeFor("ST", cmbSample)

40        sql = "UPDATE endTestDefinitions SET " & _
                "H = '" & IIf(cH = 1, 1, 0) & "', " & _
                "S = '" & IIf(cS = 1, 1, 0) & "', " & _
                "L = '" & IIf(cL = 1, 1, 0) & "', " & _
                "O = '" & IIf(cO = 1, 1, 0) & "', " & _
                "G = '" & IIf(cG = 1, 1, 0) & "', " & _
                "J = '" & IIf(cJ = 1, 1, 0) & "', " & _
                "PrintPriority = '" & Val(tPriority) & "', " & _
                "DP = '" & Val(tDP) & "', " & _
                "DoDelta = '" & oDelta & "', " & _
                "DeltaLimit ='" & Val(tdelta) & "', " & _
                "Units = '" & tUnits & "', " & _
                "BarCode = '" & lblBarCode & "', " & _
                "ShortName = '" & lblShortName & "', " & _
                "knowntoanalyser = '" & IIf(cKnown = 1, 1, 0) & "', " & _
                "CheckTime = '" & IIf(txtCheckTime = "", 1, Val(txtCheckTime)) & "', " & _
                "ShowLessThan = " & chkLowLess.Value & ", " & _
                "ShowMoreThan = " & chkHighGreater.Value & " " & _
                "WHERE LongName = '" & lstParameter & "' " & _
                "and SampleType = '" & SampleType & "' " & _
                "and Hospital = '" & HospName(0) & "' and category = '" & cCat & "'"

50        Cnxn(0).Execute sql




60        Exit Sub

SaveCommonDetails_Error:

          Dim strES As String
          Dim intEL As Integer



70        intEL = Erl
80        strES = Err.Description
90        LogError "frmEndDefaults", "SaveCommonDetails", intEL, strES, sql


End Sub

Private Sub SaveNormals()

          Dim sql As String
          Dim tb As New Recordset
          Dim AgeNumber As Long
          Dim SampleType As String


10        On Error GoTo SaveNormals_Error

20        AgeNumber = -1
30        g.Col = 0
40        For AgeNumber = 1 To g.Rows - 1
50            g.Row = AgeNumber
60            If g.CellBackColor = vbYellow Then
70                AgeNumber = AgeNumber - 1
80                Exit For
90            End If
100       Next
110       If AgeNumber = -1 Then
120           iMsg "SELECT Age Range", vbCritical
130           Exit Sub
140       End If

150       SampleType = ListCodeFor("ST", cmbSample)

160       sql = "SELECT * from endTestDefinitions WHERE " & _
                "LongName = '" & lstParameter & "' " & _
                "and AgeFromDays = '" & FromDays(AgeNumber) & "' " & _
                "and AgeToDays = '" & ToDays(AgeNumber) & "' " & _
                "and SampleType = '" & SampleType & "' " & _
                "and Hospital = '" & HospName(0) & "' " & _
                "and category = '" & cCat & "'"

170       Set tb = New Recordset
180       RecOpenServer 0, tb, sql
190       If tb.EOF Then
200           tb.AddNew
210           tb!LongName = lstParameter
220           tb!AgeFromDays = FromDays(AgeNumber)
230           tb!AgeToDays = ToDays(AgeNumber)
240           tb!SampleType = SampleType
250           tb!Hospital = HospName(0)
260           tb!Category = cCat
270           tb!h = IIf(cH = 1, 1, 0)
280           tb!s = IIf(cS = 1, 1, 0)
290           tb!l = IIf(cL = 1, 1, 0)
300           tb!o = IIf(cO = 1, 1, 0)
310           tb!g = IIf(cG = 1, 1, 0)
320           tb!J = IIf(cJ = 1, 1, 0)
330           tb!ShortName = lblShortName
340           tb!Code = lblHost
350           tb!InUse = 1
360           tb!immunocode = tImmuliteCode
370       End If

380       tb!MaleLow = Val(tMaleLow)
390       tb!MaleHigh = Val(tMaleHigh)
400       tb!FemaleLow = Val(tFemaleLow)
410       tb!FemaleHigh = Val(tFemaleHigh)

420       tb.Update

430       SaveCommonDetails




440       Exit Sub

SaveNormals_Error:

          Dim strES As String
          Dim intEL As Integer



450       intEL = Erl
460       strES = Err.Description
470       LogError "frmEndDefaults", "SaveNormals", intEL, strES, sql


End Sub

Private Sub tDelta_KeyUp(KeyCode As Integer, Shift As Integer)

10        On Error GoTo tDelta_KeyUp_Error

20        SaveCommonDetails

30        Exit Sub

tDelta_KeyUp_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmEndDefaults", "tDelta_KeyUp", intEL, strES


End Sub

Private Sub tDP_KeyUp(KeyCode As Integer, Shift As Integer)

10        On Error GoTo tDP_KeyUp_Error

20        SaveCommonDetails

30        Exit Sub

tDP_KeyUp_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmEndDefaults", "tDP_KeyUp", intEL, strES


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
60        LogError "frmEndDefaults", "tFemaleHigh_KeyUp", intEL, strES


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
60        LogError "frmEndDefaults", "tFemaleLow_KeyUp", intEL, strES


End Sub

Private Sub tfr_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

          Dim sql As String
          Dim AgeNumber As Long
          Dim SampleType As String
          Dim n As Long


10        On Error GoTo tfr_KeyUp_Error

20        If Index = 2 Or Index = 3 Then
30            tfr(Index - 2) = tfr(Index)
40        End If

50        AgeNumber = -1
60        g.Col = 0
70        For n = 1 To g.Rows - 1
80            g.Row = n
90            If g.CellBackColor = vbYellow Then
100               AgeNumber = g.TextMatrix(n, 0)
110               Exit For
120           End If
130       Next
140       If AgeNumber = -1 Then
150           iMsg "SELECT Age Range", vbCritical
160           Exit Sub
170       End If

180       SampleType = ListCodeFor("ST", cmbSample)

190       sql = "UPDATE endTestDefinitions " & _
                "Set FlagMaleLow = '" & tfr(1) & "', " & _
                "FlagMaleHigh = '" & tfr(0) & "', " & _
                "FlagFemaleLow = '" & tfr(3) & "', " & _
                "FlagFemaleHigh = '" & tfr(2) & "' WHERE " & _
                "LongName = '" & lstParameter & "' " & _
                "and AgeFromDays = '" & FromDays(AgeNumber) & "' " & _
                "and AgeToDays = '" & ToDays(AgeNumber) & "' " & _
                "and SampleType = '" & SampleType & "' " & _
                "and Hospital = '" & HospName(0) & "' and category = '" & cCat & "'"

200       Cnxn(0).Execute sql



210       Exit Sub

tfr_KeyUp_Error:

          Dim strES As String
          Dim intEL As Integer



220       intEL = Erl
230       strES = Err.Description
240       LogError "frmEndDefaults", "tfr_KeyUp", intEL, strES, sql


End Sub

Private Sub tImmuliteCode_KeyUp(KeyCode As Integer, Shift As Integer)

          Dim sql As String


10        On Error GoTo tImmuliteCode_KeyUp_Error

20        sql = "UPDATE EndTestDefinitions " & _
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
70        LogError "frmEndDefaults", "tImmuliteCode_KeyUp", intEL, strES, sql


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
60        LogError "frmEndDefaults", "tMaleHigh_KeyUp", intEL, strES


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
60        LogError "frmEndDefaults", "tMaleLow_KeyUp", intEL, strES


End Sub

Private Sub tPlausibleHigh_KeyUp(KeyCode As Integer, Shift As Integer)

          Dim sql As String


10        On Error GoTo tPlausibleHigh_KeyUp_Error

20        sql = "UPDATE EndTestDefinitions " & _
                "Set PlausibleHigh = " & Val(tPlausibleHigh) & " " & _
                "WHERE LongName = '" & lstParameter & "' " & _
                "and Hospital = '" & HospName(0) & "' and category = '" & cCat & "'"

30        Cnxn(0).Execute sql




40        Exit Sub

tPlausibleHigh_KeyUp_Error:

          Dim strES As String
          Dim intEL As Integer



50        intEL = Erl
60        strES = Err.Description
70        LogError "frmEndDefaults", "tPlausibleHigh_KeyUp", intEL, strES, sql


End Sub

Private Sub tPlausibleLow_KeyUp(KeyCode As Integer, Shift As Integer)

          Dim sql As String


10        On Error GoTo tPlausibleLow_KeyUp_Error

20        sql = "UPDATE EndTestDefinitions " & _
                "Set PlausibleLow = " & Val(tPlausibleLow) & " " & _
                "WHERE LongName = '" & lstParameter & "' " & _
                "and Hospital = '" & HospName(0) & "' and category = '" & cCat & "'"

30        Cnxn(0).Execute sql




40        Exit Sub

tPlausibleLow_KeyUp_Error:

          Dim strES As String
          Dim intEL As Integer



50        intEL = Erl
60        strES = Err.Description
70        LogError "frmEndDefaults", "tPlausibleLow_KeyUp", intEL, strES, sql


End Sub

Private Sub tPriority_KeyUp(KeyCode As Integer, Shift As Integer)

10        On Error GoTo tPriority_KeyUp_Error

20        SaveCommonDetails

30        Exit Sub

tPriority_KeyUp_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmEndDefaults", "tPriority_KeyUp", intEL, strES


End Sub

Private Sub tUnits_KeyUp(KeyCode As Integer, Shift As Integer)

10        On Error GoTo tUnits_KeyUp_Error

20        SaveCommonDetails

30        Exit Sub

tUnits_KeyUp_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmEndDefaults", "tUnits_KeyUp", intEL, strES


End Sub



Private Sub txtCheckTime_KeyPress(KeyAscii As Integer)
10        KeyAscii = VI(KeyAscii, Numeric_Only)
End Sub

Private Sub txtCheckTime_KeyUp(KeyCode As Integer, Shift As Integer)
10        SaveCommonDetails
End Sub

Private Sub UpDown1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        On Error GoTo UpDown1_MouseUp_Error

20        SaveCommonDetails

30        Exit Sub

UpDown1_MouseUp_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmEndDefaults", "UpDown1_MouseUp", intEL, strES


End Sub
