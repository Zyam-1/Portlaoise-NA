VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmImmDefaults 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Immunology Parameter Definitions"
   ClientHeight    =   8385
   ClientLeft      =   225
   ClientTop       =   765
   ClientWidth     =   9795
   Icon            =   "frmImmDefaults.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8385
   ScaleWidth      =   9795
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Age Ranges"
      Height          =   2250
      Left            =   2925
      TabIndex        =   53
      Top             =   2970
      Width           =   6795
      Begin VB.CheckBox chkHaem 
         Alignment       =   1  'Right Justify
         Caption         =   "Immunohaem"
         Height          =   195
         Left            =   4320
         TabIndex        =   71
         Top             =   675
         Width           =   1635
      End
      Begin VB.CheckBox chkVward 
         Alignment       =   1  'Right Justify
         Caption         =   "View on Ward"
         Height          =   195
         Left            =   4320
         TabIndex        =   69
         Top             =   1620
         Width           =   1635
      End
      Begin VB.Frame fraImmulite 
         Caption         =   "Immulite Code"
         Height          =   615
         Left            =   5445
         TabIndex        =   64
         Top             =   0
         Visible         =   0   'False
         Width           =   1365
         Begin VB.TextBox tImmuliteCode 
            Height          =   285
            Left            =   210
            TabIndex        =   65
            Top             =   240
            Width           =   945
         End
      End
      Begin VB.CheckBox cRR 
         Alignment       =   1  'Right Justify
         Caption         =   "Print Ref Range"
         Height          =   195
         Left            =   4320
         TabIndex        =   63
         Top             =   1395
         Width           =   1635
      End
      Begin VB.CheckBox cInUse 
         Alignment       =   1  'Right Justify
         Caption         =   "InUse"
         Height          =   195
         Left            =   4320
         TabIndex        =   61
         Top             =   1170
         Width           =   1635
      End
      Begin VB.CheckBox cEOD 
         Alignment       =   1  'Right Justify
         Caption         =   "End Of Day"
         Height          =   195
         Left            =   4320
         TabIndex        =   60
         Top             =   945
         Width           =   1635
      End
      Begin VB.CommandButton bAmendAgeRange 
         Caption         =   "Amend Age Range"
         Enabled         =   0   'False
         Height          =   480
         Left            =   4230
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   180
         Width           =   1155
      End
      Begin MSFlexGridLib.MSFlexGrid g 
         Height          =   1860
         Left            =   120
         TabIndex        =   55
         Top             =   270
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   3281
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
         Left            =   5085
         TabIndex        =   59
         Top             =   1845
         Width           =   765
      End
      Begin VB.Label Label21 
         Caption         =   "Host Code"
         Height          =   315
         Left            =   4305
         TabIndex        =   58
         Top             =   1845
         Width           =   795
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Analyser"
      Height          =   1470
      Left            =   8100
      TabIndex        =   44
      Top             =   120
      Width           =   1545
      Begin VB.OptionButton optAnalyser 
         Caption         =   "ImmunoCAP"
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   73
         Top             =   1200
         Width           =   1215
      End
      Begin VB.OptionButton optAnalyser 
         Caption         =   "None"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   72
         Top             =   945
         Width           =   1065
      End
      Begin VB.OptionButton optAnalyser 
         Caption         =   "Prospec"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   47
         Top             =   720
         Width           =   1065
      End
      Begin VB.OptionButton optAnalyser 
         Caption         =   "Immulite"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   46
         Top             =   480
         Width           =   885
      End
      Begin VB.OptionButton optAnalyser 
         Caption         =   "Roche"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   45
         Top             =   240
         Value           =   -1  'True
         Width           =   915
      End
   End
   Begin VB.Frame Frame5 
      Height          =   1605
      Left            =   2925
      TabIndex        =   34
      Top             =   1350
      Width           =   5055
      Begin VB.TextBox txtCheckTime 
         Height          =   285
         Left            =   3960
         MaxLength       =   5
         TabIndex        =   74
         Top             =   1200
         Width           =   495
      End
      Begin VB.CheckBox chkPrintable 
         Alignment       =   1  'Right Justify
         Caption         =   "Printable"
         Height          =   195
         Left            =   585
         TabIndex        =   70
         Top             =   990
         Width           =   1635
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
         Top             =   765
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
         Left            =   1980
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "1"
         Top             =   450
         Width           =   360
      End
      Begin VB.TextBox tPriority 
         Height          =   285
         Left            =   1980
         TabIndex        =   13
         Top             =   135
         Width           =   555
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   285
         Left            =   2370
         TabIndex        =   14
         Top             =   450
         Width           =   195
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   327681
         Value           =   1
         BuddyControl    =   "tDP"
         BuddyDispid     =   196630
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
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "day(s)"
         Height          =   195
         Left            =   4500
         TabIndex        =   76
         Top             =   1230
         Width           =   420
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Check Time"
         Height          =   195
         Left            =   3060
         TabIndex        =   75
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
         Top             =   495
         Width           =   1875
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Print Priority"
         Height          =   195
         Left            =   1110
         TabIndex        =   17
         Top             =   180
         Width           =   825
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1170
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
      Begin VB.Label lblSample 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1035
         TabIndex        =   67
         Top             =   810
         Width           =   1275
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Sampletype"
         Height          =   195
         Left            =   90
         TabIndex        =   66
         Top             =   855
         Width           =   825
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
      Height          =   1185
      Left            =   2880
      TabIndex        =   18
      Top             =   120
      Width           =   2235
      Begin VB.ComboBox cmbSample 
         Height          =   315
         Left            =   270
         Style           =   2  'Dropdown List
         TabIndex        =   68
         Top             =   315
         Width           =   1680
      End
      Begin VB.ComboBox cCat 
         Height          =   315
         Left            =   270
         TabIndex        =   57
         Top             =   765
         Width           =   1665
      End
   End
   Begin VB.CommandButton bexit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   735
      Left            =   8310
      Picture         =   "frmImmDefaults.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1785
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
      Height          =   8070
      IntegralHeight  =   0   'False
      Left            =   135
      TabIndex        =   1
      Top             =   150
      Width           =   2415
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2985
      Left            =   2925
      TabIndex        =   0
      Top             =   5235
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
      TabPicture(0)   =   "frmImmDefaults.frx":0614
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
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Flag Range"
      TabPicture(1)   =   "frmImmDefaults.frx":0630
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label15"
      Tab(1).Control(1)=   "Label14"
      Tab(1).Control(2)=   "Label13"
      Tab(1).Control(3)=   "Label12"
      Tab(1).Control(4)=   "Label7"
      Tab(1).Control(5)=   "tfr(3)"
      Tab(1).Control(6)=   "tfr(1)"
      Tab(1).Control(7)=   "tfr(2)"
      Tab(1).Control(8)=   "tfr(0)"
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "Masks"
      TabPicture(2)   =   "frmImmDefaults.frx":064C
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
      TabPicture(3)   =   "frmImmDefaults.frx":0668
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label8"
      Tab(3).Control(1)=   "Label9"
      Tab(3).Control(2)=   "Label10"
      Tab(3).Control(3)=   "tPlausibleHigh"
      Tab(3).Control(4)=   "tPlausibleLow"
      Tab(3).ControlCount=   5
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
         Left            =   840
         TabIndex        =   32
         Top             =   2520
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
Attribute VB_Name = "frmImmDefaults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Activated As Boolean

Private FromDays() As Long
Private ToDays() As Long

Private Analyser As String

Private Sub FillAges()

          Dim tb As New Recordset
          Dim s As String
          Dim SampleType As String
          Dim n As Long
          Dim sql As String
          Dim Cat As String

10        On Error GoTo FillAges_Error

20        If cCat = "" Then Cat = "Default" Else Cat = cCat

30        GetAnalyser

40        With g
50            .Rows = 2
60            .AddItem ""
70            .RemoveItem 1
80        End With

90        SampleType = ListCodeFor("ST", cmbSample)

100       ReDim FromDays(0 To 0)
110       ReDim ToDays(0 To 0)

120       sql = "SELECT * from ImmTestDefinitions WHERE " & _
                "LongName = '" & lstParameter & "' " & _
                "and SampleType = '" & SampleType & "' and category = '" & Cat & "' and analyser = '" & Analyser & "' " & _
                "Order by AgeFromDays"
130       Set tb = New Recordset
140       RecOpenClient 0, tb, sql

150       If tb.EOF Then Exit Sub
160       ReDim FromDays(0 To tb.RecordCount - 1)
170       ReDim ToDays(0 To tb.RecordCount - 1)
180       n = 0
190       Do While Not tb.EOF
200           FromDays(n) = tb!AgeFromDays
210           ToDays(n) = tb!AgeToDays
220           s = Format$(n) & vbTab & _
                  dmyFromCount(FromDays(n)) & vbTab & _
                  dmyFromCount(ToDays(n))
230           g.AddItem s
240           n = n + 1
250           tb.MoveNext
260       Loop

270       If g.Rows > 2 Then
280           g.RemoveItem 1
290       End If

300       g.Col = 0
310       g.Row = 1
320       g.CellBackColor = vbYellow
330       g.CellForeColor = vbBlue

340       Exit Sub

FillAges_Error:

          Dim strES As String
          Dim intEL As Integer

350       intEL = Erl
360       strES = Err.Description
370       LogError "frmImmDefaults", "FillAges", intEL, strES, sql

End Sub

Private Sub FillDetails()

          Dim tb As New Recordset
          Dim sql As String
          Dim SampleType As String
          Dim AgeNumber As Long

10        On Error GoTo FillDetails_Error

20        SampleType = ListCodeFor("ST", cmbSample)

30        GetAnalyser

40        AgeNumber = -1
50        g.Col = 0
60        For AgeNumber = 1 To g.Rows - 1
70            g.Row = AgeNumber
80            If g.CellBackColor = vbYellow Then
90                AgeNumber = AgeNumber - 1
100               Exit For
110           End If
120       Next
130       If AgeNumber = -1 Then
140           iMsg "SELECT Age Range", vbCritical
150           Exit Sub
160       End If

170       sql = "SELECT * from ImmTestDefinitions WHERE " & _
                "LongName = '" & lstParameter & "' and Analyser = '" & Analyser & "' " & _
                "and AgeFromDays = '" & FromDays(AgeNumber) & "' " & _
                "and AgeToDays = '" & ToDays(AgeNumber) & "' " & _
                "and SampleType = '" & SampleType & "' and category = '" & cCat & "' and hospital = '" & HospName(0) & "'"
180       Set tb = New Recordset
190       RecOpenClient 0, tb, sql
200       With tb

210           If Not .EOF Then
220               tPriority = !PrintPriority & ""
230               If !KnownToAnalyser & "" <> "" Then cKnown = IIf(!KnownToAnalyser, 1, 0)
240               oDelta = IIf(!DoDelta, 1, 0)
250               tDP = !DP
260               tdelta = !DeltaLimit
270               tUnits = !Units
280               lblBarCode = !BarCode & ""
290               cCat = !Category
300               tMaleHigh = !MaleHigh
310               tMaleLow = !MaleLow
320               tFemaleHigh = !FemaleHigh
330               tFemaleLow = !FemaleLow
340               tfr(0) = !FlagMaleHigh
350               tfr(1) = !FlagMaleLow
360               tfr(2) = !FlagFemaleHigh
370               tfr(3) = !FlagFemaleLow
380               cH = IIf(!h, 1, 0)
390               cS = IIf(!s, 1, 0)
400               cL = IIf(!l, 1, 0)
410               cO = IIf(!o, 1, 0)
420               cG = IIf(!g, 1, 0)
430               cJ = IIf(!J, 1, 0)
440           End If
450       End With

460       Exit Sub

FillDetails_Error:

          Dim strES As String
          Dim intEL As Integer

470       intEL = Erl
480       strES = Err.Description
490       LogError "frmImmDefaults", "FillDetails", intEL, strES, sql

End Sub

Private Sub FillAllDetails()

          Dim tb As New Recordset
          Dim sql As String
          Dim SampleType As String

10        On Error GoTo FillAllDetails_Error

20        GetAnalyser

30        tImmuliteCode = ""
40        tPriority = ""
50        cKnown = 0
60        oDelta = 0
70        tDP = 0
80        tdelta = 0
90        tUnits = ""
100       lblBarCode = ""
110       lblShortName = ""
120       tPlausibleHigh = ""
130       tPlausibleLow = ""

140       tMaleHigh = ""
150       tMaleLow = ""
160       tFemaleHigh = ""
170       tFemaleLow = ""
180       tfr(0) = ""
190       tfr(1) = ""
200       tfr(2) = ""
210       tfr(3) = ""
220       cH = 0
230       cS = 0
240       cL = 0
250       cO = 0
260       cG = 0
270       cJ = 0
280       lblHost = ""
290       cInUse = 0
300       cEOD = 0
310       lblSample = ""
320       chkHaem = 0
330       txtCheckTime = ""

340       If cCat = "" Then cCat = "Default"

350       SampleType = ListCodeFor("ST", cmbSample)

360       sql = "SELECT distinct * from ImmTestDefinitions WHERE " & _
                "LongName = '" & lstParameter & "' " & _
                "and AgeFromDays = '" & FromDays(0) & "' " & _
                "and AgeToDays = '" & ToDays(0) & "' " & _
                "and SampleType = '" & SampleType & "' and category = 'Default' and Analyser = '" & Analyser & "'"
370       Set tb = New Recordset
380       RecOpenClient 0, tb, sql
390       With tb

400           If Not .EOF Then
410               lblShortName = tb!ShortName & ""
420               tImmuliteCode = Trim$(!immunocode & "")
430               tPriority = !PrintPriority & ""
440               If Trim(!KnownToAnalyser) & "" <> "" Then cKnown = IIf(!KnownToAnalyser, 1, 0)
450               oDelta = IIf(!DoDelta, 1, 0)
460               tDP = !DP & ""
470               tdelta = !DeltaLimit & ""
480               tUnits = !Units & ""
490               lblBarCode = !BarCode & ""
500               tPlausibleHigh = !PlausibleHigh & ""
510               tPlausibleLow = !PlausibleLow & ""
520               lblSample = ListText("ST", !SampleType)
530               tMaleHigh = !MaleHigh & ""
540               tMaleLow = !MaleLow & ""
550               tFemaleHigh = !FemaleHigh & ""
560               tFemaleLow = !FemaleLow & ""
570               tfr(0) = !FlagMaleHigh & ""
580               tfr(1) = !FlagMaleLow & ""
590               tfr(2) = !FlagFemaleHigh & ""
600               tfr(3) = !FlagFemaleLow & ""
610               cH = IIf(!h, 1, 0)
620               cS = IIf(!s, 1, 0)
630               cL = IIf(!l, 1, 0)
640               cO = IIf(!o, 1, 0)
650               cG = IIf(!g, 1, 0)
660               cJ = IIf(!J, 1, 0)
670               If Trim(!PrnRR) & "" <> "" Then
680                   cRR = IIf(!PrnRR, 1, 0)
690               Else
700                   cRR = 1
710               End If
720               If Trim(!vward) & "" <> "" Then
730                   chkVward = IIf(!vward, 1, 0)
740               Else
750                   chkVward = 1
760               End If
770               If Trim(!haem) & "" <> "" Then
780                   chkHaem = IIf(!haem, 1, 0)
790               Else
800                   chkHaem = 0
810               End If
820               If Trim(!InUse) & "" <> "" Then cInUse = IIf(!InUse, 1, 0)
830               If Trim(!Eod) & "" = "" Then cEOD = 0 Else cEOD = IIf(!Eod, 1, 0)
840               lblHost = tb!Code
850               chkPrintable = IIf(!Printable, 1, 0)
860               If IsNull(!CheckTime) Then
870                   txtCheckTime = ""
880               Else
890                   txtCheckTime = !CheckTime
900               End If
910           End If
920       End With

930       If cCat <> "Default" Then
940           sql = "SELECT distinct * from ImmTestDefinitions WHERE " & _
                    "LongName = '" & lstParameter & "' " & _
                    "and AgeFromDays = '" & FromDays(0) & "' " & _
                    "and AgeToDays = '" & ToDays(0) & "' " & _
                    "and SampleType = '" & SampleType & "' " & _
                    "and category = '" & cCat & "' and analyser = '" & Analyser & "'"
950           Set tb = New Recordset
960           RecOpenClient 0, tb, sql
970           With tb
980               If Not .EOF Then
990                   tPlausibleHigh = !PlausibleHigh & ""
1000                  tPlausibleLow = !PlausibleLow & ""
1010                  tMaleHigh = !MaleHigh & ""
1020                  tMaleLow = !MaleLow & ""
1030                  tFemaleHigh = !FemaleHigh & ""
1040                  tFemaleLow = !FemaleLow & ""
1050                  tfr(0) = !FlagMaleHigh & ""
1060                  tfr(1) = !FlagMaleLow & ""
1070                  tfr(2) = !FlagFemaleHigh & ""
1080                  tfr(3) = !FlagFemaleLow & ""
1090              End If
1100          End With
1110      End If

1120      Exit Sub

FillAllDetails_Error:

          Dim strES As String
          Dim intEL As Integer

1130      intEL = Erl
1140      strES = Err.Description
1150      LogError "frmImmDefaults", "FillAllDetails", intEL, strES, sql

End Sub


Private Sub FilllParameter()

          Dim tb As New Recordset
          Dim sql As String
          Dim SampleType As String

10        On Error GoTo FilllParameter_Error

20        GetAnalyser

30        SampleType = ListCodeFor("ST", cmbSample)

40        lstParameter.Clear

50        sql = "SELECT distinct LongName from ImmTestDefinitions WHERE "
60        sql = sql & "SampleType = '" & SampleType & "'  "

70        If Analyser <> "N" Then sql = sql & " and analyser = '" & Analyser & "' "

80        sql = sql & "and hospital = '" & HospName(0) & "'"

90        Set tb = New Recordset
100       RecOpenClient 0, tb, sql
110       Do While Not tb.EOF
120           lstParameter.AddItem tb!LongName
130           tb.MoveNext
140       Loop

150       lstParameter.ListIndex = -1

160       Exit Sub

FilllParameter_Error:

          Dim strES As String
          Dim intEL As Integer

170       intEL = Erl
180       strES = Err.Description
190       LogError "frmImmDefaults", "FilllParameter", intEL, strES, sql

End Sub

Private Sub GetAnalyser()

10        If optAnalyser(0) Then
20            Select Case optAnalyser(0).Caption
                Case "Olympus": Analyser = "A"
30              Case "Integra": Analyser = "I"
40              Case "Immage": Analyser = "1"
50              Case "Best": Analyser = "2"
60            End Select
70        ElseIf optAnalyser(1) Then
80            Select Case optAnalyser(1).Caption
                Case "Olympus": Analyser = "A"
90              Case "Integra": Analyser = "I"
100             Case "Immage": Analyser = "1"
110             Case "Best": Analyser = "2"
120             Case Else: Analyser = "4"
130           End Select
140       ElseIf optAnalyser(2) Then
150           Select Case optAnalyser(2).Caption
                Case "Olympus": Analyser = "A"
160             Case "Integra": Analyser = "I"
170             Case "Immage": Analyser = "1"
180             Case "Prospec": Analyser = "3"
190             Case Else: Analyser = "8"
200           End Select
210       ElseIf optAnalyser(4) Then
220           Analyser = "ImmunoCAP"
230       Else
240           Analyser = ""
250       End If


End Sub


Private Sub SaveCommonDetails()

          Dim sql As String
          Dim SampleType As String

10        On Error GoTo SaveCommonDetails_Error

20        GetAnalyser

30        SampleType = ListCodeFor("ST", cmbSample)

40        sql = "UPDATE ImmTestDefinitions SET " & _
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
                "CheckTime = '" & IIf(txtCheckTime = "", 1, Val(txtCheckTime)) & "', " & _
                "knowntoanalyser = '" & IIf(cKnown = 1, 1, 0) & "' " & _
                "WHERE LongName = '" & lstParameter & "' " & _
                "and SampleType = '" & SampleType & "' " & _
                "and Hospital = '" & HospName(0) & "' and category = '" & cCat & "' and analyser = '" & Analyser & "'"

50        Cnxn(0).Execute sql

60        Exit Sub

SaveCommonDetails_Error:

          Dim strES As String
          Dim intEL As Integer

70        intEL = Erl
80        strES = Err.Description
90        LogError "frmImmDefaults", "SaveCommonDetails", intEL, strES, sql

End Sub


Private Sub bAmendAgeRange_Click()

10        On Error GoTo bAmendAgeRange_Click_Error

20        GetAnalyser

30        With frmAges
40            .Analyte = lstParameter
50            .SampleType = ListCodeFor("ST", cmbSample)
60            .Discipline = "Immunology"
70            .Analyser = Analyser
80            If cCat.Text = "Default" Then .Cat = "" Else .Cat = cCat
90            .Show 1
100       End With

110       FillAges

120       FillAllDetails

130       Exit Sub

bAmendAgeRange_Click_Error:

          Dim strES As String
          Dim intEL As Integer

140       intEL = Erl
150       strES = Err.Description
160       LogError "frmImmDefaults", "bAmendAgeRange_Click", intEL, strES

End Sub

Private Sub bexit_Click()

10        Unload Me

End Sub



Private Sub cCat_Click()
      'Dim sql As String
      'dim tb as new recordset
      '
      'sql = "SELECT * from immtestdefinitions " & _
        '      "WHERE LongName = '" & lstParameter & "' " & _
        '      "and Hospital = '" & Hospname(0) & "' and " & _
        '      "category = '" & cCat & "'"
      '
      'If tb.EOF Then
      '  tb.AddNew

10        On Error GoTo cCat_Click_Error

20        FillAllDetails

30        Exit Sub

cCat_Click_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmImmDefaults", "cCat_Click", intEL, strES

End Sub

Private Sub cEOD_Click()

          Dim sql As String

10        On Error GoTo cEOD_Click_Error

20        If lblHost = "" Then Exit Sub

30        GetAnalyser

40        If cEOD.Value = 1 Then
50            sql = "UPDATE immtestdefinitions set eod = 1 WHERE longname = '" & lstParameter & "' and analyser = '" & Analyser & "'"
60            Cnxn(0).Execute sql
70        ElseIf cEOD.Value = 0 Then
80            sql = "UPDATE immtestdefinitions set eod = 0 WHERE longname = '" & lstParameter & "' and analyser = '" & Analyser & "'"
90            Cnxn(0).Execute sql
100       End If

110       Exit Sub

cEOD_Click_Error:

          Dim strES As String
          Dim intEL As Integer

120       intEL = Erl
130       strES = Err.Description
140       LogError "frmImmDefaults", "cEOD_Click", intEL, strES, sql

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
60        LogError "frmImmDefaults", "cG_MouseUp", intEL, strES

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
60        LogError "frmImmDefaults", "cH_MouseUp", intEL, strES

End Sub






Private Sub chkHaem_Click()

          Dim sql As String

10        On Error GoTo chkHaem_Click_Error

20        If lblHost = "" Then Exit Sub

30        GetAnalyser

40        If chkHaem.Value = 1 Then
50            sql = "UPDATE immtestdefinitions set haem = 1 WHERE longname = '" & lstParameter & "' and analyser = '" & Analyser & "'"
60            Cnxn(0).Execute sql
70        ElseIf chkHaem.Value = 0 Then
80            sql = "UPDATE immtestdefinitions set haem = 0 WHERE longname = '" & lstParameter & "' and analyser = '" & Analyser & "'"
90            Cnxn(0).Execute sql
100       End If

110       Exit Sub

chkHaem_Click_Error:

          Dim strES As String
          Dim intEL As Integer

120       intEL = Erl
130       strES = Err.Description
140       LogError "frmImmDefaults", "chkHaem_Click", intEL, strES, sql

End Sub

Private Sub chkPrintable_Click()

          Dim sql As String

10        On Error GoTo chkPrintable_Click_Error

20        If lblHost = "" Then Exit Sub

30        If chkPrintable.Value = 1 Then
40            sql = "UPDATE Immtestdefinitions set Printable = 1 WHERE code = '" & lblHost & "'"
50            Cnxn(0).Execute sql
60        ElseIf chkPrintable.Value = 0 Then
70            sql = "UPDATE Immtestdefinitions set printable = 0 WHERE code = '" & lblHost & "'"
80            Cnxn(0).Execute sql
90        End If

100       Exit Sub

chkPrintable_Click_Error:

          Dim strES As String
          Dim intEL As Integer

110       intEL = Erl
120       strES = Err.Description
130       LogError "frmImmDefaults", "chkPrintable_Click", intEL, strES, sql

End Sub

Private Sub chkVward_Click()

          Dim sql As String

10        On Error GoTo chkVward_Click_Error

20        If lblHost = "" Then Exit Sub


30        If chkVward.Value = 1 Then
40            sql = "UPDATE immtestdefinitions set vward = 1 WHERE code = '" & lblHost & "'"
50            Cnxn(0).Execute sql
60        ElseIf chkVward.Value = 0 Then
70            sql = "UPDATE immtestdefinitions set vward = 0 WHERE code = '" & lblHost & "'"
80            Cnxn(0).Execute sql
90        End If

100       Exit Sub

chkVward_Click_Error:

          Dim strES As String
          Dim intEL As Integer

110       intEL = Erl
120       strES = Err.Description
130       LogError "frmImmDefaults", "chkVward_Click", intEL, strES, sql

End Sub

Private Sub cInUse_Click()

          Dim sql As String

10        On Error GoTo cInUse_Click_Error

20        If lblHost = "" Then Exit Sub


30        If cInUse.Value = 1 Then
40            sql = "UPDATE immtestdefinitions set inuse = 1 WHERE code = '" & lblHost & "'"
50            Cnxn(0).Execute sql
60        ElseIf cInUse.Value = 0 Then
70            sql = "UPDATE immtestdefinitions set inuse = 0 WHERE code = '" & lblHost & "'"
80            Cnxn(0).Execute sql
90        End If

100       Exit Sub

cInUse_Click_Error:

          Dim strES As String
          Dim intEL As Integer

110       intEL = Erl
120       strES = Err.Description
130       LogError "frmImmDefaults", "cInUse_Click", intEL, strES, sql

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
60        LogError "frmImmDefaults", "cJ_MouseUp", intEL, strES

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
60        LogError "frmImmDefaults", "cKnown_MouseUp", intEL, strES

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
60        LogError "frmImmDefaults", "cL_MouseUp", intEL, strES

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
60        LogError "frmImmDefaults", "cmbSample_Click", intEL, strES

End Sub
Private Sub FillSample()

          Dim sql As String
          Dim tb As New Recordset

10        On Error GoTo FillSample_Error

20        sql = "SELECT Text FROM Lists WHERE " & _
                "ListType = 'ST'"
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
130       LogError "frmImmDefaults", "FillSample", intEL, strES, sql

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
60        LogError "frmImmDefaults", "cO_MouseUp", intEL, strES

End Sub


Private Sub cRR_Click()

          Dim sql As String

10        On Error GoTo cRR_Click_Error

20        If lblHost = "" Then Exit Sub

30        GetAnalyser

40        If cRR.Value = 1 Then
50            sql = "UPDATE immtestdefinitions set PrnRR = 1 WHERE longname = '" & lstParameter & "' and analyser = '" & Analyser & "'"
60            Cnxn(0).Execute sql
70        ElseIf cRR.Value = 0 Then
80            sql = "UPDATE immtestdefinitions set PrnRR = 0 WHERE longname = '" & lstParameter & "' and analyser = '" & Analyser & "'"
90            Cnxn(0).Execute sql
100       End If

110       Exit Sub

cRR_Click_Error:

          Dim strES As String
          Dim intEL As Integer

120       intEL = Erl
130       strES = Err.Description
140       LogError "frmImmDefaults", "cRR_Click", intEL, strES, sql

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
60        LogError "frmImmDefaults", "cS_MouseUp", intEL, strES

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
100       LogError "frmImmDefaults", "Form_Activate", intEL, strES

End Sub

Private Sub FillCats()

          Dim n As Long
          Dim Found As Boolean
          Dim tb As New Recordset
          Dim sql As String

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
200       LogError "frmImmDefaults", "FillCats", intEL, strES, sql

End Sub

Private Sub Form_Load()

10        On Error GoTo Form_Load_Error

20        Activated = False

30        optAnalyser(0).Caption = "Immage"
40        optAnalyser(1).Caption = "Best"
50        optAnalyser(2).Caption = "Prospec"

60        g.Font.Bold = True

70        Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

80        intEL = Erl
90        strES = Err.Description
100       LogError "frmImmDefaults", "Form_Load", intEL, strES

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
60        LogError "frmImmDefaults", "Form_Unload", intEL, strES


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
170       LogError "frmImmDefaults", "g_MouseUp", intEL, strES

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
70        LogError "frmImmDefaults", "lblBarCode_Click", intEL, strES

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
90        LogError "frmImmDefaults", "lstParameter_Click", intEL, strES

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
60        LogError "frmImmDefaults", "odelta_MouseUp", intEL, strES

End Sub


Private Sub optAnalyser_Click(Index As Integer)

10        On Error GoTo optAnalyser_Click_Error

20        FilllParameter
30        g.Rows = 2
40        g.AddItem ""
50        g.RemoveItem 1

60        If Index = 1 Then
70            fraImmulite.Visible = True
80        Else
90            fraImmulite.Visible = False
100       End If

110       Exit Sub

optAnalyser_Click_Error:

          Dim strES As String
          Dim intEL As Integer

120       intEL = Erl
130       strES = Err.Description
140       LogError "frmImmDefaults", "optAnalyser_Click", intEL, strES

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
60        LogError "frmImmDefaults", "orm_MouseUp", intEL, strES

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
60        LogError "frmImmDefaults", "tDelta_KeyUp", intEL, strES

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
60        LogError "frmImmDefaults", "tDP_KeyUp", intEL, strES

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
60        LogError "frmImmDefaults", "tFemaleHigh_KeyUp", intEL, strES

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
60        LogError "frmImmDefaults", "tFemaleLow_KeyUp", intEL, strES

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

50        GetAnalyser

60        AgeNumber = -1
70        g.Col = 0
80        For n = 1 To g.Rows - 1
90            g.Row = n
100           If g.CellBackColor = vbYellow Then
110               AgeNumber = g.TextMatrix(n, 0)
120               Exit For
130           End If
140       Next
150       If AgeNumber = -1 Then
160           iMsg "SELECT Age Range", vbCritical
170           Exit Sub
180       End If

190       SampleType = ListCodeFor("ST", cmbSample)

200       sql = "UPDATE ImmTestDefinitions " & _
                "Set FlagMaleLow = '" & tfr(1) & "', " & _
                "FlagMaleHigh = '" & tfr(0) & "', " & _
                "FlagFemaleLow = '" & tfr(3) & "', " & _
                "FlagFemaleHigh = '" & tfr(2) & "' WHERE " & _
                "LongName = '" & lstParameter & "' " & _
                "and AgeFromDays = '" & FromDays(AgeNumber) & "' " & _
                "and AgeToDays = '" & ToDays(AgeNumber) & "' " & _
                "and SampleType = '" & SampleType & "' " & _
                "and Hospital = '" & HospName(0) & "' and category = '" & cCat & "' and analyser = '" & Analyser & "'"

210       Cnxn(0).Execute sql

220       Exit Sub

tfr_KeyUp_Error:

          Dim strES As String
          Dim intEL As Integer

230       intEL = Erl
240       strES = Err.Description
250       LogError "frmImmDefaults", "tfr_KeyUp", intEL, strES, sql

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

150       GetAnalyser

160       SampleType = ListCodeFor("ST", cmbSample)

170       sql = "SELECT * from ImmTestDefinitions WHERE " & _
                "LongName = '" & lstParameter & "' " & _
                "and AgeFromDays = '" & FromDays(AgeNumber) & "' " & _
                "and AgeToDays = '" & ToDays(AgeNumber) & "' " & _
                "and SampleType = '" & SampleType & "' " & _
                "and Hospital = '" & HospName(0) & "' " & _
                "and category = '" & cCat & "' and analyser = '" & Analyser & "'"

180       Set tb = New Recordset
190       RecOpenServer 0, tb, sql
200       If tb.EOF Then
210           tb.AddNew
220           tb!LongName = lstParameter
230           tb!AgeFromDays = FromDays(AgeNumber)
240           tb!AgeToDays = ToDays(AgeNumber)
250           tb!SampleType = SampleType
260           tb!Hospital = HospName(0)
270           tb!Category = cCat
280           tb!h = IIf(cH = 1, 1, 0)
290           tb!s = IIf(cS = 1, 1, 0)
300           tb!l = IIf(cL = 1, 1, 0)
310           tb!o = IIf(cO = 1, 1, 0)
320           tb!g = IIf(cG = 1, 1, 0)
330           tb!J = IIf(cJ = 1, 1, 0)
340           tb!ShortName = lblShortName
350           tb!Code = lblHost
360           tb!InUse = 1
370           tb!immunocode = tImmuliteCode
380       End If

390       tb!MaleLow = Val(tMaleLow)
400       tb!MaleHigh = Val(tMaleHigh)
410       tb!FemaleLow = Val(tFemaleLow)
420       tb!FemaleHigh = Val(tFemaleHigh)

430       tb.Update

440       SaveCommonDetails

450       Exit Sub

SaveNormals_Error:

          Dim strES As String
          Dim intEL As Integer

460       intEL = Erl
470       strES = Err.Description
480       LogError "frmImmDefaults", "SaveNormals", intEL, strES, sql

End Sub

Private Sub tImmuliteCode_KeyUp(KeyCode As Integer, Shift As Integer)

          Dim sql As String

10        On Error GoTo tImmuliteCode_KeyUp_Error

20        sql = "UPDATE ImmTestDefinitions " & _
                "Set ImmunoCode = '" & Trim$(tImmuliteCode) & "' " & _
                "WHERE code = '" & lblHost & "' " & _
                "and Hospital = '" & HospName(0) & "'"

30        Cnxn(0).Execute sql

40        Exit Sub

tImmuliteCode_KeyUp_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmImmDefaults", "tImmuliteCode_KeyUp", intEL, strES, sql

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
60        LogError "frmImmDefaults", "tMaleHigh_KeyUp", intEL, strES

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
60        LogError "frmImmDefaults", "tMaleLow_KeyUp", intEL, strES

End Sub


Private Sub tPlausibleHigh_KeyUp(KeyCode As Integer, Shift As Integer)

          Dim sql As String

10        On Error GoTo tPlausibleHigh_KeyUp_Error

20        sql = "UPDATE ImmTestDefinitions " & _
                "Set PlausibleHigh = " & Val(tPlausibleHigh) & " " & _
                "WHERE code = '" & lblHost & "' " & _
                "and Hospital = '" & HospName(0) & "' and category = '" & cCat & "'"

30        Cnxn(0).Execute sql

40        Exit Sub

tPlausibleHigh_KeyUp_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmImmDefaults", "tPlausibleHigh_KeyUp", intEL, strES, sql

End Sub


Private Sub tPlausibleLow_KeyUp(KeyCode As Integer, Shift As Integer)

          Dim sql As String

10        On Error GoTo tPlausibleLow_KeyUp_Error

20        sql = "UPDATE ImmTestDefinitions " & _
                "Set PlausibleLow = " & Val(tPlausibleLow) & " " & _
                "WHERE code = '" & lblHost & "' " & _
                "and Hospital = '" & HospName(0) & "' and category = '" & cCat & "'"

30        Cnxn(0).Execute sql

40        Exit Sub

tPlausibleLow_KeyUp_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmImmDefaults", "tPlausibleLow_KeyUp", intEL, strES, sql

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
60        LogError "frmImmDefaults", "tPriority_KeyUp", intEL, strES

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
60        LogError "frmImmDefaults", "tUnits_KeyUp", intEL, strES

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
60        LogError "frmImmDefaults", "UpDown1_MouseUp", intEL, strES

End Sub


