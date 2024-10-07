VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmBgaDefaults 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Blood Gas Parameter Definitions"
   ClientHeight    =   8625
   ClientLeft      =   3180
   ClientTop       =   1170
   ClientWidth     =   9795
   ClipControls    =   0   'False
   Icon            =   "frmBgaDefaults.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8625
   ScaleWidth      =   9795
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Enabled         =   0   'False
      Height          =   675
      Left            =   2520
      Picture         =   "frmBgaDefaults.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   7830
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Caption         =   "Age Ranges"
      Height          =   2115
      Left            =   2850
      TabIndex        =   38
      Top             =   2610
      Width           =   6615
      Begin VB.CheckBox cEOD 
         Alignment       =   1  'Right Justify
         Caption         =   "End Of Day"
         Height          =   195
         Left            =   4770
         TabIndex        =   46
         Top             =   1500
         Width           =   1635
      End
      Begin VB.CheckBox cInUse 
         Alignment       =   1  'Right Justify
         Caption         =   "InUse"
         Height          =   195
         Left            =   4770
         TabIndex        =   44
         Top             =   1710
         Width           =   1635
      End
      Begin VB.CommandButton cmdAmendAgeRange 
         Caption         =   "Amend Age Range"
         Enabled         =   0   'False
         Height          =   525
         Left            =   4890
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   870
         Width           =   1155
      End
      Begin MSFlexGridLib.MSFlexGrid grdAge 
         Height          =   1725
         Left            =   570
         TabIndex        =   40
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
         TabIndex        =   43
         Top             =   360
         Width           =   795
      End
      Begin VB.Label lblHost 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5610
         TabIndex        =   42
         Top             =   330
         Width           =   765
      End
   End
   Begin VB.Frame Frame5 
      Height          =   1545
      Left            =   2850
      TabIndex        =   33
      Top             =   1140
      Width           =   6585
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
         BuddyDispid     =   196622
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
      Left            =   5670
      TabIndex        =   20
      Top             =   180
      Width           =   2775
      Begin VB.ComboBox cmbUnits 
         Height          =   315
         Left            =   810
         Style           =   2  'Dropdown List
         TabIndex        =   47
         Top             =   150
         Width           =   1545
      End
      Begin VB.Label lblBarCode 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   810
         TabIndex        =   41
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
      Height          =   915
      Left            =   2850
      TabIndex        =   19
      Top             =   180
      Width           =   2235
      Begin VB.OptionButton oSU 
         Caption         =   "CSF/Other"
         Height          =   225
         Index           =   2
         Left            =   630
         TabIndex        =   45
         Top             =   540
         Width           =   1095
      End
      Begin VB.OptionButton oSU 
         Caption         =   "Urine"
         Height          =   225
         Index           =   1
         Left            =   1170
         TabIndex        =   18
         Top             =   270
         Width           =   795
      End
      Begin VB.OptionButton oSU 
         Alignment       =   1  'Right Justify
         Caption         =   "Serum"
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   37
         Top             =   240
         Value           =   -1  'True
         Width           =   825
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   630
      Left            =   6075
      Picture         =   "frmBgaDefaults.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7830
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
      TabPicture(0)   =   "frmBgaDefaults.frx":091E
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
      TabPicture(1)   =   "frmBgaDefaults.frx":093A
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
      TabCaption(2)   =   "Plausible"
      TabPicture(2)   =   "frmBgaDefaults.frx":0956
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label10"
      Tab(2).Control(1)=   "Label9"
      Tab(2).Control(2)=   "Label8"
      Tab(2).Control(3)=   "tPlausibleLow"
      Tab(2).Control(4)=   "tPlausibleHigh"
      Tab(2).ControlCount=   5
      Begin VB.TextBox tPlausibleHigh 
         Height          =   285
         Left            =   -72330
         TabIndex        =   50
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox tPlausibleLow 
         Height          =   285
         Left            =   -72330
         TabIndex        =   49
         Top             =   1290
         Width           =   1215
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
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "High"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   -72840
         TabIndex        =   53
         Top             =   750
         Width           =   330
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Low"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   -72810
         TabIndex        =   52
         Top             =   1320
         Width           =   300
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Results outside this range will be marked as implausible."
         Height          =   195
         Left            =   -73470
         TabIndex        =   51
         Top             =   2160
         Width           =   3930
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
         Top             =   2520
         Width           =   4395
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
Attribute VB_Name = "frmBgaDefaults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Activated As Boolean
Private loading As Boolean

Private FromDays() As Long
Private ToDays() As Long

Private Sub cEOD_Click()
          Dim sql As String



10        On Error GoTo cEOD_Click_Error

20        If lblHost = "" Then Exit Sub


30        If cEOD.Value = 1 Then
40            sql = "UPDATE bgatestdefinitions set eod = 1 WHERE code = '" & lblHost & "'"
50            Cnxn(0).Execute sql
60        ElseIf cEOD.Value = 0 Then
70            sql = "UPDATE bgatestdefinitions set eod = 0 WHERE code = '" & lblHost & "'"
80            Cnxn(0).Execute sql
90        End If



100       Exit Sub

cEOD_Click_Error:

          Dim strES As String
          Dim intEL As Integer

110       intEL = Erl
120       strES = Err.Description
130       LogError "frmBgaDefaults", "cEOD_Click", intEL, strES, sql


End Sub

Private Sub cInUse_Click()
          Dim sql As String



10        On Error GoTo cInUse_Click_Error

20        If lblHost = "" Then Exit Sub


30        If cInUse.Value = 1 Then
40            sql = "UPDATE bgatestdefinitions set inuse = 1 WHERE code = '" & lblHost & "'"
50            Cnxn(0).Execute sql
60        ElseIf cInUse.Value = 0 Then
70            sql = "UPDATE bgatestdefinitions set inuse = 0 WHERE code = '" & lblHost & "'"
80            Cnxn(0).Execute sql
90        End If

100       Exit Sub


110       Exit Sub

cInUse_Click_Error:

          Dim strES As String
          Dim intEL As Integer

120       intEL = Erl
130       strES = Err.Description
140       LogError "frmBgaDefaults", "cInUse_Click", intEL, strES, sql


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
70        LogError "frmBgaDefaults", "cKnown_MouseUp", intEL, strES


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
90        LogError "frmBgaDefaults", "cmbUnits_Click", intEL, strES


End Sub

Private Sub cmdAmendAgeRange_Click()



10        On Error GoTo cmdAmendAgeRange_Click_Error

20        If lstParameter = "" Then
30            iMsg "Please pick a Test!"
40            Exit Sub
50        End If

60        With frmAges
70            .Analyte = lstParameter
80            If oSU(0).Value = True Then
90                .SampleType = "S"
100           ElseIf oSU(1).Value = True Then
110               .SampleType = "U"
120           ElseIf oSU(2).Value = True Then
130               .SampleType = "C"
140           End If
150           .Discipline = "Blood Gas"
160           .Show 1
170       End With

180       FillAges

190       FillAllDetails



200       Exit Sub

cmdAmendAgeRange_Click_Error:

          Dim strES As String
          Dim intEL As Integer

210       intEL = Erl
220       strES = Err.Description
230       LogError "frmBgaDefaults", "cmdAmendAgeRange_Click", intEL, strES


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
80        LogError "frmBgaDefaults", "cmdUpdate_Click", intEL, strES


End Sub

Private Sub FillAges()

          Dim tb As New Recordset
          Dim s As String
          Dim SampleType As String
          Dim n As Long
          Dim sql As String



10        On Error GoTo FillAges_Error

20        ClearFGrid grdAge

30        If oSU(0).Value = True Then
40            SampleType = "S"
50        ElseIf oSU(1).Value = True Then
60            SampleType = "U"

70        End If

80        ReDim FromDays(0 To 0)
90        ReDim ToDays(0 To 0)


100       If oSU(2) = True Then
110           sql = "SELECT * from BgaTestDefinitions WHERE " & _
                    "LongName = '" & lstParameter & "' " & _
                    "and (SampleType <> 'S' and sampletype <> 'U')" & _
                    "Order by AgeFromDays"
120       Else
130           sql = "SELECT * from BgaTestDefinitions WHERE " & _
                    "LongName = '" & lstParameter & "' " & _
                    "and SampleType = '" & SampleType & "' " & _
                    "Order by AgeFromDays"
140       End If

150       Set tb = New Recordset
160       RecOpenClient 0, tb, sql

170       If tb.EOF Then Exit Sub
180       ReDim FromDays(0 To tb.RecordCount - 1)
190       ReDim ToDays(0 To tb.RecordCount - 1)
200       n = 0
210       Do While Not tb.EOF
220           If tb!AgeFromDays & "" = "" Then FromDays(n) = 0 Else FromDays(n) = tb!AgeFromDays
230           If tb!AgeToDays & "" = "" Then ToDays(n) = 43870 Else ToDays(n) = tb!AgeToDays
240           s = Format$(n) & vbTab & _
                  dmyFromCount(FromDays(n)) & vbTab & _
                  dmyFromCount(ToDays(n))
250           grdAge.AddItem s
260           n = n + 1
270           tb.MoveNext
280       Loop

290       FixG grdAge

300       grdAge.Col = 0
310       grdAge.Row = 1
320       grdAge.CellBackColor = vbYellow
330       grdAge.CellForeColor = vbBlue




340       Exit Sub

FillAges_Error:

          Dim strES As String
          Dim intEL As Integer

350       intEL = Erl
360       strES = Err.Description
370       LogError "frmBgaDefaults", "FillAges", intEL, strES, sql


End Sub

Private Sub FillAllDetails()

          Dim tb As New Recordset
          Dim sql As String
          Dim SampleType As String
          Dim n As Long



10        On Error GoTo FillAllDetails_Error

20        tPriority = ""
30        cKnown = 0
40        oDelta = 0
50        tDP = 0
60        tdelta = ""
70        loading = True
80        cmbUnits.ListIndex = -1
90        loading = False
100       lblBarCode = ""
110       lblHost = ""
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
220       cInUse = 0
230       cEOD = 0

240       If oSU(0).Value = True Then
250           SampleType = "S"
260       ElseIf oSU(1).Value = True Then
270           SampleType = "U"
280       End If

290       If Trim(FromDays(0) & "") = "" Then FromDays(0) = 0
300       If Trim(ToDays(0) & "") = "" Then ToDays(0) = 43830

310       If oSU(2).Value = True Then
320           sql = "SELECT * from BgaTestDefinitions WHERE " & _
                    "LongName = '" & lstParameter & "' " & _
                    "and AgeFromDays = '" & FromDays(0) & "' " & _
                    "and AgeToDays = '" & ToDays(0) & "' " & _
                    "and (SampleType <> 'S' and sampletype <> 'U')"
330       Else
340           sql = "SELECT * from BgaTestDefinitions WHERE " & _
                    "LongName = '" & lstParameter & "' " & _
                    "and AgeFromDays = '" & FromDays(0) & "' " & _
                    "and AgeToDays = '" & ToDays(0) & "' " & _
                    "and SampleType = '" & SampleType & "'"
350       End If
360       Set tb = New Recordset
370       RecOpenServer 0, tb, sql
380       With tb

390           If Not .EOF Then
400               tPriority = !PrintPriority
410               cKnown = IIf(!KnownToAnalyser, 1, 0)
420               If !DoDelta & "" <> "" Then oDelta = IIf(!DoDelta, 1, 0)
430               tDP = !DP
440               tdelta = !DeltaLimit & ""
450               If Trim(!Units) & "" <> "" Then
460                   For n = 0 To cmbUnits.ListCount
470                       If !Units & "" = cmbUnits.List(n) Then
480                           loading = True
490                           cmbUnits.ListIndex = n
500                           loading = False
510                       End If
520                   Next
530               End If

540               lblBarCode = !BarCode & ""

550               tPlausibleHigh = !PlausibleHigh
560               tPlausibleLow = !PlausibleLow
570               lblHost = !Code
580               tMaleHigh = !MaleHigh
590               tMaleLow = !MaleLow
600               tFemaleHigh = !FemaleHigh
610               tFemaleLow = !FemaleLow
620               tfr(0) = !FlagMaleHigh & ""
630               tfr(1) = !FlagMaleLow & ""
640               tfr(2) = !FlagFemaleHigh & ""
650               tfr(3) = !FlagFemaleLow & ""
660               cInUse = IIf(!InUse, 1, 0)
670               If Trim(!Eod) & "" = "" Then cEOD = 0 Else cEOD = IIf(!Eod, 1, 0)

680           End If
690       End With




700       Exit Sub

FillAllDetails_Error:

          Dim strES As String
          Dim intEL As Integer

710       intEL = Erl
720       strES = Err.Description
730       LogError "frmBgaDefaults", "FillAllDetails", intEL, strES, sql


End Sub

Private Sub FillDetails()
          Dim n As Long

          Dim tb As New Recordset
          Dim sql As String
          Dim SampleType As String
          Dim AgeNumber As Long
10        On Error GoTo FillDetails_Error

20        On Error GoTo FillDetails_Error

30        If oSU(0).Value = True Then
40            SampleType = "S"
50        ElseIf oSU(1).Value = True Then
60            SampleType = "U"

70        End If

80        AgeNumber = -1
90        grdAge.Col = 0
100       For AgeNumber = 1 To grdAge.Rows - 1
110           grdAge.Row = AgeNumber
120           If grdAge.CellBackColor = vbYellow Then
130               AgeNumber = AgeNumber - 1
140               Exit For
150           End If
160       Next
170       If AgeNumber = -1 Then
180           iMsg "SELECT Age Range", vbCritical
190           Exit Sub
200       End If

210       sql = "SELECT * from BgaTestDefinitions WHERE " & _
                "LongName = '" & lstParameter & "' " & _
                "and AgeFromDays = '" & FromDays(AgeNumber) & "' " & _
                "and AgeToDays = '" & ToDays(AgeNumber) & "' " & _
                "and SampleType = '" & SampleType & "'"
220       Set tb = New Recordset
230       RecOpenServer 0, tb, sql
240       With tb

250           If Not .EOF Then
260               tPriority = !PrintPriority
270               cKnown = IIf(!KnownToAnalyser, 1, 0)
280               If !DoDelta & "" <> "" Then oDelta = IIf(!DoDelta, 1, 0)
290               tDP = !DP
300               tdelta = !DeltaLimit

310               If Trim(!Units) & "" <> "" Then
320                   For n = 0 To cmbUnits.ListCount
330                       If !Units & "" = cmbUnits.List(n) Then
340                           loading = True
350                           cmbUnits.ListIndex = n
360                           loading = False
370                       End If
380                   Next
390               End If
400               lblBarCode = !BarCode & ""

410               tMaleHigh = !MaleHigh
420               tMaleLow = !MaleLow
430               tFemaleHigh = !FemaleHigh
440               tFemaleLow = !FemaleLow
450               tfr(0) = !FlagMaleHigh
460               tfr(1) = !FlagMaleLow
470               tfr(2) = !FlagFemaleHigh
480               tfr(3) = !FlagFemaleLow

490           End If
500       End With



510       Exit Sub

FillDetails_Error:

          Dim strES As String
          Dim intEL As Integer

520       intEL = Erl
530       strES = Err.Description
540       LogError "frmBgaDefaults", "FillDetails", intEL, strES


End Sub

Private Sub FilllParameter()

          Dim tb As New Recordset
          Dim sql As String
          Dim SampleType As String






10        On Error GoTo FilllParameter_Error

20        If oSU(0).Value = True Then
30            SampleType = "S"
40        ElseIf oSU(1).Value = True Then
50            SampleType = "U"
60        End If


70        lstParameter.Clear

80        If oSU(2).Value = True Then
90            sql = "SELECT distinct LongName, PrintPriority from BgaTestDefinitions WHERE " & _
                    "(SampleType <> 'S' and sampletype <> 'U') " & _
                    "order by PrintPriority"
100       Else
110           sql = "SELECT distinct LongName, PrintPriority from BgaTestDefinitions WHERE " & _
                    "SampleType = '" & SampleType & "' " & _
                    "order by PrintPriority"
120       End If
130       Set tb = New Recordset
140       RecOpenServer 0, tb, sql
150       Do While Not tb.EOF
160           lstParameter.AddItem tb!LongName
170           tb.MoveNext
180       Loop

190       lstParameter.ListIndex = -1



200       Exit Sub

FilllParameter_Error:

          Dim strES As String
          Dim intEL As Integer

210       intEL = Erl
220       strES = Err.Description
230       LogError "frmBgaDefaults", "FilllParameter", intEL, strES, sql


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
140       LogError "frmBgaDefaults", "FillUnits", intEL, strES, sql


End Sub

Private Sub Form_Activate()

10        On Error GoTo Form_Activate_Error

20        If Activated Then Exit Sub

30        Activated = True

40        FilllParameter
50        FillUnits

60        Exit Sub

Form_Activate_Error:

          Dim strES As String
          Dim intEL As Integer

70        intEL = Erl
80        strES = Err.Description
90        LogError "frmBgaDefaults", "Form_Activate", intEL, strES


End Sub

Private Sub Form_Load()

10        On Error GoTo Form_Load_Error

20        Activated = False


30        grdAge.Font.Bold = True

40        Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmBgaDefaults", "Form_Load", intEL, strES


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
60        LogError "frmBgaDefaults", "Form_Unload", intEL, strES


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
170       LogError "frmBgaDefaults", "grdAge_MouseUp", intEL, strES


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
80        LogError "frmBgaDefaults", "lblBarCode_Click", intEL, strES


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
90        LogError "frmBgaDefaults", "lstParameter_Click", intEL, strES


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
70        LogError "frmBgaDefaults", "odelta_MouseUp", intEL, strES


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
70        LogError "frmBgaDefaults", "orm_MouseUp", intEL, strES


End Sub

Private Sub oSU_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

10        On Error GoTo oSU_MouseUp_Error

20        FilllParameter

30        Exit Sub

oSU_MouseUp_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmBgaDefaults", "oSU_MouseUp", intEL, strES


End Sub

Private Sub SaveCommonDetails()

          Dim sql As String
          Dim SampleType As String



10        On Error GoTo SaveCommonDetails_Error

20        If oSU(0).Value = True Then
30            SampleType = "S"
40        ElseIf oSU(1).Value = True Then
50            SampleType = "U"

60        End If

70        sql = "UPDATE BgaTestDefinitions SET " & _
                "KnownToAnalyser = '" & IIf(cKnown = 1, 1, 0) & "', " & _
                "PrintPriority = '" & Val(tPriority) & "', " & _
                "DP = '" & Val(tDP) & "', " & _
                "DoDelta = '" & oDelta & "', " & _
                "DeltaLimit ='" & Val(tdelta) & "', " & _
                "Units = '" & cmbUnits & "', " & _
                "BarCode = '" & lblBarCode & "' " & _
                "WHERE LongName = '" & lstParameter & "' " & _
                "and Hospital = '" & HospName(0) & "'"

80        Cnxn(0).Execute sql


90        Exit Sub

SaveCommonDetails_Error:

          Dim strES As String
          Dim intEL As Integer

100       intEL = Erl
110       strES = Err.Description
120       LogError "frmBgaDefaults", "SaveCommonDetails", intEL, strES


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

150       If oSU(0).Value = True Then
160           SampleType = "S"
170       ElseIf oSU(1).Value = True Then
180           SampleType = "U"
190       End If

200       sql = "UPDATE BgaTestDefinitions " & _
                "Set MaleLow = '" & tMaleLow & "', " & _
                "MaleHigh = '" & tMaleHigh & "', " & _
                "FemaleLow = '" & tFemaleLow & "', " & _
                "FemaleHigh = '" & tFemaleHigh & "' WHERE " & _
                "LongName = '" & lstParameter & "' " & _
                "and AgeFromDays = '" & FromDays(AgeNumber) & "' " & _
                "and AgeToDays = '" & ToDays(AgeNumber) & "' " & _
                " " & _
                "and Hospital = '" & HospName(0) & "'"

210       Cnxn(0).Execute sql




220       Exit Sub

SaveNormals_Error:

          Dim strES As String
          Dim intEL As Integer

230       intEL = Erl
240       strES = Err.Description
250       LogError "frmBgaDefaults", "SaveNormals", intEL, strES, sql


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
70        LogError "frmBgaDefaults", "tDelta_KeyUp", intEL, strES


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
70        LogError "frmBgaDefaults", "tDP_KeyUp", intEL, strES


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
60        LogError "frmBgaDefaults", "tFemaleHigh_KeyUp", intEL, strES


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
60        LogError "frmBgaDefaults", "tFemaleLow_KeyUp", intEL, strES


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
60        grdAge.Col = 0
70        For n = 1 To grdAge.Rows - 1
80            grdAge.Row = n
90            If grdAge.CellBackColor = vbYellow Then
100               AgeNumber = grdAge.TextMatrix(n, 0)
110               Exit For
120           End If
130       Next
140       If AgeNumber = -1 Then
150           iMsg "SELECT Age Range", vbCritical
160           Exit Sub
170       End If

180       If oSU(0).Value = True Then
190           SampleType = "S"
200       ElseIf oSU(1).Value = True Then
210           SampleType = "U"
220       ElseIf oSU(2).Value = True Then
230           SampleType = "C"
240       End If

250       sql = "UPDATE BgaTestDefinitions " & _
                "Set FlagMaleLow = '" & tfr(1) & "', " & _
                "FlagMaleHigh = '" & tfr(0) & "', " & _
                "FlagFemaleLow = '" & tfr(3) & "', " & _
                "FlagFemaleHigh = '" & tfr(2) & "' WHERE " & _
                "LongName = '" & lstParameter & "' " & _
                "and AgeFromDays = '" & FromDays(AgeNumber) & "' " & _
                "and AgeToDays = '" & ToDays(AgeNumber) & "' " & _
                "and SampleType = '" & SampleType & "' " & _
                "and Hospital = '" & HospName(0) & "'"

260       Cnxn(0).Execute sql


270       Exit Sub

tfr_KeyUp_Error:

          Dim strES As String
          Dim intEL As Integer

280       intEL = Erl
290       strES = Err.Description
300       LogError "frmBgaDefaults", "tfr_KeyUp", intEL, strES, sql


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
60        LogError "frmBgaDefaults", "tMaleHigh_KeyUp", intEL, strES


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
60        LogError "frmBgaDefaults", "tMaleLow_KeyUp", intEL, strES


End Sub

Private Sub tPlausibleHigh_KeyUp(KeyCode As Integer, Shift As Integer)

          Dim sql As String



10        On Error GoTo tPlausibleHigh_KeyUp_Error

20        sql = "UPDATE BgaTestDefinitions " & _
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
70        LogError "frmBgaDefaults", "tPlausibleHigh_KeyUp", intEL, strES, sql


End Sub

Private Sub tPlausibleLow_KeyUp(KeyCode As Integer, Shift As Integer)

          Dim sql As String



10        On Error GoTo tPlausibleLow_KeyUp_Error

20        sql = "UPDATE BgaTestDefinitions " & _
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
70        LogError "frmBgaDefaults", "tPlausibleLow_KeyUp", intEL, strES, sql


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
70        LogError "frmBgaDefaults", "tPriority_KeyUp", intEL, strES


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
70        LogError "frmBgaDefaults", "UpDown1_MouseUp", intEL, strES


End Sub
