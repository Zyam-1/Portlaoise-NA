VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmCoagDefinitions 
   Caption         =   "NetAcquire - Coagulation Definitions"
   ClientHeight    =   6690
   ClientLeft      =   465
   ClientTop       =   1590
   ClientWidth     =   8400
   Icon            =   "frmCoagDefinitions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   8400
   Begin VB.Frame Frame7 
      Caption         =   "Specifics (Applies to all age ranges)"
      Height          =   1590
      Left            =   1650
      TabIndex        =   5
      Top             =   5025
      Width           =   5325
      Begin VB.TextBox lblPrint 
         Height          =   285
         Left            =   1170
         TabIndex        =   61
         Top             =   1260
         Width           =   645
      End
      Begin VB.CheckBox chkInuse 
         Alignment       =   1  'Right Justify
         Caption         =   "In Use"
         Height          =   225
         Left            =   4230
         TabIndex        =   60
         Top             =   630
         Width           =   825
      End
      Begin VB.TextBox txtAnalyser 
         Height          =   315
         Left            =   3780
         Locked          =   -1  'True
         TabIndex        =   58
         Top             =   1215
         Width           =   1335
      End
      Begin MSComCtl2.UpDown upPP 
         Height          =   285
         Left            =   1831
         TabIndex        =   57
         Top             =   1260
         Width           =   240
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         BuddyControl    =   "lblPrint"
         BuddyDispid     =   196610
         OrigLeft        =   2115
         OrigTop         =   1260
         OrigRight       =   2355
         OrigBottom      =   1545
         Max             =   1000
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.TextBox tTestName 
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   600
         Width           =   1695
      End
      Begin VB.ComboBox cmbUnits 
         Height          =   315
         Left            =   3780
         TabIndex        =   8
         Text            =   "cUnits"
         Top             =   210
         Width           =   1425
      End
      Begin VB.TextBox tCode 
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   240
         Width           =   1695
      End
      Begin VB.CheckBox cPrintable 
         Alignment       =   1  'Right Justify
         Caption         =   "Printable"
         Height          =   225
         Left            =   3060
         TabIndex        =   6
         Top             =   630
         Value           =   1  'Checked
         Width           =   915
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Analyser Code"
         Height          =   195
         Left            =   2745
         TabIndex        =   59
         Top             =   1260
         Width           =   1020
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Print Priority"
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   56
         Top             =   1290
         Width           =   825
      End
      Begin VB.Label lblBarCode 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3780
         TabIndex        =   55
         Top             =   930
         Width           =   1365
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "BarCode"
         Height          =   195
         Left            =   3120
         TabIndex        =   54
         Top             =   960
         Width           =   615
      End
      Begin VB.Label lDP 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         Height          =   285
         Left            =   1200
         TabIndex        =   14
         Top             =   930
         Width           =   855
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Decimal Points"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   13
         Top             =   960
         Width           =   1050
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Test Name"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   12
         Top             =   660
         Width           =   780
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Units"
         Height          =   195
         Left            =   3390
         TabIndex        =   11
         Top             =   270
         Width           =   360
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Code"
         Height          =   195
         Left            =   780
         TabIndex        =   10
         Top             =   270
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Age Ranges"
      Height          =   2115
      Left            =   1620
      TabIndex        =   2
      Top             =   165
      Width           =   6615
      Begin VB.CommandButton cmdAmendAgeRange 
         Caption         =   "Amend Age Range"
         Height          =   885
         Left            =   4890
         Picture         =   "frmCoagDefinitions.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   510
         Width           =   1155
      End
      Begin MSFlexGridLib.MSFlexGrid grdAge 
         Height          =   1725
         Left            =   570
         TabIndex        =   4
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
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   735
      Left            =   7110
      Picture         =   "frmCoagDefinitions.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5850
      Width           =   1155
   End
   Begin VB.ListBox lstParameter 
      Height          =   4665
      IntegralHeight  =   0   'False
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   1155
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2625
      Left            =   1650
      TabIndex        =   15
      Top             =   2325
      Width           =   5325
      _ExtentX        =   9393
      _ExtentY        =   4630
      _Version        =   393216
      Tabs            =   4
      Tab             =   3
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
      TabPicture(0)   =   "frmCoagDefinitions.frx":091E
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "tFemaleLow"
      Tab(0).Control(1)=   "tMaleLow"
      Tab(0).Control(2)=   "tFemaleHigh"
      Tab(0).Control(3)=   "tMaleHigh"
      Tab(0).Control(4)=   "Label6"
      Tab(0).Control(5)=   "Label4"
      Tab(0).Control(6)=   "Label3"
      Tab(0).Control(7)=   "Label2"
      Tab(0).Control(8)=   "Label1(1)"
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Flag Range"
      TabPicture(1)   =   "frmCoagDefinitions.frx":093A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "tFlagFemaleLow"
      Tab(1).Control(1)=   "tFlagMaleHigh"
      Tab(1).Control(2)=   "tFlagFemaleHigh"
      Tab(1).Control(3)=   "tFlagMaleLow"
      Tab(1).Control(4)=   "Label7(1)"
      Tab(1).Control(5)=   "Label12(2)"
      Tab(1).Control(6)=   "Label13(1)"
      Tab(1).Control(7)=   "Label14(2)"
      Tab(1).Control(8)=   "Label15(1)"
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "Plausible"
      TabPicture(2)   =   "frmCoagDefinitions.frx":0956
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "tPlausibleLow"
      Tab(2).Control(1)=   "tPlausibleHigh"
      Tab(2).Control(2)=   "Label10(2)"
      Tab(2).Control(3)=   "Label9(1)"
      Tab(2).Control(4)=   "Label8(1)"
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "Delta Check"
      TabPicture(3)   =   "frmCoagDefinitions.frx":0972
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Label20"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Label18"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Label21"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "oDelta"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "tDelta"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "txtCheckTime"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).ControlCount=   6
      Begin VB.TextBox txtCheckTime 
         Height          =   285
         Left            =   2310
         MaxLength       =   5
         TabIndex        =   62
         Top             =   1620
         Width           =   495
      End
      Begin VB.TextBox tFlagFemaleLow 
         Height          =   315
         Left            =   -73500
         TabIndex        =   53
         Top             =   1470
         Width           =   915
      End
      Begin VB.TextBox tfr 
         Height          =   315
         Index           =   0
         Left            =   -71640
         TabIndex        =   30
         Top             =   1200
         Width           =   915
      End
      Begin VB.TextBox tfr 
         Height          =   315
         Index           =   2
         Left            =   -73170
         TabIndex        =   29
         Top             =   1200
         Width           =   915
      End
      Begin VB.TextBox tfr 
         Height          =   315
         Index           =   1
         Left            =   -71640
         TabIndex        =   28
         Top             =   1800
         Width           =   915
      End
      Begin VB.TextBox tfr 
         Height          =   315
         Index           =   3
         Left            =   -73170
         TabIndex        =   27
         Top             =   1800
         Width           =   915
      End
      Begin VB.TextBox tFemaleLow 
         Height          =   315
         Left            =   -73500
         MaxLength       =   5
         TabIndex        =   26
         Top             =   1470
         Width           =   915
      End
      Begin VB.TextBox tMaleLow 
         Height          =   315
         Left            =   -71970
         MaxLength       =   5
         TabIndex        =   25
         Top             =   1470
         Width           =   915
      End
      Begin VB.TextBox tFemaleHigh 
         Height          =   315
         Left            =   -73500
         MaxLength       =   5
         TabIndex        =   24
         Top             =   870
         Width           =   915
      End
      Begin VB.TextBox tMaleHigh 
         Height          =   315
         Left            =   -71970
         MaxLength       =   5
         TabIndex        =   23
         Top             =   870
         Width           =   915
      End
      Begin VB.TextBox tPlausibleLow 
         Height          =   285
         Left            =   -73080
         TabIndex        =   22
         Top             =   1290
         Width           =   1215
      End
      Begin VB.TextBox tPlausibleHigh 
         Height          =   285
         Left            =   -73080
         TabIndex        =   21
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox tDelta 
         Height          =   285
         Left            =   2310
         MaxLength       =   5
         TabIndex        =   20
         Top             =   1320
         Width           =   735
      End
      Begin VB.CheckBox oDelta 
         Alignment       =   1  'Right Justify
         Caption         =   "Enabled"
         Height          =   195
         Left            =   1950
         TabIndex        =   19
         Top             =   930
         Width           =   915
      End
      Begin VB.TextBox tFlagMaleHigh 
         Height          =   315
         Left            =   -71970
         TabIndex        =   18
         Top             =   870
         Width           =   915
      End
      Begin VB.TextBox tFlagFemaleHigh 
         Height          =   315
         Left            =   -73500
         TabIndex        =   17
         Top             =   870
         Width           =   915
      End
      Begin VB.TextBox tFlagMaleLow 
         Height          =   315
         Left            =   -71970
         TabIndex        =   16
         Top             =   1470
         Width           =   915
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "day(s)"
         Height          =   195
         Left            =   2820
         TabIndex        =   64
         Top             =   1650
         Width           =   420
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Check Time"
         Height          =   195
         Left            =   1350
         TabIndex        =   63
         Top             =   1650
         Width           =   855
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Low"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   0
         Left            =   -73080
         TabIndex        =   52
         Top             =   1590
         Width           =   300
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "High"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   0
         Left            =   -73110
         TabIndex        =   51
         Top             =   1020
         Width           =   330
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Result values outside this range will be flagged as High or Low"
         Height          =   195
         Index           =   0
         Left            =   -74190
         TabIndex        =   50
         Top             =   2520
         Width           =   4410
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "These are the normal range values printed on the report forms."
         Height          =   195
         Left            =   -74490
         TabIndex        =   49
         Top             =   1980
         Width           =   4395
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Do not print the result if the sample is :-"
         Height          =   195
         Index           =   1
         Left            =   -74070
         TabIndex        =   48
         Top             =   750
         Width           =   2730
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Female"
         Height          =   195
         Index           =   1
         Left            =   -73020
         TabIndex        =   47
         Top             =   930
         Width           =   510
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Male"
         Height          =   195
         Index           =   0
         Left            =   -71370
         TabIndex        =   46
         Top             =   900
         Width           =   345
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "High"
         Height          =   195
         Index           =   1
         Left            =   -73950
         TabIndex        =   45
         Top             =   1290
         Width           =   330
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Low"
         Height          =   195
         Index           =   0
         Left            =   -73920
         TabIndex        =   44
         Top             =   1860
         Width           =   300
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Low"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   -74070
         TabIndex        =   43
         Top             =   1530
         Width           =   300
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "High"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   -74070
         TabIndex        =   42
         Top             =   930
         Width           =   330
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Female"
         Height          =   195
         Left            =   -73350
         TabIndex        =   41
         Top             =   600
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Male"
         Height          =   195
         Index           =   1
         Left            =   -71700
         TabIndex        =   40
         Top             =   570
         Width           =   345
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Results outside this range will be marked as implausible."
         Height          =   195
         Index           =   2
         Left            =   -74220
         TabIndex        =   39
         Top             =   1950
         Width           =   3930
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Low"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   1
         Left            =   -73560
         TabIndex        =   38
         Top             =   1320
         Width           =   300
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "High"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   1
         Left            =   -73590
         TabIndex        =   37
         Top             =   750
         Width           =   330
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Value"
         Height          =   195
         Left            =   1800
         TabIndex        =   36
         Top             =   1350
         Width           =   405
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Result values outside this range will be flagged as High or Low"
         Height          =   195
         Index           =   1
         Left            =   -74520
         TabIndex        =   35
         Top             =   1980
         Width           =   4410
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Female"
         Height          =   195
         Index           =   2
         Left            =   -73350
         TabIndex        =   34
         Top             =   600
         Width           =   510
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Male"
         Height          =   195
         Index           =   1
         Left            =   -71700
         TabIndex        =   33
         Top             =   570
         Width           =   345
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "High"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   2
         Left            =   -74070
         TabIndex        =   32
         Top             =   930
         Width           =   330
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Low"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   1
         Left            =   -74070
         TabIndex        =   31
         Top             =   1530
         Width           =   300
      End
   End
End
Attribute VB_Name = "frmCoagDefinitions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FromDays() As Long
Private ToDays() As Long

Private Sub chkInuse_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        On Error GoTo chkInuse_MouseUp_Error

20        SaveDetails

30        Exit Sub

chkInuse_MouseUp_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmCoagDefinitions", "chkInuse_MouseUp", intEL, strES


End Sub



Private Sub cmbUnits_KeyUp(KeyCode As Integer, Shift As Integer)

10        On Error GoTo cmbUnits_KeyUp_Error

20        SaveDetails

30        Exit Sub

cmbUnits_KeyUp_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmCoagDefinitions", "cmbUnits_KeyUp", intEL, strES


End Sub

Private Sub cmdAmendAgeRange_Click()



10        On Error GoTo cmdAmendAgeRange_Click_Error

20        If lstParameter = "" Then
30            iMsg "SELECT Parameter", vbCritical
40            Exit Sub
50        End If

60        With frmAges
70            .Analyte = lstParameter
80            .SampleType = "Coagulation"
90            .Discipline = "Coagulation"
100           .Show 1
110       End With

120       FillAges




130       Exit Sub

cmdAmendAgeRange_Click_Error:

          Dim strES As String
          Dim intEL As Integer

140       intEL = Erl
150       strES = Err.Description
160       LogError "frmCoagDefinitions", "cmdAmendAgeRange_Click", intEL, strES


End Sub

Private Sub cmdCancel_Click()

10        Unload Me

End Sub

Private Sub FillAges()

          Dim tb As New Recordset
          Dim s As String
          Dim n As Long
          Dim sql As String

10        On Error GoTo FillAges_Error

20        ClearFGrid grdAge

30        ReDim FromDays(0 To 0)
40        ReDim ToDays(0 To 0)

50        sql = "SELECT * from CoagTestDefinitions WHERE " & _
                "TestName = '" & lstParameter & "' " & _
                "and Hospital = '" & HospName(0) & "' " & _
                "Order by cast(AgetoDays as numeric) asc"
60        Set tb = New Recordset
70        RecOpenClient 0, tb, sql

80        If Not tb.EOF Then

90            ReDim FromDays(0 To tb.RecordCount - 1)
100           ReDim ToDays(0 To tb.RecordCount - 1)
110           n = 0
120           Do While Not tb.EOF
130               FromDays(n) = tb!AgeFromDays
140               ToDays(n) = tb!AgeToDays
150               s = Format$(n) & vbTab & _
                      dmyFromCount(FromDays(n)) & vbTab & _
                      dmyFromCount(ToDays(n))
160               grdAge.AddItem s
170               n = n + 1
180               tb.MoveNext
190           Loop

200       End If

210       FixG grdAge

220       grdAge.Col = 0
230       grdAge.Row = 1
240       grdAge.CellBackColor = vbYellow
250       grdAge.CellForeColor = vbBlue

260       Exit Sub

FillAges_Error:

          Dim strES As String
          Dim intEL As Integer

270       intEL = Erl
280       strES = Err.Description
290       LogError "frmCoagDefinitions", "FillAges", intEL, strES, sql

End Sub

Private Sub FillDetails()

          Dim tb As New Recordset
          Dim sql As String
          Dim Filled As Boolean
          Dim AgeNumber As Long
          Dim Y As Long



10        On Error GoTo FillDetails_Error

20        Filled = False

30        AgeNumber = -1
40        grdAge.Col = 0
50        For Y = 1 To grdAge.Rows - 1
60            grdAge.Row = Y
70            If grdAge.CellBackColor = vbYellow Then
80                AgeNumber = Y - 1
90                Exit For
100           End If
110       Next
120       If AgeNumber = -1 Then
130           iMsg "SELECT Age Range", vbCritical
140           Exit Sub
150       End If

160       tCode = ""
170       tTestName = ""
180       oDelta = 0
190       tdelta = ""
200       lDP = "0"
210       cmbUnits = ""
220       cPrintable = 0
230       tPlausibleLow = ""
240       tPlausibleHigh = ""
250       tMaleHigh = ""
260       tFemaleHigh = ""
270       tMaleLow = ""
280       tFemaleLow = ""
290       lblBarCode = ""
300       lblPrint = ""
310       txtCheckTime = ""

320       sql = "SELECT * from CoagTestDefinitions WHERE " & _
                "TestName = '" & lstParameter & "' " & _
                "and AgeFromDays = '" & FromDays(AgeNumber) & "' " & _
                "and AgeToDays = '" & ToDays(AgeNumber) & "' " & _
                "And Hospital = '" & HospName(0) & "'"
330       Set tb = New Recordset
340       RecOpenServer 0, tb, sql
350       If Not tb.EOF Then
360           tCode = tb!Code
370           tTestName = lstParameter
380           If tb!DoDelta & "" <> "" Then oDelta = IIf(tb!DoDelta, 1, 0)
390           If Not IsNull(tb!DeltaLimit) Then tdelta = tb!DeltaLimit
400           lDP = tb!DP
410           cmbUnits = tb!Units
420           If Not IsNull(tb!Printable) Then cPrintable = IIf(tb!Printable, 1, 0)
430           If Not IsNull(tb!InUse) Then chkInuse = IIf(tb!InUse, 1, 0)
440           tPlausibleLow = tb!PlausibleLow & ""
450           tPlausibleHigh = tb!PlausibleHigh & ""
460           tMaleHigh = tb!MaleHigh & ""
470           tFemaleHigh = tb!FemaleHigh & ""
480           tMaleLow = tb!MaleLow & ""
490           tFemaleLow = tb!FemaleLow & ""
500           tFlagMaleHigh = tb!FlagMaleHigh & ""
510           tFlagFemaleHigh = tb!FlagFemaleHigh & ""
520           tFlagMaleLow = tb!FlagMaleLow & ""
530           tFlagFemaleLow = tb!FlagFemaleLow & ""
540           lblBarCode = tb!BarCode & ""
550           lblPrint = tb!PrintPriority & ""
560           If IsNull(tb!CheckTime) Then
570               txtCheckTime = ""
580           Else
590               txtCheckTime = tb!CheckTime
600           End If
610       End If


620       Exit Sub

FillDetails_Error:

          Dim strES As String
          Dim intEL As Integer

630       intEL = Erl
640       strES = Err.Description
650       LogError "frmCoagDefinitions", "FillDetails", intEL, strES, sql


End Sub

Private Sub FillParameters()

          Dim InList As Boolean
          Dim n As Long
          Dim sql As String
          Dim tb As New Recordset

10        On Error GoTo FillParameters_Error

20        lstParameter.Clear


30        sql = "SELECT * from coagtestdefinitions order by PrintPriority"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        Do While Not tb.EOF
70            InList = False
80            For n = 0 To lstParameter.ListCount - 1
90                If lstParameter.List(n) = Trim(tb!TestName) Then
100                   InList = True
110                   Exit For
120               End If
130           Next
140           If Not InList Then
150               lstParameter.AddItem Trim(tb!TestName)
160           End If
170           tb.MoveNext
180       Loop

190       If lstParameter.ListCount > 0 Then
200           lstParameter.Selected(0) = True
210       End If



220       Exit Sub

FillParameters_Error:

          Dim strES As String
          Dim intEL As Integer

230       intEL = Erl
240       strES = Err.Description
250       LogError "frmCoagDefinitions", "FillParameters", intEL, strES, sql


End Sub

Private Sub FillScreen()

          Dim tb As New Recordset
          Dim sql As String

10        On Error GoTo FillScreen_Error

20        cmbUnits.Clear
30        sql = "SELECT * from lists WHERE listtype = 'UN'"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql

60        Do While Not tb.EOF
70            cmbUnits.AddItem Trim(tb!Text)
80            tb.MoveNext
90        Loop
100       If cmbUnits.ListCount > 0 Then
110           cmbUnits.ListIndex = 0
120       End If

130       FillParameters
140       FillAges
150       FillDetails

160       Exit Sub

FillScreen_Error:

          Dim strES As String
          Dim intEL As Integer

170       intEL = Erl
180       strES = Err.Description
190       LogError "frmCoagDefinitions", "FillScreen", intEL, strES, sql

End Sub

Private Sub cPrintable_Click()

'SaveDetails

End Sub

Private Sub cPrintable_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        On Error GoTo cPrintable_MouseUp_Error

20        SaveDetails

30        Exit Sub

cPrintable_MouseUp_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmCoagDefinitions", "cPrintable_MouseUp", intEL, strES


End Sub

Private Sub Form_Load()

10        On Error GoTo Form_Load_Error

20        FillScreen

30        Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmCoagDefinitions", "Form_Load", intEL, strES

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
170       LogError "frmCoagDefinitions", "grdAge_MouseUp", intEL, strES


End Sub

Private Sub lblBarCode_Click()

10        On Error GoTo lblBarCode_Click_Error

20        lblBarCode = iBOX("Scan in Bar Code", , lblBarCode)

30        SaveDetails

40        Exit Sub

lblBarCode_Click_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmCoagDefinitions", "lblBarCode_Click", intEL, strES


End Sub

Private Sub lblPrint_LostFocus()

10        On Error GoTo lblPrint_LostFocus_Error

20        If lblPrint = "" Then lblPrint = "0"
30        SaveDetails

40        Exit Sub

lblPrint_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmCoagDefinitions", "lblPrint_LostFocus", intEL, strES


End Sub

Private Sub lDP_Click()

10        On Error GoTo lDP_Click_Error

20        lDP = Format$(Val(lDP) + 1)

30        If Val(lDP) > 3 Then lDP = "0"

40        SaveDetails

50        Exit Sub

lDP_Click_Error:

          Dim strES As String
          Dim intEL As Integer

60        intEL = Erl
70        strES = Err.Description
80        LogError "frmCoagDefinitions", "lDP_Click", intEL, strES


End Sub

Private Sub lstParameter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)


10        On Error GoTo lstParameter_MouseUp_Error

20        FillAges
30        FillDetails

40        Exit Sub

lstParameter_MouseUp_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmCoagDefinitions", "lstParameter_MouseUp", intEL, strES


End Sub

Private Sub odelta_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        On Error GoTo odelta_MouseUp_Error

20        SaveDetails

30        Exit Sub

odelta_MouseUp_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmCoagDefinitions", "odelta_MouseUp", intEL, strES


End Sub

Private Sub SaveDetails()

          Dim tb As New Recordset
          Dim sql As String
          Dim Y As Long



10        On Error GoTo SaveDetails_Error

20        grdAge.Col = 0
30        For Y = 1 To grdAge.Rows - 1
40            grdAge.Row = Y
50            If grdAge.CellBackColor = vbYellow Then

60                sql = "SELECT * from CoagTestDefinitions WHERE " & _
                        "TestName = '" & lstParameter & "' " & _
                        "and AgeFromDays = '" & FromDays(Y - 1) & "' " & _
                        "and AgeToDays = '" & ToDays(Y - 1) & "'"
70                Set tb = New Recordset
80                RecOpenClient 0, tb, sql
90                With tb
100                   If .EOF Then .AddNew

110                   !MaleLow = Val(tMaleLow)
120                   !MaleHigh = Val(tMaleHigh)
130                   !FemaleLow = Val(tFemaleLow)
140                   !FemaleHigh = Val(tFemaleHigh)
150                   !FlagMaleLow = Val(tFlagMaleLow)
160                   !FlagMaleHigh = Val(tFlagMaleHigh)
170                   !FlagFemaleLow = Val(tFlagFemaleLow)
180                   !FlagFemaleHigh = Val(tFlagFemaleHigh)
190                   !PlausibleLow = Val(tPlausibleLow)
200                   !PlausibleHigh = Val(tPlausibleHigh)
210                   !AgeFromDays = FromDays(Y - 1)
220                   !AgeToDays = ToDays(Y - 1)
230                   !CheckTime = IIf(txtCheckTime = "", 1, Val(txtCheckTime))

240                   .Update
250               End With

260               sql = "UPDATE CoagTestDefinitions " & _
                        "Set Code = '" & tCode & "', " & _
                        "DoDelta = " & IIf(oDelta = 1, 1, 0) & ", " & _
                        "DeltaLimit = '" & tdelta & "', " & _
                        "DP = '" & lDP & "', " & _
                        "Units = '" & cmbUnits & "', " & _
                        "Printable = " & IIf(cPrintable = 1, 1, 0) & ", " & _
                        "inuse = " & IIf(chkInuse = 1, 1, 0) & ", " & _
                        "BarCode = '" & Trim$(lblBarCode) & "', " & _
                        "PrintPriority = '" & Trim(lblPrint) & "', " & _
                        "CheckTime = '" & IIf(txtCheckTime = "", 1, Val(txtCheckTime)) & "' " & _
                        "WHERE code = '" & tCode & "' "
270               Cnxn(0).Execute sql

280               Exit For
290           End If
300       Next






310       Exit Sub

SaveDetails_Error:

          Dim strES As String
          Dim intEL As Integer

320       intEL = Erl
330       strES = Err.Description
340       LogError "frmCoagDefinitions", "SaveDetails", intEL, strES, sql


End Sub

Private Sub tCode_KeyUp(KeyCode As Integer, Shift As Integer)

10        On Error GoTo tCode_KeyUp_Error

20        SaveDetails

30        Exit Sub

tCode_KeyUp_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmCoagDefinitions", "tCode_KeyUp", intEL, strES


End Sub

Private Sub tDelta_KeyUp(KeyCode As Integer, Shift As Integer)

10        On Error GoTo tDelta_KeyUp_Error

20        SaveDetails

30        Exit Sub

tDelta_KeyUp_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmCoagDefinitions", "tDelta_KeyUp", intEL, strES


End Sub





Private Sub tFemaleHigh_KeyUp(KeyCode As Integer, Shift As Integer)

10        On Error GoTo tFemaleHigh_KeyUp_Error

20        SaveDetails

30        Exit Sub

tFemaleHigh_KeyUp_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmCoagDefinitions", "tFemaleHigh_KeyUp", intEL, strES


End Sub

Private Sub tFemaleLow_KeyUp(KeyCode As Integer, Shift As Integer)

10        On Error GoTo tFemaleLow_KeyUp_Error

20        SaveDetails

30        Exit Sub

tFemaleLow_KeyUp_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmCoagDefinitions", "tFemaleLow_KeyUp", intEL, strES


End Sub

Private Sub tFlagFemaleHigh_KeyUp(KeyCode As Integer, Shift As Integer)

10        On Error GoTo tFlagFemaleHigh_KeyUp_Error

20        SaveDetails

30        Exit Sub

tFlagFemaleHigh_KeyUp_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmCoagDefinitions", "tFlagFemaleHigh_KeyUp", intEL, strES


End Sub

Private Sub tFlagFemaleLow_KeyUp(KeyCode As Integer, Shift As Integer)

10        On Error GoTo tFlagFemaleLow_KeyUp_Error

20        SaveDetails

30        Exit Sub

tFlagFemaleLow_KeyUp_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmCoagDefinitions", "tFlagFemaleLow_KeyUp", intEL, strES


End Sub

Private Sub tFlagMaleHigh_KeyUp(KeyCode As Integer, Shift As Integer)

10        On Error GoTo tFlagMaleHigh_KeyUp_Error

20        SaveDetails

30        Exit Sub

tFlagMaleHigh_KeyUp_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmCoagDefinitions", "tFlagMaleHigh_KeyUp", intEL, strES


End Sub

Private Sub tFlagMaleLow_KeyUp(KeyCode As Integer, Shift As Integer)

10        On Error GoTo tFlagMaleLow_KeyUp_Error

20        SaveDetails

30        Exit Sub

tFlagMaleLow_KeyUp_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmCoagDefinitions", "tFlagMaleLow_KeyUp", intEL, strES


End Sub

Private Sub tMaleHigh_KeyUp(KeyCode As Integer, Shift As Integer)

10        On Error GoTo tMaleHigh_KeyUp_Error

20        SaveDetails

30        Exit Sub

tMaleHigh_KeyUp_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmCoagDefinitions", "tMaleHigh_KeyUp", intEL, strES


End Sub

Private Sub tMaleLow_KeyUp(KeyCode As Integer, Shift As Integer)

10        On Error GoTo tMaleLow_KeyUp_Error

20        SaveDetails

30        Exit Sub

tMaleLow_KeyUp_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmCoagDefinitions", "tMaleLow_KeyUp", intEL, strES


End Sub

Private Sub tPlausibleHigh_KeyUp(KeyCode As Integer, Shift As Integer)

10        On Error GoTo tPlausibleHigh_KeyUp_Error

20        SaveDetails

30        Exit Sub

tPlausibleHigh_KeyUp_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmCoagDefinitions", "tPlausibleHigh_KeyUp", intEL, strES


End Sub

Private Sub tPlausibleLow_KeyUp(KeyCode As Integer, Shift As Integer)

10        On Error GoTo tPlausibleLow_KeyUp_Error

20        SaveDetails

30        Exit Sub

tPlausibleLow_KeyUp_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmCoagDefinitions", "tPlausibleLow_KeyUp", intEL, strES


End Sub



Private Sub txtCheckTime_KeyPress(KeyAscii As Integer)
10        KeyAscii = VI(KeyAscii, Numeric_Only)
End Sub

Private Sub txtCheckTime_KeyUp(KeyCode As Integer, Shift As Integer)
10        SaveDetails
End Sub

Private Sub upPP_Change()

10        On Error GoTo upPP_Change_Error

20        SaveDetails

30        Exit Sub

upPP_Change_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmCoagDefinitions", "upPP_Change", intEL, strES


End Sub
