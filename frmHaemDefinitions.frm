VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmHaemDefinitions 
   Caption         =   "NetAcquire - Haematology Definitions"
   ClientHeight    =   6825
   ClientLeft      =   540
   ClientTop       =   1995
   ClientWidth     =   8100
   Icon            =   "frmHaemDefinitions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   8100
   Begin VB.Frame Frame7 
      Caption         =   "Specifics (Applies to all age ranges)"
      Height          =   1245
      Left            =   1350
      TabIndex        =   9
      Top             =   5475
      Width           =   4635
      Begin VB.CheckBox chkWard 
         Caption         =   "View on Ward"
         Height          =   240
         Left            =   2295
         TabIndex        =   49
         Top             =   765
         Width           =   1410
      End
      Begin VB.TextBox tTestName 
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   420
         Width           =   2925
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Test Name"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   8
         Top             =   480
         Width           =   780
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Decimal Points"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   6
         Top             =   780
         Width           =   1050
      End
      Begin VB.Label lDP 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         Height          =   285
         Left            =   1200
         TabIndex        =   7
         Top             =   750
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Age Ranges"
      Height          =   2115
      Left            =   1350
      TabIndex        =   2
      Top             =   75
      Width           =   6615
      Begin VB.CommandButton bAmendAgeRange 
         Caption         =   "Amend Age Range"
         Height          =   1110
         Left            =   4890
         Picture         =   "frmHaemDefinitions.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   330
         Width           =   1155
      End
      Begin MSFlexGridLib.MSFlexGrid g 
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
   Begin VB.ListBox lstParameter 
      Height          =   5235
      IntegralHeight  =   0   'False
      Left            =   90
      TabIndex        =   1
      Top             =   165
      Width           =   1155
   End
   Begin VB.CommandButton bCancel 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   930
      Left            =   6675
      Picture         =   "frmHaemDefinitions.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5745
      Width           =   1290
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2985
      Left            =   1380
      TabIndex        =   10
      Top             =   2295
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   5265
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
      TabPicture(0)   =   "frmHaemDefinitions.frx":091E
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "tMaleHigh"
      Tab(0).Control(1)=   "tFemaleHigh"
      Tab(0).Control(2)=   "tMaleLow"
      Tab(0).Control(3)=   "tFemaleLow"
      Tab(0).Control(4)=   "Label1(1)"
      Tab(0).Control(5)=   "Label2"
      Tab(0).Control(6)=   "Label3"
      Tab(0).Control(7)=   "Label4"
      Tab(0).Control(8)=   "Label6"
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Flag Range"
      TabPicture(1)   =   "frmHaemDefinitions.frx":093A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "tFlagMaleLow"
      Tab(1).Control(1)=   "tFlagFemaleHigh"
      Tab(1).Control(2)=   "tFlagMaleHigh"
      Tab(1).Control(3)=   "tFlagFemaleLow"
      Tab(1).Control(4)=   "Label15(1)"
      Tab(1).Control(5)=   "Label14(2)"
      Tab(1).Control(6)=   "Label13(1)"
      Tab(1).Control(7)=   "Label12(2)"
      Tab(1).Control(8)=   "Label7(1)"
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "Plausible Range"
      TabPicture(2)   =   "frmHaemDefinitions.frx":0956
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "tPlausibleHigh"
      Tab(2).Control(1)=   "tPlausibleLow"
      Tab(2).Control(2)=   "Label8(1)"
      Tab(2).Control(3)=   "Label9(1)"
      Tab(2).Control(4)=   "Label10(2)"
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "Delta Check"
      TabPicture(3)   =   "frmHaemDefinitions.frx":0972
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Label20"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Label5"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Label16"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "tDelta"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "oDelta"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "txtCheckTime"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).ControlCount=   6
      Begin VB.TextBox txtCheckTime 
         Height          =   285
         Left            =   2790
         MaxLength       =   5
         TabIndex        =   50
         Top             =   1860
         Width           =   555
      End
      Begin VB.TextBox tFlagMaleLow 
         Height          =   315
         Left            =   -71460
         TabIndex        =   26
         Top             =   1770
         Width           =   915
      End
      Begin VB.TextBox tFlagFemaleHigh 
         Height          =   315
         Left            =   -72990
         TabIndex        =   25
         Top             =   1170
         Width           =   915
      End
      Begin VB.TextBox tFlagMaleHigh 
         Height          =   315
         Left            =   -71460
         TabIndex        =   24
         Top             =   1170
         Width           =   915
      End
      Begin VB.CheckBox oDelta 
         Alignment       =   1  'Right Justify
         Caption         =   "Enabled"
         Height          =   195
         Left            =   2430
         TabIndex        =   23
         Top             =   1140
         Width           =   915
      End
      Begin VB.TextBox tDelta 
         Height          =   285
         Left            =   2790
         MaxLength       =   5
         TabIndex        =   22
         Top             =   1530
         Width           =   795
      End
      Begin VB.TextBox tPlausibleHigh 
         Height          =   285
         Left            =   -72600
         TabIndex        =   21
         Top             =   900
         Width           =   1215
      End
      Begin VB.TextBox tPlausibleLow 
         Height          =   285
         Left            =   -72600
         TabIndex        =   20
         Top             =   1470
         Width           =   1215
      End
      Begin VB.TextBox tMaleHigh 
         Height          =   315
         Left            =   -71640
         MaxLength       =   5
         TabIndex        =   19
         Top             =   1200
         Width           =   915
      End
      Begin VB.TextBox tFemaleHigh 
         Height          =   315
         Left            =   -73170
         MaxLength       =   5
         TabIndex        =   18
         Top             =   1200
         Width           =   915
      End
      Begin VB.TextBox tMaleLow 
         Height          =   315
         Left            =   -71640
         MaxLength       =   5
         TabIndex        =   17
         Top             =   1800
         Width           =   915
      End
      Begin VB.TextBox tFemaleLow 
         Height          =   315
         Left            =   -73170
         MaxLength       =   5
         TabIndex        =   16
         Top             =   1800
         Width           =   915
      End
      Begin VB.TextBox tfr 
         Height          =   315
         Index           =   3
         Left            =   -73170
         TabIndex        =   15
         Top             =   1800
         Width           =   915
      End
      Begin VB.TextBox tfr 
         Height          =   315
         Index           =   1
         Left            =   -71640
         TabIndex        =   14
         Top             =   1800
         Width           =   915
      End
      Begin VB.TextBox tfr 
         Height          =   315
         Index           =   2
         Left            =   -73170
         TabIndex        =   13
         Top             =   1200
         Width           =   915
      End
      Begin VB.TextBox tfr 
         Height          =   315
         Index           =   0
         Left            =   -71640
         TabIndex        =   12
         Top             =   1200
         Width           =   915
      End
      Begin VB.TextBox tFlagFemaleLow 
         Height          =   315
         Left            =   -72990
         TabIndex        =   11
         Top             =   1770
         Width           =   915
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "day(s)"
         Height          =   195
         Left            =   3360
         TabIndex        =   52
         Top             =   1890
         Width           =   420
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Check Time"
         Height          =   195
         Left            =   1830
         TabIndex        =   51
         Top             =   1890
         Width           =   855
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Low"
         Height          =   195
         Index           =   1
         Left            =   -73740
         TabIndex        =   48
         Top             =   1830
         Width           =   300
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "High"
         Height          =   195
         Index           =   2
         Left            =   -73770
         TabIndex        =   47
         Top             =   1260
         Width           =   330
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Male"
         Height          =   195
         Index           =   1
         Left            =   -71190
         TabIndex        =   46
         Top             =   870
         Width           =   345
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Female"
         Height          =   195
         Index           =   2
         Left            =   -72840
         TabIndex        =   45
         Top             =   900
         Width           =   510
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Result values outside this range will be flagged as High or Low"
         Height          =   195
         Index           =   1
         Left            =   -74010
         TabIndex        =   44
         Top             =   2490
         Width           =   4410
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Value"
         Height          =   195
         Left            =   2280
         TabIndex        =   43
         Top             =   1560
         Width           =   405
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "High"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   1
         Left            =   -73110
         TabIndex        =   42
         Top             =   930
         Width           =   330
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Low"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   1
         Left            =   -73080
         TabIndex        =   41
         Top             =   1500
         Width           =   300
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Results outside this range will be marked as implausible."
         Height          =   195
         Index           =   2
         Left            =   -73740
         TabIndex        =   40
         Top             =   2340
         Width           =   3930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Male"
         Height          =   195
         Index           =   1
         Left            =   -71370
         TabIndex        =   39
         Top             =   900
         Width           =   345
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Female"
         Height          =   195
         Left            =   -73020
         TabIndex        =   38
         Top             =   930
         Width           =   510
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "High"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   -73740
         TabIndex        =   37
         Top             =   1260
         Width           =   330
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Low"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   -73710
         TabIndex        =   36
         Top             =   1860
         Width           =   300
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Low"
         Height          =   195
         Index           =   0
         Left            =   -73920
         TabIndex        =   35
         Top             =   1860
         Width           =   300
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "High"
         Height          =   195
         Index           =   1
         Left            =   -73950
         TabIndex        =   34
         Top             =   1290
         Width           =   330
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Male"
         Height          =   195
         Index           =   0
         Left            =   -71370
         TabIndex        =   33
         Top             =   900
         Width           =   345
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Female"
         Height          =   195
         Index           =   1
         Left            =   -73020
         TabIndex        =   32
         Top             =   930
         Width           =   510
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Do not print the result if the sample is :-"
         Height          =   195
         Index           =   1
         Left            =   -74070
         TabIndex        =   31
         Top             =   750
         Width           =   2730
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "These are the normal range values printed on the report forms."
         Height          =   195
         Left            =   -74160
         TabIndex        =   30
         Top             =   2520
         Width           =   4395
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Result values outside this range will be flagged as High or Low"
         Height          =   195
         Index           =   0
         Left            =   -74190
         TabIndex        =   29
         Top             =   2520
         Width           =   4410
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "High"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   0
         Left            =   -73110
         TabIndex        =   28
         Top             =   1020
         Width           =   330
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Low"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   0
         Left            =   -73080
         TabIndex        =   27
         Top             =   1590
         Width           =   300
      End
   End
End
Attribute VB_Name = "frmHaemDefinitions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Added 15/Jul/2004
'Save Details


Private FromDays() As Long
Private ToDays() As Long

Private Sub FillAges()

          Dim tb As New Recordset
          Dim s As String
          Dim n As Long
          Dim sql As String


10        On Error GoTo FillAges_Error

20        g.Rows = 2
30        g.AddItem ""
40        g.RemoveItem 1

50        ReDim FromDays(0 To 0)
60        ReDim ToDays(0 To 0)

70        sql = "SELECT * from HaemTestDefinitions WHERE " & _
                "AnalyteName = '" & lstParameter & "' " & _
                "and Hospital = '" & HospName(0) & "' " & _
                "Order by cast(AgetoDays as numeric) asc"
80        Set tb = New Recordset
90        RecOpenClient 0, tb, sql
100       If tb.EOF Then Exit Sub

110       ReDim FromDays(0 To tb.RecordCount - 1)
120       ReDim ToDays(0 To tb.RecordCount - 1)
130       n = 0
140       Do While Not tb.EOF
150           If Trim(tb!AgeFromDays) & "" <> "" Then FromDays(n) = tb!AgeFromDays & ""
160           If Trim(tb!AgeToDays) & "" <> "" Then ToDays(n) = tb!AgeToDays & ""
170           s = Format$(n) & vbTab & _
                  dmyFromCount(FromDays(n)) & vbTab & _
                  dmyFromCount(ToDays(n))
180           g.AddItem s
190           n = n + 1
200           tb.MoveNext
210       Loop

220       If g.Rows > 2 Then
230           g.RemoveItem 1
240       End If

250       g.Col = 0
260       g.Row = 1
270       g.CellBackColor = vbYellow
280       g.CellForeColor = vbBlue



290       Exit Sub

FillAges_Error:

          Dim strES As String
          Dim intEL As Integer



300       intEL = Erl
310       strES = Err.Description
320       LogError "frmHaemDefinitions", "FillAges", intEL, strES, sql


End Sub

Private Sub FillParameters()

          Dim tb As New Recordset
          Dim sql As String

10        On Error GoTo FillParameters_Error

20        lstParameter.Clear
30        sql = "SELECT distinct AnalyteName from HaemTestDefinitions " & _
                "WHERE Hospital = '" & HospName(0) & "' " & _
                "order by AnalyteName asc"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql

60        Do While Not tb.EOF
70            lstParameter.AddItem tb!AnalyteName & ""
80            tb.MoveNext
90        Loop

100       If lstParameter.ListCount > 0 Then
110           lstParameter.Selected(0) = True
120       End If

130       Exit Sub

FillParameters_Error:

          Dim strES As String
          Dim intEL As Integer



140       intEL = Erl
150       strES = Err.Description
160       LogError "frmHaemDefinitions", "FillParameters", intEL, strES, sql


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
40        g.Col = 0
50        For Y = 1 To g.Rows - 1
60            g.Row = Y
70            If g.CellBackColor = vbYellow Then
80                AgeNumber = Y - 1
90                Exit For
100           End If
110       Next
120       If AgeNumber = -1 Then
130           iMsg "SELECT Age Range", vbCritical
140           Exit Sub
150       End If

160       tTestName = ""
170       oDelta = 0
180       tdelta = ""
190       lDP = "0"
200       tPlausibleLow = ""
210       tPlausibleHigh = ""
220       tMaleHigh = ""
230       tFemaleHigh = ""
240       tMaleLow = ""
250       tFemaleLow = ""
260       txtCheckTime = ""

270       sql = "SELECT * from HaemTestDefinitions WHERE " & _
                "AnalyteName = '" & lstParameter & "' " & _
                "and AgeFromDays = '" & FromDays(AgeNumber) & "' " & _
                "and AgeToDays = '" & ToDays(AgeNumber) & "' " & _
                "And Hospital = '" & HospName(0) & "'"
280       Set tb = New Recordset
290       RecOpenServer 0, tb, sql
300       If Not tb.EOF Then
310           tTestName = lstParameter
320           oDelta = IIf(tb!DoDelta, 1, 0)
330           tdelta = tb!DeltaValue & ""
340           If Not IsNull(tb!Printformat) Then
350               lDP = tb!Printformat
360           Else
370               lDP = "1"
380           End If
390           tPlausibleLow = IIf(IsNull(tb!PlausibleLow), "0", tb!PlausibleLow)
400           tPlausibleHigh = IIf(IsNull(tb!PlausibleHigh), "9999", tb!PlausibleHigh)
410           tMaleHigh = tb!MaleHigh
420           tFemaleHigh = tb!FemaleHigh
430           tMaleLow = tb!MaleLow
440           tFemaleLow = tb!FemaleLow
450           If tb!vward = True Then chkWard.Value = 1 Else chkWard = 0
460           If IsNull(tb!CheckTime) Then
470               txtCheckTime = ""
480           Else
490               txtCheckTime = tb!CheckTime
500           End If
510       End If



520       Exit Sub

FillDetails_Error:

          Dim strES As String
          Dim intEL As Integer



530       intEL = Erl
540       strES = Err.Description
550       LogError "frmHaemDefinitions", "FillDetails", intEL, strES, sql


End Sub


Private Sub SaveDetails()

          Dim tb As New Recordset
          Dim sql As String
          Dim Y As Long

10        On Error GoTo SaveDetails_Error

20        g.Col = 0
30        For Y = 1 To g.Rows - 1
40            g.Row = Y
50            If g.CellBackColor = vbYellow Then

60                sql = "SELECT * from HaemTestDefinitions WHERE " & _
                        "AnalyteName = '" & lstParameter & "' " & _
                        "and AgeFromDays = '" & FromDays(Y - 1) & "' " & _
                        "and AgeToDays = '" & ToDays(Y - 1) & "'"
70                Set tb = New Recordset
80                RecOpenClient 0, tb, sql
90                With tb

100                   If .EOF Then .AddNew
110                   !AnalyteName = lstParameter
120                   !DoDelta = oDelta = 1
130                   !DeltaValue = Val(tdelta)
140                   !Printformat = lDP
150                   !MaleLow = Val(tMaleLow)
160                   !MaleHigh = Val(tMaleHigh)
170                   !FemaleLow = Val(tFemaleLow)
180                   !FemaleHigh = Val(tFemaleHigh)
190                   !PlausibleLow = Val(tPlausibleLow)
200                   !PlausibleHigh = Val(tPlausibleHigh)
210                   !AgeFromDays = FromDays(Y - 1)
220                   !AgeToDays = ToDays(Y - 1)
230                   !vward = chkWard.Value
240                   !Hospital = HospName(0)
250                   !CheckTime = IIf(txtCheckTime = "", 1, Val(txtCheckTime))
260                   .Update
270               End With
280               Exit For
290           End If
300       Next

          'colHaemTestDefinitions.Refresh

          'Added 15/Jul/2004

310       sql = "INSERT into UPDATEs  (upd, dtime) values ('HaemTest', '" & Format(Now, "dd/MMM/yyyy hh:mm") & "')"
320       Cnxn(0).Execute sql

330       Exit Sub

SaveDetails_Error:

          Dim strES As String
          Dim intEL As Integer



340       intEL = Erl
350       strES = Err.Description
360       LogError "frmHaemDefinitions", "SaveDetails", intEL, strES, sql


End Sub






Private Sub bAmendAgeRange_Click()

10        On Error GoTo bAmendAgeRange_Click_Error

20        If lstParameter = "" Then
30            iMsg "SELECT Parameter", vbCritical
40            Exit Sub
50        End If

60        With frmAges
70            .Analyte = lstParameter
80            .SampleType = "Haematology"
90            .Discipline = "Haematology"
100           .Show 1
110       End With

120       FillAges

130       Exit Sub

bAmendAgeRange_Click_Error:

          Dim strES As String
          Dim intEL As Integer



140       intEL = Erl
150       strES = Err.Description
160       LogError "frmHaemDefinitions", "bAmendAgeRange_Click", intEL, strES


End Sub

Private Sub bcancel_Click()

10        Unload Me

End Sub




Private Sub chkWard_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        On Error GoTo chkWard_MouseUp_Error

20        SaveDetails

30        Exit Sub

chkWard_MouseUp_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmHaemDefinitions", "chkWard_MouseUp", intEL, strES


End Sub

Private Sub Form_Load()

10        On Error GoTo Form_Load_Error

20        SSTab1.TabVisible(1) = False

30        FillParameters
40        FillAges
50        FillDetails

60        Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer



70        intEL = Erl
80        strES = Err.Description
90        LogError "frmHaemDefinitions", "Form_Load", intEL, strES


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
170       LogError "frmHaemDefinitions", "g_MouseUp", intEL, strES


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
80        LogError "frmHaemDefinitions", "lDP_Click", intEL, strES


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
70        LogError "frmHaemDefinitions", "lstParameter_MouseUp", intEL, strES


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
60        LogError "frmHaemDefinitions", "odelta_MouseUp", intEL, strES


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
60        LogError "frmHaemDefinitions", "tDelta_KeyUp", intEL, strES


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
60        LogError "frmHaemDefinitions", "tFemaleHigh_KeyUp", intEL, strES


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
60        LogError "frmHaemDefinitions", "tFemaleLow_KeyUp", intEL, strES


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
60        LogError "frmHaemDefinitions", "tFlagFemaleHigh_KeyUp", intEL, strES


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
60        LogError "frmHaemDefinitions", "tFlagFemaleLow_KeyUp", intEL, strES


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
60        LogError "frmHaemDefinitions", "tFlagMaleHigh_KeyUp", intEL, strES


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
60        LogError "frmHaemDefinitions", "tFlagMaleLow_KeyUp", intEL, strES


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
60        LogError "frmHaemDefinitions", "tMaleHigh_KeyUp", intEL, strES


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
60        LogError "frmHaemDefinitions", "tMaleLow_KeyUp", intEL, strES


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
60        LogError "frmHaemDefinitions", "tPlausibleHigh_KeyUp", intEL, strES


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
60        LogError "frmHaemDefinitions", "tPlausibleLow_KeyUp", intEL, strES


End Sub



Private Sub txtCheckTime_KeyPress(KeyAscii As Integer)
10        KeyAscii = VI(KeyAscii, Numeric_Only)
End Sub

Private Sub txtCheckTime_KeyUp(KeyCode As Integer, Shift As Integer)
10        SaveDetails
End Sub
