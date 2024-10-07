VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmBioRanges 
   Caption         =   "NetAcquire - Test Definitions"
   ClientHeight    =   8025
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12555
   LinkTopic       =   "Form1"
   ScaleHeight     =   8025
   ScaleWidth      =   12555
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAddAgeSpecific 
      Caption         =   "&Add Parameter Age Specific Range"
      Height          =   1035
      HelpContextID   =   10100
      Left            =   10635
      Picture         =   "frmBioRanges.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   3810
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.CommandButton cmdXL 
      Caption         =   "Export to Excel"
      Height          =   615
      Left            =   10635
      Picture         =   "frmBioRanges.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   5610
      Width           =   1485
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   525
      HelpContextID   =   10026
      Left            =   10635
      Picture         =   "frmBioRanges.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   7320
      Width           =   1485
   End
   Begin VB.CommandButton bPrint 
      Caption         =   "&Print"
      Height          =   585
      HelpContextID   =   10130
      Left            =   10635
      Picture         =   "frmBioRanges.frx":0DB6
      Style           =   1  'Graphical
      TabIndex        =   19
      Tag             =   "bprint"
      Top             =   6660
      Width           =   1485
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7665
      Left            =   270
      TabIndex        =   13
      Top             =   90
      Width           =   10125
      _ExtentX        =   17859
      _ExtentY        =   13520
      _Version        =   393216
      Tabs            =   5
      Tab             =   3
      TabsPerRow      =   5
      TabHeight       =   794
      TabMaxWidth     =   3528
      TabCaption(0)   =   "Codes, Units and Precision"
      TabPicture(0)   =   "frmBioRanges.frx":1420
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "gCodes"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Normal and Flag Ranges"
      TabPicture(1)   =   "frmBioRanges.frx":143C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "gNormal"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Plausible, Auto-Val and Delta Ranges"
      TabPicture(2)   =   "frmBioRanges.frx":1458
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "gPlausible"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Masks"
      TabPicture(3)   =   "frmBioRanges.frx":1474
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Label4"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "gMasks"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "Print Sequence"
      TabPicture(4)   =   "frmBioRanges.frx":1490
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "bMoveUp"
      Tab(4).Control(1)=   "bMoveDown"
      Tab(4).Control(2)=   "tmrUp"
      Tab(4).Control(3)=   "tmrDown"
      Tab(4).Control(4)=   "gSequence"
      Tab(4).ControlCount=   5
      Begin VB.CommandButton bMoveUp 
         Caption         =   "Move &Up"
         Height          =   1035
         Left            =   -68550
         Picture         =   "frmBioRanges.frx":14AC
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   1260
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton bMoveDown 
         Caption         =   "Move &Down"
         Height          =   1035
         Left            =   -68550
         Picture         =   "frmBioRanges.frx":18EE
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   2310
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Timer tmrUp 
         Enabled         =   0   'False
         Interval        =   250
         Left            =   -67920
         Top             =   1620
      End
      Begin VB.Timer tmrDown 
         Enabled         =   0   'False
         Interval        =   250
         Left            =   -67920
         Top             =   2610
      End
      Begin MSFlexGridLib.MSFlexGrid gCodes 
         Height          =   6405
         HelpContextID   =   10090
         Left            =   -74850
         TabIndex        =   14
         Top             =   1080
         Width           =   8925
         _ExtentX        =   15743
         _ExtentY        =   11298
         _Version        =   393216
         Cols            =   13
         FixedCols       =   3
         BackColor       =   -2147483624
         ForeColor       =   -2147483635
         BackColorFixed  =   -2147483647
         ForeColorFixed  =   -2147483624
         WordWrap        =   -1  'True
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   0
         GridLines       =   3
         GridLinesFixed  =   3
         AllowUserResizing=   1
         FormatString    =   $"frmBioRanges.frx":1D30
      End
      Begin MSFlexGridLib.MSFlexGrid gNormal 
         Height          =   6375
         HelpContextID   =   10090
         Left            =   -74850
         TabIndex        =   15
         Top             =   1110
         Width           =   8925
         _ExtentX        =   15743
         _ExtentY        =   11245
         _Version        =   393216
         Cols            =   19
         FixedCols       =   4
         BackColor       =   -2147483624
         ForeColor       =   -2147483635
         BackColorFixed  =   -2147483647
         ForeColorFixed  =   -2147483624
         WordWrap        =   -1  'True
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   0
         GridLines       =   3
         GridLinesFixed  =   3
         FormatString    =   $"frmBioRanges.frx":1DCF
      End
      Begin MSFlexGridLib.MSFlexGrid gPlausible 
         Height          =   6225
         HelpContextID   =   10090
         Left            =   -74850
         TabIndex        =   16
         Top             =   1260
         Width           =   8925
         _ExtentX        =   15743
         _ExtentY        =   10980
         _Version        =   393216
         Cols            =   10
         FixedCols       =   3
         BackColor       =   -2147483624
         ForeColor       =   -2147483635
         BackColorFixed  =   -2147483647
         ForeColorFixed  =   -2147483624
         WordWrap        =   -1  'True
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   0
         GridLines       =   3
         GridLinesFixed  =   3
         AllowUserResizing=   1
         FormatString    =   $"frmBioRanges.frx":1EF1
      End
      Begin MSFlexGridLib.MSFlexGrid gMasks 
         Height          =   5865
         HelpContextID   =   10090
         Left            =   780
         TabIndex        =   17
         Top             =   1530
         Width           =   5925
         _ExtentX        =   10451
         _ExtentY        =   10345
         _Version        =   393216
         Cols            =   7
         FixedCols       =   2
         BackColor       =   -2147483624
         ForeColor       =   -2147483635
         BackColorFixed  =   -2147483647
         ForeColorFixed  =   -2147483624
         WordWrap        =   -1  'True
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   0
         GridLines       =   3
         GridLinesFixed  =   3
         AllowUserResizing=   1
         FormatString    =   "<Long Name                 |<Short Name      |<Code   |^Old    |^  Lipaemic  |^  Icteric  |^  Haemolysed  "
      End
      Begin MSFlexGridLib.MSFlexGrid gSequence 
         Height          =   6045
         Left            =   -74520
         TabIndex        =   26
         Top             =   1080
         Width           =   5925
         _ExtentX        =   10451
         _ExtentY        =   10663
         _Version        =   393216
         FixedCols       =   0
         BackColor       =   -2147483624
         ForeColor       =   -2147483635
         BackColorFixed  =   -2147483647
         ForeColorFixed  =   -2147483624
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   0
         GridLines       =   3
         GridLinesFixed  =   3
         ScrollBars      =   2
         FormatString    =   "<Test Name                  |<Long Name                                        "
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
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H80000001&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Do not Print Result if >="
         ForeColor       =   &H80000018&
         Height          =   285
         Left            =   2520
         TabIndex        =   18
         Top             =   1230
         Width           =   3270
      End
   End
   Begin VB.ComboBox cmbSampleType 
      Height          =   315
      HelpContextID   =   10010
      Left            =   10635
      TabIndex        =   9
      Text            =   "cmbSampleType"
      Top             =   2340
      Width           =   1485
   End
   Begin VB.ComboBox cmbCategory 
      Height          =   315
      HelpContextID   =   10020
      Left            =   10635
      TabIndex        =   8
      Text            =   "cmbCategory"
      Top             =   2880
      Width           =   1485
   End
   Begin VB.ComboBox cmbHospital 
      Height          =   315
      HelpContextID   =   10030
      Left            =   10635
      TabIndex        =   7
      Text            =   "cmbHospital"
      Top             =   3420
      Width           =   1485
   End
   Begin VB.Frame Frame2 
      Caption         =   "Names"
      Height          =   885
      Left            =   10635
      TabIndex        =   4
      Top             =   180
      Width           =   1485
      Begin VB.OptionButton optLongOrShort 
         Caption         =   "Short Names"
         Height          =   255
         Index           =   1
         Left            =   45
         TabIndex        =   6
         Top             =   495
         Width           =   1245
      End
      Begin VB.OptionButton optLongOrShort 
         Caption         =   "Long Names"
         Height          =   255
         Index           =   0
         Left            =   60
         TabIndex        =   5
         Top             =   240
         Value           =   -1  'True
         Width           =   1245
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Discipline"
      Height          =   885
      Left            =   10635
      TabIndex        =   0
      Top             =   1110
      Width           =   1485
      Begin VB.OptionButton optDiscipline 
         Caption         =   "Biochemistry"
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   3
         Top             =   210
         Value           =   -1  'True
         Width           =   1305
      End
      Begin VB.OptionButton optDiscipline 
         Caption         =   "Immunology"
         Height          =   195
         Index           =   1
         Left            =   60
         TabIndex        =   2
         Top             =   420
         Width           =   1275
      End
      Begin VB.OptionButton optDiscipline 
         Caption         =   "Endocrinology"
         Height          =   195
         Index           =   2
         Left            =   60
         TabIndex        =   1
         Top             =   630
         Width           =   1305
      End
   End
   Begin VB.Image imgSquarePlus 
      Height          =   225
      Left            =   12285
      Picture         =   "frmBioRanges.frx":1F97
      Top             =   180
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image imgSquareMinus 
      Height          =   225
      Left            =   12555
      Picture         =   "frmBioRanges.frx":2091
      Top             =   180
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image imgSquareTick 
      Height          =   225
      Left            =   12525
      Picture         =   "frmBioRanges.frx":218B
      Top             =   780
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image imgSquareCross 
      Height          =   225
      Left            =   12525
      Picture         =   "frmBioRanges.frx":2461
      Top             =   480
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label lblExcelInfo 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exporting..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   10755
      TabIndex        =   22
      Top             =   6240
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Sample Type"
      Height          =   195
      Left            =   10575
      TabIndex        =   12
      Top             =   2115
      Width           =   930
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Category"
      Height          =   195
      Left            =   10620
      TabIndex        =   11
      Top             =   2700
      Width           =   630
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Hospital"
      Height          =   195
      Left            =   10620
      TabIndex        =   10
      Top             =   3240
      Width           =   570
   End
End
Attribute VB_Name = "frmBioRanges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private ShowAllAgeRanges As String

Private strUnits() As String


Private FireCounter As Integer

Private pViewTab As Integer

Private Sub SaveSequence()

          Dim sql As String
          Dim Y As Integer
          Dim Discipline As String

10        On Error GoTo SaveSequence_Error

20        Discipline = GetDiscipline()

30        For Y = 1 To gSequence.Rows - 1
40            sql = "UPDATE " & Discipline & "TestDefinitions " & _
                    "SET PrintPriority = '" & Y & "' " & _
                    "WHERE ShortName = '" & gSequence.TextMatrix(Y, 0) & "'"
50            Cnxn(0).Execute sql
60        Next

70        Exit Sub

SaveSequence_Error:

          Dim strES As String
          Dim intEL As Integer

80        intEL = Erl
90        strES = Err.Description
100       LogError "frmBioRanges", "SaveSequence", intEL, strES, sql


End Sub
Private Sub FillStrUnits()

          Dim sql As String
          Dim tb As Recordset
          Dim intN As Integer

10        On Error GoTo FillStrUnits_Error

20        sql = "Select Text from Lists where " & _
                "ListType = 'UN' " & _
                "order by ListOrder"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql

50        ReDim strUnits(0 To 0)
60        strUnits(0) = ""
70        intN = 1

80        Do While Not tb.EOF

90            ReDim Preserve strUnits(0 To intN) As String
100           strUnits(intN) = tb!Text & ""

110           intN = intN + 1

120           tb.MoveNext

130       Loop

140       Exit Sub

FillStrUnits_Error:

          Dim strES As String
          Dim intEL As Integer

150       intEL = Erl
160       strES = Err.Description
170       LogError "frmBioRanges", "FillStrUnits", intEL, strES, sql


End Sub

Private Sub ClearGrid()

          Dim grd As MSFlexGrid

10        On Error GoTo ClearGrid_Error

20        Set grd = GetActiveGrid

30        grd.Visible = False
40        grd.Rows = 2
50        grd.AddItem ""
60        grd.RemoveItem 1

70        Exit Sub

ClearGrid_Error:

          Dim strES As String
          Dim intEL As Integer

80        intEL = Erl
90        strES = Err.Description
100       LogError "frmBioRanges", "ClearGrid", intEL, strES


End Sub

Private Sub FillgNormal(Optional ByVal RowTop As Long)

          Dim sql As String
          Dim tb As Recordset
          Dim strS As String
          Dim intN As Integer
          Dim blnFound As Boolean
          Dim blnShowThis As Boolean
          Dim strSampleType As String
          Dim Discipline As String

10        On Error GoTo FillgNormal_Error

20        If cmbHospital = "" Then
30            cmbHospital = HospName(0)
40        End If
          'If cmbCategory = "" Then
          '  cmbCategory = "Human"
          'End If
50        If cmbSampleType = "" Then
60            cmbSampleType = "Serum"
70        End If
80        strSampleType = ListCodeFor("ST", cmbSampleType)

90        Discipline = GetDiscipline()

          ' "<Long Name             |<Short Name   |^Age From |^Age To   "               '0 to 3
          ' "|^Normal Male Low|^Normal Male High|^Normal Female Low|^Normal Female High" '4 to 7
          ' "|^Flag Male Low|^Flag Male High|^Flag Female Low|^Flag Female High"         '8 to 11
          ' "|<Code    |^Dec.Pl  "       '12 to 13
          ' "|<AgeFromDays|<AgeToDays|^Print Priority"                                   '14 to 16

100       ClearGrid

110       sql = "Select * from " & Discipline & "TestDefinitions where " & _
                "SampleType = '" & strSampleType & "' " & _
                "and Category = '" & cmbCategory & "' " & _
                "and Hospital = '" & cmbHospital & "' and inuse = 1 " & _
                "order by PrintPriority , AgeFromDays"
120       Set tb = New Recordset
130       RecOpenServer 0, tb, sql
140       Do While Not tb.EOF
150           blnShowThis = False
160           blnFound = False

170           If ShowAllAgeRanges <> tb!LongName Then  'is other ages
180               For intN = 1 To gNormal.Rows - 1
190                   If gNormal.TextMatrix(intN, 0) = tb!LongName Then
200                       blnFound = True
210                       gNormal.Col = 0
220                       gNormal.Row = intN
230                       Set gNormal.CellPicture = imgSquarePlus.Picture
240                       gNormal.CellPictureAlignment = flexAlignRightCenter
250                       Exit For
260                   End If
270               Next
280           Else
290               blnShowThis = False
300               For intN = 1 To gNormal.Rows - 1
310                   If gNormal.TextMatrix(intN, 0) = tb!LongName Then
320                       blnFound = True
330                       If ShowAllAgeRanges = tb!LongName Then
340                           blnShowThis = True
350                           Exit For
360                       End If
370                   End If
380               Next
390           End If

400           If Not blnFound Or blnShowThis Then
410               strS = tb!LongName & vbTab & tb!ShortName & vbTab & _
                         tb!Code & vbTab & tb!AgeFromDays & vbTab & tb!AgeToDays & vbTab & _
                         tb!MaleLow & vbTab & tb!MaleHigh & vbTab & _
                         tb!FemaleLow & vbTab & tb!FemaleHigh & vbTab & _
                         tb!FlagMaleLow & vbTab & tb!FlagMaleHigh & vbTab & _
                         tb!FlagFemaleLow & vbTab & tb!FlagFemaleHigh & vbTab & _
                         tb!Code & vbTab & tb!DP & vbTab & _
                         tb!AgeFromDays & vbTab & _
                         tb!AgeToDays & vbTab & _
                         tb!PrintPriority

420               gNormal.AddItem strS
430               gNormal.Row = gNormal.Rows - 1
440           End If
450           tb.MoveNext
460       Loop

470       If ShowAllAgeRanges <> "" Then
480           gNormal.Col = 0
490           For intN = 1 To gNormal.Rows - 1
500               If gNormal.TextMatrix(intN, 0) = ShowAllAgeRanges Then
510                   gNormal.Row = intN
520                   Set gNormal.CellPicture = imgSquareMinus.Picture
530                   gNormal.CellPictureAlignment = flexAlignRightCenter
540               End If
550           Next
560       End If

570       If RowTop <> 0 Then
580           gNormal.TopRow = RowTop
590       End If

600       With gNormal
610           If .Rows > 2 Then
620               .RemoveItem 1
630           End If
640           .Visible = True
650       End With

660       AdjustAgeView

670       Exit Sub

FillgNormal_Error:

          Dim strES As String
          Dim intEL As Integer

680       intEL = Erl
690       strES = Err.Description
700       LogError "frmBioRanges", "FillgNormal", intEL, strES, sql


End Sub


Private Function GetActiveGrid() As MSFlexGrid

10        On Error GoTo GetActiveGrid_Error

20        Select Case SSTab1.Tab
          Case 0: Set GetActiveGrid = gCodes
30        Case 1: Set GetActiveGrid = gNormal
40        Case 2: Set GetActiveGrid = gPlausible
50        Case 3: Set GetActiveGrid = gMasks
60        Case 4: Set GetActiveGrid = gSequence
70        End Select

80        Exit Function

GetActiveGrid_Error:

          Dim strES As String
          Dim intEL As Integer

90        intEL = Erl
100       strES = Err.Description
110       LogError "frmBioRanges", "GetActiveGrid", intEL, strES


End Function

Private Function GetDiscipline()

          Dim intN As Integer
          Dim RetVal As String

10        On Error GoTo GetDiscipline_Error

20        For intN = 0 To 2
30            If optDiscipline(intN) Then
40                RetVal = Left$(optDiscipline(intN).Caption, 3)
50            End If
60        Next

70        GetDiscipline = RetVal

80        Exit Function

GetDiscipline_Error:

          Dim strES As String
          Dim intEL As Integer

90        intEL = Erl
100       strES = Err.Description
110       LogError "frmBioRanges", "GetDiscipline", intEL, strES


End Function

Private Sub bMoveDown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        On Error GoTo bMoveDown_MouseDown_Error

20        FireDown

30        tmrDown.Interval = 250
40        FireCounter = 0

50        tmrDown.Enabled = True

60        Exit Sub

bMoveDown_MouseDown_Error:

          Dim strES As String
          Dim intEL As Integer

70        intEL = Erl
80        strES = Err.Description
90        LogError "frmBioRanges", "bMoveDown_MouseDown", intEL, strES


End Sub


Private Sub bMoveDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        On Error GoTo bMoveDown_MouseUp_Error

20        tmrDown.Enabled = False

30        SaveSequence

40        Exit Sub

bMoveDown_MouseUp_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmBioRanges", "bMoveDown_MouseUp", intEL, strES


End Sub


Private Sub bMoveUp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        On Error GoTo bMoveUp_MouseDown_Error

20        FireUp

30        tmrUp.Interval = 250
40        FireCounter = 0

50        tmrUp.Enabled = True

60        Exit Sub

bMoveUp_MouseDown_Error:

          Dim strES As String
          Dim intEL As Integer

70        intEL = Erl
80        strES = Err.Description
90        LogError "frmBioRanges", "bMoveUp_MouseDown", intEL, strES


End Sub


Private Sub bMoveUp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        On Error GoTo bMoveUp_MouseUp_Error

20        tmrUp.Enabled = False

30        SaveSequence

40        Exit Sub

bMoveUp_MouseUp_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmBioRanges", "bMoveUp_MouseUp", intEL, strES


End Sub


Private Sub cmdAddAgeSpecific_Click()

          Dim n As Integer
          Dim s As String
          Dim strSampleType As String
          Dim Discipline As String

10        On Error GoTo cmdAddAgeSpecific_Click_Error

20        If cmbHospital = "" Then
30            cmbHospital = HospName(0)
40        End If
          'If cmbCategory = "" Then
          '  cmbCategory = "Human"
          'End If
50        If cmbSampleType = "" Then
60            cmbSampleType = "Serum"
70        End If
80        strSampleType = ListCodeFor("ST", cmbSampleType)

90        Discipline = GetDiscipline()

          'eg Caption = "&Add PT Age Specific Range"
100       n = InStr(cmdAddAgeSpecific.Caption, " Age")
110       s = Mid$(cmdAddAgeSpecific.Caption, 6, n - 6)

120       With frmBioMultiAddAge
130           .Analyte = s
140           .SampleType = strSampleType
150           .Hospital = cmbHospital
160           .Category = cmbCategory
170           .Discipline = Discipline
180           .Show 1
190       End With
200       FillgNormal

210       Exit Sub

cmdAddAgeSpecific_Click_Error:

          Dim strES As String
          Dim intEL As Integer

220       intEL = Erl
230       strES = Err.Description
240       LogError "frmBioRanges", "cmdAddAgeSpecific_Click", intEL, strES


End Sub

Private Sub cmdCancel_Click()

10        Unload Me

End Sub

Private Sub cmdXL_Click()

10        On Error GoTo cmdXL_Click_Error

20        Select Case SSTab1
          Case 0: ExportFlexGrid gCodes, Me
30        Case 1: ExportFlexGrid gNormal, Me
40        Case 2: ExportFlexGrid gPlausible, Me
50        Case 3: ExportFlexGrid gMasks, Me
60        End Select

70        Exit Sub

cmdXL_Click_Error:

          Dim strES As String
          Dim intEL As Integer

80        intEL = Erl
90        strES = Err.Description
100       LogError "frmBioRanges", "cmdXL_Click", intEL, strES


End Sub

Private Sub Form_Load()

10        On Error GoTo Form_Load_Error

20        FillStrUnits
30        FillSampleTypes
40        FillCategories
50        FillHospitals

60        gNormal.ColWidth(0) = 1500
70        gNormal.ColWidth(1) = 0

80        gCodes.ColWidth(0) = 1500
90        gCodes.ColWidth(1) = 0

100       gPlausible.ColWidth(0) = 1500
110       gPlausible.ColWidth(1) = 0

120       gMasks.ColWidth(0) = 1500
130       gMasks.ColWidth(1) = 0

140       SSTab1.Tab = pViewTab

150       Select Case pViewTab
          Case 0: FillgCodes
160       Case 1: FillgNormal
170       Case 2: FillgPlausible
180       Case 3: FillgMasks
190       Case 4: FillgSequence
200       End Select

210       Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

220       intEL = Erl
230       strES = Err.Description
240       LogError "frmBioRanges", "Form_Load", intEL, strES


End Sub


Private Sub FillgSequence()

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo FillgSequence_Error

20        gSequence.Visible = False
30        gSequence.Rows = 2
40        gSequence.AddItem ""
50        gSequence.RemoveItem 1

60        sql = "Select ShortName, LongName, max(PrintPriority) as M " & _
                "from BioTestDefinitions where inuse = 1 " & _
                "GROUP BY ShortName, LongName " & _
                "Order by M"
70        Set tb = New Recordset
80        RecOpenServer 0, tb, sql
90        Do While Not tb.EOF
100           gSequence.AddItem tb!ShortName & vbTab & _
                                tb!LongName & ""
110           tb.MoveNext
120       Loop

130       If gSequence.Rows > 2 Then gSequence.RemoveItem 1
140       gSequence.Visible = True

150       Exit Sub

FillgSequence_Error:

          Dim strES As String
          Dim intEL As Integer

160       intEL = Erl
170       strES = Err.Description
180       LogError "frmBioRanges", "FillgSequence", intEL, strES, sql


End Sub

Private Sub SaveMasks()

          Dim intN As Integer
          Dim sql As String
          Dim strSampleType As String
          Dim intLIH As Integer
          Dim intO As Integer
          Dim Discipline As String
          Dim xSave As Integer
          Dim ds As Recordset
          Dim f As Field
          Dim DefIndex As Long
          Dim tb As Recordset
          Dim tempDef As Long

10        On Error GoTo SaveMasks_Error

20        If cmbHospital = "" Then
30            cmbHospital = HospName(0)
40        End If
          'If cmbCategory = "" Then
          '  cmbCategory = "Human"
          'End If
50        If cmbSampleType = "" Then
60            cmbSampleType = "Serum"
70        End If
80        strSampleType = ListCodeFor("ST", cmbSampleType)
90        Discipline = GetDiscipline()

100       intN = gMasks.Row

110       sql = "Select * from " & Discipline & "Testdefinitions order by defindex desc"
120       Set tb = New Recordset
130       RecOpenServer 0, tb, sql

140       If Not tb.EOF Then tempDef = tb!DefIndex + 1

150       DefIndex = tempDef + 1

160       sql = "Select * from " & Discipline & "Testdefinitions " & _
                "where LongName = '" & gCodes.TextMatrix(intN, 0) & "' " & _
                "and SampleType = '" & strSampleType & "' " & _
                "and Category = '" & cmbCategory & "' " & _
                "and Hospital = '" & cmbHospital & "' and inuse = 1 and code = '" & gCodes.TextMatrix(intN, 2) & "' order by agefromdays"
170       Set tb = New Recordset
180       RecOpenServer 0, tb, sql
190       Set ds = New Recordset
200       RecOpenServer 0, ds, sql
210       Do While Not tb.EOF
220           tb!InUse = 0
230           tb.Update
240           ds.AddNew
250           For Each f In ds.Fields
260               If UCase(f.Name) <> UCase("rowguid") Then
270                   ds(f.Name) = tb(f.Name)
280               End If
290           Next
300           ds!DefIndex = DefIndex
310           ds.Update
320           DefIndex = DefIndex + 1
330           tb.MoveNext
340       Loop

350       sql = "update " & Discipline & "Testdefinitions set inuse = 1 where defindex > " & tempDef & " and defindex < " & DefIndex & ""
360       Cnxn(0).Execute sql

370       xSave = gMasks.Col
380       gMasks.Col = 3    'Old
390       intO = IIf(gMasks.CellPicture = imgSquareTick, 1, 0)
400       gMasks.Col = xSave

410       sql = "Update " & Discipline & "TestDefinitions SET " & _
                "LIH = '" & intLIH & "', " & _
                "O = '" & intO & "' " & _
                "where code = '" & gMasks.TextMatrix(intN, 2) & "' and inuse = 1 " & _
                "and SampleType = '" & strSampleType & "' " & _
                "and Category = '" & cmbCategory & "' " & _
                "and Hospital = '" & cmbHospital & "'"
420       Cnxn(0).Execute sql

430       Exit Sub

SaveMasks_Error:

          Dim strES As String
          Dim intEL As Integer

440       intEL = Erl
450       strES = Err.Description
460       LogError "frmBioRanges", "SaveMasks", intEL, strES, sql

End Sub

Private Sub SavePlausible()

          Dim intN As Integer
          Dim sql As String
          Dim strSampleType As String
          Dim Discipline As String
          Dim ds As Recordset
          Dim f As Field
          Dim DefIndex As Long
          Dim tb As Recordset
          Dim tempDef As Long

10        On Error GoTo SavePlausible_Error

20        If cmbHospital = "" Then
30            cmbHospital = HospName(0)
40        End If
          'If cmbCategory = "" Then
          '  cmbCategory = "Human"
          'End If
50        If cmbSampleType = "" Then
60            cmbSampleType = "Serum"
70        End If
80        strSampleType = ListCodeFor("ST", cmbSampleType)

90        Discipline = GetDiscipline()

100       gPlausible.Col = 6
110       intN = gPlausible.Row


120       sql = "Select * from " & Discipline & "Testdefinitions order by defindex desc"
130       Set tb = New Recordset
140       RecOpenServer 0, tb, sql

150       If Not tb.EOF Then tempDef = tb!DefIndex + 1

160       DefIndex = tempDef + 1

170       sql = "Select * from " & Discipline & "Testdefinitions " & _
                "where LongName = '" & gCodes.TextMatrix(intN, 0) & "' " & _
                "and SampleType = '" & strSampleType & "' " & _
                "and Category = '" & cmbCategory & "' " & _
                "and Hospital = '" & cmbHospital & "' and inuse = 1 and code = '" & gCodes.TextMatrix(intN, 2) & "' order by agefromdays"
180       Set tb = New Recordset
190       RecOpenServer 0, tb, sql
200       Set ds = New Recordset
210       RecOpenServer 0, ds, sql
220       Do While Not tb.EOF
230           tb!InUse = 0
240           tb.Update
250           ds.AddNew
260           For Each f In ds.Fields
270               If UCase(f.Name) <> UCase("rowguid") Then
280                   ds(f.Name) = tb(f.Name)
290               End If
300           Next
310           ds!DefIndex = DefIndex
320           ds.Update
330           DefIndex = DefIndex + 1
340           tb.MoveNext
350       Loop


360       sql = "update " & Discipline & "Testdefinitions set inuse = 1 where defindex > " & tempDef & " and defindex < " & DefIndex & ""
370       Cnxn(0).Execute sql


380       sql = "Update " & Discipline & "TestDefinitions " & _
                "Set PlausibleLow = '" & Val(gPlausible.TextMatrix(intN, 3)) & "', " & _
                "PlausibleHigh = '" & Val(gPlausible.TextMatrix(intN, 4)) & "', " & _
                "AutoValLow = '" & Val(gPlausible.TextMatrix(intN, 5)) & "', " & _
                "AutoValHigh = '" & Val(gPlausible.TextMatrix(intN, 6)) & "', " & _
                "DeltaLimit = '" & Val(gPlausible.TextMatrix(intN, 8)) & "', " & _
                "DoDelta = '" & IIf(gPlausible.CellPicture = imgSquareTick.Picture, 1, 0) & "' " & _
                "where code = '" & gPlausible.TextMatrix(intN, 2) & "' and inuse = 1 " & _
                "and SampleType = '" & strSampleType & "' " & _
                "and Category = '" & cmbCategory & "' " & _
                "and Hospital = '" & cmbHospital & "'"
390       Cnxn(0).Execute sql

400       Exit Sub

SavePlausible_Error:

          Dim strES As String
          Dim intEL As Integer

410       intEL = Erl
420       strES = Err.Description
430       LogError "frmBioRanges", "SavePlausible", intEL, strES, sql


End Sub

Private Sub SaveNormal()

          Dim intN As Integer
          Dim tb As Recordset
          Dim sql As String
          Dim strSampleType As String
          Dim Discipline As String
          Dim DefIndex As Long
          Dim tempDef As Long
          Dim ds As Recordset
          Dim f As Field

10        On Error GoTo SaveNormal_Error

20        If cmbHospital = "" Then
30            cmbHospital = HospName(0)
40        End If
          'If cmbCategory = "" Then
          '  cmbCategory = "Human"
          'End If
50        If cmbSampleType = "" Then
60            cmbSampleType = "Serum"
70        End If
80        strSampleType = ListCodeFor("ST", cmbSampleType)

90        Discipline = GetDiscipline()

          'strS = "<Long Name             |<Short Name   |^Age From |^Age To   "                      '0 to 3
          'strS = strS & "|^Normal Male Low|^Normal Male High|^Normal Female Low|^Normal Female High" '4 to 7
          'strS = strS & "|^Flag Male Low|^Flag Male High|^Flag Female Low|^Flag Female High"         '8 to 11
          'strS = strS & "|<Code    |^Dec.Pl  "       '12 to 13
          'strS = strS & "|<AgeFromDays|<AgeToDays|^Print Priority"                                   '14 to 16

100       intN = gNormal.Row


110       sql = "Select * from " & Discipline & "Testdefinitions order by defindex desc"
120       Set tb = New Recordset
130       RecOpenServer 0, tb, sql

140       If Not tb.EOF Then tempDef = tb!DefIndex + 1

150       DefIndex = tempDef + 1

160       sql = "Select * from " & Discipline & "Testdefinitions " & _
                "where LongName = '" & gCodes.TextMatrix(intN, 0) & "' " & _
                "and SampleType = '" & strSampleType & "' " & _
                "and Category = '" & cmbCategory & "' " & _
                "and Hospital = '" & cmbHospital & "' " & _
                "and inuse = 1 and code = '" & gCodes.TextMatrix(intN, 2) & "' " & _
                "AND AgeFromDays = '" & gNormal.TextMatrix(intN, 15) & "' " & _
                "order by agefromdays"
170       Set tb = New Recordset
180       RecOpenServer 0, tb, sql
190       Set ds = New Recordset
200       RecOpenServer 0, ds, sql
210       Do While Not tb.EOF
220           tb!InUse = 0
230           tb.Update
240           ds.AddNew
250           For Each f In ds.Fields
260               If UCase(f.Name) <> UCase("rowguid") Then
270                   ds(f.Name) = tb(f.Name)
280               End If
290           Next
300           ds!DefIndex = DefIndex
310           ds.Update
320           tb.MoveNext
330       Loop


340       sql = "update " & Discipline & "Testdefinitions set inuse = 1 where defindex = " & DefIndex & ""
350       Cnxn(0).Execute sql


360       sql = "Update " & Discipline & "TestDefinitions SET " & _
                "MaleLow = " & Val(gNormal.TextMatrix(intN, 5)) & ", " & _
                "MaleHigh = " & Val(gNormal.TextMatrix(intN, 6)) & ", " & _
                "FemaleLow = " & Val(gNormal.TextMatrix(intN, 7)) & ", " & _
                "FemaleHigh = " & Val(gNormal.TextMatrix(intN, 8)) & ", " & _
                "FlagMaleLow = " & Val(gNormal.TextMatrix(intN, 9)) & ", " & _
                "FlagMaleHigh = " & Val(gNormal.TextMatrix(intN, 10)) & ", " & _
                "FlagFemaleLow = " & Val(gNormal.TextMatrix(intN, 11)) & ", " & _
                "FlagFemaleHigh = " & Val(gNormal.TextMatrix(intN, 12)) & " " & _
                "WHERE LongName = '" & gNormal.TextMatrix(intN, 0) & "' " & _
                "and inuse = 1 and code = '" & gCodes.TextMatrix(intN, 2) & "'" & _
                "AND SampleType = '" & strSampleType & "' " & _
                "AND Category = '" & cmbCategory & "' " & _
                "AND Hospital = '" & cmbHospital & "' " & _
                "AND AgeFromDays = '" & gNormal.TextMatrix(intN, 15) & "' and defindex = " & DefIndex & ""
370       Cnxn(0).Execute sql

          '    tb!Code = tbOrig!Code & ""
          '    tb!BarCode = tbOrig!BarCode & ""
          '    tb!ImmunoCode = tbOrig!ImmunoCode & ""
          '    tb!Units = tbOrig!Units & ""
          '    tb!DP = tbOrig!DP

          '    tb!ActiveFromDate = Format$(Now, "dd/mmm/yyyy")
          '    tb!ActiveToDate = Format$(Now, "dd/mmm/yyyy")
          '
          '    tb!AgeFromDays = gNormal.TextMatrix(intN, 21)
          '    tb!AgeToDays = gNormal.TextMatrix(intN, 22)
          '    tb!PrintPriority = tbOrig!PrintPriority
          '
          '    tb!KnownToAnalyser = tbOrig!KnownToAnalyser
          '    tb!Analyser = tbOrig!Analyser
          '    tb!InUse = tbOrig!InUse
          '    tb!SplitList = tbOrig!SplitList
          '    tb!EOD = tbOrig!EOD
          '
          '    tb!Hospital = cmbHospital
          '    tb!Category = cmbCategory
          '    tb!SampleType = strSampleType
          '
          '    tb!h = tbOrig!h
          '    tb!s = tbOrig!s
          '    tb!l = tbOrig!l
          '    tb!o = tbOrig!o
          '    tb!g = tbOrig!g
          '    tb!J = tbOrig!J
          '
          '    tb!DoDelta = tbOrig!DoDelta
          '    tb!DeltaLimit = tbOrig!DeltaLimit
          '
          '    tb!Printable = tbOrig!Printable


380       Exit Sub

SaveNormal_Error:

          Dim strES As String
          Dim intEL As Integer

390       intEL = Erl
400       strES = Err.Description
410       LogError "frmBioRanges", "SaveNormal", intEL, strES, sql


End Sub

Private Sub AdjustAgeView()

          Dim intN As Integer

10        On Error GoTo AdjustAgeView_Error

20        For intN = 1 To gNormal.Rows - 1
30            If IsNumeric(gNormal.TextMatrix(intN, 14)) Then
40                gNormal.TextMatrix(intN, 2) = dmyFromCount(gNormal.TextMatrix(intN, 14))
50            End If
60            If IsNumeric(gNormal.TextMatrix(intN, 15)) Then
70                gNormal.TextMatrix(intN, 3) = dmyFromCount(gNormal.TextMatrix(intN, 15))
80            End If
90        Next

100       Exit Sub

AdjustAgeView_Error:

          Dim strES As String
          Dim intEL As Integer

110       intEL = Erl
120       strES = Err.Description
130       LogError "frmBioRanges", "AdjustAgeView", intEL, strES


End Sub
Private Sub FillSampleTypes()

          Dim sql As String
          Dim tb As Recordset

10        On Error GoTo FillSampleTypes_Error

20        sql = "Select Text from Lists where " & _
                "ListType = 'ST' " & _
                "order by ListOrder"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql

50        cmbSampleType.Clear
60        Do While Not tb.EOF
70            cmbSampleType.AddItem tb!Text & ""
80            tb.MoveNext
90        Loop

100       cmbSampleType = "Serum"

110       Exit Sub

FillSampleTypes_Error:

          Dim strES As String
          Dim intEL As Integer

120       intEL = Erl
130       strES = Err.Description
140       LogError "frmBioRanges", "FillSampleTypes", intEL, strES, sql


End Sub

Private Sub FillHospitals()

          Dim sql As String
          Dim tb As Recordset

10        On Error GoTo FillHospitals_Error

20        sql = "Select Text from Lists where " & _
                "ListType = 'HO' " & _
                "order by ListOrder"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql

50        cmbHospital.Clear
60        Do While Not tb.EOF
70            cmbHospital.AddItem tb!Text & ""
80            tb.MoveNext
90        Loop

100       cmbHospital.Text = HospName(0)

110       Exit Sub

FillHospitals_Error:

          Dim strES As String
          Dim intEL As Integer

120       intEL = Erl
130       strES = Err.Description
140       LogError "frmBioRanges", "FillHospitals", intEL, strES, sql


End Sub

Private Sub FillCategories()

          Dim sql As String
          Dim tb As Recordset

10        On Error GoTo FillCategories_Error

20        sql = "Select Cat from Categorys " & _
                "order by ListOrder"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql

50        cmbCategory.Clear
60        Do While Not tb.EOF
70            cmbCategory.AddItem tb!Cat & ""
80            tb.MoveNext
90        Loop

100       If cmbCategory.ListCount > 0 Then
110           cmbCategory.ListIndex = 0
120       End If

130       Exit Sub

FillCategories_Error:

          Dim strES As String
          Dim intEL As Integer

140       intEL = Erl
150       strES = Err.Description
160       LogError "frmBioRanges", "FillCategories", intEL, strES, sql


End Sub

Private Sub gCodes_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

          Dim s As String
          Dim f As Form
          Dim SaveMe As Boolean
          '<Long Name  |<Short Name                         0 to 1
          '^Code |^Bar Code |^Immuno Code|^Units  |^Dec.Pl  2 to 6
          '^Printable|^Known to Analyser|^Analyser Code|    7 to 9
          '^In Use|^End of Day                              10 to 11

10        On Error GoTo gCodes_MouseUp_Error

20        SaveMe = False
30        Select Case gCodes.MouseCol

          Case 2

40        Case 3
50            gCodes.TextMatrix(gCodes.Row, 3) = iBOX("Scan Barcode for " & gCodes.TextMatrix(gCodes.Row, 0), , gCodes.TextMatrix(gCodes.Row, 3))
60            SaveMe = True

70        Case 4, 9
80            s = "Enter " & gCodes.TextMatrix(0, gCodes.Col) & " for " & gCodes.TextMatrix(gCodes.Row, 0)
90            gCodes = iBOX(s, , gCodes.TextMatrix(gCodes.Row, gCodes.Col))
100           SaveMe = True

110       Case 5
120           Set f = New fcdrDBox
130           With f
140               .Options = strUnits
150               .Prompt = "Enter Units for " & gCodes.TextMatrix(gCodes.Row, 0)
160               .Show 1
170               gCodes = .ReturnValue
180           End With
190           Unload f
200           Set f = Nothing
210           SaveMe = True

220       Case 6
230           Select Case gCodes.TextMatrix(gCodes.Row, 6)
              Case "0": gCodes.TextMatrix(gCodes.Row, 6) = "1"
240           Case "1": gCodes.TextMatrix(gCodes.Row, 6) = "2"
250           Case "2": gCodes.TextMatrix(gCodes.Row, 6) = "3"
260           Case Else: gCodes.TextMatrix(gCodes.Row, 6) = "0"
270           End Select
280           SaveMe = True


290       Case 7, 8, 10, 11
300           If gCodes.CellPicture = imgSquareCross.Picture Then
310               Set gCodes.CellPicture = imgSquareTick.Picture
320           Else
330               Set gCodes.CellPicture = imgSquareCross.Picture
340           End If
350           SaveMe = True

360       End Select

370       If SaveMe Then
380           SaveCodesNew
390       End If

400       Exit Sub

gCodes_MouseUp_Error:

          Dim strES As String
          Dim intEL As Integer

410       intEL = Erl
420       strES = Err.Description
430       LogError "frmBioRanges", "gCodes_MouseUp", intEL, strES


End Sub


Private Sub gMasks_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

          Dim xx As Integer
          Dim yy As Integer

10        On Error GoTo gMasks_MouseUp_Error

20        yy = gMasks.MouseRow
30        If yy = 0 Then Exit Sub

40        xx = gMasks.Col
50        Debug.Print xx

60        Select Case xx

          Case 3    'Old
70            If gMasks.CellPicture = imgSquareCross.Picture Then
80                Set gMasks.CellPicture = imgSquareTick.Picture
90            Else
100               Set gMasks.CellPicture = imgSquareCross.Picture
110           End If
120           SaveMasks

130       Case 4, 5, 6    'LIH
140           Select Case gMasks.TextMatrix(yy, xx)
              Case "": gMasks.TextMatrix(yy, xx) = "1+"
150           Case "1+": gMasks.TextMatrix(yy, xx) = "2+"
160           Case "2+": gMasks.TextMatrix(yy, xx) = "3+"
170           Case "3+": gMasks.TextMatrix(yy, xx) = "4+"
180           Case "4+": gMasks.TextMatrix(yy, xx) = "5+"
190           Case Else: gMasks.TextMatrix(yy, xx) = ""
200           End Select
210           SaveMasks

220       End Select

230       Exit Sub

gMasks_MouseUp_Error:

          Dim strES As String
          Dim intEL As Integer

240       intEL = Erl
250       strES = Err.Description
260       LogError "frmBioRanges", "gMasks_MouseUp", intEL, strES


End Sub


Private Sub gNormal_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

          Dim s As String
          Dim SaveMe As Boolean

10        On Error GoTo gNormal_MouseUp_Error

20        If gNormal.MouseRow = 0 Then
30            cmdAddAgeSpecific.Visible = False
40            Exit Sub
50        End If

60        cmdAddAgeSpecific.Caption = "&Add " & gNormal.TextMatrix(gNormal.Row, 0) & " Age Specific Range"
70        cmdAddAgeSpecific.Visible = True
          'strS = "<Long Name             |<Short Name   |^Age From |^Age To   "                      '0 to 3
          'strS = strS & "|^Normal Male Low|^Normal Male High|^Normal Female Low|^Normal Female High" '4 to 7
          'strS = strS & "|^Flag Male Low|^Flag Male High|^Flag Female Low|^Flag Female High"         '8 to 11
          'strS = strS & "|^Plausible Low|^Plausible High|^AutoVal Low|^AutoVal High"                 '12 to 15
          'strS = strS & "|<Code    |^Dec.Pl  "       '16 to 20
          'strS = strS & "|<AgeFromDays|<AgeToDays|^Print Priority"                                   '21 to 23
          'strS = strS & "|<RowNumber|<ModifiedFlag"                                                  '24 to 25

80        SaveMe = False

90        Debug.Print gNormal.MouseCol

100       Select Case gNormal.MouseCol

          Case 0
110           gNormal.Col = 0
120           If gNormal.CellPicture = imgSquarePlus.Picture Then
130               ShowAllAgeRanges = gNormal.TextMatrix(gNormal.MouseRow, 0)
140           Else
150               ShowAllAgeRanges = ""
160           End If
170           FillgNormal gNormal.TopRow

180       Case 2, 3
190           iMsg "Ages are not Editable!", vbExclamation

200       Case 5 To 13
210           s = "Enter " & gNormal.TextMatrix(0, gNormal.Col) & " Range for " & gNormal.TextMatrix(gNormal.Row, 0)
220           gNormal = iBOX(s, , gNormal.TextMatrix(gNormal.Row, gNormal.Col))
230           SaveMe = True
240       Case 14    'dp
250           gNormal = Val(gNormal) + 1
260           If Val(gNormal) > 4 Then
270               gNormal = 0
280           End If
290           SaveMe = True

300       End Select

310       If SaveMe Then
320           SaveNormal
330       End If

340       Exit Sub

gNormal_MouseUp_Error:

          Dim strES As String
          Dim intEL As Integer

350       intEL = Erl
360       strES = Err.Description
370       LogError "frmBioRanges", "gNormal_MouseUp", intEL, strES


End Sub


Private Sub gPlausible_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

          Dim s As String
          Dim SaveMe As Boolean

10        On Error GoTo gPlausible_MouseUp_Error

20        SaveMe = False

30        Select Case gPlausible.MouseCol

          Case 3 To 6, 8
40            s = "Enter " & gPlausible.TextMatrix(0, gPlausible.Col) & _
                  " Range for " & gPlausible.TextMatrix(gPlausible.Row, 0)
50            gPlausible = iBOX(s, , gPlausible.TextMatrix(gPlausible.Row, gPlausible.Col))
60            SaveMe = True

70        Case 7    'DoDelta
80            If gPlausible.CellPicture = imgSquareCross.Picture Then
90                Set gPlausible.CellPicture = imgSquareTick.Picture
100           Else
110               Set gPlausible.CellPicture = imgSquareCross.Picture
120           End If
130           SaveMe = True

140       End Select

150       If SaveMe Then
160           SavePlausible
170       End If

180       Exit Sub

gPlausible_MouseUp_Error:

          Dim strES As String
          Dim intEL As Integer

190       intEL = Erl
200       strES = Err.Description
210       LogError "frmBioRanges", "gPlausible_MouseUp", intEL, strES


End Sub


Private Sub gSequence_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

          Dim n As Integer
          Dim yy As Integer
          Dim ySave As Integer

10        On Error GoTo gSequence_MouseUp_Error

20        If gSequence.MouseRow = 0 Then Exit Sub

30        ySave = gSequence.Row
40        gSequence.Col = 0
50        For yy = 1 To gSequence.Rows - 1
60            gSequence.Row = yy
70            If gSequence.CellBackColor = vbYellow Then
80                For n = 0 To 1
90                    gSequence.Col = n
100                   gSequence.CellBackColor = 0
110               Next
120               Exit For
130           End If
140       Next
150       gSequence.Row = ySave
160       For n = 0 To 1
170           gSequence.Col = n
180           gSequence.CellBackColor = vbYellow
190       Next
200       bMoveUp.Visible = True
210       bMoveDown.Visible = True

220       Exit Sub

gSequence_MouseUp_Error:

          Dim strES As String
          Dim intEL As Integer

230       intEL = Erl
240       strES = Err.Description
250       LogError "frmBioRanges", "gSequence_MouseUp", intEL, strES


End Sub


Private Sub optDiscipline_Click(Index As Integer)

10        On Error GoTo optDiscipline_Click_Error

20        Me.Caption = "NetAcquire - " & optDiscipline(Index).Caption

30        Select Case SSTab1.Tab
          Case 0: FillgCodes
40        Case 1: FillgNormal
50        Case 2: FillgPlausible
60        Case 3: FillgMasks
70        Case 4: FillgSequence
80        End Select

90        Exit Sub

optDiscipline_Click_Error:

          Dim strES As String
          Dim intEL As Integer

100       intEL = Erl
110       strES = Err.Description
120       LogError "frmBioRanges", "optDiscipline_Click", intEL, strES


End Sub
Private Sub FillgMasks()

          Dim tb As Recordset
          Dim sql As String
          Dim strS As String
          Dim strSampleType As String
          Dim blnFound As Boolean
          Dim intN As Integer
          Dim intLIH As Integer
          Dim Discipline As String

10        On Error GoTo FillgMasks_Error

20        If cmbHospital = "" Then
30            cmbHospital = HospName(0)
40        End If
          'If cmbCategory = "" Then
          '  cmbCategory = "Human"
          'End If
50        If cmbSampleType = "" Then
60            cmbSampleType = "Serum"
70        End If
80        strSampleType = ListCodeFor("ST", cmbSampleType)
90        Discipline = GetDiscipline

100       ClearGrid

110       gMasks.Col = 2    'Old

120       sql = "SELECT LongName, ShortName, " & _
                "COALESCE (LIH, 0) AS LIH, " & _
                "COALESCE (O, 0) AS O " & _
                "FROM " & Discipline & "TestDefinitions WHERE " & _
                "SampleType = '" & strSampleType & "' " & _
                "AND Category = '" & cmbCategory & "' " & _
                "AND Hospital = '" & cmbHospital & "' and inuse = 1 " & _
                "ORDER BY PrintPriority"
130       Set tb = New Recordset
140       RecOpenServer 0, tb, sql
150       Do While Not tb.EOF
160           blnFound = False
170           For intN = 1 To gMasks.Rows - 1
180               If gMasks.TextMatrix(intN, 2) = tb!ShortName & "" Then
190                   blnFound = True
200                   Exit For
210               End If
220           Next
230           If Not blnFound Then
240               strS = tb!LongName & vbTab & _
                         tb!ShortName & vbTab & vbTab
                  'LIH
                  'L 100 to 600
                  'I  10 to  60
                  'H   1 to   6
250               intLIH = tb!LIH
260               If intLIH > 99 Then
270                   strS = strS & Format$(intLIH \ 100) & "+"
280                   intLIH = intLIH Mod 100
290               End If
300               strS = strS & vbTab

310               If intLIH > 9 Then
320                   strS = strS & Format$(intLIH \ 10) & "+"
330                   intLIH = intLIH Mod 10
340               End If
350               strS = strS & vbTab
360               If intLIH > 0 Then
370                   strS = strS & Format$(intLIH) & "+"
380               End If

390               gMasks.AddItem strS

400               gMasks.Row = gMasks.Rows - 1
410               Set gMasks.CellPicture = IIf(tb!o <> 0, imgSquareTick.Picture, imgSquareCross.Picture)
420               gMasks.CellPictureAlignment = flexAlignCenterCenter

430           End If

440           tb.MoveNext

450       Loop

460       If gMasks.Rows > 2 Then
470           gMasks.RemoveItem 1
480       End If
490       gMasks.Visible = True

500       Exit Sub

FillgMasks_Error:

          Dim strES As String
          Dim intEL As Integer

510       intEL = Erl
520       strES = Err.Description
530       LogError "frmBioRanges", "FillgMasks", intEL, strES, sql


End Sub

Private Sub FillgCodes()

          Dim tb As Recordset
          Dim sql As String
          Dim strS As String
          Dim strSampleType As String
          Dim blnFound As Boolean
          Dim intN As Integer
          Dim Discipline As String

10        On Error GoTo FillgCodes_Error

20        If cmbHospital = "" Then
30            cmbHospital = HospName(0)
40        End If
          'If cmbCategory = "" Then
          '  cmbCategory = "Human"
          'End If
50        If cmbSampleType = "" Then
60            cmbSampleType = "Serum"
70        End If
80        strSampleType = ListCodeFor("ST", cmbSampleType)

90        Discipline = GetDiscipline

100       ClearGrid

          '<Long Name  |<Short Name                         0 to 1
          '^Code |^Bar Code |^Immuno Code|^Units  |^Dec.Pl  2 to 6
          '^Printable|^Known to Analyser|^Analyser Code|    7 to 9
          '^In Use|^End of Day                              10 to 11

110       sql = "Select * from " & Discipline & "TestDefinitions where " & _
                "SampleType = '" & strSampleType & "' " & _
                "and Category = '" & cmbCategory & "' " & _
                "and Hospital = '" & cmbHospital & "' and inuse = 1 " & _
                "order by PrintPriority"
120       Set tb = New Recordset
130       RecOpenServer 0, tb, sql
140       Do While Not tb.EOF
150           blnFound = False
160           For intN = 1 To gCodes.Rows - 1
170               If gCodes.TextMatrix(intN, 2) = tb!Code & "" Then
180                   blnFound = True
190                   Exit For
200               End If
210           Next
220           If Not blnFound Then
230               strS = tb!LongName & vbTab & _
                         tb!ShortName & vbTab & _
                         tb!Code & vbTab & _
                         tb!BarCode & vbTab & _
                         tb!immunocode & vbTab & _
                         tb!Units & vbTab & _
                         tb!DP & vbTab & _
                         vbTab & vbTab & _
                         Trim$(tb!Analyser & "")

240               gCodes.AddItem strS

250               gCodes.Row = gCodes.Rows - 1
260               gCodes.Col = 7    'Printable
270               If Not IsNull(tb!Printable) Then
280                   Set gCodes.CellPicture = IIf(tb!Printable, imgSquareTick.Picture, imgSquareCross.Picture)
290               Else
300                   Set gCodes.CellPicture = imgSquareCross.Picture
310               End If
320               gCodes.CellPictureAlignment = flexAlignCenterCenter

330               gCodes.Col = 8    'Known to Analyser
340               If Not IsNull(tb!KnownToAnalyser) Then
350                   Set gCodes.CellPicture = IIf(tb!KnownToAnalyser, imgSquareTick.Picture, imgSquareCross.Picture)
360               Else
370                   Set gCodes.CellPicture = imgSquareCross.Picture
380               End If
390               gCodes.CellPictureAlignment = flexAlignCenterCenter

400               gCodes.Col = 10    'In Use
410               If Not IsNull(tb!InUse) Then
420                   Set gCodes.CellPicture = IIf(tb!InUse, imgSquareTick.Picture, imgSquareCross.Picture)
430               Else
440                   Set gCodes.CellPicture = imgSquareCross.Picture
450               End If
460               gCodes.CellPictureAlignment = flexAlignCenterCenter

470               gCodes.Col = 11    'End of Day
480               If Not IsNull(tb!Eod) Then
490                   Set gCodes.CellPicture = IIf(tb!Eod, imgSquareTick.Picture, imgSquareCross.Picture)
500               Else
510                   Set gCodes.CellPicture = imgSquareCross.Picture
520               End If
530               gCodes.CellPictureAlignment = flexAlignCenterCenter

540           End If

550           tb.MoveNext
560       Loop

570       If gCodes.Rows > 2 Then gCodes.RemoveItem 1
580       gCodes.Visible = True

590       Exit Sub

FillgCodes_Error:

          Dim strES As String
          Dim intEL As Integer

600       intEL = Erl
610       strES = Err.Description
620       LogError "frmBioRanges", "FillgCodes", intEL, strES, sql


End Sub

Private Sub FillgPlausible()

          Dim tb As Recordset
          Dim sql As String
          Dim strS As String
          Dim strSampleType As String
          Dim blnFound As Boolean
          Dim intN As Integer
          Dim TableName As String

10        On Error GoTo FillgPlausible_Error

20        If cmbHospital = "" Then
30            cmbHospital = HospName(0)
40        End If
          'If cmbCategory = "" Then
          '  cmbCategory = "Human"
          'End If
50        If cmbSampleType = "" Then
60            cmbSampleType = "Serum"
70        End If
80        strSampleType = ListCodeFor("ST", cmbSampleType)

90        For intN = 0 To 2
100           If optDiscipline(intN) Then
110               TableName = Left$(optDiscipline(intN).Caption, 3)
120           End If
130       Next

140       ClearGrid

150       sql = "SELECT LongName, ShortName, " & _
                "COALESCE (PlausibleLow, 0) AS PlausibleLow, " & _
                "COALESCE (PlausibleHigh, 9999) AS PlausibleHigh, " & _
                "COALESCE (AutoValLow, 0) AS AutoValLow, " & _
                "COALESCE (AutoValHigh, 9999) AS AutoValHigh, " & _
                "COALESCE (DeltaLimit, 0) AS DeltaLimit, " & _
                "COALESCE (DoDelta, 0) AS DoDelta " & _
                "FROM " & TableName & "TestDefinitions WHERE " & _
                "SampleType = '" & strSampleType & "' " & _
                "AND Category = '" & cmbCategory & "' " & _
                "AND Hospital = '" & cmbHospital & "' and inuse = 1 " & _
                "ORDER BY PrintPriority"
160       Set tb = New Recordset
170       RecOpenServer 0, tb, sql
180       Do While Not tb.EOF
190           blnFound = False
200           For intN = 1 To gPlausible.Rows - 1
210               If gPlausible.TextMatrix(intN, 0) = tb!LongName & "" Then
220                   blnFound = True
230                   Exit For
240               End If
250           Next
260           If Not blnFound Then
270               strS = tb!LongName & vbTab & _
                         tb!ShortName & vbTab & _
                         tb!PlausibleLow & vbTab & _
                         tb!PlausibleHigh & vbTab & _
                         tb!AutoValLow & vbTab & _
                         tb!AutoValHigh & vbTab & vbTab & _
                         tb!DeltaLimit

280               gPlausible.AddItem strS

290               gPlausible.Row = gPlausible.Rows - 1
300               gPlausible.Col = 6    'DoDelta
310               Set gPlausible.CellPicture = IIf(tb!DoDelta, imgSquareTick.Picture, imgSquareCross.Picture)
320               gPlausible.CellPictureAlignment = flexAlignCenterCenter

330           End If

340           tb.MoveNext
350       Loop

360       If gPlausible.Rows > 2 Then gPlausible.RemoveItem 1
370       gPlausible.Visible = True

380       Exit Sub

FillgPlausible_Error:

          Dim strES As String
          Dim intEL As Integer

390       intEL = Erl
400       strES = Err.Description
410       LogError "frmBioRanges", "FillgPlausible", intEL, strES, sql


End Sub

Private Sub optLongOrShort_Click(Index As Integer)

          Dim A As Integer
          Dim b As Integer

10        On Error GoTo optLongOrShort_Click_Error

20        A = 0: b = 1500
30        If optLongOrShort(0) Then A = 1500: b = 0

40        gNormal.ColWidth(0) = A
50        gNormal.ColWidth(1) = b

60        gCodes.ColWidth(0) = A
70        gCodes.ColWidth(1) = b

80        gPlausible.ColWidth(0) = A
90        gPlausible.ColWidth(1) = b

100       gMasks.ColWidth(0) = A
110       gMasks.ColWidth(1) = b

120       Exit Sub

optLongOrShort_Click_Error:

          Dim strES As String
          Dim intEL As Integer

130       intEL = Erl
140       strES = Err.Description
150       LogError "frmBioRanges", "optLongOrShort_Click", intEL, strES


End Sub


Private Sub SSTab1_Click(PreviousTab As Integer)

10        On Error GoTo SSTab1_Click_Error

20        Select Case SSTab1.Tab
          Case 0: FillgCodes
30        Case 1: FillgNormal
40        Case 2: FillgPlausible
50        Case 3: FillgMasks
60        Case 4: FillgSequence
70        End Select

80        Exit Sub

SSTab1_Click_Error:

          Dim strES As String
          Dim intEL As Integer

90        intEL = Erl
100       strES = Err.Description
110       LogError "frmBioRanges", "SSTab1_Click", intEL, strES


End Sub

Private Sub FireDown()

          Dim n As Integer
          Dim s As String
          Dim X As Integer
          Dim VisibleRows As Integer

10        On Error GoTo FireDown_Error

20        If gSequence.Row = gSequence.Rows - 1 Then Exit Sub
30        n = gSequence.Row

40        VisibleRows = gSequence.Height \ gSequence.RowHeight(1) - 1

50        FireCounter = FireCounter + 1
60        If FireCounter > 5 Then
70            tmrDown.Interval = 100
80        End If

90        gSequence.Visible = False

100       s = ""
110       For X = 0 To gSequence.Cols - 1
120           s = s & gSequence.TextMatrix(n, X) & vbTab
130       Next
140       s = Left$(s, Len(s) - 1)

150       gSequence.RemoveItem n
160       If n < gSequence.Rows Then
170           gSequence.AddItem s, n + 1
180           gSequence.Row = n + 1
190       Else
200           gSequence.AddItem s
210           gSequence.Row = gSequence.Rows - 1
220       End If

230       For X = 0 To gSequence.Cols - 1
240           gSequence.Col = X
250           gSequence.CellBackColor = vbYellow
260       Next

270       If Not gSequence.RowIsVisible(gSequence.Row) Or gSequence.Row = gSequence.Rows - 1 Then
280           If gSequence.Row - VisibleRows + 1 > 0 Then
290               gSequence.TopRow = gSequence.Row - VisibleRows + 1
300           End If
310       End If

320       gSequence.Visible = True

330       Exit Sub

FireDown_Error:

          Dim strES As String
          Dim intEL As Integer

340       intEL = Erl
350       strES = Err.Description
360       LogError "frmBioRanges", "FireDown", intEL, strES


End Sub


Private Sub tmrDown_Timer()

10        On Error GoTo tmrDown_Timer_Error

20        FireDown

30        Exit Sub

tmrDown_Timer_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmBioRanges", "tmrDown_Timer", intEL, strES


End Sub

Private Sub tmrUp_Timer()

10        On Error GoTo tmrUp_Timer_Error

20        FireUp

30        Exit Sub

tmrUp_Timer_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmBioRanges", "tmrUp_Timer", intEL, strES


End Sub


Private Sub FireUp()

          Dim n As Integer
          Dim s As String
          Dim X As Integer

10        On Error GoTo FireUp_Error

20        If gSequence.Row = 1 Then Exit Sub

30        FireCounter = FireCounter + 1
40        If FireCounter > 5 Then
50            tmrUp.Interval = 100
60        End If

70        n = gSequence.Row

80        gSequence.Visible = False

90        s = ""
100       For X = 0 To gSequence.Cols - 1
110           s = s & gSequence.TextMatrix(n, X) & vbTab
120       Next
130       s = Left$(s, Len(s) - 1)

140       gSequence.RemoveItem n
150       gSequence.AddItem s, n - 1

160       gSequence.Row = n - 1
170       For X = 0 To gSequence.Cols - 1
180           gSequence.Col = X
190           gSequence.CellBackColor = vbYellow
200       Next

210       If Not gSequence.RowIsVisible(gSequence.Row) Then
220           gSequence.TopRow = gSequence.Row
230       End If

240       gSequence.Visible = True

250       Exit Sub

FireUp_Error:

          Dim strES As String
          Dim intEL As Integer

260       intEL = Erl
270       strES = Err.Description
280       LogError "frmBioRanges", "FireUp", intEL, strES


End Sub

Public Property Let ViewTab(ByVal intNewValue As Integer)

10        On Error GoTo ViewTab_Error

20        pViewTab = intNewValue

30        Exit Property

ViewTab_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmBioRanges", "ViewTab", intEL, strES


End Property

Private Sub SaveCodesNew()
          Dim intN As Integer
          Dim sql As String
          Dim strSampleType As String
          Dim blnP As Boolean
          Dim blnK As Boolean
          Dim blnI As Boolean
          Dim blnE As Boolean
          Dim Discipline As String
          Dim ds As Recordset
          Dim f As Field
          Dim DefIndex As Long
          Dim tb As Recordset
          Dim tempDef As Long

10        On Error GoTo SaveCodesNew_Error

20        If cmbHospital = "" Then
30            cmbHospital = HospName(0)
40        End If
50        If cmbSampleType = "" Then
60            cmbSampleType = "Serum"
70        End If
80        strSampleType = ListCodeFor("ST", cmbSampleType)

90        Discipline = GetDiscipline()

100       intN = gCodes.Row

110       gCodes.Col = 7
120       If gCodes.CellPicture = imgSquareTick.Picture Then blnP = True Else blnP = False

130       gCodes.Col = 8
140       If gCodes.CellPicture = imgSquareTick.Picture Then blnK = True Else blnK = False

150       gCodes.Col = 10
160       If gCodes.CellPicture = imgSquareTick.Picture Then blnI = True Else blnI = False

170       gCodes.Col = 11
180       If gCodes.CellPicture = imgSquareTick.Picture Then blnE = True Else blnE = False

190       sql = "Select * from " & Discipline & "Testdefinitions order by defindex desc"
200       Set tb = New Recordset
210       RecOpenServer 0, tb, sql

220       If Not tb.EOF Then tempDef = tb!DefIndex + 1

230       DefIndex = tempDef + 1

240       sql = "Select * from " & Discipline & "Testdefinitions " & _
                "where LongName = '" & gCodes.TextMatrix(intN, 0) & "' " & _
                "and SampleType = '" & strSampleType & "' " & _
                "and Category = '" & cmbCategory & "' " & _
                "and Hospital = '" & cmbHospital & "' and inuse = 1 and code = '" & gCodes.TextMatrix(intN, 2) & "' order by agefromdays"
250       Set tb = New Recordset
260       RecOpenServer 0, tb, sql
270       Set ds = New Recordset
280       RecOpenServer 0, ds, sql
290       Do While Not tb.EOF
300           tb!InUse = 0
310           tb.Update
320           ds.AddNew
330           For Each f In ds.Fields
340               If UCase(f.Name) <> UCase("rowguid") Then
350                   ds(f.Name) = tb(f.Name)
360               End If
370           Next
380           ds!DefIndex = DefIndex
390           ds.Update
400           DefIndex = DefIndex + 1
410           tb.MoveNext
420       Loop


430       sql = "update " & Discipline & "Testdefinitions set inuse = 1 where defindex > " & tempDef & " and defindex < " & DefIndex & ""
440       Cnxn(0).Execute sql

450       sql = "Update " & Discipline & "TestDefinitions " & _
                "Set Code = '" & gCodes.TextMatrix(intN, 2) & "', " & _
                "BarCode = '" & gCodes.TextMatrix(intN, 3) & "', " & _
                "ImmunoCode = '" & gCodes.TextMatrix(intN, 4) & "', " & _
                "Units = '" & gCodes.TextMatrix(intN, 5) & "', " & _
                "DP = '" & Val(gCodes.TextMatrix(intN, 6)) & "', " & _
                "Printable = " & IIf(blnP, 1, 0) & ", " & _
                "KnownToAnalyser = " & IIf(blnK, 1, 0) & ", " & _
                "Analyser = '" & gCodes.TextMatrix(intN, 9) & "', " & _
                "EOD = " & IIf(blnE, 1, 0) & " " & _
                "where LongName = '" & gCodes.TextMatrix(intN, 0) & "' " & _
                "and SampleType = '" & strSampleType & "' " & _
                "and Category = '" & cmbCategory & "' " & _
                "and Hospital = '" & cmbHospital & "' and inuse = 1 and code = '" & gCodes.TextMatrix(intN, 2) & "'"
460       Cnxn(0).Execute sql

470       Exit Sub

SaveCodesNew_Error:

          Dim strES As String
          Dim intEL As Integer

480       intEL = Erl
490       strES = Err.Description
500       LogError "frmBioRanges", "SaveCodesNew", intEL, strES, sql


End Sub
