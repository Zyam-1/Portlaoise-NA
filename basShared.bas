Attribute VB_Name = "basShared"
Option Explicit

Public TestSys As Boolean
Public intOtherHospitalsInGroup As Long

'number of hospitals
Public Cn As Long

Public Type PhoneLog
    SampleID As Double
    Datetime As Date
    PhonedTo As String
    PhonedBy As String
    Comment As String
    Discipline As String
    Title As String
    PersonName As String
End Type

'User information
Public UserName As String
Public UserCode As String
Public UserMemberOf As String
Public UserPass As String

'logoff information
Public LogOffDelayMin As Long
Public LogOffDelaySecs As Long

'constants
Public Const gVALID = 1
Public Const gNOTVALID = 2
Public Const gPRINTED = 1
Public Const gNOTPRINTED = 2
Public Const gDONTCARE = 0

'Constants for calulating dates & ages
Public Const FORWARD = 1  'used for expiry dates etc
Public Const BACKWARD = 2    'used for DoB etc
Public Const MaxAgeToDays As Long = 43830

'connections
Public Cnxn() As Connection
Public CnxnBB(0 To 0) As Connection
Public CnxnRemote As Connection
Public CnxnRemoteBB As Connection

'hospital information
Public HospName(100) As String
Public Entity(100) As String
Public RemoteEntity(100) As String

Declare Function WinHelp Lib "User" (ByVal hWnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Any) As Long

Public colCounters As New Counters
Public colCountersb As New Counterbs
Public colBIEResults As New BIEResults

Public InterpList(0 To 24) As String

Public Remote As String

Public Const CB_FINDSTRING = &H14C
Public Const CB_FINDSTRINGEXACT = &H158
Public Const CB_SHOWDROPDOWN = &H14F
Public Const CB_GETITEMHEIGHT = &H154
Public Const CB_GETDROPPEDWIDTH = &H15F
Public Const CB_SETDROPPEDWIDTH = &H160
Private Const LB_GETITEMHEIGHT = &H1A1


Public Declare Function MoveWindow Lib "user32" _
                                   (ByVal hWnd As Long, _
                                    ByVal x As Long, ByVal Y As Long, _
                                    ByVal nWidth As Long, _
                                    ByVal nHeight As Long, _
                                    ByVal bRepaint As Long) As Long

Public Declare Function SendMessage Lib "user32" _
                                    Alias "SendMessageA" _
                                    (ByVal hWnd As Long, _
                                     ByVal wMsg As Long, _
                                     ByVal wParam As Long, _
                                     lParam As Any) As Long

Public Enum PrintAlignContants
    AlignLeft = 0
    AlignCenter = 1
    AlignRight = 2
End Enum

Public Enum TabList
    UrineTab = 1
    UrIdentTab = 2
    FaecesTab = 3
    CsTab = 4
    FobTab = 5
    RotaTab = 6
    RsTab = 7
    RsvTab = 8
    FluidsTab = 9
    CDiffTab = 10
    OpTab = 11
    BcTab = 12
    HPyloriTab = 13

End Enum

Public Enum DateType
    Earliest = 0
    Latest = 1
End Enum
Public Enum ConsultantListStatus
    ReleasedToConsultant = 0
    AuthorisedAndReleased = 1
    RevertToLab = 2
    ReleasedToWard = 3
End Enum




Public Sub Archive(ByVal n As Long, _
                   ByVal rs As Recordset, _
                   ByVal Table As String)
'      'archive tables
'      Dim ds As Recordset
'      Dim f As Field
'      Dim sql As String
'
'10    On Error GoTo Archive_Error
'
'20    sql = "SELECT * FROM " & Table & " WHERE 0 = 1"
'30    Set ds = New Recordset
'40    RecOpenServer n, ds, sql
'
'50    ds.AddNew
'60    For Each f In rs.Fields
'70        If UCase(f.Name) <> "ROWGUID" Then
'80            ds(f.Name) = rs(f.Name)
'90        End If
'100   Next
'110   ds.Update
'
'120   Exit Sub
'
'Archive_Error:
'
'      Dim strES As String
'      Dim intEL As Integer
'
'130   intEL = Erl
'140   strES = Err.Description
'150   LogError "basShared", "Archive", intEL, strES, sql

End Sub
Public Sub ArchiveHaem(ByVal SampleID As String)

          Dim sql As String

10        On Error GoTo ArchiveHaem_Error

20        sql = "IF EXISTS (SELECT * FROM HaemResults WHERE SampleID = '" & SampleID & "') " & _
                "  INSERT INTO [ArcHaemResults] " & _
                "  ([SampleID], [AnalysisError], [NegPosError], [PosDiff], [PosMorph], [PosCount], [err_F], [err_R], " & _
                "  [ipMessage], [WBC], [RBC], [Hgb], [Hct], [MCV], [MCH], [MCHC], [Plt], " & _
                "  [LymP], [MonoP], [NeutP], [EosP], [BasP], [LymA], [MonoA], [NeutA], [EosA], [BasA], " & _
                "  [RDWCV], [RDWSD], [PDW], [MPV], [PlCr], " & _
                "  [Valid], [Printed], [Retics], [MonoSpot], [WBCComment], [cESR], [cRetics], [cMonoSpot], [cCoag], " & _
                "  [md0], [md1], [md2], [md3], [md4], [md5], [RunDate], [RunDateTime], [ESR], [PT], [PTControl], [APTT], [APTTControl], [INR], " & _
                "  [FDP], [FIB], [Operator], [FAXed], [Warfarin], [DDimers], [TransmitTime], [Pct], " & _
                "  [WIC], [WOC], [gWB1], [gWB2], [gRBC], [gPlt], [gWIC], [LongError], [cFilm], " & _
                "  [RetA], [RetP], [nrbcA], [nrbcP], [Nopas], [Image], [mi], [an], [ca], [va], " & _
                "  [ho], [he], [ls], [at], [bl], [pp], [nl], [mn], [wp], [ch], [wb], " & _
                "  [lucp], [luca], [cASOT], [tASOT], [tRA], [cRA], [cMalaria], [Malaria], [cSickledex], [Sickledex], " & _
                "  [hdw], [Analyser], [hyp], [li], [mpxi], [mpo], [iG], [lptp], [pclm], [rbcf], " & _
                "  [rbcg], [lplt], [cbad], [RA], [Val1], [Val2], [Val3], [Val4], [Val5], [gRBCH], " & _
                "  [gPLTH], [gPLTF], [gV], [gC], [gS], [DF1], [IRF], [HealthLink], [ValidateTime], " & _
      "  [CD3A], [CD4A], [CD8A], [CD3P], [CD4P], [CD8P], [CD48], [WVF], [rowguid]) "
30        sql = sql & "  SELECT [SampleID], [AnalysisError], [NegPosError], [PosDiff], [PosMorph], [PosCount], [err_F], [err_R], " & _
                "  [ipMessage], [WBC], [RBC], [Hgb], [Hct], [MCV], [MCH], [MCHC], [Plt], " & _
                "  [LymP], [MonoP], [NeutP], [EosP], [BasP], [LymA], [MonoA], [NeutA], [EosA], [BasA], " & _
                "  [RDWCV], [RDWSD], [PDW], [MPV], [PlCr], " & _
                "  [Valid], [Printed], [Retics], [MonoSpot], [WBCComment], [cESR], [cRetics], [cMonoSpot], [cCoag], " & _
                "  [md0], [md1], [md2], [md3], [md4], [md5], [RunDate], [RunDateTime], [ESR], [PT], [PTControl], [APTT], [APTTControl], [INR], " & _
                "  [FDP], [FIB], [Operator], [FAXed], [Warfarin], [DDimers], [TransmitTime], [Pct], " & _
                "  [WIC], [WOC], [gWB1], [gWB2], [gRBC], [gPlt], [gWIC], [LongError], [cFilm], " & _
                "  [RetA], [RetP], [nrbcA], [nrbcP], [Nopas], [Image], [mi], [an], [ca], [va], " & _
                "  [ho], [he], [ls], [at], [bl], [pp], [nl], [mn], [wp], [ch], [wb], " & _
                "  [lucp], [luca], [casot], [tasot], [tra], [cra], [cmalaria], [malaria], [csickledex], [sickledex], " & _
                "  [hdw], [Analyser], [hyp], [li], [mpxi], [mpo], [iG], [lptp], [pclm], [rbcf], " & _
                "  [rbcg], [lplt], [cbad], [RA], [Val1], [Val2], [Val3], [Val4], [Val5], [gRBCH], " & _
                "  [gPLTH], [gPLTF], [gV], [gC], [gS], [DF1], [IRF], [HealthLink], [ValidateTime], " & _
                "  [CD3A], [CD4A], [CD8A], [CD3P], [CD4P], [CD8P], [CD48], [WVF], NEWID() " & _
                "  FROM [HaemResults] WHERE SampleID = '" & SampleID & "'"
40        Cnxn(0).Execute sql

50        Exit Sub

ArchiveHaem_Error:

          Dim strES As String
          Dim intEL As Integer

60        intEL = Erl
70        strES = Err.Description
80        LogError "basShared", "ArchiveHaem", intEL, strES, sql

End Sub


Function AreFlagsPresent(f() As Long) As Boolean

          Dim n As Long

10        On Error GoTo AreFlagsPresent_Error

20        AreFlagsPresent = False

30        For n = 0 To 5
40            If f(n) Then
50                AreFlagsPresent = True
60                Exit Function
70            End If
80        Next

90        Exit Function

AreFlagsPresent_Error:

          Dim strES As String
          Dim intEL As Integer

100       intEL = Erl
110       strES = Err.Description
120       LogError "basShared", "AreFlagsPresent", intEL, strES

End Function

Public Function AreResultsPresent(ByVal Dept As String, ByVal SampleID As String) _
       As Boolean

      'dept = "BGA", "Bio", "Haem", "Coag", "Histo", "Cyto", "Imm", "Ext"

          Dim tb As New Recordset
          Dim sql As String

10        On Error GoTo AreResultsPresent_Error

20        sql = "SELECT count(*) as tot from " & Dept & "Results WHERE " & _
                "SampleID = '" & SampleID & "'"
30        Set tb = New Recordset
40        Set tb = Cnxn(0).Execute(sql)

50        AreResultsPresent = Sgn(tb!Tot)

60        Exit Function

AreResultsPresent_Error:

          Dim strES As String
          Dim intEL As Integer

70        intEL = Erl
80        strES = Err.Description
90        LogError "basShared", "AreResultsPresent", intEL, strES, sql

End Function

Function calcmean(v() As Single) As Single

          Dim mean As Single
          Dim n As Long
          Dim entries As Long

10        On Error GoTo calcmean_Error

20        entries = (UBound(v) - LBound(v)) + 1
30        mean = 0
40        For n = LBound(v) To UBound(v)
50            mean = mean + v(n)
60        Next
70        mean = mean / entries

80        calcmean = mean

90        Exit Function

calcmean_Error:

          Dim strES As String
          Dim intEL As Integer

100       intEL = Erl
110       strES = Err.Description
120       LogError "basShared", "calcmean", intEL, strES

End Function

Function calcmean22(v() As Single)

          Dim n As Long
          Dim Min As Single
          Dim Max As Single
          Dim minpos As Long
          Dim maxpos As Long
          Dim sum As Long

10        On Error GoTo calcmean22_Error

20        Min = 9999
30        Max = 0
40        sum = 0
50        maxpos = 1
60        minpos = 1
70        For n = 1 To 22
80            If v(n) < Min Then
90                Min = v(n)
100               minpos = n
110           ElseIf v(n) > Max Then
120               Max = v(n)
130               maxpos = n
140           End If
150           sum = sum + v(n)
160       Next
170       sum = sum - v(minpos)
180       sum = sum - v(maxpos)

190       calcmean22 = sum / 20

200       Exit Function

calcmean22_Error:

          Dim strES As String
          Dim intEL As Integer

210       intEL = Erl
220       strES = Err.Description
230       LogError "basShared", "calcmean22", intEL, strES

End Function

Function calcsd(v() As Single) As Single

          Dim sumsquared As Single
          Dim squaredsum As Single
          Dim entries As Long
          Dim n As Long

10        On Error GoTo calcsd_Error

20        entries = (UBound(v) - LBound(v)) + 1
30        sumsquared = 0
40        squaredsum = 0
50        For n = LBound(v) To UBound(v)
60            sumsquared = sumsquared + v(n) * v(n)
70            squaredsum = squaredsum + v(n)
80        Next
90        squaredsum = (squaredsum * squaredsum) / entries

100       calcsd = Sqr((sumsquared - squaredsum) / (entries - 1))

110       Exit Function

calcsd_Error:

          Dim strES As String
          Dim intEL As Integer

120       intEL = Erl
130       strES = Err.Description
140       LogError "basShared", "calcsd", intEL, strES

End Function

Public Function GetLastWord(ByVal s As String) As String

          Dim RetVal As String
          Dim w() As String

10        On Error GoTo GetLastWord_Error

20        RetVal = ""

30        If Trim$(s) <> "" Then
40            w = Split(s, " ")
50            RetVal = w(UBound(w))
60        End If

70        GetLastWord = RetVal

80        Exit Function

GetLastWord_Error:

          Dim strES As String
          Dim intEL As Integer

90        intEL = Erl
100       strES = Err.Description
110       LogError "basShared", "GetLastWord", intEL, strES

End Function

Public Sub SetDatesColour(ByVal f As Form)

10        On Error GoTo SetDatesColour_Error

20        With f
30            If CheckDateSequence(.dtSampleDate, .dtRecDate, .dtRunDate, .tSampleTime, .tRecTime) Then
40                .fraDate.ForeColor = vbButtonText
50                .fraDate.Font.Bold = False
60                .lblDate(0).ForeColor = vbButtonText
70                .lblDate(0).Font.Bold = False
80                .lblDate(1).ForeColor = vbButtonText
90                .lblDate(1).Font.Bold = False
100               .lblDateError.Visible = False
110           Else
120               .fraDate.ForeColor = vbRed
130               .fraDate.Font.Bold = True
140               .lblDate(0).ForeColor = vbRed
150               .lblDate(0).Font.Bold = True
160               .lblDate(1).ForeColor = vbRed
170               .lblDate(1).Font.Bold = True
180               .lblDateError.Visible = True
190           End If
200       End With

210       Exit Sub

SetDatesColour_Error:

          Dim strES As String
          Dim intEL As Integer

220       intEL = Erl
230       strES = Err.Description
240       LogError "basShared", "SetDatesColour", intEL, strES

End Sub

Public Function CheckDateSequence(ByRef SampleDate As String, _
                                  ByRef ReceivedDate As String, _
                                  ByRef Rundate As String, _
                                  ByRef SampleTime As String, _
                                  ByRef ReceivedTime As String) _
                                  As Boolean
      'Returns True if ok
          Dim RetVal As Boolean

10        On Error GoTo CheckDateSequence_Error

20        RetVal = True

30        If DateDiff("d", ReceivedDate, Rundate) < 0 Then
40            RetVal = False
50        End If
60        If DateDiff("d", SampleDate, ReceivedDate) < 0 Then
70            RetVal = False
80        End If
90        If IsDate(SampleTime) And IsDate(ReceivedTime) Then
100           If DateDiff("n", SampleDate & " " & SampleTime, ReceivedDate & " " & ReceivedTime) < 0 Then
110               RetVal = False
120           End If
130       End If

140       CheckDateSequence = RetVal

150       Exit Function

CheckDateSequence_Error:

          Dim strES As String
          Dim intEL As Integer

160       intEL = Erl
170       strES = Err.Description
180       LogError "basShared", "CheckDateSequence", intEL, strES

End Function

Public Function ClinName(ByVal Code As String) As String

          Dim tb As New Recordset
          Dim sql As String

10        On Error GoTo ClinName_Error

20        ClinName = "?"
30        Code = UCase$(Trim$(Code))

40        sql = "SELECT Text FROM Clinicians WHERE " & _
                "Code = '" & Code & "'"
50        Set tb = New Recordset
60        RecOpenServer 0, tb, sql
70        If Not tb.EOF Then
80            ClinName = Trim(tb!Text)
90        End If

100       Exit Function

ClinName_Error:

          Dim strES As String
          Dim intEL As Integer


110       intEL = Erl
120       strES = Err.Description
130       LogError "basShared", "ClinName", intEL, strES, sql

End Function

Public Function ChangeComboHeight(cmb As ComboBox, numItemsToDisplay As Integer) As Boolean

          Dim newHeight As Long
          Dim itemHeight As Long

10        itemHeight = SendMessage(cmb.hWnd, CB_GETITEMHEIGHT, 0, ByVal 0)
20        newHeight = itemHeight * (numItemsToDisplay + 2)
30        Call MoveWindow(cmb.hWnd, cmb.Left / 15, cmb.Top / 15, cmb.Width / 15, newHeight, True)

End Function

Public Sub FillGPsClinWard(ByVal f As Form, ByVal HospitalName As String)

          Dim tb As New Recordset
          Dim sql As String
          Dim Hosp As String

10        On Error GoTo FillGPsClinWard_Error

20        Hosp = AddTicks(HospitalName)

30        sql = "SELECT DISTINCT Text, MIN(ListOrder) ListOrder FROM GPs WHERE " & _
                "HospitalCode = (SELECT Code FROM Lists WHERE " & _
                "                ListType = 'HO' " & _
                "                AND Text = '" & Hosp & "') " & _
                "AND InUse = '1' " & _
                "GROUP BY Text " & _
                "ORDER BY ListOrder;"
40        sql = sql & "SELECT DISTINCT Text, MIN(ListOrder) ListOrder FROM Clinicians WHERE " & _
                "HospitalCode = (SELECT Code FROM Lists WHERE " & _
                "                ListType = 'HO' " & _
                "                AND Text = '" & Hosp & "') " & _
                "AND InUse = '1' " & _
                "GROUP BY Text " & _
                "ORDER BY ListOrder;"
50        sql = sql & "SELECT DISTINCT Text, MIN(ListOrder) ListOrder FROM Wards WHERE " & _
                "HospitalCode = (SELECT Code FROM Lists WHERE " & _
                "                ListType = 'HO' " & _
                "                AND Text = '" & Hosp & "') " & _
                "AND InUse = '1' " & _
                "GROUP BY Text " & _
                "ORDER BY ListOrder"

60        Set tb = New Recordset
70        RecOpenServer 0, tb, sql

80        With f.cmbGP
90            .Clear
100           Do While Not tb.EOF
110               .AddItem tb!Text & ""
120               tb.MoveNext
130           Loop
140       End With

150       Set tb = tb.NextRecordset
160       With f.cmbClinician
170           Do While Not tb.EOF
180               .AddItem tb!Text & ""
190               tb.MoveNext
200           Loop
210       End With

220       Set tb = tb.NextRecordset
230       With f.cmbWard
240           .Clear
250           Do While Not tb.EOF
260               .AddItem Trim(tb!Text)
270               tb.MoveNext
280           Loop
290       End With

300       Exit Sub

FillGPsClinWard_Error:

          Dim strES As String
          Dim intEL As Integer

310       intEL = Erl
320       strES = Err.Description
330       LogError "basShared", "FillGPsClinWard", intEL, strES, sql

End Sub

Public Sub FillGPsWard(ByVal f As Form, ByVal HospitalName As String)

          Dim strHospitalCode As String
          Dim tb As New Recordset
          Dim sql As String

10        On Error GoTo FillGPsWard_Error

20        strHospitalCode = ListCodeFor("HO", HospitalName)

30        sql = "SELECT DISTINCT Text,ListOrder from GPs WHERE " & _
                "HospitalCode = '" & strHospitalCode & "' " & _
                "AND InUse = '1' " & _
                "ORDER BY ListOrder;" & _
                "SELECT DISTINCT Text, ListOrder FROM Wards WHERE " & _
                "HospitalCode = '" & strHospitalCode & "' " & _
                "AND InUse = '1' " & _
                "ORDER BY ListOrder"

40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql

60        With f.cmbGP
70            .Clear
80            Do While Not tb.EOF
90                If .List(.ListCount - 1) <> tb!Text & "" Then
100                   .AddItem tb!Text & ""
110               End If
120               tb.MoveNext
130           Loop
140       End With

150       Set tb = tb.NextRecordset

160       With f.cmbWard
170           .Clear
180           Do While Not tb.EOF
190               .AddItem Trim(tb!Text)
200               tb.MoveNext
210           Loop
220       End With

230       Exit Sub

FillGPsWard_Error:

          Dim strES As String
          Dim intEL As Integer

240       intEL = Erl
250       strES = Err.Description
260       LogError "basShared", "FillGPsWard", intEL, strES, sql

End Sub

Public Sub FillMRU(ByVal cMRU As ComboBox)

          Dim sql As String
          Dim tb As New Recordset

10        On Error GoTo FillMRU_Error

20        sql = "SELECT top 10 * from MRU WHERE " & _
                "UserCode = '" & UserCode & "' " & _
                "Order by DateTime desc"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql

50        With cMRU
60            .Clear
70            Do While Not tb.EOF
80                .AddItem Trim$(tb!SampleID & "")
90                tb.MoveNext
100           Loop
110           If .ListCount > 0 Then
120               .Text = ""
130           End If
140       End With

150       Exit Sub

FillMRU_Error:

          Dim strES As String
          Dim intEL As Integer

160       intEL = Erl
170       strES = Err.Description
180       LogError "basShared", "FillMRU", intEL, strES, sql

End Sub


Public Function GetWard(ByVal CodeOrText As String, _
                        ByVal HospCode As String) As String

          Dim tb As New Recordset
          Dim sql As String
          Dim s As String

10        On Error GoTo GetWard_Error

20        s = AddTicks(Trim$(CodeOrText))

30        sql = "SELECT [Text] from Wards WHERE " & _
                "HospitalCode = '" & HospCode & "' " & _
                "and Inuse = 1 " & _
                "and (Code = '" & s & "' or [Text] = '" & s & "')"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        If Not tb.EOF Then
70            GetWard = tb!Text & ""
80        Else
90            GetWard = CodeOrText
100       End If

110       Exit Function

GetWard_Error:

          Dim strES As String
          Dim intEL As Integer

120       intEL = Erl
130       strES = Err.Description
140       LogError "basShared", "GetWard", intEL, strES, sql

End Function

Public Function ListCodeFor(ByVal ListType As String, ByVal Text As String) As String

          Dim tb As New Recordset
          Dim sql As String

10        On Error GoTo ListCodeFor_Error

20        ListCodeFor = ""

30        sql = "SELECT Code FROM Lists WHERE " & _
                "ListType = '" & ListType & "' " & _
                "AND Text = '" & AddTicks(Text) & "'"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        If Not tb.EOF Then
70            ListCodeFor = Trim(tb!Code)
80        End If

90        Exit Function

ListCodeFor_Error:

          Dim strES As String
          Dim intEL As Integer

100       intEL = Erl
110       strES = Err.Description
120       LogError "basShared", "ListCodeFor", intEL, strES, sql

End Function

Public Function ListText(ByVal ListType As String, ByVal Code As String) As String

          Dim tb As New Recordset
          Dim sql As String

10        On Error GoTo ListText_Error

20        ListText = ""
30        Code = Trim$(Code)

40        sql = "SELECT Text FROM Lists WHERE " & _
                "ListType = '" & ListType & "' " & _
                "AND Code = '" & Code & "'"
50        Set tb = New Recordset
60        RecOpenServer 0, tb, sql
70        If Not tb.EOF Then
80            ListText = Trim(tb!Text)
90        End If

100       Exit Function

ListText_Error:

          Dim strES As String
          Dim intEL As Integer

110       intEL = Erl
120       strES = Err.Description
130       LogError "basShared", "ListText", intEL, strES, sql

End Function


Public Sub LogTimeOfPrinting(ByVal SampleID As String, _
                             ByVal Dept As String)
      'Dept is "H", "B" or "C" or "D"

          Dim sql As String

10        On Error GoTo LogTimeOfPrinting_Error

20        Select Case UCase$(Left$(Dept, 1))
          Case "H": Dept = "DateTimeHaemPrinted"
30        Case "B": Dept = "DateTimeBioPrinted"
40        Case "C": Dept = "DateTimeCoagPrinted"
50        Case "D": Dept = "DateTimeDemographics"
60        Case Else: Exit Sub
70        End Select

80        sql = "UPDATE Demographics set " & _
                Dept & " = '" & Format$(Now, "dd/mmm/yyyy hh:mm:ss") & "' " & _
                "WHERE SampleID = '" & SampleID & "' " & _
                "and " & Dept & " is null"
90        Cnxn(0).Execute sql

100       Exit Sub

LogTimeOfPrinting_Error:

          Dim strES As String
          Dim intEL As Integer

110       intEL = Erl
120       strES = Err.Description
130       LogError "basShared", "LogTimeOfPrinting", intEL, strES, sql

End Sub

Function Maximum(g() As Single) As Single

          Dim n As Long
          Dim Max As Single

10        On Error GoTo Maximum_Error

20        Max = 0
30        For n = LBound(g) To UBound(g)
40            If g(n) > Max Then Max = g(n)
50        Next

60        Maximum = Max

70        Exit Function

Maximum_Error:

          Dim strES As String
          Dim intEL As Integer

80        intEL = Erl
90        strES = Err.Description
100       LogError "basShared", "maximum", intEL, strES

End Function

Function minimum(g() As Single) As Single

          Dim n As Long
          Dim Min As Single

10        On Error GoTo minimum_Error

20        Min = 999
30        For n = LBound(g) To UBound(g)
40            If g(n) < Min Then Min = g(n)
50        Next

60        minimum = Min

70        Exit Function

minimum_Error:

          Dim strES As String
          Dim intEL As Integer

80        intEL = Erl
90        strES = Err.Description
100       LogError "basShared", "minimum", intEL, strES

End Function

Public Sub NameLostFocus(ByRef strName As String, _
                         ByRef strSex As String)

      Dim ForeName As String
      Dim tb As New Recordset
      Dim sql As String

10    On Error GoTo NameLostFocus_Error




20    strName = Replace(strName, ",", "")

30    strName = initial2upper(strName)

40    ForeName = ParseForeName(strName)

50    If ForeName = "" Then
60        strSex = ""
70        Exit Sub
80    End If



90    If GetOptionSetting("EnableSexNamesLookup", 0) = 0 Then Exit Sub


100   sql = "SELECT * from SexNames WHERE " & _
            "Name = '" & AddTicks(ForeName) & "'"
110   Set tb = New Recordset
120   RecOpenServer 0, tb, sql
130   If tb.EOF Then
140       If strSex <> "" Then
150           sql = "Insert Into SexNames (Name, Sex) Values ('@Name', '@Sex')"
160           sql = Replace(sql, "@Name", ForeName)
170           sql = Replace(sql, "@Sex", Left$(strSex, 1))
180           Cnxn(0).Execute sql
190       End If
200   Else
210       Select Case UCase(tb!sex & "")
              Case "M": strSex = "Male"
220           Case "F": strSex = "Female"
230       End Select
240   End If

250   Exit Sub

NameLostFocus_Error:

      Dim strES As String
      Dim intEL As Integer

260   intEL = Erl
270   strES = Err.Description
280   LogError "basShared", "NameLostFocus", intEL, strES, sql

End Sub

Public Function QueryKnown(ByVal ClinOrGP As String, _
                           ByVal CodeOrText As String, _
                           Optional ByVal Hospital As String) _
                           As String
      'Returns either "" = not known
      '        or CodeOrText = known

          Dim HospCode As String
          Dim Original As String
          Dim tb As New Recordset
          Dim sql As String

10        On Error GoTo QueryKnown_Error

20        QueryKnown = ""
30        Original = CodeOrText

40        CodeOrText = Trim$(UCase$(CodeOrText))
50        If CodeOrText = "" Then Exit Function

60        If ClinOrGP = "GP" Then
70            If Not IsMissing(Hospital) Then
80                HospCode = ListCodeFor("HO", Hospital)
90            Else
100               HospCode = "M"
110           End If
120           sql = "SELECT * from GPs where hospitalcode = '" & HospCode & "' And InUse = '1'"
130           Set tb = New Recordset
140           RecOpenClient 0, tb, sql
150           Do While Not tb.EOF
160               If Trim$(UCase$(tb!Code)) = CodeOrText Then
170                   QueryKnown = tb!Text
180                   Exit Function
190               ElseIf Trim$(UCase$(tb!Text)) = CodeOrText Then
200                   QueryKnown = tb!Text & ""
210                   Exit Function
220               End If
230               tb.MoveNext
240           Loop
250       ElseIf ClinOrGP = "Clin" Then
260           If Not IsMissing(Hospital) Then
270               HospCode = ListCodeFor("HO", Hospital)
280           Else
290               HospCode = "M"
300           End If
310           sql = "SELECT * from Clinicians WHERE " & _
                    "HospitalCode = '" & HospCode & "' And InUse = '1'"
320           Set tb = New Recordset
330           RecOpenServer 0, tb, sql
340           Do While Not tb.EOF
350               If Trim$(UCase$(tb!Code)) = CodeOrText Then
360                   QueryKnown = tb!Text
370                   Exit Function
380               ElseIf Trim$(UCase$(tb!Text)) = CodeOrText Then
390                   QueryKnown = tb!Text & ""
400                   Exit Function
410               End If
420               tb.MoveNext
430           Loop
440       End If

450       Exit Function

QueryKnown_Error:

          Dim strES As String
          Dim intEL As Integer

460       intEL = Erl
470       strES = Err.Description
480       LogError "basShared", "QueryKnown", intEL, strES, sql

End Function

Public Sub Set_Font(ByVal f As Form)

          Dim ax As Control

10        On Error GoTo Set_Font_Error

20        If SysOptFont(0) = "" Then Exit Sub

30        f.Font = SysOptFont(0)

40        For Each ax In f.Controls
50            If TypeOf ax Is TextBox Or TypeOf ax Is Label _
                 Or TypeOf ax Is MSFlexGrid Or TypeOf ax Is Frame _
                 Or TypeOf ax Is ComboBox Or TypeOf ax Is ListBox _
                 Or TypeOf ax Is CheckBox Or TypeOf ax Is OptionButton _
                 Or TypeOf ax Is SSPanel Or TypeOf ax Is CommandButton _
                 Or TypeOf ax Is SSTab Then
60                ax.Font = SysOptFont(0)
70            End If
80        Next

90        Exit Sub

Set_Font_Error:

          Dim strES As String
          Dim intEL As Integer

100       intEL = Erl
110       strES = Err.Description
120       LogError "basShared", "Set_Font", intEL, strES

End Sub

Public Sub SexLostFocus(ByVal tSex As TextBox, ByVal tName As TextBox)

          Dim sql As String
          Dim tb As New Recordset
          Dim ForeName As String

10        On Error GoTo SexLostFocus_Error

20        If Trim$(tSex) = "" Or GetOptionSetting("EnableSexNamesLookup", 0) = 0 Then Exit Sub
30        If UCase$(Left$(tSex, 1)) <> "F" And UCase$(Left$(tSex, 1)) <> "M" Then Exit Sub

40        ForeName = ParseForeName(tName)
50        If ForeName = "" Then Exit Sub

60        sql = "SELECT * from SexNames WHERE " & _
                "Name = '" & AddTicks(ForeName) & "'"
70        Set tb = New Recordset
80        RecOpenClient 0, tb, sql
90        If tb.EOF Then
100           tb.AddNew
110       End If

120       tb!Name = ForeName
130       tb!sex = UCase$(Left$(tSex, 1))
140       tb.Update

150       Exit Sub

SexLostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

160       intEL = Erl
170       strES = Err.Description
180       LogError "basShared", "SexLostFocus", intEL, strES, sql

End Sub

Function TechNameFor(ByVal OperCode As String) As String

          Dim sql As String
          Dim tb As New Recordset

10        On Error GoTo TechNameFor_Error

20        TechNameFor = ""

30        sql = "SELECT Name FROM Users WHERE " & _
                "Code = '" & OperCode & "'"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        If Not tb.EOF Then
70            TechNameFor = initial2upper(tb!Name)
80        End If

90        Exit Function

TechNameFor_Error:

          Dim strES As String
          Dim intEL As Integer

100       intEL = Erl
110       strES = Err.Description
120       LogError "basShared", "TechNameFor", intEL, strES, sql

End Function

Public Sub UPDATEMRU(ByVal SampleID As String, _
                     ByVal cMRU As ComboBox)

          Dim sql As String
          Dim tb As New Recordset
          Dim n As Long
          Dim Found As Boolean
          Dim NewMRU(0 To 9, 0 To 1) As String
          '(x,0) SampleID
          '(x,1) DateTime

10        On Error GoTo UPDATEMRU_Error

20        sql = "SELECT top 10 * from MRU WHERE " & _
                "UserCode = '" & UserCode & "' " & _
                "Order by DateTime desc"
30        Set tb = New Recordset
40        RecOpenClient 0, tb, sql

50        n = -1
60        Do While Not tb.EOF
70            n = n + 1
80            NewMRU(n, 0) = Trim$(tb!SampleID)
90            NewMRU(n, 1) = tb!Datetime
100           tb.MoveNext
110       Loop

120       Found = False
130       For n = 0 To 9
140           If SampleID = NewMRU(n, 0) Then
150               sql = "UPDATE MRU " & _
                        "Set DateTime = '" & Format$(Now, "dd/mmm/yyyy hh:mm:ss") & "' " & _
                        "WHERE SampleID = '" & SampleID & "' " & _
                        "and UserCode = '" & UserCode & "'"
160               Cnxn(0).Execute sql
170               Found = True
180               Exit For
190           End If
200       Next

210       If Not Found Then
220           sql = "DELETE from MRU WHERE " & _
                    "UserCode = '" & UserCode & "'"
230           Cnxn(0).Execute sql
240           For n = 0 To 8
250               If NewMRU(n, 0) <> "" Then
260                   sql = "INSERT into MRU " & _
                            "(SampleID, DateTime, UserCode ) VALUES " & _
                            "('" & NewMRU(n, 0) & "', " & _
                            "'" & Format$(NewMRU(n, 1), "dd/mmm/yyyy hh:mm:ss") & "', " & _
                            "'" & UserCode & "')"
270                   Cnxn(0).Execute sql
280               End If
290           Next
300           sql = "INSERT into MRU " & _
                    "(SampleID, DateTime, UserCode ) VALUES " & _
                    "('" & SampleID & "', " & _
                    "'" & Format$(Now, "dd/mmm/yyyy hh:mm:ss") & "', " & _
                    "'" & UserCode & "')"
310           Cnxn(0).Execute sql
320       End If

330       sql = "SELECT top 10 * from MRU WHERE " & _
                "UserCode = '" & UserCode & "' " & _
                "Order by DateTime desc"
340       Set tb = New Recordset
350       RecOpenClient 0, tb, sql

360       With cMRU
370           .Clear
380           Do While Not tb.EOF
390               .AddItem Trim$(tb!SampleID & "")
400               tb.MoveNext
410           Loop
420           If .ListCount > 0 Then
430               .Text = ""
440           End If
450       End With

460       Exit Sub

UPDATEMRU_Error:

          Dim strES As String
          Dim intEL As Integer

470       intEL = Erl
480       strES = Err.Description
490       LogError "basShared", "UPDATEMRU", intEL, strES, sql

End Sub

Public Function AddressOfGP(ByVal GPName As String) As String

          Dim sql As String
          Dim tb As Recordset
          Dim RetVal As String

10        On Error GoTo AddressOfGP_Error

20        If Trim$(GPName) = "" Or UCase$(GPName) = "GP" Then
30            RetVal = ""
40        Else
50            sql = "SELECT Addr0 FROM GPs WHERE " & _
                    "Text = '" & AddTicks(GPName) & "' AND InUse = '1'"
60            Set tb = New Recordset
70            RecOpenServer 0, tb, sql
80            If Not tb.EOF Then
90                RetVal = tb!Addr0 & ""
100           Else
110               RetVal = ""
120           End If
130       End If

140       AddressOfGP = RetVal

150       Exit Function

AddressOfGP_Error:

          Dim strES As String
          Dim intEL As Integer

160       intEL = Erl
170       strES = Err.Description
180       LogError "basShared", "AddressOfGP", intEL, strES, sql

End Function

Public Function IsFaxable(ByVal Source As String, _
                          ByVal Specific As String) As String

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo IsFaxable_Error

20        sql = "SELECT Fax FROM " & Source & " WHERE " & _
                "Text = '" & AddTicks(Specific) & "' " & _
                "AND Fax <> '' AND Fax IS NOT NULL"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql

50        If Not tb.EOF Then
60            IsFaxable = tb!FAX & ""
70        Else
80            IsFaxable = ""
90        End If

100       Exit Function

IsFaxable_Error:

          Dim strES As String
          Dim intEL As Integer

110       intEL = Erl
120       strES = Err.Description
130       LogError "basShared", "IsFaxable", intEL, strES, sql

End Function

Public Function getBatchEntryOpenStatus(ByVal strOption As String) As String

          Dim tb As New Recordset
          Dim sql As String

10        On Error GoTo getBatchEntryOpenStatus_Error

20        sql = "SELECT * from Options WHERE Description = '" & strOption & "'"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        If Not tb.EOF Then
60            If Trim$(tb!Contents) <> "" Then
70                getBatchEntryOpenStatus = Trim$(tb!UserName & "")
80            End If
90        Else
100           getBatchEntryOpenStatus = ""
110       End If

120       Exit Function

getBatchEntryOpenStatus_Error:

          Dim strES As String
          Dim intEL As Integer

130       intEL = Erl
140       strES = Err.Description
150       LogError "basShared", "getBatchOccultOpenStatus", intEL, strES, sql

End Function

Public Sub MarkBatchEntryOpen4Use(ByVal strOption As String)

          Dim sql As String
          Dim tb As Recordset

10        On Error GoTo MarkBatchEntryOpen4Use_Error

20        sql = "SELECT * FROM Options WHERE " & _
                "Description = '" & strOption & "'"

30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        If tb.EOF Then
60            tb.AddNew
70        End If
80        tb!Description = strOption
90        tb!Contents = vbGetComputerName()
100       tb!UserName = UserName
110       tb.Update

120       Exit Sub

MarkBatchEntryOpen4Use_Error:

          Dim strES As String
          Dim intEL As Integer

130       intEL = Erl
140       strES = Err.Description
150       LogError "frmBatchOccult", "MarkBatchEntryOccultBloodOpen4Use", intEL, strES, sql

End Sub

Public Function delBatchEntryOpenStatus(ByVal strOption As String) As String

          Dim sql As String

10        On Error GoTo delBatchEntryOpenStatus_Error

20        sql = "DELETE FROM Options WHERE (Description = '" & strOption & "')"
30        Cnxn(0).Execute sql

40        Exit Function

delBatchEntryOpenStatus_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "basShared", "delBatchEntryOpenStatus", intEL, strES, sql

End Function

Public Function PrintTextRTB(rtb As RichTextBox, ByVal Text As String, _
                             Optional FontSize As Integer = 9, Optional FontBold As Boolean = False, _
                             Optional FontItalic As Boolean = False, Optional FontUnderLine As Boolean = False, _
                             Optional FontColor As ColorConstants = vbBlack)


      '---------------------------------------------------------------------------------------
      ' Procedure : PrintText
      ' DateTime  : 05/06/2008 11:40
      ' Author    : Babar Shahzad
      ' Note      : Printer object needs to be set first before calling this function.
      '             Portrait mode (width X height) = 11800 X 16500
      '---------------------------------------------------------------------------------------
10        With rtb

20            .SelFontSize = FontSize
30            .SelBold = FontBold
40            .SelItalic = FontItalic
50            .SelUnderline = FontUnderLine
60            .SelColor = FontColor
70            .SelText = Text
80        End With
End Function

Public Function PrintText(ByVal Text As String, _
                          Optional FontSize As Integer = 9, Optional FontBold As Boolean = False, _
                          Optional FontItalic As Boolean = False, Optional FontUnderLine As Boolean = False, _
                          Optional FontColor As ColorConstants = vbBlack, _
                          Optional EnterCrLf As Boolean = False)


      '---------------------------------------------------------------------------------------
      ' Procedure : PrintText
      ' DateTime  : 05/06/2008 11:40
      ' Author    : Babar Shahzad
      ' Note      : Printer object needs to be set first before calling this function.
      '             Portrait mode (width X height) = 11800 X 16500
      '---------------------------------------------------------------------------------------
10        With Printer
20            .Font.Size = FontSize
30            .Font.Bold = FontBold
40            .Font.Italic = FontItalic
50            .Font.Underline = FontUnderLine
60            .ForeColor = FontColor
70            If EnterCrLf Then
80                Printer.Print Text
90            Else
100               Printer.Print Text;
110           End If
120       End With
End Function

Public Function FormatString(strDestString As String, _
                             intNumChars As Integer, _
                             Optional strSeperator As String = "", _
                             Optional intAlign As PrintAlignContants = AlignLeft) As String

      '**************intAlign = 0 --> Left Align
      '**************intAlign = 1 --> Center Align
      '**************intAlign = 2 --> Right Align
          Dim intPadding As Integer
10        intPadding = 0

20        If Len(strDestString) > intNumChars Then
30            FormatString = Mid(strDestString, 1, intNumChars) & strSeperator
40        ElseIf Len(strDestString) < intNumChars Then
              Dim i As Integer
              Dim intStringLength As String
50            intStringLength = Len(strDestString)
60            intPadding = intNumChars - intStringLength

70            If intAlign = PrintAlignContants.AlignLeft Then
80                strDestString = strDestString & String(intPadding, " ")  '& " "
90            ElseIf intAlign = PrintAlignContants.AlignCenter Then
100               If (intPadding Mod 2) = 0 Then
110                   strDestString = String(intPadding / 2, " ") & strDestString & String(intPadding / 2, " ")
120               Else
130                   strDestString = String((intPadding - 1) / 2, " ") & strDestString & String((intPadding - 1) / 2 + 1, " ")
140               End If
150           ElseIf intAlign = PrintAlignContants.AlignRight Then
160               strDestString = String(intPadding, " ") & strDestString
170           End If

180           strDestString = strDestString & strSeperator
190           FormatString = strDestString
200       Else
210           strDestString = strDestString & strSeperator
220           FormatString = strDestString
230       End If



End Function

Public Function FixComboWidth(Combo As ComboBox) As Boolean

          Dim i As Integer
          Dim ScrollWidth As Long

10        With Combo
20            For i = 0 To .ListCount
30                If .Parent.TextWidth(.List(i)) > ScrollWidth Then
40                    ScrollWidth = .Parent.TextWidth(.List(i))
50                End If
60            Next i
70            FixComboWidth = SendMessage(.hWnd, CB_SETDROPPEDWIDTH, _
                                          ScrollWidth / 15 + 30, 0) > 0

80        End With

End Function

Public Sub FixListHeight(lst As ListBox)
          Dim GetHeightOfListItem As Long
10        If lst.ListCount = 0 Then Exit Sub
20        GetHeightOfListItem = SendMessage(lst.hWnd, LB_GETITEMHEIGHT, 1, 0)
30        If GetHeightOfListItem > 0 Then
40            lst.Height = GetHeightOfListItem * (lst.ListCount + 1) * 15
50        End If

End Sub


Public Sub LoadListGeneric(cmb As ComboBox, ListType As String)

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo LoadListGeneric_Error

20        cmb.Clear

30        sql = "Select Text From Lists Where ListType = '" & ListType & "' Order By ListOrder"
40        Set tb = New Recordset
50        RecOpenClient 0, tb, sql
60        If Not tb.EOF Then

70            cmb.AddItem ""
80            While Not tb.EOF
90                cmb.AddItem tb!Text
100               tb.MoveNext
110           Wend
120       End If


130       Exit Sub

LoadListGeneric_Error:

          Dim strES As String
          Dim intEL As Integer

140       intEL = Erl
150       strES = Err.Description
160       LogError "basShared", "LoadListGeneric", intEL, strES, sql


End Sub

Public Function getBioAnalyserCode(ByVal strAnalyserName As String) As String
          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo getBioAnalyserCode_Error

20        sql = "Select Code From Lists Where ListType = 'BioAnalysers' and Text = '" & strAnalyserName & "'"
30        Set tb = New Recordset
40        RecOpenClient 0, tb, sql
50        If Not tb.EOF Then
60            getBioAnalyserCode = tb!Code & ""
70        End If

80        Exit Function

getBioAnalyserCode_Error:

          Dim strES As String
          Dim intEL As Integer

90        intEL = Erl
100       strES = Err.Description
110       LogError "basShared", "getBioAnalyserCode", intEL, strES, sql



End Function

Public Function TabExistsForSite(ByVal Site As String, ByVal TabItem As TabList) As Boolean

          Dim tb As Recordset
          Dim sql As String
          Dim ColName As String

10        On Error GoTo TabExistsForSite_Error

20        If Trim(Site) = "" Then
30            TabExistsForSite = False
40            Exit Function
50        End If

60        ColName = ""


70        Select Case TabItem
          Case TabList.BcTab:
80            ColName = "BC"
90        Case TabList.CDiffTab:
100           ColName = "cDiff"
110       Case TabList.CsTab:
120           ColName = "CS"
130       Case TabList.FaecesTab:
140           ColName = "Faeces"
150       Case TabList.FluidsTab:
160           ColName = "Fluids"
170       Case TabList.FobTab:
180           ColName = "FOB"
190       Case TabList.HPyloriTab:
200           ColName = "HPylori"
210       Case TabList.OpTab:
220           ColName = "OP"
230       Case TabList.RotaTab:
240           ColName = "Rota"
250       Case TabList.RsTab:
260           ColName = "RS"
270       Case TabList.RsvTab:
280           ColName = "RSV"
290       Case TabList.UrIdentTab:
300           ColName = "UrIdent"
310       Case TabList.UrineTab:
320           ColName = "Urine"
330       End Select




340       sql = "Select %colname From MicroSetup Where Site = '%site'"
350       sql = Replace(sql, "%colname", ColName)
360       sql = Replace(sql, "%site", Site)

370       Set tb = New Recordset
380       RecOpenClient 0, tb, sql
390       If tb.EOF Then
400           TabExistsForSite = False
410       Else
420           TabExistsForSite = tb(ColName)
430       End If

440       Exit Function

TabExistsForSite_Error:

          Dim strES As String
          Dim intEL As Integer

450       intEL = Erl
460       strES = Err.Description
470       LogError "basShared", "TabExistsForSite", intEL, strES, sql

End Function

Public Sub CycleControlValue(ByRef ListOfItems() As String, _
                             ByRef Ctrl As Control)

10        On Error GoTo CycleControlValue_Error

          Dim n As Integer
          Dim x As Integer
          Dim CtrlValue As String

20        If TypeOf Ctrl Is TextBox Then
30            CtrlValue = Ctrl.Text
40        ElseIf TypeOf Ctrl Is Label Then
50            CtrlValue = Ctrl.Caption
60        ElseIf TypeOf Ctrl Is MSFlexGrid Then
70            CtrlValue = Ctrl.TextMatrix(Ctrl.Row, Ctrl.Col)
80        Else
90            CtrlValue = ""
100       End If



110       For n = 0 To UBound(ListOfItems)
120           If UCase(CtrlValue) = UCase(ListOfItems(n)) Then
130               If n = UBound(ListOfItems) Then
140                   x = 0
150               Else
160                   x = n + 1
170               End If
180               CtrlValue = ListOfItems(x)
190               Exit For
200           End If
210       Next

220       If TypeOf Ctrl Is TextBox Then
230           Ctrl.Text = CtrlValue
240       ElseIf TypeOf Ctrl Is Label Then
250           Ctrl.Caption = CtrlValue
260       ElseIf TypeOf Ctrl Is MSFlexGrid Then
270           Ctrl.TextMatrix(Ctrl.Row, Ctrl.Col) = CtrlValue

280       End If


290       Exit Sub

CycleControlValue_Error:

          Dim strES As String
          Dim intEL As Integer

300       intEL = Erl
310       strES = Err.Description
320       LogError "basShared", "CycleControlValue", intEL, strES

End Sub

Public Function GetRunDateTime(SampleID As String, Disc As String, _
                               Optional RunDateType As DateType = DateType.Earliest) As String

          Dim tb As Recordset
          Dim sql As String
          Dim ColumnName As String

10        On Error GoTo GetRunDateTime_Error

20        If Disc = "Haem" Then
30            ColumnName = "RunDateTime"
40        Else
50            ColumnName = "RunTime"
60        End If

70        sql = "SELECT Top 1 " & ColumnName & " FROM " & Disc & "Results WHERE SampleID = '" & SampleID & "' " & _
                "AND COALESCE(" & ColumnName & ", '') <> '' "
80        Select Case RunDateType
          Case 0:
90            sql = sql & "ORDER BY " & ColumnName & " Asc"
100       Case 1:
110           sql = sql & "ORDER BY " & ColumnName & " Desc"
120       End Select
130       Set tb = New Recordset
140       RecOpenServer 0, tb, sql
150       If tb.EOF Then
160           GetRunDateTime = ""
170       Else
180           If Disc = "Haem" Then
190               GetRunDateTime = tb!RunDateTime & ""
200           Else
210               GetRunDateTime = tb!RunTime & ""
220           End If
230       End If



240       Exit Function

GetRunDateTime_Error:

          Dim strES As String
          Dim intEL As Integer

250       intEL = Erl
260       strES = Err.Description
270       LogError "basShared", "GetRunDateTime", intEL, strES, sql

End Function

Public Function RemoveLeadingCrLf(TextString As String) As String

10        On Error GoTo RemoveLeadingCrLf_Error

          Dim i As Integer

20        If TextString = "" Or Len(TextString) < 2 Then
30            RemoveLeadingCrLf = TextString
40            Exit Function
50        End If

60        i = 1

70        While i < Len(TextString)
80            If Left(TextString, 2) = vbCrLf Then
90                TextString = Right(TextString, Len(TextString) - 2)
100               i = i + 2
110           Else
120               i = Len(TextString)
130           End If

140       Wend

150       RemoveLeadingCrLf = TextString

160       Exit Function

RemoveLeadingCrLf_Error:

          Dim strES As String
          Dim intEL As Integer

170       intEL = Erl
180       strES = Err.Description
190       LogError "basShared", "RemoveLeadingCrLf", intEL, strES

End Function


Public Sub FillGenericList(ByRef cmb As ComboBox, ByVal ListType As String, Optional ByVal AddEmpty As Boolean = False)

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo FillGenericList_Error

20        cmb.Clear
30        sql = "SELECT * FROM Lists WHERE " & _
                "ListType = '" & ListType & "' " & _
                "AND InUse = 1 " & _
                "ORDER BY ListOrder"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        If AddEmpty Then cmb.AddItem ""
70        Do While Not tb.EOF
80            cmb.AddItem tb!Text & ""
90            tb.MoveNext
100       Loop

110       Exit Sub

FillGenericList_Error:

          Dim strES As String
          Dim intEL As Integer

120       intEL = Erl
130       strES = Err.Description
140       LogError "Shared", "FillGenericList", intEL, strES, sql

End Sub

Public Function EditGrid(ByVal g As MSFlexGrid, _
                         ByVal KeyCode As Integer, _
                         ByVal Shift As Integer) _
                         As Boolean

      'returns true if grid changed

          Dim ShiftDown As Boolean

10        EditGrid = False

20        If g.Row < g.FixedRows Then
30            Exit Function
40        ElseIf g.Col < g.FixedCols Then
50            Exit Function
60        End If
70        ShiftDown = (Shift And vbShiftMask) > 0


80        Select Case KeyCode
          Case vbKeyA To vbKeyZ:
90            If ShiftDown Then
100               g = g & chr(KeyCode)
110               EditGrid = True
120           Else
130               g = g & chr(KeyCode + 32)
140               EditGrid = True
150           End If

160       Case vbKey0 To vbKey9:
170           g = g & chr(KeyCode)
180           EditGrid = True

190       Case vbKeyBack:
200           If Len(g) > 0 Then
210               g = Left$(g, Len(g) - 1)
220               EditGrid = True
230           End If

240       Case &HBE, vbKeyDecimal:
250           g = g & "."
260           EditGrid = True

270       Case vbKeySpace:
280           g = g & " "
290           EditGrid = True

300       Case vbKeyNumpad0 To vbKeyNumpad9:
310       Case vbKeyDelete:
320       Case vbKeyLeft:
330       Case vbKeyRight:
340       Case vbKeyUp:
350       Case vbKeyDown:
360       Case vbKeyTab:
370       End Select

End Function

Public Function AutoComplete(cmbCombo As ComboBox, sKeyAscii As Integer, Optional bUpperCase As Boolean = True) As Integer
          Dim lngFind As Long, intPos As Integer, intLength As Integer
          Dim tStr As String
10        On Error GoTo AutoComplete_Error

20        If sKeyAscii = 8 Or sKeyAscii = 13 Then
30            AutoComplete = sKeyAscii
40            Exit Function
50        End If

60        With cmbCombo
70            If sKeyAscii = 8 Then
80                If .SelStart = 0 Then Exit Function
90                .SelStart = .SelStart - 1
100               .SelLength = 32000
110               .SelText = ""
120           Else
130               intPos = .SelStart    '// save intial cursor position
140               tStr = .Text    '// save string
150               If bUpperCase = True Then
160                   .SelText = UCase(chr(sKeyAscii))    '// change string. (uppercase only)
170               Else
180                   .SelText = chr(sKeyAscii)    '// change string. (leave case alone)
190               End If
200           End If

210           lngFind = SendMessage(.hWnd, CB_FINDSTRING, 0, ByVal .Text)    '// Find string in combobox
220           If lngFind = -1 Then    '// if string not found
230               .Text = tStr    '// set old string (used for boxes that require charachter monitoring
240               .SelStart = intPos    '// set cursor position
250               .SelLength = (Len(.Text) - intPos)    '// set selected length
260               AutoComplete = sKeyAscii    '// return 0 value to KeyAscii
270               Exit Function

280           Else    '// If string found
290               intPos = .SelStart    '// save cursor position
300               intLength = Len(.List(lngFind)) - Len(.Text)    '// save remaining highlighted text length
310               .SelText = .SelText & Right(.List(lngFind), intLength)    '// change new text in string
                  '.Text = .List(lngFind)'// Use this instead of the above .Seltext line to set the text typed to the exact case of the item selected in the combo box.
320               .SelStart = intPos    '// set cursor position
330               .SelLength = intLength    '// set selected length
340           End If
350       End With


360       Exit Function

AutoComplete_Error:

          Dim strES As String
          Dim intEL As Integer

370       intEL = Erl
380       strES = Err.Description
390       LogError "Shared", "AutoComplete", intEL, strES


End Function

Public Function QueryCombo(cmbCombo As ComboBox) As String

          Dim strFoundString As String
          Dim lngFind As Long

10        On Error GoTo QueryCombo_Error

20        With cmbCombo
30            lngFind = SendMessage(.hWnd, CB_FINDSTRINGEXACT, 0, ByVal .Text)
40            If lngFind = -1 Then
50                QueryCombo = ""
60            Else
70                QueryCombo = .List(lngFind)

80            End If
90        End With


100       Exit Function

QueryCombo_Error:

          Dim strES As String
          Dim intEL As Integer

110       intEL = Erl
120       strES = Err.Description
130       LogError "Shared", "QueryCombo", intEL, strES

End Function

Public Sub RemoveReport(ConIndex As String, SampleID As String, Department As String, Optional InterimORFinal As Integer = 0)    'Masood 19_Feb_2013

          Dim sql As String
10        On Error GoTo RemoveReport_Error

20        If InterimORFinal = 1 Then
30            sql = "Delete from Reports " & _
                    "where SampleID = '" & Val(SampleID) & "' AND Dept = '" & Department & "'"
40            Cnxn(Val(ConIndex)).Execute sql
50        ElseIf InterimORFinal = 0 Then
60            sql = "Delete from UnauthorisedReports " & _
                    "where SampleID = '" & Val(SampleID) & "' AND Dept = '" & Department & "'"
70            Cnxn(Val(ConIndex)).Execute sql
80        End If
90        Exit Sub


RemoveReport_Error:

          Dim strES As String
          Dim intEL As Integer

100       intEL = Erl
110       strES = Err.Description
120       LogError "basShared", "RemoveReport", intEL, strES, sql
End Sub

Public Sub MarkFlexGridRow(grd As MSFlexGrid, StartRow As Integer, StartCol As Integer, EndCol As Integer, RowColor As Long)

          Dim n As Integer

10        On Error GoTo MarkFlexGridRow_Error

20        With grd
30            .Row = StartRow
40            For n = StartCol To EndCol
50                .Col = n
60                .CellBackColor = RowColor
70            Next n
80        End With

90        Exit Sub

MarkFlexGridRow_Error:

          Dim strES As String
          Dim intEL As Integer

100       intEL = Erl
110       strES = Err.Description
120       LogError "basShared", "MarkFlexGridRow", intEL, strES

End Sub

'---------------------------------------------------------------------------------------
' Procedure : SetFormStyle
' Author    : Babar Shahzad
' Date      : 24/09/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub SetFormStyle(Frm As Form)

On Error GoTo SetFormStyle_Error

With Frm
    .Icon = frmMain.ImageList1.ListImages(5).Picture
End With

Exit Sub

SetFormStyle_Error:

 Dim strES As String
 Dim intEL As Integer

 intEL = Erl
 strES = Err.Description
 LogError "basShared", "SetFormStyle", intEL, strES
    
End Sub
Public Function ISITEMINLIST(Item As String, ListType As String) As Boolean
          Dim tb As Recordset
          Dim sql As String
          Dim RetVal As Boolean

10        On Error GoTo IsItemInList_Error

20        sql = "SELECT Code, ListType, Text From Lists " & _
                "WHERE InUse = 1 AND ListType = '" & ListType & " '  AND Text = '" & Item & "'"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        If (tb.EOF And tb.BOF) Then
60            RetVal = False
70        ElseIf (Trim$(tb!Text & "") = "") Then
80            RetVal = False
90        Else
100           RetVal = True
110       End If

120       ISITEMINLIST = RetVal

130       Exit Function

IsItemInList_Error:

          Dim strES As String
          Dim intEL As Integer

140       intEL = Erl
150       strES = Err.Description
160       LogError "BasShared", "IsItemInList", intEL, strES, sql

End Function

Public Function GetHospitalName(Code As String, Discipline As String) As String

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo GetHospitalName_Error

20    sql = "SELECT Hospital FROM " & Discipline & "TestDefinitions WHERE Code = '" & Code & "'"
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql
50    If tb.EOF Then
60        GetHospitalName = ""
70    Else
80        GetHospitalName = tb!Hospital & ""
90    End If

100   Exit Function

GetHospitalName_Error:

      Dim strES As String
      Dim intEL As Integer

110   intEL = Erl
120   strES = Err.Description
130   LogError "basShared", "GetHospitalName", intEL, strES, sql

End Function

Public Function GetSplitValue(Value As String, ItemNo As Integer) As String

      Dim Res() As String

10    On Error GoTo GetSplitValue_Error

20    Res = Split(Value, "|")
30    If UBound(Res) > -1 Then
40        GetSplitValue = Res(ItemNo)
50    Else
60        GetSplitValue = Value
70    End If

80    Exit Function

GetSplitValue_Error:

       Dim strES As String
       Dim intEL As Integer

90     intEL = Erl
100    strES = Err.Description
110    LogError "basShared", "GetSplitValue", intEL, strES
End Function

Public Sub UpdateConsultantList(ByVal SampleIDWithOffset As String, ByVal Dept As String, ByVal Status As ConsultantListStatus, ByVal Ack As Integer, ByVal ConAck As Integer)

      Dim sql As String

10    On Error GoTo UpdateConsultantList_Error

20    sql = "If Exists(Select 1 From ConsultantList " & _
                    "Where SampleID = " & SampleIDWithOffset & " ) " & _
                    "Begin " & _
                    "Update ConsultantList Set " & _
                    "SampleID = " & SampleIDWithOffset & ", " & _
                    "Department = '" & Dept & "', " & _
                    "Status  = '" & Status & "', " & _
                    "Username = '" & UserName & "', " & _
                    "Ack = " & Ack & ", ConAck = " & ConAck & " " & _
                    "Where SampleID = " & SampleIDWithOffset & "  " & _
                    "End  " & _
                    "Else " & _
                    "Begin  " & _
                    "Insert Into ConsultantList (SampleID,Department,Status,Username, Ack, ConAck) Values (" & SampleIDWithOffset & ",'Micro','" & Status & "','" & UserName & "', " & Ack & ", " & ConAck & ") " & _
                    "End"

30    Cnxn(0).Execute sql


40    Exit Sub

UpdateConsultantList_Error:

       Dim strES As String
       Dim intEL As Integer

50     intEL = Erl
60     strES = Err.Description
70     LogError "basShared", "UpdateConsultantList", intEL, strES, sql
          
End Sub

Public Function GetLabLinkMapping(ByVal MappingType As String, ByVal TargetHospital As String, ByVal SourceValue As String) As String

      Dim sql  As String
      Dim tb As Recordset

10    On Error GoTo GetLabLinkMapping_Error


20    sql = "SELECT * FROM LabLinkMapping WHERE UPPER(MappingType) = '" & UCase(MappingType) & "' AND UPPER(SourceValue) = '" & UCase(SourceValue) & "' " & _
      "AND UPPER(TargetHospital) = '" & UCase(TargetHospital) & "'"

30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql
50    If Not tb.EOF Then
60        GetLabLinkMapping = tb!TargetValue & ""
70    Else
80        GetLabLinkMapping = ""
90    End If


100   Exit Function

GetLabLinkMapping_Error:
      Dim strES As String
      Dim intEL As Integer

110   intEL = Erl
120   strES = Err.Description
130   LogError "basShared", "GetLabLinkMapping", intEL, strES

                      
End Function
Public Function IsResultAmended(Department As String, SampleID As Long, Code As String, Value As String) As Boolean

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo IsResultAmended_Error
20    If GetOptionSetting("EnableAmendedTest", "0") = "1" Then
30        sql = "SELECT Count(*) AS CNT  " & _
                "FROM " & Department & "ResultsAudit  WHERE SampleID = '" & SampleID & "' " & _
                "AND Code = '" & Code & "' AND ltrim(rtrim(Result)) <> (Select ltrim(rtrim(Result)) FROM " & Department & "Results  WHERE SampleID = '" & SampleID & "' AND Code = '" & Code & "' )"

40        Set tb = New Recordset
50        RecOpenClient 0, tb, sql

60        IsResultAmended = tb!Cnt > 0
70    Else
80        IsResultAmended = False
90    End If

100   Exit Function

IsResultAmended_Error:
      Dim strES As String
      Dim intEL As Integer

110   intEL = Erl
120   strES = Err.Description
130   LogError "basShared", "IsResultAmended", intEL, strES, sql

End Function
Public Function IsHaemResultAmended(SampleID As Long, TestName As String) As Boolean

      Dim tb As Recordset
      Dim sql As String


10    On Error GoTo IsHaemResultAmended_Error

20    If GetOptionSetting("EnableAmendedTest", "0") = "1" Then
30        sql = "SELECT Count(*) AS CNT  " & _
                "FROM ArcHaemResults  WHERE SampleID = '" & SampleID & "' " & _
                "AND ltrim(rtrim(" & TestName & ")) <> (Select ltrim(rtrim(" & TestName & ")) FROM  HaemResults  WHERE SampleID = '" & SampleID & "'  )"

40        Set tb = New Recordset
50        RecOpenClient 0, tb, sql

60        IsHaemResultAmended = tb!Cnt > 0
70    Else
80        IsHaemResultAmended = False
90    End If



100   Exit Function

IsHaemResultAmended_Error:
      Dim strES As String
      Dim intEL As Integer

110   intEL = Erl
120   strES = Err.Description
130   LogError "basShared", "IsHaemResultAmended", intEL, strES, sql

End Function

Public Sub ReleaseMicro(ByVal SampleID As String, ByVal Release As Integer)

      Dim sql        As String

10    On Error GoTo ReleaseMicro_Error

20    sql = "UPDATE Demographics " & _
            "SET ForMicro = '" & Release & "', " & _
            "MicroHealthLinkReleaseTime = getdate() " & _
            "WHERE SampleID = '" & SampleID & "'"
      '
      'If Release Then
      '  sql = "UPDATE Demographics " & _
         '        "SET ForMicro = '1', " & _
         '        "MicroHealthLinkReleaseTime = getdate() " & _
         '        "WHERE SampleID = '" & SampleID & "'"
      'Else
      '  sql = "UPDATE Demographics " & _
         '        "SET ForMicro = '0', " & _
         '        "MicroHealthLinkReleaseTime = NULL " & _
         '        "WHERE SampleID = '" & SampleID & "'"
      'End If

30    Cnxn(0).Execute sql

40    Exit Sub

ReleaseMicro_Error:

      Dim strES      As String
      Dim intEL      As Integer

50    intEL = Erl
60    strES = Err.Description
70    LogError "basShared", "ReleaseMicro", intEL, strES, sql

End Sub

Public Function IsMicroReleased(ByVal SampleID As String) As Boolean

      Dim sql As String
      Dim tb As Recordset

10    On Error GoTo IsMicroReleased_Error

20    IsMicroReleased = False

30    sql = "SELECT COUNT(*) Tot FROM Demographics " & _
            "WHERE COALESCE(ForMicro, 0) <> 0 " & _
            "AND SampleID = '" & SampleID & "'"
40    Set tb = Cnxn(0).Execute(sql)
50    If tb!Tot > 0 Then
60      IsMicroReleased = True
70    End If

80    Exit Function

IsMicroReleased_Error:

      Dim strES As String
      Dim intEL As Integer

90    intEL = Erl
100   strES = Err.Description
110   LogError "basShared", "IsMicroReleased", intEL, strES, sql

End Function

Public Function GetLatestRunDateTime(ByVal Disp As String, ByVal SampleID As String, ByVal RunDateTime As String) As String

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo GetLatestRunDateTime_Error

20    If GetOptionSetting("GetLatestRunDate", "0") = 0 Then
30        GetLatestRunDateTime = RunDateTime
40        Exit Function
50    End If

60    sql = "SELECT TOP 1 RunTime FROM " & Disp & "Results WHERE sampleid= '" & SampleID & "' ORDER BY Runtime"


70    Set tb = New Recordset
80    RecOpenClient 0, tb, sql
90    If tb.EOF Then
100       GetLatestRunDateTime = ""
110   Else
120       GetLatestRunDateTime = tb!RunTime
130   End If

140   Exit Function

GetLatestRunDateTime_Error:

      Dim strES As String
      Dim intEL As Integer

150   intEL = Erl
160   strES = Err.Description
170   LogError "basShared", "GetLatestRunDateTime", intEL, strES, sql

End Function
