Attribute VB_Name = "basMicro"
Option Explicit

Public Type FaecalOrder

    SampleID As Double
    cS As Boolean
    ssScreen As Boolean
    Campylobacter As Boolean
    Coli0157 As Boolean
    Cryptosporidium As Boolean
    Rota As Boolean
    Adeno As Boolean
    OB0 As Boolean
    OB1 As Boolean
    OB2 As Boolean
    OP As Boolean
    ToxinAB As Boolean
    HPylori As Boolean
    RedSub As Boolean
    CDiffCulture As Boolean
    GDH As Boolean
    PCR As Boolean
    GL As Boolean
End Type

Public Enum MicroSections
    msUrine = 1
    msIDENTIFICATION = 2
    msFAECES = 3
    msCANDS = 4
    msFOB = 5
    msROTAADENO = 6
    msRedSub = 7
    msRSV = 8
    msCSF = 9
    msCDIFF = 10
    msOP = 11
    msBLOODCULTURE = 12
    msHPylori = 13
End Enum

Public Function LeftOfBar(ByVal s As String) As String

      's="Value|ForeColour|BackColour"

          Dim v() As String
          Dim RetVal As String

10        On Error GoTo LeftOfBar_Error

20        v = Split(s, "|")

30        If UBound(v) = -1 Then
40            RetVal = Trim$(s)
50        Else
60            RetVal = Trim$(v(0))
70        End If

80        LeftOfBar = RetVal

90        Exit Function

LeftOfBar_Error:

          Dim strES As String
          Dim intEL As Integer

100       intEL = Erl
110       strES = Err.Description
120       LogError "basMicro", "LeftOfBar", intEL, strES

End Function


Public Sub SaveInitialMicroSiteDetails(ByVal Site As String, _
                                       ByVal SampleIDWithOffset As Double, _
                                       ByVal SiteDetails As String)

          Dim sql As String

10        On Error GoTo SaveInitialMicroSiteDetails_Error

          'Created on 18/02/2011 11:53:12
          'Autogenerated by SQL Scripting

20        sql = "If Exists(Select 1 From MicroSiteDetails " & _
                "Where SampleID = @SampleID0 ) " & _
                "Begin " & _
                "Update MicroSiteDetails Set " & _
                "SampleID = @SampleID0, " & _
                "Site = '@Site1', " & _
                "SiteDetails = '@SiteDetails2', " & _
                "UserName = '@UserName7' " & _
                "Where SampleID = @SampleID0  " & _
                "End  " & _
                "Else " & _
                "Begin  " & _
                "Insert Into MicroSiteDetails (SampleID, Site, SiteDetails, UserName) Values " & _
                "(@SampleID0, '@Site1', '@SiteDetails2', '@UserName7') " & _
                "End"

30        sql = Replace(sql, "@SampleID0", SampleIDWithOffset)
40        sql = Replace(sql, "@Site1", Site)
50        sql = Replace(sql, "@SiteDetails2", SiteDetails)
60        sql = Replace(sql, "@UserName7", UserName)

70        Cnxn(0).Execute sql

          '
          'sql = "Select * from MicroSiteDetails where " & _
           '      "SampleID = '" & SampleIDWithOffset & "'"
          'Set tb = New Recordset
          'RecOpenClient 0, tb, sql
          'If tb.EOF Then
          '    tb.AddNew
          'End If
          'tb!SampleID = SampleIDWithOffset
          'tb!Site = Site
          'tb!SiteDetails = SiteDetails
          'tb!UserName = UserName
          'tb.Update

80        Exit Sub

SaveInitialMicroSiteDetails_Error:

          Dim strES As String
          Dim intEL As Integer

90        intEL = Erl
100       strES = Err.Description
110       LogError "basMicro", "SaveInitialMicroSiteDetails", intEL, strES, sql

End Sub

Public Function IsAnyRecordPresent(ByVal TableName As String, _
                                   ByVal SampleID As Double) _
                                   As Boolean

          Dim sql As String
          Dim tb As Recordset

10        On Error GoTo IsAnyRecordPresent_Error

20        sql = "SELECT COUNT(*) Tot FROM " & TableName & " WHERE " & _
                "SampleID = '" & SampleID & "'"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql

50        IsAnyRecordPresent = tb!Tot > 0

60        Exit Function

IsAnyRecordPresent_Error:

          Dim strES As String
          Dim intEL As Integer

70        intEL = Erl
80        strES = Err.Description
90        LogError "basMicro", "IsAnyRecordPresent", intEL, strES, sql

End Function

Public Sub GetFaecalOrder(ByVal NoOffsetSampleID As Long, ByRef fOrder As FaecalOrder)

          Dim tb As Recordset
          Dim sql As String
          Dim SampleIDWithOffset As Double

10        On Error GoTo GetFaecalOrder_Error

20        SampleIDWithOffset = Val(NoOffsetSampleID) + SysOptMicroOffset(0)

30        sql = "SELECT * FROM FaecalRequests WHERE " & _
                "SampleID = " & SampleIDWithOffset
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        If Not tb.EOF Then
70            With fOrder
80                .cS = tb!cS
90                .ssScreen = tb!ssScreen
100               .Campylobacter = tb!Campylobacter
110               .Coli0157 = tb!Coli0157
120               .Cryptosporidium = tb!Cryptosporidium
130               .Rota = tb!Rota
140               .Adeno = tb!Adeno
150               .OB0 = tb!OB0
160               .OB1 = tb!OB1
170               .OB2 = tb!OB2
180               .OP = tb!OP
190               .ToxinAB = tb!ToxinAB
200               .HPylori = tb!HPylori
210               .RedSub = IIf(IsNull(tb!RedSub), 0, tb!RedSub)
220               .CDiffCulture = IIf(IsNull(tb!CDiff), 0, tb!CDiff)
230               .GDH = IIf(IsNull(tb!GDH), 0, tb!GDH)
240               .PCR = IIf(IsNull(tb!PCR), 0, tb!PCR)
250               .GL = IIf(IsNull(tb!GL), 0, tb!GL)
260           End With
270       Else
280           With fOrder
290               .cS = False
300               .ssScreen = False
310               .Campylobacter = False
320               .Coli0157 = False
330               .Cryptosporidium = False
340               .Rota = False
350               .Adeno = False
360               .OB0 = False
370               .OB1 = False
380               .OB2 = False
390               .OP = False
400               .ToxinAB = False
410               .HPylori = False
420               .RedSub = False
430               .CDiffCulture = False
440               .GDH = False
450               .PCR = False
460               .GL = False
470           End With
480       End If

490       Exit Sub

GetFaecalOrder_Error:

          Dim strES As String
          Dim intEL As Integer

500       intEL = Erl
510       strES = Err.Description
520       LogError "basMicro", "GetFaecalOrder", intEL, strES, sql

End Sub
Public Function IsFluid(ByVal Site As String) As Boolean

          Dim tb As Recordset
          Dim sql As String
          Dim RetVal As Boolean

10        On Error GoTo IsFluid_Error

20        sql = "SELECT COUNT(*) AS Tot FROM Lists WHERE " & _
                "ListType = 'FF' " & _
                "AND Text = '" & Site & "'"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        RetVal = tb!Tot > 0

60        IsFluid = RetVal

70        Exit Function

IsFluid_Error:

          Dim strES As String
          Dim intEL As Integer

80        intEL = Erl
90        strES = Err.Description
100       LogError "basMicro", "IsFluid", intEL, strES, sql

End Function

Public Sub SaveFaecalOrder(ByVal NoOffsetSampleID As Long, ByRef fOrder As FaecalOrder)

          Dim sql As String
          Dim SampleIDWithOffset As Double

10        On Error GoTo SaveFaecalOrder_Error

20        SampleIDWithOffset = Val(NoOffsetSampleID) + SysOptMicroOffset(0)


          'Created on 01/02/2011 16:29:02
          'Autogenerated by SQL Scripting

30        sql = "If Exists(Select 1 From FaecalRequests " & _
                "Where SampleID = @SampleID0 ) " & _
                "Begin " & _
                "Update FaecalRequests Set " & _
                "OP = @OP1, " & _
                "Rota = @Rota2, " & _
                "Adeno = @Adeno3, " & _
                "Coli0157 = @Coli01578, " & _
                "OB0 = @OB09, " & _
                "OB1 = @OB110, " & _
                "OB2 = @OB211, " & _
                "ssScreen = '@ssScreen12', " & _
                "cS = '@cS14', " & _
                "Campylobacter = '@Campylobacter15', " & _
                "Cryptosporidium = '@Cryptosporidium16', " & _
                "ToxinAB = '@ToxinAB17', " & _
      "HPylori = '@HPylori18', "
40        sql = sql & "RedSub = @RedSub19, " & _
                "CDiff = @CDiff22, " & _
                "GDH = @GDH23, " & _
                "PCR = @PCR24, " & _
                "GL=@GL," & _
                "UserName = '@UserName20', " & _
                "DateTimeOfRecord = '@DateTimeOfRecord21' " & _
                "Where SampleID = @SampleID0  " & _
                "End  " & _
                "Else " & _
                "Begin  " & _
                "Insert Into FaecalRequests (SampleID, OP, Rota, Adeno, Coli0157, OB0, OB1, OB2, ssScreen, cS, " & _
                "Campylobacter, Cryptosporidium, ToxinAB, HPylori, RedSub, CDiff, GDH, PCR, UserName, DateTimeOfRecord,GL) Values " & _
                "(@SampleID0, @OP1, @Rota2, @Adeno3, @Coli01578, @OB09, @OB110, @OB211, '@ssScreen12', " & _
                "'@cS14', '@Campylobacter15', '@Cryptosporidium16', '@ToxinAB17', '@HPylori18', @RedSub19, @CDiff22, " & _
                "@GDH23, @PCR24 ,'@UserName20', '@DateTimeOfRecord21',@GL) " & _
                "End"
50        With fOrder
60            sql = Replace(sql, "@SampleID0", SampleIDWithOffset)
70            sql = Replace(sql, "@OP1", IIf(.OP, 1, 0))
80            sql = Replace(sql, "@Rota2", IIf(.Rota, 1, 0))
90            sql = Replace(sql, "@Adeno3", IIf(.Adeno, 1, 0))
100           sql = Replace(sql, "@Coli01578", IIf(.Coli0157, 1, 0))
110           sql = Replace(sql, "@OB09", IIf(.OB0, 1, 0))
120           sql = Replace(sql, "@OB110", IIf(.OB1, 1, 0))
130           sql = Replace(sql, "@OB211", IIf(.OB2, 1, 0))
140           sql = Replace(sql, "@ssScreen12", IIf(.ssScreen, 1, 0))
150           sql = Replace(sql, "@cS14", IIf(.cS, 1, 0))
160           sql = Replace(sql, "@Campylobacter15", IIf(.Campylobacter, 1, 0))
170           sql = Replace(sql, "@Cryptosporidium16", IIf(.Cryptosporidium, 1, 0))
180           sql = Replace(sql, "@ToxinAB17", IIf(.ToxinAB, 1, 0))
190           sql = Replace(sql, "@HPylori18", IIf(.HPylori, 1, 0))
200           sql = Replace(sql, "@RedSub19", IIf(.RedSub, 1, 0))
210           sql = Replace(sql, "@CDiff22", IIf(.CDiffCulture, 1, 0))
220           sql = Replace(sql, "@GDH23", IIf(.GDH, 1, 0))
230           sql = Replace(sql, "@PCR24", IIf(.PCR, 1, 0))
240           sql = Replace(sql, "@UserName20", UserName)
250           sql = Replace(sql, "@DateTimeOfRecord21", Format(Now, "dd/MMM/yyyy hh:mm:ss"))
260           sql = Replace(sql, "@GL", IIf(.GL, 1, 0))
270       End With
280       Cnxn(0).Execute sql


          'sql = "SELECT * FROM FaecalRequests WHERE " & _
           '      "SampleID = '" & SampleIDWithOffset & "' "
          'Set tb = New Recordset
          'RecOpenServer 0, tb, sql
          'If tb.EOF Then
          '    tb.AddNew
          '    tb!SampleID = SampleIDWithOffset
          'End If
          'With fOrder
          '    tb!cS = .cS
          '    tb!ssScreen = .ssScreen
          '    tb!Campylobacter = .Campylobacter
          '    tb!Coli0157 = .Coli0157
          '    tb!Cryptosporidium = IIf(.Cryptosporidium, 1, 0)
          '    tb!Rota = .Rota
          '    tb!Adeno = .Adeno
          '    tb!OB0 = .OB0
          '    tb!OB1 = .OB1
          '    tb!OB2 = .OB2
          '    tb!OP = .OP
          '    tb!ToxinAB = IIf(.ToxinAB, 1, 0)
          '    tb!HPylori = IIf(.HPylori, 1, 0)
          '    tb!RedSub = IIf(.RedSub, 1, 0)
          '    tb!UserName = UserName
          '    tb!DateTimeOfRecord = Now
          'End With
          'tb.Update

290       Exit Sub

SaveFaecalOrder_Error:

          Dim strES As String
          Dim intEL As Integer

300       intEL = Erl
310       strES = Err.Description
320       LogError "basMicro", "SaveFaecalOrder", intEL, strES, sql

End Sub

Public Function AntibioticCodeFor(ByVal inABName As String) As String

          Dim tb As New Recordset
          Dim sql As String

10        On Error GoTo AntibioticCodeFor_Error

20        sql = "SELECT Code from Antibiotics WHERE " & _
                "AntibioticName = '" & inABName & "'"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        If tb.EOF Then
60            AntibioticCodeFor = "???"
70        Else
80            If Trim$(tb!Code & "") <> "" Then
90                AntibioticCodeFor = Trim$(tb!Code)
100           Else
110               AntibioticCodeFor = "???"
120           End If
130       End If

140       Exit Function

AntibioticCodeFor_Error:

          Dim strES As String
          Dim intEL As Integer

150       intEL = Erl
160       strES = Err.Description
170       LogError "basMicro", "AntibioticCodeFor", intEL, strES, sql

End Function

Public Function AntibioticNameFor(ByVal inABCode As String) As String

          Dim tb As New Recordset
          Dim sql As String

10        On Error GoTo AntibioticNameFor_Error

20        sql = "SELECT AntibioticName from Antibiotics WHERE " & _
                "Code = '" & inABCode & "'"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        If tb.EOF Then
60            AntibioticNameFor = inABCode
70        Else
80            AntibioticNameFor = Trim$(tb!AntibioticName & "")
90        End If

100       Exit Function

AntibioticNameFor_Error:

          Dim strES As String
          Dim intEL As Integer

110       intEL = Erl
120       strES = Err.Description
130       LogError "basMicro", "AntibioticNameFor", intEL, strES, sql

End Function

Public Function AntibioticRerportNameFor(ByVal inABCode As String) As String

          Dim tb As New Recordset
          Dim sql As String

10        On Error GoTo AntibioticRerportNameFor_Error

20        sql = "SELECT ReportName from Antibiotics WHERE " & _
                "Code = '" & inABCode & "'"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        If tb.EOF Then
60            AntibioticRerportNameFor = inABCode
70        Else
80            AntibioticRerportNameFor = Trim$(tb!ReportName & "")
90        End If

100       Exit Function

AntibioticRerportNameFor_Error:

          Dim strES As String
          Dim intEL As Integer

110       intEL = Erl
120       strES = Err.Description
130       LogError "basMicro", "AntibioticRerportNameFor", intEL, strES, sql

End Function


Public Function IdentIsSaveable(ByVal Index As Integer) As Boolean

10        On Error GoTo IdentIsSaveable_Error

20        IdentIsSaveable = False

30        With frmEditMicrobiologyNew

40            If Trim$(.cmbGram(Index).Text & _
                       .txtZN(Index).Text & _
                       .cmbWetPrep(Index).Text & _
                       .txtIndole(Index).Text & _
                       .txtCoagulase(Index).Text & _
                       .txtCatalase(Index).Text & _
                       .txtOxidase(Index).Text & _
                       .txtReincubation(Index) & _
                       .txtNotes(Index).Text) <> "" Then

50                IdentIsSaveable = True

60            End If

70        End With

80        Exit Function

IdentIsSaveable_Error:

          Dim strES As String
          Dim intEL As Integer

90        intEL = Erl
100       strES = Err.Description
110       LogError "basMicro", "IdentIsSaveable", intEL, strES

End Function

Public Sub PrintResultUrnWin(ByVal SampleID As String)

          Dim tb As New Recordset
          Dim sql As String
          Dim GP As String
          Dim Clin As String
          Dim Ward As String

10        On Error GoTo PrintResultUrnWin_Error

20        Ward = ""
30        GP = ""
40        Clin = ""

50        sql = "SELECT * FROM Demographics WHERE " & _
                "SampleID = '" & Val(SampleID) + SysOptMicroOffset(0) & "'"

60        Set tb = New Recordset
70        RecOpenServer 0, tb, sql
80        If Not tb.EOF Then
90            Ward = tb!Ward & ""
100           GP = tb!GP & ""
110           Clin = tb!Clinician & ""
120       End If

130       sql = "SELECT * FROM PrintPending WHERE " & _
                "Department = 'N' " & _
                "AND SampleID = '" & SampleID & "'"
140       Set tb = New Recordset
150       RecOpenServer 0, tb, sql
160       If tb.EOF Then
170           tb.AddNew
180       End If
190       tb!SampleID = SampleID
200       tb!Department = "N"
210       tb!Initiator = UserName
220       tb!pTime = Now
230       tb!Ward = Ward & ""
240       tb!GP = GP & ""
250       tb!Clinician = Clin & ""
260       tb.Update

270       Exit Sub

PrintResultUrnWin_Error:

          Dim strES As String
          Dim intEL As Integer



280       intEL = Erl
290       strES = Err.Description
300       LogError "basMicro", "PrintResultUrnWin", intEL, strES, sql

End Sub

Public Sub UpdatePrintValidLog(ByVal SampleID As Double, _
                               ByVal Section As String, _
                               ByVal LogAsValid As Integer, _
                               ByVal LogAsPrinted As Integer)

      '0 - not valid
      '1 - valid
      '2 - no change

          Dim tb As Recordset
          Dim sql As String
          Dim LogDept As String

10        On Error GoTo UpdatePrintValidLog_Error

20        Select Case UCase$(Section)
          Case "DEMOGRAPHICS": LogDept = ""
30        Case "URINE": LogDept = "U"
40        Case "IDENTIFICATION": LogDept = ""
50        Case "FAECES": LogDept = ""
60        Case "CANDS": LogDept = "D"
70        Case "FOB": LogDept = "F"
80        Case "ROTAADENO": LogDept = "A"
90        Case "REDSUB": LogDept = "R"
100       Case "RSV": LogDept = "V"
110       Case "CSF": LogDept = "C"
120       Case "CDIFF": LogDept = "G"
130       Case "HPYLORI": LogDept = "Y"
140       Case "OP": LogDept = "O"
150       Case "BLOODCULTURE": LogDept = "B"
160       Case "SEMEN": LogDept = "Z"
170       Case Else: LogDept = ""
180       End Select

190       If LogDept = "" Then
200           Exit Sub
210       End If
          '320   If LogAsPrinted = 0 Then
          '330     tb!Printed = 0
          '340     tb!PrintedBy = ""
          '350     tb!PrintedDateTime = Null
          '360   ElseIf LogAsPrinted = 1 Then
          '370     tb!Printed = 1
          '380     tb!PrintedBy = IIf(LogAsPrinted, AddTicks(Username), "")
          '390     tb!PrintedDateTime = IIf(LogAsPrinted, Format$(Now, "dd/MMM/yyyy hh:mm:ss"), Null)
          '400   End If
          '
          '410   If LogAsValid = 0 Then
          '420     tb!Valid = 0
          '430     tb!ValidatedBy = ""
          '440     tb!ValidatedDateTime = Null
          '450   ElseIf LogAsValid = 1 Then
          '460     tb!Valid = 1
          '470     tb!ValidatedBy = AddTicks(Username)
          '480     tb!ValidatedDateTime = Format$(Now, "dd/MMM/yyyy hh:mm:ss")
          '490   End If
          '210   If LogAsPrinted = 0 Then
          '220     P = 0
          '230     PBy = ""
          '240     PDate = "NULL"
          '250   ElseIf LogAsPrinted = 1 Then
          '260     P = 1
          '270     PBy = IIf(LogAsPrinted, AddTicks(Username), "")
          '280     PDate = IIf(LogAsPrinted, Format$(Now, "'dd/MMM/yyyy hh:mm:ss'"), "NULL")
          '290   End If
          '300   If LogAsValid = 0 Then
          '310     V = 0
          '320     VBy = ""
          '330     VDate = "NULL"
          '340   ElseIf LogAsValid = 1 Then
          '350     V = 1
          '360     VBy = AddTicks(Username)
          '370     VDate = Format$(Now, "'dd/MMM/yyyy hh:mm:ss'")
          '380   End If
          '
          '390   sql = "SET NOCOUNT ON " & _
           '      "IF EXISTS (SELECT * FROM PrintValidLog WHERE " & _
           '      "           SampleID = '" & SampleID & "' " & _
           '      "           AND Department = '" & LogDept & "') " & _
           '      "  BEGIN " & _
           '      "    INSERT INTO PrintValidLogArc " & _
           '      "    SELECT PrintValidLog.*, '" & AddTicks(Username) & "', getdate() " & _
           '      "    FROM PrintValidLog WHERE " & _
           '      "    SampleID = '" & SampleID & "' " & _
           '      "    AND Department = 'D' " & _
           '      "    UPDATE PrintValidLog " & _
           '      "    SET Printed = '" & P & "', PrintedBy = '" & PBy & "', PrintedDateTime = " & PDate & ", " & _
           '      "    Valid = '" & V & "', ValidatedBy = '" & VBy & "', ValidatedDateTime = " & VDate & " " & _
           '      "    WHERE SampleID = '" & SampleID & "' " & _
           '      "    AND Department = '" & LogDept & "' " & _
'      "  END "
          '400   sql = sql & "ELSE " & _
           '      "  INSERT INTO PrintValidLog " & _
           '      "  (SampleID, Department, Printed, Valid, PrintedBy, PrintedDateTime, ValidatedBy, ValidatedDateTime) VALUES " & _
           '      "  ('" & SampleID & "', " & _
           '      "   '" & LogDept & "', " & _
           '      "   '" & P & "', " & _
           '      "   '" & V & "', " & _
           '      "   '" & PBy & "', " & _
           '      "    " & PDate & ", " & _
           '      "   '" & VBy & "', " & _
           '      "   " & VDate & " )"
          '410   Cnxn(0).Execute sql

220       sql = "SELECT * FROM PrintValidLog WHERE " & _
                "SampleID = '" & SampleID & "' " & _
                "AND Department = '" & LogDept & "'"
230       Set tb = New Recordset
240       RecOpenClient 0, tb, sql
250       If tb.EOF Then
260           tb.AddNew
270       Else
280           sql = "INSERT INTO PrintValidLogArc " & _
                    "  SELECT PrintValidLog.*, " & _
                    "  '" & AddTicks(UserName) & "', " & _
                    "  '" & Format$(Now, "dd/MMM/yyyy hh:mm:ss") & "' " & _
                    "  FROM PrintValidLog WHERE " & _
                    "  SampleID = '" & SampleID & "' " & _
                    "  AND Department = '" & LogDept & "' "
290           Cnxn(0).Execute sql
300       End If
310       tb!SampleID = SampleID
320       tb!Department = LogDept

330       If LogAsPrinted = 0 Then
340           tb!Printed = 0
350           tb!PrintedBy = ""
360           tb!PrintedDateTime = Null
370       ElseIf LogAsPrinted = 1 Then
380           tb!Printed = 1
390           tb!PrintedBy = IIf(LogAsPrinted, AddTicks(UserName), "")
400           tb!PrintedDateTime = IIf(LogAsPrinted, Format$(Now, "dd/MMM/yyyy hh:mm:ss"), Null)
410       End If

420       If LogAsValid = 0 Then
430           tb!Valid = 0
440           tb!ValidatedBy = ""
450           tb!ValidatedDateTime = Null
460       ElseIf LogAsValid = 1 Then
470           tb!Valid = 1
480           tb!ValidatedBy = AddTicks(UserName)
490           tb!ValidatedDateTime = Format$(Now, "dd/MMM/yyyy hh:mm:ss")
500       End If

510       tb.Update

520       Exit Sub

UpdatePrintValidLog_Error:

          Dim strES As String
          Dim intEL As Integer

530       intEL = Erl
540       strES = Err.Description
550       LogError "basMicro", "UpdatePrintValidLog", intEL, strES

End Sub

Public Function ValidStatus4MicroDept(ByVal SampleIDWithOffset As Double, ByVal strDept As String) As Boolean

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo ValidStatus4MicroDept_Error

20        sql = "SELECT COALESCE(Valid, 0) AS Valid FROM PrintValidLog WHERE " & _
                "SampleID = '" & SampleIDWithOffset & "' " & _
                "AND Department = '" & strDept & "'"

30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        If tb.EOF Then
60            ValidStatus4MicroDept = False
70        Else
80            ValidStatus4MicroDept = tb!Valid
90        End If

100       Exit Function

ValidStatus4MicroDept_Error:

          Dim strES As String
          Dim intEL As Integer

110       intEL = Erl
120       strES = Err.Description
130       LogError "basMicro", "ValidStatus4MicroDept", intEL, strES, sql

End Function

Public Function GetMicroSection(MicroSection As MicroSections) As String

          Dim Section As String

10        On Error GoTo GetMicroSection_Error

20        Select Case MicroSection
          Case msUrine:
30            Section = "U"
40        Case msIDENTIFICATION:
50            Section = ""
60        Case msFAECES:
70            Section = ""
80        Case msCANDS:
90            Section = "D"
100       Case msFOB:
110           Section = "F"
120       Case msROTAADENO:
130           Section = "A"
140       Case msRedSub:
150           Section = "R"
160       Case msRSV:
170           Section = "V"
180       Case msCSF:
190           Section = "C"
200       Case msCDIFF:
210           Section = "G"
220       Case msOP:
230           Section = "O"
240       Case msBLOODCULTURE:
250           Section = "B"
260       Case msHPylori:
270           Section = "Y"
280       Case Else
290           Section = ""
300       End Select

310       GetMicroSection = Section

320       Exit Function

GetMicroSection_Error:

          Dim strES As String
          Dim intEL As Integer

330       intEL = Erl
340       strES = Err.Description
350       LogError "basMicro", "GetMicroSection", intEL, strES

End Function

Public Function IsMicroValid(SampleIDWithOffset As Double, MicroSection As MicroSections) As Boolean

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo IsMicroValid_Error

20        IsMicroValid = False

30        sql = "SELECT * FROM PrintValidLog WHERE " & _
                "SampleID = '" & SampleIDWithOffset & "' " & _
                "AND Department = '%department'"

40        sql = Replace(sql, "%department", GetMicroSection(MicroSection))

50        Set tb = New Recordset
60        RecOpenServer 0, tb, sql
70        If Not tb.EOF Then
80            If tb!Valid = 1 Then
90                IsMicroValid = True
100           End If
110       End If



120       Exit Function

IsMicroValid_Error:

          Dim strES As String
          Dim intEL As Integer

130       intEL = Erl
140       strES = Err.Description
150       LogError "basMicro", "IsMicroValid", intEL, strES, sql

End Function




