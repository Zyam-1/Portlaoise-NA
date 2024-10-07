Attribute VB_Name = "basEndocrinology"
Option Explicit

Public Function AreEndResultsPresent(ByVal SampleID As String) As Long

          Dim tb As New Recordset
          Dim sql As String

10        On Error GoTo AreEndResultsPresent_Error

20        AreEndResultsPresent = 0
30        If SampleID = "" Then Exit Function

40        sql = "SELECT count(*) as tot from EndResults WHERE " & _
                "SampleID = '" & SampleID & "'"
50        Set tb = New Recordset
60        Set tb = Cnxn(0).Execute(sql)

70        AreEndResultsPresent = Sgn(tb!Tot)

80        Exit Function

AreEndResultsPresent_Error:

          Dim strES As String
          Dim intEL As Integer

90        intEL = Erl
100       strES = Err.Description
110       LogError "basEndocrinology", "AreEndResultsPresent", intEL, strES

End Function

Public Function eCodeForLongName(ByVal LongName As String) _
       As String

          Dim sql As String
          Dim tb As New Recordset


10        On Error GoTo eCodeForLongName_Error

20        eCodeForLongName = "???"

30        sql = "SELECT * from EndTestDefinitions WHERE longname = '" & LongName & "' and inuse = 1"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        If Not tb.EOF Then
70            eCodeForLongName = Trim(tb!Code)
80        End If




90        Exit Function

eCodeForLongName_Error:

          Dim strES As String
          Dim intEL As Integer



100       intEL = Erl
110       strES = Err.Description
120       LogError "basEndocrinology", "eCodeForLongName", intEL, strES


End Function
Public Function eAnylForLong(ByVal LongName As String) _
       As String

          Dim sql As String
          Dim tb As New Recordset


10        On Error GoTo eAnylForLong_Error

20        eAnylForLong = "???"

30        sql = "SELECT * from EndTestDefinitions WHERE code = '" & LongName & "'"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        If Not tb.EOF Then
70            eAnylForLong = Trim(tb!AnalyserID)
80        End If




90        Exit Function

eAnylForLong_Error:

          Dim strES As String
          Dim intEL As Integer



100       intEL = Erl
110       strES = Err.Description
120       LogError "basEndocrinology", "eAnylForLong", intEL, strES, sql


End Function
Public Function eCodeForShortName(ByVal ShortName As String) _
       As String

          Dim tb As New Recordset
          Dim sql As String


10        On Error GoTo eCodeForShortName_Error

20        eCodeForShortName = "???"

30        sql = "SELECT * from EndTestDefinitions WHERE ShortName = '" & ShortName & "'"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        If Not tb.EOF Then
70            eCodeForShortName = Trim(tb!Code)
80        End If


90        Exit Function

eCodeForShortName_Error:

          Dim strES As String
          Dim intEL As Integer



100       intEL = Erl
110       strES = Err.Description
120       LogError "basEndocrinology", "eCodeForShortName", intEL, strES, sql


End Function
Public Function eAnylForCode(ByVal Code As String) _
       As String

          Dim tb As New Recordset
          Dim sql As String

10        On Error GoTo eAnylForCode_Error

20        eAnylForCode = "???"

30        sql = "SELECT * from EndTestDefinitions WHERE code = '" & Code & "' and inuse = 1"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        If Not tb.EOF Then
70            eAnylForCode = Trim(tb!Analyser & "")
80        End If

90        Exit Function

eAnylForCode_Error:

          Dim strES As String
          Dim intEL As Integer



100       intEL = Erl
110       strES = Err.Description
120       LogError "basEndocrinology", "eAnylForCode", intEL, strES, sql

End Function
Public Function EndLongNameFor(ByVal Code As String) As String

          Dim tb As New Recordset
          Dim sql As String


10        On Error GoTo EndLongNameFor_Error

20        EndLongNameFor = "???"

30        sql = "SELECT * from EndTestDefinitions WHERE Code = '" & Code & "'"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        If Not tb.EOF Then
70            EndLongNameFor = Trim(tb!LongName)
80        End If




90        Exit Function

EndLongNameFor_Error:

          Dim strES As String
          Dim intEL As Integer



100       intEL = Erl
110       strES = Err.Description
120       LogError "basEndocrinology", "EndLongNameFor", intEL, strES, sql


End Function

Public Function EndShortNameFor(ByVal Code As String) As String

          Dim tb As New Recordset
          Dim sql As String


10        On Error GoTo EndShortNameFor_Error

20        EndShortNameFor = "???"

30        sql = "SELECT * from EndTestDefinitions WHERE Code = '" & Code & "'"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        If Not tb.EOF Then
70            EndShortNameFor = Trim(tb!ShortName)
80        End If




90        Exit Function

EndShortNameFor_Error:

          Dim strES As String
          Dim intEL As Integer



100       intEL = Erl
110       strES = Err.Description
120       LogError "basEndocrinology", "EndShortNameFor", intEL, strES, sql


End Function

Public Function QuickInterpEnd(ByVal Result As BIEResult) _
       As String


10        On Error GoTo QuickInterpEnd_Error

20        With Result
30            If Left(.Result, 1) = ">" Then
40                If Val(Mid(.Result, 2)) > .PlausibleHigh Then
50                    QuickInterpEnd = "***"
60                ElseIf Val(Mid(.Result, 2)) < .PlausibleLow Then
70                    QuickInterpEnd = "***"
80                ElseIf Val(Mid(.Result, 2)) < .Low Then
90                    QuickInterpEnd = "Low "
100               ElseIf Val(Mid(.Result, 2)) > .High Then
110                   QuickInterpEnd = "High"
120               End If
130           ElseIf Left(.Result, 1) = "<" Then
140               If Val(Mid(.Result, 2)) > .PlausibleHigh Then
150                   QuickInterpEnd = "***"
160               ElseIf Val(Mid(.Result, 2)) < .PlausibleLow Then
170                   QuickInterpEnd = "***"
180               ElseIf Val(Mid(.Result, 2)) < .Low Then
190                   QuickInterpEnd = "Low "
200               ElseIf Val(Mid(.Result, 2)) > .High Then
210                   QuickInterpEnd = "High"
220               End If
230           ElseIf Val(.Result) > .PlausibleHigh Then
240               QuickInterpEnd = "***"
250           ElseIf Val(.Result) < .PlausibleLow Then
260               QuickInterpEnd = "***"
270           ElseIf Val(.Result) < .Low Then
280               QuickInterpEnd = "Low "
290           ElseIf Val(.Result) > .High Then
300               QuickInterpEnd = "High"
310           Else
320               QuickInterpEnd = "    "
330           End If
340       End With





350       Exit Function

QuickInterpEnd_Error:

          Dim strES As String
          Dim intEL As Integer



360       intEL = Erl
370       strES = Err.Description
380       LogError "basEndocrinology", "QuickInterpEnd", intEL, strES


End Function

Public Sub PrintResultEndWin(ByVal SampleID As String)

          Dim tb As New Recordset
          Dim sql As String
          Dim Clin As String
          Dim GP As String
          Dim Ward As String

10        On Error GoTo PrintResultEndWin_Error

20        Ward = ""
30        GP = ""
40        Clin = ""

50        sql = "SELECT * FROM Demographics WHERE " & _
                "SampleID = '" & SampleID & "'"

60        Set tb = New Recordset
70        RecOpenServer 0, tb, sql
80        If Not tb.EOF Then
90            Ward = tb!Ward & ""
100           GP = tb!GP & ""
110           Clin = tb!Clinician & ""
120       End If

130       sql = "SELECT * FROM PrintPending WHERE " & _
                "Department = 'E' " & _
                "AND SampleID = '" & SampleID & "'"
140       Set tb = New Recordset
150       RecOpenClient 0, tb, sql
160       If tb.EOF Then
170           tb.AddNew
180       End If
190       tb!SampleID = SampleID
200       tb!Department = "E"
210       tb!Initiator = Username
220       tb!pTime = Now
230       tb!Ward = Ward & ""
240       tb!GP = GP & ""
250       tb!Clinician = Clin & ""

260       tb.Update

270       Exit Sub

PrintResultEndWin_Error:

          Dim strES As String
          Dim intEL As Integer



280       intEL = Erl
290       strES = Err.Description
300       LogError "basEndocrinology", "PrintResultEndWin", intEL, strES, sql

End Sub





Public Function Check_End(ByVal Test As String, ByVal Units As String, ByVal SampleType As String) As String
          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo Check_End_Error

20        Check_End = ""

30        sql = "Select * from Endtestdefinitions where shortname = '" & Test & "'"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        If tb.EOF Then
70            Check_End = "Test Name"
80        Else
90            If ListCodeFor("ST", SampleType) <> tb!SampleType Then
100               Check_End = "Sample Type"
110           ElseIf Units <> tb!Units Then
120               Check_End = "Units"
130           End If
140       End If

150       Exit Function

Check_End_Error:

          Dim strES As String
          Dim intEL As Integer

160       intEL = Erl
170       strES = Err.Description
180       LogError "basEndocrinology", "Check_End", intEL, strES, sql

End Function

Public Function TranslateEndResultVirology(Code As String, Result As String) As String

10        On Error GoTo TranslateEndResultVirology_Error

20        If Result = "Negative" Or Result = "Positive" Or Result = "Inconclusive *" Then
30            TranslateEndResultVirology = Result
40        Else

50            Select Case Code
              Case "106":    'HBsAg
60                If Result < 1 Then
70                    Result = "Negative"
80                ElseIf Result >= 1 Then
90                    Result = "Inconclusive *"
100               End If
110           Case "118":    'AUSAB
120               If Result < 10 Then
130                   Result = "Negative"
140               ElseIf Result >= 10 Then
150                   Result = "Positive"
160               End If
170           Case "126":    'HepBCo
180               If Result >= 1.001 And Result <= 3 Then
190                   Result = "Negative"
200               ElseIf Result >= 0 And Result <= 1 Then
210                   Result = "Inconclusive *"
220               End If
230           Case "841":    'HCV
240               If Result < 1 Then
250                   Result = "Negative"
260               ElseIf Result >= 1 Then
270                   Result = "Inconclusive *"
280               End If
290           Case "817":    'HIV
300               If Result < 0.9 Then
310                   Result = "Negative"
320               ElseIf Result >= 0.9 Then
330                   Result = "Inconclusive *"
340               End If
350           End Select
360           TranslateEndResultVirology = Result
370       End If


380       Exit Function

TranslateEndResultVirology_Error:

          Dim strES As String
          Dim intEL As Integer

390       intEL = Erl
400       strES = Err.Description
410       LogError "basEndocrinology", "TranslateEndResultVirology", intEL, strES

End Function
